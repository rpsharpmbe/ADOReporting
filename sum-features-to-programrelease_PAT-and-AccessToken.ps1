<#
.SYNOPSIS
  Runs a WIQL query to find Features, sums a numeric field (Effort / Story Points),
  and writes the total into a target Program Release work item field.

.AUTH
  Supports:
   - PAT via env var ADO_PAT
   - Bearer token via env var SYSTEM_ACCESSTOKEN (Azure DevOps pipeline)
#>

param(
  [Parameter(Mandatory)] [string] $Org,
  [Parameter(Mandatory)] [string] $Project,
  [Parameter(Mandatory)] [string] $IterationPath,
  [Parameter(Mandatory)] [int]    $ReleaseId,

  [string] $PointsFieldRef = "Microsoft.VSTS.Scheduling.Effort",
  [Parameter(Mandatory)] [string] $ReleasePointsFieldRef,

  [string] $ApiVersion = "7.1"
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# AUTH (PAT or Bearer)
# ------------------------------------------------------------
$headers = @{}

if ($env:SYSTEM_ACCESSTOKEN) {
  Write-Host "Auth mode: Bearer (System.AccessToken)"
  $headers.Authorization = "Bearer $($env:SYSTEM_ACCESSTOKEN)"
}
elseif ($env:ADO_PAT) {
  Write-Host "Auth mode: PAT"
  $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($env:ADO_PAT)"))
  $headers.Authorization = "Basic $b64"
}
else {
  throw "No auth available. Set ADO_PAT or SYSTEM_ACCESSTOKEN."
}

$headers.Accept = "application/json"
$headers."X-TFS-FedAuthRedirect" = "Suppress"

# ------------------------------------------------------------
# REST helper
# ------------------------------------------------------------
function Invoke-AdoRest {
  param(
    [Parameter(Mandatory)] [ValidateSet("GET","POST","PATCH")] [string] $Method,
    [Parameter(Mandatory)] [string] $Url,
    [hashtable] $Headers,
    [string] $BodyJson,
    [string] $ContentType
  )

  try {
    $h = @{} + $Headers
    if ($ContentType) { $h["Content-Type"] = $ContentType }

    if ($BodyJson) {
      return Invoke-RestMethod -Method $Method -Uri $Url -Headers $h -Body $BodyJson
    } else {
      return Invoke-RestMethod -Method $Method -Uri $Url -Headers $h
    }
  }
  catch {
    Write-Host ""
    Write-Host "ADO REST call failed" -ForegroundColor Red
    Write-Host "Method: $Method" -ForegroundColor Red
    Write-Host "URL:    $Url" -ForegroundColor Red
    if ($BodyJson) {
      Write-Host "Body:" -ForegroundColor Red
      Write-Host $BodyJson
    }
    if ($_.Exception.Response) {
      $resp = $_.Exception.Response
      Write-Host ("HTTP: {0} {1}" -f [int]$resp.StatusCode, $resp.StatusDescription) -ForegroundColor Red
      try {
        $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
        Write-Host ($sr.ReadToEnd())
      } catch {}
    }
    throw
  }
}

# ------------------------------------------------------------
# Setup
# ------------------------------------------------------------
$base = "https://dev.azure.com/$Org/$Project"

Write-Host ""
Write-Host "Org:       $Org"
Write-Host "Project:   $Project"
Write-Host "Iteration: $IterationPath"
Write-Host "Sum field: $PointsFieldRef"
Write-Host "Target WI: #$ReleaseId -> $ReleasePointsFieldRef"
Write-Host ""

# ------------------------------------------------------------
# WIQL
# ------------------------------------------------------------
$wiql = @"
SELECT [System.Id]
FROM WorkItems
WHERE
  [System.TeamProject] = @project
  AND [System.WorkItemType] = 'Feature'
  AND [System.State] <> 'Removed'
  AND [System.IterationPath] UNDER '$IterationPath'
  AND [System.Tags] CONTAINS 'NGR-NDR-R2.0'
ORDER BY [System.ChangedDate] DESC
"@

Write-Host "WIQL:"
Write-Host $wiql
Write-Host ""

# ------------------------------------------------------------
# 1) Run WIQL
# ------------------------------------------------------------
Write-Host "1) Running WIQL..."

$wiqlUrl  = "$base/_apis/wit/wiql?api-version=$ApiVersion"
$wiqlBody = @{ query = $wiql } | ConvertTo-Json -Depth 20

$wiqlResult = Invoke-AdoRest `
  -Method POST `
  -Url $wiqlUrl `
  -Headers $headers `
  -BodyJson $wiqlBody `
  -ContentType "application/json"

$ids = @()
if ($wiqlResult.workItems) {
  $ids = $wiqlResult.workItems | ForEach-Object { $_.id }
}

Write-Host ("WIQL returned {0} work item(s)." -f $ids.Count)

# ------------------------------------------------------------
# 2) Sum field values
# ------------------------------------------------------------
$total = 0.0

if ($ids.Count -gt 0) {
  Write-Host "2) Fetching fields in batches and summing..."

  $batchUrl = "$base/_apis/wit/workitemsbatch?api-version=$ApiVersion"

  $batchBody = @{
    ids    = $ids
    fields = @($PointsFieldRef)
  } | ConvertTo-Json -Depth 20

  $batch = Invoke-AdoRest `
    -Method POST `
    -Url $batchUrl `
    -Headers $headers `
    -BodyJson $batchBody `
    -ContentType "application/json"

  foreach ($wi in $batch.value) {
    $val = 0
    if ($wi.fields.PSObject.Properties.Name -contains $PointsFieldRef) {
      $val = $wi.fields.$PointsFieldRef
    }
    if ($null -eq $val -or $val -eq "") { $val = 0 }
    $total += [double]$val
  }
}

Write-Host ("Computed total = {0}" -f $total)

# ------------------------------------------------------------
# 3) Patch Program Release (MUST be JSON Patch ARRAY)
# ------------------------------------------------------------
Write-Host "3) Writing total into Program Release #$ReleaseId..."

$getUrl = "$base/_apis/wit/workitems/${ReleaseId}?fields=$([uri]::EscapeDataString($ReleasePointsFieldRef))&api-version=$ApiVersion"
$current = Invoke-AdoRest -Method GET -Url $getUrl -Headers $headers

$op = "add"
if ($current.fields.PSObject.Properties.Name -contains $ReleasePointsFieldRef) {
  $op = "replace"
}

# Force array-of-ops in a way PS cannot collapse
$patchOps = @()
$patchOps += [pscustomobject]@{
  op    = $op
  path  = "/fields/$ReleasePointsFieldRef"
  value = $total
}

$patchJson = ($patchOps | ConvertTo-Json -Depth 20)

# Sanity: ensure it begins with '['
if (-not $patchJson.TrimStart().StartsWith("[")) {
  # Hard fallback
  $patchJson = "[" + $patchJson + "]"
}

Write-Host "PATCH document being sent:"
Write-Host $patchJson

$patchUrl  = "$base/_apis/wit/workitems/${ReleaseId}?api-version=$ApiVersion"

$result = Invoke-AdoRest `
  -Method PATCH `
  -Url $patchUrl `
  -Headers $headers `
  -BodyJson $patchJson `
  -ContentType "application/json-patch+json"

Write-Host ("Updated work item {0}. Field '{1}' set to {2}" -f $result.id, $ReleasePointsFieldRef, $total)
Write-Host ""
Write-Host "DONE ✔"
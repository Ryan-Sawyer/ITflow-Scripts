#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
<#
.SYNOPSIS
    Syncs devices from Microsoft Graph into ITflow (serial number + hostname).

.DESCRIPTION
    Prompts for all required credentials interactively at runtime, then authenticates
    to Microsoft Graph using the Microsoft.Graph PowerShell module (app-only / client
    credentials), retrieves all devices, and creates or updates matching assets in
    ITflow via its REST API.

.NOTES
    Install required modules (once):
        Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser

    Required Graph API permissions (Application):
        - Device.Read.All

    Required ITflow permissions:
        - Assets: Read + Write
#>

[CmdletBinding(SupportsShouldProcess)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Logging
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = @{ INFO = 'Cyan'; WARN = 'Yellow'; ERROR = 'Red' }[$Level]
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Interactive parameter collection
# ─────────────────────────────────────────────────────────────────────────────
function Read-RunParams {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   Microsoft Graph → ITflow Device Sync Setup     ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "── Azure App Registration ───────────────────────────" -ForegroundColor DarkCyan
    $tenantId     = (Read-Host "  Tenant ID").Trim()
    $clientId     = (Read-Host "  Client ID").Trim()
    $clientSecret = Read-Host "  Client Secret" -AsSecureString

    Write-Host ""
    Write-Host "── ITflow ───────────────────────────────────────────" -ForegroundColor DarkCyan
    $itflowUrl = (Read-Host "  Base URL (e.g. https://itflow.yourdomain.com)").Trim().TrimEnd('/')
    $itflowKey = (Read-Host "  API Key").Trim()

    $itflowClientIdRaw = (Read-Host "  Client ID (numeric)").Trim()
    while (-not ($itflowClientIdRaw -match '^\d+$')) {
        Write-Host "  ⚠  Please enter a numeric value." -ForegroundColor Yellow
        $itflowClientIdRaw = (Read-Host "  Client ID (numeric)").Trim()
    }

    Write-Host ""

    return @{
        TenantId       = $tenantId
        ClientId       = $clientId
        ClientSecret   = $clientSecret
        ITflowBaseUrl  = $itflowUrl
        ITflowApiKey   = $itflowKey
        ITflowClientId = [int]$itflowClientIdRaw
    }
}

# Collect all inputs up front, then assign to script-scoped variables
$cfg = Read-RunParams

$TenantId       = $cfg.TenantId
$ClientId       = $cfg.ClientId
$ClientSecret   = $cfg.ClientSecret
$ITflowBaseUrl  = $cfg.ITflowBaseUrl
$ITflowApiKey   = $cfg.ITflowApiKey
$ITflowClientId = $cfg.ITflowClientId

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Graph authentication via Microsoft.Graph module
# ─────────────────────────────────────────────────────────────────────────────
function Connect-ToGraph {
    Write-Log "Connecting to Microsoft Graph (app-only)..."

    $credential = [System.Management.Automation.PSCredential]::new($ClientId, $ClientSecret)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome

    $ctx = Get-MgContext
    Write-Log "Connected as '$($ctx.AppName)' to tenant '$($ctx.TenantId)'."
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Retrieve all devices via Microsoft.Graph module
# ─────────────────────────────────────────────────────────────────────────────
function Get-AllGraphDevices {
    Write-Log "Fetching devices from Microsoft Graph..."

    # serialNumber, manufacturer, and model are top-level Graph device properties
    $devices = @(Get-MgDeviceManagementManagedDevice -All)
# 'id,DeviceName,operatingSystem,operatingSystemVersion,physicalIds,serialNumber,manufacturer,model'

    Write-Log "Total devices retrieved: $($devices.Count)"
    return $devices
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Resolve serial number
# Priority: top-level serialNumber → [SERIAL_NUMBER] physicalId tag → [OrderID] tag
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# REGION: ITflow API helpers
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-ITflowApi {
    param(
        [string]    $Method,
        [string]    $Endpoint,
        [hashtable] $Body = @{}
    )

    # ITflow authenticates via api_key query param (GET) or JSON body field (POST/PUT)
    $baseUri   = "$ITflowBaseUrl/api/v1/$Endpoint"
    $separator = if ($Endpoint -like '*?*') { '&' } else { '?' }

    $headers = @{ 'Content-Type' = 'application/json' }

    $params = @{
        Method                = $Method
        Headers               = $headers
        AllowInsecureRedirect = $true
        ErrorAction           = 'Stop'
    }

    if ($Method -in 'POST','PUT','PATCH') {
        # api_key goes in the JSON body for write requests
        $Body['api_key'] = $ITflowApiKey
        $params['Uri']   = $baseUri
        $params['Body']  = ($Body | ConvertTo-Json -Depth 5)
    } else {
        # api_key goes as a query param for read requests
        $params['Uri'] = "$baseUri$($separator)api_key=$ITflowApiKey"
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        # Safely extract whatever error detail is available
        $statusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        $detail     = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Log "ITflow API error [$Method $Endpoint] HTTP ${statusCode}: $detail" -Level WARN
        return $null
    }
}

function Get-ITflowAssetLookup {
    Write-Log "Fetching existing ITflow assets (client $ITflowClientId)..."

    $result = Invoke-ITflowApi -Method GET -Endpoint "assets/read.php?client_id=$ITflowClientId&limit=1000"
    Write-Host ($result | ConvertTo-Json -Depth 3)
    $lookup = @{}

    if ($null -eq $result) {
        Write-Log "  ITflow returned no response — starting with empty asset list." -Level WARN
        return $lookup
    }

    # Debug: log the top-level keys so we can see the actual response shape
    $keys = ($result | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -join ', '
    Write-Log "  ITflow response keys: $keys"

    # ITflow wraps records in a 'data' key, but guard against other shapes
    $assets = $null
    if ($result.PSObject.Properties['data']) {
        $assets = $result.data
    } elseif ($result -is [System.Collections.IEnumerable] -and $result -isnot [string]) {
        # Response is a bare array
        $assets = $result
    } else {
        Write-Log "  Unexpected ITflow response shape — cannot parse assets. Raw: $($result | ConvertTo-Json -Depth 2 -Compress)" -Level WARN
        return $lookup
    }

    foreach ($asset in $assets) {
        if ($asset.asset_serial) { $lookup[$asset.asset_serial] = $asset }
    }

    Write-Log "  $($lookup.Count) existing assets found with serial numbers."
    return $lookup
}

function New-ITflowAsset {
    param([string]$Hostname, [string]$Serial, [string]$OS, [string]$Make, [string]$Model)

    return Invoke-ITflowApi -Method POST -Endpoint 'assets/create.php' -Body @{
        asset_name         = $Hostname
        asset_serial       = $Serial
        asset_type         = 'Workstation'   # adjust to match your ITflow asset types
        asset_os           = $OS
        asset_make         = $Make
        asset_model        = $Model
        client_id          = $ITflowClientId
    }
}

function Update-ITflowAsset {
    param([int]$AssetId, [string]$Hostname, [string]$Serial, [string]$OS, [string]$Make, [string]$Model)

    return Invoke-ITflowApi -Method PUT -Endpoint "assets/update.php?asset_id=$AssetId" -Body @{
        asset_name         = $Hostname
        asset_serial       = $Serial
        asset_os           = $OS
        asset_make         = $Make
        asset_model        = $Model
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Main sync
# ─────────────────────────────────────────────────────────────────────────────
function Start-DeviceSync {
    Write-Log "=== Starting Microsoft Graph → ITflow device sync ==="

    try {
        Connect-ToGraph

        $graphDevices = Get-AllGraphDevices
        $itflowAssets = Get-ITflowAssetLookup

        $stats = @{ Created = 0; Updated = 0; Skipped = 0; Errors = 0 }

        foreach ($device in $graphDevices) {

            $hostname = $device.DeviceName
            # AdditionalProperties holds fields returned by -Property that aren't
            # native PS object members (serialNumber, manufacturer, model)
            $ap = @{}
	    if ($device.AdditionalProperties) {
    	      $ap = $device.AdditionalProperties
	    }

	    Write-Log "DEBUG RAW DEVICE: $($device | ConvertTo-Json -Depth 3 -Compress)"
	    Write-Log "DEBUG AdditionalProperties: $($device.AdditionalProperties | ConvertTo-Json -Depth 3 -Compress)"

	    $serial = $device.SerialNumber
	        #-SerialNumber ($ap['serialNumber'] ?? $ap['SerialNumber']) `
	        #-PhysicalIds  $device.PhysicalIds

	    $make  = if ($ap['manufacturer'] ?? $ap['Manufacturer']) { 
	        [string]($ap['manufacturer'] ?? $ap['Manufacturer']) 
	    } else { '' }

	    $model = if ($ap['model'] ?? $ap['Model']) { 
	        [string]($ap['model'] ?? $ap['Model']) 
	    } else { '' }

	    $os = ''

	    if ($device.PSObject.Properties['OperatingSystem']) {
	        $os += $device.OperatingSystem
	    }

	    if ($device.PSObject.Properties['OsVersion']) {
	        $os += " $($device.OsVersion)"
	    }

$os = $os.Trim()

	    Write-Output $ap['serialNumber']
            if (-not $hostname) {
                Write-Log "Device '$($device.Id)' has no DisplayName — skipping." -Level WARN
                $stats.Skipped++; continue
            }

            if (-not $serial) {
                Write-Log "Device '$hostname' has no parseable serial number — skipping." -Level WARN
                $stats.Skipped++; continue
            }

            Write-Log "Processing: '$hostname' | Serial: $serial | Make: $make | Model: $model | OS: $os"

            if ($itflowAssets.ContainsKey($serial)) {
                $existing = $itflowAssets[$serial]
                if ($PSCmdlet.ShouldProcess("ITflow asset $($existing.asset_id)", "Update '$hostname'")) {
                    $r = Update-ITflowAsset -AssetId $existing.asset_id -Hostname $hostname -Serial $serial -OS $os -Make $make -Model $model
                    if ($r) { Write-Log "  ✔ Updated asset ID $($existing.asset_id)"; $stats.Updated++ }
                    else    { Write-Log "  ✘ Update failed for '$hostname'" -Level ERROR; $stats.Errors++ }
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess('ITflow', "Create asset '$hostname'")) {
                    $r = New-ITflowAsset -Hostname $hostname -Serial $serial -OS $os -Make $make -Model $model
                    if ($r) { Write-Log "  ✔ Created asset for '$hostname'"; $stats.Created++ }
                    else    { Write-Log "  ✘ Create failed for '$hostname'" -Level WARN; $stats.Errors++ }
                }
            }
        }

        Write-Log "=== Sync Complete ==="
        Write-Log "  Created : $($stats.Created)"
        Write-Log "  Updated : $($stats.Updated)"
        Write-Log "  Skipped : $($stats.Skipped)"
        Write-Log "  Errors  : $($stats.Errors)"
    }
    finally {
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected from Microsoft Graph."
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
Start-DeviceSync

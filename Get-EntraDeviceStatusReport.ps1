#Requires -Version 7

<#
.SYNOPSIS
    Builds a report of Microsoft Entra devices and their join status.

.DESCRIPTION
    Connects to Microsoft Graph, pulls all devices in the tenant, maps the device trust type
    to a friendly Entra status, and removes duplicate records by keeping the newest entry for
    each device key.

.PARAMETER ExportPath
    Optional CSV path to export the final report.

.PARAMETER CsvPath
    Alias for ExportPath to make CSV export easier to discover.

.PARAMETER UseDeviceCode
    Use device code authentication instead of interactive sign-in.

.PARAMETER DeduplicationProperty
    Property used to group duplicate devices before keeping the newest record.
    DeviceId is the safest default. DisplayName is useful when your tenant has repeated
    device registrations with the same name.

.PARAMETER ShowGridView
    Opens the final report in Out-GridView when available.

.EXAMPLE
    .\Get-EntraDeviceStatusReport.ps1

.EXAMPLE
    .\Get-EntraDeviceStatusReport.ps1 -ExportPath C:\Reports\EntraDevices.csv

.EXAMPLE
    .\Get-EntraDeviceStatusReport.ps1 -UseDeviceCode -DeduplicationProperty DisplayName

.NOTES
    Required Graph scopes: Device.Read.All, Directory.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [Alias('CsvPath')]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceCode,

    [Parameter(Mandatory = $false)]
    [ValidateSet('DeviceId', 'DisplayName')]
    [string]$DeduplicationProperty = 'DeviceId',

    [Parameter(Mandatory = $false)]
    [switch]$ShowGridView
)

$RequiredScopes = @('Device.Read.All', 'Directory.Read.All')

foreach ($RequiredCommand in @('Connect-MgGraph', 'Disconnect-MgGraph', 'Invoke-MgGraphRequest')) {
    if (-not (Get-Command $RequiredCommand -ErrorAction SilentlyContinue)) {
        throw "Required Microsoft Graph command '$RequiredCommand' was not found. Import Microsoft.Graph.Authentication before running this script."
    }
}

function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param()

    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan

    try {
        if ($UseDeviceCode) {
            Connect-MgGraph -Scopes $RequiredScopes -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
        }
        else {
            Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop | Out-Null
        }
    }
    catch {
        throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }
}

function Get-EntraDeviceStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$TrustType = $null
    )

    if ([string]::IsNullOrWhiteSpace($TrustType)) {
        return 'Unknown'
    }

    switch ($TrustType) {
        'Workplace' { 'Entra ID Registered' }
        'AzureAd' { 'Entra ID Joined' }
        'ServerAd' { 'Entra ID Hybrid Joined' }
        default { 'Unknown' }
    }
}

function Get-DeviceGroupingKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Device
    )

    $value = $null

    switch ($DeduplicationProperty) {
        'DeviceId' {
            if ($Device.DeviceId) {
                $value = $Device.DeviceId
            }
            elseif ($Device.DisplayName) {
                $value = $Device.DisplayName.Trim().ToLowerInvariant()
            }
        }
        'DisplayName' {
            if ($Device.DisplayName) {
                $value = $Device.DisplayName.Trim().ToLowerInvariant()
            }
            elseif ($Device.DeviceId) {
                $value = $Device.DeviceId
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = $Device.Id
    }

    return $value
}

function Get-AllGraphDevices {
    [CmdletBinding()]
    param()

    $uri = 'https://graph.microsoft.com/v1.0/devices?$select=id,displayName,deviceId,trustType,registrationDateTime,approximateLastSignInDateTime&$top=999'
    $devices = @()

    while ($uri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        if ($response.value) {
            $devices += $response.value
        }

        $uri = $response.'@odata.nextLink'
    }

    return $devices
}

Connect-ToMicrosoftGraph

Write-Host 'Retrieving devices from Microsoft Entra ID...' -ForegroundColor Cyan

try {
    $RawDevices = Get-AllGraphDevices
}
catch {
    Disconnect-MgGraph | Out-Null
    throw "Failed to retrieve devices: $($_.Exception.Message)"
}

if (-not $RawDevices) {
    Write-Warning 'No devices were returned from Microsoft Graph.'
    Disconnect-MgGraph | Out-Null
    return
}

$Report = foreach ($Device in $RawDevices) {
    [PSCustomObject]@{
        GroupingKey          = Get-DeviceGroupingKey -Device $Device
        Id                   = $Device.Id
        DeviceId             = $Device.DeviceId
        DisplayName          = $Device.DisplayName
        TrustType            = $Device.TrustType
        Status               = Get-EntraDeviceStatus -TrustType $Device.TrustType
        RegistrationDateTime = $Device.RegistrationDateTime
        LastSignInDateTime   = $Device.approximateLastSignInDateTime
    }
}

$DedupedReport = $Report |
Group-Object -Property GroupingKey |
ForEach-Object {
    $duplicateCount = $_.Count - 1
    $_.Group |
    Sort-Object -Property @{ Expression = {
            if ($_.LastSignInDateTime) { $_.LastSignInDateTime }
            elseif ($_.RegistrationDateTime) { $_.RegistrationDateTime }
            else { [datetime]::MinValue }
        } 
    } -Descending |
    Select-Object -First 1 |
    Select-Object *,
    @{ Name = 'HasDuplicates'; Expression = { $duplicateCount -gt 0 } },
    @{ Name = 'DuplicateCount'; Expression = { $duplicateCount } }
} |
Sort-Object -Property DisplayName, RegistrationDateTime

$DuplicateCount = $Report.Count - $DedupedReport.Count

Write-Host "Retrieved $($RawDevices.Count) device records." -ForegroundColor Green
Write-Host "Final report contains $($DedupedReport.Count) unique devices." -ForegroundColor Green
if ($DuplicateCount -gt 0) {
    Write-Host "Removed $DuplicateCount duplicate record(s) by keeping the newest registration entry per $DeduplicationProperty." -ForegroundColor Yellow
}

$FinalReport = $DedupedReport |
Select-Object DisplayName, DeviceId, Id, Status, HasDuplicates, DuplicateCount, TrustType, RegistrationDateTime, LastSignInDateTime

if ($ShowGridView) {
    try {
        $FinalReport | Out-GridView -Title 'Entra ID Device Status Report'
    }
    catch {
        Write-Warning 'Out-GridView is not available in this session.'
    }
}
else {
    $FinalReport |
    Format-Table -Property DisplayName, DeviceId, Status, HasDuplicates, DuplicateCount, TrustType, RegistrationDateTime, LastSignInDateTime -AutoSize |
    Out-String -Width 300 |
    Write-Output
}

if ($ExportPath) {
    $exportDirectory = Split-Path -Path $ExportPath -Parent
    if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
        New-Item -ItemType Directory -Path $exportDirectory -Force | Out-Null
    }

    $FinalReport | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported to $ExportPath" -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null
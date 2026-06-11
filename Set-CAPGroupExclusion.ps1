<#
.SYNOPSIS
    Adds a group as an exclusion to all Conditional Access Policies in the tenant.

.DESCRIPTION
    Iterates through all Conditional Access Policies and ensures the specified group
    is included in the excluded groups list of each policy.

.PARAMETER GroupId
    The Object ID of the Entra ID group to exclude from all Conditional Access Policies.

.PARAMETER AuthMethod
    Authentication method to use when connecting to Microsoft Graph.
    - "Interactive": Opens a browser window for interactive login (default).
    - "DeviceCode": Uses the device code flow (useful for headless or remote sessions).

.EXAMPLE
    .\Set-CAPGroupExclusion.ps1 -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    .\Set-CAPGroupExclusion.ps1 -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -AuthMethod DeviceCode

.NOTES
    Requires Microsoft.Graph PowerShell module.
    Required scopes: Policy.ReadWrite.ConditionalAccess, Policy.Read.All
#>

#Requires -Version 7
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Object ID of the group to exclude from all CAPs")]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$GroupId,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Interactive", "DeviceCode")]
    [string]$AuthMethod = "Interactive"
)

$RequiredScopes = @("Policy.ReadWrite.ConditionalAccess", "Policy.Read.All")

Write-Host "Connecting to Microsoft Graph using $AuthMethod authentication..." -ForegroundColor Cyan

if ($AuthMethod -eq "DeviceCode") {
    Connect-MgGraph -Scopes $RequiredScopes -UseDeviceCode
}
else {
    Connect-MgGraph -Scopes $RequiredScopes
}

# Verify the group exists before iterating all policies
Write-Host "`nValidating group '$GroupId'..." -ForegroundColor Cyan
try {
    $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    Write-Host "Group found: '$($group.DisplayName)'" -ForegroundColor Green
}
catch {
    Write-Error "Group with ID '$GroupId' was not found in the tenant. Aborting.`n$_"
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Retrieve all Conditional Access Policies
Write-Host "`nRetrieving all Conditional Access Policies..." -ForegroundColor Cyan
$policies = Get-MgIdentityConditionalAccessPolicy -All

if (-not $policies -or $policies.Count -eq 0) {
    Write-Warning "No Conditional Access Policies found in the tenant."
    Disconnect-MgGraph | Out-Null
    exit 0
}

Write-Host "Found $($policies.Count) Conditional Access Policies.`n" -ForegroundColor Green

$updated = 0
$skipped = 0
$failed = 0

foreach ($policy in $policies) {
    $policyName = $policy.DisplayName
    $policyId = $policy.Id

    # Current list of excluded groups (may be null)
    $existingExclusions = @($policy.Conditions.Users.ExcludeGroups)

    if ($existingExclusions -contains $GroupId) {
        Write-Host "  [SKIP]    '$policyName' — group already excluded." -ForegroundColor Yellow
        $skipped++
        continue
    }

    $newExclusions = $existingExclusions + $GroupId

    $body = @{
        conditions = @{
            users = @{
                excludeGroups = $newExclusions
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($policyName, "Add group exclusion '$($group.DisplayName)'")) {
        try {
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -BodyParameter $body -ErrorAction Stop
            Write-Host "  [UPDATED] '$policyName'" -ForegroundColor Green
            $updated++
        }
        catch {
            Write-Host "  [FAILED]  '$policyName' — $_" -ForegroundColor Red
            $failed++
        }
    }
}

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "  Updated : $updated" -ForegroundColor Green
Write-Host "  Skipped : $skipped" -ForegroundColor Yellow
Write-Host "  Failed  : $failed"  -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Gray" })

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph." -ForegroundColor Cyan

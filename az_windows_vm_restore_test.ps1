param(
    [Parameter(Mandatory = $true)]
    [string[]]$VMNames,

    [Parameter(Mandatory = $true)]
    [string]$SourceResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$TargetResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$TargetVNetName,

    [Parameter(Mandatory = $true)]
    [string]$TargetVNetResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$TargetSubnetName,

    [string]$SubscriptionId,

    [string]$TenantId,

    [string]$RecoveryPoint = "Latest",

    [datetime]$RecoveryPointDateTime,

    [int]$RestoreTimeoutMinutes = 120,

    [int]$VMStartupTimeoutMinutes = 20,

    [switch]$Cleanup,

    [switch]$GenerateHtmlReport
)

# ============================================================
# Restore-VM-Integrity-Drill
#
# PowerShell 7.4
#
# Purpose:
#   1. Locate Azure Backup recovery point
#   2. Restore VM as a new VM
#   3. Wait for restore completion
#   4. Start restored VM
#   5. Run Windows integrity tests
#   6. Produce JSON/HTML results
#   7. Optionally delete restored VM
#
# Execution model:
#   Designed to run from a local workstation or developer laptop
#   instead of an Azure Automation Account.
# ============================================================

$ErrorActionPreference = "Stop"

function Assert-AzModule {
    param(
        [string[]]$ModuleNames
    )

    foreach ($ModuleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            throw "Required Az module '$ModuleName' is not installed. Install it with: Install-Module $ModuleName -Scope CurrentUser"
        }
    }
}

function Connect-LocalAzureContext {
    param(
        [string]$SubscriptionId,
        [string]$TenantId
    )

    Disable-AzContextAutosave -Scope Process

    $Context = Get-AzContext

    if (-not $Context) {
        if ($SubscriptionId) {
            if ($TenantId) {
                $Context = Connect-AzAccount -Subscription $SubscriptionId -Tenant $TenantId
            }
            else {
                $Context = Connect-AzAccount -Subscription $SubscriptionId
            }
        }
        else {
            $Context = Connect-AzAccount
        }
    }
    elseif ($SubscriptionId) {
        Set-AzContext -SubscriptionId $SubscriptionId -Tenant $TenantId -ErrorAction Stop | Out-Null
        $Context = Get-AzContext
    }

    if (-not $Context) {
        throw "Unable to authenticate to Azure from this workstation."
    }

    return $Context
}

$ExecutionId = Get-Date -Format "yyyyMMdd-HHmmss"
$ExecutionStart = Get-Date

$Results = [System.Collections.Generic.List[object]]::new()

Write-Output "============================================================"
Write-Output " Azure VM Restore Integrity Drill"
Write-Output " Execution ID: $ExecutionId"
Write-Output " Run mode: Local workstation"
Write-Output "============================================================"

# ------------------------------------------------------------
# Authentication
# ------------------------------------------------------------

Write-Output "Validating Azure PowerShell modules..."
Assert-AzModule -ModuleNames @(
    "Az.Accounts",
    "Az.Compute",
    "Az.Network",
    "Az.RecoveryServices",
    "Az.Resources"
)

Write-Output "Authenticating to Azure using your local account..."
$AzureContext = Connect-LocalAzureContext -SubscriptionId $SubscriptionId -TenantId $TenantId

Write-Output "Authenticated as: $($AzureContext.Account.Id)"
if ($AzureContext.Subscription) {
    Write-Output "Subscription: $($AzureContext.Subscription.Name) ($($AzureContext.Subscription.Id))"
}

# ------------------------------------------------------------
# Helper: Add result
# ------------------------------------------------------------

function Add-TestResult {
    param(
        [string]$VMName,
        [string]$Test,
        [string]$Status,
        [string]$Details,
        [object]$Value = $null
    )

    $Results.Add(
        [PSCustomObject]@{
            VM        = $VMName
            Test      = $Test
            Status    = $Status
            Details   = $Details
            Value     = $Value
            Timestamp = (Get-Date).ToString("s")
        }
    )
}

# ------------------------------------------------------------
# Helper: Find backup item
# ------------------------------------------------------------

function Get-VMBackupItem {
    param(
        [object]$Vault,
        [string]$VMName
    )

    Write-Output "Searching backup item for VM: $VMName"

    $items = Get-AzRecoveryServicesBackupItem `
        -BackupManagementType AzureVM `
        -WorkloadType AzureVM `
        -VaultId $Vault.Id

    $item = $items |
        Where-Object {
            $_.FriendlyName -eq $VMName -or
            $_.Name -eq $VMName
        } |
        Select-Object -First 1

    return $item
}

# ------------------------------------------------------------
# Helper: Select recovery point
# ------------------------------------------------------------

function Get-RecoveryPointForVM {
    param(
        [object]$BackupItem,
        [object]$Vault,
        [string]$PreferredRecoveryPoint,
        [datetime]$PreferredRecoveryPointTime
    )

    $EndDate = Get-Date
    $StartDate = $EndDate.AddDays(-30)

    Write-Output "Searching recovery points..."

    $RecoveryPoints = Get-AzRecoveryServicesBackupRecoveryPoint `
        -Item $BackupItem `
        -StartDate $StartDate.ToUniversalTime() `
        -EndDate $EndDate.ToUniversalTime() `
        -VaultId $Vault.Id

    if (-not $RecoveryPoints) {
        throw "No recovery points found."
    }

    switch ($PreferredRecoveryPoint.ToLowerInvariant()) {
        "latest" {
            return $RecoveryPoints |
                Sort-Object RecoveryPointTime -Descending |
                Select-Object -First 1
        }
        "earliest" {
            return $RecoveryPoints |
                Sort-Object RecoveryPointTime |
                Select-Object -First 1
        }
        default {
            $TargetDate = $null

            if ($PreferredRecoveryPointTime) {
                $TargetDate = $PreferredRecoveryPointTime
            }
            else {
                try {
                    $TargetDate = [datetime]::Parse($PreferredRecoveryPoint)
                }
                catch {
                    throw "RecoveryPoint '$PreferredRecoveryPoint' is not valid. Use 'Latest', 'Earliest', or a valid datetime string."
                }
            }

            $ChosenPoint = $RecoveryPoints |
                Where-Object { $_.RecoveryPointTime -le $TargetDate } |
                Sort-Object RecoveryPointTime -Descending |
                Select-Object -First 1

            if (-not $ChosenPoint) {
                throw "No recovery point found on or before '$TargetDate'."
            }

            return $ChosenPoint
        }
    }
}

# ------------------------------------------------------------
# Helper: Wait for restore job
# ------------------------------------------------------------

function Wait-ForRestoreJob {
    param(
        [object]$Job,
        [object]$Vault,
        [int]$TimeoutMinutes
    )

    $Start = Get-Date

    do {

        Start-Sleep -Seconds 30

        $CurrentJob = Get-AzRecoveryServicesBackupJob `
            -VaultId $Vault.Id `
            -JobId $Job.JobId

        Write-Output ("Restore job status: {0}" -f $CurrentJob.Status)

        if ($CurrentJob.Status -in @(
            "Completed",
            "Failed",
            "Cancelled",
            "Canceled"
        )) {
            return $CurrentJob
        }

        if (((Get-Date) - $Start).TotalMinutes -gt $TimeoutMinutes) {
            throw "Restore operation exceeded timeout of $TimeoutMinutes minutes."
        }

    } while ($true)
}

# ------------------------------------------------------------
# Helper: Wait for VM
# ------------------------------------------------------------

function Wait-ForVMReady {
    param(
        [string]$ResourceGroup,
        [string]$VMName,
        [int]$TimeoutMinutes
    )

    $Start = Get-Date

    do {

        Start-Sleep -Seconds 20

        try {

            $Status = Get-AzVM `
                -ResourceGroupName $ResourceGroup `
                -Name $VMName `
                -Status

            $PowerState = (
                $Status.Statuses |
                Where-Object {
                    $_.Code -like "PowerState/*"
                }
            ).DisplayStatus

            $ProvisioningState = (
                $Status.Statuses |
                Where-Object {
                    $_.Code -like "ProvisioningState/*"
                }
            ).DisplayStatus

            Write-Output "VM state: $PowerState / $ProvisioningState"

            if (
                $PowerState -eq "VM running" -and
                $ProvisioningState -eq "Provisioning succeeded"
            ) {
                return $true
            }

        }
        catch {
            Write-Output "VM not ready yet: $($_.Exception.Message)"
        }

        if (((Get-Date) - $Start).TotalMinutes -gt $TimeoutMinutes) {
            return $false
        }

    } while ($true)
}

# ------------------------------------------------------------
# Windows validation script
#
# This script runs INSIDE the restored VM using Azure VM Run
# Command. It does not require WinRM/RDP/network connectivity.
# ------------------------------------------------------------

$WindowsValidationScript = @'

$ErrorActionPreference = "Continue"

$Results = [System.Collections.Generic.List[object]]::new()

function Add-Result {
    param(
        [string]$Test,
        [string]$Status,
        [string]$Details,
        [object]$Value = $null
    )

    $Results.Add(
        [PSCustomObject]@{
            Test    = $Test
            Status  = $Status
            Details = $Details
            Value   = $Value
        }
    )
}

# ------------------------------------------------------------
# OS
# ------------------------------------------------------------

try {

    $OS = Get-CimInstance Win32_OperatingSystem

    if ($OS) {

        Add-Result `
            -Test "Windows OS" `
            -Status "PASS" `
            -Details "Windows operating system detected." `
            -Value $OS.Caption

    }
    else {

        Add-Result `
            -Test "Windows OS" `
            -Status "FAIL" `
            -Details "Unable to query Windows operating system."

    }

}
catch {

    Add-Result `
        -Test "Windows OS" `
        -Status "FAIL" `
        -Details $_.Exception.Message
}

# ------------------------------------------------------------
# Computer name
# ------------------------------------------------------------

try {

    $ComputerName = $env:COMPUTERNAME

    Add-Result `
        -Test "Computer Name" `
        -Status "PASS" `
        -Details "Computer name successfully retrieved." `
        -Value $ComputerName

}
catch {

    Add-Result `
        -Test "Computer Name" `
        -Status "FAIL" `
        -Details $_.Exception.Message
}

# ------------------------------------------------------------
# Last boot
# ------------------------------------------------------------

try {

    $OS = Get-CimInstance Win32_OperatingSystem

    $BootTime = $OS.LastBootUpTime

    if ($BootTime) {

        Add-Result `
            -Test "Windows Boot" `
            -Status "PASS" `
            -Details "Windows successfully booted." `
            -Value $BootTime

    }
    else {

        Add-Result `
            -Test "Windows Boot" `
            -Status "FAIL" `
            -Details "Unable to determine Windows boot time."

    }

}
catch {

    Add-Result `
        -Test "Windows Boot" `
        -Status "FAIL" `
        -Details $_.Exception.Message
}

# ------------------------------------------------------------
# Volumes
# ------------------------------------------------------------

try {

    $Volumes = Get-Volume |
        Where-Object {
            $_.DriveLetter
        }

    foreach ($Volume in $Volumes) {

        if ($Volume.HealthStatus -eq "Healthy") {

            Add-Result `
                -Test "Volume $($Volume.DriveLetter):" `
                -Status "PASS" `
                -Details "Volume is healthy." `
                -Value "$($Volume.FileSystem) / $($Volume.SizeRemaining) free"

        }
        else {

            Add-Result `
                -Test "Volume $($Volume.DriveLetter):" `
                -Status "FAIL" `
                -Details "Volume health status: $($Volume.HealthStatus)"

        }

    }

}
catch {

    Add-Result `
        -Test "Disk / Volume Check" `
        -Status "FAIL" `
        -Details $_.Exception.Message
}

# ------------------------------------------------------------
# Windows system directory
# ------------------------------------------------------------

try {

    if (Test-Path "C:\Windows") {

        Add-Result `
            -Test "Windows File System" `
            -Status "PASS" `
            -Details "C:\Windows exists."

    }
    else {

        Add-Result `
            -Test "Windows File System" `
            -Status "FAIL" `
            -Details "C:\Windows does not exist."

    }

}
catch {

    Add-Result `
        -Test "Windows File System" `
        -Status "FAIL" `
        -Details $_.Exception.Message
}

# ------------------------------------------------------------
# Critical Windows services
# ------------------------------------------------------------

$CriticalServices = @(
    "EventLog",
    "RpcSs",
    "Winmgmt",
    "Schedule"
)

foreach ($ServiceName in $CriticalServices) {

    try {

        $Service = Get-Service -Name $ServiceName -ErrorAction Stop

        if ($Service.Status -eq "Running") {

            Add-Result `
                -Test "Service: $ServiceName" `
                -Status "PASS" `
                -Details "Service is running."

        }
        else {

            Add-Result `
                -Test "Service: $ServiceName" `
                -Status "FAIL" `
                -Details "Service state: $($Service.Status)"

        }

    }
    catch {

        Add-Result `
            -Test "Service: $ServiceName" `
            -Status "FAIL" `
            -Details $_.Exception.Message
    }
}

# ------------------------------------------------------------
# Windows Event Log
# ------------------------------------------------------------

try {

    $Events = Get-WinEvent `
        -FilterHashtable @{
            LogName = "System"
            Level = 2
            StartTime = (Get-Date).AddMinutes(-15)
        } `
        -MaxEvents 10 `
        -ErrorAction SilentlyContinue

    if ($Events.Count -eq 0) {

        Add-Result `
            -Test "Recent System Errors" `
            -Status "PASS" `
            -Details "No critical System events detected in the last 15 minutes."

    }
    else {

        Add-Result `
            -Test "Recent System Errors" `
            -Status "WARNING" `
            -Details "$($Events.Count) critical System events detected."

    }

}
catch {

    Add-Result `
        -Test "Recent System Errors" `
        -Status "WARNING" `
        -Details "Unable to query System event log."
}

# ------------------------------------------------------------
# Azure VM Agent
# ------------------------------------------------------------

try {

    $AgentService = Get-Service `
        -Name "WindowsAzureGuestAgent" `
        -ErrorAction SilentlyContinue

    if ($AgentService.Status -eq "Running") {

        Add-Result `
            -Test "Azure VM Agent" `
            -Status "PASS" `
            -Details "Azure VM Agent is running."

    }
    else {

        Add-Result `
            -Test "Azure VM Agent" `
            -Status "FAIL" `
            -Details "Azure VM Agent is not running."

    }

}
catch {

    Add-Result `
        -Test "Azure VM Agent" `
        -Status "WARNING" `
        -Details "Unable to determine Azure VM Agent state."
}

# ------------------------------------------------------------
# Return results
# ------------------------------------------------------------

$Results | ConvertTo-Json -Depth 5

'@

# ------------------------------------------------------------
# Main VM loop
# ------------------------------------------------------------

foreach ($VMName in $VMNames) {

    $VMStart = Get-Date
    $RestoredVMName = $null
    $RestoreJob = $null

    Write-Output ""
    Write-Output "============================================================"
    Write-Output "Processing VM: $VMName"
    Write-Output "============================================================"

    try {

        # ----------------------------------------------------
        # Find source VM
        # ----------------------------------------------------

        $SourceVM = Get-AzVM `
            -ResourceGroupName $SourceResourceGroup `
            -Name $VMName

        if (-not $SourceVM) {
            throw "Source VM not found."
        }

        Add-TestResult `
            -VMName $VMName `
            -Test "Source VM Discovery" `
            -Status "PASS" `
            -Details "Source VM found." `
            -Value $SourceVM.Id

        # ----------------------------------------------------
        # Find Recovery Services Vault
        # ----------------------------------------------------

        Write-Output "Finding Recovery Services vault..."

        $Vaults = Get-AzRecoveryServicesVault

        $BackupItem = $null
        $Vault = $null

        foreach ($CandidateVault in $Vaults) {

            try {

                $CandidateItem = Get-VMBackupItem `
                    -Vault $CandidateVault `
                    -VMName $VMName

                if ($CandidateItem) {

                    $BackupItem = $CandidateItem
                    $Vault = $CandidateVault
                    break
                }

            }
            catch {
                continue
            }
        }

        if (-not $BackupItem) {
            throw "No Azure Backup item found for VM $VMName."
        }

        Add-TestResult `
            -VMName $VMName `
            -Test "Backup Configuration" `
            -Status "PASS" `
            -Details "Backup item found." `
            -Value $Vault.Name

        # ----------------------------------------------------
        # Recovery point
        # ----------------------------------------------------

        $RP = Get-RecoveryPointForVM `
            -BackupItem $BackupItem `
            -Vault $Vault `
            -PreferredRecoveryPoint $RecoveryPoint `
            -PreferredRecoveryPointTime $RecoveryPointDateTime

        if (-not $RP) {
            throw "No recovery point found."
        }

        Write-Output "Selected recovery point:"
        Write-Output $RP.RecoveryPointTime

        Add-TestResult `
            -VMName $VMName `
            -Test "Recovery Point" `
            -Status "PASS" `
            -Details "Recovery point selected." `
            -Value $RP.RecoveryPointTime

        # ----------------------------------------------------
        # Generate test VM name
        # ----------------------------------------------------

        $RestoredVMName = "$VMName-RESTORETEST-$ExecutionId"

        Write-Output "Restored VM name: $RestoredVMName"

        # ----------------------------------------------------
        # Restore
        # ----------------------------------------------------

        Write-Output "Starting VM restore..."

        $RestoreJob = Restore-AzRecoveryServicesBackupItem `
            -RecoveryPoint $RP `
            -TargetResourceGroupName $TargetResourceGroup `
            -StorageAccountName $null `
            -StorageAccountResourceGroupName $null `
            -TargetVMName $RestoredVMName `
            -TargetVNetName $TargetVNetName `
            -TargetVNetResourceGroup $TargetVNetResourceGroup `
            -TargetSubnetName $TargetSubnetName `
            -VaultId $Vault.Id `
            -VaultLocation $Vault.Location

        Add-TestResult `
            -VMName $VMName `
            -Test "Restore Started" `
            -Status "PASS" `
            -Details "Azure Backup restore job started." `
            -Value $RestoreJob.JobId

        # ----------------------------------------------------
        # Wait for restore
        # ----------------------------------------------------

        $CompletedJob = Wait-ForRestoreJob `
            -Job $RestoreJob `
            -Vault $Vault `
            -TimeoutMinutes $RestoreTimeoutMinutes

        if ($CompletedJob.Status -ne "Completed") {

            Add-TestResult `
                -VMName $VMName `
                -Test "VM Restore" `
                -Status "FAIL" `
                -Details "Restore job ended with status $($CompletedJob.Status)." `
                -Value $CompletedJob.JobId

            throw "Restore failed."
        }

        Add-TestResult `
            -VMName $VMName `
            -Test "VM Restore" `
            -Status "PASS" `
            -Details "Restore completed successfully." `
            -Value $CompletedJob.JobId

        # ----------------------------------------------------
        # Verify VM exists
        # ----------------------------------------------------

        $RestoredVM = Get-AzVM `
            -ResourceGroupName $TargetResourceGroup `
            -Name $RestoredVMName `
            -ErrorAction Stop

        Add-TestResult `
            -VMName $VMName `
            -Test "Restored VM Created" `
            -Status "PASS" `
            -Details "Restored VM exists." `
            -Value $RestoredVM.Id

        # ----------------------------------------------------
        # Start VM
        # ----------------------------------------------------

        Write-Output "Starting restored VM..."

        Start-AzVM `
            -ResourceGroupName $TargetResourceGroup `
            -Name $RestoredVMName `
            -NoWait

        # ----------------------------------------------------
        # Wait for VM
        # ----------------------------------------------------

        $VMReady = Wait-ForVMReady `
            -ResourceGroup $TargetResourceGroup `
            -VMName $RestoredVMName `
            -TimeoutMinutes $VMStartupTimeoutMinutes

        if (-not $VMReady) {

            Add-TestResult `
                -VMName $VMName `
                -Test "VM Startup" `
                -Status "FAIL" `
                -Details "Restored VM did not reach running state within timeout."

            throw "Restored VM failed to start."
        }

        Add-TestResult `
            -VMName $VMName `
            -Test "VM Startup" `
            -Status "PASS" `
            -Details "Restored VM reached running state."

        # ----------------------------------------------------
        # Azure VM Run Command
        # ----------------------------------------------------

        Write-Output "Running Windows integrity tests..."

        $RunCommandResult = Invoke-AzVMRunCommand `
            -ResourceGroupName $TargetResourceGroup `
            -VMName $RestoredVMName `
            -CommandId "RunPowerShellScript" `
            -ScriptString $WindowsValidationScript

        $RawOutput = (
            $RunCommandResult.Value |
            Where-Object {
                $_.Code -eq "ComponentStatus/StdOut/succeeded"
            }
        ).Message

        if (-not $RawOutput) {
            throw "No output returned from Windows validation."
        }

        Write-Output "Windows validation completed."

        try {

            $WindowsResults = $RawOutput | ConvertFrom-Json

            foreach ($TestResult in $WindowsResults) {

                Add-TestResult `
                    -VMName $VMName `
                    -Test $TestResult.Test `
                    -Status $TestResult.Status `
                    -Details $TestResult.Details `
                    -Value $TestResult.Value
            }

        }
        catch {

            Add-TestResult `
                -VMName $VMName `
                -Test "Windows Validation Parsing" `
                -Status "FAIL" `
                -Details "Unable to parse Run Command output."

        }

    }
    catch {

        Write-Error "VM $VMName failed: $($_.Exception.Message)"

        Add-TestResult `
            -VMName $VMName `
            -Test "Overall VM Test" `
            -Status "FAIL" `
            -Details $_.Exception.Message
    }
    finally {

        # ----------------------------------------------------
        # Cleanup
        # ----------------------------------------------------

        if (
            $Cleanup.IsPresent -and
            $RestoredVMName
        ) {

            Write-Output "Cleaning up restored VM: $RestoredVMName"

            try {

                Remove-AzVM `
                    -ResourceGroupName $TargetResourceGroup `
                    -Name $RestoredVMName `
                    -Force

                Add-TestResult `
                    -VMName $VMName `
                    -Test "Cleanup" `
                    -Status "PASS" `
                    -Details "Restored VM deleted."

            }
            catch {

                Add-TestResult `
                    -VMName $VMName `
                    -Test "Cleanup" `
                    -Status "FAIL" `
                    -Details $_.Exception.Message
            }
        }
        elseif ($RestoredVMName) {

            Add-TestResult `
                -VMName $VMName `
                -Test "Cleanup" `
                -Status "SKIPPED" `
                -Details "Cleanup disabled. Restored VM retained." `
                -Value $RestoredVMName
        }
    }
}

# ============================================================
# Summary
# ============================================================

$ExecutionEnd = Get-Date

$PassCount = @(
    $Results | Where-Object Status -eq "PASS"
).Count

$FailCount = @(
    $Results | Where-Object Status -eq "FAIL"
).Count

$WarningCount = @(
    $Results | Where-Object Status -eq "WARNING"
).Count

$OverallStatus = if ($FailCount -eq 0) {
    "PASS"
}
else {
    "FAIL"
}

$Report = [PSCustomObject]@{
    ExecutionId   = $ExecutionId
    Started       = $ExecutionStart
    Completed     = $ExecutionEnd
    Duration      = ($ExecutionEnd - $ExecutionStart).ToString()
    OverallStatus = $OverallStatus
    VMCount       = $VMNames.Count
    TestsPassed   = $PassCount
    TestsFailed   = $FailCount
    Warnings      = $WarningCount
    Results       = $Results
}

# ============================================================
# JSON report
# ============================================================

$JsonReport = $Report | ConvertTo-Json -Depth 10

Write-Output ""
Write-Output "============================================================"
Write-Output "RESTORE DRILL SUMMARY"
Write-Output "============================================================"

Write-Output "Execution ID : $ExecutionId"
Write-Output "Status       : $OverallStatus"
Write-Output "VMs tested   : $($VMNames.Count)"
Write-Output "Tests passed : $PassCount"
Write-Output "Tests failed : $FailCount"
Write-Output "Warnings     : $WarningCount"
Write-Output "Duration     : $($ExecutionEnd - $ExecutionStart)"

Write-Output ""
Write-Output "JSON REPORT"
Write-Output $JsonReport

# ============================================================
# HTML report
# ============================================================

if ($GenerateHtmlReport.IsPresent) {

    $Rows = foreach ($Result in $Results) {

        $StatusClass = switch ($Result.Status) {
            "PASS"    { "pass" }
            "FAIL"    { "fail" }
            "WARNING" { "warning" }
            default   { "other" }
        }

        @"
<tr>
<td>$($Result.VM)</td>
<td>$($Result.Test)</td>
<td class="$StatusClass">$($Result.Status)</td>
<td>$($Result.Details)</td>
<td>$($Result.Value)</td>
</tr>
"@
    }

    $HtmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Azure VM Restore Integrity Drill</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }
        th { background: #f2f2f2; }
        .pass { color: green; font-weight: bold; }
        .fail { color: red; font-weight: bold; }
        .warning { color: orange; font-weight: bold; }
        .other { color: gray; }
    </style>
</head>
<body>
    <h1>Azure VM Restore Integrity Drill</h1>
    <p><strong>Execution ID:</strong> $ExecutionId</p>
    <p><strong>Status:</strong> $OverallStatus</p>
    <p><strong>VMs tested:</strong> $($VMNames.Count)</p>
    <p><strong>Tests passed:</strong> $PassCount</p>
    <p><strong>Tests failed:</strong> $FailCount</p>
    <p><strong>Warnings:</strong> $WarningCount</p>
    <p><strong>Duration:</strong> $($ExecutionEnd - $ExecutionStart)</p>

    <table>
        <thead>
            <tr>
                <th>VM</th>
                <th>Test</th>
                <th>Status</th>
                <th>Details</th>
                <th>Value</th>
            </tr>
        </thead>
        <tbody>
            $Rows
        </tbody>
    </table>
</body>
</html>
"@

    $HtmlPath = Join-Path -Path (Get-Location) -ChildPath "vm-restore-integrity-$ExecutionId.html"
    $HtmlReport | Set-Content -Path $HtmlPath -Encoding UTF8
    Write-Output "HTML report written to: $HtmlPath"
}

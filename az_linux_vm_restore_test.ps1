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
# Restore-VM-Integrity-Drill (Linux)
#
# PowerShell 7.4, run from a local workstation.
# Restores Azure Backup VMs, starts them, and runs Linux integrity
# checks using Azure VM Run Command (RunShellScript).
# ============================================================

$ErrorActionPreference = "Stop"

function Assert-AzModule {
    param([string[]]$ModuleNames)

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
        $ConnectParameters = @{}
        if ($SubscriptionId) { $ConnectParameters.Subscription = $SubscriptionId }
        if ($TenantId) { $ConnectParameters.Tenant = $TenantId }
        $Context = Connect-AzAccount @ConnectParameters
    }
    elseif ($SubscriptionId) {
        $SetContextParameters = @{ SubscriptionId = $SubscriptionId; ErrorAction = 'Stop' }
        if ($TenantId) { $SetContextParameters.Tenant = $TenantId }
        Set-AzContext @SetContextParameters | Out-Null
        $Context = Get-AzContext
    }

    if (-not $Context) {
        throw "Unable to authenticate to Azure from this workstation."
    }

    return $Context
}

function Add-TestResult {
    param(
        [string]$VMName,
        [string]$Test,
        [string]$Status,
        [string]$Details,
        [object]$Value = $null
    )

    $Results.Add([PSCustomObject]@{
        VM        = $VMName
        Test      = $Test
        Status    = $Status
        Details   = $Details
        Value     = $Value
        Timestamp = (Get-Date).ToString('s')
    })
}

function Get-VMBackupItem {
    param([object]$Vault, [string]$VMName)

    Write-Output "Searching backup item for VM: $VMName"
    $Items = Get-AzRecoveryServicesBackupItem `
        -BackupManagementType AzureVM `
        -WorkloadType AzureVM `
        -VaultId $Vault.Id

    return $Items |
        Where-Object { $_.FriendlyName -eq $VMName -or $_.Name -eq $VMName } |
        Select-Object -First 1
}

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

    if (-not $RecoveryPoints) { throw "No recovery points found in the last 30 days." }

    switch ($PreferredRecoveryPoint.ToLowerInvariant()) {
        'latest' {
            return $RecoveryPoints | Sort-Object RecoveryPointTime -Descending | Select-Object -First 1
        }
        'earliest' {
            return $RecoveryPoints | Sort-Object RecoveryPointTime | Select-Object -First 1
        }
        default {
            if ($PSBoundParameters.ContainsKey('PreferredRecoveryPointTime') -and $PreferredRecoveryPointTime) {
                $TargetDate = $PreferredRecoveryPointTime
            }
            else {
                try { $TargetDate = [datetime]::Parse($PreferredRecoveryPoint) }
                catch { throw "RecoveryPoint '$PreferredRecoveryPoint' is invalid. Use Latest, Earliest, or a datetime." }
            }

            $ChosenPoint = $RecoveryPoints |
                Where-Object { $_.RecoveryPointTime -le $TargetDate } |
                Sort-Object RecoveryPointTime -Descending |
                Select-Object -First 1

            if (-not $ChosenPoint) { throw "No recovery point found on or before '$TargetDate'." }
            return $ChosenPoint
        }
    }
}

function Wait-ForRestoreJob {
    param([object]$Job, [object]$Vault, [int]$TimeoutMinutes)

    $Start = Get-Date
    do {
        Start-Sleep -Seconds 30
        $CurrentJob = Get-AzRecoveryServicesBackupJob -VaultId $Vault.Id -JobId $Job.JobId
        Write-Output ("Restore job status: {0}" -f $CurrentJob.Status)

        if ($CurrentJob.Status -in @('Completed', 'Failed', 'Cancelled', 'Canceled')) {
            return $CurrentJob
        }
        if (((Get-Date) - $Start).TotalMinutes -gt $TimeoutMinutes) {
            throw "Restore operation exceeded timeout of $TimeoutMinutes minutes."
        }
    } while ($true)
}

function Wait-ForVMReady {
    param([string]$ResourceGroup, [string]$VMName, [int]$TimeoutMinutes)

    $Start = Get-Date
    do {
        Start-Sleep -Seconds 20
        try {
            $Status = Get-AzVM -ResourceGroupName $ResourceGroup -Name $VMName -Status
            $PowerState = ($Status.Statuses | Where-Object { $_.Code -like 'PowerState/*' }).DisplayStatus
            $ProvisioningState = ($Status.Statuses | Where-Object { $_.Code -like 'ProvisioningState/*' }).DisplayStatus
            Write-Output "VM state: $PowerState / $ProvisioningState"

            if ($PowerState -eq 'VM running' -and $ProvisioningState -eq 'Provisioning succeeded') {
                return $true
            }
        }
        catch {
            Write-Output "VM not ready yet: $($_.Exception.Message)"
        }

        if (((Get-Date) - $Start).TotalMinutes -gt $TimeoutMinutes) { return $false }
    } while ($true)
}

$ErrorActionPreference = 'Stop'
$ExecutionId = Get-Date -Format 'yyyyMMdd-HHmmss'
$ExecutionStart = Get-Date
$Results = [System.Collections.Generic.List[object]]::new()

Write-Output '============================================================'
Write-Output ' Azure Linux VM Restore Integrity Drill'
Write-Output " Execution ID: $ExecutionId"
Write-Output ' Run mode: Local workstation'
Write-Output '============================================================'

Write-Output 'Validating Azure PowerShell modules...'
Assert-AzModule -ModuleNames @('Az.Accounts', 'Az.Compute', 'Az.Network', 'Az.RecoveryServices', 'Az.Resources')
Write-Output 'Authenticating to Azure using your local account...'
$AzureContext = Connect-LocalAzureContext -SubscriptionId $SubscriptionId -TenantId $TenantId
Write-Output "Authenticated as: $($AzureContext.Account.Id)"
if ($AzureContext.Subscription) {
    Write-Output "Subscription: $($AzureContext.Subscription.Name) ($($AzureContext.Subscription.Id))"
}

# This script is sent to the restored Linux guest by Azure VM Run Command.
$LinuxValidationScript = @'
#!/bin/bash
set +e

RESULTS=""
add_result() {
    local test="$1" status="$2" details="$3" value="$4"
    local item
    item=$(TEST_NAME="$test" TEST_STATUS="$status" TEST_DETAILS="$details" TEST_VALUE="$value" python3 - <<'PY'
import json, os
print(json.dumps({
    "Test": os.environ.get("TEST_NAME", ""),
    "Status": os.environ.get("TEST_STATUS", ""),
    "Details": os.environ.get("TEST_DETAILS", ""),
    "Value": os.environ.get("TEST_VALUE", "") or None
}))
PY
)
    if [ -n "$RESULTS" ]; then RESULTS="$RESULTS,$item"; else RESULTS="$item"; fi
}

# OS identity
if [ -r /etc/os-release ]; then
    . /etc/os-release
    add_result "Linux OS" "PASS" "Linux operating system detected." "${PRETTY_NAME:-${NAME:-Linux}}"
else
    add_result "Linux OS" "FAIL" "Unable to read /etc/os-release." ""
fi

# Hostname
HOSTNAME_VALUE=$(hostname 2>/dev/null)
if [ -n "$HOSTNAME_VALUE" ]; then
    add_result "Hostname" "PASS" "Hostname successfully retrieved." "$HOSTNAME_VALUE"
else
    add_result "Hostname" "FAIL" "Unable to retrieve hostname." ""
fi

# Kernel and boot time
KERNEL=$(uname -r 2>/dev/null)
if [ -n "$KERNEL" ]; then
    add_result "Linux Kernel" "PASS" "Kernel version successfully retrieved." "$KERNEL"
else
    add_result "Linux Kernel" "FAIL" "Unable to retrieve kernel version." ""
fi
if [ -r /proc/uptime ]; then
    UPTIME_SECONDS=$(awk '{print int($1)}' /proc/uptime 2>/dev/null)
    add_result "Linux Boot" "PASS" "System is booted; uptime retrieved." "${UPTIME_SECONDS} seconds"
else
    add_result "Linux Boot" "FAIL" "Unable to read /proc/uptime." ""
fi

# Root filesystem and mounted filesystems
if command -v df >/dev/null 2>&1 && df -P / >/dev/null 2>&1; then
    ROOT_USAGE=$(df -hP / | tail -n 1 | awk '{print "Filesystem=" $1 ", Size=" $2 ", Used=" $3 ", Available=" $4 ", Use=" $5}')
    add_result "Root Filesystem" "PASS" "Root filesystem is mounted and readable." "$ROOT_USAGE"
else
    add_result "Root Filesystem" "FAIL" "Root filesystem check failed." ""
fi

if command -v findmnt >/dev/null 2>&1; then
    MOUNT_COUNT=$(findmnt -rn -o TARGET 2>/dev/null | wc -l | tr -d ' ')
    if [ "${MOUNT_COUNT:-0}" -gt 0 ]; then
        add_result "Mounted Filesystems" "PASS" "Mounted filesystem table is available." "$MOUNT_COUNT mounts"
    else
        add_result "Mounted Filesystems" "WARNING" "No mounted filesystems were reported." ""
    fi
else
    add_result "Mounted Filesystems" "WARNING" "findmnt is not installed; mounted filesystems were not enumerated." ""
fi

# Essential paths
for path in / /etc /var /usr; do
    if [ -d "$path" ] && [ -r "$path" ]; then
        add_result "Filesystem Path: $path" "PASS" "Path exists and is readable." ""
    else
        add_result "Filesystem Path: $path" "FAIL" "Path is missing or unreadable." ""
    fi
done

# Essential services, using systemd where present
if command -v systemctl >/dev/null 2>&1; then
    for service in sshd ssh waagent; do
        if systemctl list-unit-files "${service}.service" --no-legend 2>/dev/null | grep -q "${service}.service"; then
            SERVICE_STATE=$(systemctl is-active "$service" 2>/dev/null)
            if [ "$SERVICE_STATE" = "active" ]; then
                add_result "Service: $service" "PASS" "Service is active." "$SERVICE_STATE"
            else
                add_result "Service: $service" "WARNING" "Service is installed but not active." "${SERVICE_STATE:-unknown}"
            fi
        fi
    done
else
    add_result "Service Manager" "WARNING" "systemctl is unavailable; service checks were skipped." ""
fi

# Azure Linux Agent, if installed
if command -v waagent >/dev/null 2>&1 || systemctl list-unit-files 'waagent.service' --no-legend 2>/dev/null | grep -q 'waagent.service'; then
    if systemctl is-active --quiet waagent 2>/dev/null; then
        add_result "Azure Linux Agent" "PASS" "Azure Linux Agent service is active." "waagent"
    elif pgrep -x waagent >/dev/null 2>&1; then
        add_result "Azure Linux Agent" "PASS" "Azure Linux Agent process is running." "waagent"
    else
        add_result "Azure Linux Agent" "WARNING" "Azure Linux Agent is installed but does not appear active." ""
    fi
else
    add_result "Azure Linux Agent" "WARNING" "Azure Linux Agent (WALinuxAgent) was not detected." ""
fi

# Recent critical kernel/system messages, where journalctl is available
if command -v journalctl >/dev/null 2>&1; then
    ERROR_COUNT=$(journalctl -p err --since '15 minutes ago' -n 100 --no-pager 2>/dev/null | grep -cve '^-- No entries --$' || true)
    if [ "${ERROR_COUNT:-0}" -eq 0 ]; then
        add_result "Recent System Errors" "PASS" "No recent error-priority journal entries were detected." "0"
    else
        add_result "Recent System Errors" "WARNING" "Recent error-priority journal entries were found." "$ERROR_COUNT entries (capped at 100)"
    fi
else
    add_result "Recent System Errors" "WARNING" "journalctl is unavailable; recent system log check was skipped." ""
fi

printf '{"results":[%s]}\n' "$RESULTS"
'@

foreach ($VMName in $VMNames) {
    $RestoredVMName = $null

    Write-Output ''
    Write-Output '============================================================'
    Write-Output "Processing Linux VM: $VMName"
    Write-Output '============================================================'

    try {
        $SourceVM = Get-AzVM -ResourceGroupName $SourceResourceGroup -Name $VMName -ErrorAction Stop
        if (-not $SourceVM) { throw 'Source VM not found.' }
        if ($SourceVM.StorageProfile.OsDisk.OsType -ne 'Linux') {
            throw "VM '$VMName' is reported as '$($SourceVM.StorageProfile.OsDisk.OsType)', not Linux."
        }

        Add-TestResult -VMName $VMName -Test 'Source VM Discovery' -Status 'PASS' -Details 'Source Linux VM found.' -Value $SourceVM.Id

        Write-Output 'Finding Recovery Services vault...'
        $Vaults = Get-AzRecoveryServicesVault
        $BackupItem = $null
        $Vault = $null

        foreach ($CandidateVault in $Vaults) {
            try {
                $CandidateItem = Get-VMBackupItem -Vault $CandidateVault -VMName $VMName
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

        if (-not $BackupItem) { throw "No Azure Backup item found for VM '$VMName'." }
        Add-TestResult -VMName $VMName -Test 'Backup Configuration' -Status 'PASS' -Details 'Backup item found.' -Value $Vault.Name

        $RP = Get-RecoveryPointForVM `
            -BackupItem $BackupItem `
            -Vault $Vault `
            -PreferredRecoveryPoint $RecoveryPoint `
            -PreferredRecoveryPointTime $RecoveryPointDateTime
        if (-not $RP) { throw 'No recovery point found.' }

        Write-Output "Selected recovery point: $($RP.RecoveryPointTime)"
        Add-TestResult -VMName $VMName -Test 'Recovery Point' -Status 'PASS' -Details 'Recovery point selected.' -Value $RP.RecoveryPointTime

        $RestoredVMName = "$VMName-RESTORETEST-$ExecutionId"
        Write-Output "Restored VM name: $RestoredVMName"
        Write-Output 'Starting VM restore...'

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

        Add-TestResult -VMName $VMName -Test 'Restore Started' -Status 'PASS' -Details 'Azure Backup restore job started.' -Value $RestoreJob.JobId

        $CompletedJob = Wait-ForRestoreJob -Job $RestoreJob -Vault $Vault -TimeoutMinutes $RestoreTimeoutMinutes
        if ($CompletedJob.Status -ne 'Completed') {
            Add-TestResult -VMName $VMName -Test 'VM Restore' -Status 'FAIL' -Details "Restore job ended with status $($CompletedJob.Status)." -Value $CompletedJob.JobId
            throw 'Restore failed.'
        }
        Add-TestResult -VMName $VMName -Test 'VM Restore' -Status 'PASS' -Details 'Restore completed successfully.' -Value $CompletedJob.JobId

        $RestoredVM = Get-AzVM -ResourceGroupName $TargetResourceGroup -Name $RestoredVMName -ErrorAction Stop
        Add-TestResult -VMName $VMName -Test 'Restored VM Created' -Status 'PASS' -Details 'Restored VM exists.' -Value $RestoredVM.Id

        Write-Output 'Starting restored VM...'
        Start-AzVM -ResourceGroupName $TargetResourceGroup -Name $RestoredVMName -NoWait | Out-Null

        $VMReady = Wait-ForVMReady -ResourceGroup $TargetResourceGroup -VMName $RestoredVMName -TimeoutMinutes $VMStartupTimeoutMinutes
        if (-not $VMReady) {
            Add-TestResult -VMName $VMName -Test 'VM Startup' -Status 'FAIL' -Details 'Restored VM did not reach running state within timeout.'
            throw 'Restored VM failed to start.'
        }
        Add-TestResult -VMName $VMName -Test 'VM Startup' -Status 'PASS' -Details 'Restored VM reached running state.'

        Write-Output 'Running Linux integrity tests...'
        $RunCommandResult = Invoke-AzVMRunCommand `
            -ResourceGroupName $TargetResourceGroup `
            -VMName $RestoredVMName `
            -CommandId 'RunShellScript' `
            -ScriptString $LinuxValidationScript

        $RawOutput = ($RunCommandResult.Value | Where-Object { $_.Code -eq 'ComponentStatus/StdOut/succeeded' }).Message
        if (-not $RawOutput) { throw 'No output returned from Linux validation.' }

        try {
            $ParsedOutput = $RawOutput | ConvertFrom-Json -ErrorAction Stop
            $LinuxResults = @($ParsedOutput.results)
            if ($LinuxResults.Count -eq 0) { throw 'Validation returned no test results.' }

            foreach ($TestResult in $LinuxResults) {
                Add-TestResult `
                    -VMName $VMName `
                    -Test $TestResult.Test `
                    -Status $TestResult.Status `
                    -Details $TestResult.Details `
                    -Value $TestResult.Value
            }
        }
        catch {
            Add-TestResult -VMName $VMName -Test 'Linux Validation Parsing' -Status 'FAIL' -Details "Unable to parse Run Command output: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Error "VM $VMName failed: $($_.Exception.Message)"
        Add-TestResult -VMName $VMName -Test 'Overall VM Test' -Status 'FAIL' -Details $_.Exception.Message
    }
    finally {
        if ($Cleanup.IsPresent -and $RestoredVMName) {
            Write-Output "Cleaning up restored VM: $RestoredVMName"
            try {
                # Remove the VM object only. Azure-managed NICs and disks may remain;
                # review the target resource group and remove leftovers deliberately.
                Remove-AzVM -ResourceGroupName $TargetResourceGroup -Name $RestoredVMName -Force -ErrorAction Stop
                Add-TestResult -VMName $VMName -Test 'Cleanup' -Status 'PASS' -Details 'Restored VM object deleted. Review the target resource group for attached disks/NICs.'
            }
            catch {
                Add-TestResult -VMName $VMName -Test 'Cleanup' -Status 'FAIL' -Details $_.Exception.Message
            }
        }
        elseif ($RestoredVMName) {
            Add-TestResult -VMName $VMName -Test 'Cleanup' -Status 'SKIPPED' -Details 'Cleanup disabled. Restored VM retained.' -Value $RestoredVMName
        }
    }
}

$ExecutionEnd = Get-Date
$PassCount = @($Results | Where-Object Status -eq 'PASS').Count
$FailCount = @($Results | Where-Object Status -eq 'FAIL').Count
$WarningCount = @($Results | Where-Object Status -eq 'WARNING').Count
$OverallStatus = if ($FailCount -eq 0) { 'PASS' } else { 'FAIL' }

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

$JsonReport = $Report | ConvertTo-Json -Depth 10
Write-Output ''
Write-Output '============================================================'
Write-Output 'RESTORE DRILL SUMMARY'
Write-Output '============================================================'
Write-Output "Execution ID : $ExecutionId"
Write-Output "Status       : $OverallStatus"
Write-Output "VMs tested   : $($VMNames.Count)"
Write-Output "Tests passed : $PassCount"
Write-Output "Tests failed : $FailCount"
Write-Output "Warnings     : $WarningCount"
Write-Output "Duration     : $($ExecutionEnd - $ExecutionStart)"
Write-Output ''
Write-Output 'JSON REPORT'
Write-Output $JsonReport

if ($GenerateHtmlReport.IsPresent) {
    $Rows = foreach ($Result in $Results) {
        $StatusClass = switch ($Result.Status) {
            'PASS' { 'pass' }
            'FAIL' { 'fail' }
            'WARNING' { 'warning' }
            default { 'other' }
        }

        $Cells = @($Result.VM, $Result.Test, $Result.Status, $Result.Details, $Result.Value) |
            ForEach-Object { [System.Net.WebUtility]::HtmlEncode([string]$_) }

        "<tr><td>$($Cells[0])</td><td>$($Cells[1])</td><td class=`"$StatusClass`">$($Cells[2])</td><td>$($Cells[3])</td><td>$($Cells[4])</td></tr>"
    }

    $HtmlReport = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>Azure Linux VM Restore Integrity Drill</title>
<style>
body { font-family: Arial, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }
th { background: #f2f2f2; }
.pass { color: green; font-weight: bold; }
.fail { color: red; font-weight: bold; }
.warning { color: darkorange; font-weight: bold; }
.other { color: gray; }
</style>
</head>
<body>
<h1>Azure Linux VM Restore Integrity Drill</h1>
<p><strong>Execution ID:</strong> $ExecutionId</p>
<p><strong>Status:</strong> $OverallStatus</p>
<p><strong>VMs tested:</strong> $($VMNames.Count)</p>
<p><strong>Tests passed:</strong> $PassCount</p>
<p><strong>Tests failed:</strong> $FailCount</p>
<p><strong>Warnings:</strong> $WarningCount</p>
<p><strong>Duration:</strong> $($ExecutionEnd - $ExecutionStart)</p>
<table>
<thead><tr><th>VM</th><th>Test</th><th>Status</th><th>Details</th><th>Value</th></tr></thead>
<tbody>
$($Rows -join "`n")
</tbody>
</table>
</body>
</html>
"@

    $HtmlPath = Join-Path -Path (Get-Location) -ChildPath "linux-vm-restore-integrity-$ExecutionId.html"
    $HtmlReport | Set-Content -Path $HtmlPath -Encoding UTF8
    Write-Output "HTML report written to: $HtmlPath"
}

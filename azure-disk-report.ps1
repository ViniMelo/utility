param(
    [Parameter(Mandatory = $true)]
    [string]$DiskName,

    [int]$DaysBack = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-AzSession {
    try {
        $null = Get-AzContext -ErrorAction Stop
        return
    }
    catch {
        $azAccountsModule = Get-Module -ListAvailable -Name Az.Accounts

        if (-not $azAccountsModule) {
            Write-Warning "Azure PowerShell module 'Az.Accounts' is not installed."

            $installChoice = Read-Host "Would you like to install the Az.Accounts module from PSGallery now? [Y/N]"
            if ($installChoice -match '^(Y|YES)$') {
                try {
                    Install-Module -Name Az.Accounts -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
                }
                catch {
                    throw "Unable to install Az.Accounts: $($_.Exception.Message)"
                }
            }
            else {
                throw "Azure PowerShell is not installed. Please install Az.Accounts and rerun the script."
            }
        }

        try {
            Import-Module Az.Accounts -ErrorAction Stop
        }
        catch {
            throw "Unable to import Az.Accounts: $($_.Exception.Message)"
        }

        $loginChoice = Read-Host "Azure session is not logged in. Would you like to attempt a login now? [Y/N]"
        if ($loginChoice -match '^(Y|YES)$') {
            try {
                Connect-AzAccount -UseDeviceAuthentication
            }
            catch {
                throw "Azure login attempt failed: $($_.Exception.Message)"
            }
        }
        else {
            throw "Azure session is not logged in. Run Connect-AzAccount and try again."
        }
    }
}

function Get-LocalizedValue {
    param(
        [Parameter(Mandatory = $false)]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return $Value
    }

    if ($Value.PSObject.Properties.Name -contains 'LocalizedValue') {
        return $Value.LocalizedValue
    }

    return $Value
}

Ensure-AzSession

$StartDate = (Get-Date).AddDays(-$DaysBack)

try {
    Write-Host "Searching for disk '$DiskName'..." -ForegroundColor Cyan

    $matchingDisks = @(Get-AzDisk | Where-Object { $_.Name -ieq $DiskName })

    if ($matchingDisks.Count -eq 0) {
        throw "Disk '$DiskName' was not found."
    }

    if ($matchingDisks.Count -gt 1) {
        $diskList = $matchingDisks |
            Select-Object Name, ResourceGroupName, Location, @{N='DiskId';E={$_.Id}} |
            Format-Table -AutoSize |
            Out-String

        throw "Multiple disks matched '$DiskName'. Please use a unique disk name.`n$diskList"
    }

    $Disk = $matchingDisks[0]

    Write-Host "Disk found." -ForegroundColor Green

    try {
        $Logs = @(Get-AzLog `
            -ResourceId $Disk.Id `
            -StartTime $StartDate `
            -DetailedOutput `
            -ErrorAction Stop)
    }
    catch {
        Write-Warning "Failed to retrieve activity logs for disk '$DiskName'."
        $Logs = @()
    }

    $CurrentAttachment = if ($Disk.ManagedBy) {
        $Disk.ManagedBy
    }
    else {
        "UNATTACHED"
    }

    $LastVM = "Unknown"
    $LastAttachDate = $null
    $LastAttachUser = $null
    $LastAttachStatus = $null

    if (-not $Disk.ManagedBy) {
        Write-Host "Disk is currently orphaned/unattached. Attempting to identify previous usage..." -ForegroundColor Yellow

        $LastRelevantEvent = $Logs |
            Sort-Object EventTimestamp -Descending |
            Where-Object {
                $resourceId = $_.ResourceId
                $opName = Get-LocalizedValue $_.OperationName

                $resourceMatchesCompute = ($resourceId -match "Microsoft.Compute")
                $operationMatchesDisk = ($opName -match "disk")
                $operationMatchesVm = ($opName -match "virtual machine")

                $resourceMatchesCompute -or $operationMatchesDisk -or $operationMatchesVm
            } |
            Select-Object -First 1

        if ($LastRelevantEvent) {
            $LastAttachDate = $LastRelevantEvent.EventTimestamp
            $LastAttachUser = $LastRelevantEvent.Caller
            $LastAttachStatus = Get-LocalizedValue $LastRelevantEvent.Status

            if ($LastRelevantEvent.ResourceId) {
                $LastVM = $LastRelevantEvent.ResourceId
            }
        }
    }
    else {
        $LastVM = $Disk.ManagedBy
    }

    $CreationEvent = $Logs |
        Sort-Object EventTimestamp |
        Where-Object {
            $opName = Get-LocalizedValue $_.OperationName
            $opName -match "create"
        } |
        Select-Object -First 1

    $FailedEvents = @($Logs | Where-Object {
            $status = Get-LocalizedValue $_.Status
            $status -and $status -ne "Succeeded"
        })

    $Summary = [PSCustomObject]@{
        DiskName                = $Disk.Name
        ResourceGroup           = $Disk.ResourceGroupName
        SubscriptionId          = ($Disk.Id -split "/")[2]

        Location                = $Disk.Location
        AvailabilityZone        = ($Disk.Zones -join ",")

        DiskState               = $Disk.DiskState
        ManagedBy               = $CurrentAttachment

        OSType                  = $Disk.OsType
        DiskSizeGB              = $Disk.DiskSizeGB

        StorageType             = if ($Disk.Sku) { $Disk.Sku.Name } else { $null }
        SecurityType            = if ($Disk.SecurityProfile) { $Disk.SecurityProfile.SecurityType } else { $null }

        TimeCreated             = $Disk.TimeCreated

        LastKnownVM             = $LastVM
        LastKnownActivity       = $LastAttachDate
        LastKnownActivityBy     = $LastAttachUser
        LastKnownActivityStatus = $LastAttachStatus

        CreatedBy               = if ($CreationEvent) { $CreationEvent.Caller } else { $null }
        CreateTime              = if ($CreationEvent) { $CreationEvent.EventTimestamp } else { $null }

        AuditEventsFound        = $Logs.Count
        FailedEventsFound       = $FailedEvents.Count

        Tags = if ($Disk.Tags) {
            ($Disk.Tags.GetEnumerator() |
                ForEach-Object {
                    "$($_.Key)=$($_.Value)"
                }) -join "; "
        }
        else {
            ""
        }
    }

    Write-Host ""
    Write-Host "=== DISK SUMMARY ===" -ForegroundColor Green
    $Summary | Format-List

    $AuditExport = $Logs | Select-Object `
        EventTimestamp,
        Caller,
        CorrelationId,
        ResourceId,
        @{N="Operation";E={Get-LocalizedValue $_.OperationName}},
        @{N="Status";E={Get-LocalizedValue $_.Status}},
        @{N="SubStatus";E={Get-LocalizedValue $_.SubStatus}},
        @{N="Category";E={Get-LocalizedValue $_.Category}}

    $OutputFile = "C:\Temp\$($Disk.Name)_${DaysBack}Day_Audit.csv"

    $AuditExport |
        Sort-Object EventTimestamp -Descending |
        Export-Csv $OutputFile -NoTypeInformation

    Write-Host ""
    Write-Host "Audit events exported to:" -ForegroundColor Green
    Write-Host $OutputFile

    if ($FailedEvents.Count -gt 0) {
        Write-Host ""
        Write-Host "=== FAILED OPERATIONS DETECTED ===" -ForegroundColor Yellow

        $FailedEvents |
            Select-Object `
                EventTimestamp,
                Caller,
                @{N='Operation';E={Get-LocalizedValue $_.OperationName}},
                @{N='Status';E={Get-LocalizedValue $_.Status}} |
            Format-Table -AutoSize
    }
    else {
        Write-Host ""
        Write-Host "No failed operations found in the last $DaysBack days." -ForegroundColor Green
    }
}
catch {
    Write-Error $_.Exception.Message
}

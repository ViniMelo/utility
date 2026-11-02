param(
    [Parameter(Mandatory = $true)]
    [string]$DiskName,

    [int]$DaysBack = 30
)

$StartDate = (Get-Date).AddDays(-$DaysBack)

try {

    Write-Host "Searching for disk '$DiskName'..." -ForegroundColor Cyan

    $Disk = Get-AzDisk | Where-Object Name -eq $DiskName

    if (!$Disk) {
        throw "Disk '$DiskName' was not found."
    }

    Write-Host "Disk found." -ForegroundColor Green

    # Get activity logs
    try {
        $Logs = Get-AzLog `
            -ResourceId $Disk.Id `
            -StartTime $StartDate `
            -DetailedOutput `
            -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve activity logs."
        $Logs = @()
    }

    #
    # Determine current state
    #
    $CurrentAttachment = if ($Disk.ManagedBy) {
        $Disk.ManagedBy
    }
    else {
        "UNATTACHED"
    }

    #
    # Determine last VM relationship for orphaned disks
    #
    $LastVM = "Unknown"
    $LastAttachDate = $null
    $LastAttachUser = $null
    $LastAttachStatus = $null

    if (-not $Disk.ManagedBy) {

        Write-Host "Disk is currently orphaned/unattached. Attempting to identify previous usage..." -ForegroundColor Yellow

        $LastRelevantEvent = $Logs |
            Sort-Object EventTimestamp -Descending |
            Where-Object {
                $_.ResourceId -match "Microsoft.Compute" -or
                $_.OperationName.LocalizedValue -match "disk" -or
                $_.OperationName.LocalizedValue -match "virtual machine"
            } |
            Select-Object -First 1

        if ($LastRelevantEvent) {

            $LastAttachDate = $LastRelevantEvent.EventTimestamp
            $LastAttachUser = $LastRelevantEvent.Caller
            $LastAttachStatus = $LastRelevantEvent.Status.LocalizedValue

            if ($LastRelevantEvent.ResourceId) {
                $LastVM = $LastRelevantEvent.ResourceId
            }
        }
    }
    else {

        $LastVM = $Disk.ManagedBy
    }

    #
    # Creation event
    #
    $CreationEvent = $Logs |
        Sort-Object EventTimestamp |
        Where-Object {
            $_.OperationName.LocalizedValue -match "create"
        } |
        Select-Object -First 1

    #
    # Failed operations
    #
    $FailedEvents = $Logs |
        Where-Object {
            $_.Status.LocalizedValue -ne "Succeeded"
        }

    #
    # Disk summary object
    #
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

        StorageType             = $Disk.Sku.Name
        SecurityType            = $Disk.SecurityProfile.SecurityType

        TimeCreated             = $Disk.TimeCreated

        LastKnownVM             = $LastVM
        LastKnownActivity       = $LastAttachDate
        LastKnownActivityBy     = $LastAttachUser
        LastKnownActivityStatus = $LastAttachStatus

        CreatedBy               = $CreationEvent.Caller
        CreateTime              = $CreationEvent.EventTimestamp

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

    #
    # Display summary
    #
    Write-Host ""
    Write-Host "=== DISK SUMMARY ===" -ForegroundColor Green
    $Summary | Format-List

    #
    # Export detailed audit log
    #
    $AuditExport = $Logs | Select-Object `
        EventTimestamp,
        Caller,
        CorrelationId,
        ResourceId,
        @{N="Operation";E={$_.OperationName.LocalizedValue}},
        @{N="Status";E={$_.Status.LocalizedValue}},
        @{N="SubStatus";E={$_.SubStatus.LocalizedValue}},
        @{N="Category";E={$_.Category.LocalizedValue}}

    $OutputFile = "C:\Temp\$($Disk.Name)_30Day_Audit.csv"

    $AuditExport |
        Sort-Object EventTimestamp -Descending |
        Export-Csv $OutputFile -NoTypeInformation

    Write-Host ""
    Write-Host "Audit events exported to:" -ForegroundColor Green
    Write-Host $OutputFile

    #
    # Failed operations summary
    #
    if ($FailedEvents.Count -gt 0) {

        Write-Host ""
        Write-Host "=== FAILED OPERATIONS DETECTED ===" -ForegroundColor Yellow

        $FailedEvents |
            Select-Object `
                EventTimestamp,
                Caller,
                @{N='Operation';E={$_.OperationName.LocalizedValue}},
                @{N='Status';E={$_.Status.LocalizedValue}} |
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

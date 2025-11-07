<# 

#>

param(
    [string]$ServerName,
    [string]$Memory,
    [string]$CpuCount,
    [string]$Timezone
)

# =========================
# CONFIG SECTION (EDIT ME)
# =========================

# VMM server and filters
$VmmServerName              = "SCVMM-SERVER-NAME"          
$ClusterNameFilter          = "*Cluster*"                        # what clusters/hosts to target
$VmHostNameFilter           = "*node*"                           # host selection filter

# Logging / working paths
$BaseWorkFolder             = "C:\VMBuild"                  
$LogFilePath                = Join-Path $BaseWorkFolder "trigger.log"
$CleanupFilePath            = Join-Path $BaseWorkFolder "cleanup.txt"
$CleanupArchiveFolder       = Join-Path $BaseWorkFolder "archive"
$CleanupArchiveFileName     = "cleanup.txt"

# Slack / webhook (placeholder)
$SlackEnabled               = $true
$SlackWebhookUrl            = "https://hooks.slack.com/services/......."
$SlackChannelTitle          = "VM Create status"

# Existing template / OS / answer file / networks – these MUST match your VMM
$ExistingTemplateId         = "00000000-0000-0000-0000-000000000000" # <-- replace with real template ID
$OperatingSystemId          = "00000000-0000-0000-0000-000000000000" # <-- replace with real OS ID
$VmNetworkName              = "VM-NETWORK-NAME"                      # e.g. "10.5.10.x"
$VmNetworkId                = "00000000-0000-0000-0000-000000000000" # if you must pin an ID
$StaticIPv4PoolName         = "STATIC-POOL-NAME"                     # e.g. "10.5.14.x"
$StaticIPv4PoolId           = "00000000-0000-0000-0000-000000000000"

# Library / answer file path in VMM
$UnattendSharePath          = "\\VMM-LIBRARY-SERVER\Library\Win2022\unattend.xml"

# Storage selection thresholds
$VolumeNameExclusionPattern = "*FS0*"        # pattern to exclude some CSVs
$VolumeFullThreshold        = 16             # VMs per volume
$VolumeMinFreeSpaceGB       = 600            # minimum free space

# =========================
# IMPORT VMM MODULE
# =========================
ipmo 'virtualmachinemanager\virtualmachinemanager.psd1'

# =========================
# HELPER: WRITE TO LOG
# =========================
function Write-BuildLog {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts $Message" | Tee-Object -FilePath $LogFilePath -Append
}

# =========================
# HELPER: SEND SLACK
# =========================
function Send-SlackMessage {
    param(
        [string]$Text,
        [string]$Color = "#152911"
    )
    if (-not $SlackEnabled) { return }
    if (-not $SlackWebhookUrl) { return }

    $body = @{
        title   = $SlackChannelTitle
        pretext = $SlackChannelTitle
        text    = $Text
        color   = $Color
    } | ConvertTo-Json

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $body -ContentType 'application/json' | Out-Null
    } catch {
        Write-BuildLog "Slack notification failed: $_"
    }
}

# make sure working folders exist
if (-not (Test-Path $BaseWorkFolder))       { New-Item -ItemType Directory -Force -Path $BaseWorkFolder | Out-Null }
if (-not (Test-Path $CleanupArchiveFolder)) { New-Item -ItemType Directory -Force -Path $CleanupArchiveFolder | Out-Null }

Write-BuildLog "SCRIPT STARTED"

try {
    # normalize params
    [int]$MemoryMB    = $Memory
    [int]$CpuCountInt = $CpuCount
    [int]$TimezoneInt = $Timezone

    # =========================
    # CLUSTER OVERCOMMIT CHECK
    # =========================
    $clusters = Get-SCVMHostCluster -VMMServer $VmmServerName | Where-Object { $_.Name -like $ClusterNameFilter }

    foreach ($cluster in $clusters) {
        if ($cluster.ClusterReserveState -eq "Over") {
            $msg = "Rejecting VM request – cluster overcommitted: $($cluster.Name) Res: $($cluster.ClusterReserve)"
            Write-BuildLog $msg
            Send-SlackMessage -Text $msg -Color "#ff0000"
            throw "Script stopped due to overcommitted cluster: $($cluster.Name)"
        } else {
            Write-BuildLog "$($cluster.Name): Not overcommitted"
        }
    }

    # =========================
    # CLEAN UP PREVIOUS TEMPLATE
    # =========================
    if (Test-Path $CleanupFilePath) {
        $cleanupContent = Get-Content -Path $CleanupFilePath

        $templateValue = $cleanupContent | Select-String -Pattern "Template: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
        $profileValue  = $cleanupContent | Select-String -Pattern "Profile: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }

        if ($profileValue) {
            $HardwareProfileToDelete = Get-SCHardwareProfile -VMMServer $VmmServerName | Where-Object { $_.Name -eq $profileValue }
            if ($HardwareProfileToDelete) {
                Write-BuildLog "Cleaning up hardware profile $profileValue"
                Remove-SCHardwareProfile -HardwareProfile $HardwareProfileToDelete
            }
        }

        if ($templateValue) {
            $TemplateToDelete = Get-SCVMTemplate -VMMServer $VmmServerName -All | Where-Object { $_.Name -eq $templateValue }
            if ($TemplateToDelete) {
                Write-BuildLog "Cleaning up template $templateValue"
                Remove-SCVMTemplate -VMTemplate $TemplateToDelete
            }
        }

        $archiveFile = (Get-Date -Format yyyyMMddHHmmss) + "_" + $CleanupArchiveFileName
        $archivePath = Join-Path $CleanupArchiveFolder $archiveFile
        Move-Item $CleanupFilePath $archivePath
    } else {
        Write-BuildLog "No previous cleanup file found."
    }

    # =========================
    # START BUILD
    # =========================
    [Guid]$JobGroup = [Guid]::NewGuid()

    # SCSI adapter
    New-SCVirtualScsiAdapter -VMMServer $VmmServerName -JobGroup $JobGroup -AdapterID 7 -ShareVirtualScsiAdapter:$false -ScsiControllerType DefaultTypeNoType

    # VM network
    $VMNetwork = Get-SCVMNetwork -VMMServer $VmmServerName -Name $VmNetworkName
    # If you *must* pin the ID, uncomment:
    # $VMNetwork = Get-SCVMNetwork -VMMServer $VmmServerName | Where-Object { $_.Id -eq $VmNetworkId }

    New-SCVirtualNetworkAdapter `
        -VMMServer $VmmServerName `
        -JobGroup $JobGroup `
        -MACAddress "00:00:00:00:00:00" `
        -MACAddressType Static `
        -VirtualNetwork $VmNetworkName `
        -VLanEnabled $true `
        -VLanID 5 `
        -Synthetic `
        -EnableVMNetworkOptimization $false `
        -EnableMACAddressSpoofing $false `
        -EnableGuestIPNetworkVirtualizationUpdates $false `
        -IPv4AddressType Static `
        -IPv6AddressType Dynamic `
        -VMNetwork $VMNetwork `
        -DevicePropertiesAdapterNameMode Disabled

    # CPU type (make this generic)
    $CPUType = Get-SCCPUType -VMMServer $VmmServerName | Select-Object -First 1

    # Hardware profile name
    $hardwareGuid = "Profile-" + [guid]::NewGuid().ToString()

    New-SCHardwareProfile -VMMServer $VmmServerName `
        -CPUType $CPUType `
        -Name $hardwareGuid `
        -Description "Profile used to create a VM/Template" `
        -CPUCount $CpuCountInt `
        -MemoryMB $MemoryMB `
        -DynamicMemoryEnabled:$false `
        -HighlyAvailable:$true `
        -SecureBootEnabled:$true `
        -SecureBootTemplate "MicrosoftWindows" `
        -Generation 2 `
        -JobGroup $JobGroup

    # Base template / OS / unattend
    $TemplateOrig    = Get-SCVMTemplate -VMMServer $VmmServerName -ID $ExistingTemplateId
    $HardwareProfile = Get-SCHardwareProfile -VMMServer $VmmServerName | Where-Object { $_.Name -eq $hardwareGuid }
    $AnswerFile      = Get-SCScript -VMMServer $VmmServerName | Where-Object { $_.SharePath -eq $UnattendSharePath }
    $OperatingSystem = Get-SCOperatingSystem -VMMServer $VmmServerName -ID $OperatingSystemId

    $TemplateID = "Template-" + [guid]::NewGuid().ToString()

    New-SCVMTemplate -Name $TemplateID `
        -Template $TemplateOrig `
        -EnableNestedVirtualization:$false `
        -HardwareProfile $HardwareProfile `
        -JobGroup (New-Guid) `
        -ComputerName $ServerName `
        -TimeZone $TimezoneInt `
        -AnswerFile $AnswerFile `
        -OperatingSystem $OperatingSystem `
        -UpdateManagementProfile $null

    $template = Get-SCVMTemplate -VMMServer $VmmServerName -All | Where-Object { $_.Name -eq $TemplateID }

    # =========================
    # STORAGE SELECTION (CSV)
    # =========================
    $vmHosts = Get-SCVMHost -VMMServer $VmmServerName | Where-Object { $_.Name -like $VmHostNameFilter }

    try {
        $allClusterVolumes = $vmHosts |
            Get-SCStorageVolume |
            Where-Object { $_.Name -like "*ClusterStorage*" } |
            Where-Object { $_.Name -notlike $VolumeNameExclusionPattern }
    } catch {
        Write-BuildLog "Failed to retrieve storage volumes: $_"
        throw
    }

    $volumeCount = @{}
    $allVolumes  = @()
    $vmVolumes   = @()

    foreach ($vmHost in $vmHosts) {
        try {
            $storageVolumes = $vmHost |
                Get-SCStorageVolume |
                Where-Object { $_.Name -like "*ClusterStorage*" } |
                Where-Object { $_.Name -notlike $VolumeNameExclusionPattern }
        } catch {
            Write-BuildLog "Failed to retrieve storage for host $($vmHost.Name): $_"
            continue
        }

        $allVolumes += $storageVolumes

        try {
            $vms = Get-SCVirtualMachine -VMHost $vmHost
        } catch {
            Write-BuildLog "Failed to retrieve VMs for host $($vmHost.Name): $_"
            continue
        }

        foreach ($vm in $vms) {
            $volumePath = ($vm.Location -split '\\')[0..2] -join '\'
            if (![string]::IsNullOrEmpty($volumePath) -and $volumePath -like "*ClusterStorage*") {
                if ($volumeCount.ContainsKey($volumePath)) {
                    $volumeCount[$volumePath].Count++
                } else {
                    $volumeCount[$volumePath] = [PSCustomObject]@{
                        Count = 1
                        Space = 0
                    }
                }
                $vmVolumes += $volumePath
            }
        }
    }

    foreach ($volume in $allClusterVolumes) {
        $volumePath = $volume.Name
        if ($volumeCount.ContainsKey($volumePath)) {
            $volumeCount[$volumePath].Space = $volume.FreeSpace
        } else {
            $volumeCount[$volumePath] = [PSCustomObject]@{
                Count = 0
                Space = $volume.FreeSpace
            }
        }
    }

    $volumeCountDisplay = $volumeCount.GetEnumerator() | ForEach-Object {
        $vmCount = $_.Value.Count
        $spaceGB = [math]::Round($_.Value.Space / 1GB, 2)
        $status = if ($spaceGB -lt $VolumeMinFreeSpaceGB) {
            "Too Low on Space"
        } elseif ($vmCount -gt $VolumeFullThreshold) {
            "Overcommitted"
        } elseif ($vmCount -eq $VolumeFullThreshold) {
            "Full"
        } else {
            "Usable"
        }

        [PSCustomObject]@{
            VolumePath      = $_.Key
            VMCount         = $vmCount
            FreeSpaceGB     = $spaceGB
            Status          = $status
            AdditionalVMs   = if ($status -eq "Too Low on Space") { 0 } else { $VolumeFullThreshold - $vmCount }
            OvercommittedBy = if ($status -eq "Overcommitted") { $vmCount - $VolumeFullThreshold } else { 0 }
        }
    } | Sort-Object -Property VMCount, FreeSpaceGB -Descending

    $targetVolume = $volumeCountDisplay |
        Where-Object { $_.Status -notin @("Full","Overcommitted","Too Low on Space") } |
        Sort-Object -Property VMCount, FreeSpaceGB -Descending |
        Select-Object -First 1

    if (-not $targetVolume) {
        throw "No suitable volume found."
    }

    Write-BuildLog "Next volume to be used: $($targetVolume.VolumePath)"

    # =========================
    # HOST SELECTION
    # =========================
    $hosts = Get-SCVMHost -VMMServer $VmmServerName | Where-Object { $_.Name -like $VmHostNameFilter }
    $hosts = $hosts | Sort-Object -Property AvailableMemoryMB -Descending
    $vmHost = $hosts[0]

    # =========================
    # VM CONFIGURATION
    # =========================
    $vmConfiguration = New-SCVMConfiguration -VMTemplate $template -Name $ServerName

    Set-SCVMConfiguration -VMConfiguration $vmConfiguration -VMHost $vmHost
    Update-SCVMConfiguration -VMConfiguration $vmConfiguration
    Set-SCVMConfiguration -VMConfiguration $vmConfiguration -ComputerName $ServerName
    Set-SCVMConfiguration -VMConfiguration $vmConfiguration -VMLocation $targetVolume.VolumePath -PinVMLocation $true

    # network pinning
    $AllNICConfigurations = Get-SCVirtualNetworkAdapterConfiguration -VMConfiguration $vmConfiguration
    $NicConfiguration     = $AllNICConfigurations[0]
    $StaticIPv4Pool       = Get-SCStaticIPAddressPool -VMMServer $VmmServerName -Name $StaticIPv4PoolName
    if ($StaticIPv4PoolId -and -not $StaticIPv4Pool) {
        $StaticIPv4Pool = Get-SCStaticIPAddressPool -VMMServer $VmmServerName | Where-Object { $_.Id -eq $StaticIPv4PoolId }
    }
    if ($StaticIPv4Pool) {
        Set-SCVirtualNetworkAdapterConfiguration -VirtualNetworkAdapterConfiguration $NicConfiguration `
            -IPv4AddressPool $StaticIPv4Pool `
            -PinIPv4AddressPool $true `
            -PinIPv6AddressPool $false `
            -PinMACAddressPool $false
    }

    Update-SCVMConfiguration -VMConfiguration $vmConfiguration

    # write cleanup file for next run
    "Template: $TemplateID`nProfile: $hardwareGuid" | Out-File -FilePath $CleanupFilePath -Force

    # create VM
    New-SCVirtualMachine -Name $ServerName -VMConfiguration $vmConfiguration -JobGroup $JobGroup -ReturnImmediately `
        -StartAction "TurnOnVMIfRunningWhenVSStopped" -StopAction "SaveVM" -StartVM

    $okText = "VM $ServerName is being deployed on host $($vmHost.Name) volume $($targetVolume.VolumePath)"
    Write-BuildLog $okText
    Send-SlackMessage -Text $okText -Color "#142911"

    Write-BuildLog "SCRIPT ENDED with Success"
    exit 0
}
catch {
    Write-BuildLog "SCRIPT ENDED with Errors: $_"
    Send-SlackMessage -Text "VM creation failed: $_" -Color "#ff0000"
    exit 1
}
finally {
    # optional cleanup
}

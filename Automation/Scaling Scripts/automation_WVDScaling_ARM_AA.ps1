

param(
  [Parameter(mandatory = $false)]
  [object]$WebHookData
)
# If the runbook was called from a Webhook, the WebhookData will not be null.
if ($WebHookData) {

  # Collect properties of WebhookData
  $WebhookName = $WebHookData.WebhookName
  $WebhookHeaders = $WebHookData.RequestHeader
  $WebhookBody = $WebHookData.RequestBody

  # Collect individual headers. Input converted from JSON.
  $From = $WebhookHeaders.From
  $Input = (ConvertFrom-Json -InputObject $WebhookBody)
}
else {
  Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}

$AADTenantId = $Input.AADTenantId
$SubscriptionID = $Input.SubscriptionID
$ResourceGroupName = $Input.ResourceGroupName
$HostpoolName = $Input.hostPoolName
$WorkDays = $Input.workDays
$BeginPeakTime = $Input.beginPeakTime
$EndPeakTime = $Input.endPeakTime
$TimeDifferenceInHours = $Input.TimeDifferenceInHours
$PeakLoadBalancingType = $Input.PeakLoadBalancingType
$OffPeakLoadBalancingType = $Input.OffPeakLoadBalancingType
$peakMaxSessions = $Input.peakMaxSessions
$offpeakMaxSessions = $Input.OffPeakMaxSessions
$peakScaleFactor = $Input.peakScaleFactor
$offpeakScaleFactor = $Input.offpeakScaleFactor
$peakMinimumNumberOfRDSH = $Input.peakMinimumNumberOfRDSH
$offpeakMinimumNumberOfRDSH = $Input.offpeakMinimumNumberOfRDSH
$minimumNumberFastScale = $Input.minimumNumberFastScale
$jobTimeout = $Input.jobTimeout
$LimitSecondsToForceLogOffUser = $Input.LimitSecondsToForceLogOffUser
$LogOffMessageTitle = $Input.LogOffMessageTitle
$LogOffMessageBody = $Input.LogOffMessageBody
$MaintenanceTagName = $Input.MaintenanceTagName
$LogAnalyticsWorkspaceId = $Input.LogAnalyticsWorkspaceId
$LogAnalyticsPrimaryKey = $Input.LogAnalyticsPrimaryKey
$ConnectionAssetName = $Input.ConnectionAssetName

Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process -Force -Confirm:$false
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force -Confirm:$false

# Setting ErrorActionPreference to stop script execution when error occurs
$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Function for converting UTC to Local time
function Convert-UTCtoLocalTime {
  param(
    $TimeDifferenceInHours
  )

  $UniversalTime = (Get-Date).ToUniversalTime()
  $TimeDifferenceMinutes = 0
  if ($TimeDifferenceInHours -match ":") {
    $TimeDifferenceHours = $TimeDifferenceInHours.Split(":")[0]
    $TimeDifferenceMinutes = $TimeDifferenceInHours.Split(":")[1]
  }
  else {
    $TimeDifferenceHours = $TimeDifferenceInHours
  }
  #Azure is using UTC time, justify it to the local time
  $ConvertedTime = $UniversalTime.AddHours($TimeDifferenceHours).AddMinutes($TimeDifferenceMinutes)
  return $ConvertedTime
}

# Function to add logs to Log Analytics Workspace
function Add-LogEntry {
  param(
    [Object]$LogMessageObj,
    [string]$LogAnalyticsWorkspaceId,
    [string]$LogAnalyticsPrimaryKey,
    [string]$LogType,
    $TimeDifferenceInHours
  )

  if ($LogAnalyticsWorkspaceId -ne $null) {

    foreach ($Key in $LogMessage.Keys) {
      switch ($Key.substring($Key.Length - 2)) {
        '_s' { $sep = '"'; $trim = $Key.Length - 2 }
        '_t' { $sep = '"'; $trim = $Key.Length - 2 }
        '_b' { $sep = ''; $trim = $Key.Length - 2 }
        '_d' { $sep = ''; $trim = $Key.Length - 2 }
        '_g' { $sep = '"'; $trim = $Key.Length - 2 }
        default { $sep = '"'; $trim = $Key.Length }
      }
      $LogData = $LogData + '"' + $Key.substring(0, $trim) + '":' + $sep + $LogMessageObj.Item($Key) + $sep + ','
    }

    $TimeStamp = Convert-UTCtoLocalTime -TimeDifferenceInHours $TimeDifferenceInHours
    $LogData = $LogData + '"TimeStamp":"' + $timestamp + '"'
    $json = "{$($LogData)}"
    $PostResult = Send-OMSAPIIngestionFile -CustomerId $LogAnalyticsWorkspaceId -SharedKey $LogAnalyticsPrimaryKey -Body "$json" -LogType $LogType -TimeStampField "TimeStamp"
    
    if ($PostResult -ne "Accepted") {
      Write-Error "Error posting to OMS - $PostResult"
    }
  }
}

# Construct Begin time and End time for the Peak/Off-Peak periods from UTC to local time
$TimeDifference = [string]$TimeDifferenceInHours
$CurrentDateTime = Convert-UTCtoLocalTime -TimeDifferenceInHours $TimeDifference

# Collect the credentials from Azure Automation Account Assets
$Connection = Get-AutomationConnection -Name $ConnectionAssetName

# Authenticate to Azure 
Clear-AzContext -Force
$AZAuthentication = Connect-AzAccount -ApplicationId $Connection.ApplicationId -TenantId $AADTenantId -CertificateThumbprint $Connection.CertificateThumbprint -ServicePrincipal
if ($AZAuthentication -eq $null) {
  Write-Error "Failed to authenticate to Azure using the Automation Account $($_.exception.message)"
  exit
} 
else {
  $AzObj = $AZAuthentication | Out-String
  Write-Output "Authenticated to Azure using the Automation Account `n$AzObj"
}

# Set the Azure context with Subscription
$AzContext = Set-AzContext -SubscriptionId $SubscriptionID
if ($AzContext -eq $null) {
  Write-Error "Subscription '$SubscriptionID' does not exist. Ensure that you have entered the correct values"
  exit
} 
else {
  $AzSubObj = $AzContext | Out-String
  Write-Output "Set the Azure Context to the correct Subscription `n$AzSubObj"
}

# Convert Datetime format
$BeginPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $BeginPeakTime)
$EndPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $EndPeakTime)

# Check the calculated end peak time is later than begin peak time in case of going between days
if ($EndPeakDateTime -lt $BeginPeakDateTime) {
  if ($CurrentDateTime -lt $EndPeakDateTime) { $BeginPeakDateTime = $BeginPeakDateTime.AddDays(-1) } else { $EndPeakDateTime = $EndPeakDateTime.AddDays(1) }
}

# Create the time period for the Peak to Off Peak Transition period
$peakToOffPeakTransitionTime = $EndPeakDateTime.AddMinutes(15)

# Check given hostpool name exists
$HostpoolInfo = Get-AzWvdHostPool -ResourceGroupName $ResourceGroupName -Name $HostpoolName
if ($HostpoolInfo -eq $null) {
  Write-Error "Hostpoolname '$HostpoolName' does not exist. Ensure that you have entered the correct values"
  exit
}	

# Get todays day of week for comparing to Work Days
$today = (Get-Date).DayOfWeek

# Compare Work Days and Peak Hours, and set up appropriate load balancing type based on PeakLoadBalancingType & OffPeakLoadBalancingType
if (($CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime) -and ($WorkDays -contains $today)) {

  if ($HostpoolInfo.LoadBalancerType -ne $PeakLoadBalancingType) {
    Write-Output "Changing Hostpool Load Balance Type to: $PeakLoadBalancingType Load Balancing"

    if ($PeakLoadBalancingType -eq "DepthFirst") {                
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'DepthFirst' -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    else {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'BreadthFirst' -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    Write-Output "Hostpool Load Balance Type in Peak Hours is: $PeakLoadBalancingType Load Balancing"
  }
  # Compare MaxSessionLimit of hostpool to peakMaxSessions value and adjust if necessary
  if ($HostpoolInfo.MaxSessionLimit -ne $peakMaxSessions) {
    Write-Output "Changing Hostpool Peak MaxSessionLimit to: $peakMaxSessions"

    if ($PeakLoadBalancingType -eq "DepthFirst") {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'DepthFirst' -MaxSessionLimit $peakMaxSessions
    }
    else {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'BreadthFirst' -MaxSessionLimit $peakMaxSessions
    }
  }
}
else {
  if ($HostpoolInfo.LoadBalancerType -ne $OffPeakLoadBalancingType) {
    Write-Output "Changing Hostpool Load Balance Type to: $OffPeakLoadBalancingType Load Balancing"
        
    if ($OffPeakLoadBalancingType -eq "DepthFirst") {                
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'DepthFirst' -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    else {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'BreadthFirst' -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    Write-Output "Hostpool Load Balance Type in Off-Peak Hours is: $OffPeakLoadBalancingType Load Balancing"
  }
  # Compare MaxSessionLimit of hostpool to offpeakMaxSessions value and adjust if necessary
  if ($HostpoolInfo.MaxSessionLimit -ne $offpeakMaxSessions) {
    Write-Output "Changing Hostpool Off-Peak MaxSessionLimit to: $offpeakMaxSessions"

    if ($PeakLoadBalancingType -eq "DepthFirst") {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'DepthFirst' -MaxSessionLimit $offpeakMaxSessions
    }
    else {
      Update-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $HostpoolName -LoadBalancerType 'BreadthFirst' -MaxSessionLimit $offpeakMaxSessions
    }
  }
}

# Check for VM's with maintenance tag set to True & ensure connections are set as not allowed
$AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName

foreach ($SessionHost in $AllSessionHosts) {

  $SessionHostName = $SessionHost.Name
  $SessionHostName = $SessionHostName.Split("/")[1]
  $VMName = $SessionHostName.Split(".")[0]
  $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

  if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
    Write-Output "The host $VMName is in Maintenance mode, so is not allowing any further connections"
    Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$False -ErrorAction SilentlyContinue
  }
  else {
    Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
  }
}

Write-Output "Starting WVD Hosts Scale Optimization"

# Check the Hostpool Load Balancer type
$HostpoolInfo = Get-AzWvdHostPool -ResourceGroupName $resourceGroupName -Name $hostPoolName
Write-Output "Hostpool Load Balancing Type is: $($HostpoolInfo.LoadBalancerType)"

# Check if it's peak hours
if (($CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime) -and ($WorkDays -contains $today)) {

  # Gather hostpool maximum sessions and calculate Scalefactor for each host.										  
  $HostpoolMaxSessionLimit = $HostpoolInfo.MaxSessionLimit
  $ScaleFactorEachHost = $HostpoolMaxSessionLimit * $peakScaleFactor
  $SessionhostLimit = [math]::Floor($ScaleFactorEachHost)

  Write-Output "It is currently: Peak hours"
  Write-Output "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)"
  Write-Output "Checking current Host availability and workloads..."

  # Get all session hosts in the host pool
  $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
  if ($AllSessionHosts -eq $null) {
    Write-Error "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it"
    exit
  }
  
  # Check the number of available running session hosts
  $NumberOfRunningHost = 0
  foreach ($SessionHost in $AllSessionHosts) {

    $SessionHostName = $SessionHost.Name
    $SessionHostName = $SessionHostName.Split("/")[1]
    $VMName = $SessionHostName.Split(".")[0]
    Write-Output "Host:$VMName, Current sessions:$($SessionHost.Session), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)"

    if ($SessionHost.Status -eq "Available" -and $SessionHost.AllowNewSession -eq $True) {
      $NumberOfRunningHost = $NumberOfRunningHost + 1
    }
  }
  Write-Output "Current number of available running hosts: $NumberOfRunningHost"

  # Start more hosts if available host number is less than the specified Peak minimum number of hosts
  if ($NumberOfRunningHost -lt $peakMinimumNumberOfRDSH) {
    Write-Output "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) - Need to start additional hosts"

    $global:peakMinRDSHcapacityTrigger = $True

    :peakMinStartupLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($NumberOfRunningHost -ge $peakMinimumNumberOfRDSH) {

        if ($minimumNumberFastScale -eq $True) {
          Write-Output "The number of available running hosts should soon equal the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)"
          break peakMinStartupLoop
        }
        else {
          Write-Output "The number of available running hosts ($NumberOfRunningHost) now equals the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)"
          break peakMinStartupLoop
        }
      }

      # Check the session hosts status to determine it's healthy before starting it
      if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
        $SessionHostName = $SessionHost.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

        # Check to see if the Session host is in maintenance mode
        if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
          Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
          continue
        }

        # Ensure the host has allow new connections set to True
        if ($SessionHost.AllowNewSession = $False) {
          try {
            Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
          }
          catch {
            Write-Error "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
            exit 1
          }
        }
        if ($minimumNumberFastScale -eq $True) {

          # Start the Azure VM in Fast-Scale Mode for parallel processing
          try {
            Write-Output "Starting host $VMName in fast-scale mode..."
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -AsJob

          }
          catch {
            Write-Error "Failed to start host $VMName with error: $($_.exception.message)"
            exit
          }
        }
        if ($minimumNumberFastScale -eq $False) {

          # Start the Azure VM
          try {
            Write-Output "Starting host $VMName and waiting for it to complete..."
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName

          }
          catch {
            Write-Error "Failed to start host $VMName with error: $($_.exception.message)"
            exit
          }
          # Wait for the session host to become available
          $IsHostAvailable = $false
          while (!$IsHostAvailable) {

            $SessionHostStatus = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName

            if ($SessionHostStatus.Status -eq "Available") {
              $IsHostAvailable = $true

            }
          }
        }
        $NumberOfRunningHost = $NumberOfRunningHost + 1
        $global:spareCapacity = $True
      }
    }
  }
  else {
    $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name

    :mainLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($SessionHost.Session -le $HostpoolMaxSessionLimit -or $SessionHost.Session -gt $HostpoolMaxSessionLimit) {
        if ($SessionHost.Session -ge $SessionHostLimit) {
          $SessionHostName = $SessionHost.Name
          $SessionHostName = $SessionHostName.Split("/")[1]
          $VMName = $SessionHostName.Split(".")[0]

          # Check if a hosts sessions have exceeded the Peak scale factor
          if (($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) -and ($global:capacityTrigger -eq $False -or !$global:capacityTrigger)) {
            Write-Output "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..."
            $global:capacityTrigger = $True
          }

          :startupLoop  foreach ($SessionHost in $AllSessionHosts) {

            # Check the existing session hosts spare capacity before starting another host
            if ($SessionHost.Status -eq "Available" -and ($SessionHost.Session -ge 0 -and $SessionHost.Session -lt $SessionHostLimit) -and $SessionHost.AllowNewSession -eq $True) {
              $SessionHostName = $SessionHost.Name
              $SessionHostName = $SessionHostName.Split("/")[1]
              $VMName = $SessionHostName.Split(".")[0]

              if ($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) {
                Write-Output "Host $VMName has spare capacity so don't need to start another host. Continuing now..."
                $global:exceededHostCapacity = $True
                $global:spareCapacity = $True
              }
              break startupLoop
            }

            # Check the session hosts status to determine it's healthy before starting it
            if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
              $SessionHostName = $SessionHost.Name
              $SessionHostName = $SessionHostName.Split("/")[1]
              $VMName = $SessionHostName.Split(".")[0]
              $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

              # Check to see if the Session host is in maintenance mode
              if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                continue
              }

              # Ensure the host has allow new connections set to True
              if ($SessionHost.AllowNewSession = $False) {
                try {
                  Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
                }
                catch {
                  Write-Error "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
                  exit 1
                }
              }

              # Start the Azure VM
              try {
                Write-Output "There is not enough spare capacity on other active hosts. A new host will now be started..."
                Write-Output "Starting host $VMName and waiting for it to complete..."
                Start-AzVM -Name $VMName -ResourceGroupName $VMInfo.ResourceGroupName
              }
              catch {
                Write-Error "Failed to start host $VMName with error: $($_.exception.message)"
                exit
              }

              # Wait for the session host to become available
              $IsHostAvailable = $false
              while (!$IsHostAvailable) {

                $SessionHostStatus = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName

                if ($SessionHostStatus.Status -eq "Available") {
                  $IsHostAvailable = $true
                }
              }
              $NumberOfRunningHost = $NumberOfRunningHost + 1
              $global:spareCapacity = $True
              Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"
              break mainLoop

            }
          }
        }
        # Shut down hosts utilizing unnecessary resource
        $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }
        $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
        :shutdownLoop foreach ($ActiveHost in $ActiveHostsZeroSessions) {
          
          $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }

          # Ensure there is at least the peakMinimumNumberOfRDSH sessions available
          if ($NumberOfRunningHost -le $peakMinimumNumberOfRDSH) {
            Write-Output "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Peak Minimum Number of RDSH = $peakMinimumNumberOfRDSH"
            break mainLoop
          }

          # Check for session capacity on other active hosts before shutting the free host down
          else {
            $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }
            :shutdownLoopTier2 foreach ($ActiveHost in $ActiveHostsZeroSessions) {

              $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
              foreach ($SessionHost in $AllSessionHosts) {

                if ($SessionHost.Status -eq "Available" -and ($SessionHost.Session -ge 0 -and $SessionHost.Session -lt $SessionHostLimit -and $SessionHost.AllowNewSession -eq $True)) {
                  if ($SessionHost.Name -ne $ActiveHost.Name) {
                    $ActiveHostName = $ActiveHost.Name
                    $ActiveHostName = $ActiveHostName.Split("/")[1]
                    $VMName = $ActiveHostName.Split(".")[0]
                    $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

                    # Check if the Session host is in maintenance
                    if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                      Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                      continue
                    }

                    Write-Output "Identified free host $VMName with $($ActiveHost.Session) sessions that can be shut down to save resource"

                    # Ensure the running Azure VM is set into drain mode
                    try {
                      Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName -AllowNewSession:$False -ErrorAction SilentlyContinue
                    }
                    catch {
                      Write-Error "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)"
                      exit
                    }
                    try {
                      Write-Output "Stopping host $VMName and waiting for it to complete..."
                      Stop-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
                    }
                    catch {
                      Write-Error "Failed to stop host $VMName with error: $($_.exception.message)"
                      exit
                    }
                    # Check if the session host server is healthy before enable allowing new connections
                    if ($ActiveHost.UpdateState -eq "Succeeded") {
                      # Ensure Azure VMs that are stopped have the allowing new connections state True
                      try {
                        Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
                      }
                      catch {
                        Write-Output "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
                        exit
                      }
                    }

                    # Wait after shutting down Host until it's Status returns as Unavailable
                    $IsShutdownHostUnavailable = $false
                    while (!$IsShutdownHostUnavailable) {

                      $shutdownHost = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName

                      if ($shutdownHost.Status -eq "Unavailable") {
                        $IsShutdownHostUnavailable = $true
                      }
                    }

                    # Decrement the number of running session hosts
                    $NumberOfRunningHost = $NumberOfRunningHost - 1
                    Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"

                    break shutdownLoop
                  }
                }
              }     
            }
          }  
        }
      }
    }
  }
  # Get all available hosts and write to WVDAvailableHosts log
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$NumberOfRunningHost" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDAvailableHosts_CL" -TimeDifferenceInHours $TimeDifference
}
else {

  # Gathering hostpool maximum session and calculating Scalefactor for each host.										  
  $HostpoolMaxSessionLimit = $HostpoolInfo.MaxSessionLimit
  $ScaleFactorEachHost = $HostpoolMaxSessionLimit * $offpeakScaleFactor
  $SessionhostLimit = [math]::Floor($ScaleFactorEachHost)

  Write-Output "It is currently: Off-Peak hours"
  Write-Output "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)"
  Write-Output "Checking current Host availability and workloads..."

  # Get all session hosts in the host pool
  $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
  if ($AllSessionHosts -eq $null) {
    Write-Error "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it"
    exit
  }

  # Check the number of running session hosts
  $NumberOfRunningHost = 0
  foreach ($SessionHost in $AllSessionHosts) {

    $SessionHostName = $SessionHost.Name
    $SessionHostName = $SessionHostName.Split("/")[1]
    $VMName = $SessionHostName.Split(".")[0]
    Write-Output "Host:$VMName, Current sessions:$($SessionHost.Session), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)"

    if ($SessionHost.Status -eq "Available" -and $SessionHost.AllowNewSession -eq $True) {
      $NumberOfRunningHost = $NumberOfRunningHost + 1
    }
  }
  Write-Output "Current number of Available Running Hosts: $NumberOfRunningHost"

  # Check if it is within PeakToOffPeakTransitionTime after the end of Peak time and set the Peak to Off-Peak transition trigger if true
  $peakToOffPeakTransitionTrigger = $false

  if (($CurrentDateTime -ge $EndPeakDateTime) -and ($CurrentDateTime -le $peakToOffPeakTransitionTime)){
    $peakToOffPeakTransitionTrigger = $True
  }

  # Check if user logoff is turned on in off peak
  if ($LimitSecondsToForceLogOffUser -ne 0 -and $peakToOffPeakTransitionTrigger -eq $True) {
    Write-Output "The Hostpool has recently transitioned to Off-Peak from Peak and force logging-off of users in Off-Peak is enabled. Checking if any resource can be saved..."

    if ($NumberOfRunningHost -gt $offpeakMinimumNumberOfRDSH) {
      Write-Output "The number of available running hosts is greater than the Off-Peak Minimum Number of RDSH. Logging-off procedure will now be started..."

      foreach ($SessionHost in $AllSessionHosts) {

        $SessionHostName = $SessionHost.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }
        
        if ($SessionHost.Status -eq "Available") {

          # Get the User sessions in the hostPool
          try {
            $HostPoolUserSessions = Get-AzWvdUserSession -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName
          }
          catch {
            Write-Error "Failed to retrieve user sessions in hostPool $($HostpoolName) with error: $($_.exception.message)"
            exit
          }

          Write-Output "Current sessions running on host $VMName : $($SessionHost.Session)"
        }
      } 
      
      Write-Output "Sending log off message to users..."
      
      $HostPoolUserSessions = Get-AzWvdUserSession -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName
      $ExistingSession = 0
      foreach ($Session in $HostPoolUserSessions) {
  
        $SessionHostName = $Session.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $SessionId = $Session.Id
        $SessionId = $SessionId.Split("/")[12]
        
        # Notify user to log off their session
        try {
          Send-AzWvdUserSessionMessage -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -SessionHostName $SessionHostName -UserSessionId $SessionId -MessageTitle $LogOffMessageTitle -MessageBody "$($LogOffMessageBody) - You will logged off in $($LimitSecondsToForceLogOffUser) seconds"
        }
        catch {
          Write-Error "Failed to send message to user with error: $($_.exception.message)"
        exit
        }
        $ExistingSession = $ExistingSession + 1
      }
      # List User Session count
      Write-Output "Logoff messages were sent to $ExistingSession user(s)"

      # Set all Available session hosts into drain mode to stop any more connections
      Write-Output "Setting all available hosts into Drain mode to stop any further connections whilst logging-off procedure is running..."
      $forceLogoffSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Status -eq "Available" }
      foreach ($SessionHost in $forceLogoffSessionHosts) {
        
        $SessionHostName = $SessionHost.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

        # Check to see if the Session host is in maintenance
        if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
          Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
          $NumberOfRunningHost = $NumberOfRunningHost - 1
          continue
        }
        try {
          Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$False -ErrorAction SilentlyContinue
        }
        catch {
          Write-Error "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)"
          exit
        }
      }
            
      # Wait for n seconds to log off users
      Write-Output "Waiting for $LimitSecondsToForceLogOffUser seconds before logging off users..."
      Start-Sleep -Seconds $LimitSecondsToForceLogOffUser

      # Force Users to log off
      Write-Output "Forcing users to log off now..."

      try {
        $HostPoolUserSessions = Get-AzWvdUserSession -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName
      }
      catch {
        Write-Error "Failed to retrieve list of user sessions in HostPool $HostpoolName with error: $($_.exception.message)"
        exit
      }
      $ExistingSession = 0
      foreach ($Session in $HostPoolUserSessions) {

        $SessionHostName = $Session.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $SessionId = $Session.Id
        $SessionId = $SessionId.Split("/")[12]

        # Log off user
        try {
          Remove-AzWvdUserSession -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -SessionHostName $SessionHostName -Id $SessionId
        }
        catch {
          Write-Error "Failed to log off user session $($Session.UserSessionid) on host $SessionHostName with error: $($_.exception.message)"
          exit
        }
        $ExistingSession = $ExistingSession + 1
      }

      # List User Logoff count
      Write-Output "$ExistingSession user(s) were logged off"

      foreach ($SessionHost in $forceLogoffSessionHosts) {
        if ($NumberOfRunningHost -gt $offpeakMinimumNumberOfRDSH) {

          $SessionHostName = $SessionHost.Name
          $SessionHostName = $SessionHostName.Split("/")[1]
          $VMName = $SessionHostName.Split(".")[0]
          $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

          # Wait for the drained sessions to update on the WVD service
          $HaveSessionsDrained = $false
          while (!$HaveSessionsDrained) {

            $SessionHostStatus = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -SessionHostName $SessionHostName

            if ($SessionHostStatus.Session -eq 0) {
              $HaveSessionsDrained = $true
              Write-Output "Host $VMName now has 0 sessions"
            }
          }

          # Shutdown the Azure VM
          try {
            Write-Output "Stopping host $VMName and waiting for it to complete..."
            Stop-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
          }
          catch {
            Write-Error "Failed to stop host $VMName with error: $($_.exception.message)"
            exit
          }
          
          # Check if the session host is healthy before allowing new connections
          if ($SessionHost.UpdateState -eq "Succeeded") {
            # Ensure Azure VMs that are stopped have the allow new connections state set to True
            try {
              Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
            }
            catch {
              Write-Error "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
              exit 1
            }
          }
          # Decrement the number of running session host
          $NumberOfRunningHost = $NumberOfRunningHost - 1
        }
      }
    }
  }

  #Get Session Hosts again in case force Log Off users has changed their state
  $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name

  if ($NumberOfRunningHost -lt $offpeakMinimumNumberOfRDSH) {
    Write-Output "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) - Need to start additional hosts"
    $global:offpeakMinRDSHcapacityTrigger = $True

    :offpeakMinStartupLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($NumberOfRunningHost -ge $offpeakMinimumNumberOfRDSH) {

        if ($minimumNumberFastScale -eq $True) {
          Write-Output "The number of available running hosts should soon equal the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)"
          break offpeakMinStartupLoop
        }
        else {
          Write-Output "The number of available running hosts ($NumberOfRunningHost) now equals the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)"
          break offpeakMinStartupLoop
        }
      }

      # Check the session host status and if the session host is healthy before starting the host
      if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
        $SessionHostName = $SessionHost.Name
        $SessionHostName = $SessionHostName.Split("/")[1]
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

        # Check to see if the Session host is in maintenance
        if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
          Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
          continue
        }

        # Ensure Azure VMs that are stopped have the allowing new connections state set to True
        if ($SessionHost.AllowNewSession = $False) {
          try {
            Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
          }
          catch {
            Write-Error "Unable to set it to allow connections on host $VMName with error: $($_.exception.message)"
            exit 1
          }
        }
        if ($minimumNumberFastScale -eq $True) {

          # Start the Azure VM in Fast-Scale Mode for parallel processing
          try {
            Write-Output "Starting host $VMName in fast-scale mode..."
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -AsJob

          }
          catch {
            Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
            exit
          }
        }
        if ($minimumNumberFastScale -eq $False) {

          # Start the Azure VM
          try {
            Write-Output "Starting host $VMName and waiting for it to complete..."
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName

          }
          catch {
            Write-Error "Failed to start host $VMName with error: $($_.exception.message)"
            exit
          }
          # Wait for the sessionhost to become available
          $IsHostAvailable = $false
          while (!$IsHostAvailable) {

            $SessionHostStatus = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName

            if ($SessionHostStatus.Status -eq "Available") {
              $IsHostAvailable = $true

            }
          }
        }
        $NumberOfRunningHost = $NumberOfRunningHost + 1
        $global:spareCapacity = $True
      }
    }
  }
  else {
    $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name

    :mainLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($SessionHost.Session -le $HostpoolMaxSessionLimit -or $SessionHost.Session -gt $HostpoolMaxSessionLimit) {
        if ($SessionHost.Session -ge $SessionHostLimit) {
          $SessionHostName = $SessionHost.Name
          $SessionHostName = $SessionHostName.Split("/")[1]
          $VMName = $SessionHostName.Split(".")[0]

          if (($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) -and ($global:capacityTrigger -eq $False -or !$global:capacityTrigger)) {
            Write-Output "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..."
            $global:capacityTrigger = $True
          }

          :startupLoop  foreach ($SessionHost in $AllSessionHosts) {
            # Check the existing session hosts and session availability before starting another session host
            if ($SessionHost.Status -eq "Available" -and ($SessionHost.Session -ge 0 -and $SessionHost.Session -lt $SessionHostLimit) -and $SessionHost.AllowNewSession -eq $True) {
              $SessionHostName = $SessionHost.Name
              $SessionHostName = $SessionHostName.Split("/")[1]
              $VMName = $SessionHostName.Split(".")[0]

              if ($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) {
                Write-Output "Host $VMName has spare capacity so don't need to start another host. Continuing now..."

                $global:exceededHostCapacity = $True
                $global:spareCapacity = $True
              }
              break startupLoop
            }

            # Check the session host status and if the session host is healthy before starting the host
            if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
              $SessionHostName = $SessionHost.Name
              $SessionHostName = $SessionHostName.Split("/")[1]
              $VMName = $SessionHostName.Split(".")[0]
              $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

              # Check if the session host is in maintenance
              if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                continue
              }

              # Ensure Azure VMs that are stopped have the allowing new connections state set to True
              if ($SessionHost.AllowNewSession = $False) {
                try {
                  Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
                }
                catch {
                  Write-Error "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)"
                  exit 1
                }
              }

              # Start the Azure VM
              try {
                Write-Output "There is not enough spare capacity on other active hosts. A new host will now be started..."
                Write-Output "Starting host $VMName and waiting for it to complete..."
                Start-AzVM -Name $VMName -ResourceGroupName $VMInfo.ResourceGroupName
              }
              catch {
                Write-Error "Failed to start host $VMName with error: $($_.exception.message)"
                exit
              }
              # Wait for the sessionhost to become available
              $IsHostAvailable = $false
              while (!$IsHostAvailable) {

                $SessionHostStatus = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $SessionHostName

                if ($SessionHostStatus.Status -eq "Available") {
                  $IsHostAvailable = $true
                }
              }
              $NumberOfRunningHost = $NumberOfRunningHost + 1
              $global:spareCapacity = $True
              Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"
              break mainLoop
            }
          }
        }
        # Shut down hosts utilizing unnecessary resource
        $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }
        $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
        :shutdownLoop foreach ($ActiveHost in $ActiveHostsZeroSessions) {
          
          $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }

          # Ensure there is at least the offpeakMinimumNumberOfRDSH sessions available
          if ($NumberOfRunningHost -le $offpeakMinimumNumberOfRDSH) {
            Write-Output "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Off-Peak Minimum Number of Hosts = $offpeakMinimumNumberOfRDSH"
            break mainLoop
          }

          # Check for session capacity on other active hosts before shutting the free host down
          else {
            $ActiveHostsZeroSessions = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Where-Object { $_.Session -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True }
            :shutdownLoopTier2 foreach ($ActiveHost in $ActiveHostsZeroSessions) {
              $AllSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Sort-Object Status, Name
              foreach ($SessionHost in $AllSessionHosts) {
                if ($SessionHost.Status -eq "Available" -and ($SessionHost.Session -ge 0 -and $SessionHost.Session -lt $SessionHostLimit -and $SessionHost.AllowNewSession -eq $True)) {
                  if ($SessionHost.Name -ne $ActiveHost.Name) {
                    $ActiveHostName = $ActiveHost.Name
                    $ActiveHostName = $ActiveHostName.Split("/")[1]
                    $VMName = $ActiveHostName.Split(".")[0]
                    $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

                    # Check if the Session host is in maintenance
                    if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                      Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                      continue
                    }

                    Write-Output "Identified free Host $VMName with $($ActiveHost.Session) sessions that can be shut down to save resource"

                    # Ensure the running Azure VM is set as drain mode
                    try {
                      Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName -AllowNewSession:$False -ErrorAction SilentlyContinue
                    }
                    catch {
                      Write-Error "Unable to set 'Allow New Sessions' to False on Host $VMName with error: $($_.exception.message)"
                      exit
                    }
                    try {
                      Write-Output "Stopping host $VMName and waiting for it to complete ..."
                      Stop-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
                    }
                    catch {
                      Write-Error "Failed to stop host $VMName with error: $($_.exception.message)"
                      exit
                    }
                    # Check if the session host server is healthy before enable allowing new connections
                    if ($SessionHost.UpdateState -eq "Succeeded") {
                      # Ensure Azure VMs that are stopped have the allowing new connections state True
                      try {
                        Update-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName -AllowNewSession:$True -ErrorAction SilentlyContinue
                      }
                      catch {
                        Write-Error "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)"
                        exit
                      }
                    }
                    # Wait after shutting down ActiveHost until it's Status returns as Unavailable
                    $IsShutdownHostUnavailable = $false
                    while (!$IsShutdownHostUnavailable) {

                      $shutdownHost = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName -Name $ActiveHostName

                      if ($shutdownHost.Status -eq "Unavailable") {
                        $IsShutdownHostUnavailable = $true
                      }
                    }
                    # Decrement the number of running session host
                    $NumberOfRunningHost = $NumberOfRunningHost - 1
                    Write-Output "Current Number of Available Running Hosts is now: $NumberOfRunningHost"
                    break shutdownLoop
                  }
                }
              }     
            }
          }  
        }
      }
    }
  }
  # Get all available hosts and write to WVDAvailableHosts log
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$NumberOfRunningHost" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDAvailableHosts_CL" -TimeDifferenceInHours $TimeDifference
}

if (($global:spareCapacity -eq $False -or !$global:spareCapacity) -and ($global:capacityTrigger -eq $True)) { 
  Write-Warning "WARNING - All available running hosts have surpassed the Scale Factor of $SessionHostLimit and there are no additional hosts available to start"
}

if (($global:spareCapacity -eq $False -or !$global:spareCapacity) -and ($global:peakMinRDSHcapacityTrigger -eq $True)) { 
  Write-Warning "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) but there are no additional hosts available to start"
}

if (($global:spareCapacity -eq $False -or !$global:spareCapacity) -and ($global:offpeakMinRDSHcapacityTrigger -eq $True)) { 
  Write-Warning "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) but there are no additional hosts available to start"
}

Write-Output "Waiting for any outstanding jobs to complete..."
Get-Job | Wait-Job -Timeout $jobTimeout

$timedoutJobs = Get-Job -State Running
$failedJobs = Get-Job -State Failed

foreach ($job in $timedoutJobs) {
  Write-Warning "Error - The job $($job.Name) timed out"
}

foreach ($job in $failedJobs) {
  Write-Error "Error - The job $($job.Name) failed"
}

Write-Output "All job checks completed"
Write-Output "Ending Hosts Scale Optimization"
Write-Output "Writing to User/Host logs in Log Analytics"

# Get all active users and write to WVDUserSessions log
$CurrentActiveUsers = Get-AzWvdUserSession -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName | Select-Object UserPrincipalName, SessionHostName, SessionState | Sort-Object SessionHostName | Out-String
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$CurrentActiveUsers" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDUserSessions_CL" -TimeDifferenceInHours $TimeDifference

# Get all active hosts regardless of Maintenance Mode and write to WVDActiveHosts log
$RunningSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $resourceGroupName -HostPoolName $HostpoolName
$NumberOfRunningSessionHost = 0
foreach ($RunningSessionHost in $RunningSessionHosts) {

  if ($RunningSessionHost.Status -eq "Available") {
    $NumberOfRunningSessionHost = $NumberOfRunningSessionHost + 1
  }
}

$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$NumberOfRunningSessionHost" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDActiveHosts_CL" -TimeDifferenceInHours $TimeDifference


Write-Output "-------------------- Ending script --------------------"

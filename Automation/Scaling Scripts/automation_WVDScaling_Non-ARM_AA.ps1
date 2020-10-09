

param(
	[Parameter(mandatory = $false)]
	[object]$WebHookData
)
# If runbook was called from Webhook, WebhookData will not be null.
if ($WebHookData) {

	# Collect properties of WebhookData
	$WebhookName = $WebHookData.WebhookName
	$WebhookHeaders = $WebHookData.RequestHeader
	$WebhookBody = $WebHookData.RequestBody

	# Collect individual headers. Input converted from JSON.
	$From = $WebhookHeaders.From
	$Input = (ConvertFrom-Json -InputObject $WebhookBody)
}
else
{
	Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}

$AADTenantId = $Input.AADTenantId
$SubscriptionID = $Input.SubscriptionID
$TenantName = $Input.TenantName
$TenantGroupName = $Input.TenantGroupName
$ResourceGroupName = $Input.ResourceGroupName
$HostpoolName = $Input.hostPoolName
$BeginPeakTime = $Input.beginPeakTime
$EndPeakTime = $Input.endPeakTime
$TimeDifferenceInHours = $Input.TimeDifferenceInHours
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
$RDBrokerURL = $Input.RDBrokerURL
$AutomationAccountName = $Input.AutomationAccountName
$ConnectionAssetName = $Input.ConnectionAssetName

Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process -Force -Confirm:$false
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force -Confirm:$false

# Setting ErrorActionPreference to stop script execution when error occurs
$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Function for converting UTC to Local time
function ConvertUTCtoLocal {
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
function Add-LogEntry
{
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
      $LogData = $LogData + '"' + $Key.substring(0,$trim) + '":' + $sep + $LogMessageObj.Item($Key) + $sep + ','
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
$CurrentDateTime = ConvertUTCtoLocal -TimeDifferenceInHours $TimeDifference

# Collect the credentials from Azure Automation Account Assets
$Connection = Get-AutomationConnection -Name $ConnectionAssetName

# Authenticate to Azure 
Clear-AzContext -Force
$AZAuthentication = Connect-AzAccount -ApplicationId $Connection.ApplicationId -TenantId $AADTenantId -CertificateThumbprint $Connection.CertificateThumbprint -ServicePrincipal
if ($AZAuthentication -eq $null) {
  Write-Output "Failed to authenticate to Azure using a service principal: $($_.exception.message)"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to authenticate to Azure: $($_.exception.message)" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  exit
} 
else {
  $AzObj = $AZAuthentication | Out-String
  Write-Output "Authenticating using a service principal to Azure. Result: `n$AzObj"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Authenticating as service principal to Azure. Result: `n$AzObj" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

# Set the Azure context with Subscription
$AzContext = Set-AzContext -SubscriptionId $SubscriptionID
if ($AzContext -eq $null) {
  Write-Error "Please provide a valid subscription"
  exit
} 
else {
  $AzSubObj = $AzContext | Out-String
  Write-Output "Setting the Azure subscription. Result: `n$AzSubObj"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Sets the Azure subscription. Result: `n$AzSubObj" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

# Authenticate to WVD
try {
  $WVDAuthentication = Add-RdsAccount -DeploymentUrl $RDBrokerURL -ApplicationId $Connection.ApplicationId -CertificateThumbprint $Connection.CertificateThumbprint -AADTenantId $AadTenantId
}
catch {
  Write-Output "Failed to authenticate to WVD using a service principal: $($_.exception.message)"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to authenticate WVD: $($_.exception.message)" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  exit
}
$WVDObj = $WVDAuthentication | Out-String
Write-Output "Authenticating using a service principal to WVD. Result: `n$WVDObj"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Authenticating as service principal for WVD. Result: `n$WVDObj" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

# Set context to the appropriate tenant group
$CurrentTenantGroupName = (Get-RdsContext).TenantGroupName
if ($TenantGroupName -ne $CurrentTenantGroupName) {
  Write-Output "Switching to the $TenantGroupName context"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Switching to the $TenantGroupName context" }
	Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  Set-RdsContext -TenantGroupName $TenantGroupName
}

# Convert Datetime format
$BeginPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $BeginPeakTime)
$EndPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $EndPeakTime)

# Check the calculated end time is later than begin time in case of time zone
if ($EndPeakDateTime -lt $BeginPeakDateTime) {
  if ($CurrentDateTime -lt $EndPeakDateTime) { $BeginPeakDateTime = $BeginPeakDateTime.AddDays(-1) } else { $EndPeakDateTime = $EndPeakDateTime.AddDays(1) }
}

# Checking given hostpool name exists in Tenant
$HostpoolInfo = Get-RdsHostPool -TenantName $TenantName -Name $HostpoolName
if ($HostpoolInfo -eq $null) {
    Write-Output "Hostpoolname '$HostpoolName' does not exist in the tenant of '$TenantName'. Ensure that you have entered the correct values"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpoolname '$HostpoolName' does not exist in the tenant of '$TenantName'. Ensure that you have entered the correct values" }
		Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    exit
}	
        
# Compare beginpeaktime and endpeaktime hours and set up appropriate load balacing type based on PeakLoadBalancingType & OffPeakLoadBalancingType
if ($CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime) {

  if ($HostpoolInfo.LoadBalancerType -ne $PeakLoadBalancingType) {
      Write-Output "Changing Hostpool Load Balance Type to: $PeakLoadBalancingType Load Balancing"
      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Changing Hostpool Load Balance Type to: $PeakLoadBalancingType Load Balancing" }
      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

    if ($PeakLoadBalancingType -eq "DepthFirst") {                
      Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -DepthFirstLoadBalancer -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    else {
      Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -BreadthFirstLoadBalancer -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
    }
    Write-Output "Hostpool Load Balance Type in Peak Hours is: $PeakLoadBalancingType Load Balancing"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool Load Balance Type in Peak Hours is: $PeakLoadBalancingType Load Balancing" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  }
  # Compare MaxSessionLimit of hostpool to peakMaxSessions value and adjust if necessary
  if ($HostpoolInfo.MaxSessionLimit -ne $peakMaxSessions) {
    Write-Output "Changing Hostpool Peak MaxSessionLimit to: $peakMaxSessions"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Changing Hostpool Peak MaxSessionLimit to: $peakMaxSessions" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

    if ($PeakLoadBalancingType -eq "DepthFirst") {
      Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -DepthFirstLoadBalancer -MaxSessionLimit $peakMaxSessions
    }
    else {
      Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -BreadthFirstLoadBalancer -MaxSessionLimit $peakMaxSessions
    }
  }
}
else{
    if ($HostpoolInfo.LoadBalancerType -ne $OffPeakLoadBalancingType) {
        Write-Output "Changing Hostpool Load Balance Type to: $OffPeakLoadBalancingType Load Balancing"
        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Changing Hostpool Load Balance Type to: $OffPeakLoadBalancingType Load Balancing" }
        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
        
        if ($OffPeakLoadBalancingType -eq "DepthFirst") {                
            Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -DepthFirstLoadBalancer -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
        }
        else {
            Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -BreadthFirstLoadBalancer -MaxSessionLimit $HostpoolInfo.MaxSessionLimit
        }
        Write-Output "Hostpool Load Balance Type in Off-Peak Hours is: $OffPeakLoadBalancingType Load Balancing"
        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool Load Balance Type in Off-Peak Hours is: $OffPeakLoadBalancingType Load Balancing" }
        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    }
    # Compare MaxSessionLimit of hostpool to offpeakMaxSessions value and adjust if necessary
    if ($HostpoolInfo.MaxSessionLimit -ne $offpeakMaxSessions) {
      Write-Output "Changing Hostpool Off-Peak MaxSessionLimit to: $offpeakMaxSessions"
      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Changing Hostpool Off-Peak MaxSessionLimit to: $offpeakMaxSessions" }
      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

      if ($PeakLoadBalancingType -eq "DepthFirst") {
        Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -DepthFirstLoadBalancer -MaxSessionLimit $offpeakMaxSessions
      }
      else {
        Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -BreadthFirstLoadBalancer -MaxSessionLimit $offpeakMaxSessions
    }
  }
}

# Check for VM's with maintenance tag set to True & ensure connections are set as not allowed

$AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName

foreach ($SessionHost in $AllSessionHosts) {

  $SessionHostName = $SessionHost.SessionHostName | Out-String
  $VMName = $SessionHostName.Split(".")[0]
  $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

  if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
    Write-Output "The host $VMName is in Maintenance mode, so is not allowing any further connections"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The host $VMName is in Maintenance mode, so is not allowing any further connections" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $False -ErrorAction SilentlyContinue
  }
  else {
    Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
  }
}

Write-Output "Starting WVD Hosts Scale Optimization"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting WVD Tenant Hosts Scale Optimization" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

# Check the Hostpool Load Balancer type
$HostpoolInfo = Get-RdsHostPool -TenantName $tenantName -Name $hostPoolName
Write-Output "Hostpool Load Balance Type is: $($HostpoolInfo.LoadBalancerType)"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool Load Balance Type is: $($HostpoolInfo.LoadBalancerType)" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

# Check if it's peak hours
if ($CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime) {

  # Gathering hostpool maximum session and calculating Scalefactor for each host.										  
  $HostpoolMaxSessionLimit = $HostpoolInfo.MaxSessionLimit
  $ScaleFactorEachHost = $HostpoolMaxSessionLimit * $peakScaleFactor
  $SessionhostLimit = [math]::Floor($ScaleFactorEachHost)

  Write-Output "It is currently: Peak hours"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "It is currently: Peak hours" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  Write-Output "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  Write-Output "Checking current Host availability and workloads..."
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Checking current Host availability and workloads..." }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

  # Get all session hosts in the host pool
  $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
  if ($AllSessionHosts -eq $null) {
    Write-Output "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    exit
  }
  
  # Check the number of running session hosts
  $NumberOfRunningHost = 0
  foreach ($SessionHost in $AllSessionHosts) {

    $SessionHostName = $SessionHost.SessionHostName | Out-String
    $VMName = $SessionHostName.Split(".")[0]
    Write-Output "Host:$VMName, Current sessions:$($SessionHost.Sessions), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host:$VMName, Current sessions:$($SessionHost.Sessions), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

    if ($SessionHost.Status -eq "Available" -and $SessionHost.AllowNewSession -eq $True) {
      $NumberOfRunningHost = $NumberOfRunningHost + 1
    }
  }
  Write-Output "Current number of available running hosts: $NumberOfRunningHost"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of available running hosts: $NumberOfRunningHost" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  if ($NumberOfRunningHost -lt $peakMinimumNumberOfRDSH) {
    Write-Output "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) - Need to start additional hosts"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) - Need to start additional hosts" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    $global:peakMinRDSHcapacityTrigger = $True

    :peakMinStartupLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($NumberOfRunningHost -ge $peakMinimumNumberOfRDSH) {

        if ($minimumNumberFastScale -eq $True) {
          Write-Output "The number of available running hosts should soon equal the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The number of available running hosts should soon equal the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          break peakMinStartupLoop
        }
        else {
          Write-Output "The number of available running hosts ($NumberOfRunningHost) now equals the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The number of available running hosts ($NumberOfRunningHost) now equals the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH)" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          break peakMinStartupLoop
        }
      }

      # Check the session host status and if the session host is healthy before starting the host
      if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
        $SessionHostName = $SessionHost.SessionHostName | Out-String
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

        # Check to see if the Session host is in maintenance
        if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
          Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          continue
        }

        # Ensure Azure VMs that are stopped have the allowing new connections state set to True
        if ($SessionHost.AllowNewSession = $False) {
          try {
            Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
          }
          catch {
            Write-Output "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit 1
          }
        }
        if ($minimumNumberFastScale -eq $True) {

          # Start the Azure VM in Fast-Scale Mode for parallel processing
          try {
            Write-Output "Starting host $VMName in fast-scale mode..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName in fast-scale mode..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -AsJob

          }
          catch {
            Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to start host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit
          }
        }
        if ($minimumNumberFastScale -eq $False) {

          # Start the Azure VM
          try {
            Write-Output "Starting host $VMName and waiting for it to complete..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName and waiting for it to complete..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName

          }
          catch {
            Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to start host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit
          }
          # Wait for the sessionhost to become available
          $IsHostAvailable = $false
          while (!$IsHostAvailable) {

            $SessionHostStatus = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName

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
    $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName

    :mainLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($SessionHost.Sessions -le $HostpoolMaxSessionLimit -or $SessionHost.Sessions -gt $HostpoolMaxSessionLimit) {
        if ($SessionHost.Sessions -ge $SessionHostLimit) {
          $SessionHostName = $SessionHost.SessionHostName | Out-String
          $VMName = $SessionHostName.Split(".")[0]

          if (($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) -and ($global:capacityTrigger -eq $False -or !$global:capacityTrigger)) {
            Write-Output "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            $global:capacityTrigger = $True
          }

          :startupLoop  foreach ($SessionHost in $AllSessionHosts) {
            # Check the existing session hosts and session availability before starting another session host
            if ($SessionHost.Status -eq "Available" -and ($SessionHost.Sessions -ge 0 -and $SessionHost.Sessions -lt $SessionHostLimit) -and $SessionHost.AllowNewSession -eq $True) {
              $SessionHostName = $SessionHost.SessionHostName | Out-String
              $VMName = $SessionHostName.Split(".")[0]

              if ($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) {
                Write-Output "Host $VMName has spare capacity so don't need to start another host. Continuing now..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName has spare capacity so don't need to start another host. Continuing now..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                $global:exceededHostCapacity = $True
                $global:spareCapacity = $True
              }
              break startupLoop
            }

            # Check the session host status and if the session host is healthy before starting the host
            if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
              $SessionHostName = $SessionHost.SessionHostName | Out-String
              $VMName = $SessionHostName.Split(".")[0]
              $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

              # Check if the session host is in maintenance
              if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                continue
              }

              # Ensure Azure VMs that are stopped have the allowing new connections state set to True
              if ($SessionHost.AllowNewSession = $False) {
                try {
                  Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
                }
                catch {
                  Write-Output "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
                  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)" }
                  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                  exit 1
                }
              }

              # Start the Azure VM
              try {
                Write-Output "There is not enough spare capacity on other active hosts. A new host will now be started..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "There is not enough spare capacity on other active hosts. A new host will now be started..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                Write-Output "Starting host $VMName and waiting for it to complete..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName and waiting for it to complete..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                Start-AzVM -Name $VMName -ResourceGroupName $VMInfo.ResourceGroupName
              }
              catch {
                Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to start host $VMName with error: $($_.exception.message)" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                exit
              }

              # Wait for the sessionhost to become available
              $IsHostAvailable = $false
              while (!$IsHostAvailable) {

                $SessionHostStatus = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName

                if ($SessionHostStatus.Status -eq "Available") {
                  $IsHostAvailable = $true
                }
              }
              $NumberOfRunningHost = $NumberOfRunningHost + 1
              $global:spareCapacity = $True
              Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of Available Running Hosts is now: $NumberOfRunningHost" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              break mainLoop

            }
          }
        }
        # Shut down hosts utilizing unnecessary resource
        $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}
        $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
        :shutdownLoop foreach ($ActiveHost in $ActiveHostsZeroSessions) {
          
          $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}

          # Ensure there is at least the peakMinimumNumberOfRDSH sessions available
          if ($NumberOfRunningHost -le $peakMinimumNumberOfRDSH) {
            Write-Output "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Peak Minimum Number of RDSH = $peakMinimumNumberOfRDSH"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Peak Minimum Number of RDSH = $peakMinimumNumberOfRDSH" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            break mainLoop
          }

          # Check for session capacity on other active hosts before shutting the free host down
          else {
            $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}
            :shutdownLoopTier2 foreach ($ActiveHost in $ActiveHostsZeroSessions) {
              $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
              foreach ($SessionHost in $AllSessionHosts) {
                  if ($SessionHost.Status -eq "Available" -and ($SessionHost.Sessions -ge 0 -and $SessionHost.Sessions -lt $SessionHostLimit -and $SessionHost.AllowNewSession -eq $True)) {
                    if ($SessionHost.SessionHostName -ne $ActiveHost.SessionHostName) {
                      $ActiveHostName = $ActiveHost.SessionHostName | Out-String
                      $VMName = $ActiveHostName.Split(".")[0]
                      $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

                      # Check if the Session host is in maintenance
                      if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                        Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
                        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                      continue
                      }

                      Write-Output "Identified free host $VMName with $($ActiveHost.Sessions) sessions that can be shut down to save resource"
                      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Identified free host $VMName with $($ActiveHost.Sessions) sessions that can be shut down to save resource" }
                      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

                      # Ensure the running Azure VM is set as drain mode
                      try {
                        Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $ActiveHost.SessionHostName -AllowNewSession $False -ErrorAction SilentlyContinue
                      }
                      catch {
                        Write-Output "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)"
                        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)" }
                        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                      exit
                      }
                      try {
                        Write-Output "Stopping host $VMName and waiting for it to complete..."
                        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Stopping host $VMName and waiting for it to complete..." }
                        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                        Stop-AzureRmVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
                      }
                      catch {
                        Write-Output "Failed to stop host $VMName with error: $($_.exception.message)"
                        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to stop host $VMName with error: $($_.exception.message)" }
                        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                      exit
                      }
                      # Check if the session host server is healthy before enable allowing new connections
                      if ($SessionHost.UpdateState -eq "Succeeded") {
                        # Ensure Azure VMs that are stopped have the allowing new connections state True
                        try {
                          Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $ActiveHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
                        }
                        catch {
                          Write-Output "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
                          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)" }
                          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                          exit
                        }
                      }
                    # Decrement the number of running session hosts
                    $NumberOfRunningHost = $NumberOfRunningHost - 1
                    Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"
                    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of Available Running Hosts is now: $NumberOfRunningHost" }
                    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
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
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "It is currently: Off-Peak hours" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  Write-Output "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool Maximum Session Limit: $($HostpoolMaxSessionLimit)" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
  Write-Output "Checking current Host availability and workloads..."
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Checking current Host availability and workloads..." }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

  # Get all session hosts in the host pool
  $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
  if ($AllSessionHosts -eq $null) {
    Write-Output "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "No Session Hosts exist within the Hostpool '$HostpoolName'. Ensure that the Hostpool has hosts within it" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    exit
  }

  # Check the number of running session hosts
  $NumberOfRunningHost = 0
  foreach ($SessionHost in $AllSessionHosts) {

    $SessionHostName = $SessionHost.SessionHostName | Out-String
    $VMName = $SessionHostName.Split(".")[0]
    Write-Output "Host:$VMName, Current sessions:$($SessionHost.Sessions), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host:$VMName, Current sessions:$($SessionHost.Sessions), Status:$($SessionHost.Status), Allow New Sessions:$($SessionHost.AllowNewSession)" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

    if ($SessionHost.Status -eq "Available" -and $SessionHost.AllowNewSession -eq $True) {
      $NumberOfRunningHost = $NumberOfRunningHost + 1
    }
  }
  Write-Output "Current number of Available Running Hosts: $NumberOfRunningHost"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of Available Running Hosts: $NumberOfRunningHost" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

  # Check if user logoff is turned on in off peak
  if ($LimitSecondsToForceLogOffUser -ne 0) {
    Write-Output "Force Logging-off of Users in Off-Peak is enabled in order to consolidate sessions to minimal hosts. Checking if any resource can be saved..."
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Force Logging-off of Users in Off-Peak is enabled in order to consolidate sessions to minimal hosts. Checking if any resource can be saved..." }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

    if ($NumberOfRunningHost -gt $offpeakMinimumNumberOfRDSH) {
      Write-Output "The number of available running hosts is greater than the Off-Peak Minimum Number of RDSH sessions. Logging-off procedure will now be started..."
      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The number of available running hosts is greater than the Off-Peak Minimum Number of RDSH sessions. Logging-off procedure will now be started..." }
      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

      foreach ($SessionHost in $AllSessionHosts) {

        $SessionHostName = $SessionHost.SessionHostName | Out-String
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }
        
        if ($NumberOfRunningHost -gt $offpeakMinimumNumberOfRDSH) {
          if ($SessionHost.Status -eq "Available") {

            # Get the User sessions in the hostPool
            try {
              $HostPoolUserSessions = Get-RdsUserSession -TenantName $TenantName -HostPoolName $HostpoolName
            }
            catch {
              Write-Output "Failed to retrieve user sessions in hostPool $($HostpoolName) with error: $($_.exception.message)"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to retrieve user sessions in hostPool $($HostpoolName) with error: $($_.exception.message)" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              exit
            }
            $HostUserSessionCount = ($HostPoolUserSessions | Where-Object -FilterScript { $_.SessionHostName -eq $SessionHost }).Count
            Write-Output "Current sessions running on the host $VMName :$HostUserSessionCount"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current sessions running on the host $VMName :$HostUserSessionCount" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

            $ExistingSession = 0
            foreach ($Session in $HostPoolUserSessions) {
              if ($Session.SessionHostName -eq $SessionHost) {
                if ($LimitSecondsToForceLogOffUser -ne 0) {
                  # Notify user to log off their session
                  Write-Output "Sending log off message to users..."
                  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Sending log off message to users..." }
                  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                  try {
                    Send-RdsUserSessionMessage -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $Session.SessionHostName -SessionId $Session.SessionId -MessageTitle $LogOffMessageTitle -MessageBody "$($LogOffMessageBody) You will logged off in $($LimitSecondsToForceLogOffUser) seconds." -NoUserPrompt
                  }
                  catch {
                    Write-Output "Failed to send message to user with error: $($_.exception.message)"
                    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to send message to user with error: $($_.exception.message)" }
                    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                    exit
                  }
                }
                $ExistingSession = $ExistingSession + 1
              }
            }
            # Wait for n seconds to log off user
            Write-Output "Waiting for $LimitSecondsToForceLogOffUser seconds before logging off users..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Waiting for $LimitSecondsToForceLogOffUser seconds before logging off users..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            Start-Sleep -Seconds $LimitSecondsToForceLogOffUser
            if ($LimitSecondsToForceLogOffUser -ne 0) {
              # Force Users to log off
              Write-Output "Forcing users to log off now..."
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Forcing users to log off now..." }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              try {
                $HostPoolUserSessions = Get-RdsUserSession -TenantName $TenantName -HostPoolName $HostpoolName
              }
              catch {
                Write-Output "Failed to retrieve list of user sessions in HostPool $HostpoolName with error: $($_.exception.message)"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to retrieve list of user sessions in HostPool $HostpoolName with error: $($_.exception.message)" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                exit
              }
              foreach ($Session in $HostPoolUserSessions) {
                if ($Session.SessionHostName -eq $SessionHost) {
                  # Log off user
                  try {
                    Invoke-RdsUserSessionLogoff -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $Session.SessionHostName -SessionId $Session.SessionId -NoUserPrompt
                    $ExistingSession = $ExistingSession - 1
                  }
                  catch {
                    Write-Output "Failed to log off user with error: $($_.exception.message)"
                    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to log off user with error: $($_.exception.message)" }
                    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                    exit
                  }
                }
              }
            }

            $SessionHostName = $SessionHost.SessionHostName | Out-String
            $VMName = $SessionHostName.Split(".")[0]
            $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }
            
            # Check to see if the Session host is in maintenance
            if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
              Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              $NumberOfRunningHost = $NumberOfRunningHost - 1
              continue
            }

            # Ensure the running Azure VM is set as drain mode
            try {
              Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $False -ErrorAction SilentlyContinue
            }
            catch {
              Write-Output "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to False on host $VMName with error: $($_.exception.message)" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit
            }

            # Check the session count before shutting down the VM
            if ($SessionHost.Sessions -eq 0) {
              Write-Output "Host $VMName now has 0 sessions"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName now has 0 sessions" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              # Shutdown the Azure VM
              try {
                Write-Output "Stopping host $VMName and waiting for it to complete..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Stopping host $VMName and waiting for it to complete..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                Stop-AzureRmVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
              }
              catch {
                Write-Output "Failed to stop host $VMName with error: $($_.exception.message)"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to stop host $VMName with error: $($_.exception.message)" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                exit
              }
            }

            # Check if the session host server is healthy before enable allowing new connections
            if ($SessionHost.UpdateState -eq "Succeeded") {
              # Ensure Azure VMs that are stopped have the allowing new connections state True
              try {
                Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
              }
              catch {
                Write-Output "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on host $VMName with error: $($_.exception.message)" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                exit 1
              }
            }
            # Decrement the number of running session host
            $NumberOfRunningHost = $NumberOfRunningHost - 1
          }
        }
      }
    }
  }
  if ($NumberOfRunningHost -lt $offpeakMinimumNumberOfRDSH) {
    Write-Output "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) - Need to start additional hosts"
    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) - Need to start additional hosts" }
    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
    $global:offpeakMinRDSHcapacityTrigger = $True

    :offpeakMinStartupLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($NumberOfRunningHost -ge $offpeakMinimumNumberOfRDSH) {

        if ($minimumNumberFastScale -eq $True) {
          Write-Output "The number of available running hosts should soon equal the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The number of available running hosts should soon equal the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          break offpeakMinStartupLoop
        }
        else {
          Write-Output "The number of available running hosts ($NumberOfRunningHost) now equals the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "The number of available running hosts ($NumberOfRunningHost) now equals the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH)" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          break offpeakMinStartupLoop
        }
      }

      # Check the session host status and if the session host is healthy before starting the host
      if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
        $SessionHostName = $SessionHost.SessionHostName | Out-String
        $VMName = $SessionHostName.Split(".")[0]
        $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

        # Check to see if the Session host is in maintenance
        if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
          Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
          $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
          Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
          continue
        }

        # Ensure Azure VMs that are stopped have the allowing new connections state set to True
        if ($SessionHost.AllowNewSession = $False) {
          try {
            Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
          }
          catch {
            Write-Output "Unable to set it to allow connections on host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set it to allow connections on host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit 1
          }
        }
        if ($minimumNumberFastScale -eq $True) {

          # Start the Azure VM in Fast-Scale Mode for parallel processing
          try {
            Write-Output "Starting host $VMName in fast-scale mode..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName in fast-scale mode..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -AsJob

          }
          catch {
            Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to start host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit
          }
        }
        if ($minimumNumberFastScale -eq $False) {

          # Start the Azure VM
          try {
            Write-Output "Starting host $VMName and waiting for it to complete..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName and waiting for it to complete..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            Start-AzVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName

          }
          catch {
            Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to start host $VMName with error: $($_.exception.message)" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            exit
          }
          # Wait for the sessionhost to become available
          $IsHostAvailable = $false
          while (!$IsHostAvailable) {

            $SessionHostStatus = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName

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
    $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName

    :mainLoop foreach ($SessionHost in $AllSessionHosts) {

      if ($SessionHost.Sessions -le $HostpoolMaxSessionLimit -or $SessionHost.Sessions -gt $HostpoolMaxSessionLimit) {
        if ($SessionHost.Sessions -ge $SessionHostLimit) {
          $SessionHostName = $SessionHost.SessionHostName | Out-String
          $VMName = $SessionHostName.Split(".")[0]

          if (($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) -and ($global:capacityTrigger -eq $False -or !$global:capacityTrigger)) {
            Write-Output "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..."
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "One or more hosts have surpassed the Scale Factor of $SessionHostLimit. Checking other active host capacities now..." }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            $global:capacityTrigger = $True
          }

          :startupLoop  foreach ($SessionHost in $AllSessionHosts) {
            # Check the existing session hosts and session availability before starting another session host
            if ($SessionHost.Status -eq "Available" -and ($SessionHost.Sessions -ge 0 -and $SessionHost.Sessions -lt $SessionHostLimit) -and $SessionHost.AllowNewSession -eq $True) {
              $SessionHostName = $SessionHost.SessionHostName | Out-String
              $VMName = $SessionHostName.Split(".")[0]

              if ($global:exceededHostCapacity -eq $False -or !$global:exceededHostCapacity) {
                Write-Output "Host $VMName has spare capacity so don't need to start another host. Continuing now..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName has spare capacity so don't need to start another host. Continuing now..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                $global:exceededHostCapacity = $True
                $global:spareCapacity = $True
              }
              break startupLoop
            }

            # Check the session host status and if the session host is healthy before starting the host
            if (($SessionHost.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable") -and ($SessionHost.UpdateState -eq "Succeeded")) {
              $SessionHostName = $SessionHost.SessionHostName | Out-String
              $VMName = $SessionHostName.Split(".")[0]
              $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

              # Check if the session host is in maintenance
              if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                continue
              }

              # Ensure Azure VMs that are stopped have the allowing new connections state set to True
              if ($SessionHost.AllowNewSession = $False) {
                try {
                  Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
                }
                catch {
                  Write-Output "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)"
                  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)" }
                  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                  exit 1
                }
              }

              # Start the Azure VM
              try {
                Write-Output "There is not enough spare capacity on other active hosts. A new host will now be started..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "There is not enough spare capacity on other active hosts. A new host will now be started..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                Write-Output "Starting host $VMName and waiting for it to complete..."
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Starting host $VMName and waiting for it to complete..." }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                Start-AzVM -Name $VMName -ResourceGroupName $VMInfo.ResourceGroupName
              }
              catch {
                Write-Output "Failed to start host $VMName with error: $($_.exception.message)"
                $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of Available Running Hosts is now: $NumberOfRunningHost" }
                Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                exit
              }
              # Wait for the sessionhost to become available
              $IsHostAvailable = $false
              while (!$IsHostAvailable) {

                $SessionHostStatus = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost.SessionHostName

                if ($SessionHostStatus.Status -eq "Available") {
                  $IsHostAvailable = $true
                }
              }
              $NumberOfRunningHost = $NumberOfRunningHost + 1
              $global:spareCapacity = $True
              Write-Output "Current number of Available Running Hosts is now: $NumberOfRunningHost"
              $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current number of Available Running Hosts is now: $NumberOfRunningHost" }
              Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
              break mainLoop
            }
          }
        }
        # Shut down hosts utilizing unnecessary resource
        $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}
        $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
        :shutdownLoop foreach ($ActiveHost in $ActiveHostsZeroSessions) {
          
          $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}

          # Ensure there is at least the offpeakMinimumNumberOfRDSH sessions available
          if ($NumberOfRunningHost -le $offpeakMinimumNumberOfRDSH) {
            Write-Output "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Off-Peak Minimum Number of RDSH = $offpeakMinimumNumberOfRDSH"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Found no available resource to save as the number of Available Running Hosts = $NumberOfRunningHost and the specified Off-Peak Minimum Number of RDSH = $offpeakMinimumNumberOfRDSH" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
            break mainLoop
          }

          # Check for session capacity on other active hosts before shutting the free host down
          else {
            $ActiveHostsZeroSessions = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object {$_.Sessions -eq 0 -and $_.Status -eq "Available" -and $_.AllowNewSession -eq $True}
            :shutdownLoopTier2 foreach ($ActiveHost in $ActiveHostsZeroSessions) {
              $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Status,SessionHostName
              foreach ($SessionHost in $AllSessionHosts) {
                if ($SessionHost.Status -eq "Available" -and ($SessionHost.Sessions -ge 0 -and $SessionHost.Sessions -lt $SessionHostLimit -and $SessionHost.AllowNewSession -eq $True)) {
                  if ($SessionHost.SessionHostName -ne $ActiveHost.SessionHostName) {
                    $ActiveHostName = $ActiveHost.SessionHostName | Out-String
                    $VMName = $ActiveHostName.Split(".")[0]
                    $VmInfo = Get-AzVM | Where-Object { $_.Name -eq $VMName }

                    # Check if the Session host is in maintenance
                    if ($VMInfo.Tags.ContainsKey($MaintenanceTagName) -and $VMInfo.Tags.ContainsValue($True)) {
                      Write-Output "Host $VMName is in Maintenance mode, so this host will be skipped"
                      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Host $VMName is in Maintenance mode, so this host will be skipped" }
                      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                    continue
                    }

                    Write-Output "Identified free Host $VMName with $($ActiveHost.Sessions) sessions that can be shut down to save resource"
                    $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Identified free Host $VMName with $($ActiveHost.Sessions) sessions that can be shut down to save resource" }
                    Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

                    # Ensure the running Azure VM is set as drain mode
                    try {
                      Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $ActiveHost.SessionHostName -AllowNewSession $False -ErrorAction SilentlyContinue
                    }
                    catch {
                      Write-Output "Unable to set 'Allow New Sessions' to False on Host $VMName with error: $($_.exception.message)"
                      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to False on Host $VMName with error: $($_.exception.message)" }
                      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                    exit
                    }
                    try {
                      Write-Output "Stopping host $VMName and waiting for it to complete ..."
                      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Stopping host $VMName and waiting for it to complete ..." }
                      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                      Stop-AzureRmVM -Name $VMName -ResourceGroupName $VmInfo.ResourceGroupName -Force
                    }
                    catch {
                      Write-Output "Failed to stop host $VMName with error: $($_.exception.message)"
                      $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to stop host $VMName with error: $($_.exception.message)" }
                      Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                    exit
                    }
                    # Check if the session host server is healthy before enable allowing new connections
                    if ($SessionHost.UpdateState -eq "Succeeded") {
                      # Ensure Azure VMs that are stopped have the allowing new connections state True
                      try {
                        Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $ActiveHost.SessionHostName -AllowNewSession $True -ErrorAction SilentlyContinue
                      }
                      catch {
                        Write-Output "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)"
                        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Unable to set 'Allow New Sessions' to True on Host $VMName with error: $($_.exception.message)" }
                        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
                        exit
                      }
                    }
                  # Decrement the number of running session host
                  $NumberOfRunningHost = $NumberOfRunningHost - 1
                  Write-Output "Current Number of Available Running Hosts is now: $NumberOfRunningHost"
                  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Current Number of Available Running Hosts is now: $NumberOfRunningHost" }
                  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
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
  Write-Output "WARNING - All available running hosts have surpassed the Scale Factor of $SessionHostLimit and there are no additional hosts available to start"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "WARNING - All available running hosts have surpassed the Scale Factor of $SessionHostLimit and there are no additional hosts available to start" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

if (($global:spareCapacity -eq $False -or !$global:spareCapacity) -and ($global:peakMinRDSHcapacityTrigger -eq $True)) { 
  Write-Output "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) but there are no additional hosts available to start"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Peak Minimum Number of RDSH ($peakMinimumNumberOfRDSH) but there are no additional hosts available to start" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

if (($global:spareCapacity -eq $False -or !$global:spareCapacity) -and ($global:offpeakMinRDSHcapacityTrigger -eq $True)) { 
  Write-Output "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) but there are no additional hosts available to start"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "WARNING - Current number of available running hosts ($NumberOfRunningHost) is less than the specified Off-Peak Minimum Number of RDSH ($offpeakMinimumNumberOfRDSH) but there are no additional hosts available to start" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

Write-Output "Waiting for any outstanding jobs to complete..."
Get-Job | Wait-Job -Timeout $jobTimeout

$timedoutJobs = Get-Job -State Running
$failedJobs = Get-Job -State Failed

foreach ($job in $timedoutJobs) {
  Write-Output "Error - The job $($job.Name) timed out"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Error - The job $($job.Name) timed out" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

foreach ($job in $failedJobs) {
  Write-Output "Error - The job $($job.Name) failed"
  $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Error - The job $($job.Name) failed" }
  Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
}

Write-Output "All job checks completed"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "All job checks completed" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
Write-Output "Ending WVD Tenant Scale Optimization"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Ending WVD Tenant Scale Optimization" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference
Write-Output "Writing to User/Host logs"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Writing to User/Host logs" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

# Get all active users and write to WVDUserSessions log
$CurrentActiveUsers = Get-RdsUserSession  -TenantName $TenantName -HostPoolName $HostpoolName | Select-Object UserPrincipalName, SessionHostName, SessionState | Sort-Object SessionHostName | Out-String
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$CurrentActiveUsers" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDUserSessions_CL" -TimeDifferenceInHours $TimeDifference

# Get all active hosts regardless of Maintenance Mode and write to WVDActiveHosts log
$RunningSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName
$NumberOfRunningSessionHost = 0
foreach ($RunningSessionHost in $RunningSessionHosts) {

  if ($RunningSessionHost.Status -eq "Available") {
    $NumberOfRunningSessionHost = $NumberOfRunningSessionHost + 1
  }
}

$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "$NumberOfRunningSessionHost" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDActiveHosts_CL" -TimeDifferenceInHours $TimeDifference


Write-Output "-------------------- Ending script --------------------"
$LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "-------------------- Ending script --------------------" }
Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -LogType "WVDScaling_CL" -TimeDifferenceInHours $TimeDifference

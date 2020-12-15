#Disclaimer
The sample scripts are not supported under any Microsoft standard support program or service. 
The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims 
all implied warranties including, without limitation, any implied warranties of merchantability or
of fitness for a particular purpose. The entire risk arising out of the use or performance of the 
sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or 
anyone else involved in the creation, production, or delivery of the scripts be liable for any 
damages whatsoever (including, without limitation, damages for loss of business profits, 
business interruption, loss of business information, or other pecuniary loss) arising out of 
the use of or inability to use the sample scripts or documentation, even if Microsoft has been 
advised of the possibility of such damages.#


#Add SharePoint PowerShell Snap-In
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {  
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"  
}

## Changes - Added Worflow Engine Status,Search Crawl History
### Changes 9/1/2017 - SQL Backup history and Upgrade Status have been added but are not rendered in the report for now.
### Changes 11/22/2017 - Added Health Analyzer and OWA to report

 
 
[xml]$xml = get-content D:\scripts\variables.xml

 
#########################################  SharePoint Status Results   ###########################################

#Disabled Status
$DisabledStatus='<td>Disabled</td>'
$DisabledColorStyle='<td style="background-color:#C55A11 !important;color:white !important" >Disabled</td>'

#Offline Status
$OfflineStatus='<td>Offline</td>'
$OfflineColorStyle='<td style="background-color:#FF0000 !important;color:white !important" >Offline</td>'

#Unprovisioning Status
$UnprovisioningStatus='<td>Unprovisioning</td>'
$UnprovisioningColorStyle='<td style="background-color:#FFC000 !important;color:white !important" >Unprovisioning</td>'

#Provisioning Status
$ProvisioningStatus='<td>Provisioning</td>'
$ProvisioningColorStyle='<td style="background-color:#92D050 !important;color:white !important" >Provisioning</td>'

#Upgrading Status
$UpgradingStatus='<td>Upgrading</td>'
$UpgradingColorStyle='<td style="background-color:#00B0F0 !important;color:white !important" >Upgrading</td>'

#########################################  System Status Result   #################################################

#Error Status
$ErrorStatus='<td>Error</td>'
$ErrorColorStyle='<td style="background-color:#FF0000 !important;color:white !important" >Error</td>'

#Degraded Status
$DegradedStatus='<td>Degraded</td>'
$DegradedColorStyle='<td style="background-color:#C00000 !important;color:white !important" >Degraded</td>'

#Unknown Status
$UnknownStatus='<td>Unknown</td>'
$UnknownColorStyle='<td style="background-color:#7030A0 !important;color:white !important" >Unknown</td>'

#Pred Fail Status
$PredFailStatus='<td>Pred Fail</td>'
$PredFailColorStyle='<td style="background-color:#CC0066 !important;color:white !important" >Pred Fail</td>'

#Starting Status
$StartingStatus='<td>Starting</td>'
$StartingColorStyle='<td style="background-color:#92D050 !important;color:white !important" >Starting</td>'

#Stopping Status
$StoppingStatus='<td>Stopping</td>'
$StoppingColorStyle='<td style="background-color:#C55A11 !important;color:white !important" >Stopping</td>'

#Service Status
$ServiceStatus='<td>Service</td>'
$ServiceColorStyle='<td style="background-color:#548235 !important;color:white !important" >Service</td>'

#Stressed Status
$StressedStatus='<td>Stressed</td>'
$StressedColorStyle='<td style="background-color:#CC0066 !important;color:white !important" >Stressed</td>'

#NonRecover Status
$NonRecoverStatus='<td>NonRecover</td>'
$NonRecoverColorStyle='<td style="background-color:#FF0000 !important;color:white !important" >NonRecover</td>'

#NoContact Status
$NoContactStatus='<td>NoContact</td>'
$NoContactColorStyle='<td style="background-color:#FFC000 !important;color:white !important" >NoContact</td>'

#LostCom Status
$LostComStatus='<td>LostCom</td>'
$LostComColorStyle='<td style="background-color:#C55A11 !important;color:white !important" >LostCom</td>'

######################################### Windows Service Status   ##############################################

#Pending Status
$PendingStatus='<td>Pending</td>'
$PendingColorStyle='<td style="background-color:#92D050 !important;color:white !important" >Pending</td>'

#Paused Status
$PausedStatus='<td>Paused</td>'
$PausedColorStyle='<td style="background-color:#C55A11 !important;color:white !important" >Paused</td>'

#Stopped Status
$StoppedStatus='<td>Stopped</td>'
$StoppedColorStyle='<td style="background-color:#FF0000 !important;color:white !important" >Stopped</td>'

########################################  IIS App Pool Status  ##################################################

#starting Status
$startingPoolStatus='<td>1</td>'
$startingPoolColorStyle='<td style="background-color:#92D050 !important;color:white !important" >1</td>'

#stopping Status
$stoppingPoolStatus='<td>3</td>'
$stoppingPoolColorStyle='<td style="background-color:#C55A11 !important;color:white !important" >3</td>'

#Stopped Status
$StoppedPoolStatus='<td>4</td>'
$StoppedPoolColorStyle='<td style="background-color:#FF0000 !important;color:white !important" >4</td>'

########################################   Set Status Color Code Functions  #######################################
#region    Set colors
function SetSystemStatusColor
{
  $SystemStatushtml = $args[0]
     
  #Error
  $SystemStatushtml=$SystemStatushtml -replace $ErrorStatus,$ErrorColorStyle
   
  #Degraded
  $SystemStatushtml=$SystemStatushtml -replace $DegradedStatus,$DegradedColorStyle
   
  #Unknown
  $SystemStatushtml=$SystemStatushtml -replace $UnknownStatus,$UnknownColorStyle
   
  #Pred Fail
  $SystemStatushtml=$SystemStatushtml -replace $PredFailStatus,$PredFailColorStyle
   
  #Starting
  $SystemStatushtml=$SystemStatushtml -replace $StartingStatus,$StartingColorStyle
   
  #Stopping
  $SystemStatushtml=$SystemStatushtml -replace $StoppingStatus,$StoppingColorStyle
   
  #Service
  $SystemStatushtml=$SystemStatushtml -replace $ServiceStatus,$ServiceColorStyle
   
  #Stressed
  $SystemStatushtml=$SystemStatushtml -replace $StressedStatus,$StressedColorStyle
   
  #NonRecover
  $SystemStatushtml=$SystemStatushtml -replace $NonRecoverStatus,$NonRecoverColorStyle
   
  #NoContact
  $SystemStatushtml=$SystemStatushtml -replace $NoContactStatus,$NoContactColorStyle
   
  #LostCom 
  $SystemStatushtml=$SystemStatushtml -replace $LostComStatus,$LostComColorStyle
     
  $SystemStatushtml
}

function SetSharePointStatusColor
{
  $SharePointStatushtml = $args[0]
     
  #Disabled
  $SharePointStatushtml=$SharePointStatushtml -replace $DisabledStatus,$DisabledColorStyle
   
  #Offline
  $SharePointStatushtml=$SharePointStatushtml -replace $OfflineStatus,$OfflineColorStyle
   
  #Unprovisioning
  $SharePointStatushtml=$SharePointStatushtml -replace $UnprovisioningStatus,$UnprovisioningColorStyle
   
  #Provisioning
  $SharePointStatushtml=$SharePointStatushtml -replace $ProvisioningStatus,$ProvisioningColorStyle
   
  #Upgrading
  $SharePointStatushtml=$SharePointStatushtml -replace $UpgradingStatus,$UpgradingColorStyle
   
  $SharePointStatushtml
}

function SetWinServiceStatusColor
{
  $WinServiceStatushtml = $args[0]
       
  #Pending
  $WinServiceStatushtml=$WinServiceStatushtml -replace $PendingStatus,$PendingColorStyle
   
  #Paused
  $WinServiceStatushtml=$WinServiceStatushtml -replace $PausedStatus,$PausedColorStyle
   
  #Stopped
  $WinServiceStatushtml=$WinServiceStatushtml -replace $StoppedStatus,$StoppedColorStyle  
     
  $WinServiceStatushtml
}

function SetAppPoolStatusColor
{
  $AppPoolStatushtml = $args[0]
       
  #1 - starting
  $AppPoolStatushtml=$AppPoolStatushtml -replace $startingPoolStatus,$startingPoolColorStyle
   
  #3 - stopping
  $AppPoolStatushtml=$AppPoolStatushtml -replace $stoppingPoolStatus,$stoppingPoolColorStyle
   
  #4 - stopped
  $AppPoolStatushtml=$AppPoolStatushtml -replace $StoppedPoolStatus,$StoppedPoolColorStyle  
     
  $AppPoolStatushtml
}

#endregion

#region SharePoint Servers
########################################  Get SharePoint Servers  ###################################################

#Define servers
$ServersinFarm = Get-SPServer |  where {$_.Role -ne "Invalid"}
#endregion

#region    CPU, Memory and Disk Utilization of Servers

#######################################  CPU Utilization Status Details  ##################################################
$CPUDataCol = $object = Get-WmiObject -Class Win32_Processor -ComputerName $ServersinFarm.name | Select @{Name='Server Name';Expression={$_.SystemName}}, LoadPercentage ,

Status   
 
$CPU = $CPUDataCol | ConvertTo-Html -Fragment
$CPU = SetSystemStatusColor $CPU
write-host Received CPU Utilization Status Details -foregroundcolor "green"
#endregion


#######################################  Memory Utilization Status Details  #################################################

$MemoryCol = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ServersinFarm.name | Select @{Name='Server Name';Expression={$_.CsName}} ,  @{Name="Total Memory
(GB)";Expression={[math]::round($_.TotalVisibleMemorySize/1mb,2)}} ,@{Name="Free Memory(GB)";Expression={[math]::round($_.FreePhysicalMemory/1mb,2)}}  , @{Name = "Memory Usage
(%)"; Expression ={“{0:N1}” -f ((([math]::round($_.TotalVisibleMemorySize,2) - [math]::round($_.FreePhysicalMemory,2))*100)/[math]::round($_.TotalVisibleMemorySize,2))}}

$Memory = $MemoryCol | ConvertTo-Html -Fragment
$Memory = SetSystemStatusColor $Memory
write-host Received Memory Utilization Status Details -foregroundcolor "green"

#######################################  Disk Utilization Status Details  ##################################################

$DiskCol = $object3 = Get-WmiObject -Class Win32_LogicalDisk -filter DriveType=3 -ComputerName $ServersinFarm.name | Select @{Name='Server Name';Expression={$_.SystemName}} ,
DeviceID , @{Name="Size(GB)";Expression={"{0:N1}" -f($_.size/1gb)}}, @{Name="Free  Space(GB)";Expression={"{0:N1}" -f($_.freespace/1gb)}} , @{Name = "Free Space(%)"; Expression = {“{0:N2}” -f (((($_.freespace))/ $_.size)*100) }}

 
$Disk = $DiskCol | ConvertTo-Html -Fragment

write-host Received SharePoint Server Disk Status -foregroundcolor "green"

#endregion

#region SharePoint Farm and Servers Information

########################################  SharePoint Farm Configuration Status  ######################################

$SPFarm = Get-SPFarm | select Name ,NeedsUpgrade , Status , @{Name='SharePoint Verison Number';Expression={$_.BuildVersion}} |ConvertTo-Html -Fragment
$SPFarm = SetSharePointStatusColor $SPFarm
write-host Received SharePoint Farm Status Details -foregroundcolor "green"

 
########################################  SharePoint Server Status Details  ###########################################

$SPServersInfo = Get-SPServer | Select Name,Role,Status,CanUpgrade,NeedsUpgrade | ConvertTo-Html -Fragment
$SPServersInfo = SetSharePointStatusColor $SPServersInfo
write-host Received SharePoint Server Status Farm Details -foregroundcolor "green"

########################################  Server in Farm  ##################################################

#SharePoint servers Status Details 
 
$SPServerDetails = Get-SPServer | select Address , Status , Farm | ConvertTo-Html -Fragment
$SPServerDetails = SetSystemStatusColor $SPServerDetails

write-host Received All SharePoint Server Farm Info -foregroundcolor "green"

#endregion
########################################  Web Template usage in farm  ##################################################

#Web Template usage in farm 
$Webt = Get-SPSite -Limit All | ? { -not $_.IsReadLocked } | Get-SPWeb -Limit All | GROUP WebTemplate | FT Name, Count -AutoSize| ConvertTo-Html -Fragment



#region Web Applications Status

########################################  SharePoint Web Application Status Details  ###################################

$WebApplication = Get-SPWebApplication | Select Name , Url , Status | ConvertTo-Html -Fragment
$WebApplication = SetSharePointStatusColor $WebApplication
write-host Received Web Application Status Details -foregroundcolor "green"

 
 
########################################  Site Collection Response Time Details  ##################################################

 
$ExemptedWebApps = $xml.VARIABLES.excludewebapp

Try{

    $CustomResult=@()
    $WebApps = Get-SPWebApplication
    #Run for each web application
    foreach($WebApp in $WebApps){

        #skips exempted web applications
        if($ExemptedWebApps -match $($WebApp.url)){
            $message += "<p>Skipping web application $($WebApp.Url)<p>"
            continue
        }
    
        $message +=  "Web Application: $($WebApp.url)<br/>"

        #Run for each site collection
        $Sites = Get-SPSite -WebApplication $WebApp -Limit All
        foreach($Site in $Sites){
       
      
            #Run for each site
            $Webs = Get-SPWeb -Site $Site -Limit All
            foreach($Web in $Webs){
                $uri = $Web.Url
                $time = try{
                            $request = $null
                
                            ## Request the URI, and measure how long the response took.
                            $result1 = Measure-Command { $request = Invoke-WebRequest -Uri $uri -UseDefaultCredentials }
                            $result1.TotalMilliseconds
                        } 
                        catch{
                            # If the request generated an exception
                            $request = $_.Exception.Response
                            $time = -1
                        }
                $CustomResult += [PSCustomObject] @{
                    Time = Get-Date;
                    Uri = $uri;
                    StatusCode = [int] $request.StatusCode;
                    StatusDescription = $request.StatusDescription;
                    ResponseLength = $request.RawContentLength;
                    TimeTaken =  $time; 
                }

             }
        
        } 
    }

   
    $WebResponseTime += (($CustomResult | Select-Object    @{Expression={$_.Uri};Label="URL"},`
                                                             @{Expression={$_.StatusDescription};Label="Status"},`
                                                             @{Expression={$_.StatusCode};Label="Status Code"},`
                                                             @{Expression={$_.TimeTaken};Label="Response Time(ms)"},`
                                                             @{Expression={$_.Time};Label="Time Checked"}`
                                         | ConvertTo-HTML -Fragment) `
)
                                                
 
 
}

 Catch{
    Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Error "Exception Message: $($_.Exception.Message)"   
}

Write-Host Receceived Site Collection Status -foregroundcolor "green"
#endregion

 
 
######################################################################Search Crawl History############################################
#region Search Crawl History

$ssa = Get-SPEnterpriseSearchServiceApplication | Where-Object {$_.Name -like "*Search*"}
$ContentSources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $SSA #| where {$_.Name -eq $ContentSourceName}
$crawl = $ContentSources | Select Name, SuccessCount, WarningCount, ErrorCount,CrawlStarted,CrawlCompleted,  @{label="CrawlDuration";expression={$_.CrawlCompleted - $_.CrawlStarted}} | ConvertTo-Html -Fragment

write-host Received SharePoint Crawl History -foregroundcolor "green"

#endregion
#region Usage and Users Information

  
########################################  Windows sharePoint Services Status Details  ##################################

$SPServices = invoke-command -computername $ServersinFarm.Name {get-Service -Name 'SPAdminV4' , 'SPTimerV4' , 'SPTraceV4' , 'SPUserCodeV4' , 'SMTPSVC' , 'OSearch15' , 'W3SVC' ,'IISADMIN' ,'FIMSynchronizationService', 'FIMService','AppFabricCachingService','WorkflowServiceBackend' | Select  @{Label ="ServiceS"; Expression = {$_.DisplayName}} ,Status, @{Name ="Server Name";Expression = {$_.PSComputerName}} -Exclude PSShowComputerName,RunspaceID } | Sort-Object Status
$SPServices =  $SPServices | Select  Services , Status,  @{Name ="Server Name"; Expression = {$_.PSComputerName}}|ConvertTo-Html -Fragment

$SPServices = SetWinServiceStatusColor $SPServices
write-host Received SharePoint Windows Service Status -foregroundcolor "green"

 
#######################################  IIS Application Pool Status Details  ##########################################

function Get-AppPoolStatus
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$false)][string]$ApplicationPoolName
    )
    
    begin
    {
        $filter = { $true }
        [Reflection.Assembly]::LoadWithPartialName('Microsoft.Web.Administration') | Out-Null
    }
    process
    {

        if( $ApplicationPoolName )
        {
            $filter =  { $_.Name -eq $ApplicationPoolName }
        }
        
        $serverManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote( $ComputerName )
        
        if( $serverManager.ApplicationPools | ? $filter )
        {
            foreach( $applicationPool in $serverManager.ApplicationPools | ? $filter )
            {
                [PSCustomObject] @{
                    ComputerName          = $ComputerName
                    ApplicationPoolName   = $applicationPool.Name
                    ApplicationPoolStatus = $applicationPool.State
                }
            }
        }
        else
        {
            [PSCustomObject] @{
                ComputerName          = $ComputerName
                ApplicationPoolName   = $ApplicationPoolName
                ApplicationPoolStatus = "Not Found"
            }        
        }
    }
    end
    {
    }
}

$results = @()

foreach( $serviceApplicationPool in Get-SPServiceApplicationPool )
{
    foreach( $server in Get-SPServer |? { $_.Role -ne "Invalid" } )
    {
        # get the named app pools
        $results += Get-AppPoolStatus -ComputerName $server.Name -ApplicationPoolName $serviceApplicationPool.Name
        
        # get the app pools that use the GUID (w/o the - chars)
        $results += Get-AppPoolStatus -ComputerName $server.Name -ApplicationPoolName $serviceApplicationPool.Id.ToString().Replace("-", "") -Verbose
    }
}

$results | SORT ComputerName, ApplicationPoolName | FT ComputerName, ApplicationPoolName, ApplicationPoolStatus -AutoSize
$resultsapp = $results | ConvertTo-Html -Fragment

#endregion

#region SharePoint Service Applications Status

#######################################  SharePoint Service Application Status Details  ##################################

$ServiceAppplications = Get-SPServiceApplication | Select DisplayName , ApplicationVersion , Status , NeedsUpgrade | Convertto-Html -Fragment
$ServiceAppplications = SetSharePointStatusColor $ServiceAppplications
write-host Received Service Application Status Details -foregroundcolor "green"

#######################################  SharePoint Service Application Proxy Status Details  ###########################

$ApplicationProxies = Get-SPServiceApplicationProxy | Select TypeName ,Status , NeedsUpgrade | ConvertTo-Html -Fragment
$ApplicationProxies = SetSharePointStatusColor $ApplicationProxies
write-host Received Application Proxy Status Details -foregroundcolor "green"

#######################################  SharePoint Service Instances Status Details  ###################################

$SPServiceInstanceCol = @()
foreach($Name in $ServersinFarm)
{
$SPServiceInstanceCol += Get-SPServiceInstance -Server $Name | Select Server,TypeName,Status,NeedsUpgrade | Sort-Object Status
}
$SPServiceInstances = $SPServiceInstanceCol | Convertto-Html -Fragment
$SPServiceInstances = SetSharePointStatusColor $SPServiceInstances
write-host Received Service Instances Status  Details -foregroundcolor "green"

#endregion

#region SharePoint Custom Solution Status

#######################################  SharePoint Custom Solution Status Details  ######################################

$SPSolutions = Get-SPSolution | Select Name , Deployed , Status | ConvertTo-Html -Fragment
$SPSolutions = SetSharePointStatusColor $SPSolutions
write-host Received SharePoint Custom Solution Status Details -foregroundcolor "green"

#endregion

#region SharePoint Databases status

#######################################  SharePoint Database Status Details  #############################################

$DBCheck = Get-SPDatabase | Select Name ,NormalizedDataSource,WebApplication,CurrentSiteCount, Status , NeedsUpgrade , @{Label ="Size in MB"; Expression = {$_.disksizerequired/1024/1024}}

$SPDBReport = $DBCheck | ConvertTo-Html -Fragment
$SPDBReport = SetSharePointStatusColor $SPDBReport

#Get Database Information
$DatabaseCol = "" |  select "NoOfDatabase","TotalDatabaseSize"
$DatabaseCol.NoOfDatabase = ($DBCheck | Measure-Object).Count
$DatabaseCol.TotalDatabaseSize = ($DBCheck | Measure-Object "Size in MB" -Sum).Sum
$DatabaseColReport =  $DatabaseCol | ConvertTo-Html -Fragment
$DatabaseColReport = SetSharePointStatusColor $DatabaseColReport

write-host Received Database Status Details -foregroundcolor "green"

#endregion

 
#region SharePoint Native Backup History

#######################################  SharePoint Native Backup History   ##################################################

$backup = $xml.variables.spbkloc

$SPBackup = Get-SPBackupHistory -Directory $backup  | select Method , Starttime, Endtime
$SPBackup = $SPBackup | ConvertTo-Html -Fragment
write-host Received SharePoint Native Backup History Details -foregroundcolor "green"

#endregion

#region SharePoint Search Component Staus

#######################################  SharePoint Search Component Status  ##################################################

 
 
$SPSStatus = Get-SPEnterpriseSearchServiceApplication | Get-SPEnterpriseSearchStatus | select Name, state
$SPSStatus = $SPSStatus | ConvertTo-Html -Fragment
write-host Received SharePoint Search Component Status -foregroundcolor "green"

#endregion
#######################################  SharePoint Workflow Engine Status ##################################################

#region SharePoint Workflow Engine Status

Add-PSSnapin microsoft.sbconsole.powershell -ea 0

$wfestatus = get-WFfarmStatus | ConvertTo-Html -Fragment

 
write-host Received SharePoint Workflow Engine Status -foregroundcolor "green"
#endregion

#######################################  Workflow Service Status ##################################################

#region SharePoint Workflow Service Status

function Get-WorkflowServiceStatus
{
    param ($sSiteCUrl,$sWebAppUrl,$sOperationType)

    try
    {
        switch ($sOperationType) 
        { 
        "SC" {
            Write-Host "Getting Workflow Status for Site Collection $sSiteCUrl" -ForegroundColor Green                        
            Get-SPWorkflowConfig -SiteCollection $sSiteCUrl
            } 
        "WA" {
            Write-Host "Getting Workflow Status for Web Application $sWebAppUrl" -ForegroundColor Green              
            Get-SPWorkflowConfig -WebApplication $sWebAppUrl
            }        
        default {
            Write-Host "Requested Operation not valid!!" -ForegroundColor DarkBlue            
            }
        }    
    }
    catch [System.Exception]
    {
        write-host -f red $_.Exception.ToString()
    }
}

#Required Parameters
$sSiteUrl = $SC
$sWebAppUrl = $webapp
$sOperationType="WA"

 
 
$SPCstatus = Get-WorkflowServiceStatus -sSiteCUrl $sSiteUrl -sWebAppUrl $sWebAppUrl -sOperationType $sOperationType | ConvertTo-Html -Fragment

 
write-host Received SharePoint Workflow Service Status -foregroundcolor "green"
#endregion

#######################################  AppFarbic Service Status ##################################################

#region AppFarbic Status

Use-CacheCluster
$APFstatus = Get-CacheHost | ConvertTo-Html -Fragment

 
write-host Received AppFarhic Cache Status -foregroundcolor "green"
#endregion

 
 
####################################### SharePoint Update Status ##################################################

#region SharePoint Update Status

$patch = get-spproduct
$spupdate = $patch.servers | select Servername, RequiredButMissingPatchableUnits,RequiredButMissingPatches, RequiredButMissingProducts | ConvertTo-Html -Fragment

 
write-host Received SharePoint Update Status -foregroundcolor "green"
#endregion

 
 
 
########################################SQL Server backup#####################################################

 
$ConfigDB = Get-SPDatabase | Where-Object{$_.Type -eq "Configuration Database"}

 
$ServerList = $ConfigDB.Server

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
foreach ($address in $ServerList)
{
       
    $SQLServer = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $address 
    foreach($Database in $SQLServer.Databases)
    {
        $SQLBKP += Select  Name, RecoveryModel,LastBackupDate,LastDifferentialBackupDate,LastLogBackupDate

    }
}

$SQLBKP = $SQLBKP | ConvertTo-Html -Fragment

 
#######################################Health Analyzer###############################################################

#regin Health Analyzer

$ReportsList = [Microsoft.SharePoint.Administration.Health.SPHealthReportsList]::Local
$FormUrl = '{0}{1}?id=' -f $ReportsList.ParentWeb.Url, $ReportsList.Forms.List.DefaultDisplayFormUrl
[array]$body = $Null

$ReportsList.Items | Where-Object {$_['Severity'] -ne '4 - Success' -or $_['Title'] -contains '*Missing server side dependencies*'}  | foreach-Object {
   $item = New-Object PSObject
   #$item | Add-Member NoteProperty Severity $_['Severity']
  # $item | Add-Member NoteProperty Category $_['Category']
   $item | Add-Member NoteProperty Title $_['Title']
  $item | Add-Member NoteProperty Explanation $_['Explanation']
  #$item | Add-Member NoteProperty Modified $_['Modified']
   $item | Add-Member NoteProperty FailingServers $_['Failing Servers']
   $item | Add-Member NoteProperty FailingServices $_['Failing Services']
  #$item | Add-Member NoteProperty Remedy $_['Remedy']
  
   $body += $item
   }
  
 
 
$SPFarmHA = $body | ConvertTo-Html -Fragment

#### Workflow Manager Status
$wf = Get-WFFarmStatus | SORT HostName | SELECT HostName, ServiceName, ServiceStatus  | ConvertTo-Html -Fragment

#### ServiceBus Status
$sb = (Get-SBFarmStatus).GetEnumerator() | SORT HostName | SELECT HostId, HostName, Servicename, Status |ConvertTo-Html -Fragment 

 
############# Using Memory ##############

 
$OWAServer = Get-SPWOPIZone | get-spwopibinding | select-object servername -first 1
$OWAServer
Invoke-Command -ComputerName $OWAServer.ServerName -ScriptBlock{

$owadata = (Get-OfficeWebAppsFarm).Machines
}

$owadata = $owadata | ConvertTo-Html -Fragment



############################ Time Stamp##############################################
$ReportDate = Get-Date |select Day , DayOfWeek , Year | ConvertTo-Html -Fragment

#Get Start Time
$StartDate = Get-Date
$StartDateTime = $StartDate.ToUniversalTime()

#######################################  Convert Report to HTML #################################################################

 $Header = @"
<style>
body {font-family:Calibri; font-size:10pt;}
th {background-color:#045FB4;color:white;}
td {background-color:#D8D8D8;color:black;}
</style>
"@

 
ConvertTo-Html -Head $header -Body "
<font color = blue><H1><B>SharePoint Monitoring Dashboard Report Date</B></H1></font>
<font color = blue>Time Stamp -$StartDateTime</font>
<font color = blue><H4><B>CPU Utilization</B></H4></font>$CPU
<font color = blue><H4><B>Memory Utilization</B></H4></font>$Memory
<font color = blue><H4><B>Disk Utilization</B></H4></font>$Disk
<font color = blue><H4><B>SharePoint Health Analyzer</B></H4></font>$SPFarmHA
<font color = blue><H4><B>SharePoint Farm Status</B></H4></font>$SPFarm
<font color = blue><H4><B>SharePoint servers status</B></H4></font>$SPServersInfo
<font color = blue><H4><B>SharePoint Farm Server List</B></H4></font>$SPServerDetails
<font color = blue><H4><B>SharePoint Web Application Status</B></H4>$WebApplication
<font color = blue><H4><B>Site Collection Status</B></H4></font>$WebResponseTime
<font color = blue><H4><B>SharePoint Windows Services Status</B></H4></font>$SPServices
<font color = blue><H4><B>IIS Application Pool Status</B></H4>$resultsapp
<font color = blue><H4><B>SharePoint Service Application Status</B></H4>$ServiceAppplications
<font color = blue><H4><B>SharePoint Service Application Proxy Status</B></H4>$ApplicationProxies
<font color = blue><H4><B>SharePoint Service Instances Status</B></H4>$SPServiceInstances
<font color = blue><H4><B>SharePoint Custom Solution Status</B></H4></font>$SPSolutions
<font color = blue><H4><B>SharePoint Content Database Status</B></H4></font>$SPDBReport
<font color = blue><H4><B>SharePoint Database Information</B></H4></font>$DatabaseColReport
<font color = blue><H4><B>SharePoint Native Backup History</B></H4></font>$SPBackup
<font color = blue><H4><B>SharePoint Search Component Status</B></H4></font>$SPSStatus
<font color = blue><H4><B>SharePoint Search History</B></H4></font>$crawl
<font color = blue><H4><B>Workflow Engine Status</B></H4></font>$wf
<font color = blue><H4><B>Service Bus Status</B></H4></font>$sb
<font color = blue><H4><B>SharePoint AppFarbic Cache Status</B></H4></font>$APFstatus
<font color = blue><H4><B>Office Web Application Status</B></H4></font>$owadata
<font color = blue>Verison 3.5</font

 
<font color = blue><H4><B>" -Title "SharePoint Farm Status Report"  | Out-File "D:\scripts\FarmReport.html" -Encoding ascii

#######################################  Send Email ###########################################################################

 
 
 
$fromaddress = $xml.variables.fromaddress
$toaddress = $xml.variables.toaddress
$Subject = $xml.variables.subject
$smtpserver = $xml.variables.smtpserver
$body = Get-Content "D:\scripts\FarmReport.html"
#$attachment = 'D:\Scripts\FarmReport.html'

 
 
 
$message = new-object System.Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
$message.From = $fromaddress
$message.To.Add($toaddress)
$message.IsBodyHtml = $TRUE
$message.Subject =$Subject

 
 
$message.body = $body
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
$smtp.Send($message)
#$attachment.Dispose();
$message.Dispose();

 
#######################################  Monitoring Script End #################################################################

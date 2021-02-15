<#
    .SYNOPSIS
        THOTH - Ultimate System Inventory & Log Collector: GUI

    .DESCRIPTION
        With more and more end users working from home, remotely, the IT departments have started to face a new challenge when the same users are experiencing technical issues with their local computers.
        Some times a remote shared session is not possible due to the issue per-se, and end users are not capable to provide deep technical information in order to help on the troubleshooting process.
        This script has been built in a way that it can be directly shared with end-users, where the only thing they have to do, is to download and execute the THOTH-FORM.PS1 as Administrator.
        It will compile a simple GUI to instruct the end-user through the process; 
        The script will collect basic and advanced logs from the local computer, creates a HTML report and zips everything into a single file that can then, be shared with the user's IT department.
        Hopefully, the collected logs and data will give us enough information to, at least, restabilish remote connectivity where we can then, jump in to a remote shared session, in order to finish the break-fix process.

        Results and files are placed at: C:\LogCollector

        CUSTOMIZATIOM & BRANDING:
        You can change the variables at the begining of the code in order to customize the look and feel of the HTML report.
        Customizable items and their default values are:
            Title: THOTH
            Logo: An image of THOTH, God of Record Keeping (Embedded in the code via BASE64 image string)
            Footer: Created by Rafael Machado (rafaelcarneiromachado@gmail.com)
            Main Background Color = "#3B3B3B"
            Main Font Color = "#FFFFFF"
        
        DATA INVENTORIED & COLLECTED:
        1. Hardware Inventory;
        2. Network Inventory, including ISP information and Internet Speed Test (Download Throughput and Average Ping);
        3. Driver Inventory;
        4. Windows Services Statuses;
        5. Resultant Set of Policies (GPOs / GPResult);
        6. Software Inventory;
        7. Windows Update: Hotfixes Installed and Missing;
        8. EventViewer Events from Application and System since last boot time;
        9. Overall System Performance Report considering a sample of 60 seconds, taken during the data gathering process;
        10. Mobile Device Management Logs (if the device is corporate managed via a MDM platform);
        11. CBS Log;
        12. SCCM Logs (If SCCM Agent installed);
        13. VMWare Horizon Logs for Horizon Agent and Horizon Client (If installed);

    .EXAMPLE
        Thoth-Form.ps1

    .NOTES
        When reviewing the code, press CTRL + M to EXPAND/COLLPASE the PS Regions. That will make it easier to understand the code structure.

        This script THOTH-FORM.PS1 has everything embedded, including the INVOKE-THOTHSCAN.PS1, which is the script that performs the actual scan;
        I've put that way in order to make it easier to send just one single file to end users or to compile the PS script into an EXE file;
        If you don't want to use the GUI, you can extract (copy/paste) the function code of INVOKE-THOTHSCAN and use it as a stand-alone PS function;
        
        PS1 to EXE:
        I'd recommend to compile the Thoth-Form.ps1 into an .EXE file in order to make it more user-friendly;
        In order to do that, you can use Inno Setup: https://jrsoftware.org/isinfo.php
        The project page at GitHub has the InnoSetup script file, which you can import, modify as you wish, and compile/re-compile the PS1 into an EXE file

    .AUTHOR
        Rafael Carneiro Machado          

    .LINK
        https://github.com/rafaelcarneiromachado
        https://www.linkedin.com/in/rafaelcarneiromachado/

#>

#-------------------------------------------------------------------------------------------
# CUSTOMIZATION / BRANDING: Here you can change the following values in order to customize the look & feel of the tool:
$mainTitle = 'THOTH'                                                         # Main Title
$logoFile = 'C:\Example\Logo.png'                                            # In case you want to replace the default Logo. Make sure to type a full path to a .PNG file, transparent and with a 150x150px size
$footer = 'Created by Rafael Machado (rafaelcarneiromachado@gmail.com)'      # Information placed at the end of the HTML Report
$backgroundColor = '#3B3B3B'                                                 # Hex value for the main background color that will be applied to the elements of the HTML report
$foregroundColor = '#FFFFFF'                                                 # Hex value for the main foreground color (Font Colors) that will be applied to the elements of the HTML report
$itDept = 'your IT Department'
$email = 'support@yourITdepartment.com'
#-------------------------------------------------------------------------------------------


#----DO NOT CHANGE ANYTHING FROM THIS POINT ONWARDS, UNLESS YOU KNOW WHAT YOU ARE DOING ----
$Invocation = (Get-Variable MyInvocation -Scope 0).Value
$script:defaultPath = Split-Path $Invocation.MyCommand.Path

#-------------------------------------------------------------------------------------------
#region FORM: Check if running as Administrator

$nonAdminPopup = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$nonAdminMessage = "Looks like the tool is not running with local administrator privileges:`n`n`n1. If you are using a corporate computer, please contact your IT Department;`n`n2. If you are using a personal computer, make sure you run the EXE file as Administrator (Locate the EXE file, with your mouse pointer, RIGHT CLICK over it and select RUN AS ADMINISTRATOR)"
Try {
    $myWindowsID=[Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
    $adminRole=[Security.Principal.WindowsBuiltInRole]::Administrator
    If (-not $myWindowsPrincipal.IsInRole($adminRole)) {
        $nonAdminPopup.Popup($nonAdminMessage,0,'Attention',0+48)
        Return
    }
} Catch {
    $nonAdminPopup.Popup($nonAdminMessage,0,'Attention',0+48)
    Return
}
#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Functions

function script:Run {
    Add-Type -AssemblyName System.Windows.Forms
    $script:abort = $false
    $buttonRun.Hide()
    $buttonQuit.Hide()
    $buttonAbort.Show()
    $logOutput.Show()
    $progressBar.Show()
    $h3.Text = 'Processing ...'
    $h3.ForeColor = 'Red'
    $h3.Location = '7,135'
    $progressBar.Visible = $True

    $script:processRunspace =[runspacefactory]::CreateRunspace()
    $script:processRunspace.ApartmentState = 'STA'
    $script:processRunspace.ThreadOptions = 'ReuseThread'          
    $script:processRunspace.Open()
    $script:processRunspace.SessionStateProxy.SetVariable('logOutput',$logOutput)
    $script:processRunspace.SessionStateProxy.SetVariable('defaultPath',$script:defaultPath)
    $script:processRunspace.SessionStateProxy.SetVariable('mainTitle',$mainTitle)
    $script:processRunspace.SessionStateProxy.SetVariable('logoFile',$logoFile)
    $script:processRunspace.SessionStateProxy.SetVariable('footer',$footer)
    $script:processRunspace.SessionStateProxy.SetVariable('backgroundColor',$backgroundColor)
    $script:processRunspace.SessionStateProxy.SetVariable('foregroundColor',$foregroundColor)
    try {
        $script:mainScript = Get-Content function:\Invoke-ThothScan
        $script:psCmd = [PowerShell]::Create().AddScript($script:mainScript)
        $script:psCmd.Runspace = $script:processRunspace 
        $handle = $script:psCmd.BeginInvoke()
        do {
            [Windows.Forms.Application]::DoEvents()
        } until ($handle.IsCompleted)
        $null = $script:processRunspace.Close()
        $null = $script:processRunspace.Dispose()
        $null = $script:psCmd.Stop()
        $null = $script:psCmd.Dispose()
    } catch {
        $message = ('Error while executing Invoke-ThothScan.ps1: {0}' -f $_)
        Write-Host $message
        $logOutput.AppendText($message)
    }
    if ($script:abort) {
        $progressBar.Visible = $false
        $progressBar.Hide()
        $h3.Location = '170,180'
        $h3.ForeColor = 'Orange'
        $h3.Text = 'Process Aborted'
        $buttonAbort.Hide()
        $buttonQuit.Show()
        $buttonRun.Show()
        Write-Host $logOutput.Text      
    } else {
        $progressBar.Visible = $false
        $progressBar.Hide()
        $h3.Location = '170,180'
        $h3.ForeColor = 'Green'
        $h3.Text = 'Process Completed'
        $buttonAbort.Hide()
        $buttonQuit.Location = '170,400'
        $buttonQuit.Show()
        Write-Host $logOutput.Text
        try {
            $script:lastReportFolder = Get-ChildItem -Path 'C:\LogCollector' | Sort-Object LastWriteTime -descending | Select-Object -ExpandProperty fullname -first 1
            $lastReportName = Get-ChildItem -Path $LastReportFolder -Name *.html
            $lastReport = "$lastReportFolder\$lastReportName"
            $zipFile = Get-ChildItem -Path $LastReportFolder -Name *.zip
            $finalMessage = "Please, send the file $zipFile to $itDept via: $email`r`n`r`nMind the file size as it may be bigger than the limit allowed by your e-mail policies"
            [Windows.Forms.MessageBox]::Show($finalMessage,'Process Completed','OK','Information')
            Start-Process $lastReport
            & "$env:windir\explorer.exe" $script:lastReportFolder
        } catch {
            $message = "Error on the last steps of the process: $_"
            Write-Host $message
            $logOutput.AppendText($message)
        }
    }
}

function script:Abort {
    $script:abort = $True
    $script:processRunspace.Close()
    $script:psCmd.Stop()
    $script:psCmd.Dispose()
    $logOutput.AppendText("`r`nProcess aborted by the user")
}

function Invoke-ThothScan {
    <#
        .SYNOPSIS
            THOTH - Ultimate System Inventory & Log Collector: Raw Script

        .DESCRIPTION
            The script can or cannot be used in combination with Thor-Form.ps1, which is an user-friendly GUI responsible to trigger this script, although the THOTH-FORM.ps1 is not required;
            It will perform a deep scan in the local computer, inventorying all relevant information plus it will collect useful data and logs in order to help on remote troubleshooting.
            At the end, the script creates an HTML report and a ZIP file containing all information retrieved and collected during the process.

            Files are placed at: C:\LogCollector

            CUSTOMIZATIOM & BRANDING:
            You can change the variables at the begining of the code in order to customize the look and feel of the HTML report.
            Customizable items and their default values are:
                Title: THOTH
                Logo: An image of THOTH, God of Recoord Keeping (Embedded in the code via BASE64 image string)
                Footer: Created by Rafael Machado (rafaelcarneiromachado@gmail.com)
                Main Background Color = "#3B3B3B"
                Main Font Color = "#FFFFFF"
        
            DATA INVENTORIED & COLLECTED:
            1. Hardware Inventory;
            2. Network Inventory, including ISP information and Internet Speed Test (Download Throughput and Average Ping);
            3. Driver Inventory;
            4. Windows Services Statuses;
            5. Resultant Set of Policies (GPOs / GPResult);
            6. Software Inventory;
            7. Windows Update - Hotfixes Installed;
            8. EventViewer Events from Application and System since last boot time;
            9. Overall System Performance Report considering a sample of 60 seconds, taken during the data gathering process;
            10. Mobile Device Management Logs (if the device is corporate managed via a MDM platform);
            11. CBS Log;
            12. SCCM Logs (If SCCM Agent installed);
            13. VMWare Horizon Logs for Horizon Agent and Horizon Client (If installed);

        .EXAMPLE
            Invoke-ThothScan.ps1

        .NOTES
            When reviewing the code, press CTRL + M to EXPAND/COLLPASE the PS Regions. That will make it easier to understand the code structure.

        .AUTHOR
            Rafael Carneiro Machado          

        .LINK
            https://github.com/rafaelcarneiromachado
            https://www.linkedin.com/in/rafaelcarneiromachado/

    #>

    #region Check if Customization & Branding is coming from Thor-Form.ps1. Otherwise, accept Customization & Branding defined here.
    [int]$brandingFromForm = 0
    if (Test-Path variable:\mainTitle) {$brandingFromForm = $brandingFromForm +1}
    if (Test-Path variable:\logoFile) {$brandingFromForm = $brandingFromForm +1}
    if (Test-Path variable:\footer) {$brandingFromForm = $brandingFromForm +1}
    if (Test-Path variable:\backgroundColor) {$brandingFromForm = $brandingFromForm +1}
    if (Test-Path variable:\foregroundColor) {$brandingFromForm = $brandingFromForm +1}
    If ($brandingFromForm -eq 0) {
    #endregion

    #-------------------------------------------------------------------------------------------
    # CUSTOMIZATION / BRANDING: Here you can change the following values in order to customize the look & feel of the tool:
    $mainTitle = 'THOTH'                                                         # Main Title
    $logoFile = 'C:\Example\Logo.png'                                            # In case you want to replace the default Logo. Make sure to type a full path to a .PNG file, transparent and with a 150x150px size
    $footer = 'Created by Rafael Machado (rafaelcarneiromachado@gmail.com)'      # Information placed at the end of the HTML Report
    $backgroundColor = '#3B3B3B'                                                 # Hex value for the main background color that will be applied to the elements of the HTML report
    $foregroundColor = '#FFFFFF'                                                 # Hex value for the main foreground color (Font Colors) that will be applied to the elements of the HTML report
    #-------------------------------------------------------------------------------------------

    }

    #----DO NOT CHANGE ANYTHING FROM THIS POINT ONWARDS, UNLESS YOU KNOW WHAT YOU ARE DOING ----
    if (!(Test-Path variable:\defaultPath)) {
        $Invocation = (Get-Variable MyInvocation -Scope 0).Value
        $defaultPath = Split-Path $Invocation.MyCommand.Path
    }

    #-------------------------------------------------------------------------------------------
    #region FORM: Check if running as Administrator

    $nonAdminMessage = 'You must run the script AS ADMINISTRATOR!'
    try {
        $myWindowsID=[Security.Principal.WindowsIdentity]::GetCurrent()
        $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
        $adminRole=[Security.Principal.WindowsBuiltInRole]::Administrator
        If (-not $myWindowsPrincipal.IsInRole($adminRole)) {
            Write-Host $nonAdminMessage
            Return
        }
    } catch {
        Write-Host $nonAdminMessage
        Return
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Define Helper Functions

    function Get-TimeStamp {
        return '[{0:dd/MM/yyyy} {0:HH:mm:ss}]' -f (Get-Date)
    }
    
    function script:Transcript {
        Param ([Parameter(Mandatory=$true)]$Section,[switch]$SubHeader)
        $message = ''
        $message = "$(Get-TimeStamp) $Section"
        if (Get-Variable logOutput -ErrorAction SilentlyContinue) {
            $logOutput.AppendText("`r`n$message")
        }
        Write-Host $message -ErrorAction Continue
       
        If ($subHeader) {
            $script:subHeader = '<hr><p id="ListHeader"><b>' + $section + '</b></p>'
        } Else {
            $script:preContent = ''
            $script:postContent = ''
            $script:preContent = '<button class="collapsible">' + $section + ':</button><div class="content">'
            $script:postContent = '</div>'
        }
        $script:noData = '<p><font color="red">Unable to retrieve information</font></p>'
    }

    function script:LicenseStatus {
        $script:licenseStatus = ''
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" -Property LicenseStatus -ErrorAction Ignore
        } catch {
            $script:licenseStatus = 'Unable to Identify' 
        }

        if ($wpa) {
            :outer foreach ($item in $wpa) {
                switch ($item.LicenseStatus) {
                    0 {$script:licenseStatus = 'Unlicensed'}
                    1 {$script:licenseStatus = 'Licensed'; break outer}
                    2 {$script:licenseStatus = 'Out-Of-Box Grace Period'; break outer}
                    3 {$script:licenseStatus = 'Out-Of-Tolerance Grace Period'; break outer}
                    4 {$script:licenseStatus = 'Non-Genuine Grace Period'; break outer}
                    5 {$script:licenseStatus = 'Notification'; break outer}
                    6 {$script:licenseStatus = 'Extended Grace'; break outer}
                }
            }
        } else {
            $licenseStatus = 'Unable to Identify'
        }
    }

    function script:Hyperlink {
        Param ([Parameter(Mandatory=$true)]$linkPath)
        $script:htmlLink = 'START' + $linkPath + 'MIDDLE' + $linkPath + 'END'
    }

    function script:Internet {
        $internetCheck = 0
        $netProfiles = Get-NetConnectionProfile
        foreach ($netProfile in $netProfiles) {
            $netCount = if ($netProfiles.IPv4Connectivity -eq 'Internet') {1} else {0}
            $internetCheck = $internetCheck + $netCount
        }
        if ($internetCheck -gt 0) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $internetSecondCheck = if ((Invoke-RestMethod -Uri 'http://ip-api.com/json').status -eq 'success') {$True} Else {$False}
            if ($internetSecondCheck) {
                $script:internet = $True
            } else {
                $script:internet = $False
            }
        } else {
        $script:internet = $False
        }
    }

    Function script:DownloadSpeed { 
        Param ([Parameter(Mandatory=$true)]$url)
        $downloadSpeed = ''
        try {
            $col = new-object System.Collections.Specialized.NameValueCollection 
            $wc = new-object system.net.WebClient 
            $wc.QueryString = $col 
            $downloadElaspedTime = (measure-command {$webpage1 = $wc.DownloadData($url)}).totalmilliseconds
            $downSize = ($webpage1.length + $webpage2.length) / 1Mb
            $downloadSize = [Math]::Round($downSize, 2)
            $downloadTimeSec = $downloadElaspedTime * 0.001
            $downSpeed = ($downloadSize / $downloadTimeSec) * 8
            $downloadSpeed = [Math]::Round($downSpeed, 2)
            if ($downloadSpeed -match '^[\d\.]+$') {
                $downloadSpeed
            } else {
                $downloadSpeed = $null
                $downloadSpeed
            }
        } catch {
            $downloadSpeed = $null
            $downloadSpeed
        }
    }

    Function script:Ping {
        Param ([Parameter(Mandatory=$true)]$pingServer)
        $pingResult = ''
        try {
            $pingResult = (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$pingServer'").ResponseTime
            if ($pingResult -match '^[\d\.]+$') {
                $pingResult
            } else {
                $pingResult = $null
                $pingResult
            }
        } catch {
            $pingResult = $null
            $pingResult
        }
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Preparing the environment
    
    #Prepare Folders and Timestamps
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $culture = New-Object system.globalization.cultureinfo('en-GB')
    $currentDateTime = Get-Date -format 'yyyy-MM-ddTHH.mm.ss'
    $bootTime = Get-WmiObject -class Win32_OperatingSystem | Select-Object @{n='Last Boot Time';e={Get-Date($_.ConvertToDateTime($_.LastBootUpTime))}}
    $bootTimestampMinus2hours = (Get-Date -Date $bootTime.'Last Boot Time').AddHours(-2)
    $bootTimestampForLogs = Get-Date $bootTimestampMinus2hours -Format 'yyyy-MM-ddTHH:mm:ss'
    $currentTimestampForLogs = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
    $rootPath = "$env:SystemDrive\LogCollector\$currentDateTime"

    #Start Transcript
    $section = 'Preparing the environment'
    Start-Transcript -Path "$rootPath\0.LogCollector_ExecutionLog.log" -Force
    Transcript -Section $section

    New-Item -ItemType Directory "$env:windir\Temp\LogCollector" -Force
    $null = New-Item -Path $rootPath -ItemType Directory -ErrorAction Ignore | Select-Object -ExpandProperty Name
    $logFolder = "$rootPath\LogFiles"
    $null = New-Item -Path $logFolder -ItemType Directory -ErrorAction Ignore | Select-Object -ExpandProperty Name
    $perfFolder = "$rootPath\Performance"
    $null = New-Item -Path $perfFolder -ItemType Directory -ErrorAction Ignore | Select-Object -ExpandProperty Name
    try {
        $currentLoggedUser = ((& "$env:windir\system32\quser.exe") -replace '^>', '') -replace '\s{2,}', ',' | ConvertFrom-Csv | Where-Object {$_.state -eq 'Active'} | Select-Object -ExpandProperty username
    } catch {
        $currentLoggedUser = $script:noData
    }
  $perfXML = @'
<?xml version="1.0" encoding="UTF-16"?>
<DataCollectorSet>
<Status>0</Status>
<Duration>60</Duration>
<Description>Generate a report detailing the status of local hardware resources, system response times, and processes on the local computer. Use this information to identify possible causes of performance issues. Membership in the local Administrators group, or equivalent, is the minimum required to run this Data Collector Set.</Description>
<DescriptionUnresolved>Generate a report detailing the status of local hardware resources, system response times, and processes on the local computer. Use this information to identify possible causes of performance issues. Membership in the local Administrators group, or equivalent, is the minimum required to run this Data Collector Set.</DescriptionUnresolved>
<DisplayName>
</DisplayName>
<DisplayNameUnresolved>
</DisplayNameUnresolved>
<SchedulesEnabled>-1</SchedulesEnabled>
<Keyword>CPU</Keyword>
<Keyword>Memory</Keyword>
<Keyword>Disk</Keyword>
<Keyword>Network</Keyword>
<Keyword>Performance</Keyword>
<LatestOutputLocation>
</LatestOutputLocation>
<Name>Perf</Name>
<OutputLocation>C:\Windows\Temp\LogCollector\</OutputLocation>
<RootPath>C:\Windows\Temp\LogCollector</RootPath>
<Segment>0</Segment>
<SegmentMaxDuration>0</SegmentMaxDuration>
<SegmentMaxSize>0</SegmentMaxSize>
<SerialNumber>1</SerialNumber>
<Server>
</Server>
<Subdirectory>
</Subdirectory>
<SubdirectoryFormat>1</SubdirectoryFormat>
<SubdirectoryFormatPattern>
</SubdirectoryFormatPattern>
<Task>
</Task>
<TaskRunAsSelf>0</TaskRunAsSelf>
<TaskArguments>
</TaskArguments>
<TaskUserTextArguments>
</TaskUserTextArguments>
<UserAccount>SYSTEM</UserAccount>
<Security></Security>
<StopOnCompletion>0</StopOnCompletion>
<TraceDataCollector>
	<DataCollectorType>1</DataCollectorType>
	<Name>NT Kernel</Name>
	<FileName>NtKernel</FileName>
	<FileNameFormat>0</FileNameFormat>
	<FileNameFormatPattern>
	</FileNameFormatPattern>
	<LogAppend>0</LogAppend>
	<LogCircular>0</LogCircular>
	<LogOverwrite>0</LogOverwrite>
	<LatestOutputLocation>
	</LatestOutputLocation>
	<Guid>{00000000-0000-0000-0000-000000000000}</Guid>
	<BufferSize>64</BufferSize>
	<BuffersLost>0</BuffersLost>
	<BuffersWritten>0</BuffersWritten>
	<ClockType>1</ClockType>
	<EventsLost>0</EventsLost>
	<ExtendedModes>0</ExtendedModes>
	<FlushTimer>0</FlushTimer>
	<FreeBuffers>0</FreeBuffers>
	<MaximumBuffers>200</MaximumBuffers>
	<MinimumBuffers>0</MinimumBuffers>
	<NumberOfBuffers>0</NumberOfBuffers>
	<PreallocateFile>0</PreallocateFile>
	<ProcessMode>0</ProcessMode>
	<RealTimeBuffersLost>0</RealTimeBuffersLost>
	<SessionName>NT Kernel Logger</SessionName>
	<SessionThreadId>0</SessionThreadId>
	<StreamMode>1</StreamMode>
	<TraceDataProvider>
		<DisplayName>{9E814AAD-3204-11D2-9A82-006008A86939}</DisplayName>
		<FilterEnabled>0</FilterEnabled>
		<FilterType>0</FilterType>
		<Level>
			<Description>Events up to this level are enabled</Description>
			<ValueMapType>1</ValueMapType>
			<Value>0</Value>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x0</Value>
			</ValueMapItem>
		</Level>
		<KeywordsAny>
			<Description>Events with any of these keywords are enabled</Description>
			<ValueMapType>2</ValueMapType>
			<Value>0x10303</Value>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x1</Value>
			</ValueMapItem>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x2</Value>
			</ValueMapItem>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x100</Value>
			</ValueMapItem>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x200</Value>
			</ValueMapItem>
			<ValueMapItem>
				<Key>
				</Key>
				<Description>
				</Description>
				<Enabled>-1</Enabled>
				<Value>0x10000</Value>
			</ValueMapItem>
		</KeywordsAny>
		<KeywordsAll>
			<Description>Events with all of these keywords are enabled</Description>
			<ValueMapType>2</ValueMapType>
			<Value>0x0</Value>
		</KeywordsAll>
		<Properties>
			<Description>These additional data fields will be collected with each event</Description>
			<ValueMapType>2</ValueMapType>
			<Value>0</Value>
		</Properties>
		<Guid>{9E814AAD-3204-11D2-9A82-006008A86939}</Guid>
	</TraceDataProvider>
</TraceDataCollector>
<PerformanceCounterDataCollector>
	<DataCollectorType>0</DataCollectorType>
	<Name>Performance Counter</Name>
	<FileName>Performance Counter</FileName>
	<FileNameFormat>0</FileNameFormat>
	<FileNameFormatPattern>
	</FileNameFormatPattern>
	<LogAppend>0</LogAppend>
	<LogCircular>0</LogCircular>
	<LogOverwrite>0</LogOverwrite>
	<LatestOutputLocation>
	</LatestOutputLocation>
	<DataSourceName>
	</DataSourceName>
	<SampleInterval>1</SampleInterval>
	<SegmentMaxRecords>0</SegmentMaxRecords>
	<LogFileFormat>3</LogFileFormat>
	<Counter>\Process(*)\*</Counter>
	<Counter>\PhysicalDisk(*)\*</Counter>
	<Counter>\Processor(*)\*</Counter>
	<Counter>\Processor Performance(*)\*</Counter>
	<Counter>\Memory\*</Counter>
	<Counter>\System\*</Counter>
	<Counter>\Server\*</Counter>
	<Counter>\Network Interface(*)\*</Counter>
	<Counter>\UDPv4\*</Counter>
	<Counter>\TCPv4\*</Counter>
	<Counter>\IPv4\*</Counter>
	<Counter>\UDPv6\*</Counter>
	<Counter>\TCPv6\*</Counter>
	<Counter>\IPv6\*</Counter>
	<CounterDisplayName>\Process(*)\*</CounterDisplayName>
	<CounterDisplayName>\PhysicalDisk(*)\*</CounterDisplayName>
	<CounterDisplayName>\Processor(*)\*</CounterDisplayName>
	<CounterDisplayName>\Processor Performance(*)\*</CounterDisplayName>
	<CounterDisplayName>\Memory\*</CounterDisplayName>
	<CounterDisplayName>\System\*</CounterDisplayName>
	<CounterDisplayName>\Server\*</CounterDisplayName>
	<CounterDisplayName>\Network Interface(*)\*</CounterDisplayName>
	<CounterDisplayName>\UDPv4\*</CounterDisplayName>
	<CounterDisplayName>\TCPv4\*</CounterDisplayName>
	<CounterDisplayName>\IPv4\*</CounterDisplayName>
	<CounterDisplayName>\UDPv6\*</CounterDisplayName>
	<CounterDisplayName>\TCPv6\*</CounterDisplayName>
	<CounterDisplayName>\IPv6\*</CounterDisplayName>
</PerformanceCounterDataCollector>
<DataManager>
	<Enabled>-1</Enabled>
	<CheckBeforeRunning>-1</CheckBeforeRunning>
	<MinFreeDisk>200</MinFreeDisk>
	<MaxSize>1024</MaxSize>
	<MaxFolderCount>100</MaxFolderCount>
	<ResourcePolicy>0</ResourcePolicy>
	<ReportFileName>report.html</ReportFileName>
	<RuleTargetFileName>report.xml</RuleTargetFileName>
	<EventsFileName>
	</EventsFileName>
	<Rules>
		<Logging level="15" file="rules.log">
		</Logging>
		<Import file="%systemroot%\pla\rules\Rules.System.Common.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.Summary.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.Performance.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.CPU.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.Network.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.Disk.xml">
		</Import>
		<Import file="%systemroot%\pla\rules\Rules.System.Memory.xml">
		</Import>
	</Rules>
	<ReportSchema>
		<Report name="systemPerformance" version="1" threshold="100">
			<Import file="%systemroot%\pla\reports\Report.System.Common.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.Summary.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.Performance.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.CPU.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.Network.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.Disk.xml">
			</Import>
			<Import file="%systemroot%\pla\reports\Report.System.Memory.xml">
			</Import>
		</Report>
	</ReportSchema>
	<FolderAction>
		<Size>0</Size>
		<Age>1</Age>
		<Actions>27</Actions>
		<SendCabTo>
		</SendCabTo>
	</FolderAction>
</DataManager>
<Value name="PerformanceMonitorView" type="document">
	<OBJECT ID="DISystemMonitor" CLASSID="CLSID:C4D2D8E0-D1DD-11CE-940F-008029004347">
		<PARAM NAME="CounterCount" VALUE="4">
		</PARAM>
		<PARAM NAME="Counter00001.Path" VALUE="\Processor(_Total)\% Processor Time">
		</PARAM>
		<PARAM NAME="Counter00001.Color" VALUE="255">
		</PARAM>
		<PARAM NAME="Counter00001.Width" VALUE="2">
		</PARAM>
		<PARAM NAME="Counter00001.LineStyle" VALUE="0">
		</PARAM>
		<PARAM NAME="Counter00001.ScaleFactor" VALUE="0">
		</PARAM>
		<PARAM NAME="Counter00001.Show" VALUE="1">
		</PARAM>
		<PARAM NAME="Counter00001.Selected" VALUE="1">
		</PARAM>
		<PARAM NAME="Counter00002.Path" VALUE="\Memory\Pages/sec">
		</PARAM>
		<PARAM NAME="Counter00002.Color" VALUE="65280">
		</PARAM>
		<PARAM NAME="Counter00002.Width" VALUE="1">
		</PARAM>
		<PARAM NAME="Counter00003.Path" VALUE="\PhysicalDisk(_Total)\Avg. Disk sec/Read">
		</PARAM>
		<PARAM NAME="Counter00003.Color" VALUE="16711680">
		</PARAM>
		<PARAM NAME="Counter00003.Width" VALUE="1">
		</PARAM>
		<PARAM NAME="Counter00004.Path" VALUE="\PhysicalDisk(_Total)\Avg. Disk sec/Write">
		</PARAM>
		<PARAM NAME="Counter00004.Color" VALUE="55295">
		</PARAM>
		<PARAM NAME="Counter00004.Width" VALUE="1">
		</PARAM>
	</OBJECT>
</Value>
</DataCollectorSet>
'@
    
    #Internet Speed Test Variables
    $minimumDownloadSpeed = 10
    $minimumPing = 230
    $urlDownload = 'https://file-examples-com.github.io/uploads/2017/02/zip_10MB.zip'
    $serverSA = 'ftp.br.debian.org'
    $serverNA = 'ftp.us.debian.org'
    $serverEU = 'ftp.uk.debian.org'
    $serverAS = 'ftp.cn.debian.org'
    $serverOC = 'ftp.au.debian.org'
    $serverAF = 'ftp.is.co.za'

    #Branding/Customization
    $section = 'Branding & Customization'
    Transcript -Section $section
  $defaultLogo = @'
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
'@
  $cssStyles = @"
<style>

html{box-sizing:border-box}*,*:before,*:after{box-sizing:inherit}
html{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}body{margin:0}
html,body{font-family:verdana;font-size:12px;line-height:1.5;}html{overflow-x:hidden}
hr{border:0;border-top:1px solid #eee;margin:20px 0}

.headerContainer{position:relative;letter-spacing:4px;max-width:5500px;height: 150px; background-image: linear-gradient(to bottom right, white, #E3E3E3)}
.Image{position:absolute;top:50%;right:0%;transform:translate(0%,-50%);margin-right:16px;text-align:center}
.TitleSubtitle{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center}
.Title{font-size:36px;padding:8px 16px;background-color:HexMainBackColor;color:HexMainFontColor}
.Subtitle{font-size:36px;font-family:"Segoe UI";font-weight:400;color:#3a3a3a}

button{font:inherit;margin:0}
.collapsible {background-color: HexMainBackColor;color: HexMainFontColor;cursor: pointer;padding: 10px;width: 100%;border: 1px solid white;border-radius: 3px;text-align: left;outline: none;font-size: 14px;font-weight: bold}
.active, .collapsible:hover {filter: contrast(60%);}
.collapsible:after {content: '\002B';color: white;font-weight: bold;float: right;margin-left: 5px;}
.active:after {content: "\2212"}

.content {padding: 0 18px;max-height: 0;overflow: hidden;transition: max-height 0.2s ease-out;background-color: white}
#ListHeader {font-size: 14px;text-decoration: underline;}
a:link, a:visited {color: blue}

table {table-layout: auto;width: 100%;} 
td {font-size:12px;padding: 4px;margin: 0px;border: 0}
th {font-size:11px;background: HexMainBackColor;color: HexMainFontColor;text-transform: uppercase;padding: 10px 15px;vertical-align: middle}
tbody tr:nth-child(even) {background: #f0f0f2}

#Footer {color: #2f2f2f}

</style>
"@
  $cssScript = @'
<script>
    var coll = document.getElementsByClassName("collapsible");
    var i;
    for (i = 0; i < coll.length; i++) {
        coll[i].addEventListener("click", function() {this.classList.toggle("active");
            var content = this.nextElementSibling;
            if (content.style.maxHeight){
                content.style.maxHeight = null;
            } else {
            content.style.maxHeight = content.scrollHeight + "px";
            }
        });
    }
</script>
'@
    $customLogo = $logoFile
    if (Test-Path $customLogo) {
        $logoBase64 = [convert]::ToBase64String((get-content $customLogo -encoding byte))
    } else {
        $logoBase64 = $defaultLogo
    }
    $customHexMainBackColor = $backgroundColor + ';'
    $customHexMainFontColor = $foregroundColor + ';'
    $cssStyles = $cssStyles -replace 'HexMainBackColor',$customHexMainBackColor
    $cssStyles = $cssStyles -replace 'HexMainFontColor',$customHexMainFontColor
    $header1 = '<header class="headerContainer"><div class="Image"><img src="data:image/png;base64, ' + $logoBase64
    $header2 = '" alt="logo.png" /></div><div class="TitleSubtitle"><span class="Title"><b>' + $mainTitle
    $header3 = '</b></span><span class="Subtitle">Ultimate LOG Collector</span></div></header>'
    $header = $header1 + $header2 + $header3
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Basic Info
    $section = 'Basic Information'
    Transcript -Section $section

    try {
        $wmiComputer = Get-WmiObject Win32_ComputerSystem
        $wmiOS = Get-WmiObject Win32_OperatingSystem
        $wmiBios = Get-WmiObject Win32_Bios
        $cbsRebootPending = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction Ignore
        $wuRebootRequired = Test-Path -Path 'HKLM:\\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        $fileRenameRebootRequired = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Ignore
    
        $basicHW = $wmiComputer | Select-Object -Property `
            @{n='Computer Name'; e={$env:computername}},
            @{n='Manufacturer'; e={$wmiComputer.Manufacturer}},
            @{n='Model'; e={$wmiComputer.Model}},
            @{n='Serial Number'; e={$wmiBios.SerialNumber}},
            @{n='BIOS Version'; e={$wmiBios.SMBIOSBIOSVersion}},
            @{n='BIOS Install Date'; e={$wmiBios.InstallDate}},
            @{n='Number of Physical Processors'; e={$wmiComputer.NumberOfProcessors}},
            @{n='Total Physical Memory (GB)'; e={[math]::Round($wmiComputer.TotalPhysicalMemory / 1GB)}}
        $basicHWHTML = $basicHW | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent
   
        $OS = $wmiOS | Select-Object -Property `
            @{n='OS'; e={$wmiOS.Caption}},
            @{n='OS Build'; e={$wmiOS.BuildNumber}},
            @{n='OS License Status'; e={LicenseStatus ; $script:licenseStatus}},
            @{n='Current User'; e={$currentLoggedUser}},
            @{n='Last Boot Time';e={Get-Date $bootTime.'Last Boot Time' -Format 'dd/MM/yyyy HH:mm:ss'}},
            @{n='CBS - Reboot Pending'; e={If ($cbsRebootPending) {'True'} Else {'False'}}},
            @{n='Windows Update - Reboot Required'; e={$wuRebootRequired}},
            @{n='Pending Reboot for File Rename Operations'; e={If ($fileRenameRebootRequired) {'True'} Else {'False'}}}
        $OSHTML = $OS | ConvertTo-Html -As Table -Fragment -PostContent $script:postContent
    
        $basicInfoHTML = $basicHWHTML + $OSHTML
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Local Users
    $section = 'Local Users'
    Transcript -Section $section

    try {
        $wmiAdministrators = (Get-LocalGroupMember -Name 'Administrators').Name | ForEach-Object {(($_) -split '\\')[1]}
        $wmiLocalUsers = Get-WmiObject Win32_UserAccount -Filter "Domain = '$env:computername'" | Where-Object {$_.Name -ne 'WDAGUtilityAccount'}
    
        $localUsers = foreach ($wmiLocalUser in $wmiLocalUsers) {
            $wmiLocalUser | Select-Object -Property `
            @{n='User Account'; e={$wmiLocalUser.Name}},
            @{n='Full Name'; e={$wmiLocalUser.FullName}},
            @{n='Description'; e={$wmiLocalUser.Description}},
            @{n='Domain'; e={$wmiLocalUser.Domain}},
            @{n='Account Status'; e={if ($wmiLocalUser.Disabled) {'Disabled'} else {'Enabled'}}},
            @{n='Account Lockout'; e={$wmiLocalUser.Lockout}},
            @{n='User Can Change Password'; e={$wmiLocalUser.PasswordChangeable}},
            @{n='Password Expires'; e={$wmiLocalUser.PasswordExpires}},
            @{n='Local Administrator'; e={If ($wmiAdministrators -contains $wmiLocalUser.Name) {'Yes'} Else {'No'}}}
        }
    
        $localUsersHTML = $localUsers | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Network Information
    $section = 'Network Information'
    Transcript -Section $section

    try {
        $section = 'Network Interfaces & IP Addresses'
        Transcript -Section $section -SubHeader
    
        $wmiOs = Get-WmiObject Win32_OperatingSystem
        $netAdapters = Get-NetAdapter | Where-Object {$_.Status -ne 'Not Present' -and $_.Virtual -ne 'True'}
        $netIPConfig = Get-NetIPConfiguration | Select-Object -ExpandProperty IPv4Address
        $wmiNetAdapterConfig = Get-WmiObject Win32_NetworkAdapterConfiguration
        $netProfile = Get-NetConnectionProfile
    
        $ipconfig = foreach ($netAdapter in $netAdapters) {
            $netAdapter  | Select-Object -Property `
            @{n='NIC Name'; e={$netAdapter.Name}},
            @{n='NIC Description'; e={$netAdapter.InterfaceDescription}},
            @{n='NIC Status'; e={$netAdapter.Status}},
            @{n='Media Type'; e={$netAdapter.MediaType}},
            @{n='MAC Address'; e={$netAdapter.MacAddress}},
            @{n='IPv4'; e={($netIPConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty IPAddress) -join ', '}},
            @{n='Subnet'; e={$wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty IPSubnet | Select-Object -First 1}},
            @{n='DNS Servers'; e={($wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DNSServerSearchOrder) -join ', '}},
            @{n='DNS Domain'; e={$wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DNSDomain}},
            @{n='DNS Suffixes'; e={($wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DNSDomainSuffixSearchOrder) -join ', '}},
            @{n='Gateway'; e={$wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DefaultIPGateway | Select-Object -First 1}},
            @{n='DHCP Enabled'; e={$wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DHCPEnabled}},
            @{n='DHCP Server'; e={$wmiNetAdapterConfig | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty DHCPServer}},
            @{n='Link Speed'; e={$netAdapter.LinkSpeed}},
            @{n='Connectivity'; e={$netProfile | Where-Object {$_.InterfaceIndex -eq $netAdapter.InterfaceIndex} | Select-Object -ExpandProperty IPv4Connectivity}}
        }
    
        $ipconfigHTML = $ipconfig | ConvertTo-Html -As Table -Fragment -PreContent "$script:preContent $script:subHeader"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Wifi Information'
        Transcript -Section $section -SubHeader
    
        $netshWlanInterfaces = & "$env:windir\system32\netsh.exe" wlan show interfaces

        $wifi = $netshWlanInterfaces | Select-Object -First 1 -Property `
            @{n='Interface Name'; e={(($netshWlanInterfaces | Select-String -Pattern 'Name') -split ': ')[1]}},
            @{n='SSID'; e={(($netshWlanInterfaces | Select-String -Pattern 'SSID') -split ': ')[1]}},
            @{n='State'; e={$tmpString = (($netshWlanInterfaces | Select-String -Pattern 'State') -split ': ')[1] ; (Get-Culture).TextInfo.ToTitleCase($tmpString)}},
            @{n='Wifi Signal'; e={(($netshWlanInterfaces | Select-String -Pattern 'Signal') -split ': ')[1]}},
            @{n='Network Type'; e={(($netshWlanInterfaces | Select-String -Pattern 'Network Type') -split ': ')[1]}},
            @{n='Radio Type'; e={(($netshWlanInterfaces | Select-String -Pattern 'Radio Type') -split ': ')[1]}},
            @{n='Authentication'; e={(($netshWlanInterfaces | Select-String -Pattern 'Authentication') -split ': ')[1]}},
            @{n='Receive Rate (Mbps)'; e={(($netshWlanInterfaces | Select-String -Pattern 'Receive rate \(Mbps\)') -split ': ')[1]}},
            @{n='Transmit Rate (Mbps)'; e={(($netshWlanInterfaces | Select-String -Pattern 'Transmit rate \(Mbps\)') -split ': ')[1]}}
        if ($wifi.'Interface Name' -ne $null) {
            $wifiHTML = $wifi | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader
        } else {
            $wifiHTML = $script:subHeader + $script:noData
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Public IP Address and ISP Information'
        Transcript -Section $section -SubHeader
    
        Internet    
        if ($script:internet) {
            $ipApi = Invoke-RestMethod -Uri 'http://ip-api.com/json'
            $publicIP = $ipApi | Select-Object -Property `
                @{n='Public IP'; e={$ipApi.query}},
                @{n='Country'; e={$ipApi.country}},
                @{n='City'; e={$ipApi.city}},
                @{n='ISP'; e={$ipApi.isp}},
                @{n='Organization'; e={$ipApi.org}}
            $publicIPHTML = $publicIP | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader
        } else {
            $publicIPHTML = $script:subHeader + $script:noData
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Internet Connection Speed and Ping Test'
        Transcript -Section $section -SubHeader
    
        Internet
        if ($script:internet) {
            $speedResult = DownloadSpeed($urlDownload)

            $pingSA = Ping($serverSA)
            $pingNA = Ping($serverNA)
            $pingEU = Ping($serverEU)
            $pingAS = Ping($serverAS)
            $pingOC = Ping($serverOC)
            $pingAF = Ping($serverAF)

            $finalSpeedResult = New-Object PSObject
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Downdload Speed (Avg)' -value {if ($SpeedResult -eq $null) {'Test Failed'} ElseIf ($SpeedResult -ge $minimumDownloadSpeed) {"$SpeedResult Mb/Sec - GOOD"} else {"$SpeedResult Mb/Sec - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(South America)' -value {if ($pingSA -eq $null) {'Test Failed'} Elseif ($pingSA -le $minimumPing) {"$pingSA ms - GOOD"} else {"$pingSA ms - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(North America)' -value {if ($pingNA -eq $null) {'Test Failed'} Elseif ($pingNA -le $minimumPing) {"$pingNA ms - GOOD"} else {"$pingNA ms - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(Europe)' -value {if ($pingEU -eq $null) {'Test Failed'} Elseif ($pingEU -le $minimumPing) {"$pingEU ms - GOOD"} else {"$pingEU ms - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(Asia)' -value {if ($pingAS -eq $null) {'Test Failed'} Elseif ($pingAS -le $minimumPing) {"$pingAS ms - GOOD"} else {"$pingAS ms - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(Oceania)' -value {if ($pingOC -eq $null) {'Test Failed'} Elseif ($pingOC -le $minimumPing) {"$pingOC ms - GOOD"} else {"$pingOC ms - BAD"}}
            Add-Member -inputObject $finalSpeedResult -memberType ScriptProperty -name 'Ping(Africa)' -value {if ($pingAF -eq $null) {'Test Failed'} Elseif ($pingAF -le $minimumPing) {"$pingAF ms - GOOD"} else {"$pingAF ms - BAD"}}
            $speedTestHTML = $finalSpeedResult | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader -PostContent $script:postContent
        } else {
            $speedTestHTML = $script:subHeader + $script:noData + $script:postContent
        }

        $netHTML = $ipconfigHTML + $wifiHTML + $publicIPHTML + $speedTestHTML
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get CPU Info
    $section = 'CPU Information'
    Transcript -Section $section

    try {
        $wmiProcs = Get-WmiObject Win32_Processor
    
        $cpu = foreach ($wmiProc in $wmiProcs) {
            $wmiProc | Select-Object -Property `
            @{n='CPU ID'; e={$_.DeviceID}},
            @{n='CPU Model'; e={$_.Name}},
            @{n='CPU Manufacturer'; e={If ($_.Manufacturer -like '*intel*') {'Intel'} Else {$_.Manufacturer}}},
            @{n='Number of Cores'; e={$_.NumberOfCores}},
            @{n='Number of Logical Processors'; e={$_.NumberOfLogicalProcessors}}
        }
        $cpuHTML = $cpu | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Storage Info
    $section = 'Storage Information'
    Transcript -Section $section

    try {
        $wmiDisks = Get-WmiObject -Class Win32_LogicalDisk
    
        $disks = foreach ($wmiDisk in $wmiDisks) {
            $wmiDisk | Select-Object -Property `
            @{n='Disk ID'; e={$_.DeviceID}},
            @{n='Volume Name'; e={$_.VolumeName}},
            @{n='Disk Size (GB)'; e={[int]($_.Size/1GB)}},
            @{n='Disk Free Space (GB)'; e={[int]($_.Freespace/1GB)}},
            @{n='Disk Free Space (%)'; e={[int]($_.Freespace*100/$_.Size)}}
        }
        $disksHTML = $disks | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Graphic Cards Info
    $section = 'Graphics Information'
    Transcript -Section $section

    try {
        $wmiGPUs = Get-WmiObject -Class Win32_VideoController

        $graphics = foreach ($wmiGPU in $wmiGPUs) {
            $wmiGPU | Select-Object -Property `
            @{n='Graphic Card Name'; e={$_.Name}},
            @{n='Resolution'; e={$_.VideoModeDescription}},
            @{Expression={[math]::Round($_.AdapterRAM / 1GB)};Label='Graphics Memory (GB)'},
            @{n='Graphic Card Status'; e={$_.Status}}
        }
        $graphicsHTML = $graphics | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Audio/Video Devices Info
    $section = 'Audio/Video Devices'
    Transcript -Section $section

    try {
        $section = 'WebCameras'
        Transcript -Section $section -SubHeader

        $webcam =  Get-WmiObject Win32_PnPEntity | Where-Object {($_.PNPClass -like 'camera') -or ($_.PNPClass -eq 'Image' -and $_.Name -like '*Camera*') -or ($_.PNPClass -eq 'Image' -and $_.Name -like '*Webcam*')} | Select-Object Name, Description, Manufacturer, Present, Status -Unique
        if ($webcam) {
            $webcamHTML = $webcam | ConvertTo-Html -As Table -Fragment -PreContent "$script:preContent $script:subHeader"
        } else {
            $webcamHTML = $script:preContent + $script:subHeader + $script:noData
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Speakers'
        Transcript -Section $section -SubHeader
           
        $speakerDevice =  Get-WmiObject Win32_PnPEntity | Where-Object {($_.PNPClass -like 'audio*') -and ($_.Name -like '*Speaker*' -or $_.Name -like 'Remote Audio Device')}| Select-Object Name, Description, Manufacturer, Present, Status -Unique
        if ($speakerDevice) {
        $speakerDeviceHTML = $speakerDevice | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader
        } else {
            $speakerDeviceHTML = $script:subHeader + $script:noData
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Microphones'
        Transcript -Section $section -SubHeader
                        
        $micDevice =   Get-WmiObject Win32_PnPEntity | Where-Object {($_.PNPClass -like 'audio*') -and ($_.Name -like '*Microphone*'-or $_.Name -like 'Remote Audio Device')} | Select-Object Name, Description, Manufacturer, Present, Status -Unique
        if ($micDevice) {
        $micDeviceHTML = $micDevice | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader -PostContent $script:postContent
        } else {
            $micDeviceHTML = $script:subHeader  + $script:noData + $script:postContent
        }        
        $audioVideoDevicesHTML = $webcamHTML + $speakerDeviceHTML + $micDeviceHTML
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Plug & Play Devices
    $section = 'Plug & Play Devices'
    Transcript -Section $section

    try {
        $pnpDevices = Get-PnpDevice | Where-Object {$_.friendlyname -ne ''} | Select-Object FriendlyName,Class,Status -Unique | Sort-Object Class
    
        $pnp = foreach ($pnpDevice in $pnpDevices) {
            $pnpDevice | Select-Object -Property `
            @{n='Friendly Name'; e={$pnpDevice.friendlyname}},
            @{n='Class'; e={$pnpDevice.Class}},
            @{n='Status'; e={$pnpDevice.Status}}
        }
        $pnpHTML = $pnp | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Driver Info
    $section = 'Driver Information'
    Transcript -Section $section

    try {
        $wmiOS = Get-WmiObject -Class Win32_OperatingSystem
        $wmiDrivers = Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.driverprovidername -notlike '' -or $_.devicename -notlike ''} | Sort-Object DeviceName
    
        $drivers = foreach ($wmiDriver in $wmiDrivers) {
            $wmiDriver | Select-Object -Property `
            @{n='Device Name'; e={$wmiDriver.devicename}},
            @{n='Driver Version'; e={$wmiDriver.driverversion}},
            @{n='Driver Date'; e={$wmiOS.ConvertToDateTime($_.DriverDate).ToString('dd-MM-yyyy')}},
            @{n='Driver Provider Name'; e={`
                If ($wmiDriver.driverprovidername -like '*intel*') {
                    'Intel'
                } ElseIf ($wmiDriver.driverprovidername -like '*realtek*') {
                    'Realtek'
                } Else {
                    $wmiDriver.driverprovidername
                }
            }}
        }
        $driversHTML = $drivers | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Services
    $section = 'Services Statuses'
    Transcript -Section $section
    try {
        $services = Get-WmiObject win32_service | Select-Object DisplayName, Name, StartMode, State | Sort-Object StartMode, State, DisplayName
        $servicesHTML = $services | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Group Policy Results Report
    $section = 'Resultant Set of Policies (GPOs / GPResult)'
    Transcript -Section $section
    try {
        $null = New-Item -Path "$logFolder\Windows" -ItemType Directory
        & "$env:windir\system32\gpresult.exe" /USER $currentLoggedUser /H "$logFolder\Windows\GPResult.html"
        $gpoHTML = $script:preContent + '<iframe src="LogFiles/Windows/GPResult.html" width="100%" height="500" style="border:none;"></iframe></div>'
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Installed Software and Versions
    $section = 'Installed Software and Versions'
    Transcript -Section $section

    try {
        $tempSW1 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -notlike ''} | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        $tempSW2 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -notlike ''} | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        $tempSW = $tempSW1 + $tempSW2
    
        if ($tempSW) {
            $installedApps = $tempSW | Where-Object {$_.InstallDate -notlike ''} | Select-Object `
                @{n='Name'; e={$_.DisplayName}},
                @{n='Version'; e={$_.DisplayVersion}},
                @{n='Publisher'; e={$_.Publisher}},
                @{n='Installed On'; e={Get-Date ([datetime]::parseexact($_.InstallDate, 'yyyyMMdd', $null)) -Format $culture.DateTimeFormat.UniversalSortableDateTimePattern}} -Unique | `
                Sort-Object Name
            $installedAppsHTML = $installedApps | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
        } else {
            $installedAppsHTML = $script:preContent + '<p>Not Software Installed</p>' + $script:postContent
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Windows Update Status
    $section = 'Windows Update Status'
    Transcript -Section $section

    try {
        $section = 'Hotfixes Installed'
        Transcript -Section $section -SubHeader
    
        $kbs = Get-HotFix

        $hotFixes =  foreach ($kb in $kbs) {
            $kb | Select-Object -Property `
            @{n='KB'; e={$kb.HotFixID}},
            @{n='KB Type'; e={$kb.Description}},
            @{n='Installed On'; e={Get-Date $kb.installedon -Format $culture.DateTimeFormat.UniversalSortableDateTimePattern}} ,
            @{n='Installed By'; e={$kb.InstalledBy}} | Sort-Object -Descending 'Installed On'
        }
        $hotFixesHTML = $hotFixes | ConvertTo-Html -As Table -Fragment -PreContent "$script:preContent $script:subHeader"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Pending Updates'
        Transcript -Section $section -SubHeader

        Internet
        if ($script:internet) {

            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $searchResults = $searcher.Search('IsInstalled=0')
            if ($searchResults) {
                if ($missingUpdates) {
                    $missingUpdates = $searchResults.RootCategories | ForEach-Object {
                        foreach ($update in $_.Updates) {
                            $update | Select-Object -Property `
                            @{n='KB'; e={[Regex]::Match($update.Title, '^.*\b(KB[0-9]+)\b.*$').Groups[1].Value}},
                            @{n='Category'; e={$_.Name}},
                            @{n='Title'; e={$update.Title}},
                            @{n='Type'; e={$update.Type}},
                            @{n='Downloaded'; e={$update.IsDownloaded}}
                        }
                    }
                    $missingUpdatesHTML = $missingUpdates | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader -PostContent $script:postContent
                } else {
                    $missingUpdatesHTML = $script:subHeader + '<p>No Updates Missing</p>' + $script:postContent
                }
            } else {
                $missingUpdatesHTML = $script:subHeader + $script:noData + $script:postContent
            }
        } else {
            $missingUpdatesHTML = $script:subHeader + $script:noData + $script:postContent
        }
        $windowsUpdateHTML = $hotFixesHTML + $missingUpdatesHTML
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Event Viewers Alerts Since Last Boot Time
    $section = 'Event Viewers Alerts Since Last Boot Time'
    Transcript -Section $section

    try {
        $tempLogsApp = Get-EventLog -LogName Application -EntryType 'Error','Warning' -After $bootTime.'Last Boot Time' | Select-Object -property @{n='Log'; e={'Application'}}, EventID, EntryType, Source, TimeGenerated, Message
        $tempLogsSys = Get-EventLog -LogName System -EntryType 'Error','Warning' -After $bootTime.'Last Boot Time' | Select-Object -property @{n='Log'; e={'System'}}, EventID, EntryType, Source, TimeGenerated, Message
        $tempLogs = $tempLogsApp + $tempLogsSys
        $eventLogs = $tempLogs | Select-Object -Property `
            @{n='Log'; e={$_.Log}},
            @{n='Event ID'; e={$_.EventID}},
            @{n='Event Type'; e={$_.EntryType}},
            @{n='Event Source'; e={$_.Source}},
            @{n='Time Generated'; e={Get-Date $_.TimeGenerated -Format $culture.DateTimeFormat.UniversalSortableDateTimePattern}},
            @{n='Message'; e={$_.Message}} -Unique | Sort-Object -Descending 'Time Generated'
        $eventLogsHTML = $eventLogs | ConvertTo-Html -As Table -Fragment -PreContent $script:preContent -PostContent $script:postContent
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get Performance Report
    $section = 'Performance Report'
    Transcript -Section $section

    try {
        $parameters2remove = 'delete -n Perf'
        $perfXML | Out-File $perfFolder\Perf.xml
        $parameters2import = "import -n Perf -xml $perfFolder\Perf.xml"
        $parameters2start = 'start -n Perf'
        Remove-Item -Recurse "$env:windir\Temp\LogCollector\*" -Force -ErrorAction Ignore
        Start-Process "$env:windir\system32\logman.exe" -WindowStyle hidden $parameters2remove -verb runas -Wait -ErrorAction Ignore
        Start-Process "$env:windir\system32\logman.exe" -WindowStyle hidden $parameters2import -verb runas -Wait
        Start-Process "$env:windir\system32\logman.exe" -WindowStyle hidden $parameters2start -verb runas -Wait
        Start-Sleep -Seconds 75 #Sleep for 60 seconds to run the Perfmon + 15 seconds to compile the data into the destination folder
        Start-Process "$env:windir\system32\logman.exe" -WindowStyle hidden $parameters2remove -verb runas -Wait
        Copy-Item -Path "$env:windir\Temp\LogCollector\*" -Destination "$perfFolder" -Recurse
        Remove-Item -Recurse "$env:windir\Temp\LogCollector\*" -Force
        Remove-Item $perfFolder\Perf.xml -Force
        $perfmonHTML = $script:preContent + '<iframe src="Performance/report.html" width="100%" height="500" style="border:none;"></iframe></div>'
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Get MDM Information
    $section = 'MDM Information'
    Transcript -Section $section

    try {
    
        $section = 'MDM Summary'
        Transcript -Section $section -SubHeader
        $MDMServer = Get-CimInstance -Namespace root\cimv2\mdm -ClassName MDM_MgmtAuthority -ErrorAction Ignore | Select-Object -ExpandProperty AuthorityName
        if ($MDMServer) {
            $mdmDiagParam = "-out $env:windir\Temp\LogCollector\MDM"
            Start-Process "$env:windir\system32\MdmDiagnosticsTool.exe" -WindowStyle hidden $mdmDiagParam -verb runas -Wait
            $MDMDiagFiles = Get-ChildItem "$env:windir\Temp\LogCollector\MDM" -ErrorAction Ignore
            if ($MDMDiagFiles) {
                $null = New-Item -Path "$logFolder\MDM" -ItemType Directory
                Copy-Item -Path "$env:windir\Temp\LogCollector\MDM\*" -Destination "$logFolder\MDM\" -Force
                $MDMDiagReportPath = Get-ChildItem "$env:windir\Temp\LogCollector\MDM\*.html" | Select-Object -ExpandProperty FullName -ErrorAction Ignore
            }
            $MDMSummary = New-Object PSObject
            Add-Member -inputObject $MDMSummary -memberType NoteProperty -name 'MDM Authority' -value $MDMServer
            if ($MDMDiagReportPath) {
                Add-Member -inputObject $MDMSummary -memberType ScriptProperty -name 'MDM Diagnostic Report' -value {Hyperlink -linkPath "$logFolder\MDM\MDMDiagReport.html" ; $script:htmlLink}
            } else {
                Add-Member -inputObject $MDMSummary -memberType NoteProperty -name 'MDM Diagnostic Report' -value 'Not Available'
            }
            $MDMDiagReportHTML = $MDMSummary | ConvertTo-Html -As List -Fragment -PreContent "$script:preContent $script:subHeader"

            $section = 'MDM Alerts'
            Transcript -Section $section -SubHeader
            
            if (Test-Path -Path "$env:windir\Temp\LogCollector\MDM\DeviceManagement-Enterprise-Diagnostics-Provider.evtx") {
                $MDMAlerts = Get-WinEvent -Path "$env:windir\Temp\LogCollector\MDM\DeviceManagement-Enterprise-Diagnostics-Provider.evtx" | Where-Object {$_.LevelDisplayName -ne 'Information'} | Select-Object -Property `
                @{n='Log'; e={'MDM'}},
                @{n='Event ID'; e={$_.Id}},
                @{n='Event Type'; e={$_.LevelDisplayName}},
                @{n='Time Generated'; e={Get-Date $_.TimeCreated -Format $culture.DateTimeFormat.UniversalSortableDateTimePattern}},
                @{n='Message'; e={$_.Message}} -Unique | Sort-Object -Descending 'Time Generated'
                $MDMAlertsHTML = $MDMAlerts | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader -PostContent $script:postContent
            } else {
                $MDMAlertsHTML = $script:subHeader + $script:noData + $script:postContent
            }
        } else {
            $MDMDiagReportHTML = $script:preContent + '<p>Not MDM Managed</p>' + $script:postContent
            $MDMAlertsHTML = ''
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Collect CBS Log
    $section = 'CBS Log'
    Transcript -Section $section

    try {
        If (Test-Path "$env:windir\Logs\CBS\CBS.log") {
            Copy-Item -Path "$env:windir\Logs\CBS\CBS.log" -Destination "$logFolder\Windows\" -Force
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Collect Windows Update Log
    $section = 'Windows Update Log'
    Transcript -Section $section
    try {
        $Job1 = Start-Job -ScriptBlock {Get-WindowsUpdateLog -LogPath "$logFolder\Windows\WindowsUpdate.log"}
        $Job1 | Wait-Job | Remove-Job
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Collect Event Viewer Logs from System and Application filtered based on events since last boot time
    $section = 'Event Viewer: Application Events Since Last Reboot'
    Transcript -Section $section

    try {
        $appLog = $logFolder + '\Windows\EventViewer - Application.evtx'
        & "$env:windir\system32\wevtutil.exe" epl Application $appLog "/q:*[System[TimeCreated[@SystemTime>='$bootTimestampForLogs' and @SystemTime<='$currentTimestampForLogs']]]"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    $section = 'Event Viewer: System Events Since Last Reboot'
    Transcript -Section $section

    try {
        $sysLog = $logFolder + '\Windows\EventViewer - System.evtx'
        & "$env:windir\system32\wevtutil.exe" epl System $sysLog "/q:*[System[TimeCreated[@SystemTime>='$bootTimestampForLogs' and @SystemTime<='$currentTimestampForLogs']]]"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Collect SCCM Logs
    $section = 'SCCM Logs'
    Transcript -Section $section

    try {
        If (Test-Path "$env:windir\CCM\Logs") {
            $null = mkdir "$logFolder\SCCM"
            Copy-Item -Path "$env:windir\CCM\Logs\*" -Destination "$logFolder\SCCM\" -Force
        } Else {
            Transcript -Section 'No SCCM logs available'
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Collect VMWare Logs
    $section = 'VMWare Logs'
    Transcript -Section $section

    try {
        $vmwareClientLogsPath = "$env:SystemDrive\Users\$currentLoggedUser\AppData\Local\VMware\VDM\Logs"
        $vmwareAgentLogsPath = "$env:ProgramData\VMware\VDM\logs"
        If (Test-Path $vmwareClientLogsPath) {
            $null = New-Item -Path "$logFolder\VMWareHorizon\Client\" -ItemType Directory
            Copy-Item -Path "$vmwareClientLogsPath\*" -Destination "$logFolder\VMWareHorizon\Client\" -Container -Recurse
        }
        Else {
            Transcript -Section 'No VMWare Horizon Client logs available'
        }
        If (Test-Path $vmwareAgentLogsPath) {
            $null = New-Item -Path "$logFolder\VMWareHorizon\Agent\" -ItemType Directory
            Copy-Item -Path "$vmwareAgentLogsPath\*" -Destination "$logFolder\VMWareHorizon\Agent\" -Container -Recurse
        }
        Else {
            Transcript -Section 'No VMWare Horizon Agent logs available'
        }
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Compile a List of Logs Collected
    $section = 'List of Logs Collected'
    Transcript -Section $section

    try {
        $section = 'Logs Available Offline'
        Transcript -Section $section -SubHeader
    
        $otherLogs = New-Object PSObject
        Add-Member -inputObject $otherLogs -memberType ScriptProperty -name 'Other Windows Logs' -value {If (Test-Path "$logFolder\Windows") {Hyperlink -linkPath "$logFolder\Windows" ; $script:htmlLink} Else {'Not Available'}}
        Add-Member -inputObject $otherLogs -memberType ScriptProperty -name 'MDM Logs' -value {If (Test-Path "$logFolder\MDM") {Hyperlink -linkPath "$logFolder\MDM" ; $script:htmlLink} Else {'Not Available'}}
        Add-Member -inputObject $otherLogs -memberType ScriptProperty -name 'SCCM Logs' -value {If (Test-Path "$logFolder\SCCM") {Hyperlink -linkPath "$logFolder\SCCM" ; $script:htmlLink} Else {'Not Available'}}
        Add-Member -inputObject $otherLogs -memberType ScriptProperty -name 'VMWare Logs' -value {If (Test-Path "$logFolder\VMWareHorizon") {Hyperlink -linkPath "$logFolder\VMWareHorizon" ; $script:htmlLink} Else {'Not Available'}}
        $otherLogsHTMl = $otherLogs | ConvertTo-Html -As List -Fragment -PreContent "$script:preContent $script:subHeader"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }

    try {
        $section = 'Full Memory Dumps (Only Listing - Not Collected Due to Their File Size)'
        Transcript -Section $section -SubHeader
        $fullDumpPath = "$env:SystemRoot\*.dmp"
        If (Test-Path $fullDumpPath) {
            $memoryDump = Get-ChildItem -Path $fullDumpPath | Select-Object Name,Directory,Length,CreationTime
            $memoryDumpHTML = $memoryDump | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader
        } Else {
            $memoryDumpHTML = $script:subHeader + '<p>No Full Memory Dump Stored</p>'
        }

        $section = 'Mini Memory Dumps (Only Listing - Not Collected Due to Their File Size)'
        Transcript -Section $section -SubHeader
        $miniDumpPath = "$env:SystemRoot\minidump\*.dmp"
        If (Test-Path $miniDumpPath) {
            $miniMemoryDump = Get-ChildItem -Path $miniDumpPath | Select-Object Name,Directory,Length,CreationTime
            $miniMemoryDumpHTML = $miniMemoryDump | ConvertTo-Html -As Table -Fragment -PreContent $script:subHeader -PostContent $script:postContent
        } Else {
            $miniMemoryDumpHTML = $script:subHeader + '<p>No Mini Memory Dumps Stored</p>' + $script:postContent
        }

        $LogsHTMl = $otherLogsHTMl + $memoryDumpHTML + $miniMemoryDumpHTML
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Generate HTML Report
    $section = 'Compiling HTML Report'
    Transcript -Section $section

    $section = 'Conditional Formatting'
    Transcript -Section $section
    $basicInfoHTML = $basicInfoHTML -replace 'True','<font color="red">True</font>'
    $localUsersHTML = $localUsersHTML -replace 'Disabled','<font color="red">Disabled</font>'
    $localUsersHTML = $localUsersHTML -replace 'Enabled','<font color="green">Enabled</font>'
    $localUsersHTML = $localUsersHTML -replace 'Yes','<b><font color="red">Yes</font></b>'
    $netHTML = $netHTML -replace 'GOOD','<b><font color="green">GOOD</font></b>'
    $netHTML = $netHTML -replace 'BAD','<b><font color="red">BAD</font></b>'
    $netHTML = $netHTML -replace 'Test Failed','<font color="red">Test Failed</font>'
    $pnpHTML = $pnpHTML -replace 'OK','<font color="green">OK</font>'
    $pnpHTML = $pnpHTML -replace 'Unknown','<font color="orange">Unknown</font>'
    $pnpHTML = $pnpHTML -replace 'Error','<font color="red">Error</font>'
    $servicesHTML = $servicesHTML -replace 'Stopped','<font color="red">Stopped</font>'
    $servicesHTML = $servicesHTML -replace 'Running','<font color="green">Running</font>'
    $eventLogsHTML = $eventLogsHTML -replace 'Error','<font color="red">Error</font>'
    $eventLogsHTML = $eventLogsHTML -replace 'Warning','<font color="Orange">Warning</font>'
    $MDMDiagReportHTML = $MDMDiagReportHTML -replace 'START','<a target="_blank" href="'
    $MDMDiagReportHTML = $MDMDiagReportHTML -replace 'MIDDLE','">'
    $MDMDiagReportHTML = $MDMDiagReportHTML -replace 'END','</a>'
    $MDMAlertsHTML = $MDMAlertsHTML -replace 'Error','<font color="red">Error</font>'
    $MDMAlertsHTML = $MDMAlertsHTML -replace 'Warning','<font color="Orange">Warning</font>'
    $LogsHTMl = $LogsHTMl -replace 'START','<a target="_blank" href="'
    $LogsHTMl = $LogsHTMl -replace 'MIDDLE','">'
    $LogsHTMl = $LogsHTMl -replace 'END','</a>'

    try {
        $section = 'Assembling HTML Skeleton'
        Transcript -Section $section
    
        #Putting all Together
        $body = "$header $basicInfoHTML $localUsersHTML $netHTML $cpuHTML $disksHTML $graphicsHTML $audioVideoDevicesHTML $pnpHTML $driversHTML $servicesHTML $gpoHTML $installedAppsHTML $windowsUpdateHTML $eventLogsHTML $perfmonHTML $MDMDiagReportHTML $MDMAlertsHTML $LogsHTMl"
        $footer = "<p id='Footer'>Report Creation Date: $(Get-Date -Format 'dd-MM-yyyy HH:mm')<br>$footer</p>"
 
        $section = 'Generating Final HTML'
        Transcript -Section $section
       
        ConvertTo-HTML -Title $mainTitle -Head $cssStyles -Body $body -PostContent "$footer $cssScript" |  Out-File "$rootPath\Report - $env:computername - $currentDateTime.html"
    
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Stop Transcript: Must be stopped at this point in order to allow creation of ZIP archive
    Transcript -Section 'Finishing'
    Stop-Transcript
    #endregion

    #-------------------------------------------------------------------------------------------
    #region Compressing all collected files and creating the ZIP archive
    try {
        $section = 'Compiling Results'
        Transcript -Section $section
        $zipName = Get-Item $rootPath | Select-Object -ExpandProperty Name
        Start-Sleep -Seconds 5 #Just to make sure the script has enough time to close all handles
        Compress-Archive -Path $rootPath -DestinationPath "$rootPath\$zipName.zip" -Force
        Transcript -Section "Zip Archive Created at: $rootPath\$zipName.zip"
    } catch {
        Transcript -Section "Failed: $section `n$_"
    }
    #endregion
    Transcript -Section 'Completed'
}

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Required Assemblies & Preparation

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$defaultLogo = @'
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
'@
$customLogo = $logoFile
if (Test-Path $customLogo) {
    $logoBase64 = [convert]::ToBase64String((get-content $customLogo -encoding byte))
} else {
    $logoBase64 = $defaultLogo
}
$customHexMainBackColor = $backgroundColor
$customHexMainFontColor = $foregroundColor

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Create GUI Interface

Add-Type -AssemblyName System.Windows.Forms
$objForm = New-Object system.Windows.Forms.Form
$objForm.Text = "$mainTitle - Ultimate Log Collector"
$objForm.Width = 523
$objForm.Height = 520
$objForm.MinimizeBox = $False
$objForm.MaximizeBox = $False
$objForm.FormBorderStyle = 'Fixed3D'
$objForm.StartPosition = 'CenterScreen'

$Icon = [drawing.icon]::ExtractAssociatedIcon($PSHOME + '\powershell.exe')
$objForm.Icon = $Icon

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Title and Logo
$img = [Drawing.Bitmap]::FromStream([IO.MemoryStream][Convert]::FromBase64String($logoBase64))
[Windows.Forms.Application]::EnableVisualStyles();
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = '340,5'
$pictureBox.Width =  $img.Size.Width
$pictureBox.Height =  $img.Size.Height
$pictureBox.Image = $img
$objform.controls.add($pictureBox)

$fontH1 = New-Object System.Drawing.Font('Segoe UI',22,[Drawing.FontStyle]::Bold)
$h1 = New-Object System.Windows.Forms.Label
$h1.ForeColor = $customHexMainBackColor
$h1.Location = '25,25'
$h1.Font = $fontH1
$h1.Text = $mainTitle
$h1.AutoSize = $True
$objForm.Controls.Add($h1)

$fontH2 = New-Object System.Drawing.Font('Segoe UI',16,[Drawing.FontStyle]::Regular)
$h2 = New-Object System.Windows.Forms.Label
$h2.ForeColor = 'Gray'
$h2.Location = '25,75'
$h2.Font = $fontH2
$h2.Text = 'Ultimate LOG Collector'
$h2.AutoSize = $True
$objForm.Controls.Add($h2)

$fontH3 = New-Object System.Drawing.Font('Verdana',12,[Drawing.FontStyle]::Bold)
$h3 = New-Object System.Windows.Forms.Label
$h3.ForeColor = 'Red'
$h3.Location = New-Object System.Drawing.Point(15,240)
$h3.Font = $fontH3
$h3.Text = 'This process may take up to 10 minutes to complete'
$h3.AutoSize = $True
$objForm.Controls.Add($h3)

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq 'Enter') {Run}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq 'Escape') {$objForm.Close()}})

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Log Progress Bar and Output

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 180)
$progressBar.Size = New-Object System.Drawing.Size(480, 40)
$progressBar.Style = 'Marquee'
$progressBar.MarqueeAnimationSpeed = 20
$progressBar.Hide()
$objForm.Controls.Add($progressBar)

$fontLogOutput = New-Object System.Drawing.Font('Arial',9,[Drawing.FontStyle]::Regular)
$logOutput = [hashtable]::Synchronized(@{})
$logOutput = New-Object System.Windows.Forms.TextBox
$logOutput.Location = New-Object System.Drawing.Size(10,245) 
$logOutput.Size = New-Object System.Drawing.Size(480,115)
$logOutput.Multiline = $True
$logOutput.ScrollBars = 'Vertical'
$logOutput.AcceptsReturn = $True
$logOutput.WordWrap = $True
$logOutput.Font = $fontLogOutput
$logOutput.ReadOnly = $True
$logOutput.Text = ''
$logOutput.Hide()
$objForm.Controls.Add($logOutput)

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Buttons

$lineButtons = New-Object System.Windows.Forms.Label
$lineButtons.ForeColor = 'Gray'
$lineButtons.Location = '5,383'
$lineButtons.Height = 2
$lineButtons.Width = 490
$lineButtons.BorderStyle = 'Fixed3D'
$lineButtons.AutoSize = $False
$objForm.Controls.Add($lineButtons)

$buttonRun = New-Object System.Windows.Forms.Button 
$buttonRun.Location = New-Object System.Drawing.Size(50,400) 
$buttonRun.Size = New-Object System.Drawing.Size(150,30) 
$buttonRun.Text = 'Run' 
$buttonRun.Add_Click({Run}) 
$objForm.Controls.Add($buttonRun)

$buttonQuit = New-Object System.Windows.Forms.Button
$buttonQuit.Location = New-Object System.Drawing.Point(290,400)
$buttonQuit.Size = New-Object System.Drawing.Size(150,30)
$buttonQuit.Text = 'Close'
$buttonQuit.add_Click({Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Undefined -Force ; $objform.Close()})
$objForm.Controls.Add($buttonQuit)

$buttonAbort = New-Object System.Windows.Forms.Button
$buttonAbort.Location = New-Object System.Drawing.Size(170,400)
$buttonAbort.Size = New-Object System.Drawing.Size(150,30)
$buttonAbort.Text = 'Abort'
$buttonAbort.add_Click({Abort})
$buttonAbort.Hide()
$objForm.Controls.Add($buttonAbort)

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Footer

$lineFooter = New-Object System.Windows.Forms.Label
$lineFooter.ForeColor = 'Gray'
$lineFooter.Location = '5,445'
$lineFooter.Height = 2
$lineFooter.Width = 490
$lineFooter.BorderStyle = 'Fixed3D'
$lineFooter.AutoSize = $False
$objForm.Controls.Add($lineFooter)

$fontFooter = New-Object System.Drawing.Font('Verdana',8,[Drawing.FontStyle]::Regular)
$labelFooter = New-Object System.Windows.Forms.Label
$labelFooter.ForeColor = 'Gray'

$labelFooter.Font = $fontFooter
$labelFooter.Text = "$footer`n "
$labelFooter.AutoSize = $False
$labelFooter.TextAlign = 'BottomCenter'
$labelFooter.Dock = 'Fill'
$objForm.Controls.Add($labelFooter)

#endregion

#-------------------------------------------------------------------------------------------
#region FORM: Activate the Form

$objForm.Topmost = $False
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()

#endregion


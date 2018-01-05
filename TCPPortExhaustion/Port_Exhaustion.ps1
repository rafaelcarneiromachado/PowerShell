<# 

File Name: Port_Exhaustion.ps1   
Version: 1.0 
Author: Rafael Carneiro Machado
E-Mail: rafaelcarneiromachado@gmail.com 
Web: https://www.linkedin.com/in/rafaelcarneiromachado/

.SYNOPSIS 
 
        Test/Log TCP Connections in order to identify Port Exhaustion.

.DESCRIPTION
        
        A LOG file will be created at C:\Temp\PortExhaustion[DATE/TIME].
#>


#region ::Format the log file

If (Test-Path C:\Temp) {} Else { New-Item C:\Temp -ItemType Directory }
    Get-ChildItem C:\Temp\PortExhaustion*.txt | Sort-Object CreationTime -Descending | Select -Skip 32 | Remove-Item -Force
    $Date = Get-Date -Format g
    $DateLog = Get-Date -Format MMddyyyy\THHmmss
    $LogName = "PortExhaustion_$DateLog.txt"
"============================================================================================================================" | Out-File C:\Temp\$LogName
"                                             PORT EXHAUSTION TOOL ($Date)                                                   " | Out-File C:\Temp\$LogName -Append
"============================================================================================================================" | Out-File C:\Temp\$LogName -Append
#endregion


#region ::Get TCP Port Usage Summary by IP Addresses

<#

IMPORTANT: This part of the script was copied from https://blogs.msdn.microsoft.com/debuggingtoolbox/2010/10/11/powershell-script-troubleshooting-for-port-exhaustion-using-netstat/
Author: frank.taglianetti@microsoft.com

#>

function MaxNumOfTcpPorts  #helper function to retrive number of ports per address
{
param 
    (
        [parameter(Mandatory=$true)]
         $tcpParams
    )
    #  Returns the maximum number of ports per TCP address
    #  Check for Windows Vista and later
    $IsVistaOrLater = Get-WmiObject -Class Win32_OperatingSystem | %{($_.Version -match "6\.\d+")}
    if($isVistaOrLater)
    {
        # Use netsh to retrieve the number of ports and parse out the string of numbers after "Number of Ports : "
        $maxPorts = netsh int ip show dynamicport tcp |
            Select-String -Pattern "Number of Ports : (\d*)"|
            %{$_.matches[0].Groups[1].Value}
        # Convert string to integer
        $maxPorts = [int32]::Parse($maxPorts)
        #  modify the PSCustomObject to simulate the MaxUserPort value for printout
        Add-Member -InputObject $tcpParams -MemberType NoteProperty -Name MaxUserPort -Value $maxPorts 
    }
    else  # this is Windows XP or older
    {
        # check of emphermal ports modified in registry
        $maxPorts = $($tcpParams | Select-Object MaxUserPort).MaxUserPort
        if($maxPorts -eq $null)
        {
            $maxPorts = 5000 - 1kb    #Windows Default range is from 1025 to 5000 inclusive
            Add-Member -InputObject $tcpParams -MemberType NoteProperty -Name MaxUserPort -Value $maxPorts
        }
    }
    return $maxPorts
}
function New-Port  # helper function to track number of ports per IP address
{
    Param
    (
        [string] $IPAddress = [String]::EmptyString,
        [int32] $PortsWaiting = 0,
        [int32] $MaxUserPort = 3976
    )

    $newPort = New-Object PSObject

    Add-Member -InputObject $newPort -MemberType NoteProperty -Name IPAddress -Value $IPAddress
    Add-Member -InputObject $newPort -MemberType NoteProperty -Name PortsUsed -Value 1
    Add-Member -InputObject $newPort -MemberType ScriptProperty -Name PercentUsed -Value {$this.PortsUsed / $this.MaxUserPort}
    Add-Member -InputObject $newPort -MemberType NoteProperty -Name PortsWaiting -Value $portsWaiting
    Add-Member -InputObject $newPort -MemberType ScriptProperty -Name PercentWaiting -Value {$this.PortsWaiting / [Math]::Max(1,$this.PortsUsed)}
    Add-Member -InputObject $newPort -MemberType NoteProperty -Name MaxUserPort -Value $maxUserPort
    return $newPort
}

######################### Beginning of the main routine ##########################

# Store MaxUserPort for percentage used calculations
$tcpParams = Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters
$maxPorts = MaxNumOfTcpPorts($tcpParams)   # call function to return max # ports as per OS version
$tcpTimedWaitDelay = $($tcpParams | Select-Object TcpTimedWaitDelay).TcpTimedWaitDelay

if($tcpTimedWaitDelay -eq $Null)           #Value wasn't configured in registry
{
    $tcpTimedWaitDelay = 240               #Default Value if registry value doesn't exist
    Add-Member -InputObject $tcpParams -MemberType NoteProperty -Name TcpTimedWaitDelay -Value $tcpTimedWaitDelay  #fake reg value for output
}

# Display the MaxUserPort and TcpTimedWaitDelay settings in the registry if available
"                                                    CURRENT TCP PARAMETERS                                                  " | Out-File C:\Temp\$LogName -Append
$tcpParams | Format-List MaxUserPort,TcpTimedWaitDelay | Out-File C:\Temp\$LogName -Append

# collection of IP Address and port counts
[System.Collections.HashTable] $ports = New-Object System.Collections.HashTable

[int32] $intWait = 0

netstat -an | 
Select-String "TCP\s+.+\:.+\s+(.+)\:(\d+)\s+(\w+)" | 
ForEach-Object {
    $key = $_.matches[0].Groups[1].value      # use the IP address as hash key
    $Status = $_.matches[0].Groups[3].value   # Last group contains port status
    if("TIME_WAIT" -like $Status)
    {
        $intWait = 1                          # incr count
    }
    else
    {
        $intWait = 0                          # don't incr count
    }
    if(-not $ports.ContainsKey($key))         #IP Address not yet counted
    {
        $port = New-Port -IPAddress $key -PortsWaiting $intWait -MaxUserPort $maxPorts    #intialize new tracking object
        $ports.Add($key,$port)                #Add the tracking object to hashtable
    }
    else                                      #otherwise a tracking object exists for this IP
    {
        $port = $ports[$key]                  #retrieve the tracking object
        $port.PortsUsed ++                    # increment the port count (PortsUsed)
        $port.PortsWaiting += $intWait        # increment PortsWaiting if status is TIME_WAIT
    }
}

 
 
#Format-Table -InputObject $ports.Values -auto

"----------------------------------------------------------------------------------------------------------------------------" | Out-File C:\Temp\$LogName -Append
"                                                    PORT USAGE SUMMARY                                                      " | Out-File C:\Temp\$LogName -Append
$ports.Values | 
    Sort-Object -Property PortsUsed, PortsWaiting -Descending  |
    Format-Table -Property IPAddress,PortsWaiting,
        @{Name='%Waiting';Expression ={"{0:P}" -f $_.PercentWaiting};Alignment="Right"},
        PortsUsed,
        @{Name='%Used';Expression ={"{0:P}" -f $_.PercentUsed}; Alignment="Right"} -Auto | Out-File C:\Temp\$LogName -Append

Remove-Variable -Name "ports"
#endregion


#region ::Retrieve all the TCP Connections (equivalent to netstat, sorted by LocalPort)
"----------------------------------------------------------------------------------------------------------------------------" | Out-File C:\Temp\$LogName -Append
"                                                    TCP CONNECTIONS                                                         " | Out-File C:\Temp\$LogName -Append
Get-NetTCPConnection | Select LocalAddress, LocalPort, RemoteAddress, RemotePort, State, @{l="ProcessID";e={$_.Owningprocess}}, @{l="ProcessName";e={(get-process -ID $_.Owningprocess).processname}} | Sort LocalPort, ProcessID | ft -AutoSize | Out-File C:\Temp\$LogName -Append
#endregion


#region ::Count TCP Connections per Process Name
"----------------------------------------------------------------------------------------------------------------------------" | Out-File C:\Temp\$LogName -Append
"                                               TCP CONNECTIONS/PROCESS NAME                                                 " | Out-File C:\Temp\$LogName -Append
Get-NetTCPConnection | Select @{l="ProcessName";e={(get-process -ID $_.Owningprocess).processname}} | Group-Object ProcessName -NoElement | sort -descending Count | ft -AutoSize | Out-File C:\Temp\$LogName -Append
#endregion


#region ::Retrieve all the TCP connections for the top 5 Processes
"----------------------------------------------------------------------------------------------------------------------------" | Out-File C:\Temp\$LogName -Append
"                                              TCP CONNECTIONS/TOP 5 PROCESSES                                               " | Out-File C:\Temp\$LogName -Append
$top5process = Get-NetTCPConnection | Select @{l="ProcessName";e={(get-process -ID $_.Owningprocess).processname}} | Group-Object ProcessName -NoElement | sort -descending Count | select -first 5
Get-NetTCPConnection | Select LocalAddress, LocalPort, RemoteAddress, RemotePort, State, @{l="ProcessID";e={$_.Owningprocess}}, @{l="ProcessName";e={(get-process -ID $_.Owningprocess).processname}} | ? {$_.Processname -in (($top5process).name)} | sort ProcessName | ft -GroupBy ProcessName -AutoSize  | Out-File C:\Temp\$LogName -Append
#endregion


#region ::Retrieve Event Viewer logs (System) from the past 3 hours
"----------------------------------------------------------------------------------------------------------------------------" | Out-File C:\Temp\$LogName -Append
"                                     GET EVENT VIEWER LOGS (SYSTEM) FROM THE LAST 3 HOURS                                   " | Out-File C:\Temp\$LogName -Append
Get-EventLog SYSTEM -After (Get-Date).AddHours(-3) | Select-Object TimeGenerated, EventID, MachineName, EntryType, Source, UserName, Message | Format-Table -AutoSize | Out-File C:\Temp\$LogName -Append -Width 5000
#endregion

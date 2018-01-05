<# 

File Name: Test-SMB-Path.ps1   
Version: 1.0 
Author: Rafael Carneiro Machado
E-Mail: rafaelcarneiromachado@gmail.com 
Web: https://www.linkedin.com/in/rafaelcarneiromachado/

.SYNOPSIS 
 
        Test Share/SMB availability.
#>

#Fixed Values
$strHostname = hostname
$strOutputFile = "C:\Windows\Temp\Test-SMB-Path.log"

#Variables
$strPath = "\\Server\Share" #>> Type the share as \\SERVER\SHARE
$strInterval = 1 #>> Type an interval in seconds to monitor the Share
$strOutputMode = 1 #>> Type "1" for Console, Type 2 for Text/Log File

While($True) {
    If ($strOutputMode -eq 1) {
        If (Test-Path $strPath) {Write-Host (Get-Date -Format G) "| INFO | Path $strPath is reacheable from Host $strHostname" -foregroundcolor "DarkGreen"} Else {Write-Host (Get-Date -Format G) "| ERROR | Path $strPath is NOT reacheable from Host $strHostname" -foregroundcolor "Red"}
    }
    If ($strOutputMode -eq 2) {
        If (Test-Path $strPath) {$(Get-Date -Format G) + " | INFO | Path $strPath is reacheable from Host $strHostname" | Out-File $strOutputFile -Append} Else {$(Get-Date -Format G) + " | ERROR | Path $strPath is NOT reacheable from Host $strHostname" | Out-File $strOutputFile -Append}
    }
    Start-Sleep -Seconds $strInterval
}

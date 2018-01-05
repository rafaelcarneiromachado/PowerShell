<# 

File Name: BeHolder_Code_v5.ps1   
Version: 5.0 
Author: Rafael Carneiro Machado
E-Mail: rafaelcarneiromachado@gmail.com 
Web: https://www.linkedin.com/in/rafaelcarneiromachado/

.SYNOPSIS 
 
        User friendly GUI to set performance counters manually or based on templates.

.DESCRIPTION/HOW TO USE
        
        1. Unzip all files in C:\Temp\BeHolder
        2. Make sure you have the following files inside the C:\Temp\BeHolder folder:
                a) C:\Temp\BeHolder\BeHolder_Code_v5.ps1
                b) C:\Temp\BeHolder\beholder4.png
                c) C:\Temp\BeHolder\Templates\ (All PerfMon templates that you want, using .XML extension)
        3. Right click on C:\Temp\BeHolder\BeHolder_Code_v5.ps1 and select "Run with PowerShell" (If does not work, run the script directly from an ELEVATED (AS ADMINISTRATOR) PowerShell prompt;
#>

#-------------------------------------------------------------------------------------------
#Get/Set Execution Policy
#-------------------------------------------------------------------------------------------
Set-ExecutionPolicy -Scope CurrentUser Bypass -Force

#-------------------------------------------------------------------------------------------
#Get path of BeHolder
#-------------------------------------------------------------------------------------------
$Invocation = (Get-Variable MyInvocation -Scope 0).Value
$DefaultPath = Split-Path $Invocation.MyCommand.Path
$DefaultPath | Out-File c:\temp\test.txt

#-------------------------------------------------------------------------------------------
#Run as Administrator
#-------------------------------------------------------------------------------------------
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   $Host.UI.RawUI.BackgroundColor = "Gray"
   $Host.UI.RawUI.ForegroundColor = "Black"
   clear-host
   }
else
   {
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   $newProcess.Verb = "runas";
   [System.Diagnostics.Process]::Start($newProcess)
   exit
   }

#-------------------------------------------------------------------------------------------
#Create GUI Interface
#-------------------------------------------------------------------------------------------

    Add-Type -AssemblyName System.Windows.Forms
    $objForm = New-Object system.Windows.Forms.Form
    $objForm.Text = "BeHolder - Performance Tool"

    $objForm.Width = 1170
    $objForm.Height = 465
    $objForm.MinimizeBox = $True
    $objForm.MaximizeBox = $False
    $objForm.FormBorderStyle = 'Fixed3D'
    $objForm.StartPosition = "CenterScreen"

    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $objForm.Icon = $Icon

#-------------------------------------------------------------------------------------------
#Title and Logo
#-------------------------------------------------------------------------------------------

    $file = (get-item $DefaultPath\beholder4.png)
    $img = [System.Drawing.Image]::Fromfile($file);
    [System.Windows.Forms.Application]::EnableVisualStyles();
    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.Location = '1,1'
    $pictureBox.Width =  $img.Size.Width;
    $pictureBox.Height =  $img.Size.Height;
    $pictureBox.Image = $img;
    $objform.controls.add($pictureBox)

    $Font = New-Object System.Drawing.Font("Verdana",18,[System.Drawing.FontStyle]::Regular)
    $Label = New-Object System.Windows.Forms.Label
    $Label.ForeColor = "Gray"
    $Label.Location = '565,30'
    $Label.Font = $Font
    $Label.Text = "BeHolder"
    $Label.AutoSize = $True
    $objForm.Controls.Add($Label)

    $Font2 = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
    $Label2 = New-Object System.Windows.Forms.Label
    $Label2.ForeColor = "Gray"
    $Label2.Location = '515,65'
    $Label2.Font = $Font2
    $Label2.Text = "Tool to set Performance Counters"
    $Label2.AutoSize = $True
    $objForm.Controls.Add($Label2)

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$x=$objListBox.SelectedItem;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close()}})

#-------------------------------------------------------------------------------------------
#First Block (OPTIONS)
#-------------------------------------------------------------------------------------------

    $Line1 = New-Object System.Windows.Forms.Label
    $Line1.ForeColor = "Gray"
    $Line1.Location = '5,100'
    $Line1.Height = 2
    $Line1.Width = 1150
    $Line1.BorderStyle = "Fixed3D"
    $Line1.AutoSize = $False
    $objForm.Controls.Add($Line1)

    $NameText = New-Object System.Windows.Forms.Label
    $NameText.Location = '8,115'
    $NameText.AutoSize = $True
    $NameText.Text = "Counter Set Name:"
    $NameControl = New-Object System.Windows.Forms.TextBox
    $NameControl.Location = New-Object System.Drawing.Size(10,130) 
    $NameControl.Size = New-Object System.Drawing.Size(250,220)
    $NameControl.Text = "Type a Name"
    $objForm.Controls.Add($NameControl)
    $objForm.Controls.Add($NameText)

    $LogLocationText = New-Object System.Windows.Forms.Label
    $LogLocationText.Location = '298,115'
    $LogLocationText.AutoSize = $True
    $LogLocationText.Text = "Path for Log(s):"
    $LogLocationControl = New-Object System.Windows.Forms.TextBox
    $LogLocationControl.Location = New-Object System.Drawing.Size(300,130) 
    $LogLocationControl.Size = New-Object System.Drawing.Size(350,220)
    $LogLocationControl.Text = "C:\PerfLogs\Admin"
    $objForm.Controls.Add($LogLocationControl)
    $objForm.Controls.Add($LogLocationText)

    $IntervalText = New-Object System.Windows.Forms.Label
    $IntervalText.Location = '930,115'
    $IntervalText.AutoSize = $True
    $IntervalText.Text = "Interval (Seconds):"
    $IntervalControl = New-Object System.Windows.Forms.TextBox
    $IntervalControl.Location = New-Object System.Drawing.Size(955,130) 
    $IntervalControl.Size = New-Object System.Drawing.Size(50,220)
    $IntervalControl.TextAlign = "Center"
    $IntervalControl.Text = "15"
    $objForm.Controls.Add($IntervalControl)
    $objForm.Controls.Add($IntervalText)

    $FrequencyText = New-Object System.Windows.Forms.Label
    $FrequencyText.Location = '688,115'
    $FrequencyText.AutoSize = $True
    $FrequencyText.Text = "Run For:"
    $FrequencyControl = New-Object System.Windows.Forms.ComboBox
    $FrequencyControl.Location = New-Object System.Drawing.Size(690,130) 
    $FrequencyControl.Size = New-Object System.Drawing.Size(200,220)
    $FrequencyControl.Items.AddRange(("Stop Manually", "60 Seconds", "5 Minutes", "15 Minutes", "1 Hour", "4 Hours", "8 Hours", "1 Day", "1 Week", "1 Month"))
    $FrequencyControl.SelectedText = "Stop Manually"
    $objForm.Controls.Add($FrequencyControl)
    $objForm.Controls.Add($FrequencyText)

    $CheckboxText = New-Object System.Windows.Forms.Label
    $CheckboxText.Location = '1077,115'
    $CheckboxText.AutoSize = $True
    $CheckboxText.Text = "Start Now?"
    $Checkbox = New-Object System.Windows.Forms.Checkbox
    $Checkbox.Location = New-Object System.Drawing.Size(1100,130) 
    $Checkbox.Size = New-Object System.Drawing.Size(20,20)
    $Checkbox.TabIndex = 3
    $objForm.Controls.Add($Checkbox)
    $objForm.Controls.Add($CheckboxText)

#-------------------------------------------------------------------------------------------
#Second Block (TEMPLATES/OPERATION MODE)
#-------------------------------------------------------------------------------------------

    $Line = New-Object System.Windows.Forms.Label
    $Line.ForeColor = "Gray"
    $Line.Location = '5,160'
    $Line.Height = 2
    $Line.Width = 1150
    $Line.BorderStyle = "Fixed3D"
    $Line.AutoSize = $False
    $objForm.Controls.Add($Line)

    $Font2 = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
    $LabelTemplates = New-Object System.Windows.Forms.Label
    $LabelTemplates.ForeColor = "Black"
    $LabelTemplates.Location = '453,173'
    $LabelTemplates.Font = $Font2
    $LabelTemplates.Text = "..:: Set Based on a Template ::.."
    $LabelTemplates.AutoSize = $True
    $objForm.Controls.Add($LabelTemplates)

    #-------------------------------------------------------------------------------------------
    #Set Operation Mode

    $groupBoxOperationMode = New-Object System.Windows.Forms.GroupBox
    $groupBoxOperationMode.Location = New-Object System.Drawing.Size(30,210)
    $groupBoxOperationMode.size = New-Object System.Drawing.Size(140,100)
    $groupBoxOperationMode.text = "Operation Mode:"
    $objForm.Controls.Add($groupBoxOperationMode)

    $OperationModeText1 = New-Object System.Windows.Forms.Label
    $OperationModeText1.Location = '50,30'
    $OperationModeText1.AutoSize = $True
    $OperationModeText1.Text = "Basic"
    $OperationModeControl1 = New-Object System.Windows.Forms.RadioButton
    $OperationModeControl1.Location = New-Object System.Drawing.Size(30,27) 
    $OperationModeControl1.Size = New-Object System.Drawing.Size(20,20)
    $OperationModeControl1.TabIndex = 3
    $OperationModeControl1.Checked = "$True"
    $OperationModeControl1.Add_Click({Basic})
    $groupBoxOperationMode.Controls.Add($OperationModeControl1)
    $groupBoxOperationMode.Controls.Add($OperationModeText1)

    $OperationModeText2 = New-Object System.Windows.Forms.Label
    $OperationModeText2.Location = '50,60'
    $OperationModeText2.AutoSize = $True
    $OperationModeText2.Text = "Expert"
    $OperationModeControl2 = New-Object System.Windows.Forms.RadioButton
    $OperationModeControl2.Location = New-Object System.Drawing.Size(30,57) 
    $OperationModeControl2.Size = New-Object System.Drawing.Size(20,20)
    $OperationModeControl2.TabIndex = 3
    $OperationModeControl2.Add_Click({Expert})
    $groupBoxOperationMode.Controls.Add($OperationModeControl2)
    $groupBoxOperationMode.Controls.Add($OperationModeText2)

    Function Expert {
        $ListBox0.Size = New-Object System.Drawing.Size(890,70)
        $groupBox0.Controls.Add($ButtonCountersTemplates1)
        $groupBox0.Controls.Add($ButtonCountersTemplates2)
        $groupBox0.Controls.Add($ButtonCountersTemplates3)
        $objForm.Controls.Add($Line3)
        $objForm.Controls.Add($LabelCustomCounters)
        $objForm.Controls.Add($groupBox)
        $objForm.Controls.Add($groupBox2)
        $objForm.Controls.Add($groupBox3)
        $objForm.Controls.Add($Button3)
        $objForm.Controls.Add($Button4)        
        $Line4.Location = '5,600'
        $ButtonSet.Location = New-Object System.Drawing.Size(70,610) 
        $ButtonStop.Location = New-Object System.Drawing.Size(290,610) 
        $ButtonPerfMon.Location = New-Object System.Drawing.Size(510,610) 
        $ButtonLog.Location = New-Object System.Drawing.Size(730,610) 
        $buttonQuit.Location = New-Object System.Drawing.Size(950, 610)
        $Line5.Location = '5,650'
        $LabelFooter.Location = '380,657'
        $objForm.Height = 715
        $objForm.Refresh()
    }

    Function Basic {
        $ListBox0.Size = New-Object System.Drawing.Size(890,107)
        $groupBox0.Controls.Remove($ButtonCountersTemplates1)
        $groupBox0.Controls.Remove($ButtonCountersTemplates2)
        $groupBox0.Controls.Remove($ButtonCountersTemplates3)
        $objForm.Controls.Remove($Line3)
        $objForm.Controls.Remove($LabelCustomCounters)
        $objForm.Controls.Remove($groupBox)
        $objForm.Controls.Remove($groupBox2)
        $objForm.Controls.Remove($groupBox3)
        $objForm.Controls.Remove($Button3)
        $objForm.Controls.Remove($Button4)
        $Line4.Location = '5,350'
        $ButtonSet.Location = New-Object System.Drawing.Size(70,360) 
        $ButtonStop.Location = New-Object System.Drawing.Size(290,360) 
        $ButtonPerfMon.Location = New-Object System.Drawing.Size(510,360) 
        $ButtonLog.Location = New-Object System.Drawing.Size(730,360) 
        $buttonQuit.Location = New-Object System.Drawing.Size(950,360)
        $Line5.Location = '5,400'
        $LabelFooter.Location = '380,407'
        $objForm.Height = 465
        $objForm.Refresh()
    }

    #-------------------------------------------------------------------------------------------

    $groupBox0 = New-Object System.Windows.Forms.GroupBox
    $groupBox0.Location = New-Object System.Drawing.Size(200,195)
    $groupBox0.size = New-Object System.Drawing.Size(900,130)
    $groupBox0.text = "Select a Template:"
    $objForm.Controls.Add($groupBox0)

    $ListBox0 = New-Object System.Windows.Forms.ListBox 
    $ListBox0.Location = New-Object System.Drawing.Size(5,20) 
    $ListBox0.Size = New-Object System.Drawing.Size(890,107) 
    get-childitem $DefaultPath\Templates\*.xml | foreach {$_.name} | Out-File $DefaultPath\Templates\Templates.txt -force
    Get-Content $DefaultPath\Templates\Templates.txt | ForEach-Object {[void] $ListBox0.Items.Add($_)}
    $groupBox0.Controls.Add($ListBox0)

    Function GetTemplateCounters {
        If ($ListBox0.SelectedIndex.Equals(-1)) {[System.Windows.Forms.MessageBox]::Show('Please, select a Template','Attention','OK','Information')} Else {
            $varSelectedTemplate = $ListBox0.SelectedItem
            $arrayTemplateCounters = @()
            [xml]$xmlTemplate = Get-Content $DefaultPath\Templates\$varSelectedTemplate
            $arrayTemplateCountersTemp = $xmlTemplate.GetElementsByTagName("Counter")
            for ($j=0; $j -lt $arrayTemplateCountersTemp.Count; $j++)
            {
                $ListBox3.Items.AddRange($arrayTemplateCountersTemp.Item($j).InnerText)
            }
        }
    }

    $ButtonCountersTemplates1 = New-Object System.Windows.Forms.Button 
    $ButtonCountersTemplates1.Location = New-Object System.Drawing.Size(5,100) 
    $ButtonCountersTemplates1.Size = New-Object System.Drawing.Size(250,20) 
    $ButtonCountersTemplates1.Text = "Import New Template" 
    $ButtonCountersTemplates1.Add_Click({Invoke-Item $DefaultPath\Templates}) 

    $ButtonCountersTemplates2 = New-Object System.Windows.Forms.Button 
    $ButtonCountersTemplates2.Location = New-Object System.Drawing.Size(325,100) 
    $ButtonCountersTemplates2.Size = New-Object System.Drawing.Size(250,20) 
    $ButtonCountersTemplates2.Text = "Reload Template List" 
    $ButtonCountersTemplates2.Add_Click({
        $ListBox0.Items.Clear()
        get-childitem $DefaultPath\Templates\*.xml | foreach {$_.name} | Out-File $DefaultPath\Templates\Templates.txt -force
        Get-Content $DefaultPath\Templates\Templates.txt | ForEach-Object {[void] $ListBox0.Items.Add($_)}
        })
    
    $ButtonCountersTemplates3 = New-Object System.Windows.Forms.Button 
    $ButtonCountersTemplates3.Location = New-Object System.Drawing.Size(645,100) 
    $ButtonCountersTemplates3.Size = New-Object System.Drawing.Size(250,20) 
    $ButtonCountersTemplates3.Text = "Get Counters From Selected Template" 
    $ButtonCountersTemplates3.Add_Click({GetTemplateCounters}) 

#-------------------------------------------------------------------------------------------
#Third Block (CUSTOM COUNTERS / EXPERT)
#-------------------------------------------------------------------------------------------

    $Line3 = New-Object System.Windows.Forms.Label
    $Line3.ForeColor = "Gray"
    $Line3.Location = '5,338'
    $Line3.Height = 2
    $Line3.Width = 1150
    $Line3.BorderStyle = "Fixed3D"
    $Line3.AutoSize = $False

    $Font2 = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
    $LabelCustomCounters = New-Object System.Windows.Forms.Label
    $LabelCustomCounters.ForeColor = "Black"
    $LabelCustomCounters.Location = '453,350'
    $LabelCustomCounters.Font = $Font2
    $LabelCustomCounters.Text = "..:: Set Custom/Specific Counters ::.."
    $LabelCustomCounters.AutoSize = $True

    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Location = New-Object System.Drawing.Size(10,370)
    $groupBox.size = New-Object System.Drawing.Size(350,220)
    $groupBox.text = "Counter Sets:"

    $ListBox = New-Object System.Windows.Forms.ListBox 
    $ListBox.Location = New-Object System.Drawing.Size(12,20) 
    $ListBox.Size = New-Object System.Drawing.Size(325,190) 
    $ListBox.Height = 190
    get-counter -listset * | Sort-Object CounterSetName | foreach {[void] $ListBox.Items.Add($_.CounterSetName)}
    $groupBox.Controls.Add($ListBox)

    $groupBox2 = New-Object System.Windows.Forms.GroupBox
    $groupBox2.Location = New-Object System.Drawing.Size(380,370)
    $groupBox2.size = New-Object System.Drawing.Size(350,220)
    $groupBox2.text = "Counters:"

    $ListBox2 = New-Object System.Windows.Forms.ListBox 
    $ListBox2.Location = New-Object System.Drawing.Size(12,20) 
    $ListBox2.Size = New-Object System.Drawing.Size(325,170) 
    $ListBox2.Height = 170
    $ListBox2.SelectionMode = "MultiExtended"
        function LoadCounters {
            $ListBox2.Items.Clear()
            If ($ListBox.SelectedIndex.Equals(-1)) {[System.Windows.Forms.MessageBox]::Show('Please, select a Counter Set','Attention','OK','Information')} Else {
                $CounterSetSelected = $ListBox.SelectedItem.ToString()
                (get-counter -list $CounterSetSelected).counter | foreach {[void] $ListBox2.Items.Add($_)}
            }
        }
    $groupBox2.Controls.Add($ListBox2)

    $Button2 = New-Object System.Windows.Forms.Button 
    $Button2.Location = New-Object System.Drawing.Size(15,188) 
    $Button2.Size = New-Object System.Drawing.Size(320,20) 
    $Button2.Text = "Load Counters" 
    $Button2.Add_Click({LoadCounters}) 
    $groupBox2.Controls.Add($Button2)

    $Button3 = New-Object System.Windows.Forms.Button 
    $Button3.Location = New-Object System.Drawing.Size(750,440) 
    $Button3.Size = New-Object System.Drawing.Size(30,30) 
    $Button3.Text = ">>" 
    $Button3.Add_Click({AddCounters})
        function AddCounters {
            $ListBox3.Items.AddRange($ListBox2.SelectedItems)
        } 

    $Button4 = New-Object System.Windows.Forms.Button 
    $Button4.Location = New-Object System.Drawing.Size(750,470) 
    $Button4.Size = New-Object System.Drawing.Size(30,30)
    $Button4.Text = "<<" 
    $Button4.Add_Click({RemoveCounters}) 

    $groupBox3 = New-Object System.Windows.Forms.GroupBox
    $groupBox3.Location = New-Object System.Drawing.Size(800,370)
    $groupBox3.size = New-Object System.Drawing.Size(350,220)
    $groupBox3.text = "Selected Counters:"

    $ListBox3 = New-Object System.Windows.Forms.ListBox 
    $ListBox3.Location = New-Object System.Drawing.Size(12,20) 
    $ListBox3.Size = New-Object System.Drawing.Size(325,190)
    $ListBox3.SelectionMode = "MultiExtended" 
    $groupBox3.Controls.Add($ListBox3)

        function RemoveCounters {
            for ($i = $ListBox3.SelectedIndices.Count ; $i -gt 0 ; $i--) {
                $ListBox3.Items.RemoveAt($ListBox3.SelectedIndex)
            }
        }

#-------------------------------------------------------------------------------------------
#Fourth Block (BUTTONS)
#-------------------------------------------------------------------------------------------

    $Line4 = New-Object System.Windows.Forms.Label
    $Line4.ForeColor = "Gray"
    $Line4.Location = '5,350'
    $Line4.Height = 2
    $Line4.Width = 1150
    $Line4.BorderStyle = "Fixed3D"
    $Line4.AutoSize = $False
    $objForm.Controls.Add($Line4)

    $ButtonSet = New-Object System.Windows.Forms.Button 
    $ButtonSet.Location = New-Object System.Drawing.Size(70,360) 
    $ButtonSet.Size = New-Object System.Drawing.Size(150,30) 
    $ButtonSet.Text = "Set" 
    $ButtonSet.Add_Click({SetCounters}) 
    $objForm.Controls.Add($ButtonSet)

    $ButtonStop = New-Object System.Windows.Forms.Button 
    $ButtonStop.Location = New-Object System.Drawing.Size(290,360) 
    $ButtonStop.Size = New-Object System.Drawing.Size(150,30) 
    $ButtonStop.Text = "Stop" 
    $ButtonStop.Add_Click({stopcounters}) 
    $objForm.Controls.Add($ButtonStop)

    $ButtonPerfMon = New-Object System.Windows.Forms.Button 
    $ButtonPerfMon.Location = New-Object System.Drawing.Size(510,360) 
    $ButtonPerfMon.Size = New-Object System.Drawing.Size(150,30) 
    $ButtonPerfMon.Text = "Open PerfMon" 
    $ButtonPerfMon.Add_Click({Invoke-Item c:\windows\system32\perfmon.exe}) 
    $objForm.Controls.Add($ButtonPerfMon)

    $ButtonLog = New-Object System.Windows.Forms.Button 
    $ButtonLog.Location = New-Object System.Drawing.Size(730,360) 
    $ButtonLog.Size = New-Object System.Drawing.Size(150,30) 
    $ButtonLog.Text = "Open Log Folder" 
    $ButtonLog.Add_Click({
        If ($LogLocationControl.Text.Equals('')) {
            [System.Windows.Forms.MessageBox]::Show('Please, type a path to the Logs','Attention','OK','Information')
        } Else {
            If (Test-Path $LogLocationControl.Text) {
                Invoke-Item $LogLocationControl.Text
            } Else {
                [System.Windows.Forms.MessageBox]::Show('The Log Path informed does not exist!','Attention','OK','Information')
            }
        }
    }) 
    $objForm.Controls.Add($ButtonLog)

    $buttonQuit = New-Object -TypeName System.Windows.Forms.Button
    $buttonQuit.Location = New-Object System.Drawing.Size(950,360)
    $buttonQuit.Size = New-Object System.Drawing.Size(150,30)
    $buttonQuit.Text = 'Quit'
    $buttonQuit.add_Click({Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Undefined -Force ; $objform.Close()})
    $objForm.Controls.Add($buttonQuit)

#-------------------------------------------------------------------------------------------
#Footer
#-------------------------------------------------------------------------------------------

    $Line5 = New-Object System.Windows.Forms.Label
    $Line5.ForeColor = "Gray"
    $Line5.Location = '5,400'
    $Line5.Height = 2
    $Line5.Width = 1150
    $Line5.BorderStyle = "Fixed3D"
    $Line5.AutoSize = $False
    $objForm.Controls.Add($Line5)

    $Font3 = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Regular)
    $LabelFooter = New-Object System.Windows.Forms.Label
    $LabelFooter.ForeColor = "Gray"
    $LabelFooter.Location = '380,407'
    $LabelFooter.Font = $Font3
    $LabelFooter.Text = "Author: Rafael Carneiro Machado (rafaelcarneiromachado@gmail.com)"
    $LabelFooter.AutoSize = $True
    $objForm.Controls.Add($LabelFooter)

#-------------------------------------------------------------------------------------------
#Call Function SETCOUNTERS
#-------------------------------------------------------------------------------------------

Function SetCounters {

If ($NameControl.Text.Equals('Type a Name') -or $NameControl.Text.Equals('')) {[System.Windows.Forms.MessageBox]::Show('Please, type a name to the Counter Set','Attention','OK','Information')} Else {
If ($LogLocationControl.Text.Equals('')) {[System.Windows.Forms.MessageBox]::Show('Please, type a path to the Logs','Attention','OK','Information')} Else {
If ($IntervalControl.Text.Equals('')) {[System.Windows.Forms.MessageBox]::Show('Please, type a Interval in seconds on which the counters will be collected','Attention','OK','Information')} Else {

    $str2compareselectedcounters = $ListBox3.Items.Count
    If ($ListBox0.SelectedIndex.Equals(-1) -and $str2compareselectedcounters -eq 0) {[System.Windows.Forms.MessageBox]::Show('There is no Template or Counter(s) selected!','Attention','OK','Information')} Else {

        #-------------------------------------------------------------------------------------------
        #Get Variables
        #-------------------------------------------------------------------------------------------

        $varCounterSetName = $NameControl.Text
        If (Test-Path $LogLocationControl.Text) {} Else {md $LogLocationControl.Text}
        $varPathLogs = $LogLocationControl.Text + "\" + $NameControl.Text
        $varFrequencyTemp = $FrequencyControl.Text
            If ($varFrequencyTemp -eq "Stop Manually") {$varFrequency = "0"}
            If ($varFrequencyTemp -eq "60 Seconds") {$varFrequency = "00:01:00"}
            If ($varFrequencyTemp -eq "5 Minutes") {$varFrequency = "00:05:00"}
            If ($varFrequencyTemp -eq "15 Minutes") {$varFrequency = "00:15:00"}
            If ($varFrequencyTemp -eq "1 Hour") {$varFrequency = "01:00:00"}
            If ($varFrequencyTemp -eq "4 Hours") {$varFrequency = "04:00:00"}
            If ($varFrequencyTemp -eq "8 Hours") {$varFrequency = "08:00:00"}
            If ($varFrequencyTemp -eq "1 Day") {$varFrequency = "24:00:00"}
            If ($varFrequencyTemp -eq "1 Week") {$varFrequency = "168:00:00"}
            If ($varFrequencyTemp -eq "1 Month") {$varFrequency = "720:00:00"}
        $varInterval = $IntervalControl.Text
        $varStartNow = $Checkbox.CheckState

        If (test-path $DefaultPath\Counters.txt) {remove-item -path $DefaultPath\Counters.txt}

        If ($ListBox0.SelectedIndex.Equals(-1)) {} Else {
            $varSelectedTemplate = $ListBox0.SelectedItem
            $arrayTemplateCounters = @()
            [xml]$xmlTemplate = Get-Content $DefaultPath\Templates\$varSelectedTemplate
            $arrayTemplateCountersTemp = $xmlTemplate.GetElementsByTagName("Counter")
            for ($j=0; $j -lt $arrayTemplateCountersTemp.Count; $j++)
            {
                $arrayTemplateCounters += $arrayTemplateCountersTemp.Item($j).InnerText + "`n" | Out-File $DefaultPath\Counters.txt -Append
            }
        }

        If ($str2compareselectedcounters -eq 0) {} Else {
        $arraySelectedCounters = @()
            foreach ($ListBox3Item in $ListBox3.Items)
                {$arraySelectedCounters += $ListBox3Item + "`n" | Out-File $DefaultPath\Counters.txt -Append}
            Get-Content $DefaultPath\Counters.txt | Sort | Get-Unique > $DefaultPath\Counters2.txt #TO REMOVE DUPLICITIES
            Remove-Item -path $DefaultPath\Counters.txt
            Get-Content $DefaultPath\Counters2.txt | ? {$_.trim() -ne ""} | Set-Content $DefaultPath\Counters.txt #TO REMOVE BLANK SPACES/LINES
            Remove-Item -path $DefaultPath\Counters2.txt
        }

        #-------------------------------------------------------------------------------------------
        #Set Counters
        #-------------------------------------------------------------------------------------------
   
        $parameters2set = ""
            If ($varFrequency -ne "0") {                #SET COUNTERS TO STOP AFTER A SPECIFIC TIME
                $parameters2set = 'create counter -n "' + $varCounterSetName + '" -o "' + $varPathLogs + '" -si '  + $varInterval + ' -rf ' + $varFrequency + ' -cf ' + $DefaultPath + '\Counters.txt'
                Start-Process c:\windows\system32\logman.exe -WindowStyle hidden $parameters2set -verb runas -Wait 
                #Write-Host "1: "Start-Process c:\windows\system32\logman.exe $parameters2set -Wait
                remove-item -path $DefaultPath\Counters.txt
                [System.Windows.Forms.MessageBox]::Show("All Good! :-)`nCheck PerfMon",'Attention','OK','Information')
            } Else {                                    #SET COUNTERS TO DO NOT STOP AUTOMATICALLY
                $parameters2set = 'create counter -n "' + $varCounterSetName + '" -o "' + $varPathLogs + '" -si '  + $varInterval + ' -cf ' + $DefaultPath + '\Counters.txt'
                Start-Process c:\windows\system32\logman.exe -WindowStyle hidden $parameters2set -verb runas -Wait
                #Write-Host "2: "Start-Process c:\windows\system32\logman.exe $parameters2set -Wait
                remove-item -path $DefaultPath\Counters.txt
                [System.Windows.Forms.MessageBox]::Show("All Good! :-)`nCheck PerfMon",'Attention','OK','Information')
            }
            If ($varStartNow -eq 'Checked') {               #SET COUNTERS AND START IT STRAIGH AWAY
                $CurrentLogTemp = (logman query "type a name" | findstr /b /o "Output Location")
                If ([string]::IsNullOrEmpty($CurrentLogTemp)) {} Else {
                    $CurrentLog = ($CurrentLogTemp).substring(26)
                    If (Test-Path "$CurrentLog") {Remove-Item "$CurrentLog"}
                }
                $parameters2start = 'start "' + $NameControl.Text + '"'
                Start-Process c:\windows\system32\logman.exe -WindowStyle hidden $parameters2start -verb runas -Wait
            }
    }
}
}
}
}

#-------------------------------------------------------------------------------------------
#Call Function STOPCOUNTERS
#-------------------------------------------------------------------------------------------
Function StopCounters {

If ($NameControl.Text.Equals('Type a Name') -or $NameControl.Text.Equals('')) {[System.Windows.Forms.MessageBox]::Show('Please, type the name of the Counter Set to be stopped','Attention','OK','Information')} Else {
    $str2checkCounterSet2StopTemp = (logman query $NameControl.Text | findstr /b /o "Status:")
    $str2checkCounterSet2Stop = ($str2checkCounterSet2StopTemp).substring(25)
    If ($str2checkCounterSet2Stop -eq '' -or $str2checkCounterSet2Stop -eq 'Stopped') {[System.Windows.Forms.MessageBox]::Show('Counter Set not found or it is stopped already','Attention','OK','Information')} Else {


        #-------------------------------------------------------------------------------------------
        #Get Variables
        #-------------------------------------------------------------------------------------------

        $varCounterSetName = $NameControl.Text

        #-------------------------------------------------------------------------------------------
        #Stop Counters
        #-------------------------------------------------------------------------------------------

        $parameters2stop = 'stop "' + $NameControl.Text + '"'
        Start-Process c:\windows\system32\logman.exe -WindowStyle hidden $parameters2stop -Wait
        [System.Windows.Forms.MessageBox]::Show('Counter Set stopped, check PerfMon','Attention','OK','Information')
    }
}
}

#-------------------------------------------------------------------------------------------
#Activate the Form
#-------------------------------------------------------------------------------------------

$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()

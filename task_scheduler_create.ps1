###########################################
# Script Name: taskscheduler_create.ps1
# Version: 1.0
###########################################


$script_path="<>"

######################### LOG Functions ############################
$runuser=whoami
Function Start-Log {
    
    #Check if file exists and delete if it does
    If((Test-Path -Path $LogFile)){
        Remove-Item -Path $LogFile -Force
    }
 
    #Create file and start logging
    New-Item -Path $LogPath -Name $LogName –ItemType File |Out-Null
 
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
    Add-Content -Path $LogFile -Value "Started Script $($MyInvocation.MyCommand.Name) at [$([DateTime]::Now)]."
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
}

Function Log {
param(
    [string]$In
)
    Add-Content -Path $LogFile -Value $In
	#write-host "$In"
}

Function End-Log {
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
    Add-Content -Path $LogFile -Value "Finished processing at [$([DateTime]::Now)]."
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
}



$fileName="TaskSchedule_create"

$LogName = "$($fileName)_$(get-date -Format "yyyyMMdd-HHmm").log"

if(!(Test-Path "$script_path\Logs"))
{
   New-Item -Path "$script_path" -Name "Logs" -ItemType Directory -Force 
}
$LogPath = "$script_path\Logs"
$LogFile = $LogPath + "\" + $LogName

Start-Log

$data1 =  Import-Csv -Path "$script_path\Frequency.csv"
$freq_type = $data1.Frequency | select -Unique
[array]$dropDownArray1 = $freq_type

$data2 =  Import-Csv -Path "$script_path\DaysOfWeek.csv"
$dow_type = $data2.Days | select -Unique
[array]$dropDownArray2 = $dow_type

$data3 =  Import-Csv -Path "$script_path\Month.csv"
$month_type = $data3.Month | select -Unique
[array]$dropDownArray3 = $month_type

$data4 =  Import-Csv -Path "$script_path\Days.csv"
$day_type = $data4.Days | select -Unique
[array]$dropDownArray4 = $day_type

$data5 =  Import-Csv -Path "$script_path\Time.csv"
$time_type = $data5.Time | select -Unique
[array]$dropDownArray5 = $time_type

[void][System.Reflection.Assembly]::LoadWithPartialName( “System.Windows.Forms”)
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName( “Microsoft.VisualBasic”)

$form = New-Object System.Windows.Forms.Form
$form.Width = 1200;
$form.Height = 600;
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.Text = "Task Scheduler Creation Form";
$form.BackColor = "Lightgray"
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

$serverlabel = New-Object “System.Windows.Forms.Label”;
$serverlabel.Left = 30;
$serverlabel.Top = 20;
$serverlabel.AutoSize  = $True
$serverlabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$serverlabel.Font = $Font
$serverlabel.Text = "Enter Target Server";
$form.Controls.Add($serverlabel);

$namelabel = New-Object “System.Windows.Forms.Label”;
$namelabel.Left = 30;
$namelabel.Top = 60;
$namelabel.AutoSize  = $True
$namelabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$namelabel.Font = $Font
$namelabel.Text = "Enter Task Scheduler Name";
$form.Controls.Add($namelabel);

$scriptlabel = New-Object “System.Windows.Forms.Label”;
$scriptlabel.Left = 30;
$scriptlabel.Top = 100;
$scriptlabel.AutoSize  = $True
$scriptlabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$scriptlabel.Font = $Font
$scriptlabel.Text = "Enter Full Path of script";
$form.Controls.Add($scriptlabel);

$userlabel = New-Object “System.Windows.Forms.Label”;
$userlabel.Left = 30;
$userlabel.Top = 140;
$userlabel.AutoSize  = $True
$userlabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$userlabel.Font = $Font
$userlabel.Text = "Enter UserName";
$form.Controls.Add($userlabel);

$passwdlabel = New-Object “System.Windows.Forms.Label”;
$passwdlabel.Left = 30;
$passwdlabel.Top = 180;
$passwdlabel.AutoSize  = $True
$passwdlabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$passwdlabel.Font = $Font
$passwdlabel.Text = "Enter Password";
$form.Controls.Add($passwdlabel);

$freqlabel = New-Object “System.Windows.Forms.Label”;
$freqlabel.Left = 30;
$freqlabel.Top = 220;
$freqlabel.AutoSize  = $True
$freqlabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$freqlabel.Font = $Font
$freqlabel.Text = "Select Trigger Frequency";
$form.Controls.Add($freqlabel);

$timelabel = New-Object “System.Windows.Forms.Label”;
$timelabel.Left = 30;
$timelabel.Top = 260;
$timelabel.AutoSize  = $True
$timelabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$timelabel.Font = $Font
$timelabel.Text = "Enter Time";
$form.Controls.Add($timelabel);


############################################################################
$textboxsserver = New-Object System.Windows.Forms.TextBox
$textboxsserver.Left = 500;
$textboxsserver.Top = 20;
$textboxsserver.width = 300;
$textboxsserver.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxsserver.Font = $Font
$form.Controls.Add($textboxsserver)

$textboxsname = New-Object System.Windows.Forms.TextBox
$textboxsname.Left = 500;
$textboxsname.Top = 60;
$textboxsname.width = 300;
$textboxsname.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxsname.Font = $Font
$form.Controls.Add($textboxsname)

$textboxscript = New-Object System.Windows.Forms.TextBox
$textboxscript.Left = 500;
$textboxscript.Top = 100;
$textboxscript.width = 300;
$textboxscript.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxscript.Font = $Font
$form.Controls.Add($textboxscript)


$textboxuser = New-Object System.Windows.Forms.TextBox
$textboxuser.Left = 500;
$textboxuser.Top = 140;
$textboxuser.width = 300;
$textboxuser.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxuser.Font = $Font
$form.Controls.Add($textboxuser)


$textboxpasswd = New-Object System.Windows.Forms.TextBox
$textboxpasswd.PasswordChar = '*'
$textboxpasswd.Left = 500;
$textboxpasswd.Top = 180;
$textboxpasswd.width = 300;
$textboxpasswd.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxpasswd.Font = $Font
$form.Controls.Add($textboxpasswd)

$textBoxfreq = New-Object “System.Windows.Forms.ComboBox”;
$textBoxfreq.Left = 500;
$textBoxfreq.height = 300
$textBoxfreq.AutoSize  = $True
$textBoxfreq.Top = 220;
$textBoxfreq.width = 100;
$textBoxfreq.DropDownStyle =[System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxfreq.Font = $Font

$textBoxdow = New-Object “System.Windows.Forms.ComboBox”;
$textBoxdow.Left = 630;
$textBoxdow.height = 300
$textBoxdow.AutoSize  = $True
$textBoxdow.Top = 220;
$textBoxdow.width = 100;
$textBoxdow.DropDownStyle =[System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxdow.Font = $Font

$textBoxmonth = New-Object “System.Windows.Forms.ComboBox”;
$textBoxmonth.Left = 740;
$textBoxmonth.height = 300
$textBoxmonth.AutoSize  = $True
$textBoxmonth.Top = 220;
$textBoxmonth.width = 100;
$textBoxmonth.DropDownStyle =[System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxmonth.Font = $Font

$textBoxday = New-Object “System.Windows.Forms.ComboBox”;
$textBoxday.Left = 860;
$textBoxday.height = 300
$textBoxday.AutoSize  = $True
$textBoxday.Top = 220;
$textBoxday.width = 100;
$textBoxday.DropDownStyle =[System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxday.Font = $Font


$textboxtime = New-Object System.Windows.Forms.TextBox
$textboxtime.Left = 500;
$textboxtime.Top = 260;
$textboxtime.width = 100;
$textboxtime.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textboxtime.Font = $Font
$form.Controls.Add($textboxtime)

$textBoxtimezone = New-Object “System.Windows.Forms.ComboBox”;
$textBoxtimezone.Left = 610;
$textBoxtimezone.height = 300
$textBoxtimezone.AutoSize  = $True
$textBoxtimezone.Top = 260;
$textBoxtimezone.width = 100;
$textBoxtimezone.DropDownStyle =[System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxtimezone.Font = $Font



ForEach ($item in $dropDownArray1) {
     [void] $textBoxfreq.Items.Add($item)
    } 

$form.Controls.Add($textBoxfreq);

ForEach ($item in $dropDownArray2) {
     [void] $textBoxdow.Items.Add($item)
    } 

$form.Controls.Add($textBoxdow);

ForEach ($item in $dropDownArray3) {
     [void] $textBoxmonth.Items.Add($item)
    } 

$form.Controls.Add($textBoxmonth);

ForEach ($item in $dropDownArray4) {
     [void] $textBoxday.Items.Add($item)
    } 

$form.Controls.Add($textBoxday);

ForEach ($item in $dropDownArray5) {
     [void] $textBoxtimezone.Items.Add($item)
    } 

$form.Controls.Add($textBoxtimezone);


#Add a OK button
$okButton1 = New-Object System.Windows.Forms.Button
$okButton1.Left = 400;
$okButton1.Top = 500;
$okButton1.Width = 170;
$okButton1.Height = 30;
$okButton1.Text = 'OK'
$Font = New-Object System.Drawing.Font("Calibri",13,[System.Drawing.FontStyle]::Bold)
$okButton1.Font = $Font
$okButton1.DialogResult=[System.Windows.Forms.DialogResult]::OK
$okButton1.ForeColor = "Black"
$okButton1.BackColor = "PaleGoldenrod"
$form.Controls.Add($okButton1)


#Add a cancel button
$cancelButton1 = New-Object System.Windows.Forms.Button
$cancelButton1.Left = 700;
$cancelButton1.Top = 500;
$cancelButton1.Width = 170;
$cancelButton1.Height = 30;
$cancelButton1.Text = "Cancel"
$Font = New-Object System.Drawing.Font("Calibri",13,[System.Drawing.FontStyle]::Bold)
$cancelButton1.Font = $Font
$cancelButton1.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
$cancelButton1.ForeColor = "Black"
$cancelButton1.BackColor = "PaleGoldenrod"
$form.Controls.Add($cancelButton1)
$cancelButton1.add_Click({$form.Close()})

$box = $form.ShowDialog()


$server=$textboxsserver.Text
$name=$textboxsname.Text
$scriptdetails=$textboxscript.Text
$username=$textboxuser.Text
$password=$textboxpasswd.Text
$frequency=$textBoxfreq.Text
$dayofweek=$textBoxdow.Text
$month=$textBoxmonth.Text
$dayofmonth=$textBoxday.Text
$time=$textboxtime.Text
$zone=$textBoxtimezone.Text

Log "Run User: $runuser"
Log "TaskScheduler Name: $name"
Log "Script Details: $scriptdetails"
Log "Frequency: $frequency"
Log "Time of Task Scheduler Run: $time"
Log " "



################################################################

if((Test-Connection -ComputerName $server))
{

$regex1='(\.ps1)'
$regex2='(\.bat)'


if($scriptdetails -match $regex1)
{
$scriptdetails="powershell.exe $scriptdetails"
}
elseif($scriptdetails -match $regex2)
{
$scriptdetails=$scriptdetails
}
else
{
Write-Host "unknown script file extension."
}



if ("$frequency" -eq "Daily")
{
$TriggerFrequency = @{ $frequency = $true }
$Action = New-ScheduledTaskAction -Execute $scriptdetails
$Trigger = New-ScheduledTaskTrigger @TriggerFrequency -At $time$zone
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
#Register-ScheduledTask -TaskName $name -InputObject $Task -User $usrename -Password $password -Force
Invoke-Command -ComputerName $server -ScriptBlock {Register-ScheduledTask -TaskName $using:name -InputObject $using:Task -User $using:username -Password $using:password -Force}
Log "Task Scheduler has been created."
}

if ("$frequency" -eq "Weekly" -and $dayofweek -ne " ")
{
$TriggerFrequency = @{ $frequency = $true }
$Action = New-ScheduledTaskAction -Execute $scriptdetails
$Trigger = New-ScheduledTaskTrigger @TriggerFrequency -At $time$zone -DaysOfWeek $dayofweek
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
#Register-ScheduledTask -TaskName $name -InputObject $Task -User $usrename -Password $password -Force
Invoke-Command -ComputerName $server -ScriptBlock {Register-ScheduledTask -TaskName $using:name -InputObject $using:Task -User $using:username -Password $using:password -Force}
Log "Task Scheduler has been created."
}



if ($frequency -eq "Monthly" -and $month -ne "" -and $dayofmonth -ne "")
{
$suffix=":00"
if ($zone -eq "AM")
{
$time=$time
}
elseif($zone -eq "PM")
{
$time=12+$time
}
$month=$month.Substring(0,3).ToUpper()
Invoke-Command -ComputerName $server -ScriptBlock {schtasks.exe /Create /SC $using:frequency /M $using:month /D $using:dayofmonth /TN $using:name /ST $using:time+$using:suffix /TR $using:scriptdetails /F /RU $using:username /RP $using:password }
}

End-Log
}

else
{
Write-Host "Target server $server is not reachable. Exiting..." -ForegroundColor Red -BackgroundColor white
Log "Run user: $runuser"
Log "Target server $server is not reachable. Exiting..."
End-Log
sleep 2
exit
}

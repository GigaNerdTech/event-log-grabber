# Get Windows Event logs for a time range

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Event Log Grabber" Height="500" Width="450" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="ServerToCheck" Content="Server to check:" />
        <TextBox x:Name="ServerTextBox" />
        <Label x:Name="StartDateLabel" Content="Start Time (MM/DD/YYYY HH:MM:SS):" />
        <TextBox x:Name="StartDate" />
        <Label x:Name="Interval" Content="Interval to check (minutes):" />
        <TextBox x:Name="Minutes" />
        <Label x:Name="CSV" Content="CSV to save to:" />
        <TextBox x:Name="Path" />
        <CheckBox x:Name="OSOnly" Content="Operating System Logs Only" />
        <Button x:Name="EventLogButton" Content="Grab Event Logs" Margin="10,10,10,0" VerticalAlignment="Top" Height="25" />
        <Label x:Name="Status" Content="Status: READY" />
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Controls find
$serverbox = $global:Form.FindName('ServerTextBox')
$startdatebox = $global:Form.FindName('StartDate')
$intervalbox = $global:Form.FindName('Minutes')
$pathbox = $global:Form.FindName('Path')
$grabberbutton = $global:Form.FindName('EventLogButton')
$os_only_check = $global:Form.FindName('OSOnly')
$status = $global:Form.FindName('Status')

$grabberbutton.Add_Click({

    # Default log path
    $capture_path = "C:\temp\"

    # Get Date from text box

    $desired_date = $startdatebox.Text.ToString()

    # Parse date
    if ($desired_date -match "(?<month>0[1-9]|1[012])[- /.](?<day>\d\d)[- /.](?<year>[0-9]{4}) (?<hour>\d\d):(?<minute>\d\d):(?<second>\d\d)") {
        Write-Host "Date matched!"
        $month=$Matches['month']
        $day=$Matches['day']
        $year=$Matches['year']
        $hour=$Matches['hour']
        $minute=$Matches['minute']
        $second=$Matches['second']
        $date_to_test = (Get-Date -Year $year -Day $day -Month $month -Hour $hour -Minute $minute -Second $second)
    }
    else {
        $date_to_test = ([datetime]::Now).AddMinutes(-10)
    }
    
   
    # Get check interval
    if (-Not ([string]::IsNullOrEmpty($intervalbox.Text))) {
        $check_interval = $intervalbox.Text
    }
    else {
        $check_interval =  10
    }

    # Get server name
    if (-Not ([string]::IsNullOrEmpty($serverbox.Text))) {
        $server_to_check = $serverbox.Text
    }
    else {
        $server_to_check = "localhost"
    }


    # Get File path
    if (-Not ([string]::IsNullOrEmpty($pathbox.Text))) {
        $filepath = $pathbox.Text
    }
    else {
        $log_time = [datetime]::Now
        $log_stamp = $log_time.ToString('yyyyMMdd-hhmmss')
        $filepath = "$capture_path\$server_to_check-event-$log_stamp.log"
    }

    # Debug

    Write-Host "Hostname: $server_to_check"
    Write-Host "Start date: $month/$day/$year $hour : $minute : $second"
    Write-Host "Interval: $check_interval"
    Write-Host "Capture path: $filepath"

    # Capture event logs, can take a looooong time

    # Change status
    $status.Content = "Status: RUNNING"


    if ($os_only_check.IsChecked) {
        Get-WinEvent -ListLog "Application","Security","Setup","System" -ErrorVariable err -ea 0 -ComputerName $server_to_check | Where-Object {$_.RecordCount -And $_.LastWriteTime -gt $date_to_test } | ForEach-Object { Write-Host ("Log processing: " + $_.LogName); Get-WinEvent -ComputerName $server_to_check -LogName $_.LogName -ErrorAction SilentlyContinue | 
        Where-Object { ($_.TimeCreated -le ($date_to_test.AddMinutes($check_interval))) -And ($_.TimeCreated -ge $date_to_test) } }| Select -Property * | Export-Csv -Path $filepath
    }
    else {
        Get-WinEvent -ListLog * -ErrorVariable err -ea 0 -ComputerName $server_to_check | Where-Object {$_.RecordCount -And $_.LastWriteTime -gt $date_to_test } | ForEach-Object { Write-Host ("Log processing: " + $_.LogName); Get-WinEvent -ComputerName $server_to_check -LogName $_.LogName -ErrorAction SilentlyContinue | 
        Where-Object { ($_.TimeCreated -le ($date_to_test.AddMinutes($check_interval))) -And ($_.TimeCreated -ge $date_to_test) } }| Select -Property * | Export-Csv -Path $filepath
    }

    $status.Content = "Status: READY"

    Write-Host "Done."
})

# Show GUI
$global:Form.ShowDialog() | out-null
# ($_.TimeCreated -le ($date_to_test + $check_interval)) -And
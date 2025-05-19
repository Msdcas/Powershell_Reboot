Add-Type -AssemblyName System.Windows.Forms

[int] $TresholdHours = 24*7
[string] $TaskNamePCReboot = "ForcedReboot_viaScript"

[long] $FormCloseSecondTimeout = 60*30
[bool] $IsFormTimeoutVisible = $false

[bool] $IsWriteLogs = $true
[string] $LogFolderName = "ScriptAutoReboot_logs"
[string] $LogDirectory = Join-Path -Path $env:APPDATA -ChildPath $LogFolderName
[int32] $LogFileMaxSize = 2MB

#internal variables, don't change their values without knowledge about that
[string] $StrTimeoutCancel = "TimeoutCancel"

function GetUserSelectedTimeSpan {
    param
    (
        [string[]] $times,
        [string] $infoWorkedTimeStr
    )
    
    [bool] $isOkButtonClicked = $false
    [bool] $isTimeoutEnds = $false

    #region GUI
    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Выбор нерабочего времени для перезапуска ПК"
    $form.Size            = New-Object System.Drawing.Size(550, 200)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false

    $Label1               = New-Object system.Windows.Forms.Label
    $Label1.text          = "Ваш ПК следует перезагрузить т.к. он работает больше " + $infoWorkedTimeStr + "`n"
    $Label1.text          += "Вам следует закрыть все программы и файлы к выбранному времени"
    $Label1.AutoSize      = $true
    $Label1.Size          = New-Object System.Drawing.Size(400, 30)
    $Label1.location      = New-Object System.Drawing.Point(12,9)
    $Label1.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    $form.Controls.Add($Label1)

    $comboBox             = New-Object System.Windows.Forms.ComboBox
    $comboBox.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    $comboBox.Location    = New-Object System.Drawing.Point(78, 58)
    $comboBox.Size        = New-Object System.Drawing.Size(286, 23)
    $combobox.DropDownStyle = 'DropDownList'
    $comboBox.Items.AddRange($times)
    $form.Controls.Add($comboBox)
    $comboBox.SelectedItem = 0

    $okButton             = New-Object System.Windows.Forms.Button
    $okButton.Location    = New-Object System.Drawing.Point(400, 103)
    $okButton.Size        = New-Object System.Drawing.Size(112, 41)
    $okButton.Text        = "ОК"
    $form.Controls.Add($okButton)

    $Label3               = New-Object system.Windows.Forms.Label
    $Label3.text          = "Executg time"
    $Label3.Name          = "lTime"
    $Label3.AutoSize      = $true
    $Label3.Visible       = $IsFormTimeoutVisible
    $Label3.Enabled       = $false
    $Label3.Size          = New-Object System.Drawing.Size(10, 100)
    $Label3.location      = New-Object System.Drawing.Point(12,133)
    $Label3.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',11)
    $form.Controls.Add($Label3)
    #endregion GUI

    $okButton.Add_Click({
        if ($comboBox.SelectedItem -ne $null) {
            $ReturnedStr = "$($times[$comboBox.SelectedIndex])"

            $isOkButtonClicked = $true
            [void] $form.Close()
        } 
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите время.")
        }
    })

    $form.Add_FormClosing({
        if ($_.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {

            if (-not $isTimeoutEnds) {
                if (-not $isOkButtonClicked) {
                    $_.Cancel = $true
                    [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите время и нажмите ОК.")
                }
            }            
        }

    })

    function FormCloseDelayed {
        param
        (
            [System.Windows.Forms.Form] $f,
            [long] $timeoutSeconds,
            [ref] $isTimeoutEnds
        )

        $label = $f.Controls.Find('lTime', $false)[0]

        for ($i = $timeoutSeconds; $i -gt 0; $i--) {
            $time = [TimeSpan]::FromSeconds($i)
            $label.Text = "Until auto closed $($time.Minutes.ToString()) : $($time.Seconds.ToString()) ..."
            Start-Sleep -Seconds 1
        }

        $isTimeoutEnds.Value = $true
        $f.Close()
    }

    try{
        Invoke-CommandAsync $function:FormCloseDelayed -ArgumentList @($form, $FormCloseSecondTimeout, [ref] $isTimeoutEnds)
        $form.ShowDialog() | Out-Null
    }
    catch{
        Write-Log -message $_.Exception.ToString()
        return $StrTimeoutCancel
    }

    if ($isTimeoutEnds){
        return $StrTimeoutCancel
    }
    else {
        return "$($times[$comboBox.SelectedIndex])"
    }
}

function Invoke-CommandAsync {
    param(
        [Parameter(Mandatory)]
        [scriptblock] $Action,

        [Parameter()]
        [object[]] $ArgumentList)
    
    try{
        $ps = [powershell]::Create().AddScript($Action)
        foreach ($arg in $ArgumentList) {
            $ps = $ps.AddArgument($arg)
        }

        $registerObjectEventSplat = @{
            InputObject = $ps
            EventName   = 'InvocationStateChanged'
        }

        # Auto-dispose mechanism
        $null = Register-ObjectEvent @registerObjectEventSplat -Action {
            param(
                [object] $s,
                [System.Management.Automation.PSInvocationStateChangedEventArgs] $e)

            if ($e.InvocationStateInfo.State -ne 'Running') {
                $s.Dispose()
                Unregister-Event -SourceIdentifier $Event.SourceIdentifier
            }
        }

        $null = $ps.BeginInvoke()
    }
    catch{
        Write-Log -message $_.Exception.ToString()
    }
}

function CreateScheduleRebootTask {
    param
    (
        [string] $taskName, 
        [string] $timeExecuteStr
    )
    try{
        $action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /f /t 0"
        $trigger = New-ScheduledTaskTrigger -At $timeExecuteStr -Once
        #$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount    # need adm access
        $currentUser = $env:USERNAME
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive
        Register-ScheduledTask  -Action $action -Trigger $trigger -Principal $principal -TaskName $taskName -Description "Создано автоматически. Принудительный перезапуск ПК"
    }
    catch{
        Write-Log -message $_.Exception.ToString()
    }
}

function IsScheduleExist ([string] $taskName) {
    try
    {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
        if ($task -eq $null) {
            return $False
        }
    }
    catch
    {
        return $False
    }
    return $True
}

function RemoveSchedule ([string] $taskName) {
    try{
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Log -message "Reboot task successfully removed"
    }
    catch{
        Write-Log -message $_.Exception.ToString()
    }
}

function Get-TimeDescriptionArray {
    param (
        [int]$startHour,
        [int]$endHour,
        [string] $dayDescription
    )

    if ($startHour -lt 0 -or $startHour -gt 23) {
        Write-Log -message "The hour must be between 0 and 23"
        return $null
    }

    if ($endHour -lt 0 -or $endHour -gt 23) {
        Write-Log -message "The hour must be between 0 and 23"
        return $null
    }

    if ($startHour -ge $endHour) {
        Write-Log -message "The starting hour must be less than the ending hour"
        return $null
    }

    $Array = @()

    for ($hour = $startHour; $hour -lt $endHour; $hour++) {
        $Array += "$dayDescription $($hour):00"
    }

    return $Array
}

function Write-Log {
    param (
        [string]$message
    )

    if (-not $IsWriteLogs)
    {
        return
    }

    $logFileName = "log.txt"
    

    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory | Out-Null
    }

    $logFilePath = Join-Path -Path $LogDirectory -ChildPath $logFileName

    if (Test-Path -Path $logFilePath) {
        $fileInfo = Get-Item -Path $logFilePath
    }

    if ($fileInfo.Length -gt $LogFileMaxSize) {
        Remove-Item -Path $logFilePath -Force
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    $logMessage = "$timestamp | $message"
    Add-Content -Path $logFilePath -Value $logMessage
}


try{
    # Получаем последние событие с идентификатором 6005 (система была запущена)
    $timeLastBootEvent = Get-WinEvent -FilterHashtable @{LogName='System'; Id=6005} -MaxEvents 1 | 
        Select-Object TimeCreated, Id, Message #, @{Expression={$_.TimeCreated.ToString('yyyy-MM-dd-HH-mm-ss')}}

    # Получаем последние событие с идентификатором 6006 (система была выключена)
    $timeLastShutdownEvent = Get-WinEvent -FilterHashtable @{LogName='System'; Id=6006} -MaxEvents 1 | 
        Select-Object TimeCreated, Id, Message #, @{Expression={$_.TimeCreated.ToString('yyyy-MM-dd-HH-mm-ss')}}
}
catch{
        Write-Log -message $_.Exception.ToString()
}

if ($timeLastShutdownEvent.TimeCreated -gt $timeLastBootEvent.TimeCreated )
{
    #Write-Host "Data error. Last shutdown event is younger than last startup event."
    return
}

$workTime = ( $(Get-Date) - $timeLastBootEvent.TimeCreated)
if ($workTime.TotalHours -le $TresholdHours)
{
    #Write-Host "The PC is in operation less than the threshold value specified"

    # del old reboot task
    if (IsScheduleExist -taskName $TaskNamePCReboot)
    {
        RemoveSchedule -taskName $TaskNamePCReboot
    }
    return
}

if (IsScheduleExist -taskName $TaskNamePCReboot){
    #Write-Host "Reboot task already exist"
    return
}


# при добавлении текстовых полей, нужно добавить логику в switch ($userSelTime)
$TimesToUserSelect = @("Через 5 минут", "Через 15 минут")
$todayPeriod = TimeDescriptionArray -startHour (([DateTime]::Now).Hour+1) -endHour 23 -dayDescription "Сегодня"
$tommorowPeriod = TimeDescriptionArray -startHour 5 -endHour 23 -dayDescription "Завтра"
$afterTomomorow = TimeDescriptionArray -startHour 5 -endHour 23 -dayDescription "Послезавтра"

$TimesToUserSelect = $TimesToUserSelect + $todayPeriod + $tommorowPeriod + $afterTomomorow

$currentWorkedTimeStr = "$($workTime.Days) дней и $($workTime.Hours) часов"
$userSelTime = GetUserSelectedTimeSpan -times $TimesToUserSelect -infoWorkedTimeStr $currentWorkedTimeStr


$rebootDateTime = [DateTime]::Now
switch ($userSelTime){
    
    { $_.StartsWith("Через 5 минут") } 
        {
        $timeSpan = New-TimeSpan -Minutes 5
        $rebootDateTime += $timeSpan
        }

    { $_.StartsWith("Через 15 минут") } 
        {
        $timeSpan = New-TimeSpan -Minutes 15
        $rebootDateTime += $timeSpan
        }

    { $_.StartsWith("Сегодня") }
        { 
            $rebootDateTime = [DateTime]::Now.Date
            $timeSpan = [TimeSpan]::Zero
            if (-Not [TimeSpan]::TryParse($userSelTime.Replace("Сегодня ",""), [ref]$timeSpan)){
                Write-Log -message "Error converting time period for selected period in range 'Today'"
                return
            }

            $rebootDateTime += $timeSpan
        }

    { $_.StartsWith("Завтра") }
        { 
            $rebootDateTime = [DateTime]::Now.Date.Add( $(New-TimeSpan -Days 1) )
            $timeSpan = [TimeSpan]::Zero
            if (-Not [TimeSpan]::TryParse($userSelTime.Replace("Завтра ",""), [ref]$timeSpan)){
                Write-Log -message "Error converting time period for selected period in range 'Tomorrow'"
                return
            }

            $rebootDateTime += $timeSpan
        }

    { $_.StartsWith("Послезавтра") }
        { 
            $rebootDateTime = [DateTime]::Now.Date.Add( $(New-TimeSpan -Days 2) )
            $timeSpan = [TimeSpan]::Zero
            if (-Not [TimeSpan]::TryParse($userSelTime.Replace("Послезавтра ",""), [ref]$timeSpan)){
                Write-Log -message "Error converting time period for selected period in range 'Day after tomorrow'"
                return
            }
            
            $rebootDateTime += $timeSpan
        }

    { $_.StartsWith($StrTimeoutCancel) }
        {
        Write-Log "The script was terminated automatically because the user did not select a time on the form within the specified time period = $FormCloseSecondTimeout [sec]"
        return
        }
    default 
    {
        Write-Log -message "Exception in switch(userSetlTime). Input value = $userSelTime"
        return
    }
}
$rebootTimeStr = $rebootDateTime.ToString("yyyy-MM-ddTHH:mm:ss")


CreateScheduleRebootTask -timeExecuteStr $rebootTimeStr -taskName $TaskNamePCReboot

if (IsScheduleExist -taskName $TaskNamePCReboot) {
    [System.Windows.Forms.MessageBox]::Show("Задача успешно создана!")
    Write-Log -message "Created reboot task on date = $rebootTimeStr"
    }
else {
    [System.Windows.Forms.MessageBox]::Show("Ошибка создания задачи. Пожалуйста, обратитесь к администратору для выяснения причины")
    Write-Log -message "Error creating reboot task on date = $rebootTimeStr"
}
    

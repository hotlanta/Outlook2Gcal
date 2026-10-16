Use Windows Task Scheduler to run the script automatically

1. Create a batch file, for example, run_google_calendar.bat with the following content:
@echo off
cd /d C:\Users\<username>\<location of python script>
C:\Users\<username>\<location of python executable>\python.exe outlook2gcal_api.py

2. In PowerShell, create the task, for example:
$action = New-ScheduledTaskAction `
    -Execute "C:\Users\<username>\<location of batch file>\run_google_calendar.bat"

$trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
    -At 6:00AM

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "Run Google Calendar" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -User $env:USERNAME

 3. After creating the task:
      Open Task Scheduler.
      Open Run Google Calendar.
      Go to Conditions.
      Check: Wake the computer to run this task

4. Under Settings, check Run task as soon as possible after a scheduled start is missed.

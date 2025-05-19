# Powershell_Reboot
The script is planned to be published by politicians and run approximately once an hour. To do this, you will need to develop another script that will run the one described here on schedule.
The script compares the current PC operating time with the threshold value (7 days by default). And if it exceeds the threshold value, a form is created in which the user is asked to select a time to schedule a restart. At the selected time, a task to restart the PC is created on behalf of the user in the task scheduler.
The form cannot be closed without selecting a time. The form will close automatically after half an hour - this is necessary to avoid a conflict when creating a task in the task scheduler.
The time intervals available for selection are presented as an hourly selection for the current day, the next day, and the following day.
The script will not be run again if the restart task already exists. The script does not create more than 2 tasks and always tries to delete the old one if it finds one.
The cycle methods have logging. The log file is stored by default in the "ScriptAutoReboot_logs" subfolder of the AppData folder of the user under whose session the script is launched.

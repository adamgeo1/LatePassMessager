import win32com.client
import os
import sys
import datetime
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-python-path", default=sys.executable, help="Path to Python executable")
args = parser.parse_args()

python_exe = args.python_path
script_path = os.path.abspath("Main.py")

task_name = "SendLatePassEmails"
start_time = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(hour=0, minute=1, second=0, microsecond=0)
#start_time = datetime.datetime.now() + datetime.timedelta(minutes=1)

scheduler = win32com.client.Dispatch("Schedule.Service")
scheduler.Connect()

root_folder = scheduler.GetFolder("\\")
task_def = scheduler.NewTask(0)

task_def.RegistrationInfo.Description = "Send late pass emails for CS 270 automatically every Saturday at 12:01 AM"
task_def.Settings.Enabled = True
task_def.Settings.StopIfGoingOnBatteries = False
task_def.Settings.DisallowStartIfOnBatteries = False
task_def.Settings.RunOnlyIfNetworkAvailable = True

trigger = task_def.Triggers.Create(3)
trigger.StartBoundary = start_time.strftime("%Y-%m-%dT%H:%M:%S")
trigger.DaysOfWeek = 64
trigger.WeeksInterval = 1
trigger.Enabled = True

action = task_def.Actions.Create(0)
action.Path = python_exe
action.Arguments = f'"{script_path}"'
action.WorkingDirectory = os.path.dirname(script_path)

task_def.Principal.RunLevel = 1

root_folder.RegisterTaskDefinition(task_name, task_def, 6, "", "", 0)

print(f"✅ Scheduled task '{task_name}' created using Python at {python_exe}")

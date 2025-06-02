import os
import platform
import subprocess
import sys
import datetime
import argparse

def setup_windows(python_exe, script_path):
    import win32com.client

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

def setup_unix(python_exe, script_path):
    cron_entry = f"1 0 * * 6 {python_exe} \"{script_path}\""

    try:
        result = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
        exisiting_cron = result.stdout if result.returncode == 0 else ""
        if cron_entry in exisiting_cron:
            print("Cron job already exists")
            return

        new_cron = exisiting_cron + f"\n{cron_entry}\n"
        process = subprocess.run(["crontab", "-"], input=new_cron, text=True)

        if process.returncode == 0:
            print("Cron job created (macOS/Linux")
        else:
            print("Failed to create cron job")

    except Exception as e:
        print("Error setting up cron:", str(e))

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-python-path", default=sys.executable, help="Path to Python executable")
    args = parser.parse_args()

    python_exe = args.python_path
    script_path = os.path.abspath("Main.py")

    current_os = platform.system()
    print(f"Detected OS: {current_os}")

    if current_os == "Windows":
        setup_windows(python_exe, script_path)
    elif current_os in ["Linux", "Darwin"]:
        setup_unix(python_exe, script_path)
    else:
        print("Unsupported operating system")

if __name__ == "__main__":
    main()
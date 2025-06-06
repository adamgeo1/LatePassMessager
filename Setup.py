import os
import platform
import subprocess
import sys
import datetime
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true", default=False,
                        help="Has user enter paths to testing Google Sheets, does not run scheduler")
    parser.add_argument("--python-path", default=sys.executable, help="Path to Python executable")
    parser.add_argument("--test", action="store_true", default=False,
                        help="Schedule Main.py to run 1 minute from Script.py run")
    args = parser.parse_args()

    if args.setup:
        new_keys = ["RESPONSES_ID", "LATE_PASSES_ID", "TEST_RESPONSES_ID", "TEST_LATE_PASSES_ID", "GOOGLE_CREDS_ID"]
        existing = {}

        try:
            with open(".env", "r") as env:
                lines = env.readlines()
                for line in lines:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, val = line.split("=", 1)
                        existing[key] = val
        except FileNotFoundError:
            print("⚠️ No .env found. Generating .env, note that you will need to add Google Cloud credentials before "
                  "running Main.py")
            pass

        updated_lines = []
        processed_keys = set()

        for line in lines:
            stripped = line.strip()
            if stripped and not stripped.startswith("#") and "=" in stripped:
                key = stripped.split("=", 1)[0]
                if key in new_keys:
                    choice = input(f"{key} already exists with value '{existing[key]}'. Override? (Y/N): ").lower()
                    while choice not in {"y", "n", "yes", "no"}:
                        choice = input("You must say 'y' or 'n': ").lower()
                    if choice.startswith("y"):
                        new_val = input(f"Enter new value for {key}: ")
                        updated_lines.append(f"{key}={new_val}\n")
                    else:
                        updated_lines.append(line)
                    processed_keys.add(key)
                else:
                    updated_lines.append(line)
            else:
                updated_lines.append(line)

        for key in new_keys:
            if key not in processed_keys:
                new_val = input(f"Enter new value for new key {key} (or blank to skip): ")
                if new_val != "":
                    updated_lines.append(f"{key}={new_val}\n")

        with open(".env", "w") as env:
            env.writelines(updated_lines)

        print("✅ .env file updated.")

    if args.test:
        print("Testing Mode")

    python_exe = args.python_path
    script_path = os.path.abspath("Main.py")

    current_os = platform.system()
    print(f"Detected OS: {current_os}")

    if current_os == "Windows":
        import win32com.client

        task_name = "SendLatePassEmails"
        start_time = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(hour=0, minute=1, second=0,
                                                                                    microsecond=0) if not args.test else datetime.datetime.now() + datetime.timedelta(
            minutes=1)

        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()

        root_folder = scheduler.GetFolder("\\")
        task_def = scheduler.NewTask(0)

        task_def.RegistrationInfo.Description = "Send late pass emails for CS 270 automatically every Saturday at 12:01 AM"
        task_def.Settings.Enabled = True
        task_def.Settings.StopIfGoingOnBatteries = False
        task_def.Settings.DisallowStartIfOnBatteries = False
        task_def.Settings.RunOnlyIfNetworkAvailable = True

        trigger = task_def.Triggers.Create(3 if not args.test else 1)
        trigger.StartBoundary = start_time.strftime("%Y-%m-%dT%H:%M:%S")
        if args.test:
            trigger.DaysOfWeek = 64
            trigger.WeeksInterval = 1
        trigger.Enabled = True

        action = task_def.Actions.Create(0)
        action.Path = python_exe
        action.Arguments = f'"{script_path} --test"'
        action.WorkingDirectory = os.path.dirname(script_path)

        task_def.Principal.RunLevel = 1

        root_folder.RegisterTaskDefinition(task_name, task_def, 6, "", "", 0)

        print(f"✅ Scheduled task '{task_name}' created using Python at {python_exe}")

    elif current_os in ["Linux", "Darwin"]:
        if args.test:
            try:
                run_time = (datetime.datetime.now() + datetime.timedelta(minutes=1)).strftime("%H:%M")
                command = f'echo "{python_exe} \\"{script_path}\\" --test" | at {run_time}'
                process = subprocess.run(command, shell=True, text=True, capture_output=True)

                if process.returncode == 0:
                    print("One-time job scheduled with 'at'")
                else:
                    print("Failed to schedule job:", process.stderr.strip())

            except Exception as e:
                print("Error scheduling with 'at':", str(e))
        else:
            cron_entry = f"1 0 * * 6 {python_exe} \"{script_path}\""
            try:
                result = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
                existing_cron = result.stdout if result.returncode == 0 else ""
                if cron_entry in existing_cron:
                    print("Cron job already exists")
                    return

                new_cron = existing_cron + f"\n{cron_entry}\n"
                process = subprocess.run(["crontab", "-"], input=new_cron, text=True)

                if process.returncode == 0:
                    print("Cron job created (macOS/Linux")
                else:
                    print("Failed to create cron job")

            except Exception as e:
                print("Error setting up cron:", str(e))
    else:
        print("Unsupported operating system")


if __name__ == "__main__":
    main()

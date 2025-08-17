"""
Task Scheduler utilities using pywin32 (Schedule.Service COM API).

Features:
- Register or update a task (daily / once / at logon)
- Run a task on demand
- List tasks under any folder (default "\")
- Delete a task
- Sensible defaults and logging; sets working dir and highest privileges
"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
import argparse
import logging
import sys
import win32com.client as win32

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# Task Scheduler constants
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_INTERACTIVE_TOKEN = 3           # Run as current user (no password)
TASK_LOGON_PASSWORD = 1                    # Requires username/password (domain or local)
TASK_LOGON_SERVICE_ACCOUNT = 5             # e.g., SYSTEM (requires admin)
TASK_TRIGGER_TIME = 1
TASK_TRIGGER_DAILY = 2
TASK_TRIGGER_LOGON = 9
TASK_ACTION_EXEC = 0


@dataclass
class TaskParams:
    name: str
    exe: Path
    arguments: str = ""
    start_when_available: bool = True
    run_with_highest: bool = True
    wake_to_run: bool = False
    working_dir: Path | None = None
    folder: str = "D:\\LifeLongLearning\\windows-automation-python"
    # logon choice & creds
    logon_type: int = TASK_LOGON_INTERACTIVE_TOKEN
    username: str | None = None
    password: str | None = None


def _connect() -> any:
    svc = win32.Dispatch("Schedule.Service")
    svc.Connect()
    return svc


def _get_folder(svc, folder_path: str):
    try:
        return svc.GetFolder(folder_path)
    except Exception:
        # Create nested folders as needed
        parent = svc.GetFolder("\\")
        for part in [p for p in folder_path.strip("\\").split("\\") if p]:
            try:
                parent = parent.GetFolder("\\" + part)  # check if exists as absolute under root
            except Exception:
                parent = parent.CreateFolder(part)
        return parent


def _base_task_def(svc, params: TaskParams):
    td = svc.NewTask(0)

    # Registration info
    td.RegistrationInfo.Description = f"Created by Python for {params.exe.name}"
    td.RegistrationInfo.Author = params.username or "CurrentUser"

    # Settings
    s = td.Settings
    s.Enabled = True
    s.StartWhenAvailable = params.start_when_available
    s.Hidden = False
    s.RunOnlyIfIdle = False
    s.WakeToRun = params.wake_to_run
    s.MultipleInstances = 0  # Parallel disallowed (use 3 to queue)

    # Highest privileges (like “Run with highest privileges” checkbox)
    td.Principal.RunLevel = 1 if params.run_with_highest else 0  # 1=highest
    # (Optional) set td.Principal.UserId when using PASSWORD or SERVICE_ACCOUNT

    # Action (the thing to run)
    act = td.Actions.Create(TASK_ACTION_EXEC)
    act.Path = str(params.exe)
    act.Arguments = params.arguments or ""
    if params.working_dir:
        act.WorkingDirectory = str(params.working_dir)

    return td


def add_daily_trigger(td, start_time: datetime):
    t = td.Triggers.Create(TASK_TRIGGER_DAILY)
    t.DaysInterval = 1
    t.StartBoundary = start_time.strftime("%Y-%m-%dT%H:%M:%S")


def add_once_trigger(td, run_at: datetime):
    t = td.Triggers.Create(TASK_TRIGGER_TIME)
    t.StartBoundary = run_at.strftime("%Y-%m-%dT%H:%M:%S")


def add_logon_trigger(td):
    td.Triggers.Create(TASK_TRIGGER_LOGON)


def register_task(params: TaskParams, trigger: str, when: datetime | None = None) -> None:
    svc = _connect()
    folder = _get_folder(svc, params.folder)
    td = _base_task_def(svc, params)

    # Triggers
    if trigger == "daily":
        add_daily_trigger(td, when or (datetime.now() + timedelta(minutes=2)))
    elif trigger == "once":
        add_once_trigger(td, when or (datetime.now() + timedelta(minutes=2)))
    elif trigger == "logon":
        add_logon_trigger(td)
    else:
        raise ValueError("trigger must be one of: daily | once | logon")

    user_id = params.username or ""  # empty means current user for INTERACTIVE_TOKEN
    pwd = params.password or ""

    # For SERVICE_ACCOUNT (e.g., SYSTEM), set user_id="SYSTEM" and no password.
    # For PASSWORD, provide username and password (DOMAIN\\user or .\\localuser).
    folder.RegisterTaskDefinition(
        params.name, td, TASK_CREATE_OR_UPDATE,
        user_id, pwd, params.logon_type
    )
    logging.info("Registered/updated task '%s' in '%s'", params.name, params.folder)


def run_task(name: str, folder: str = "\\") -> None:
    svc = _connect()
    f = _get_folder(svc, folder)
    task = f.GetTask(name)
    task.Run("")  # parameters for COM tasks are rarely used
    logging.info("Triggered run for '%s'", name)


def delete_task(name: str, folder: str = "\\") -> None:
    svc = _connect()
    f = _get_folder(svc, folder)
    f.DeleteTask(name, 0)
    logging.info("Deleted task '%s'", name)


def list_tasks(folder: str = "\\") -> None:
    svc = _connect()
    f = _get_folder(svc, folder)
    for t in f.GetTasks(0):
        # State: 0=Unknown 1=Disabled 2=Queued 3=Ready 4=Running
        logging.info("Task: %-30s | State=%s | NextRun=%s | Path=%s",
                     t.Name, t.State, t.NextRunTime, t.Path)

# ---------------- CLI ---------------- #


def main() -> int:
    p = argparse.ArgumentParser(description="Windows Task Scheduler automation")
    sub = p.add_subparsers(dest="cmd", required=True)

    # register
    pr = sub.add_parser("register", help="Register or update a task")
    pr.add_argument("--name", required=True)
    pr.add_argument("--exe", required=True, help="Path to EXE or script interpreter")
    pr.add_argument("--args", default="", help="Arguments passed to exe")
    pr.add_argument("--folder", default="\\")
    pr.add_argument("--trigger", choices=["daily", "once", "logon"], required=True)
    pr.add_argument("--at", help="When to run (YYYY-mm-ddTHH:MM:SS). If omitted, +2 minutes.")
    pr.add_argument("--workdir", help="Working directory (important for scripts)")
    pr.add_argument("--highest", action="store_true", help="Run with highest privileges")
    pr.add_argument("--wake", action="store_true", help="Wake the computer to run")
    # logon options
    pr.add_argument("--logon", choices=["interactive", "password", "service"], default="interactive")
    pr.add_argument("--username")
    pr.add_argument("--password")

    # run
    prun = sub.add_parser("run", help="Run a task now")
    prun.add_argument("--name", required=True)
    prun.add_argument("--folder", default="\\")

    # delete
    pd = sub.add_parser("delete", help="Delete a task")
    pd.add_argument("--name", required=True)
    pd.add_argument("--folder", default="\\")

    # list
    pl = sub.add_parser("list", help="List tasks in a folder")
    pl.add_argument("--folder", default="\\")

    args = p.parse_args()

    if args.cmd == "register":
        dt = None
        if args.at:
            dt = datetime.strptime(args.at, "%Y-%m-%dT%H:%M:%S")
        logon_map = {
            "interactive": TASK_LOGON_INTERACTIVE_TOKEN,
            "password": TASK_LOGON_PASSWORD,
            "service": TASK_LOGON_SERVICE_ACCOUNT,
        }
        workdir = Path(args.workdir) if args.workdir else None

        params = TaskParams(
            name=args.name,
            exe=Path(args.exe),
            arguments=args.args,
            working_dir=workdir,
            folder=args.folder,
            run_with_highest=args.highest,
            wake_to_run=args.wake,
            logon_type=logon_map[args.logon],
            username=args.username,
            password=args.password,
        )
        register_task(params, trigger=args.trigger, when=dt)
        return 0

    if args.cmd == "run":
        run_task(args.name, folder=args.folder)
        return 0

    if args.cmd == "delete":
        delete_task(args.name, folder=args.folder)
        return 0

    if args.cmd == "list":
        list_tasks(folder=args.folder)
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
    """
    python schedule_task.py register 
  --name "Daily_BrightDay" 
  --exe "D:\LifeLongLearning\windows-automation-python\.venv\Scripts\python.exe" 
  --args "D:\Jobs\myscript.py" 
  --workdir "D:\Jobs" 
  --trigger daily 
  --at 2025-08-18T09:00:00 
  --highest
    """
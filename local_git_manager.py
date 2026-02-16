from __future__ import annotations
import os
import subprocess
import sys
import shutil
import ctypes
from ctypes import wintypes
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from typing import Optional, Dict, Any
import json


# ==============================
# CONFIG & GLOBALS (SELF-CONTAINED)
# ==============================

SCRIPT_DIR = Path(__file__).resolve().parent
APP_NAME = "LocalGitControlPanel"


def _default_data_dir() -> Path:
    """
    Where we store config/launchers/venv by default.
    This avoids 'script folder not writable' problems (Program Files, etc).
    Set PORTABLE=1 env var if you want everything next to the script.
    """
    portable = os.environ.get("PORTABLE", "").strip().lower() in ("1", "true", "yes")
    if portable:
        return SCRIPT_DIR

    localapp = os.environ.get("LOCALAPPDATA")
    if localapp:
        return Path(localapp) / APP_NAME

    return Path.home() / f".{APP_NAME.lower()}"


DATA_DIR = _default_data_dir()
CONFIG_FILE = DATA_DIR / "config.json"

DEFAULT_CONFIG = {
    "base_dir": str((DATA_DIR / "projects").resolve()),
    "editor_path": None,
    "global_venv_path": str((DATA_DIR / "global_venv").resolve()),
    "launchers_dir": str((DATA_DIR / "_launchers").resolve()),
}

# runtime globals
BASE_DIR = Path(DEFAULT_CONFIG["base_dir"])
EDITOR_PATH: Optional[str] = None
_CFG: Optional[Dict[str, Any]] = None  # set in main()
LAUNCHERS_DIR = Path(DEFAULT_CONFIG["launchers_dir"])
_CFG: dict | None = None  # set in main()


def load_config() -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=2), encoding="utf-8")
        return dict(DEFAULT_CONFIG)

    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        merged = dict(DEFAULT_CONFIG)
        if isinstance(cfg, dict):
            # merge defaults so missing keys never crash you
            merged.update(cfg)
        return merged
    except Exception:
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=2), encoding="utf-8")
        return dict(DEFAULT_CONFIG)


def save_config(cfg: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def apply_config(cfg: dict) -> None:
    global BASE_DIR, EDITOR_PATH, LAUNCHERS_DIR
    BASE_DIR = Path(cfg.get("base_dir") or DEFAULT_CONFIG["base_dir"])
    EDITOR_PATH = cfg.get("editor_path")
    LAUNCHERS_DIR = Path(cfg.get("launchers_dir") or DEFAULT_CONFIG["launchers_dir"])


def init_config() -> dict:
    cfg = load_config()
    apply_config(cfg)
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    LAUNCHERS_DIR.mkdir(parents=True, exist_ok=True)
    return cfg


# ==============================
# GLOBAL VENV (uses config.json)
# ==============================

def get_global_venv_path() -> Path:
    global _CFG
    p = (_CFG or DEFAULT_CONFIG).get("global_venv_path") or DEFAULT_CONFIG["global_venv_path"]
    return Path(p).resolve()


def save_global_venv_path(path: Path):
    global _CFG
    if _CFG is None:
        _CFG = init_config()
    _CFG["global_venv_path"] = str(path.resolve())
    save_config(_CFG)


def is_valid_venv(venv_path: Path) -> bool:
    if not venv_path or not venv_path.exists():
        return False
    scripts = venv_path / "Scripts"
    return scripts.exists() and (scripts / "activate.bat").exists() and (scripts / "python.exe").exists()


def get_global_python() -> Optional[Path]:
    venv = get_global_venv_path()
    return (venv / "Scripts" / "python.exe") if is_valid_venv(venv) else None


def find_system_python() -> Optional[str]:
    # Prefer current interpreter if it's not clearly inside a venv
    if sys.executable and Path(sys.executable).exists():
        exe = Path(sys.executable).resolve()
        if "venv" not in (p.lower() for p in exe.parts) and ".venv" not in (p.lower() for p in exe.parts):
            return str(exe)

    for cmd in ["python", "python3", "py"]:
        found = shutil.which(cmd)
        if found:
            return found

    username = os.environ.get("USERNAME", "")
    for ver in ["313", "312", "311", "310", "39"]:
        for loc in [
            f"C:\\Python{ver}\\python.exe",
            f"C:\\Program Files\\Python{ver}\\python.exe",
            f"C:\\Users\\{username}\\AppData\\Local\\Programs\\Python\\Python{ver}\\python.exe",
        ]:
            if Path(loc).exists():
                return loc
    return None


def create_global_venv() -> bool:
    venv_path = get_global_venv_path()
    if is_valid_venv(venv_path):
        return True

    if venv_path.exists():
        try:
            shutil.rmtree(venv_path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot remove:\n{venv_path}\n\n{e}")
            return False

    python_exe = find_system_python()
    if not python_exe:
        messagebox.showerror("Error", "Python not found!")
        return False

    progress = tk.Toplevel()
    progress.title("Creating Global Venv")
    progress.geometry("380x100")
    progress.resizable(False, False)
    progress.grab_set()
    progress.update_idletasks()
    x, y = (progress.winfo_screenwidth() - 380) // 2, (progress.winfo_screenheight() - 100) // 2
    progress.geometry(f"380x100+{x}+{y}")
    tk.Label(progress, text="Creating global venv...", font=("", 10), pady=30).pack()
    progress.update()

    try:
        result = subprocess.run([python_exe, "-m", "venv", str(venv_path)], capture_output=True, text=True, timeout=180)
        success = result.returncode == 0 and is_valid_venv(venv_path)
    except Exception:
        success = False

    progress.destroy()

    if success:
        save_global_venv_path(venv_path)
    else:
        messagebox.showerror("Failed", "Venv creation failed")
    return success


def ensure_global_venv() -> bool:
    if is_valid_venv(get_global_venv_path()):
        return True
    if messagebox.askyesno("Global Venv", "No global venv found.\n\nCreate now?"):
        return create_global_venv()
    return False


def install_ipykernel():
    python = get_global_python()
    if not python:
        messagebox.showerror("Error", "Global venv not found")
        return

    progress = tk.Toplevel()
    progress.title("Installing Kernel")
    progress.geometry("300x80")
    progress.grab_set()
    tk.Label(progress, text="Installing ipykernel...", pady=25).pack()
    progress.update()

    try:
        subprocess.run([str(python), "-m", "pip", "install", "ipykernel"], capture_output=True, timeout=300)
        subprocess.run(
            [str(python), "-m", "ipykernel", "install", "--user", "--name=global_venv", "--display-name=Python (Global Venv)"],
            capture_output=True,
            timeout=60,
        )
        progress.destroy()
        messagebox.showinfo("Done", "Kernel installed!\nSelect 'Python (Global Venv)' in notebooks.")
    except Exception as e:
        progress.destroy()
        messagebox.showwarning("Error", str(e))


# ==============================
# HELPERS
# ==============================

def get_parent_pid():
    try:
        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
        if snapshot == -1:
            return None
        try:
            pe32 = PROCESSENTRY32()
            pe32.dwSize = ctypes.sizeof(PROCESSENTRY32)
            if not kernel32.Process32First(snapshot, ctypes.byref(pe32)):
                return None
            pid = os.getpid()
            while True:
                if pe32.th32ProcessID == pid:
                    return pe32.th32ParentProcessID
                if not kernel32.Process32Next(snapshot, ctypes.byref(pe32)):
                    break
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception:
        pass
    return None


def discover_editor_candidates():
    candidates = {}
    username = os.environ.get("USERNAME", "")
    for path in [
        r"C:\Program Files\cursor\Cursor.exe",
        fr"C:\Users\{username}\AppData\Local\Programs\Cursor\Cursor.exe",
        fr"C:\Users\{username}\AppData\Local\Programs\Microsoft VS Code\Code.exe",
        r"C:\Program Files\Microsoft VS Code\Code.exe",
    ]:
        if Path(path).exists():
            candidates.setdefault(Path(path).stem, path)
    return candidates


def set_editor_path():
    global EDITOR_PATH, _CFG
    path = filedialog.askopenfilename(title="Select editor", filetypes=[("Programs", ".exe")])
    if path:
        EDITOR_PATH = path
        _CFG["editor_path"] = path
        save_config(_CFG)


def get_editor_path_or_ask():
    global EDITOR_PATH, _CFG

    if EDITOR_PATH and Path(EDITOR_PATH).exists():
        return EDITOR_PATH

    candidates = discover_editor_candidates()

    if len(candidates) == 1:
        EDITOR_PATH = next(iter(candidates.values()))
        _CFG["editor_path"] = EDITOR_PATH
        save_config(_CFG)
        return EDITOR_PATH

    if candidates:
        win = tk.Toplevel()
        win.title("Choose Editor")
        tk.Label(win, text="Select:").pack(padx=10, pady=5)
        var = tk.StringVar(value=list(candidates.values())[0])

        for name, path in candidates.items():
            tk.Radiobutton(win, text=name, variable=var, value=path).pack(anchor="w", padx=10)

        def ok():
            global EDITOR_PATH, _CFG
            EDITOR_PATH = var.get()
            _CFG["editor_path"] = EDITOR_PATH
            save_config(_CFG)
            win.destroy()

        tk.Button(win, text="OK", command=ok).pack(pady=10)
        win.grab_set()
        win.wait_window()
        return EDITOR_PATH

    set_editor_path()
    return EDITOR_PATH


def run_cmd(cmd, cwd=None, allow_fail=False):
    try:
        return subprocess.run(cmd, cwd=cwd, check=not allow_fail, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        if not allow_fail:
            messagebox.showerror("Error", f"{' '.join(cmd)}\n\n{e.stderr}")
    except FileNotFoundError:
        messagebox.showerror("Error", f"Not found: {cmd[0]}")
    return None


def ensure_dirs():
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    LAUNCHERS_DIR.mkdir(parents=True, exist_ok=True)


def list_projects():
    ensure_dirs()
    return sorted(p.name for p in BASE_DIR.iterdir() if p.is_dir() and (p / "origin.git").exists())


def list_workspaces(project: str):
    ws_dir = BASE_DIR / project / "workspaces"
    return sorted(d.name for d in ws_dir.iterdir() if d.is_dir()) if ws_dir.exists() else []


def list_launchers():
    return sorted(f.name for f in LAUNCHERS_DIR.iterdir() if f.suffix.lower() == ".bat") if LAUNCHERS_DIR.exists() else []


# ==============================
# WORKSPACE FILES GENERATION
# ==============================

def write_workspace_files(project: str, ws_dir: Path):
    """Generate all workspace files: start_session.bat, notebook_helper.py, .vscode/settings.json"""
    venv = get_global_venv_path()
    activate = venv / "Scripts" / "activate.bat"
    python_exe = venv / "Scripts" / "python.exe"
    ws_abs = ws_dir.resolve()

    # 1) start_session.bat
    start_session = ws_dir / "start_session.bat"
    start_session.write_text(
        f'''@echo off
cd /d "{ws_abs}"

echo.
echo ================================================================
echo  PROJECT:   {project}
echo  WORKSPACE: {ws_dir.name}
echo ================================================================
echo.

REM Activate global venv
if exist "{activate}" (
    call "{activate}"
    echo  [VENV] Activated: {venv.name}
) else (
    echo  [!] Global venv not found: {venv}
)

echo.
echo ----------------------------------------------------------------
echo  GIT BRANCH:
echo ----------------------------------------------------------------
git branch
echo.
echo ----------------------------------------------------------------
echo  GIT STATUS:
echo ----------------------------------------------------------------
git status
echo.
echo ================================================================
echo  Session ready!
echo ================================================================
echo.
''',
        encoding="utf-8",
    )

    # 2) notebook_helper.py
    notebook_helper = ws_dir / "notebook_helper.py"
    notebook_helper.write_text(
        f'''"""
Notebook Helper - Run setup_venv() in your first cell

Usage:
    from notebook_helper import setup_venv
    setup_venv()
"""
import sys
import os
from pathlib import Path

VENV_PATH = Path(r"{venv}")
VENV_SCRIPTS = VENV_PATH / "Scripts"
VENV_SITE_PACKAGES = VENV_PATH / "Lib" / "site-packages"

def setup_venv():
    """Add global venv to path for this notebook session."""
    if not VENV_PATH.exists():
        print(f"[!] Venv not found: {{VENV_PATH}}")
        return False

    if not VENV_SITE_PACKAGES.exists():
        print(f"[!] site-packages not found: {{VENV_SITE_PACKAGES}}")
        return False

    sp = str(VENV_SITE_PACKAGES)
    if sp not in sys.path:
        sys.path.insert(0, sp)

    os.environ["VIRTUAL_ENV"] = str(VENV_PATH)
    current_path = os.environ.get("PATH", "")
    if str(VENV_SCRIPTS) not in current_path:
        os.environ["PATH"] = str(VENV_SCRIPTS) + os.pathsep + current_path

    print(f"[OK] Venv active: {{VENV_PATH.name}}")
    return True

def pip_install(*packages):
    """Install packages to global venv."""
    import subprocess
    python = VENV_SCRIPTS / "python.exe"
    cmd = [str(python), "-m", "pip", "install"] + list(packages)
    print(f"Running: {{' '.join(cmd)}}")
    return subprocess.call(cmd)
''',
        encoding="utf-8",
    )

    # 3) .vscode/settings.json
    vscode_dir = ws_dir / ".vscode"
    vscode_dir.mkdir(exist_ok=True)
    settings_file = vscode_dir / "settings.json"

    settings = {"python.defaultInterpreterPath": str(python_exe)}

    if settings_file.exists():
        try:
            existing = json.loads(settings_file.read_text(encoding="utf-8"))
            if isinstance(existing, dict):
                existing.update(settings)
                settings = existing
        except Exception:
            pass

    settings_file.write_text(json.dumps(settings, indent=4), encoding="utf-8")
    return True


# ==============================
# IMPORT FOLDER TO GIT
# ==============================

def import_folder_to_project():
    """Import an existing folder into the git system as a new project."""
    source = filedialog.askdirectory(title="Select folder to import")
    if not source:
        return None
    source = Path(source)

    if not source.exists():
        messagebox.showerror("Error", "Folder not found")
        return None

    default_name = source.name.replace(" ", "_")
    project_name = simpledialog.askstring("Project Name", "Name for this project:", initialvalue=default_name)
    if not project_name:
        return None
    project_name = project_name.strip().replace(" ", "_")

    project_dir = BASE_DIR / project_name
    if project_dir.exists():
        base_name = project_name + "_local"
        candidate = BASE_DIR / base_name
        counter = 2

        while candidate.exists():
            candidate = BASE_DIR / f"{base_name}{counter}"
            counter += 1

        project_name = candidate.name
        project_dir = candidate

    existing_venvs = []
    for venv_name in ["venv", ".venv", "env", ".env", "virtualenv"]:
        venv_path = source / venv_name
        if venv_path.exists() and (venv_path / "Scripts").exists():
            existing_venvs.append(venv_name)

    if existing_venvs:
        msg = "Found existing venv folder(s):\n• " + "\n• ".join(existing_venvs)
        msg += "\n\nThese will be KEPT for backup.\nYou can delete them manually later."
        msg += "\n\nContinue?"
        if not messagebox.askyesno("Existing Venv Found", msg):
            return None

    try:
        project_dir.mkdir(parents=True)
        origin_dir = project_dir / "origin.git"
        workspaces_dir = project_dir / "workspaces"
        workspaces_dir.mkdir()
        run_cmd(["git", "init", "--bare", str(origin_dir)], cwd=str(project_dir))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create project:\n{e}")
        return None

    ws_name = simpledialog.askstring("Workspace Name", "Name for workspace:", initialvalue="main") or "main"
    ws_name = ws_name.strip().replace(" ", "_")
    ws_dir = workspaces_dir / ws_name

    progress = tk.Toplevel()
    progress.title("Importing")
    progress.geometry("350x100")
    progress.resizable(False, False)
    progress.grab_set()
    progress.update_idletasks()
    x, y = (progress.winfo_screenwidth() - 350) // 2, (progress.winfo_screenheight() - 100) // 2
    progress.geometry(f"350x100+{x}+{y}")
    tk.Label(progress, text=f"Copying files...\n\n{source.name}", font=("", 9), pady=20).pack()
    progress.update()

    try:
        shutil.copytree(source, ws_dir)
    except Exception as e:
        progress.destroy()
        messagebox.showerror("Error", f"Failed to copy:\n{e}")
        return None

    progress.destroy()

    existing_git = ws_dir / ".git"
    if existing_git.exists():
        try:
            shutil.rmtree(existing_git)
        except Exception:
            pass

    run_cmd(["git", "init"], cwd=str(ws_dir))
    run_cmd(["git", "remote", "add", "origin", str(origin_dir)], cwd=str(ws_dir))
    run_cmd(["git", "add", "."], cwd=str(ws_dir))
    run_cmd(["git", "commit", "-m", "Initial import"], cwd=str(ws_dir), allow_fail=True)
    run_cmd(["git", "branch", "-M", ws_name], cwd=str(ws_dir), allow_fail=True)
    run_cmd(["git", "push", "-u", "origin", ws_name], cwd=str(ws_dir), allow_fail=True)

    write_workspace_files(project_name, ws_dir)

    msg = "Project imported!\n\n"
    msg += f"Project: {project_name}\n"
    msg += f"Workspace: {ws_name}\n\n"
    msg += "Files created:\n"
    msg += "• start_session.bat\n"
    msg += "• notebook_helper.py\n"
    msg += "• .vscode/settings.json\n"
    if existing_venvs:
        msg += "\nOld venvs kept:\n• " + "\n• ".join(existing_venvs)

    messagebox.showinfo("Import Complete", msg)
    return project_name


# ==============================
# LAUNCHER
# ==============================

def generate_launcher(project: str, ws_name: str) -> Path | None:
    editor = get_editor_path_or_ask()
    if not editor:
        return None

    ws_dir = BASE_DIR / project / "workspaces" / ws_name
    if not ws_dir.exists():
        return None

    LAUNCHERS_DIR.mkdir(parents=True, exist_ok=True)
    launcher = LAUNCHERS_DIR / f"{project}__{ws_name}.bat"

    launcher.write_text(
        f'''@echo off
start "" "{editor}" "{ws_dir}"
exit
''',
        encoding="utf-8",
    )
    return launcher


def generate_launcher_selected(lb_p, lb_w, lb_l):
    sel_p, sel_w = lb_p.curselection(), lb_w.curselection()
    if not sel_p or not sel_w:
        messagebox.showerror("Error", "Select project and workspace")
        return
    project, ws = lb_p.get(sel_p[0]), lb_w.get(sel_w[0])
    launcher = generate_launcher(project, ws)
    if launcher:
        messagebox.showinfo("Created", f"Launcher: {launcher.name}")
        refresh_launchers(lb_l)

def is_git_repo(path: Path) -> bool:
    return path.is_dir() and (path / ".git").exists()

def get_current_branch(repo_dir: Path) -> str:
    r = run_cmd(["git", "rev-parse", "--abbrev-ref", "HEAD"], cwd=str(repo_dir), allow_fail=True)
    if r and r.returncode == 0:
        b = (r.stdout or "").strip()
        if b and b != "HEAD":
            return b
    return "main"

def repo_has_remote(repo_dir: Path) -> bool:
    r = run_cmd(["git", "remote"], cwd=str(repo_dir), allow_fail=True)
    if not r or r.returncode != 0:
        return False
    return bool((r.stdout or "").strip())

from typing import Optional  # only needed if not already imported

def adopt_existing_repo_folder(repo_dir: Path) -> Optional[str]:
    """
    Convert a plain 'git clone' folder into Control Panel project format:
      BASE_DIR/<project>_local*/origin.git
      BASE_DIR/<project>_local*/workspaces/<branch>/
    Works reliably even when repo_dir is inside BASE_DIR by staging into a temp folder first.
    """
    repo_dir = repo_dir.resolve()

    if not repo_dir.exists() or not is_git_repo(repo_dir):
        messagebox.showerror("Error", f"Not a git repo:\n{repo_dir}")
        return None

    project_name = repo_dir.name.replace(" ", "_").strip()
    if not project_name:
        messagebox.showerror("Error", "Invalid folder name")
        return None

    # Decide target project folder (avoid collision)
    project_dir = (BASE_DIR / project_name).resolve()
    if project_dir.exists():
        base_name = project_name + "_local"
        candidate = (BASE_DIR / base_name).resolve()
        counter = 2
        while candidate.exists():
            candidate = (BASE_DIR / f"{base_name}{counter}").resolve()
            counter += 1
        project_name = candidate.name
        project_dir = candidate

    # Determine workspace name from current branch
    branch = get_current_branch(repo_dir)
    ws_name = (branch.replace("/", "_") or "main").strip()

    # Optional fetch if remote exists
    if repo_has_remote(repo_dir):
        if messagebox.askyesno("Fetch remote?", f"Repo '{repo_dir.name}' has a remote.\n\nFetch --all before adopting?"):
            run_cmd(["git", "fetch", "--all"], cwd=str(repo_dir), allow_fail=True)

    # ---- STAGE REPO INTO TEMP LOCATION (fixes "adopt from root" issues) ----
    staged_repo = repo_dir
    temp_repo = None

    try:
        # If repo is directly under BASE_DIR, stage it to a temp name first
        # (this avoids Windows move/merge edge cases)
        try:
            is_under_root = staged_repo.parent.resolve() == BASE_DIR.resolve()
        except Exception:
            is_under_root = str(staged_repo.parent).lower() == str(BASE_DIR).lower()

        if is_under_root:
            temp_repo = (BASE_DIR / f"__adopt_tmp__{staged_repo.name}").resolve()
            # ensure temp name is unique
            k = 2
            while temp_repo.exists():
                temp_repo = (BASE_DIR / f"__adopt_tmp__{staged_repo.name}_{k}").resolve()
                k += 1

            # rename (fast) instead of move
            staged_repo.rename(temp_repo)
            staged_repo = temp_repo

        # Now create managed structure
        project_dir.mkdir(parents=True, exist_ok=False)
        origin_dir = project_dir / "origin.git"
        workspaces_dir = project_dir / "workspaces"
        workspaces_dir.mkdir()

        run_cmd(["git", "init", "--bare", str(origin_dir)], cwd=str(project_dir))

        ws_dir = workspaces_dir / ws_name

        # Move staged repo into workspace
        shutil.move(str(staged_repo), str(ws_dir))

        # Reset origin to local bare
        run_cmd(["git", "remote", "remove", "origin"], cwd=str(ws_dir), allow_fail=True)
        run_cmd(["git", "remote", "add", "origin", str(origin_dir)], cwd=str(ws_dir), allow_fail=True)

        # Ensure branch exists + push to local bare
        run_cmd(["git", "checkout", "-B", ws_name], cwd=str(ws_dir), allow_fail=True)
        run_cmd(["git", "push", "-u", "origin", ws_name], cwd=str(ws_dir), allow_fail=True)

        # Generate helper files
        write_workspace_files(project_name, ws_dir)

        messagebox.showinfo(
            "Adopted!",
            f"Imported existing repo into Control Panel format:\n\n"
            f"Project: {project_name}\nWorkspace: {ws_name}\n\n"
            f"Created:\n• origin.git\n• workspaces/{ws_name}\n• start_session.bat + notebook_helper.py + .vscode/settings.json"
        )
        return project_name
    
    except PermissionError as e:
        messagebox.showerror(
            "Adopt failed (file lock)",
            f"Windows is blocking the move/rename.\n\n"
            f"Close any editor/terminal opened inside the repo folder and try again.\n\n"
            f"Details:\n{e}"
        )
        return None

    except Exception as e:
        messagebox.showerror("Adopt failed", f"{e}")
        return None

def find_unmanaged_git_repos() -> list[Path]:
    """
    Unmanaged repo = folder directly under BASE_DIR containing .git,
    but NOT already a managed project (doesn't have origin.git).
    """
    ensure_dirs()
    unmanaged = []
    for p in BASE_DIR.iterdir():
        if not p.is_dir():
            continue
        # skip managed projects
        if (p / "origin.git").exists():
            continue
        # detect plain clones in root
        if is_git_repo(p):
            unmanaged.append(p)
    return unmanaged


# ==============================
# ACTIONS
# ==============================

def refresh_launchers(lb):
    lb.delete(0, tk.END)
    for name in list_launchers():
        lb.insert(tk.END, name)

def adopt_repo_button(lb_p, lb_w, status):
    repo = filedialog.askdirectory(title="Select an existing git repo folder (.git)")
    if not repo:
        return
    repo_dir = Path(repo)
    adopted = adopt_existing_repo_folder(repo_dir)
    if adopted:
        refresh_projects(lb_p, lb_w, status)

def run_launcher_and_exit(lb_l, root):
    sel = lb_l.curselection()
    if not sel:
        messagebox.showerror("Error", "Select a launcher")
        return
    bat = LAUNCHERS_DIR / lb_l.get(sel[0])
    if not bat.exists():
        messagebox.showerror("Error", "Not found")
        return

    parent = get_parent_pid()
    subprocess.Popen(f'cmd /c "{bat}"', creationflags=subprocess.CREATE_NO_WINDOW)

    try:
        root.quit()
        root.destroy()
    except Exception:
        pass

    if parent:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(parent)], capture_output=True, creationflags=0x08000000)
        except Exception:
            pass

    os._exit(0)


def set_root_folder(lb_p, lb_w, status):
    global BASE_DIR, _CFG
    if path := filedialog.askdirectory(title="Projects folder"):
        BASE_DIR = Path(path)
        BASE_DIR.mkdir(parents=True, exist_ok=True)
        _CFG["base_dir"] = str(BASE_DIR.resolve())
        save_config(_CFG)
        refresh_projects(lb_p, lb_w, status)


def refresh_projects(lb_p, lb_w, status):
    lb_p.delete(0, tk.END)
    lb_w.delete(0, tk.END)
    for p in list_projects():
        lb_p.insert(tk.END, p)
    update_status(status)

    unmanaged = find_unmanaged_git_repos()
    if unmanaged:
        if messagebox.askyesno(
            "Unmanaged repos found",
            f"Found {len(unmanaged)} git repo folder(s) directly in the root.\n\n"
            "Do you want to adopt them into the Control Panel format now?"
        ):
            for repo_dir in unmanaged:
                if messagebox.askyesno("Adopt repo?", f"Adopt this repo?\n\n{repo_dir.name}"):
                    adopted = adopt_existing_repo_folder(repo_dir)
                    if adopted:
                        lb_p.delete(0, tk.END)
                        for p in list_projects():
                            lb_p.insert(tk.END, p)


def update_status(status, project=None):
    venv_ok = "✓" if is_valid_venv(get_global_venv_path()) else "✗"
    status.config(text=f"{'Project: ' + project if project else 'Root: ' + str(BASE_DIR)}  |  Venv: {venv_ok}")


def on_project_select(e, lb_p, lb_w, status):
    lb_w.delete(0, tk.END)
    if sel := lb_p.curselection():
        project = lb_p.get(sel[0])
        for ws in list_workspaces(project):
            lb_w.insert(tk.END, ws)
        update_status(status, project)


def create_project(lb_p, lb_w, lb_l, status):
    if not ensure_global_venv():
        return

    name = simpledialog.askstring("New Project", "Name:")
    if not name:
        return
    name = name.strip().replace(" ", "_")

    pdir = BASE_DIR / name
    if pdir.exists():
        messagebox.showerror("Error", "Already exists")
        return

    origin = pdir / "origin.git"
    wsdir = pdir / "workspaces"
    try:
        wsdir.mkdir(parents=True)
        run_cmd(["git", "init", "--bare", str(origin)], cwd=str(pdir))
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

    ws_name = simpledialog.askstring("Workspace", "First workspace:", initialvalue="main") or "main"
    ws_name = ws_name.strip().replace(" ", "_")
    ws = wsdir / ws_name

    run_cmd(["git", "clone", str(origin), str(ws)], cwd=str(pdir))
    run_cmd(["git", "switch", "-c", ws_name], cwd=str(ws), allow_fail=True)
    run_cmd(["git", "push", "-u", "origin", ws_name], cwd=str(ws), allow_fail=True)

    write_workspace_files(name, ws)

    messagebox.showinfo(
        "Created",
        f"Project: {name}\nWorkspace: {ws_name}\n\nFiles:\n• start_session.bat\n• notebook_helper.py\n• .vscode/settings.json",
    )
    refresh_projects(lb_p, lb_w, status)


def create_workspace(lb_p, lb_w, lb_l, status):
    if not (sel := lb_p.curselection()):
        messagebox.showerror("Error", "Select project first")
        return

    project = lb_p.get(sel[0])
    pdir = BASE_DIR / project
    origin = pdir / "origin.git"
    wsdir = pdir / "workspaces"
    wsdir.mkdir(exist_ok=True)

    ws_name = simpledialog.askstring("Workspace", "Name:")
    if not ws_name:
        return
    ws_name = ws_name.strip().replace(" ", "_")

    ws = wsdir / ws_name
    if ws.exists():
        messagebox.showerror("Error", "Already exists")
        return

    run_cmd(["git", "clone", str(origin), str(ws)], cwd=str(pdir))
    run_cmd(["git", "switch", "-c", ws_name], cwd=str(ws), allow_fail=True)
    run_cmd(["git", "push", "-u", "origin", ws_name], cwd=str(ws), allow_fail=True)

    write_workspace_files(project, ws)

    messagebox.showinfo(
        "Created",
        f"Workspace: {ws_name}\n\nFiles:\n• start_session.bat\n• notebook_helper.py\n• .vscode/settings.json",
    )

    lb_w.delete(0, tk.END)
    for w in list_workspaces(project):
        lb_w.insert(tk.END, w)


def regen_workspace_files(lb_p, lb_w):
    sel_p, sel_w = lb_p.curselection(), lb_w.curselection()
    if not sel_p or not sel_w:
        messagebox.showerror("Error", "Select project and workspace")
        return

    project = lb_p.get(sel_p[0])
    ws_name = lb_w.get(sel_w[0])
    ws_dir = BASE_DIR / project / "workspaces" / ws_name

    if not ws_dir.exists():
        messagebox.showerror("Error", f"Workspace not found:\n{ws_dir}")
        return

    write_workspace_files(project, ws_dir)
    messagebox.showinfo(
        "Regenerated",
        f"Files regenerated for {ws_name}:\n\n• start_session.bat\n• notebook_helper.py\n• .vscode/settings.json",
    )


def do_import(lb_p, lb_w, status):
    project_name = import_folder_to_project()
    if project_name:
        refresh_projects(lb_p, lb_w, status)


def setup_venv_btn(status):
    venv = get_global_venv_path()
    if is_valid_venv(venv):
        if messagebox.askyesno("Venv Exists", "Reinstall Jupyter kernel?"):
            install_ipykernel()
    else:
        if create_global_venv():
            if messagebox.askyesno("Kernel", "Install Jupyter kernel?"):
                install_ipykernel()
    update_status(status)


def show_venv_status():
    venv = get_global_venv_path()
    msg = f"Path: {venv}\n\n"
    msg += f"Valid: {'Yes' if is_valid_venv(venv) else 'No'}\n\n"
    if venv.exists():
        scripts = venv / "Scripts"
        msg += f"Scripts/: {'Yes' if scripts.exists() else 'No'}\n"
        if scripts.exists():
            msg += f"activate.bat: {'Yes' if (scripts / 'activate.bat').exists() else 'No'}\n"
            msg += f"python.exe: {'Yes' if (scripts / 'python.exe').exists() else 'No'}"
    messagebox.showinfo("Global Venv", msg)


def export_requirements(lb_p, lb_w):
    sel_p, sel_w = lb_p.curselection(), lb_w.curselection()
    if not sel_p or not sel_w:
        messagebox.showerror("Error", "Select project and workspace")
        return

    project = lb_p.get(sel_p[0])
    ws_name = lb_w.get(sel_w[0])
    ws_dir = BASE_DIR / project / "workspaces" / ws_name
    if not ws_dir.exists():
        messagebox.showerror("Error", f"Workspace not found:\n{ws_dir}")
        return

    venv = get_global_venv_path()
    python_exe = venv / "Scripts" / "python.exe"
    if not python_exe.exists():
        messagebox.showerror(
            "Error",
            f"Global venv python not found:\n{python_exe}\n\nCreate the global venv first.",
        )
        return

    out_file = ws_dir / "requirements.txt"
    if out_file.exists():
        if not messagebox.askyesno("Overwrite?", f"{out_file.name} already exists.\n\nOverwrite it?"):
            return

    try:
        r = subprocess.run(
            [str(python_exe), "-m", "pip", "freeze"],
            cwd=str(ws_dir),
            capture_output=True,
            text=True,
            timeout=120,
        )
        if r.returncode != 0:
            stderr = (r.stderr or "").strip()
            messagebox.showerror("pip freeze failed", stderr or "Unknown error running pip freeze.")
            return

        lines = [ln.strip() for ln in (r.stdout or "").splitlines() if ln.strip()]
        out_file.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
        messagebox.showinfo("Exported", f"Saved:\n{out_file}\n\nPackages: {len(lines)}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def open_in_editor(lb_p, lb_w):
    sel_p, sel_w = lb_p.curselection(), lb_w.curselection()
    if not sel_p or not sel_w:
        return
    editor = get_editor_path_or_ask()
    if editor:
        ws = BASE_DIR / lb_p.get(sel_p[0]) / "workspaces" / lb_w.get(sel_w[0])
        if ws.exists():
            subprocess.Popen([editor, str(ws)], creationflags=0x00000008)


# ==============================
# MAIN
# ==============================

def main():
    global _CFG
    _CFG = init_config()

    root = tk.Tk()
    root.title("Local Git Control Panel")
    root.geometry("780x540")

    f = tk.Frame(root, padx=10, pady=10)
    f.pack(fill=tk.BOTH, expand=True)

    tk.Label(f, text="Projects", font=("", 10, "bold")).grid(row=0, column=0, sticky="w")
    lb_p = tk.Listbox(f, height=14, width=20, exportselection=False)
    lb_p.grid(row=1, column=0, sticky="nsew", padx=(0, 5))

    tk.Label(f, text="Workspaces", font=("", 10, "bold")).grid(row=0, column=1, sticky="w")
    lb_w = tk.Listbox(f, height=14, width=24, exportselection=False)
    lb_w.grid(row=1, column=1, sticky="nsew", padx=(0, 5))

    tk.Label(f, text="Launchers", font=("", 10, "bold")).grid(row=0, column=2, sticky="w")
    lb_l = tk.Listbox(f, height=14, width=28, exportselection=False)
    lb_l.grid(row=1, column=2, sticky="nsew")

    f.grid_rowconfigure(1, weight=1)
    for i in range(3):
        f.grid_columnconfigure(i, weight=1)

    # Row 1: Basic actions
    b1 = tk.Frame(f)
    b1.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 3))
    tk.Button(b1, text="Set Root", command=lambda: set_root_folder(lb_p, lb_w, status)).pack(side=tk.LEFT, padx=2)
    tk.Button(b1, text="Refresh", command=lambda: [refresh_projects(lb_p, lb_w, status), refresh_launchers(lb_l)]).pack(side=tk.LEFT, padx=2)
    tk.Button(b1, text="Project", command=lambda: create_project(lb_p, lb_w, lb_l, status)).pack(side=tk.LEFT, padx=2)
    tk.Button(b1, text="Workspace", command=lambda: create_workspace(lb_p, lb_w, lb_l, status)).pack(side=tk.LEFT, padx=2)
    tk.Button(b1, text="Editor", command=set_editor_path).pack(side=tk.LEFT, padx=2)

    # Row 2: Import, Venv, Regen, Launcher
    b2 = tk.Frame(f)
    b2.grid(row=3, column=0, columnspan=3, sticky="ew", pady=3)
    tk.Button(b2, text="Import Folder", command=lambda: do_import(lb_p, lb_w, status), bg="#3498db", fg="white").pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Adopt Repo", command=lambda: adopt_repo_button(lb_p, lb_w, status),).pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Global Venv", command=lambda: setup_venv_btn(status)).pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Venv Info", command=show_venv_status).pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Export reqs", command=lambda: export_requirements(lb_p, lb_w)).pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Regen Files", command=lambda: regen_workspace_files(lb_p, lb_w)).pack(side=tk.LEFT, padx=2)
    tk.Button(b2, text="Gen Launcher", command=lambda: generate_launcher_selected(lb_p, lb_w, lb_l)).pack(side=tk.LEFT, padx=2)

    # Row 3: RUN
    b3 = tk.Frame(f)
    b3.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    tk.Button(
        b3,
        text="RUN LAUNCHER",
        command=lambda: run_launcher_and_exit(lb_l, root),
        bg="#27ae60",
        fg="white",
        font=("", 12, "bold"),
        height=2,
        cursor="hand2",
    ).pack(fill=tk.X, padx=40, pady=5)

    status = tk.Label(root, text="", anchor="w", padx=10, pady=5, relief=tk.SUNKEN, bg="#f0f0f0")
    status.pack(fill=tk.X, side=tk.BOTTOM)

    lb_p.bind("<<ListboxSelect>>", lambda e: on_project_select(e, lb_p, lb_w, status))
    lb_w.bind("<Double-1>", lambda e: open_in_editor(lb_p, lb_w))
    lb_l.bind("<Double-1>", lambda e: run_launcher_and_exit(lb_l, root))

    refresh_projects(lb_p, lb_w, status)
    refresh_launchers(lb_l)

    root.mainloop()


if __name__ == "__main__":
    main()
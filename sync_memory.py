# Git Sync Script for OpenClaw Workspace & Config
import os
import subprocess
import datetime
import shutil

WORKSPACE_DIR = r"C:\Users\307984\.openclaw\workspace"
CONFIG_DIR = r"C:\Users\307984\.openclaw"
REPO_URL = "https://github.com/tytymang/OC-Workspace.git"

# Files to backup from the config directory to the workspace/config folder
CONFIG_FILES_TO_BACKUP = [
    os.path.join(CONFIG_DIR, "cron", "jobs.json"),
    os.path.join(CONFIG_DIR, "identity", "device.json"),
    os.path.join(CONFIG_DIR, "devices", "paired.json"),
]

def run_git(args):
    try:
        result = subprocess.run(["git"] + args, cwd=WORKSPACE_DIR, capture_output=True, text=True, check=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")
        return None

def sync():
    os.chdir(WORKSPACE_DIR)
    
    # Ensure backup directory exists
    backup_path = os.path.join(WORKSPACE_DIR, "system_config_backup")
    if not os.path.exists(backup_path):
        os.makedirs(backup_path)

    # Copy system config files to workspace for git tracking
    print("Backing up system configuration files...")
    for src in CONFIG_FILES_TO_BACKUP:
        if os.path.exists(src):
            dest = os.path.join(backup_path, os.path.basename(src))
            shutil.copy2(src, dest)
            print(f"Copied {src} to {dest}")

    # Initialize if not already
    if not os.path.exists(os.path.join(WORKSPACE_DIR, ".git")):
        print("Initializing git...")
        run_git(["init"])
        run_git(["remote", "add", "origin", REPO_URL])
        run_git(["branch", "-M", "main"])

    # Pull latest changes
    print("Pulling latest data...")
    run_git(["pull", "origin", "main"])

    # Push current changes
    print("Syncing to GitHub...")
    run_git(["add", "."])
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_git(["commit", "-m", f"Auto-sync memory & config: {timestamp}"])
    run_git(["push", "origin", "main"])
    print("Sync complete!")

if __name__ == "__main__":
    sync()

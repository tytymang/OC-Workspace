# Git Sync Script for OpenClaw Workspace
import os
import subprocess
import datetime

WORKSPACE_DIR = r"C:\Users\307984\.openclaw\workspace"
REPO_URL = "https://github.com/tytymang/OC-Workspace.git"

def run_git(args):
    try:
        result = subprocess.run(["git"] + args, cwd=WORKSPACE_DIR, capture_output=True, text=True, check=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")
        return None

def sync():
    os.chdir(WORKSPACE_DIR)
    
    # Initialize if not already
    if not os.path.exists(os.path.join(WORKSPACE_DIR, ".git")):
        print("Initializing git...")
        run_git(["init"])
        run_git(["remote", "add", "origin", REPO_URL])
        run_git(["branch", "-M", "main"])

    # Pull latest changes
    print("Pulling latest memory...")
    run_git(["pull", "origin", "main"])

    # Push current changes
    print("Backing up memory...")
    run_git(["add", "."])
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_git(["commit", "-m", f"Auto-sync memory: {timestamp}"])
    run_git(["push", "origin", "main"])
    print("Sync complete!")

if __name__ == "__main__":
    sync()

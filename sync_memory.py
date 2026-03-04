import os
import shutil
import subprocess
import datetime

# Paths
WORKSPACE_DIR = r"C:\Users\307984\.openclaw\workspace"
CONFIG_DIR = r"C:\Users\307984\.openclaw"
REPO_URL = "https://github.com/tytymang/OC-Workspace.git"

# Backup sub-directories inside workspace
BACKUP_MAP = {
    "memory": os.path.join(WORKSPACE_DIR, "memory_backup"),
    "skills": os.path.join(WORKSPACE_DIR, "skills_backup"),
    "config": os.path.join(WORKSPACE_DIR, "system_config_backup")
}

def run_git(args):
    try:
        result = subprocess.run(["git"] + args, cwd=WORKSPACE_DIR, capture_output=True, text=True, check=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Git Error: {e.stderr}")
        return None

def organize_and_sync():
    # 1. Create directories
    for path in BACKUP_MAP.values():
        if not os.path.exists(path):
            os.makedirs(path)
            print(f"Created: {path}")

    # 2. Copy Memory (Long-term & Daily)
    print("Organizing Memory...")
    # Long-term memory
    lt_memory = os.path.join(WORKSPACE_DIR, "MEMORY.md")
    if os.path.exists(lt_memory):
        shutil.copy2(lt_memory, os.path.join(BACKUP_MAP["memory"], "MEMORY.md"))
    
    # Daily memories (if folder exists)
    daily_mem_dir = os.path.join(WORKSPACE_DIR, "memory")
    if os.path.exists(daily_mem_dir):
        dest_daily = os.path.join(BACKUP_MAP["memory"], "daily")
        if os.path.exists(dest_daily): shutil.rmtree(dest_daily)
        shutil.copytree(daily_mem_dir, dest_daily)

    # 3. Copy Skills
    print("Organizing Skills...")
    skills_dir = os.path.join(WORKSPACE_DIR, "skills")
    if os.path.exists(skills_dir):
        if os.path.exists(BACKUP_MAP["skills"]): shutil.rmtree(BACKUP_MAP["skills"])
        shutil.copytree(skills_dir, BACKUP_MAP["skills"])

    # 4. Copy System Configs (Gateway/Identity/Devices)
    print("Organizing System Configs...")
    config_targets = [
        (os.path.join(CONFIG_DIR, "cron", "jobs.json"), "jobs.json"),
        (os.path.join(CONFIG_DIR, "identity", "device.json"), "device.json"),
        (os.path.join(CONFIG_DIR, "devices", "paired.json"), "paired.json"),
    ]
    for src, name in config_targets:
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(BACKUP_MAP["config"], name))

    # 5. Git Sync
    print("Syncing to GitHub...")
    run_git(["add", "."])
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_git(["commit", "-m", f"Organized Sync: {timestamp}"])
    run_git(["push", "origin", "main"])
    print("Sync Complete!")

if __name__ == "__main__":
    organize_and_sync()

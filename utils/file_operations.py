from datetime import datetime
from pathlib import Path
import os
import shutil

def get_timestamped_filename(filepath):
    """Create a timestamped backup filename if the file exists."""
    if os.path.exists(filepath):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path_obj = Path(filepath)
        backup_name = f"{path_obj.stem}_{timestamp}{path_obj.suffix}"
        return backup_name
    return filepath

def backup_existing_files(cache_file, bilingual_csv, audit_json, log_file):
    """Backup existing output files with timestamps."""
    files_backed_up = []

    for filepath in [cache_file, bilingual_csv, audit_json, log_file]:
        if os.path.exists(filepath):
            backup_name = get_timestamped_filename(filepath)
            shutil.move(filepath, backup_name)
            files_backed_up.append(f"{filepath} -> {backup_name}")

    if files_backed_up:
        print("Backed up existing files:")
        for backup in files_backed_up:
            print(f"  {backup}")
        print()

    return files_backed_up

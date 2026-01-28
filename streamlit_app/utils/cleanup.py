"""
Utility functions for cleaning up temporary files
"""

import os
import shutil
import time
from pathlib import Path
from datetime import datetime, timedelta


def cleanup_old_temp_files(temp_dir, hours=24):
    """
    Delete temporary subdirectories older than specified hours
    
    Args:
        temp_dir: Path to the temporary directory
        hours: Delete files/folders older than this many hours (default: 24)
    
    Returns:
        tuple: (deleted_count, error_count)
    """
    temp_path = Path(temp_dir)
    
    if not temp_path.exists():
        return 0, 0
    
    cutoff_time = time.time() - (hours * 3600)
    deleted_count = 0
    error_count = 0
    
    try:
        for item in temp_path.iterdir():
            if item.is_dir():
                try:
                    # Get the modification time of the directory
                    item_mtime = item.stat().st_mtime
                    
                    if item_mtime < cutoff_time:
                        shutil.rmtree(item, ignore_errors=True)
                        deleted_count += 1
                        print(f"Deleted old temp directory: {item.name}")
                except Exception as e:
                    error_count += 1
                    print(f"Error deleting {item.name}: {e}")
    
    except Exception as e:
        print(f"Error accessing temp directory: {e}")
        return deleted_count, error_count + 1
    
    return deleted_count, error_count


def cleanup_session_files(temp_dir, session_id):
    """
    Delete files for a specific session
    
    Args:
        temp_dir: Path to the temporary directory
        session_id: Session ID to clean up
    
    Returns:
        bool: True if successful, False otherwise
    """
    session_path = Path(temp_dir) / session_id
    
    if session_path.exists() and session_path.is_dir():
        try:
            shutil.rmtree(session_path, ignore_errors=True)
            return True
        except Exception as e:
            print(f"Error deleting session directory {session_id}: {e}")
            return False
    
    return True  # Already deleted or doesn't exist


def get_temp_dir_size(temp_dir):
    """
    Calculate total size of temporary directory in MB
    
    Args:
        temp_dir: Path to the temporary directory
    
    Returns:
        float: Size in megabytes
    """
    temp_path = Path(temp_dir)
    
    if not temp_path.exists():
        return 0.0
    
    total_size = 0
    try:
        for item in temp_path.rglob('*'):
            if item.is_file():
                total_size += item.stat().st_size
    except Exception as e:
        print(f"Error calculating directory size: {e}")
    
    return total_size / (1024 * 1024)  # Convert to MB


def get_temp_dir_info(temp_dir):
    """
    Get information about temporary directory
    
    Args:
        temp_dir: Path to the temporary directory
    
    Returns:
        dict: Information about the temporary directory
    """
    temp_path = Path(temp_dir)
    
    if not temp_path.exists():
        return {
            'exists': False,
            'folder_count': 0,
            'total_size_mb': 0.0,
            'oldest_folder': None
        }
    
    folders = [f for f in temp_path.iterdir() if f.is_dir()]
    oldest_folder = None
    oldest_time = None
    
    for folder in folders:
        try:
            mtime = folder.stat().st_mtime
            if oldest_time is None or mtime < oldest_time:
                oldest_time = mtime
                oldest_folder = folder.name
        except:
            pass
    
    return {
        'exists': True,
        'folder_count': len(folders),
        'total_size_mb': get_temp_dir_size(temp_dir),
        'oldest_folder': oldest_folder,
        'oldest_time': datetime.fromtimestamp(oldest_time) if oldest_time else None
    }

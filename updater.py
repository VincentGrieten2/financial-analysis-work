import os
import sys
import json
import requests
from version import VERSION, GITHUB_REPO

def check_for_updates():
    """
    Check GitHub releases for newer versions.
    Returns: (bool has_update, str latest_version, str download_url)
    """
    try:
        # Get latest release from GitHub
        response = requests.get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest")
        if response.status_code != 200:
            return False, VERSION, None
        
        release_data = response.json()
        latest_version = release_data['tag_name'].replace('v', '')
        
        # Compare versions (simple string comparison, assuming semantic versioning)
        if latest_version > VERSION:
            # Get the .zip asset download URL
            for asset in release_data['assets']:
                if asset['name'].endswith('.zip'):
                    return True, latest_version, asset['browser_download_url']
        
        return False, VERSION, None
    
    except Exception as e:
        print(f"Error checking for updates: {str(e)}")
        return False, VERSION, None

def get_update_status():
    """Returns a user-friendly update status message"""
    has_update, latest_version, _ = check_for_updates()
    if has_update:
        return f"Update available! Current version: {VERSION}, Latest version: {latest_version}"
    return f"You are running the latest version ({VERSION})" 
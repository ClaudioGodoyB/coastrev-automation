#!/usr/bin/env python3
"""
Path Configuration System for CoastRev VM Migration
Centralizes Windows->Linux path conversion and Google Drive integration
"""
import os
from pathlib import Path

class PathManager:
    """Manages path conversion between Windows and Linux environments"""
    
    def __init__(self):
        # Base directories for VM environment
        self.VM_BASE = "/home/user/coastrev"
        
        # Windows paths to Linux path mappings
        self.WINDOWS_TO_LINUX = {
            r"/home/user/coastrev/data/downloads": f"{self.VM_BASE}/data/downloads",
            r"/home/user/coastrev/data/extracts": f"{self.VM_BASE}/data/extracts", 
            r"/home/user/coastrev/data/daily_details": f"{self.VM_BASE}/data/daily_details",
            r"/home/user/coastrev/templates/excel": f"{self.VM_BASE}/templates/excel",
            r"/home/user/coastrev/data": f"{self.VM_BASE}/data",
            r"/home/user/coastrev": f"{self.VM_BASE}",
        }
        
        # Google Drive folder mappings (from drive_folder_ids.json)
        self.DRIVE_FOLDERS = {
            "root": "17MjrR75Tuud0dmIVYzqYkg15b9lij9OI",
            "data_inputs": "1GMXCQocVPOR_qHC3YZNpJeXk4iKdjHN4", 
            "downloads_staging": "1ez3qGyUzDEm2wMe_TkckH7YFbhk40IVc",
            "email_attachments": "1jqAM9q_XIYHBe-Of1FB-c8zXb38RQj47",
            "email_attachments_properties": {
                "BWOF": "1cZM06utHuKWDgmn6GL2YShbSxIBlMPU2",
                "BWSM": "1LgeDcjOFAas34oT-uu2UTUc4DXjquudu", 
                "CHA": "1FTQ_RYGCb3I2LxHe4jZtt8G4FBMRhPIq",
                "CSI": "1QYRyY1XShJrGXr8365-yQ3LxG1oJYB0V",
                "HATH": "1sot24F_WTBe0EYd1rFcgvNQ1o0A-uQwu",
                "HBSB": "1JqWlul3riAqU3kuuPV4stQUMvpRxODb-",
                "ISLO": "1LM9YL4aPZYZdtNyKlNMNOnam7ljUgrPi", 
                "LOL": "1QsyyIW9kR5Yu6w99hM6YFFTYem8taHUw",
                "LSMB": "1Fe-DEjaK7MgUeFu74cqXD5WtjbuzBXJ1",
                "MPMS": "1RiR9gay9d3KBjs7T5fPeNJvPPtRGToz1",
                "MSI": "15F29IB7anvZ5y83fdip1KxRZZCxPZ9rg",
                "RRM": "1FqYK3jenguB3VfBJKU1r20BQZYwpNDsz",
                "TLMB": "1VRAnwqUBun5BO-oWoirLD0HpsnNbbgK-", 
                "TRH": "1lRerON12GbZGc5nbi-ZZGqCaqXJ2wJtV"
            }
        }
        
        # Property list for validation
        self.PROPERTIES = ["BWOF", "BWSM", "CHA", "CSI", "HATH", "HBSB", 
                          "ISLO", "LOL", "LSMB", "MPMS", "MSI", "RRM", "TLMB", "TRH"]

    def convert_windows_path(self, windows_path):
        """Convert Windows path to Linux equivalent"""
        # Normalize path separators
        normalized_path = windows_path.replace('\\', '/')
        
        # Find the longest matching Windows path
        best_match = ""
        best_replacement = ""
        
        for win_path, linux_path in self.WINDOWS_TO_LINUX.items():
            win_normalized = win_path.replace('\\', '/')
            if normalized_path.startswith(win_normalized) and len(win_normalized) > len(best_match):
                best_match = win_normalized
                best_replacement = linux_path
        
        if best_match:
            # Replace the matched portion
            relative_part = normalized_path[len(best_match):].lstrip('/')
            if relative_part:
                return os.path.join(best_replacement, relative_part).replace('\\', '/')
            else:
                return best_replacement
        else:
            # No match found - return as-is with warning
            print(f"WARNING: No path mapping found for: {windows_path}")
            return windows_path

    def get_drive_folder_id(self, folder_type, property_code=None):
        """Get Google Drive folder ID for specific folder type"""
        if folder_type in self.DRIVE_FOLDERS:
            folder_data = self.DRIVE_FOLDERS[folder_type]
            if isinstance(folder_data, dict) and property_code:
                return folder_data.get(property_code)
            else:
                return folder_data
        return None

    def get_property_drive_folder(self, property_code):
        """Get email attachments folder ID for specific property"""
        return self.get_drive_folder_id("email_attachments_properties", property_code)

    def ensure_local_directory(self, path):
        """Ensure local directory exists"""
        Path(path).mkdir(parents=True, exist_ok=True)
        return path

    def get_dated_folder_path(self, base_path, date_str=None):
        """Create dated folder path (e.g., Extract 2025-09-10)"""
        if date_str is None:
            from datetime import datetime
            date_str = datetime.today().strftime("%Y-%m-%d")
        
        if "Extract" in base_path or "extract" in base_path.lower():
            folder_name = f"Extract {date_str}"
        elif "Detail" in base_path or "detail" in base_path.lower():  
            folder_name = f"Daily Detail {date_str}"
        else:
            folder_name = date_str
            
        return os.path.join(base_path, folder_name)

# Global instance for easy importing
path_manager = PathManager()

# Convenience functions
def convert_path(windows_path):
    """Quick path conversion function"""
    return path_manager.convert_windows_path(windows_path)

def get_drive_folder(folder_type, property_code=None):
    """Quick drive folder ID lookup"""
    return path_manager.get_drive_folder_id(folder_type, property_code)

def ensure_dir(path):
    """Quick directory creation"""
    return path_manager.ensure_local_directory(path)

# Test the configuration
if __name__ == "__main__":
    # Test path conversions
    test_paths = [
        r"/home/user/coastrev/data/downloads",
        r"/home/user/coastrev/data/extracts",
        r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details\Daily Detail 2025-09-10\Misc',
        r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\BWOF.xlsx'
    ]
    
    print("Path Conversion Tests:")
    print("="*50)
    for path in test_paths:
        converted = convert_path(path)
        print(f"OLD: {path}")
        print(f"NEW: {converted}")
        print()
    
    print("Google Drive Folder Tests:")  
    print("="*50)
    print(f"Root folder: {get_drive_folder('root')}")
    print(f"BWOF email folder: {get_drive_folder('email_attachments_properties', 'BWOF')}")
    print(f"Downloads staging: {get_drive_folder('downloads_staging')}")
#!/usr/bin/env python3
"""
Bulk Path Update Script for CoastRev VM Migration
Automatically replaces Windows paths with Linux equivalents across the entire codebase
"""
import os
import re
import glob
from pathlib import Path
import shutil
from datetime import datetime

# Import our path configuration
import sys
sys.path.append('../config')
from paths import path_manager

class BulkPathUpdater:
    """Updates all Windows paths in Python files to Linux equivalents"""
    
    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        self.backup_dir = self.root_dir / "migration_backups" / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        self.updated_files = []
        self.failed_files = []
        
    def create_backup(self, file_path):
        """Create backup of file before modification"""
        try:
            relative_path = file_path.relative_to(self.root_dir)
            backup_path = self.backup_dir / relative_path
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(file_path, backup_path)
            return True
        except Exception as e:
            print(f"Failed to backup {file_path}: {e}")
            return False
    
    def update_file_paths(self, file_path):
        """Update Windows paths in a single file"""
        try:
            # Read file content
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            original_content = content
            
            # Find Windows paths using simple search and replace
            # Look for the specific patterns we know exist
            windows_paths = [
                r'C:\Users\johnj\Downloads',
                r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Extracts', 
                r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details',
                r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates',
                r'C:\Users\johnj\Desktop\CoastRev\Reporting',
                r'C:\Users\johnj\Desktop\CoastRev'
            ]
            
            # Track if any changes were made
            changes_made = False
            
            # Process each Windows path (longest paths first to avoid partial replacements)
            for win_path in sorted(windows_paths, key=len, reverse=True):
                linux_path = path_manager.convert_windows_path(win_path)
                
                if linux_path != win_path:
                    # Find all variations of this path in the content
                    variations = [
                        f'r"{win_path}"',
                        f"r'{win_path}'", 
                        f'"{win_path}"',
                        f"'{win_path}'",
                        f'rf"{win_path}"',
                        f"rf'{win_path}'"
                    ]
                    
                    for old_var in variations:
                        if old_var in content:
                            # Determine the new format
                            if old_var.startswith('rf'):
                                new_var = f'rf"{linux_path}"'
                            elif old_var.startswith('r'):
                                new_var = f'r"{linux_path}"'  
                            else:
                                new_var = f'"{linux_path}"'
                                
                            content = content.replace(old_var, new_var)
                            changes_made = True
                            print(f"  {win_path} -> {linux_path}")
            
            # Write back only if changes were made
            if changes_made and content != original_content:
                # Create backup first
                if self.create_backup(file_path):
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    return True
                else:
                    print(f"Skipping {file_path} - backup failed")
                    return False
            
            return not changes_made  # Return True if no changes needed
            
        except Exception as e:
            print(f"Error updating {file_path}: {e}")
            return False
    
    def find_python_files(self):
        """Find all Python files in the codebase"""
        python_files = []
        
        # Get all .py files recursively
        for file_path in self.root_dir.rglob("*.py"):
            # Skip backup directories and __pycache__
            if "migration_backups" in str(file_path) or "__pycache__" in str(file_path):
                continue
            python_files.append(file_path)
            
        return python_files
    
    def update_all_files(self):
        """Update all Python files in the codebase"""
        print(f"Starting bulk path update for: {self.root_dir}")
        print(f"Backup directory: {self.backup_dir}")
        print("="*60)
        
        python_files = self.find_python_files()
        print(f"Found {len(python_files)} Python files to process")
        print()
        
        # Create backup directory
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        
        for i, file_path in enumerate(python_files, 1):
            print(f"[{i:3d}/{len(python_files)}] Processing: {file_path.name}")
            
            success = self.update_file_paths(file_path)
            
            if success:
                self.updated_files.append(file_path)
            else:
                self.failed_files.append(file_path)
        
        # Print summary
        print()
        print("="*60)
        print("MIGRATION SUMMARY")
        print("="*60)
        print(f"Total files processed: {len(python_files)}")
        print(f"Successfully updated: {len(self.updated_files)}")
        print(f"Failed updates: {len(self.failed_files)}")
        
        if self.failed_files:
            print("\nFailed files:")
            for failed_file in self.failed_files:
                print(f"  - {failed_file}")
        
        print(f"\nBackups stored in: {self.backup_dir}")
        print("Path migration completed!")

def main():
    """Main execution function"""
    # Get the root directory of the codebase
    script_dir = Path(__file__).parent
    root_dir = script_dir.parent  # Go up one level from migration_scripts/
    
    print("CoastRev VM Migration - Bulk Path Updater")
    print("="*50)
    
    # Create updater and run
    updater = BulkPathUpdater(root_dir)
    updater.update_all_files()

if __name__ == "__main__":
    main()
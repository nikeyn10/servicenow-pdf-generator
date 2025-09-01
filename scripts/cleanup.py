#!/usr/bin/env python3
"""
ServiceNow PDF Generator - Cleanup Utility
Helps maintain the project by cleaning up temporary files and old downloads.
"""

import os
import shutil
import argparse
from pathlib import Path
from datetime import datetime, timedelta

def cleanup_system_files(project_root):
    """Remove system files like .DS_Store"""
    print("🧹 Removing system files...")
    for root, dirs, files in os.walk(project_root):
        for file in files:
            if file == '.DS_Store':
                file_path = os.path.join(root, file)
                os.remove(file_path)
                print(f"  Removed: {file_path}")

def cleanup_temp_files(project_root):
    """Remove temporary and debug files"""
    print("🗑️  Removing temporary files...")
    temp_patterns = [
        'debug_*.py',
        'test_*.py',
        '*.tmp',
        'debug_*.txt',
        '*_analysis.txt',
        'analyze_*.py',
        'check_*.py',
        'validate_*.py'
    ]
    
    # Note: This is a placeholder - would need glob pattern matching
    # For now, just list what could be cleaned
    temp_files = [
        'debug_pagination.py',
        'test_date_columns.py',
        'debug_june_tickets.txt',
        'june_attachments_analysis.txt',
        'analyze_july_attachments.py'
    ]
    
    for temp_file in temp_files:
        file_path = os.path.join(project_root, temp_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"  Removed: {temp_file}")

def archive_old_downloads(project_root, months_to_keep=2):
    """Archive downloads older than specified months"""
    print(f"📦 Archiving downloads older than {months_to_keep} months...")
    downloads_dir = os.path.join(project_root, 'data', 'downloads')
    archive_dir = os.path.join(project_root, 'archive', 'downloads')
    
    if not os.path.exists(downloads_dir):
        print("  No downloads directory found")
        return
    
    os.makedirs(archive_dir, exist_ok=True)
    
    cutoff_date = datetime.now() - timedelta(days=30 * months_to_keep)
    
    for folder in os.listdir(downloads_dir):
        folder_path = os.path.join(downloads_dir, folder)
        if os.path.isdir(folder_path):
            try:
                # Parse folder name as YYYY-MM
                folder_date = datetime.strptime(folder, '%Y-%m')
                if folder_date < cutoff_date:
                    archive_path = os.path.join(archive_dir, folder)
                    shutil.move(folder_path, archive_path)
                    print(f"  Archived: {folder} -> archive/downloads/")
            except ValueError:
                print(f"  Skipping folder with invalid date format: {folder}")

def get_folder_sizes(project_root):
    """Display folder sizes for manual review"""
    print("📊 Folder sizes:")
    
    folders_to_check = [
        'data/downloads',
        'output/merged',
        'archive'
    ]
    
    for folder in folders_to_check:
        folder_path = os.path.join(project_root, folder)
        if os.path.exists(folder_path):
            size = get_directory_size(folder_path)
            print(f"  {folder}: {format_size(size)}")

def get_directory_size(path):
    """Calculate total size of directory"""
    total = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            try:
                total += os.path.getsize(filepath)
            except OSError:
                pass
    return total

def format_size(size_bytes):
    """Format byte size in human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

def main():
    parser = argparse.ArgumentParser(description='ServiceNow PDF Generator Cleanup Utility')
    parser.add_argument('--dry-run', action='store_true', help='Show what would be cleaned without doing it')
    parser.add_argument('--archive-old', action='store_true', help='Archive old download folders')
    parser.add_argument('--months-to-keep', type=int, default=2, help='Months of downloads to keep (default: 2)')
    
    args = parser.parse_args()
    
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(f"🏠 Project root: {project_root}")
    
    if args.dry_run:
        print("🔍 DRY RUN - No files will be modified")
    
    print("\n" + "="*50)
    
    get_folder_sizes(project_root)
    
    if not args.dry_run:
        print("\n" + "="*50)
        cleanup_system_files(project_root)
        
        if args.archive_old:
            archive_old_downloads(project_root, args.months_to_keep)
    
    print("\n✅ Cleanup complete!")

if __name__ == '__main__':
    main()

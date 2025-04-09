import os
import sys
import shutil
import re
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("reorganization.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

def reorganize_output_folder(output_dir):
    """
    Reorganize PDF files in the output directory into categorized subfolders:
    - Country Associated Documents: Files with COUNTRY_ prefix or project ID (P######)
    - Unknown Countries: Files with UNKNOWN_ prefix
    - Failed Conversions and Renaming: All other files
    
    Args:
        output_dir: Path to the output directory containing PDF files
    """
    if not os.path.exists(output_dir):
        logging.error(f"Output directory does not exist: {output_dir}")
        return
    
    print("\n" + "="*80)
    print(" REORGANIZING OUTPUT FILES INTO CATEGORIES ".center(80, "="))
    print("="*80 + "\n")
    
    # Create the three category folders
    country_folder = os.path.join(output_dir, "Country Associated Documents")
    unknown_folder = os.path.join(output_dir, "Unknown Countries")
    failed_folder = os.path.join(output_dir, "Failed Conversions and Renaming")
    
    # Create folders if they don't exist
    os.makedirs(country_folder, exist_ok=True)
    os.makedirs(unknown_folder, exist_ok=True)
    os.makedirs(failed_folder, exist_ok=True)
    
    # Track counts for reporting
    country_count = 0
    unknown_count = 0
    failed_count = 0
    
    # Regex for project ID pattern (P followed by 6 digits)
    pid_pattern = re.compile(r'^P\d{6}')
    
    # Process each PDF file
    pdf_files = []
    for root, _, files in os.walk(output_dir):
        # Skip the category folders themselves
        if root in [country_folder, unknown_folder, failed_folder]:
            continue
            
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    # Use a progress counter
    total_files = len(pdf_files)
    processed = 0
    
    print(f"Found {total_files} PDF files to organize")
    
    for pdf_file in pdf_files:
        # Get just the filename
        filename = os.path.basename(pdf_file)
        
        # Determine the target folder
        if filename.startswith("COUNTRY_") or pid_pattern.match(filename):
            target_folder = country_folder
            country_count += 1
        elif filename.startswith("UNKNOWN_"):
            target_folder = unknown_folder
            unknown_count += 1
        else:
            target_folder = failed_folder
            failed_count += 1
        
        # Create destination path ensuring no conflicts
        dest_path = os.path.join(target_folder, filename)
        
        # Handle potential filename conflicts
        if os.path.exists(dest_path):
            base, ext = os.path.splitext(filename)
            counter = 1
            while True:
                new_filename = f"{base}_{counter:02d}{ext}"
                dest_path = os.path.join(target_folder, new_filename)
                if not os.path.exists(dest_path):
                    break
                counter += 1
        
        # Move the file
        try:
            shutil.move(pdf_file, dest_path)
            processed += 1
            
            # Show progress every 50 files
            if processed % 50 == 0 or processed == total_files:
                print(f"Processed {processed}/{total_files} files ({processed/total_files*100:.1f}%)")
                
        except Exception as e:
            logging.error(f"Error moving file {pdf_file}: {str(e)}")
    
    # Print summary
    print("\n" + "-"*80)
    print(" FILE ORGANIZATION SUMMARY ".center(80, "-"))
    print("-"*80)
    print(f"Total files processed: {processed}")
    print(f"Files in 'Country Associated Documents': {country_count}")
    print(f"Files in 'Unknown Countries': {unknown_count}")
    print(f"Files in 'Failed Conversions and Renaming': {failed_count}")
    print("-"*80 + "\n")
    
    return country_count, unknown_count, failed_count

if __name__ == "__main__":
    # If called directly, check for output directory argument
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]
        reorganize_output_folder(output_dir)
    else:
        print("Please provide the output directory as an argument.")
        print("Usage: python reorganize_output.py <output_directory>")
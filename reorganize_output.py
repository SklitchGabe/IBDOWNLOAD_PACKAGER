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

def reorganize_output_folder(output_dir, document_type=None):
    """
    Reorganize PDF files in the output directory into categorized subfolders:
    - Country Associated Documents: Files with COUNTRY_ prefix or project ID (P######)
    - Unknown Countries: Files with UNKNOWN_ prefix
    - Failed Conversions and Renaming: All other files
    
    Args:
        output_dir: Path to the output directory containing PDF files
        document_type: Type of document (e.g., icrr, aidememoire, pad) to add to filenames
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
        
        # Add document type information if provided
        if document_type:
            # Split the filename into base and extension
            base, ext = os.path.splitext(filename)
            
            # Check if the base already has a numeric suffix (like _01)
            if re.search(r'_\d+$', base):
                # Insert the document type before the numeric suffix
                match = re.search(r'(_\d+)$', base)
                suffix = match.group(1)
                base = base[:-len(suffix)]
                new_filename = f"{base}_{document_type}{suffix}{ext}"
            else:
                # Just append the document type to the base
                new_filename = f"{base}_{document_type}{ext}"
                
            # Create a temporary path for the renamed file
            temp_path = os.path.join(os.path.dirname(pdf_file), new_filename)
            
            # Rename the file
            try:
                os.rename(pdf_file, temp_path)
                # Update pdf_file to point to the renamed file
                pdf_file = temp_path
                filename = new_filename
                logging.info(f"Added document type to filename: {filename}")
            except Exception as e:
                logging.error(f"Error adding document type to filename: {filename} - {str(e)}")
        
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
    if document_type:
        print(f"Added document type '{document_type}' to all filenames")
    print("-"*80 + "\n")
    
    # Further organize the Country Associated Documents folder by country
    if country_count > 0:
        organize_by_country(country_folder)
    
    return country_count, unknown_count, failed_count

def organize_by_country(country_folder):
    """
    Further organize the "Country Associated Documents" folder into country-specific subfolders.
    Handles countries with multiple words and special characters.
    Consolidates similar country names (with spaces vs underscores).
    
    Args:
        country_folder: Path to the "Country Associated Documents" folder
    """
    if not os.path.exists(country_folder):
        logging.error(f"Country Associated Documents folder does not exist: {country_folder}")
        return
    
    print("\n" + "="*80)
    print(" ORGANIZING DOCUMENTS BY COUNTRY ".center(80, "="))
    print("="*80 + "\n")
    
    # Get all PDF files
    pdf_files = []
    for file in os.listdir(country_folder):
        file_path = os.path.join(country_folder, file)
        if file.lower().endswith('.pdf') and os.path.isfile(file_path):
            pdf_files.append(file_path)
    
    if not pdf_files:
        print("No PDF files found in Country Associated Documents folder")
        return
    
    print(f"Found {len(pdf_files)} PDF files to organize by country")
    
    # Function to standardize country names
    def standardize_country_name(country_name):
        """Standardize country name format for consistent folder naming"""
        if not country_name:
            return country_name
            
        # 1. Replace underscores with spaces for standardization
        standardized = country_name.replace('_', ' ')
        # 2. Normalize spaces (remove extra spaces)
        standardized = ' '.join(standardized.split())
        # 3. Capitalize each word for consistency
        standardized = standardized.title()
        
        return standardized
    
    # Function to create a filesystem-safe folder name
    def safe_folder_name(country_name):
        """Create a filesystem-safe folder name from a country name"""
        # Replace spaces with underscores and remove problematic characters
        return re.sub(r'[\\/*?:"<>|,]', '_', country_name.replace(' ', '_'))
    
    # Track countries and their document counts
    # We'll use a dict where key=standardized name, value=dict of original names and files
    country_docs = {}
    
    # Define improved regex patterns that handle numeric suffixes and special characters
    # For Project ID filenames with potential document type
    pid_pattern = re.compile(r'^P\d{6}_(.+?)_(EN|NON|UNK|OCR)(?:_[^_]+)?(?:_\d+)?\.pdf$')
    
    # For Country prefix filenames with potential document type
    country_prefix_pattern = re.compile(r'^COUNTRY_(.+?)_(EN|NON|UNK|OCR)(?:_[^_]+)?(?:_\d+)?\.pdf$')
    
    # First pass - identify countries and count documents
    for pdf_file in pdf_files:
        filename = os.path.basename(pdf_file)
        country = None
        
        # Extract country from filename
        if filename.startswith('P'):
            match = pid_pattern.search(filename)
            if match:
                country = match.group(1)
        elif filename.startswith('COUNTRY_'):
            match = country_prefix_pattern.search(filename)
            if match:
                country = match.group(1)
        
        if country:
            # Standardize the country name
            standard_country = standardize_country_name(country)
            
            # Store with standardized formatting 
            if standard_country not in country_docs:
                country_docs[standard_country] = {
                    'original_names': set(),  # Track original variants
                    'files': []  # Track files for this country
                }
            
            country_docs[standard_country]['original_names'].add(country)
            country_docs[standard_country]['files'].append(pdf_file)
    
    if not country_docs:
        print("Could not identify countries from filenames")
        return
    
    print(f"Identified {len(country_docs)} countries with documents")
    
    # Show any countries that had multiple name variants
    for country, data in country_docs.items():
        if len(data['original_names']) > 1:
            print(f"Consolidated country variants: {data['original_names']} → '{country}'")
    
    # Create country folders and move files
    processed = 0
    for country, data in country_docs.items():
        # Create a safe folder name
        folder_name = safe_folder_name(country)
        country_subfolder = os.path.join(country_folder, folder_name)
        
        # Create country subfolder
        os.makedirs(country_subfolder, exist_ok=True)
        
        # Move files to country subfolder
        for pdf_file in data['files']:
            filename = os.path.basename(pdf_file)
            dest_path = os.path.join(country_subfolder, filename)
            
            # Handle potential filename conflicts
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                counter = 1
                while True:
                    new_filename = f"{base}_{counter:02d}{ext}"
                    dest_path = os.path.join(country_subfolder, new_filename)
                    if not os.path.exists(dest_path):
                        break
                    counter += 1
            
            # Move the file
            try:
                shutil.move(pdf_file, dest_path)
                processed += 1
            except Exception as e:
                logging.error(f"Error moving file to country folder: {pdf_file} - {str(e)}")
        
        print(f"Created '{country}' folder with {len(data['files'])} documents")
    
    print(f"\nSuccessfully organized {processed} documents into {len(country_docs)} country folders")
    
    # Now organize by document type within each country folder
    organize_by_document_type(country_folder)

def organize_by_document_type(country_folder):
    """
    Further organize files within each country folder into document type subfolders.
    
    Args:
        country_folder: Path to the "Country Associated Documents" folder
    """
    if not os.path.exists(country_folder):
        logging.error(f"Country folder does not exist: {country_folder}")
        return
    
    print("\n" + "="*80)
    print(" ORGANIZING COUNTRY DOCUMENTS BY TYPE ".center(80, "="))
    print("="*80 + "\n")
    
    # Get all country subfolders
    country_subfolders = []
    for item in os.listdir(country_folder):
        item_path = os.path.join(country_folder, item)
        if os.path.isdir(item_path):
            country_subfolders.append(item_path)
    
    if not country_subfolders:
        print("No country subfolders found to organize")
        return
    
    print(f"Found {len(country_subfolders)} country folders to organize by document type")
    
    # Track document types and counts
    doc_type_pattern = re.compile(r'.*?_([a-z]+)(?:_\d+)?\.pdf$', re.IGNORECASE)
    total_organized = 0
    country_organized = 0
    
    # Process each country folder
    for country_subfolder in country_subfolders:
        country_name = os.path.basename(country_subfolder)
        
        # Get all PDF files in this country folder
        pdf_files = []
        for file in os.listdir(country_subfolder):
            file_path = os.path.join(country_subfolder, file)
            if file.lower().endswith('.pdf') and os.path.isfile(file_path):
                pdf_files.append(file_path)
        
        if not pdf_files:
            continue
        
        # Track document types for this country
        country_doc_types = {}
        
        # First pass - identify document types
        for pdf_file in pdf_files:
            filename = os.path.basename(pdf_file)
            
            # Try to extract document type from filename
            match = doc_type_pattern.search(filename.lower())
            if match:
                doc_type = match.group(1)
                # Don't treat language markers as document types
                if doc_type not in ['en', 'non', 'unk', 'ocr']:
                    if doc_type not in country_doc_types:
                        country_doc_types[doc_type] = []
                    country_doc_types[doc_type].append(pdf_file)
        
        # Only create doc type folders if we found multiple document types
        if len(country_doc_types) > 1:
            print(f"Organizing '{country_name}' into {len(country_doc_types)} document types")
            
            # Create document type folders and move files
            country_file_count = 0
            for doc_type, files in country_doc_types.items():
                # Create document type subfolder with proper casing
                doc_type_folder = os.path.join(country_subfolder, doc_type.upper())
                os.makedirs(doc_type_folder, exist_ok=True)
                
                # Move files to document type subfolder
                for pdf_file in files:
                    filename = os.path.basename(pdf_file)
                    dest_path = os.path.join(doc_type_folder, filename)
                    
                    # Handle potential filename conflicts
                    if os.path.exists(dest_path):
                        base, ext = os.path.splitext(filename)
                        counter = 1
                        while True:
                            new_filename = f"{base}_{counter:02d}{ext}"
                            dest_path = os.path.join(doc_type_folder, new_filename)
                            if not os.path.exists(dest_path):
                                break
                            counter += 1
                    
                    # Move the file
                    try:
                        shutil.move(pdf_file, dest_path)
                        country_file_count += 1
                    except Exception as e:
                        logging.error(f"Error moving file to document type folder: {pdf_file} - {str(e)}")
                
                print(f"  - Created '{doc_type.upper()}' folder with {len(files)} documents")
            
            total_organized += country_file_count
            country_organized += 1
        else:
            print(f"Skipping '{country_name}' - only one document type found")
    
    print(f"\nSuccessfully organized {total_organized} documents into document type folders")
    print(f"Created document type folders in {country_organized} of {len(country_subfolders)} countries")

if __name__ == "__main__":
    # If called directly, check for output directory argument
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]
        document_type = sys.argv[2] if len(sys.argv) > 2 else None
        reorganize_output_folder(output_dir, document_type)
    else:
        print("Please provide the output directory and document type as arguments.")
        print("Usage: python reorganize_output.py <output_directory> <document_type>")
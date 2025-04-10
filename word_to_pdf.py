import os
import time
import concurrent.futures
from pathlib import Path
import platform
import subprocess
from tqdm import tqdm
import sys
import shutil
import logging
import psutil
import argparse
import re
import PyPDF2
from langdetect import detect, LangDetectException
import pandas as pd

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_conversion.log"),
        # Only log warnings and errors to console to reduce clutter
        logging.StreamHandler(sys.stdout)
    ]
)

# Set the console handler to only show warnings and errors
for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.StreamHandler) and handler.stream == sys.stdout:
        handler.setLevel(logging.WARNING)

def convert_with_word(input_file, output_file=None, retries=2):
    """Convert doc/docx to PDF using Microsoft Word (Windows only)"""
    if output_file is None:
        output_file = str(Path(input_file).with_suffix('.pdf'))
    
    # Only import win32com if we're using this function
    import win32com.client
    import pythoncom
    import time
    
    # Initialize COM in this thread
    pythoncom.CoInitialize()
    
    for attempt in range(retries + 1):
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # Don't show alerts
            
            # Set these additional properties for corporate environments
            word.Options.CheckGrammarAsYouType = False
            word.Options.CheckSpellingAsYouType = False
            
            # For OneDrive files, use a more robust approach
            if "OneDrive" in input_file:
                # Try different opening methods in case of issues
                try:
                    # Method 1: Open with ReadOnly flag to avoid lock issues
                    doc = word.Documents.Open(
                        os.path.abspath(input_file), 
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False
                    )
                    
                    # Try export method instead of SaveAs for OneDrive files
                    doc.ExportAsFixedFormat(
                        OutputFileName=os.path.abspath(output_file),
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,    # wdExportOptimizeForPrint
                        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
                        DocStructureTags=True
                    )
                    doc.Close(SaveChanges=False)
                    
                except Exception as e:
                    # If the first method fails, try a different approach
                    print(f"  First OneDrive method failed: {str(e)}")
                    print("  Trying alternative method...")
                    
                    # Force close any open documents
                    try:
                        for doc in word.Documents:
                            doc.Close(SaveChanges=False)
                    except:
                        pass
                    
                    # Method 2: Copy the file to temp directory first
                    import tempfile
                    import shutil
                    
                    temp_dir = tempfile.gettempdir()
                    temp_file = os.path.join(temp_dir, f"temp_{os.path.basename(input_file)}")
                    
                    try:
                        # Copy to temp location
                        shutil.copy2(input_file, temp_file)
                        
                        # Try with the temp file
                        doc = word.Documents.Open(temp_file)
                        doc.SaveAs(os.path.abspath(output_file), FileFormat=17)
                        doc.Close()
                        
                        # Clean up temp file
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    except Exception as temp_error:
                        raise Exception(f"Both OneDrive methods failed: {str(temp_error)}")
            else:
                # Standard approach for non-OneDrive files
                doc = word.Documents.Open(os.path.abspath(input_file))
                doc.SaveAs(os.path.abspath(output_file), FileFormat=17)  # 17 is PDF format
                doc.Close(SaveChanges=False)
                
            return output_file
            
        except Exception as e:
            if attempt < retries:
                print(f"  Attempt {attempt+1} failed for {os.path.basename(input_file)}: {str(e)}")
                # Wait before retrying
                time.sleep(3)  # Increased wait time for corporate environments
                
                # Force close any hanging Word instances before retrying
                try:
                    # First try to close Word gracefully if we still have a reference
                    if word:
                        try:
                            word.Quit()
                        except:
                            pass
                    
                    # Then use taskkill as a last resort
                    import subprocess
                    subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                                  stdout=subprocess.DEVNULL, 
                                  stderr=subprocess.DEVNULL)
                    time.sleep(2)  # Give system more time to close Word
                except:
                    pass
            else:
                # All retries exhausted
                raise Exception(f"MS Word conversion failed after {retries+1} attempts: {str(e)}")
        finally:
            # Clean up COM resources
            if word:
                try:
                    word.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
    
    # This should not be reached, but just in case
    raise Exception("Unknown error in Word conversion")

def extract_project_id(pdf_path, max_pages=10):
    """
    Extract the first occurrence of a World Bank project ID from a PDF file.
    Project IDs are in the format P followed by 6 digits (e.g., P123456).
    Also handles cases where 0 is transcribed as O.
    
    Args:
        pdf_path: Path to the PDF file
        max_pages: Maximum number of pages to search (default: 10)
        
    Returns:
        The project ID if found, None otherwise
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            # Limit the number of pages to search
            pages_to_search = min(len(reader.pages), max_pages)
            
            # Regular expression pattern for project ID (P followed by any 6 chars that could be digits or letter O)
            pattern = r'P[0-9O]{6}'
            
            # Search through pages
            for page_num in range(pages_to_search):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                # Find all matches in the page
                matches = re.findall(pattern, text)
                if matches:
                    # Get the first match and fix any O's that should be 0's
                    pid = matches[0]
                    # Replace letter 'O' with digit '0' in the project ID (starting after the P)
                    corrected_pid = 'P' + pid[1:].replace('O', '0')
                    logging.info(f"Found project ID: {pid}, corrected to: {corrected_pid}")
                    return corrected_pid
                    
        return None
        
    except Exception as e:
        logging.error(f"Error processing {pdf_path}: {str(e)}")
        return None

def get_unique_filename(base_path):
    """Generate a unique filename by adding a numeric suffix"""
    if not os.path.exists(base_path):
        return base_path
    
    directory = os.path.dirname(base_path)
    filename = os.path.basename(base_path)
    name, ext = os.path.splitext(filename)
    
    counter = 1
    while True:
        new_filename = f"{name}_{counter:02d}{ext}"
        new_path = os.path.join(directory, new_filename)
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def process_file(file_path, output_dir, input_dir, rename_with_pid=True, country_mapping=None):
    """Process a single file conversion with error handling and optional PID renaming"""
    # Initialize COM for this thread - critical for worker processes
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception as e:
        logging.error(f"COM initialization failed: {str(e)}")
    
    try:
        input_path = os.path.abspath(file_path)
        
        # Get the relative path more carefully
        rel_path = os.path.dirname(os.path.relpath(input_path, input_dir))
        
        if rel_path and rel_path != '.':
            target_dir = os.path.join(output_dir, rel_path)
            os.makedirs(target_dir, exist_ok=True)
        else:
            target_dir = output_dir
            
        output_name = Path(file_path).stem + ".pdf"
        output_path = os.path.join(target_dir, output_name)
        
        # Always check if output file already exists, regardless of the source
        if os.path.exists(output_path):
            logging.debug(f"File already exists, creating unique name: {output_path}")
            output_path = get_unique_filename(output_path)
            logging.debug(f"Using unique name: {output_path}")
        
        # Log the paths to help debug
        logging.debug(f"Converting: {input_path} -> {output_path}")
        
        # Ensure we're on Windows since Word is required
        if platform.system() != "Windows":
            raise Exception("Microsoft Word conversion requires Windows")
            
        # Convert using Word
        pdf_path = convert_with_word(input_path, output_path, retries=2)
        
        # If renaming with project ID is requested
        if rename_with_pid and pdf_path:
            # Extract project ID from the converted PDF
            project_id = extract_project_id(pdf_path)
            
            # If no project ID found in PDF content, check the filename
            if not project_id:
                logging.debug(f"No project ID found in PDF content, checking filename: {os.path.basename(file_path)}")
                project_id = extract_project_id_from_filename(os.path.basename(file_path))
                if project_id:
                    logging.debug(f"Found project ID in filename: {project_id}")
            
            if project_id:
                # Detect language
                language_suffix = detect_language(pdf_path)
                
                # Get country if available - make this more rigorous
                country = ""
                if country_mapping and project_id in country_mapping:
                    country = country_mapping[project_id]
                    country = country.replace(" ", "_")  # Replace spaces with underscores
                    logging.info(f"Found country '{country}' for project ID: {project_id}")
                else:
                    logging.debug(f"No country mapping found for project ID: {project_id}")
                
                # Create new filename with project ID, country (if available), and language
                if country:
                    pid_filename = f"{project_id}_{country}_{language_suffix}.pdf"
                else:
                    pid_filename = f"{project_id}_{language_suffix}.pdf"
                
                pid_path = os.path.join(target_dir, pid_filename)
                
                # Handle duplicate filenames
                if os.path.exists(pid_path):
                    counter = 1
                    while True:
                        if country:
                            temp_name = f"{project_id}_{country}_{language_suffix}_{counter:02d}.pdf"
                        else:
                            temp_name = f"{project_id}_{language_suffix}_{counter:02d}.pdf"
                        temp_path = os.path.join(target_dir, temp_name)
                        if not os.path.exists(temp_path):
                            pid_path = temp_path
                            pid_filename = temp_name
                            break
                        counter += 1
                
                # Rename the file
                try:
                    os.rename(pdf_path, pid_path)
                    logging.info(f"Renamed to: {pid_filename}")
                    return (file_path, True, None, project_id)
                except Exception as e:
                    logging.error(f"Error renaming: {str(e)}")
                    return (file_path, True, None, None)
            else:
                logging.warning(f"No project ID found in PDF or filename: {pdf_path}")
                
                # *** SECONDARY MATCH SECTION - ENHANCED WITH OCR DETECTION ***
                # First, check if this is an OCR/scanned document
                language_suffix = detect_language(pdf_path)
                
                if language_suffix == "OCR":
                    # This is an OCR/scanned document with no text
                    ocr_filename = f"SCAN_OCR_DOCUMENT.pdf"
                    ocr_path = os.path.join(target_dir, ocr_filename)
                    
                    # Handle duplicate filenames
                    if os.path.exists(ocr_path):
                        counter = 1
                        while True:
                            temp_name = f"SCAN_OCR_DOCUMENT_{counter:02d}.pdf"
                            temp_path = os.path.join(target_dir, temp_name)
                            if not os.path.exists(temp_path):
                                ocr_path = temp_path
                                ocr_filename = temp_name
                                break
                            counter += 1
                    
                    # Rename the file
                    try:
                        os.rename(pdf_path, ocr_path)
                        logging.info(f"Renamed as OCR document: {ocr_filename}")
                        return (file_path, True, None, "SCAN_OCR_DOCUMENT")
                    except Exception as e:
                        logging.error(f"Error renaming as OCR document: {str(e)}")
                        return (file_path, True, None, None)
                
                # If not OCR, continue with country-based matching
                if country_mapping:
                    # Extract unique country names from mapping
                    unique_countries = extract_unique_countries(country_mapping)
                    
                    # Try to find country in the PDF
                    country = extract_country_from_pdf(pdf_path, unique_countries)
                    
                    if country:
                        # Create filename with country and language
                        country_filename = f"COUNTRY_{country}_{language_suffix}.pdf"
                        country_path = os.path.join(target_dir, country_filename)
                        
                        # Handle duplicate filenames
                        if os.path.exists(country_path):
                            counter = 1
                            while True:
                                temp_name = f"COUNTRY_{country}_{language_suffix}_{counter:02d}.pdf"
                                temp_path = os.path.join(target_dir, temp_name)
                                if not os.path.exists(temp_path):
                                    country_path = temp_path
                                    country_filename = temp_name
                                    break
                                counter += 1
                        
                        # Copy the file
                        try:
                            shutil.copy2(pdf_path, country_path)
                            logging.info(f"Copied with country match: {country_filename}")
                            return (file_path, True, None, f"COUNTRY_{country}")
                        except Exception as e:
                            logging.error(f"Error copying with country match: {str(e)}")
                            return (file_path, True, None, None)
                
                # If we reach here, neither project ID nor country was found, and it's not OCR
                # Just mark with language tag
                unknown_filename = f"UNKNOWN_{language_suffix}.pdf"
                unknown_path = os.path.join(target_dir, unknown_filename)
                
                # Handle duplicate filenames
                if os.path.exists(unknown_path):
                    counter = 1
                    while True:
                        temp_name = f"UNKNOWN_{language_suffix}_{counter:02d}.pdf"
                        temp_path = os.path.join(target_dir, temp_name)
                        if not os.path.exists(temp_path):
                            unknown_path = temp_path
                            unknown_filename = temp_name
                            break
                        counter += 1
                
                # Rename the file
                try:
                    os.rename(pdf_path, unknown_path)
                    logging.info(f"Renamed as unknown document: {unknown_filename}")
                    return (file_path, True, None, f"UNKNOWN_{language_suffix}")
                except Exception as e:
                    logging.error(f"Error renaming as unknown document: {str(e)}")
                    return (file_path, True, None, None)
        
        return (file_path, True, None, None)
    except Exception as e:
        logging.error(f"Process file error: {str(e)}")
        return (file_path, False, str(e), None)
    finally:
        # Always uninitialize COM before exiting
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass

def copy_existing_pdfs(input_dir, output_dir, overwrite=False, rename_with_pid=True, country_mapping=None):
    """Copy all existing PDF files from input directory to output directory using a sequential classification system"""
    pdf_files = []
    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return 0, {}
    
    print(f"\nCopying {len(pdf_files)} existing PDF files to output directory")
    
    copied = 0
    skipped = 0
    pid_mapping = {}  # To store file -> project ID mapping
    
    # Extract unique country names from mapping (do this once)
    unique_countries = extract_unique_countries(country_mapping) if country_mapping else set()
    
    with tqdm(total=len(pdf_files), unit="file", desc="Copying PDFs", ncols=100, bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]') as pbar:
        for pdf_file in pdf_files:
            # Create the relative path for maintaining folder structure
            rel_path = os.path.relpath(os.path.dirname(pdf_file), start=input_dir)
            if rel_path != '.':
                target_dir = os.path.join(output_dir, rel_path)
                os.makedirs(target_dir, exist_ok=True)
            else:
                target_dir = output_dir
            
            # Flag to track if we've processed this file
            file_processed = False
            
            try:
                # ALWAYS process through the renaming pipeline, regardless of rename_with_pid setting
                # Only use rename_with_pid to determine if we should try to extract project IDs
                
                # PRIORITY 1: Extract project ID (only if rename_with_pid is True)
                if rename_with_pid:
                    # Extract project ID from the PDF or filename
                    project_id = extract_project_id(pdf_file)
                    
                    if not project_id:
                        # Check filename if not found in content
                        project_id = extract_project_id_from_filename(os.path.basename(pdf_file))
                    
                    if project_id:
                        # We found a project ID - use this classification
                        language_suffix = detect_language(pdf_file)
                        
                        # Get country if available
                        country = ""
                        if country_mapping and project_id in country_mapping:
                            country = country_mapping[project_id]
                            country = country.replace(" ", "_")  # Replace spaces with underscores
                        
                        # Create new filename with project ID
                        if country:
                            new_filename = f"{project_id}_{country}_{language_suffix}.pdf"
                        else:
                            new_filename = f"{project_id}_{language_suffix}.pdf"
                        
                        new_path = os.path.join(target_dir, new_filename)
                        
                        # Handle duplicate filenames
                        if os.path.exists(new_path):
                            counter = 1
                            while True:
                                if country:
                                    temp_name = f"{project_id}_{country}_{language_suffix}_{counter:02d}.pdf"
                                else:
                                    temp_name = f"{project_id}_{language_suffix}_{counter:02d}.pdf"
                                temp_path = os.path.join(target_dir, temp_name)
                                if not os.path.exists(temp_path):
                                    new_path = temp_path
                                    new_filename = temp_name
                                    break
                                counter += 1
                        
                        # Copy the file
                        shutil.copy2(pdf_file, new_path)
                        logging.info(f"Copied with project ID: {new_filename}")
                        copied += 1
                        pid_mapping[new_path] = project_id
                        file_processed = True
                
                # PRIORITY 2: If no project ID or rename_with_pid is False, check if it's an OCR document
                if not file_processed:
                    language_suffix = detect_language(pdf_file)
                    
                    if language_suffix == "OCR":
                        # This is an OCR document with no selectable text
                        ocr_filename = "SCAN_OCR_DOCUMENT.pdf"
                        ocr_path = os.path.join(target_dir, ocr_filename)
                        
                        # Handle duplicate filenames
                        if os.path.exists(ocr_path):
                            counter = 1
                            while True:
                                temp_name = f"SCAN_OCR_DOCUMENT_{counter:02d}.pdf"
                                temp_path = os.path.join(target_dir, temp_name)
                                if not os.path.exists(temp_path):
                                    ocr_path = temp_path
                                    ocr_filename = temp_name
                                    break
                                counter += 1
                        
                        # Copy the file
                        shutil.copy2(pdf_file, ocr_path)
                        logging.info(f"Copied as OCR document: {ocr_filename}")
                        copied += 1
                        pid_mapping[ocr_path] = "SCAN_OCR_DOCUMENT"
                        file_processed = True
                
                # PRIORITY 3: If not OCR, check for country match (only if we have country mappings)
                if not file_processed and unique_countries:
                    country = extract_country_from_pdf(pdf_file, unique_countries)
                    
                    if country:
                        language_suffix = detect_language(pdf_file)
                        # Create filename with country and language
                        country_filename = f"COUNTRY_{country}_{language_suffix}.pdf"
                        country_path = os.path.join(target_dir, country_filename)
                        
                        # Handle duplicate filenames
                        if os.path.exists(country_path):
                            counter = 1
                            while True:
                                temp_name = f"COUNTRY_{country}_{language_suffix}_{counter:02d}.pdf"
                                temp_path = os.path.join(target_dir, temp_name)
                                if not os.path.exists(temp_path):
                                    country_path = temp_path
                                    country_filename = temp_name
                                    break
                                counter += 1
                        
                        # Copy the file
                        shutil.copy2(pdf_file, country_path)
                        logging.info(f"Copied with country match: {country_filename}")
                        copied += 1
                        pid_mapping[country_path] = f"COUNTRY_{country}"
                        file_processed = True
                
                # PRIORITY 4: Last resort - mark as unknown
                if not file_processed:
                    language_suffix = detect_language(pdf_file) if language_suffix not in ["EN", "NON", "UNK", "OCR"] else language_suffix
                    unknown_filename = f"UNKNOWN_{language_suffix}.pdf"
                    unknown_path = os.path.join(target_dir, unknown_filename)
                    
                    # Handle duplicate filenames
                    if os.path.exists(unknown_path):
                        counter = 1
                        while True:
                            temp_name = f"UNKNOWN_{language_suffix}_{counter:02d}.pdf"
                            temp_path = os.path.join(target_dir, temp_name)
                            if not os.path.exists(temp_path):
                                unknown_path = temp_path
                                unknown_filename = temp_name
                                break
                            counter += 1
                    
                    # Copy the file
                    shutil.copy2(pdf_file, unknown_path)
                    logging.info(f"Copied as unknown document: {unknown_filename}")
                    copied += 1
                    pid_mapping[unknown_path] = f"UNKNOWN_{language_suffix}"
                    file_processed = True
                
            except Exception as e:
                error_msg = f"Error copying {pdf_file}: {str(e)}"
                logging.error(error_msg)
                skipped += 1
            
            pbar.update(1)
    
    print(f"PDF copying complete. Copied: {copied}, Skipped: {skipped}")
    return copied, pid_mapping

def convert_folder_to_pdf(rename_with_pid=True, country_mapping=None, workers=None):
    """Convert all Word documents in a folder to PDF"""
    # Check if we're on Windows
    if platform.system() != "Windows":
        print("Error: This script requires Windows with Microsoft Word installed")
        return 1
    
    # Prompt user for the input folder path
    print("Please enter the path to the folder containing Word documents:")
    input_dir = input().strip()
    
    # Strip quotes if the user included them
    input_dir = input_dir.strip('"\'')
    
    # Validate input directory
    if not os.path.isdir(input_dir):
        print(f"Error: '{input_dir}' is not a valid directory")
        return 1
    
    # Prompt for country spreadsheet if not provided
    if rename_with_pid and country_mapping is None:
        print("Do you want to use a spreadsheet to map project IDs to countries? (y/n):")
        use_country_mapping = input().strip().lower() == 'y'
        
        if use_country_mapping:
            print("Please enter the path to the spreadsheet file (Excel or CSV):")
            spreadsheet_path = input().strip().strip('"\'')
            
            if os.path.exists(spreadsheet_path):
                country_mapping = load_project_country_mapping(spreadsheet_path)
                if not country_mapping:
                    print("Warning: No valid project ID to country mappings found in the spreadsheet.")
                    print("Files will be renamed with project IDs only.")
            else:
                print(f"Warning: Spreadsheet file not found: {spreadsheet_path}")
                country_mapping = {}
    
    # Prompt user for the output folder path
    print("Please enter the path to the output folder for PDF files:")
    output_dir = input().strip()
    
    # Strip quotes if the user included them
    output_dir = output_dir.strip('"\'')
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Find all .docx and .doc files
    word_files = []
    # Find all .pdf files
    pdf_files = []
    
    print("\nScanning input directory for documents...")
    for root, _, files in os.walk(input_dir):
        for file in files:
            lower_file = file.lower()
            if lower_file.endswith('.docx') or lower_file.endswith('.doc'):
                word_files.append(os.path.join(root, file))
            elif lower_file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    # Dictionary to store project ID mappings
    project_id_mappings = {}
    
    # Initialize counters BEFORE the conditional block
    successful = 0
    failed = 0
    
    if not word_files:
        print("No Word documents (.doc or .docx) found in the input directory")
        # Even if no Word files are found, we'll still copy PDFs
    else:
        # Use provided worker count or determine optimal count
        if workers is not None:
            max_workers = workers
        else:
            max_workers = get_optimal_worker_count(len(word_files))
        
        print(f"Found {len(word_files)} Word documents to convert")
        print(f"Using Microsoft Word for conversion with {max_workers} worker processes")
        
        # Initialize timing
        start_time = time.time()
        
        # Determine batch size based on worker count and system memory
        batch_size = get_optimal_batch_size(max_workers)
        
        print(f"Using batch size: {batch_size}")
        
        # Calculate total files to process
        total_files = len(word_files) + len(pdf_files)
        print(f"Total files to process: {total_files}")
        
        # Process in smaller batches to prevent memory issues
        for i in range(0, len(word_files), batch_size):
            batch = word_files[i:i+batch_size]
            
            print(f"\nProcessing batch {i//batch_size + 1} of {(len(word_files) + batch_size - 1) // batch_size} ({len(batch)} files)")
            
            # Clean up any existing Word processes before each batch
            try:
                subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                             stdout=subprocess.DEVNULL, 
                             stderr=subprocess.DEVNULL)
                time.sleep(1)  # Give system time to close Word
            except:
                pass
            
            # Add a check for available memory before starting a new batch
            mem = psutil.virtual_memory()
            if mem.percent > 90:  # If memory usage is very high
                print("System memory usage high ({}%). Waiting for 10 seconds before continuing...".format(mem.percent))
                time.sleep(10)  # Wait for memory to potentially free up
            
            # Use ThreadPoolExecutor instead for better COM compatibility
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit jobs for this batch with input_dir as an additional argument
                future_to_file = {
                    executor.submit(process_file, file, output_dir, input_dir, rename_with_pid, country_mapping): file
                    for file in batch
                }
                
                # Process results
                with tqdm(total=len(batch), unit="file", desc="Converting", ncols=100, bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]') as pbar:
                    for future in concurrent.futures.as_completed(future_to_file):
                        try:
                            file_path, success, error, project_id = future.result()
                            if success:
                                successful += 1
                                if project_id:
                                    project_id_mappings[file_path] = project_id
                            else:
                                failed += 1
                                logging.error(f"Error converting {file_path}: {error}")
                        except Exception as exc:
                            failed += 1
                            file_path = future_to_file[future]
                            logging.error(f"Exception during conversion of {file_path}: {str(exc)}")
                        finally:
                            pbar.update(1)
            
            # Clean up after batch
            try:
                subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                             stdout=subprocess.DEVNULL, 
                             stderr=subprocess.DEVNULL)
                time.sleep(1)
            except:
                pass
        
        # Report results
        elapsed_time = time.time() - start_time
        files_per_second = len(word_files) / elapsed_time if elapsed_time > 0 else 0
        
        print(f"\nWord conversion complete in {elapsed_time:.2f} seconds ({files_per_second:.2f} files/sec)")
        print(f"Successfully converted: {successful}")
        print(f"Failed conversions: {failed}")
        
        # Add success rate report
        if successful + failed > 0:
            print(f"Success rate: {successful/(successful+failed)*100:.1f}%")
        else:
            print("Success rate: N/A (no files processed)")
    
    # Always set overwrite=False to ensure no files are overwritten
    copied_pdfs, pdf_pid_mappings = copy_existing_pdfs(input_dir, output_dir, overwrite=False, rename_with_pid=rename_with_pid, country_mapping=country_mapping)
    
    # Merge the project ID mappings
    project_id_mappings.update(pdf_pid_mappings)
    
    print(f"\nAll operations complete. Output files saved to: {output_dir}")
    
    # Verify the output directory contents
    output_files = []
    for root, _, files in os.walk(output_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                output_files.append(os.path.join(root, file))
    
    print(f"Actual PDF files in output directory: {len(output_files)}")
    
    if len(output_files) < (successful + copied_pdfs):
        print("WARNING: Some files may have been overwritten due to naming conflicts.")
    
    # Create a summary of project IDs found
    if rename_with_pid:
        project_ids = list(set(project_id_mappings.values()))
        
        # Count how many are project IDs vs country-based vs OCR vs unknown
        regular_pids = [pid for pid in project_ids if not pid.startswith(("COUNTRY_", "SCAN_", "UNKNOWN_"))]
        country_based = [pid for pid in project_ids if pid.startswith("COUNTRY_")]
        ocr_documents = [pid for pid in project_ids if pid.startswith("SCAN_OCR")]
        unknown_docs = [pid for pid in project_ids if pid.startswith("UNKNOWN_")]
        
        print(f"\nFound {len(regular_pids)} files with unique project IDs")
        if regular_pids:
            print("Sample of project IDs found:")
            for pid in regular_pids[:5]:  # Show first 5
                print(f"  - {pid}")
            if len(regular_pids) > 5:
                print(f"  ...and {len(regular_pids) - 5} more")
        
        if country_based:
            print(f"\nFound {len(country_based)} files renamed by country match (no project ID)")
            print("Sample of country-based names:")
            for country_pid in country_based[:5]:  # Show first 5
                print(f"  - {country_pid}")
            if len(country_based) > 5:
                print(f"  ...and {len(country_based) - 5} more")
        
        if ocr_documents:
            print(f"\nFound {len(ocr_documents)} scanned/OCR documents with no extractable text")
        
        if unknown_docs:
            print(f"\nFound {len(unknown_docs)} documents with no identifiable project ID or country")
    
    print("\nLanguage suffix explanation:")
    print("  EN: Document is primarily in English")
    print("  NON: Document is primarily in a non-English language")
    print("  UNK: Language could not be determined (insufficient text for detection)")
    print("  OCR: Document appears to be scanned/image-based with no selectable text")
    
    print("\nFilename format explanation:")
    print("  PROJECT_ID_COUNTRY_LANGUAGE.pdf: Files with detected project IDs")
    print("  COUNTRY_COUNTRYNAME_LANGUAGE.pdf: Files with no project ID but detected country")
    print("  SCAN_OCR_DOCUMENT_##.pdf: Scanned documents with no extractable text")
    print("  UNKNOWN_LANGUAGE_##.pdf: Files with no identifiable project ID or country")
    
    return 0, output_dir  # Return tuple with exit code and output directory

def verify_pdf(pdf_path):
    """Verify that the created PDF is valid"""
    try:
        # Use PyPDF2 to check PDF validity
        import PyPDF2
        with open(pdf_path, 'rb') as file:
            try:
                pdf = PyPDF2.PdfReader(file)
                # Try to access pages to ensure it's readable
                num_pages = len(pdf.pages)
                return True
            except Exception:
                return False
    except ImportError:
        # If PyPDF2 is not installed, just check file size
        return os.path.getsize(pdf_path) > 100  # Assume valid if > 100 bytes

def is_file_locked(file_path):
    """Check if a file is locked (in use by another process)"""
    try:
        with open(file_path, 'r+b') as f:
            return False
    except IOError:
        return True

def get_optimal_batch_size(worker_count=4):
    """Determine optimal batch size based on available system memory and CPU cores
    
    Args:
        worker_count: Number of concurrent worker processes
        
    Returns:
        Optimal batch size for processing
    """
    mem = psutil.virtual_memory()
    
    # Base batch size on available memory
    if mem.total < 8 * 1024 * 1024 * 1024:  # 8 GB
        base_batch = 5
    elif mem.total < 16 * 1024 * 1024 * 1024:  # 16 GB
        base_batch = 10
    else:
        base_batch = 20
    
    # Scale batch size inversely with worker count to avoid memory pressure
    # More workers = smaller batches per worker
    adjusted_batch = max(5, int(base_batch * (4 / max(1, worker_count))))
    
    return adjusted_batch

def parse_args():
    parser = argparse.ArgumentParser(description='Convert Word documents to PDF and rename with Project IDs')
    parser.add_argument('--input', '-i', help='Input directory containing Word documents')
    parser.add_argument('--output', '-o', help='Output directory for PDF files')
    parser.add_argument('--rename', '-r', action='store_true', help='Rename files with project IDs', default=True)
    parser.add_argument('--no-rename', action='store_false', dest='rename', help="Don't rename files with project IDs")
    parser.add_argument('--workers', '-w', type=int, help='Number of worker processes to use (default: auto)', default=None)
    return parser.parse_args()

def normalize_path(path):
    """Ensure path is in a format Word can handle"""
    # Convert UNC paths to mapped drives if needed
    if path.startswith('\\\\'):
        # For UNC paths, consider using temporary local copies
        # or map a drive letter temporarily
        pass
    return os.path.abspath(path)

def extract_project_id_from_filename(filename):
    """
    Extract a project ID from a filename if present.
    Project IDs are in the format P followed by 6 digits,
    followed by either a hyphen or underscore (e.g., P123456- or P123456_).
    
    Args:
        filename: The filename to check
        
    Returns:
        The project ID if found, None otherwise
    """
    try:
        # Regular expression pattern for project ID in filename
        # Looks for P + 6 digits + (- or _)
        pattern = r'P\d{6}[-_]'
        
        # Find all matches in the filename
        matches = re.findall(pattern, filename)
        if matches:
            # Return the first match without the trailing - or _
            return matches[0][:-1]
                
        return None
        
    except Exception as e:
        logging.error(f"Error processing filename {filename}: {str(e)}")
        return None

def detect_language(pdf_path, pages_to_check=3):
    """
    Detect if a PDF document is primarily in English or not.
    Also detect if the document has no selectable text (scanned).
    
    Args:
        pdf_path: Path to the PDF file
        pages_to_check: Number of pages to analyze for language detection
        
    Returns:
        "EN" if English is detected
        "NON" if a non-English language is detected
        "UNK" if language could not be determined
        "OCR" if the document appears to be scanned/image-based with no selectable text
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Limit the number of pages to check
            pages_to_check = min(len(reader.pages), pages_to_check)
            
            # Concatenate text from multiple pages for better detection
            all_text = ""
            for page_num in range(pages_to_check):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text:
                    all_text += text
                    # Once we have a decent amount of text, we can stop
                    if len(all_text) > 1000:
                        break
            
            # Check if document appears to be scanned (no selectable text)
            if len(all_text.strip()) < 50 and len(reader.pages) > 0:
                # Document has pages but very little text - likely scanned/image-based
                logging.info(f"Document appears to be scanned or image-based: {pdf_path}")
                return "OCR"
            
            # If we have enough text to detect language
            if len(all_text) > 100:
                try:
                    lang = detect(all_text)
                    return "EN" if lang == "en" else "NON"
                except LangDetectException:
                    logging.warning(f"Could not detect language in {pdf_path}")
                    return "UNK"  # Unknown language
            else:
                if len(all_text) > 0:
                    # Some text, but not enough for confident detection
                    logging.warning(f"Not enough text for language detection in {pdf_path}")
                    return "UNK"  # Unknown due to insufficient text
                else:
                    # No text at all
                    return "OCR"  # Likely scanned
                    
    except Exception as e:
        logging.error(f"Error detecting language in {pdf_path}: {str(e)}")
        return "UNK"  # Unknown due to error

def load_project_country_mapping(spreadsheet_path, pid_column=None, country_column=None):
    """
    Load project ID to country mapping from a spreadsheet.
    
    Args:
        spreadsheet_path: Path to the spreadsheet file (Excel or CSV)
        pid_column: Column name/index containing project IDs
        country_column: Column name/index containing country names
        
    Returns:
        Dictionary mapping project IDs to countries
    """
    try:
        # Check file extension
        file_ext = os.path.splitext(spreadsheet_path)[1].lower()
        
        # Load the spreadsheet based on file type
        if file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(spreadsheet_path)
        elif file_ext == '.csv':
            df = pd.read_csv(spreadsheet_path)
        else:
            logging.error(f"Unsupported file format: {file_ext}")
            return {}
        
        # Convert all column names to strings to avoid any type issues
        df.columns = [str(col) for col in df.columns]
        
        # Show the column names to the user if not specified
        if pid_column is None or country_column is None:
            print("\nAvailable columns in the spreadsheet:")
            for i, col in enumerate(df.columns):
                print(f"{i}: {col}")
            
            if pid_column is None:
                pid_column = input("\nEnter the number or name of the column containing Project IDs: ").strip()
                # Try to convert to integer if it's a number
                try:
                    pid_column = int(pid_column)
                    pid_column = df.columns[pid_column]
                except ValueError:
                    # If not an integer, use as column name
                    pass
                print(f"Using column '{pid_column}' for Project IDs")
            
            if country_column is None:
                country_column = input("Enter the number or name of the column containing Countries: ").strip()
                # Try to convert to integer if it's a number
                try:
                    country_column = int(country_column)
                    country_column = df.columns[country_column]
                except ValueError:
                    # If not an integer, use as column name
                    pass
                print(f"Using column '{country_column}' for Countries")
        
        # Ensure the columns exist
        if pid_column not in df.columns:
            logging.error(f"Project ID column '{pid_column}' not found in spreadsheet")
            print(f"Error: Project ID column '{pid_column}' not found in spreadsheet")
            return {}
        
        if country_column not in df.columns:
            logging.error(f"Country column '{country_column}' not found in spreadsheet")
            print(f"Error: Country column '{country_column}' not found in spreadsheet")
            return {}
        
        # Create the mapping
        mapping = {}
        for _, row in df.iterrows():
            # Convert to string and strip whitespace to ensure consistency
            project_id = str(row[pid_column]).strip()
            country = str(row[country_column]).strip()
            
            # Skip empty values
            if not project_id or not country or project_id.lower() == 'nan' or country.lower() == 'nan':
                continue
            
            # Handle project IDs that may not start with 'P'
            if project_id and not project_id.startswith('P') and project_id.isdigit():
                project_id = f"P{project_id}"
            
            # Clean project ID to ensure it follows the P###### format
            # Remove any non-alphanumeric characters
            project_id = re.sub(r'[^P0-9]', '', project_id)
            
            # Make sure it matches our expected format
            if re.match(r'P\d{6}', project_id):
                mapping[project_id] = country
        
        print(f"Loaded {len(mapping)} project ID to country mappings")
        
        # Show a sample of the mappings
        if mapping:
            sample_size = min(5, len(mapping))
            print("Sample mappings:")
            sample_items = list(mapping.items())[:sample_size]
            for pid, country in sample_items:
                print(f"  {pid} -> {country}")
        
        return mapping
    
    except Exception as e:
        logging.error(f"Error loading project country mapping: {str(e)}")
        print(f"Error loading spreadsheet: {str(e)}")
        return {}

def get_optimal_worker_count(file_count):
    """Determine the optimal number of worker processes based on system resources
    
    Args:
        file_count: Number of files to be processed
        
    Returns:
        Optimal number of worker processes
    """
    # Get system information
    cpu_count = os.cpu_count() or 4  # Default to 4 if we can't determine
    mem = psutil.virtual_memory()
    
    # Base worker count on CPU cores
    # Start with a base value of CPU cores with some headroom for system
    base_workers = max(1, cpu_count - 1)
    
    # Adjust for memory constraints - Word can use significant memory
    # For systems with less memory, reduce worker count
    if mem.total < 8 * 1024 * 1024 * 1024:  # 8 GB
        mem_factor = 0.5
    elif mem.total < 16 * 1024 * 1024 * 1024:  # 16 GB
        mem_factor = 0.75
    else:
        mem_factor = 1.0
    
    # Consider the number of files - no need for many workers with few files
    # Don't create more workers than there are files to process
    file_limit = max(1, file_count // 2)
    
    # Calculate final worker count, ensuring we have at least 1
    worker_count = max(1, min(base_workers, file_limit, int(base_workers * mem_factor)))
    
    # Cap at a reasonable maximum to prevent system overload
    # Word processing can be very resource-intensive
    return min(worker_count, 8)

def is_selectable_text_pdf(pdf_path):
    """
    Check if a PDF has selectable text or is a scanned/image document
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        True if the PDF has selectable text, False if it appears to be scanned/image-based
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Check up to 3 pages or all pages if fewer
            pages_to_check = min(len(reader.pages), 3)
            
            # If document has no pages, return True (not our concern)
            if pages_to_check == 0:
                return True
                
            # Keep track of text content
            total_text = 0
            
            # Check each page for text content
            for i in range(pages_to_check):
                page = reader.pages[i]
                text = page.extract_text()
                
                if text and len(text.strip()) > 50:  # A reasonable text page should have more than 50 chars
                    return True  # Found a page with significant text
                    
                total_text += len(text.strip())
            
            # If we checked multiple pages and found very little text, it's likely a scanned document
            if total_text < 100 and pages_to_check > 0:
                return False
                
            # Default to True if we can't be sure
            return True
            
    except Exception as e:
        logging.error(f"Error checking if PDF has selectable text: {str(e)}")
        return True  # Default to True on error

def extract_unique_countries(country_mapping):
    """
    Extract all unique country names from the project ID to country mapping,
    excluding 'World' from the results
    
    Args:
        country_mapping: Dictionary mapping project IDs to country names
        
    Returns:
        Set of unique country names (excluding 'World')
    """
    if not country_mapping:
        return set()
        
    # Extract all unique country names from the mapping
    countries = set()
    for country in country_mapping.values():
        # Skip any variant of "World" (case-insensitive)
        if country.lower().strip() != "world":
            countries.add(country)
    
    return countries

def extract_country_from_pdf(pdf_path, country_names, max_pages=10):
    """
    Search for country names in a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        country_names: Set of country names to search for
        max_pages: Maximum number of pages to search
        
    Returns:
        The country name if found, None otherwise
    """
    if not country_names:
        return None
        
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            pages_to_search = min(len(reader.pages), max_pages)
            
            # Create a list of country names sorted by length (descending)
            # This prioritizes longer names (e.g., "United States" over "States")
            sorted_countries = sorted(country_names, key=len, reverse=True)
            
            # Create patterns for each country name
            # This handles case-insensitive matching and boundary detection
            country_patterns = {}
            for country in sorted_countries:
                # Create regex pattern with word boundaries
                pattern = r'\b' + re.escape(country) + r'\b'
                country_patterns[country] = re.compile(pattern, re.IGNORECASE)
                
                # Also add pattern with underscores replaced by spaces
                space_country = country.replace('_', ' ')
                if space_country != country:
                    pattern = r'\b' + re.escape(space_country) + r'\b'
                    country_patterns[country] = re.compile(pattern, re.IGNORECASE)
            
            # Search through pages
            for page_num in range(pages_to_search):
                page = reader.pages[page_num]
                text = page.extract_text()
                if not text:
                    continue
                    
                # Check each country pattern
                for country, pattern in country_patterns.items():
                    if pattern.search(text):
                        logging.info(f"Found country: {country} in document: {pdf_path}")
                        return country
                        
        return None
        
    except Exception as e:
        logging.error(f"Error searching for country in {pdf_path}: {str(e)}")
        return None

if __name__ == "__main__":
    args = parse_args()
    exit_code = 0
    output_dir = None  # Initialize output_dir variable
    
    # Prompt for document type at the beginning
    print("\n" + "-"*80)
    print(" DOCUMENT TYPE ".center(80, "-"))
    print("-"*80)
    print("What type of documents are being processed?")
    print("Examples: icrr, aidememoire, pad, esrs, etc.")
    print("This will be added to the filenames for better tracking.")
    document_type = input("Enter document type: ").strip().lower()
    
    # Validate input
    if not document_type:
        print("No document type entered. Proceeding without adding document type to filenames.")
    else:
        print(f"Using '{document_type}' as the document type identifier.")
    
    if args.input and args.output:
        # TODO: Add command-line mode implementation
        output_dir = args.output  # Use output directory from command line arguments
    else:
        # Modify convert_folder_to_pdf to return both exit_code and output_dir
        exit_code, output_dir = convert_folder_to_pdf(rename_with_pid=True, workers=args.workers)
    
    try:
        import reorganize_output
        if output_dir and os.path.exists(output_dir):
            print("\nStarting file reorganization...")
            # Pass document_type to reorganize_output_folder
            reorganize_output.reorganize_output_folder(output_dir, document_type)
        else:
            print("\nSkipping reorganization - output directory not available")
    except Exception as e:
        print(f"Error during reorganization: {str(e)}")
    
    sys.exit(exit_code) 
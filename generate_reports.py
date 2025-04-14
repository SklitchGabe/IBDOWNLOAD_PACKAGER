import os
import sys
import re
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("report_generation.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

def generate_reports(main_output_dir, portfolio_spreadsheet=None):
    """
    Generate reports about document types found for each project ID and country.
    
    Args:
        main_output_dir: The main output directory containing the organized documents
        portfolio_spreadsheet: Optional path to the user's portfolio spreadsheet
    """
    print("\n" + "="*80)
    print(" GENERATING DOCUMENT REPORTS ".center(80, "="))
    print("="*80 + "\n")
    
    # Path to the country documents folder
    country_folder = os.path.join(main_output_dir, "Country Associated Documents")
    
    if not os.path.exists(country_folder):
        print(f"Country Associated Documents folder not found at: {country_folder}")
        return
    
    # Get the current date for the report filenames
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Initialize data structures to store document information
    # Format: {project_id: {"country": country_name, "document_types": set(), "count": int}}
    master_data = {}
    
    # Format: {country_name: {project_id: {"document_types": set(), "files": list()}}}
    country_data = {}
    
    # Track all unique document types across all files
    all_document_types = set()
    
    # Regular expressions to extract information from filenames
    pid_pattern = re.compile(r'^(P\d{6})_(.+?)_(EN|NON|UNK|OCR)_([a-z]+)(?:_\d+)?\.pdf$', re.IGNORECASE)
    country_pattern = re.compile(r'^COUNTRY_(.+?)_(EN|NON|UNK|OCR)_([a-z]+)(?:_\d+)?\.pdf$', re.IGNORECASE)
    
    # Load project ID to country mapping from portfolio if available
    portfolio_mapping = {}
    if portfolio_spreadsheet and os.path.exists(portfolio_spreadsheet):
        try:
            print(f"Loading project portfolio from: {portfolio_spreadsheet}")
            
            # Determine file type and read
            if portfolio_spreadsheet.lower().endswith('.csv'):
                df = pd.read_csv(portfolio_spreadsheet)
            else:
                df = pd.read_excel(portfolio_spreadsheet)
            
            # Prompt for column names if we have a portfolio
            print("\nPlease enter the column name containing Project IDs:")
            pid_column = input().strip()
            
            print("Please enter the column name containing Countries:")
            country_column = input().strip()
            
            # Validate columns exist
            if pid_column in df.columns and country_column in df.columns:
                # Create mapping and initialize master data for all project IDs
                for _, row in df.iterrows():
                    pid = str(row[pid_column]).strip()
                    country = str(row[country_column]).strip()
                    
                    # Only add valid project IDs (starting with P followed by digits)
                    if re.match(r'^P\d+$', pid):
                        # Standardize to P followed by exactly 6 digits
                        if len(pid) > 1:
                            # Extract just the digits and ensure 6 digits with leading zeros
                            digits = re.search(r'\d+', pid).group()
                            pid = f"P{int(digits):06d}"
                            
                            portfolio_mapping[pid] = country
                            # Initialize in master data with empty document types
                            master_data[pid] = {
                                "country": country,
                                "document_types": set(),
                                "count": 0
                            }
                
                print(f"Loaded {len(portfolio_mapping)} project IDs from portfolio")
            else:
                print(f"Warning: Columns '{pid_column}' or '{country_column}' not found in spreadsheet")
        except Exception as e:
            print(f"Error loading portfolio spreadsheet: {str(e)}")
    
    # Process each country subfolder
    print("\nScanning country folders for documents...")
    country_count = 0
    document_count = 0
    
    for country_dir in os.listdir(country_folder):
        country_path = os.path.join(country_folder, country_dir)
        if os.path.isdir(country_path):
            country_count += 1
            country_name = country_dir.replace('_', ' ')
            
            # Initialize country data
            if country_name not in country_data:
                country_data[country_name] = {}
            
            # First look in the country root folder
            for file in os.listdir(country_path):
                file_path = os.path.join(country_path, file)
                if file.lower().endswith('.pdf') and os.path.isfile(file_path):
                    document_count += 1
                    process_document(file, file_path, country_name, master_data, country_data, all_document_types)
            
            # Then look in any document type subfolders
            for subdir in os.listdir(country_path):
                subdir_path = os.path.join(country_path, subdir)
                if os.path.isdir(subdir_path):
                    for file in os.listdir(subdir_path):
                        file_path = os.path.join(subdir_path, file)
                        if file.lower().endswith('.pdf') and os.path.isfile(file_path):
                            document_count += 1
                            process_document(file, file_path, country_name, master_data, country_data, all_document_types)
    
    print(f"Processed {document_count} documents across {country_count} countries")
    print(f"Found {len(all_document_types)} unique document types: {', '.join(sorted(all_document_types))}")
    
    # Generate the master report with document type columns
    generate_master_report(master_data, portfolio_mapping, main_output_dir, today, all_document_types)
    
    # Generate country-specific reports
    generate_country_reports(country_data, country_folder, today)
    
    print("\nReport generation complete!")
    print(f"Master report saved to: {os.path.join(main_output_dir, f'document_inventory_{today}.xlsx')}")

def process_document(filename, file_path, country_name, master_data, country_data, all_document_types):
    """Process a document file and update the data structures"""
    # Extract project ID and document type from filename
    pid_match = re.search(r'(P\d{6})', filename)
    doc_type_match = re.search(r'_([a-z]+)(?:_\d+)?\.pdf$', filename.lower())
    
    project_id = pid_match.group(1) if pid_match else None
    doc_type = doc_type_match.group(1) if doc_type_match else "unknown"
    
    # Skip language markers as document types
    if doc_type in ['en', 'non', 'unk', 'ocr']:
        # Try to find a secondary document type
        secondary_match = re.search(r'_([a-z]+)_(?:en|non|unk|ocr)', filename.lower())
        doc_type = secondary_match.group(1) if secondary_match else "unknown"
    
    # Add to all document types
    all_document_types.add(doc_type)
    
    # Document information
    doc_info = {
        "filename": filename,
        "document_type": doc_type,
        "country": country_name
    }
    
    # Update master data if we have a project ID
    if project_id:
        if project_id not in master_data:
            master_data[project_id] = {
                "country": country_name,
                "document_types": set(),
                "count": 0
            }
        
        master_data[project_id]["document_types"].add(doc_type)
        master_data[project_id]["count"] += 1
    
    # Update country data
    country_dict = country_data.get(country_name, {})
    
    # Use project ID if available, otherwise use filename as key
    doc_key = project_id if project_id else f"No_PID_{os.path.basename(filename)}"
    
    if doc_key not in country_dict:
        country_dict[doc_key] = {
            "document_types": set(),
            "files": []
        }
    
    country_dict[doc_key]["document_types"].add(doc_type)
    country_dict[doc_key]["files"].append(doc_info)
    
    # Ensure the country dict is in the country data
    country_data[country_name] = country_dict

def generate_master_report(master_data, portfolio_mapping, output_dir, date_str, all_document_types):
    """Generate the master report showing all project IDs and their document types"""
    print("\nGenerating master document inventory report...")
    
    # Sort document types for consistent column order
    all_doc_types_sorted = sorted(all_document_types)
    
    # Create report data structure - initialize with one row per project ID
    report_data = {}
    
    # Process each project ID from the portfolio first
    for pid, country in portfolio_mapping.items():
        report_row = {
            "Project ID": pid,
            "Country": country,
            "Has Documents": "No"
        }
        
        # Add binary flags for each document type (0 = No, 1 = Yes)
        for doc_type in all_doc_types_sorted:
            report_row[doc_type] = 0
        
        # Update if we have document data
        if pid in master_data:
            doc_info = master_data[pid]
            report_row["Has Documents"] = "Yes"
            report_row["Document Count"] = doc_info["count"]
            
            # Set flags for document types that exist
            for doc_type in doc_info["document_types"]:
                report_row[doc_type] = 1
        else:
            report_row["Document Count"] = 0
        
        report_data[pid] = report_row
    
    # Process any additional project IDs found in the documents but not in portfolio
    for pid, doc_info in master_data.items():
        if pid not in portfolio_mapping:
            report_row = {
                "Project ID": pid,
                "Country": doc_info["country"],
                "Has Documents": "Yes", 
                "Document Count": doc_info["count"],
                "Note": "Not in original portfolio"
            }
            
            # Add binary flags for each document type
            for doc_type in all_doc_types_sorted:
                report_row[doc_type] = 1 if doc_type in doc_info["document_types"] else 0
                
            report_data[pid] = report_row
    
    # Convert to DataFrame
    if report_data:
        df = pd.DataFrame(list(report_data.values()))
        
        # Sort by country and project ID
        df = df.sort_values(by=["Country", "Project ID"])
        
        # Save to Excel file
        output_file = os.path.join(output_dir, f"document_inventory_{date_str}.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Document Inventory", index=False)
            
            # Auto-adjust column widths
            for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
                col_idx = df.columns.get_loc(column) + 1
                writer.sheets["Document Inventory"].column_dimensions[chr(64 + col_idx)].width = column_width
        
        print(f"Master inventory report created with {len(report_data)} project entries")
    else:
        print("No data to include in master report")

def generate_country_reports(country_data, country_folder, date_str):
    """Generate country-specific reports for each country folder"""
    print("\nGenerating country-specific reports...")
    
    country_count = 0
    
    for country_name, projects in country_data.items():
        if not projects:
            continue
            
        # Create a safe folder name (matches the folder structure)
        safe_country = country_name.replace(' ', '_')
        country_path = os.path.join(country_folder, safe_country)
        
        # Skip if the folder doesn't exist
        if not os.path.exists(country_path):
            continue
            
        country_count += 1
        
        # Create report data - two sections
        # 1. Project IDs with document counts by type
        pid_data = []
        # 2. All documents with details
        doc_data = []
        
        # First, aggregate document types by project ID
        project_doc_counts = {}
        
        for project_id, info in projects.items():
            is_pid = re.match(r'^P\d{6}$', project_id)
            
            if is_pid:
                # Count documents by type for this project
                doc_type_counts = {}
                for file_info in info["files"]:
                    doc_type = file_info["document_type"]
                    doc_type_counts[doc_type] = doc_type_counts.get(doc_type, 0) + 1
                
                project_doc_counts[project_id] = {
                    "Project ID": project_id,
                    "Total Documents": len(info["files"]),
                    **doc_type_counts
                }
            
            # For each file, create a row in the documents table
            for file_info in info["files"]:
                doc_data.append({
                    "Project ID" if is_pid else "Filename": project_id if is_pid else os.path.basename(file_info["filename"]),
                    "Document Type": file_info["document_type"],
                    "Filename": os.path.basename(file_info["filename"])
                })
        
        # Create project ID summary data
        for pid, counts in project_doc_counts.items():
            pid_data.append(counts)
        
        # Only create report if we have data
        if pid_data or doc_data:
            # Create Excel writer
            output_file = os.path.join(country_path, f"{safe_country}_documents_{date_str}.xlsx")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # First sheet: Project ID summary
                if pid_data:
                    pid_df = pd.DataFrame(pid_data)
                    pid_df = pid_df.sort_values(by="Project ID")
                    pid_df.to_excel(writer, sheet_name="Project Summary", index=False)
                    
                    # Auto-adjust column widths
                    for column in pid_df:
                        column_width = max(pid_df[column].astype(str).map(len).max(), len(column)) + 2
                        col_idx = pid_df.columns.get_loc(column) + 1
                        writer.sheets["Project Summary"].column_dimensions[chr(64 + col_idx)].width = column_width
                
                # Second sheet: All documents
                if doc_data:
                    doc_df = pd.DataFrame(doc_data)
                    if "Project ID" in doc_df.columns:
                        doc_df = doc_df.sort_values(by=["Project ID", "Document Type"])
                    else:
                        doc_df = doc_df.sort_values(by=["Filename", "Document Type"])
                    
                    doc_df.to_excel(writer, sheet_name="All Documents", index=False)
                    
                    # Auto-adjust column widths
                    for column in doc_df:
                        column_width = max(doc_df[column].astype(str).map(len).max(), len(column)) + 2
                        col_idx = doc_df.columns.get_loc(column) + 1
                        writer.sheets["All Documents"].column_dimensions[chr(64 + col_idx)].width = column_width
    
    print(f"Created {country_count} country-specific reports")

if __name__ == "__main__":
    # If called directly, check for required arguments
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]
        portfolio_path = sys.argv[2] if len(sys.argv) > 2 else None
        generate_reports(output_dir, portfolio_path)
    else:
        print("Please provide the output directory and optional portfolio spreadsheet path.")
        print("Usage: python generate_reports.py <output_directory> [portfolio_spreadsheet]") 
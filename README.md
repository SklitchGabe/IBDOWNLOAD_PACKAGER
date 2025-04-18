# IBDOWNLOAD_PACKAGER

## Overview
IBDOWNLOAD_PACKAGER is a comprehensive document processing system designed to organize World Bank project documents by country and document type. It automates the conversion, processing, categorization, and reporting of project documents for improved document management.

## Features
- Converts various document formats (Word, PDF, etc.) to searchable PDFs
- Automatically identifies countries associated with documents based on project ID and a portfolio spreadsheet
- Organizes documents by country and document type
- Handles project ID identification and standardization (P###### format)
- Recognizes country names in multiple languages and variants
- Supports both single-folder and multi-folder processing modes
- Generates detailed inventory reports for processed documents
- Handles file naming conflicts automatically

## Requirements
- Python 3.7 or higher
- Required Python packages:
  - pandas
  - openpyxl
  - PyPDF2
  - psutil
  - python-docx (for Word document processing)

## Installation
1. Clone this repository to your local machine
2. Install required Python packages:
   ```
   pip install pandas openpyxl PyPDF2 psutil python-docx
   ```
3. Ensure you have access to the country_variants.py file for country name recognition

## Usage

### Single-Folder Processing Mode
Process a single folder of documents with the same document type:

1. Run `python main.py`
2. Select "n" when prompted about processing multiple folders
3. Enter the document type (e.g., "icrr", "pad", "aidememoire")
4. Follow the interactive prompts to select input/output folders
5. Optionally provide a spreadsheet to map project IDs to countries

### Multi-Folder Processing Mode
Process multiple folders with different document types at once:

1. Run `python main.py`
2. Select "y" when prompted about processing multiple folders
3. Specify the number of folders to process
4. For each folder, provide the folder path and document type
5. Specify a main output directory for all processed documents
6. Optionally provide a spreadsheet to map project IDs to countries

### Using Project ID Mapping
For better country association, provide a spreadsheet (Excel or CSV) with:
- A column containing Project IDs (format: P######)
- A column containing corresponding Country names

The system will prompt you to identify these columns in your spreadsheet.

## Output Folder Structure
After processing, documents are organized as follows:

```
Output Directory/
├── Country Associated Documents/
│   ├── Country Name 1/
│   │   ├── DOCUMENT_TYPE_1/
│   │   │   ├── P123456_Country Name 1_EN_document_type_1.pdf
│   │   │   └── ...
│   │   ├── DOCUMENT_TYPE_2/
│   │   │   └── ...
│   │   └── Country Name 1_documents_YYYY-MM-DD.xlsx (report)
│   ├── Country Name 2/
│   │   └── ...
│   └── ...
├── Unknown Countries/
│   ├── UNKNOWN_file1.pdf
│   └── ...
├── Failed Conversions and Renaming/
│   ├── failed_file1.pdf
│   └── ...
└── document_inventory_YYYY-MM-DD.xlsx (master report)
```
 
## Main Components

### main.py
The main script that handles:
- User interface and argument parsing
- Document conversion and processing
- Project ID extraction and standardization
- Country detection in documents
- Process coordination for single and multi-folder modes

### reorganize_output.py
Handles file organization after initial processing:
- Sorts files into country/unknown/failed categories
- Creates country-specific subfolders
- Organizes by document type within country folders
- Handles file naming conflicts

### generate_reports.py
Creates inventory reports of processed documents:
- Master inventory of all documents
- Country-specific document inventories
- Project ID summaries with document type counts

### country_variants.py
Provides country name variants for better country detection:
- Standard English country names
- Local language variants
- Alternative spellings and historical names

## Document Processing Flow

1. **Collection**: Documents from input folder(s) are collected
2. **Conversion**: Non-PDF documents are converted to PDF format
3. **Country Detection**:
   - First attempt: Extract from project ID using mapping spreadsheet
   - Second attempt: Search for country names in document content
4. **Organization**: Files sorted into country folders and document type subfolders
5. **Reporting**: Generation of inventory reports

## Troubleshooting

### Major Cases with Issues that cannot be fixed or mitigated
- **First Project ID mentioned in document text or the filename is not the Project ID of the World Bank document**
- **No Project ID found in document text or the filename**
- **Country first detected in project text is not the country of the World Bank document**
- **No Country found in document text or the filename**

### Logs
The system generates several log files:
- `reorganization.log`: Details about file organization
- `report_generation.log`: Information about report creation
- Console output: Provides real-time processing information

## Examples

### Example 1: Single-Folder Processing

```
python main.py
> n  # Single folder mode
> icrr  # Document type
> [input folder path]
> [output folder path]
> y  # Use project mapping
> [spreadsheet path]
```

### Example 2: Multi-Folder Processing

```
python main.py
> y  # Multi-folder mode
> 3  # Process 3 folders
> [folder1 path]
> icrr
> [folder2 path]
> pad
> [folder3 path]
> esrs
> [main output folder path]
```
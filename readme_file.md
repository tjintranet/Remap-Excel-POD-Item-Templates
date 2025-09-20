# Excel Column Remapper for POD Template Manager

A web-based tool for transforming arbitrary Excel files into the standardized format required by the POD (Print-on-Demand) Template Manager system.

## Overview

This application serves as a preprocessing tool that allows users to upload Excel files with any column structure and remap them to match the exact format expected by the POD Template Manager. It includes intelligent value mapping, data validation, and configuration management for streamlined repeated use.

## Features

### Core Functionality
- **File Upload**: Simple file input supporting .xls, .xlsx, and .xlsm formats
- **Column Mapping**: Interactive interface to map source columns to required POD format
- **Value Mapping**: Smart remapping of data values to standardized POD options
- **Data Validation**: Comprehensive validation with detailed error reporting
- **Export**: Generate properly formatted Excel files ready for POD import

### Advanced Features
- **Auto-mapping**: Intelligent column matching based on name similarity
- **Configuration Management**: Save and load mapping configurations for reuse
- **Real-time Preview**: Live preview of mapped data during configuration
- **Validation Statistics**: Detailed reporting on data quality and mapping completeness

## Required POD Format

The tool maps data to this exact column structure:

1. **ISBN** (Required) - 13-digit ISBN number
2. **Title** (Required) - Book title (max 58 characters)
3. **Trim Height** (Required) - Height in mm
4. **Trim Width** (Required) - Width in mm
5. **Paper Type** (Required) - Paper specification
6. **Binding Style** (Required) - Limp or Cased
7. **Page Extent** (Required) - Number of pages
8. **Lamination** (Required) - Gloss, Matt, or None
9. **Plate Section 1** (Optional) - Format: "Insert after p[page]-[pages]pp-[paper type]"
10. **Plate Section 2** (Optional) - Same format as Plate Section 1

*Note: Spine Size is automatically calculated by the main POD application and is not included in the mapping.*

## Valid Options

### Paper Types
- Amber Preprint 80 gsm
- Woodfree 80 gsm
- Munken Print Cream 70 gsm
- Munken Print Cream 80 gsm
- Navigator 80 gsm
- LetsGo Silk 90 gsm
- Matt 115 gsm
- Holmen Book Cream 60 gsm
- Enso 70 gsm
- Holmen Bulky 52 gsm
- Holmen Book 55 gsm
- Holmen Cream 65 gsm
- Holmen Book 52 gsm
- Premium Mono 90 gsm
- Premium Colour 90 gsm

### Binding Styles
- Limp
- Cased

### Lamination Options
- Gloss
- Matt
- None

## Usage Instructions

### Step 1: Upload Excel File
1. Click "Choose File" and select your Excel file
2. The app will analyze the file and display column information

### Step 2: Map Columns
1. For each required POD field, select the corresponding column from your Excel file
2. The app will attempt to auto-map columns based on name similarity
3. Preview data will show sample values from your selected columns

### Step 3: Map Values (Optional)
1. For Paper Type, Binding Style, and Lamination columns, map your data values to standard POD options
2. Use "Auto-map similar values" for automatic matching
3. Manually adjust mappings as needed

### Step 4: Validate and Export
1. Click "Validate Data" to check for errors and warnings
2. Review validation results and fix any issues
3. Click "Export Remapped Excel File" to download the transformed file

## Configuration Management

### Saving Configurations
1. Once you've set up column and value mappings, click "Save Config"
2. Enter a descriptive name and optional description
3. Download the .json configuration file

### Loading Configurations
1. Click "Load Config" and select a previously saved .json file
2. The app will automatically apply all saved mappings
3. Mappings that don't match current data will be highlighted

This feature is particularly useful for:
- Regular processing of files from the same source
- Standardizing team workflows
- Maintaining consistent mapping rules across projects

## File Structure

```
excel-remapper/
├── index.html          # Main application interface
├── script.js           # Application logic and functionality
└── README.md           # This documentation
```

## Technical Requirements

### Dependencies
- Bootstrap 5.3.2 (CSS framework)
- Bootstrap Icons 1.11.3 (Icon library)
- SheetJS/xlsx 0.18.5 (Excel file processing)

### Browser Support
- Modern browsers with ES6+ support
- Local file access capabilities
- No server required - runs entirely client-side

## Data Processing Notes

### Validation Rules
- **ISBN**: Must be exactly 13 digits
- **Title**: Flagged as error if over 58 characters (not auto-truncated)
- **Dimensions**: Must be positive numbers
- **Paper Type**: Must match approved POD specifications
- **Binding Style**: Must be "Limp" or "Cased"
- **Lamination**: Must be "Gloss", "Matt", or "None"

### Export Behavior
- Maintains original data integrity (no automatic truncation)
- Applies value mappings if enabled
- Preserves data validation flags for user review
- Generates timestamped filenames

## Troubleshooting

### Common Issues

**Column not auto-mapping correctly:**
- Manually select the correct source column from the dropdown
- Save a configuration to reuse the mapping

**Value mapping not working:**
- Ensure your data values exactly match or use the auto-map feature
- Check for extra spaces or different capitalization

**Validation errors:**
- Review the detailed error messages in the validation section
- Fix source data or adjust mappings as needed

**Configuration not loading:**
- Ensure the .json file is a valid POD mapping configuration
- Check that column names in the config match your current data

### Support
This tool processes data entirely in your browser - no data is sent to external servers. All file processing happens locally for privacy and security.

## Version History

- **v1.0**: Initial release with core mapping and validation functionality
- **v1.1**: Added value mapping capabilities
- **v1.2**: Added configuration management system

---

*This tool is designed to work with the POD Template Manager system. Ensure exported files meet all POD requirements before importing into the main application.*
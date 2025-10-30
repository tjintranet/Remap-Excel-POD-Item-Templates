# Excel Column Remapper for POD Template Manager

A web-based tool for transforming arbitrary Excel files into the standardized format required by the POD (Print-on-Demand) Template Manager system.

## Overview

This application serves as a preprocessing tool that allows users to upload Excel files with any column structure and remap them to match the exact format expected by the POD Template Manager. It includes intelligent value mapping, conditional paper type mapping based on product type, data validation, fixed value assignment, and configuration management for streamlined repeated use.

## Features

### Core Functionality
- **File Upload**: Simple file input supporting .xls, .xlsx, and .xlsm formats
- **Column Mapping**: Interactive interface to map source columns to required POD format
- **Fixed Value Assignment**: Set fixed values for columns that don't exist in source data
- **Value Mapping**: Smart remapping of data values to standardized POD options
- **Conditional Paper Type Mapping**: Map different paper types based on Product Type (Standard vs Premium)
- **Bulk Premium Paper Setting**: One-click option to set all Premium papers to the same value
- **Data Validation**: Comprehensive validation with detailed error reporting
- **Export**: Generate properly formatted Excel files ready for POD import with correct ISBN formatting

### Advanced Features
- **Auto-mapping**: Intelligent column matching based on name similarity
- **Configuration Management**: Save and load mapping configurations for reuse
- **Real-time Preview**: Live preview of mapped data during configuration
- **Validation Statistics**: Detailed reporting on data quality and mapping completeness
- **Number Formatting**: ISBN values exported as numbers with no decimals
- **Product Type Detection**: Automatically handles files with mixed Standard and Premium products

## Required POD Format

The tool maps data to this exact column structure for export:

1. **ISBN** (Required) - 13-digit ISBN number (exported as number format)
2. **Title** (Required) - Book title (max 58 characters)
3. **Trim Height** (Required) - Height in mm
4. **Trim Width** (Required) - Width in mm
5. **Paper Type** (Required) - Paper specification
6. **Binding Style** (Required) - Limp or Cased
7. **Page Extent** (Required) - Number of pages
8. **Lamination** (Required) - Gloss, Matt, or None
9. **Plate Section 1** (Optional) - Format: "Insert after p[page]-[pages]pp-[paper type]"
10. **Plate Section 2** (Optional) - Same format as Plate Section 1

**Note**: Product Type can be mapped internally for conditional paper type mapping but is NOT included in the exported file. Spine Size is automatically calculated by the main POD application and is not included in the mapping.

## Valid Options

### Product Types (Internal Use Only - Not Exported)
- **Standard** - Used for conditional mapping to standard paper types
- **Premium** - Used for conditional mapping to premium paper types

### Paper Types

#### Standard Product Papers
- Amber Preprint 80 gsm
- Woodfree 80 gsm
- Munken Print Cream 70 gsm
- LetsGo Silk 90 gsm
- Matt 115 gsm
- Holmen Book Cream 60 gsm
- Enso 70 gsm
- HolmenBulky 52 gsm
- HolmenBook 55 gsm
- HolmenCream 65 gsm
- HolmenBook 52 gsm
- Navigator 80 gsm
- MunkenCream 80 gsm
- Clairjet 90 gsm
- Magno 90 gsm

#### Premium Product Papers
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
3. **If your file has a Product Type column** (with values like "Standard"/"Premium"), map it to the "Product Type" field - this enables conditional paper mapping
4. If a column doesn't exist in your source data, select "Use fixed value..." to assign a constant value to all rows
5. Preview data will show sample values from your selected columns or display fixed values

### Step 3: Map Values (Optional but Recommended)

#### Product Type Value Mapping
If you mapped a Product Type column, first map your source values to Standard/Premium:
- Map your source "Standard" values → "Standard"
- Map your source "Premium" values → "Premium"

#### Conditional Paper Type Mapping
When both Product Type AND Paper Type are mapped, you'll see the **Paper Type Conditional Mapping** section with:

**For Standard Products:**
- Map each of your paper type values to the appropriate standard paper stocks
- Example: "White 80gsm" → "Navigator 80 gsm", "Cream 80gsm" → "MunkenCream 80 gsm"

**For Premium Products:**
- Use the **"Quick Set"** feature: Select a premium paper (e.g., "Magno 90 gsm") and click "Apply" to automatically set ALL premium papers to that value
- OR manually map each paper type individually if different premium papers are needed

**If you don't have Product Type mapped:**
- You'll see standard Paper Type value mapping where you map each paper once

#### Other Value Mappings
- Map Binding Style values (if needed)
- Map Lamination values (if needed)
- Use "Auto-map similar values" for automatic matching

### Step 4: Validate and Export
1. Click "Validate Data" to check for errors and warnings
2. Review validation results and fix any issues
3. Click "Export Remapped Excel File" to download the transformed file
4. **The exported file will NOT include the Product Type column** - it's only used internally for conditional mapping
5. ISBN values will be automatically formatted as numbers with no decimals in the exported file

## Conditional Paper Type Mapping Explained

This powerful feature allows you to use different paper types for Standard vs Premium products within the same file.

### How It Works

**Scenario**: Your source file has:
- A "Product Type" column with values: "Standard", "Premium"
- A "Paper Type" column with values: "White 80gsm", "Cream 80gsm", "Matte Coated 90gsm"

**Setup**:
1. Map both columns in Step 2
2. In Step 3, map Product Type values to Standard/Premium
3. In the Conditional Paper Type Mapping section:
   - **Standard section**: Map "White 80gsm" → "Navigator 80 gsm", "Cream 80gsm" → "MunkenCream 80 gsm", "Matte Coated 90gsm" → "Clairjet 90 gsm"
   - **Premium section**: Use Quick Set to apply "Magno 90 gsm" to all three

**Result in Export**:
- Row with Standard + White 80gsm → Paper Type becomes "Navigator 80 gsm"
- Row with Premium + White 80gsm → Paper Type becomes "Magno 90 gsm"
- Row with Standard + Matte Coated 90gsm → Paper Type becomes "Clairjet 90 gsm"
- Row with Premium + Matte Coated 90gsm → Paper Type becomes "Magno 90 gsm"

**The exported file only shows the final Paper Type** - Product Type is not included in the export.

### Quick Set Feature

For Premium products, the **Quick Set** feature makes mapping foolproof:
1. In the "When Product Type = Premium" section, locate the blue "Quick Set" box
2. Select your desired premium paper from the dropdown (e.g., "Premium Colour 90 gsm")
3. Click "Apply"
4. All paper type values for Premium products are instantly set to your selection

This eliminates the risk of missing a mapping and ensures consistency across all Premium products.

## Using Fixed Values

When your source Excel file doesn't contain a column you need to map, you can use fixed values:

1. In the mapping dropdown, select **"Use fixed value..."**
2. An input field will appear where you can enter the value
3. This value will be applied to **all rows** in the exported file
4. Fixed values are shown with a yellow badge in previews

**Example Use Cases:**
- All books have the same binding style (e.g., "Limp")
- All books use the same paper type when not using conditional mapping
- Setting a default lamination for all products (e.g., "Matt")
- Adding a constant trim size when not present in source data

## Configuration Management

### Saving Configurations
1. Once you've set up column mappings, fixed values, value mappings, and conditional paper mappings, click "Save Config"
2. Enter a descriptive name and optional description
3. Download the .json configuration file

**Configurations save:**
- All column mappings
- All fixed values
- All value mappings (including Product Type mappings)
- All conditional paper mappings (Standard and Premium)
- Export settings

### Loading Configurations
1. Click "Load Config" and select a previously saved .json file
2. The app will automatically apply all saved settings
3. Mappings that don't match current data will be highlighted

**This feature is particularly useful for:**
- Regular processing of files from the same source with consistent Standard/Premium products
- Standardizing team workflows with reusable conditional mappings
- Maintaining consistent mapping rules across projects
- Quickly setting up complex conditional mappings without manual configuration

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
- **Product Type**: If mapped, must be "Standard" or "Premium" (internal use only)
- **Paper Type**: Must match approved POD specifications
- **Dimensions**: Must be positive numbers
- **Binding Style**: Must be "Limp" or "Cased"
- **Lamination**: Must be "Gloss", "Matt", or "None"

### Export Behavior
- Maintains original data integrity (no automatic truncation)
- Applies value mappings if enabled
- Applies conditional paper mappings based on Product Type
- Applies fixed values to all rows where configured
- **Product Type column is NOT exported** - only used for conditional mapping logic
- **ISBN formatted as number with no decimals** for proper Excel display
- Preserves data validation flags for user review
- Generates timestamped filenames

### Conditional Mapping Logic
1. If Product Type is mapped and conditional paper mappings exist, uses conditional mapping
2. Falls back to regular Paper Type value mapping if no conditional mapping exists
3. If neither exists, keeps original source value

## Troubleshooting

### Common Issues

**Column not auto-mapping correctly:**
- Manually select the correct source column from the dropdown
- Save a configuration to reuse the mapping

**Need to add a column that doesn't exist in source data:**
- Use "Use fixed value..." option in the mapping dropdown
- Enter the value that should appear in all rows

**Value mapping not working:**
- Ensure your data values exactly match or use the auto-map feature
- Check for extra spaces or different capitalization

**Conditional paper mapping not applying:**
- Make sure both Product Type AND Paper Type columns are mapped
- Verify Product Type values are mapped to "Standard" or "Premium" in the Product Type Value Mapping section
- Use the Quick Set feature for Premium products to ensure all papers are mapped
- Check the browser console (F12) during export for debugging information

**Premium products showing wrong paper type:**
- Use the Quick Set feature in the Premium section to apply the same paper to all rows
- This ensures no Premium mappings are missed

**Product Type appearing in exported file:**
- Product Type should NOT appear in the export - if it does, please report this as a bug
- The column is used only for internal conditional mapping logic

**Validation errors:**
- Review the detailed error messages in the validation section
- Fix source data or adjust mappings as needed

**Configuration not loading:**
- Ensure the .json file is a valid POD mapping configuration
- Check that column names in the config match your current data
- Conditional paper mappings will load even if column names don't match

**ISBN showing in scientific notation:**
- This has been fixed in the latest version
- ISBNs are now exported as numbers with no decimal places

### Support
This tool processes data entirely in your browser - no data is sent to external servers. All file processing happens locally for privacy and security.

## Version History

- **v1.0**: Initial release with core mapping and validation functionality
- **v1.1**: Added value mapping capabilities
- **v1.2**: Added configuration management system
- **v1.3**: Added fixed value assignment and improved ISBN number formatting
- **v1.4**: Added Product Type field as optional internal column for conditional mapping
- **v1.5**: Added conditional paper type mapping with Quick Set feature for Premium products; Product Type excluded from export

---

*This tool is designed to work with the POD Template Manager system. Ensure exported files meet all POD requirements before importing into the main application.*

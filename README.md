## Getting Started

To use this script:
1. Install latest Tabular Editor 2 (https://github.com/TabularEditor/TabularEditor/releases)
2. Create C:/Power BI Backups folder
3. Place Config folder from Repo into C:/Power BI Backups
4. Open Powershell, copy text from Final PS Script and paste. No Admin Access is required and the script will check and install any modules required (at the user level, not admin).
5. Sign into Power BI via Pop-up.
6. Wait for script to complete. If the script goes beyond 45 minutes (dependent on the number of workspaces, models, and reports), a new sign-in will pop up to avoid an expired token. 
7. Open Power BI Governance Model.pbit and the model will refresh with the new data. Save as PBIX


## Features

### 1. Workspace and Metadata Extraction
- Retrieves information about Power BI workspaces, datasets, data sources, reports, report pages, and apps.
- Exports the extracted metadata into a structured Excel workbook with separate worksheets for each entity.

### 2. Model Backup and Export
- Checks for existing model backups and exports models from specified datasets using Tabular Editor.
- Saves exported models in a structured folder hierarchy based on workspace and dataset names.

### 3. Report Backup and Extraction
- Backs up reports from Power BI workspaces, cleaning report names and determining file types (`.pbix` or `.rdl`) for export.
- Extracts backed-up reports to generate `.bim` files using `pbi-tools`, ensuring proper model serialization.

### 4. Visual Object Layer (VOL) Processing (credit to m-kovalsky for initial work on this)
- Executes a predefined Tabular Editor script to extract Visual Object Layers (VOL) from reports.
- Consolidates VOL extraction results into an Excel workbook for further analysis.

### 5. Intermediate Cleanup
- Cleans temporary extraction folders and intermediate files to maintain directory hygiene.

### 6. Model Processing
- Runs additional processing scripts on exported `.bim` files using Tabular Editor.
- Aggregates CSV outputs into a single `ModelDetail.xlsx` workbook, consolidating details about semantic models and measure dependencies.

### 7. Dataflow Extraction
- Extracts dataflows from Power BI workspaces, formatting and organizing their contents, including query details.
- Exports dataflow details into a structured Excel file.

### 8. Dataflow Master File Update
- Combines newly exported dataflow details with existing records in a master `DataflowDetail.xlsx` file.
- Ensures the master file reflects the latest dataflow information across all relevant workspaces.

### 9. Token Management
- Manages Power BI API authentication by refreshing tokens in a background job.
- Cleans up and stops background token refresh jobs upon script completion.

### 10. Comprehensive Cleanup and Finalization
- Removes temporary files and folders after processing.
- Saves all generated outputs, including Excel workbooks and model files, to designated locations for future reference.

## Getting Started

To use this script:
1. Clone this repository.
2. Update the script variables (e.g., `baseFolderPath`, workspace IDs) to suit your environment.
3. Run the script in a PowerShell environment with necessary permissions.

## Prerequisites

- Power BI Management PowerShell module (`MicrosoftPowerBIMgmt`)
- Tabular Editor
- `pbi-tools`
- Excel Export module (`ImportExcel`)

## Output

- **Excel Files**: Detailed metadata, model details, VOL extractions, and dataflow information.
- **Model Files**: Serialized `.bim` files for semantic models.
- **Log Files**: Execution logs for tracking progress and debugging.

## License

This script is open-source and available under the [MIT License](LICENSE).

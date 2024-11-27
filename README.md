## This all-in-one solution is designed to be ran by anyone who uses Power BI and will work on any computer, no matter the permissions. 
- Everything within the script is limited to your access within the Power BI environment ('My Workspace') is not included. 
- All computer requirements are at the user level and do not require admin privileges. 

## Getting Started

### Instructions:
1. Install latest Tabular Editor 2 (https://github.com/TabularEditor/TabularEditor/releases)
2. Create a new folder on your C: drive called 'Power BI Backups' (C:/Power BI Backups)
3. Place Config folder and the contents from Repo into C:/Power BI Backups
4. Run 'Final PS Script' within PowerShell (either via copy/paste or renaming format to .ps1 and executing)
5. Once complete, open 'Power BI Governance Model.pbit' and the model will refresh with your data. All relationships, Visuals, and Measures are set up. Save as PBIX.


** If any modules are required, PowerShell will request to install (user level, no admin access required) **


## Features

### 1. Workspace and Metadata Extraction
- Retrieves information about Power BI workspaces, datasets, data sources, reports, report pages, and apps.
- Exports the extracted metadata into a structured Excel workbook with separate worksheets for each entity.

### 2. Model Backup and Metadata Extract
- Saves exported models in a structured folder hierarchy based on workspace and dataset names.
- Leverages Tabular Editor 2 and C# to extract the metadata and output within an Excel File.

### 3. Report Backup and Metadata Extract
- Backs up reports from Power BI workspaces, cleaning report names and determining file types (`.pbix` or `.rdl`) for export.
- Leverages Tabular Editor 2 and C# to extract the Visual Object Layer metadata and output within an Excel File (credit to m-kovalsky for initial work on this)

### 4. Dataflow Backup and Metadata Extract
- Extracts dataflows from Power BI workspaces, formatting and organizing their contents, including query details.
- Leverages PowerShell to parse and extract the metadata and output within an Excel File.
  
### 5. Power BI Governance Model
- Combines extracts into a Semantic Model to allow easy exploring, impact analysis, and governance of all Power BI Reports, Models, and Dataflows across all Workspaces

## Getting Started

To use this script:
1. Install latest Tabular Editor 2 (https://github.com/TabularEditor/TabularEditor/releases)
2. Create C:/Power BI Backups folder
3. Place Config folder from Repo into C:/Power BI Backups
4. Run Final PS Script wihtin PowerShell
5. Sign into Power BI when pop-up appears
7. Open Power BI Governance Model.pbit and the model will refresh with the new data. Save as PBIX


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

## This all-in-one solution is designed to be ran by anyone. 
- Everything within the script is limited to your access within the Power BI environment.
- All computer requirements are at the user level and do not require admin privileges.

# Power BI Governance & Impact Analysis Solution

## What It Does
This provides a quick and automated way to identify where and how specific fields, measures, and tables are used across Power BI reports in all workspaces by analyzing the visual object layer. It also backs up and breaks down the details of your models, reports, and dataflows for easy review, giving you an all-in-one **Power BI Governance** solution.

### Key Features:
- **Impact Analysis**: Fully understand the downstream impact of data model changes, ensuring you don’t accidentally break visuals or dashboards—especially when reports connected to a model span multiple workspaces.
- **Comprehensive Environment Overview**: Gain a clear, detailed view of your entire Power BI environment, including complete breakdowns of your models, reports, and dataflows and their dependencies. 
- **Backup Solution**: Automatically backs up every model, report, and dataflow for safekeeping.
- **User-Friendly Output**: Results are presented in a Power BI model, making them easy to explore, analyze, and share with your team.

 
 
## Getting Started

### Instructions:
1. Install latest Tabular Editor 2 (https://github.com/TabularEditor/TabularEditor/releases)
2. Create a new folder on your C: drive called 'Power BI Backups' (C:/Power BI Backups)
3. Place Config folder and the contents from Repo into C:/Power BI Backups
4. Open PowerShell and run 'Final PS Script' (either via copy/paste or renaming the format from .txt to .ps1 and executing)
5. Once complete, open 'Power BI Governance Model.pbit' and the model will refresh with your data. All relationships, Visuals, and Measures are set up. Save as PBIX.


- ** If any modules are required, PowerShell will request to install (user level, no admin access required) **

- ** If Tabular Editor is not installed (or can't be), the final Governance model and report/model metadata extraction won't work - BUT the workspace metadata extraction, all backups, and the dataflow metadata extraction will still work **


## Features

### 1. Workspace and Metadata Extraction
- Leverages Power BI REST API to gather information about Power BI workspaces, datasets, data sources, reports, report pages, and apps.
- Exports the extracted metadata into a structured Excel workbook with separate worksheets for each entity.
- You must have at least read acess within workspaces. 'My Workspace' not included.
- <img width="1255" alt="image" src="https://github.com/user-attachments/assets/515ce3e5-ec56-467a-a421-9da05889eaa5">


### 2. Model Backup and Metadata Extract
- Saves exported models in a structured folder hierarchy based on workspace and dataset names.
- Leverages Tabular Editor 2 and C# to extract the metadata and output within an Excel File.
- All backups are saved with the following format: Workspace Name ~ Model Name.
- You must have edit rights on the related model. Works with all Pro, Premium Capacity, Fabric Capacity workspaces. Both XMLA and non-XMLA models. 'My Workspace' not included.
<img width="695" alt="image" src="https://github.com/user-attachments/assets/c3e021b8-6dfe-40c9-bfa5-b9d4471a8fa3">


### 3. Report Backup and Metadata Extract
- Backs up Power BI and Paginated Reports from Power BI workspaces, cleaning report names and determining file types (`.pbix` or `.rdl`) for export.
- Leverages Tabular Editor 2 and C# to extract the Visual Object Layer metadata and output within an Excel File (credit to @m-kovalsky for initial work on this)
- Paginated Reports are only backed up (no metadata extraction).
- All backups are saved with the following format: Workspace Name ~ Report Name.
- You must have edit rights on the related report. Works with all Pro, Premium Capacity, Fabric Capacity workspaces. 'My Workspace' not included.
- <img width="554" alt="image" src="https://github.com/user-attachments/assets/cf88aac7-6f32-445a-96c7-6bc36fcab9aa">


### 4. Dataflow Backup and Metadata Extract
- Extracts dataflows from Power BI workspaces, formatting and organizing their contents, including query details.
- Leverages PowerShell to parse and extract the metadata and output within an Excel File.
- All backups are saved with the following format: Workspace Name ~ Dataflow Name.
- Must have edit rights on the related dataflow. 'Ownership' of the Dataflow is not required. Works with all Pro, Premium Capacity, Fabric Capacity workspaces. 'My Workspace' not included.
- <img width="542" alt="image" src="https://github.com/user-attachments/assets/67e83016-4bc7-4cf5-8d94-1a9779aad6d8">

  
### 5. Power BI Governance Model
- Combines extracts into a Semantic Model to allow easy exploring, impact analysis, and governance of all Power BI Reports, Models, and Dataflows across all Workspaces
- Works for anyone who runs the script and has at least 1 model and report. Dataflow not required.
- Public example (limited due to no filter pane): https://app.powerbi.com/view?r=eyJrIjoiN2Y3OTZlOTEtOWQ4Yi00M2UyLTk3MGQtMDA4OTRjY2M4ZGJlIiwidCI6ImUyY2Y4N2QyLTYxMjktNGExYS1iZTczLTEzOGQyY2Y5OGJlMiJ9)

## Special Notes
- The script has a built-in timer to ensure the API bearer token does not expire. It is defaulted to require logging in every 55 minutes. This is only applicable if you have a large number of reports and models (150+)
- This defaults to looping across all workspaces. If you only want to run this for a specific workspace, you can enter a workspace ID within the quotation marks in $reportSpecificWorkspaceId and/or $modelSpecificWorkspaceId (these are in the first 20 lines of the script)
..
..

<img width="1235" alt="image" src="https://github.com/user-attachments/assets/33101f84-b567-4a45-9729-09303eeb50fb">
<img width="1259" alt="image" src="https://github.com/user-attachments/assets/87d23e7e-5f9b-4883-8c58-f102033be5e0">
<img width="1221" alt="image" src="https://github.com/user-attachments/assets/e120c1bb-b52a-4197-aeb3-2a6ddbb67a9f">
<img width="1241" alt="image" src="https://github.com/user-attachments/assets/9d814034-494d-478b-b231-f759d7eebeab">

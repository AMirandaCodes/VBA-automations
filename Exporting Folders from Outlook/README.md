# Exporting Folders from Outlook

## Project Brief
The company requires an automated solution to export all Outlook Public Folders to a local drive. Public Folders are shared email storage locations accessible to all users within the organisation through Microsoft Outlook. This one-off extraction is for data backup and organisation.

## Project Requirements
- **Folder structure**: The original structure of the Public Folders must remain intact in the exported location.
- **Email format**: All emails must be exported in `.msg` format to preserve metadata and compatibility.
- **Scope**: All emails across all Public Folders must be included in the extraction.
- **Execution**: This is a one-off extraction.

## Project Deliverables
- A VBA code solution/s capable of:
  1. Extracting all Public Folders to a desired location on the local drive.
  2. Preserving the original folder structure.
  3. Exporting emails in `.msg` format.

## Instructions
1. Open Outlook and ensure access to all necessary Public Folders.
2. Run the VBA script following the provided instructions.
3. Verify the exported files in the target location.

# Code 1.1: Exporting All Folders to a Local Drive

## Code Description
This VBA script automates the export of a selected Outlook folder, along with its subfolders and emails, to a specified location on the local drive. Each email is saved in `.msg` format with its received date prepended to the filename. The script ensures that the folder structure remains intact and generates an error log for any issues encountered during the process.

## Code Features
- **Folder Selection**: Allows the user to choose an Outlook folder for export.
- **Destination Selection**: Lets the user specify a target location on the local drive.
- **Recursive Export**: Processes all subfolders within the selected folder.
- **Email Naming**: Prepends the received date to the filename for easy sorting.
- **Error Logging**: Creates a detailed log file to capture issues during execution.

## Code Usage Instructions
1. Open Microsoft Outlook.
2. Run the VBA script from the VBA editor.
3. Select the folder you wish to export in the prompt.
4. Choose a destination folder on your local drive.
5. Check the selected destination folder for the exported emails and the error log (`EmailExport_Log.txt`).

![image](https://github.com/user-attachments/assets/1ac82252-c16c-45ef-9cd3-2375eed8c07b)


## Limitations
- This script is designed for one-off exports. For repetitive tasks, consider modifying the code or creating a scheduled macro.
- Long file paths or filenames exceeding the operating system's limit (255 characters) may cause errors.



# Code 1.2: Extracting Email Metadata of a Folder (Subject, Received Date, Sender) to an Excel File

## Code Description
This VBA script automates the extraction of email metadata from a selected Outlook folder. It retrieves the **Subject**, **Received Date (with time)**, and **Sender Name** of each email and exports this information into an Excel file. The script creates the Excel file in the user’s Desktop and names it after the selected Outlook folder.

## Code Features
- **Folder Selection**: Prompts the user to choose an Outlook folder for metadata extraction.
- **Desktop as Destination**: Automatically saves the Excel file on the user’s Desktop.
- **Metadata Details**: Each email's **Subject**, **Received Date**, and **Sender Name** are included in the Excel file.
- **Naming Consistency**: The Excel file is named after the selected folder, with illegal characters stripped for compatibility.

![image](https://github.com/user-attachments/assets/7382df08-ccd6-4a77-be83-14d8bf067bae)

## Code Prerequisites
- **Microsoft Outlook**: Ensure that Outlook is installed and configured on your system.
- **Microsoft Excel**: Ensure that Excel is installed to create and save the output file.
- **Folder Access**: Verify access to the desired Outlook folder.
- **Sufficient Storage**: Ensure adequate disk space for the Excel file.

## Code Usage Instructions
1. Open Microsoft Outlook.
2. Access the VBA editor and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the Outlook folder you wish to extract metadata from when prompted.
5. Check your Desktop for the Excel file, which will have the same name as the selected Outlook folder.

## Limitations
- **Single Folder Processing**: The script only processes the selected folder. Subfolders are not included in this version.
- **One-Time Execution**: Designed for one-off extractions. For repetitive tasks, consider automating the script further or creating scheduled macros.

# Code 1.3: Extracting Email Metadata of Subfolders (Subject, Received Date, Sender) to Excel Files on Desktop

## Code Description
This VBA script automates the extraction of email metadata (Subject, Received Date with time, and Sender Name) from all subfolders of a selected Outlook folder. For each subfolder, an individual Excel file is created and saved in a designated "Email Info" folder on the user’s Desktop.

## Code Features
- **Folder Selection**: Enables the user to choose an Outlook folder for subfolder metadata extraction.
- **Desktop Destination**: Automatically creates a "Email Info" folder on the Desktop to store the Excel files.
- **Recursive Extraction**: Processes all subfolders within the selected folder.
- **Detailed Metadata**: Captures each email’s **Subject**, **Received Date** (including time), and **Sender Name**.
- **Completion Notification**: Displays a confirmation message once the extraction is complete.

![image](https://github.com/user-attachments/assets/899bd407-7b5f-432a-a93f-a78c1b590b5a)

## Code Prerequisites
- **Microsoft Outlook**: Ensure Outlook is installed and configured on your system.
- **Folder Access**: Verify access to the desired folders in Outlook.
- **Sufficient Storage**: Ensure adequate disk space for the exported Excel files.

## Code Usage Instructions
1. Open Microsoft Outlook.
2. Access the VBA editor and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the Outlook folder whose subfolders you wish to extract email metadata from.
5. Check your Desktop for the "Email Info" folder containing the Excel files.

## Limitations
- **No Metadata for Selected Folder**: Only subfolders are processed; emails in the selected folder itself are not included.
- **One-Off Execution**: Designed for single-use exports.

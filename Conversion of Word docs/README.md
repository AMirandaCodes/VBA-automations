# Conversion of Word documents to the latest Word format

## Project Brief
The company has old Word documents, Word version 97-2003 (`.doc` files), and requires an automated solution to convert any historical files to the latest Word format (`.docx`) which will enable the latest Word's features.

## Project Requirements
- **Contents unchanged**: The original contents and structure of the Word documents must remain unchanged after convertion.
- **All in `.docx` format**: All converted word documents must be `.docx`.
- **Process an entire folder**: The solution should be able to process a specific folder, and should run with folders that have a combination of old or current Word formats.

## Project Deliverables
- A VBA code solution capable of:
  1. Converting any Word documents to `.docx`.
  2. Preserving the original Word document structure and filename.

### Code Description
This VBA script automates the process of converting all Word documents in a selected folder from the `.doc` format to the latest `.docx` format. The original `.doc` files are replaced by the newly converted `.docx` versions.

### Code Features
- **Folder Selection**: Prompts the user to select a folder containing the Word documents for conversion.
- **Same Destination**: The converted `.docx` files are saved in the same folder as the original `.doc` files.
- **Automatic Replacement**: Deletes the original `.doc` files after successful conversion, ensuring no duplicates.

![image](https://github.com/user-attachments/assets/227efdc5-724a-4e6f-b40b-58ffca445dd7)

### Code Usage Instructions
1. Open Microsoft Word.
2. Access the VBA editor (Alt + F11) and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the folder containing the `.doc` files when prompted.
5. The converted `.docx` files will replace the originals in the same folder.

### Limitations
- **One-Off Execution**: Designed for single-use conversions. For repeated or scheduled tasks, consider automation enhancements.
- **Layout Adjustments**: Minor layout changes may occur during the conversion process, especially with older documents containing complex formatting.
- **File Deletion**: The script deletes original `.doc` files after conversion. Ensure you have backups if required.

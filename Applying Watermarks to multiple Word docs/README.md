# Project 3: Applying a Watermark Automatically to Multiple Word Documents

## Project Brief
The company requires a one-off automated solution to apply a custom watermark to multiple Word documents. Moving forward, users will manually apply the watermark to any new documents as needed. 

## Project Requirements
- **Preservation of Content and Format**: The original content and format of the documents must remain unchanged after the watermark is applied.
- **Consistent Watermark**: Only the specified watermark should be applied across all documents.
- **Standard Word Watermark**: The watermark must be implemented using Word's built-in watermark feature (not as a Building Block or alternative approach), ensuring compatibility with Word's native watermark settings.

## Project Deliverables
- A VBA code solution that can:
  1. Open Word files from a specified folder.
  2. Apply the specified watermark to each document.
  3. Save the changes and close the documents automatically.

# Code 1: Applying a Watermark Automatically to Multiple Word Documents

## Code Description
This VBA script automatically applies a specified watermark to the Word documents of a predefined folder.

## Code Features
- **Pre-set watermark and folder selection**: The code need to be updated with the custom fields of the user's watermark and folder, then when the code is ran it automatically applies the watermark as instructed (so the user is not prompted to choose a location)

## Code Usage Instructions
1. Open Word.
2. Open the VBA script from the VBA editor.
3. Replace the following parts of the code with your custom names:
- Replace `"C:\Path\To\Your\Documents\"` with the folder path containing your Word documents.
- Replace `"Watermark Name"` with the exact name of your watermark as saved in the gallery.
- Replace `"C:\Path\To\Watermark Template"` with the full path to your watermark template.
4. Run the VBA script.
5. Check that the documents in that folder have the watermark applied.

## Limitations
- This script is designed to work with the watermark name, watermark template and watermark template file path that is specified in the code. If any of these change, the code needs to be updated.

# dnd_convert_character_pdf_to_sheet
A utility to migrate an exported DnD character sheet to a Google Docs Sheet

## Directions
1. Contact me to get a copy of the "5E Character Sheet 2024 (Manual)" Google Sheet
1. Follow the Google Cloud setup directions in this document: https://spreadsheetpoint.com/connect-python-and-google-sheets-15-minute-guide/
1. Rename the `credentials.json` file to `google-credentials.json`.
1. Go to DNDBeyond and export your charater sheet as a PDF and copy that file into this folder
1. Create a virtual python environment, activate it and run `pip install -r requirements.txt`
1. Run `python dndbeyond_pdf_to_sheet.py --character_sheet [sheet name] --google_sheet_name [google sheet name]`

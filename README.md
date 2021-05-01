# screenshot_extractor
A script created to extract full resolution images from an .xlsx file

## Instructions:
1. Read the functionality specs below and prepare your .xlsx file accordingly.
2. Double click on the screenshot extractor Python script. A file selector window will appear.
3. Select the .xlsx file you want to extract screenshots from.
4. Press enter.
5. A .zip file will be created within the same directory as the .xlsx file.
6. Unzip the file to view/edit the images.

## Functionality Specs:
- All images are exported into one .zip file.
- Keys cannot exceed 255 characters and cannot contain any illegal characters/symbols.
- If your .xlsx file was exported directly from Feishu sheets, open the .xlsx file, click on one of the sheets, and save it once using Microsoft Excel.
- To match one screenshot to multiple keys, you can either use ⌘+c/ctrl+c to copy the screenshot to another row/sheet individually OR you can adjust the height of the image (dragging the bottom or upper edge) so that the screenshot spans the keys you want it to match with.
- If key column is left empty, the exported screenshots will be named no_key1, no_key2, etc.
- Column headers should be located in row 1
- Key column should be labelled "Key", "key", "Keys", "keys" without any leading or trailing spaces
- Only use one column to record the keys.

# Document Image Access

Document Image Access tells you which images and charts a document contains, where each one is, and what its alt text says.

Without it, a screen reader user can open a document and never learn that it holds pictures. With it, you can find every picture and chart, then open it, copy it, save it, or have it described by the tool of your choice, such as the Be My Eyes app.

## What it does

- Gives a one-key summary of the images in the current document, such as "7 images, 3 without alt text".
- Shows a full list of the images and charts, with the location and the alt text of each.
- Opens, copies, saves and exports images.
- Sends an image, a chart, or a whole page or slide to Be My Eyes for a description.
- Jumps to an image in the open document.
- Writes a new alt text into the open document.

## Requirements

- NVDA 2026.1 or later. It was tested with NVDA 2026.3.1.
- Microsoft Office is optional. Reading a Word, Excel, PowerPoint or OpenDocument file that you select in File Explorer needs no Office at all. Office is needed only for the things listed under "Waiting times" and "The buttons" below that work on a document open in Word, Excel or PowerPoint.
- The Be My Eyes app for Windows is optional. It is needed only for "Send to Be My Eyes". When it is not installed, the Send button is hidden.

## Supported files

- Word: docx, docm, dotx, dotm. The older doc and dot files are converted first, using Word, or LibreOffice when Word is missing.
- PowerPoint: pptx, pptm, potx, ppsx.
- Excel: xlsx, xlsm, xltx.
- OpenDocument: odt, ods, odp, ott.

Not supported:

- The old PowerPoint and Excel formats (ppt and xls).
- PDF files.
- Password-protected files.

## The keyboard shortcut

### NVDA+Shift+I

- Press it once for a quick summary.
- Press it twice quickly to open the full list.
- It works on the document open in Word, Excel or PowerPoint, and on the first file selected in File Explorer.
- Anywhere else, a dialog lets you choose a file.
- If a document has unsaved changes, the add-on reads the last saved version and tells you so.
- If a document is stored online, or has never been saved, the add-on tells you what to do.

### Changing the shortcut

If NVDA+Shift+I conflicts with another add-on or with your own shortcuts, you can change it:

1. Open the NVDA menu, then Preferences, then Input gestures.
2. Find the category "Document Image Access".
3. Select the command, choose Add shortcut (or Remove to delete the current one), and press the keys you want.

## The quick summary

- Tells how many images the document has and how many of them have no alt text.
- Says "No images found in" the file name when there are none.
- Charts are counted separately, for example "1 chart, which is not an image."
- If the file stores pictures that the add-on cannot place, it says so, for example "2 picture files in the document could not be located."

## The image list

### What each row tells you

Every row is read as four short parts:

- The image number, for example "Image 7". Charts are "Chart 1" and so on, after the images.
- The location.
- The alt text, or "No alt text".
- The text around the image ("Near: ..."), or for a chart its title ("Title: ...").

The location is one of these:

- Word: "Page 6, line 12". If Word could not be used, you hear "Page about 6" and the line is not known.
- Word headers, footers, footnotes and endnotes: "In header", "In footer", "In footnotes", "In endnotes".
- Images in a text box: "In a text box." is added to the row.
- PowerPoint: "Slide 3".
- Excel: "Sheet Budget, cell B3".
- Excel pictures placed inside a cell: "Sheet Budget, cell B3, in cell".
- Excel sheet backgrounds and page headers or footers: "Sheet Budget, background" and "Sheet Budget, left header".
- Excel chart sheets: "Chart sheet Chart1".
- OpenDocument: the slide or the sheet, when there is one.

### Where the list starts

- When you open the list from the open document, it starts on the image closest to where you are: the cursor in Word, the active cell in Excel, the slide or the selected shape in PowerPoint.
- When nothing has moved since you closed the list, it starts on the row you left.
- When you open it from File Explorer, it starts on the row you left, or on the first row.

### Keys in the list

- Enter opens the selected image.
- Escape closes the list.
- Alt+I moves to the list.
- Each button has an access key, shown in its section below (English interface).

## The buttons

### Open (Alt+O)

- Opens the image in the program that Windows uses for that kind of picture.
- For a chart, the add-on first exports it as an image (see "Charts").

### Save as (Alt+S)

- Saves the image as a file. The suggested name carries the location, such as p003_image_007.png.

### Copy image (Alt+Y)

- Puts the image on the clipboard.
- Pictures in formats that programs often cannot paste are converted to PNG first.
- You can then paste it into any other program, or use it with another image description add-on.

### Send to Be My Eyes (Alt+B)

- Opens the Be My Eyes app with the image, so that you can ask for a description.
- Pictures in other formats are converted to PNG first, using the Pillow library when NVDA provides it.
- The button is shown only when the Be My Eyes app is found on the computer. When it is not found, the button is hidden and disabled, and the same choice is left out of the menu of "Whole page or slide". The add-on tells you so once per NVDA session.
- See "Be My Eyes" below for the details.

### Go to in document (Alt+G)

- Moves the cursor to the image in the document that is open.
- Word: the cursor goes in front of the image. If the document changed since the list was made, the position may be approximate and the add-on says so.
- Excel: the cursor goes to the cell of the image, or the chart sheet opens.
- PowerPoint: the slide opens and the image is selected.
- It works only when you opened the list from that open document.
- Images in Word headers and footers, and Excel backgrounds and page headers, cannot be reached this way.

### Edit alt text (Alt+E)

- Opens a box with the current alt text. Change it and choose OK to write it into the document.
- Word: pictures and charts in the body of the document. The document must be saved and unchanged since the list was made.
- Excel: floating pictures and charts. Not pictures inside cells, backgrounds, page headers or footers, or chart sheets.
- PowerPoint: pictures and charts.
- Save the document afterwards to keep the change.
- If the document refuses the change, the reason is spoken. See "When a change is refused".

### Whole page or slide (Alt+W)

- Makes a picture of the whole page (Word) or the whole slide (PowerPoint) that holds the selected image, so that Be My Eyes can describe the overall layout.
- Then a menu offers Open, Save as, Copy to the clipboard and, when Be My Eyes is found, Send to Be My Eyes.
- It is not available for Excel sheets. You hear a message saying so.
- In Word, the page must be known exactly. The picture shows the content of the page and may differ from the printed page.
- It works only when you opened the list from that open document.

### Export all (Alt+X)

- Saves every image to a folder that you choose.
- File names carry the location, for example p003_image_007.png, slide02_image_004.jpg or Budget_B3_image_001.png.
- Charts are exported too, when the document is open in its program.

### Close (Alt+C)

- Closes the list.

### After a button

- The result is spoken first, for example "Alt text updated.", and then the focus goes back to the list.

## Charts

- A chart is a drawing made from data, not a picture. It is not stored as an image, so a screen reader and most tools cannot see it.
- The add-on lists charts after the images, with their title and alt text.
- To make an image of a chart, the add-on asks Excel, PowerPoint or Word to export it. This works only when the document is open in that program.
- When you open the list from File Explorer, the charts are listed but cannot be exported.
- Once exported, a chart can be opened, copied, saved, or sent to Be My Eyes like any image.

## Waiting times

Some work takes a moment. While the add-on works, you hear a short beep once a second.

- The quick summary of docx, pptx, xlsx and OpenDocument files is fast, because only the file is read.
- Legacy doc files are converted first, using Word or LibreOffice. The add-on says "Converting the .doc file with Word." The summary and the list both wait for it.
- The full list of a Word document also asks a hidden copy of Word for the exact page and line of each image. It takes a few seconds, longer for big documents.
- If Word cannot be used, the pages are estimated from the page breaks stored in the file, the line numbers are not shown, and the add-on tells you so.
- The first time you use Send to Be My Eyes, finding the app may take a few seconds.

The hidden copy of Word opens the file read-only, never runs macros, and never changes your file.

## When a change is refused

Editing the alt text, and exporting a chart, can be refused by the program. The reason is spoken:

- Protected View: choose Enable Editing in the program, then open the list again.
- The document is read-only: reopen it with editing enabled.
- The document is marked as final: choose Edit Anyway.
- Editing is restricted, or the sheet is protected: the change cannot be made.
- The program is busy, for example while you are typing in an Excel cell: finish that and try again.
- The document changed since the list was made, or was never saved: save it and open the list again.

## Be My Eyes

### Installation check

- When you install or upgrade the add-on, it looks for the Be My Eyes app.
- If the app is found, nothing is asked. The add-on remembers it, so that the first "Send to Be My Eyes" starts at once.
- If the app is missing, a dialog offers four things: open the Microsoft Store page to install it, copy the link, check again after installing, and continue without it.
- The dialog has a check box "Do not remind me again".

### When the app is missing

- The Send to Be My Eyes button is hidden and disabled, and the same choice is left out of the menu of "Whole page or slide".
- The add-on says so once per NVDA session, when the list opens.
- It checks again each time the list opens. A found app is remembered for ten minutes and a missing one for 30 seconds, so an app that you install while NVDA is running is noticed the next time you open the list.
- If the check itself cannot run, the button is kept.

### Setting the app by hand

- The add-on looks for an app whose name contains "Be My Eyes".
- If your copy has another name, create or edit the file docImages.json in your NVDA configuration folder (usually %APPDATA%\nvda) and add a line such as {"aumid": "the application ID"}.
- An application ID written by hand is always trusted: the add-on does not search, and the button is shown.

## Privacy and files

- Images are extracted to a temporary folder named nvda_docImages, inside the Windows temporary folder. It is emptied when NVDA starts and when it exits.
- Nothing is sent over the network by the add-on. An image leaves your computer only when you send it to Be My Eyes, and then Be My Eyes handles it.
- The add-on does not change your documents, except when you choose Edit alt text.

## Limits

- Pictures that are linked and not stored in the file, SmartArt diagrams, drawn shapes and the Excel IMAGE function are not listed.
- A picture used as the fill of a shape is stored in the file but not placed. It is counted in "could not be located".
- Pictures marked as decorative count as "without alt text".
- Copying, sending, or picturing a whole page converts pictures in formats other than PNG, JPEG and GIF (for example metafiles and TIFF) with the Pillow library. If your NVDA does not include Pillow, you are told so, and the whole page picture of Word is not available. Opening, saving and exporting still work.
- OpenDocument files have no page numbers.
- Excel files with pictures placed inside cells may not open in older versions of Excel. The add-on can still list them when you select the file in File Explorer.
- Behavior of Word, Excel and PowerPoint differs between versions. If something does not work with your version, see "Troubleshooting".

## Troubleshooting

- "Could not reach the open document": choose the file with the dialog that follows, or select it in File Explorer.
- A button is unavailable: the list was opened from File Explorer, or the image belongs to another program than the open document.
- There is no Send to Be My Eyes button: the Be My Eyes app was not found. Install it from the Microsoft Store, then open the list again. If it is installed under another name, see "Setting the app by hand".
- Page numbers say "about": Word could not be used to place the images. The estimate comes from the page breaks stored in the file.
- Something does not work: open the NVDA log (NVDA+F1) and search for "docImages". The messages there tell which step failed. Include them when you report a problem.

## Reporting problems

Report problems and ideas on the project page: https://github.com/NVDATH/document-image-access/issues

Please say which versions of NVDA and Office you use, and include the NVDA log lines that start with "docImages".

## About this add-on

Document Image Access was developed jointly by NVDA_TH and Claude, an AI assistant made by Anthropic.

- NVDA_TH had the idea, decided how the add-on should work, and tested every version with NVDA and real documents.
- Claude wrote the code and the documentation.
- Because an AI wrote the code, errors are possible. Testing was done by one person with one setup, so behavior with other versions of Office may differ. Please report what you find.

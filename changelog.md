## 0.6.2

### New Features and Improvements

* First public release of Document Image Access. It lists the images and charts of Word, Excel, PowerPoint and OpenDocument files with their location and alt text, and lets you open, copy, save, export, describe (with Be My Eyes), jump to and relabel them. Press NVDA+Shift+I once for a summary and twice quickly for the full list. The entries below show how each part came about.
* The **Send to Be My Eyes** button is hidden and disabled when the Be My Eyes app is not installed, and the same choice is left out of the menu of **Whole page or slide**. The add-on says so once per NVDA session and checks again each time the list opens. When the check itself cannot run, the button is kept.
* An application ID written by hand in docImages.json is trusted without a search. The ID that the add-on finds itself is kept under its own name, so the two no longer mix. The installation check leaves you alone when a hand-written ID is already set.
* Requires NVDA 2026.1 or later.

### Bug Fixes

* When NVDA does not provide the Pillow library, copying or sending a picture in a format other than PNG, JPEG or GIF, and picturing a Word page, now say what is missing instead of failing without a reason.

## 0.6.1

### New Features and Improvements

* When the add-on is installed or upgraded, it looks for the Be My Eyes app. If the app is missing, a dialog offers to open its Microsoft Store page, copy the link, check again, or stop reminding you. When the app is found, its ID is remembered so that the first "Send to Be My Eyes" starts without a search.
* "Send to Be My Eyes" now offers to open the Microsoft Store page when the app is not installed. It also notices an app that was uninstalled after its ID had been saved.

## 0.6.0

### New Features and Improvements

* Charts in Word documents are listed with their page, line, title and alt text, and can be exported as images, opened, copied, saved or sent to Be My Eyes.
* New **Whole page or slide** button. It pictures the page (Word) or the slide (PowerPoint) that holds the selected image, then offers Open, Save as, Copy to the clipboard and Send to Be My Eyes, so that Be My Eyes can describe the overall layout. It is not available for Excel sheets.
* Pictures and charts that share a Word paragraph are told apart when their position is looked up.
* The result of a button is now spoken first and the focus goes back to the list afterwards, without a delay.
* PowerPoint charts fall back to the chart's own exporter when exporting the shape fails.

## 0.5.2

### New Features and Improvements

* After any button action, the focus goes back to the list so that you can keep reading it.

## 0.5.1

### New Features and Improvements

* The list starts on the image closest to where you are in the open document: the caret in Word, the active cell in Excel, the selected shape or the slide in PowerPoint. When nothing has moved since you closed the list, it starts on the row you left.
* The Edit alt text box shows the alt text as it is in the open document now, not as it was when the list was made.

### Bug Fixes

* A document open in Protected View is reported as such, with the advice to choose Enable Editing, instead of a vaguer "read-only" message.
* Exporting a chart from a document in Protected View explains the real cause instead of Excel's misleading "dimension" error.
* The alt text of charts can be changed in Excel.

## 0.5.0

### New Features and Improvements

* **Go to in document** and **Edit alt text** now work in Excel and PowerPoint. Shapes are found by the name stored in the file, so no guessing is needed. When a change is refused, the reason is spoken: read-only, protected sheet, marked as final, or the application is busy (for example while a cell is being edited).
* Charts of Excel (including chart sheets) and PowerPoint are listed with their title and alt text, and can be exported as images through the application, then opened, copied, saved or sent to Be My Eyes.

## 0.4.5

### New Features and Improvements

* Native charts are counted and announced separately, for example "1 chart, which is not an image", in Excel (including chart sheets), Word and PowerPoint.

## 0.4.4

### New Features and Improvements

* Excel pictures placed inside cells, pictures in the page header and footer, and sheet background pictures are listed.
* A picture file used many times is extracted once.
* The add-on announces how many picture files it could not locate, instead of staying silent about them.

## 0.4.3

### New Features and Improvements

* When Word refuses to change an alt text, the reason is spoken (Protected View, read-only, editing restrictions, marked as final) and the details are written to the NVDA log.
* Word's COM objects are released as soon as the list is closed, or ten seconds after a quick summary, so that Word can close normally.

## 0.4.2

### New Features and Improvements

* The open document is looked up in three steps: the objects NVDA's own Office support holds, the object model of the application window, and finally the Windows running object table. If all fail, the log tells what was tried.
* Opening a legacy .doc file announces that it is being converted by Word.

## 0.4.1

### New Features and Improvements

* Open documents in Excel and PowerPoint are found, with the same checks as Word: unsaved changes, documents stored online, and files that were never saved.
* The progress beep repeats every second instead of every half second.

### Bug Fixes

* Fixed the add-on failing to load because of a wrong import.

## 0.4.0

### New Features and Improvements

* One shortcut, two levels: press NVDA+Shift+I once for a quick summary and twice quickly for the full list.
* PowerPoint, Excel and OpenDocument files (odt, ods, odp) are supported, and Word headers, footers, footnotes and endnotes are listed.
* New buttons: **Copy image**, **Export all** (file names carry the location, such as p003_image_007.png) and **Edit alt text** (writes a new alt text into the open Word document).
* All messages can be translated.

## 0.3.0

### New Features and Improvements

* A beep repeats while the add-on is working, for example while it waits for Word.
* Every cell of the list describes itself, such as "Page 6, line 12", because NVDA does not read column headers for this kind of list.
* New **Go to in Word** button.
* The list title tells how many images have no alt text.

## 0.2.0

### New Features and Improvements

* The list is a virtual list with columns, which stays fast with many images.
* Word images show their page and line. The numbers come from a hidden copy of Word. Without Word, the page is estimated from the page breaks stored in the file.
* The document open in Word is found through NVDA's own Word object.

## 0.1.0

### New Features and Improvements

* First test build. Lists the images of a .docx file with their alt text and the text around them. Works from the document open in Word, from a file selected in File Explorer, or from a file you choose.
* Open an image or save it as a file. Send an image to the Be My Eyes app, which is found automatically and remembered.
* Legacy .doc files are converted with Word, or LibreOffice when Word is missing.
* The work runs in the background, and temporary files are removed.

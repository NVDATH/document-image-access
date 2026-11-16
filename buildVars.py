# -*- coding: UTF-8 -*-
# buildVars.py - variables used by SCons when building the addon.

def _(x):
    return x

# Add-on information variables
addon_info = {
    # The add-on ID. It cannot change after the first release: users would stop receiving updates.
    "addon_name": "docImages",
    "addon_version": "0.6.2",
    # Translators: Summary for this add-on
    "addon_summary": _("Document Image Access"),
    # Translators: Long description to be shown for this add-on on add-on information from add-ons manager
    "addon_description": _("""Document Image Access tells you which images and charts a document contains, where they are, and what their alt text says, so that you can reach them with an image description tool.

Supported files: Word (docx and legacy doc), Excel, PowerPoint and OpenDocument (odt, ods, odp).

Features:
• Press NVDA+Shift+I for a quick summary, such as "7 images, 3 without alt text". Press it twice quickly to open the full list.
• For each image: its location (page and line in Word, slide in PowerPoint, sheet and cell in Excel), its alt text and the text around it. Pictures in headers, footers, cells, backgrounds and text boxes are listed, and charts too.
• Open an image, copy it, save it, or export all images to a folder, with file names that carry the location.
• Send an image to the Be My Eyes app for a description, or the whole page or slide that holds it. The Send button appears only when Be My Eyes is installed, and the add-on offers to open its Microsoft Store page when you install it.
• Charts are exported as images through Excel, PowerPoint or Word, so that they can be described too.
• From the list, jump to the image in the open document, and write a new alt text back into it.
• Works with Microsoft Office when it is installed. Reading a file from File Explorer needs no Office at all."""),
    "addon_author": "NVDA_TH <nvdainth@gmail.com>, assisted by A.I.",
    "addon_url": "https://github.com/NVDATH/document-image-access",
    "addon_docFileName": "readme.html",
    "addon_minimumNVDAVersion": "2026.1",
    "addon_lastTestedNVDAVersion": "2026.3",
    "addon_updateChannel": "stable",
    # Translators: What's new text for this add-on version, shown in NVDA's Add-on Store.
    # NOTE: content between the CHANGELOG markers below is auto-replaced by
    # scripts/sync_changelog.py using the latest "## " section of changelog.md.
    # Do not edit by hand; edit changelog.md instead and re-run the sync script.
    "addon_changelog": _(
        # CHANGELOG-START
        """### New Features and Improvements

* First public release of Document Image Access. It lists the images and charts of Word, Excel, PowerPoint and OpenDocument files with their location and alt text, and lets you open, copy, save, export, describe (with Be My Eyes), jump to and relabel them. Press NVDA+Shift+I once for a summary and twice quickly for the full list. The entries below show how each part came about.
* The **Send to Be My Eyes** button is hidden and disabled when the Be My Eyes app is not installed, and the same choice is left out of the menu of **Whole page or slide**. The add-on says so once per NVDA session and checks again each time the list opens. When the check itself cannot run, the button is kept.
* An application ID written by hand in docImages.json is trusted without a search. The ID that the add-on finds itself is kept under its own name, so the two no longer mix. The installation check leaves you alone when a hand-written ID is already set.
* Requires NVDA 2026.1 or later.

### Bug Fixes

* When NVDA does not provide the Pillow library, copying or sending a picture in a format other than PNG, JPEG or GIF, and picturing a Word page, now say what is missing instead of failing without a reason."""
        # CHANGELOG-END
    ),
}

pythonSources = [
    "addon/globalPlugins",
    "addon/installTasks.py",
]

i18nSources = [
    "buildVars.py",
    "addon/installTasks.py",
    "addon/globalPlugins/docImages/*.py",
    "addon/globalPlugins/docImages/engine/*.py",
]

docFiles = ["readme.html"]

tests = []
excludedFiles = []
baseLanguage = "en"
markdownExtensions = []

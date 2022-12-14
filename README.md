# CJK Material Processing
This is a macro script for screen scraping data form OCLC's WorldCat.org or FirstSearch database. It's primary purpose is to quickly get bibliographic data into a spreadsheet ot help with processing library material requests and orders. It works for materials in any language with a special emphasis on Japanese in particular. The macro is written in [AutoHotKey](https://www.autohotkey.com/) (AHK) and is designed to work in tandem with custom designed spreadsheets.

## Basic Requirements for Use
1. PC with Windows OS.
   - This macro has only been tested in Windows 10. It should work with previous versions as far back as Windows 8.
2. The AHK executible file: *BibData to Spreadsheet.exe*
   - Download it at the top of this page.
3. One of two compatiable Spreadsheets:
   - [Donation Intake Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm)
   - Ordering Materials Template (Not yet available)

This macro has been tested extensively in FireFox while using the spreadsheet as an Excel File. It has successfully worked in Google Chrome and Microsoft Edge, though has not undergone extensive testing in those browsers. Additionally, The Excel File may be used as a spreadsheet in Office 365 or Google Drive. Again, testing has been very limited in these alternative environments, but successful.

## Must Knows
- The "kill switch" for the macro is the backslash "\\" (above the "enter" key on most keyboards). Unusual behavior is always possible and this hotkey will stop the macro.
- Slower internet connections may cause the macro to malfunction and stop.

## Hotkeys to Activate Macro
Several keyboard keys are repurposed to start and stop the macro, referred to as "Hotkeys." F1 through F12 and the numpad keys are repuprosed for quick and easy use of the macro. It is *highly* recommended that you use a keyboard with a numpad as it is much each easier to use. In case you do not have a keyboard with a numpad, the function keys will suffice.

- Pause / F12
  - Deativates hotkeys so you can use your keyboard like normal. Press again to reactivate hotkey functionality.
- Numpad 1 / F1
  - Starts the macro.
- Numpad 4 / F4
  - On a loaded WorldCat.org / FirstSearch record, pulls the Subject Heading data into the spreadsheet.
- Numpad 5 / F5
  - On a loaded WorldCat.org / FirstSearch record, pulls the OCLC# into the spreadsheet.
- Numpad 6 / F6
  - On a loaded WorldCat.org / FirstSearch record, pulls "all" data into the spreadsheet that is available, including:
    - ISBN 10
    - ISBN 13
    - OCLC#
    - Title (Romanized)
    - Title (Non-Romanized)
    - Author/Creator (Romanized)
    - Author/Creator (Non-Romanized)
    - Series Title (Romanized)
    - Series Title (Non-Romanized)
    - Publisher (Romanized)
    - Publisher (Non-Romanized)
    - Date of Publication (Romanized)
    - Subject Headings

## How the Spreadsheet(s) Works 📊
There are to spreadsheet templates to choose from:
- (Collection Development - Donations - Template.xslx)[https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm]
- (Collection Development - Orders - Template.xslx)[https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Orders%20-%20Template.xlsm]

## How the Macro Works
The macro works in tandem with the speadsheet ((see more below)[https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#how-the-spreadsheet-works]) to present bibliographic data in a way that makes tracking library materials, either as orders or donations, or any other number of use cases, easier to manage.

### The Interface & Setup
Content pending, follow the walkthrough for now.

### Basic Walkthrough
_Double click_ the .exe file to start the program.
- An interface will appear, provide the title of your spreadsheet and the name of your Library as it appears in WorldCat.org
- _Click_ "Update Settings"

Open the spreadsheet, there will be several rows of samples to try in the _Donation_ sheet. Alternatively, you would enter an OCLC#, ISBN10, ISBN13, or title in the appropriate column and then run the macro to find a record in WorldCat.org or FirstSearch (WC/FS from now on) and bring all of the associated metadata into the spreadsheet. We will follow a simple example:

1. A donation has arrvied. Using a scan gun, the barcode on a book is scanned to populate the ISBN-13 column.
2. With any cell in the same row as the ISBN13 highlighted, _press numpad1/F1_ to run the macro.
Further Steps Pending

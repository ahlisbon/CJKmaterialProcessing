# CJK Material Processing
This is a macro script for screen scraping data form OCLC's WorldCat.org or FirstSearch database. It's primary purpose is to quickly get bibliographic data into a spreadsheet ot help with processing library material requests and orders. It works for materials in any language with a special emphasis on Japanese in particular. The macro is written in [AutoHotKey](https://www.autohotkey.com/) (AHK) and is designed to work in tandem with custom designed spreadsheets.

## 🔰 Basic Requirements for Use
1. PC with Windows OS.
   - This macro has only been tested in Windows 10. It should work with previous versions as far back as Windows 8.
2. The AHK executible file: *BibData to Spreadsheet.exe*
   - Download it at the top of this page.
3. One of three compatible Excel spreadsheets:
   - [Collection Development / Ordering Materials Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Orders%20-%20Template.xlsm)
      - Optimzed for general collection development.
   - [Donation Intake Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm)
      - Optimized for processing donations of recieved materials.
   - [Users Select Materials Template]() (Not Yet Available)
      - Optimized for letting a Librarian share lists of items for their users to select.

This macro has been tested extensively in FireFox while using the spreadsheet as an Excel File. It has successfully worked in Google Chrome and Microsoft Edge, though has not undergone extensive testing in those browsers. Additionally, The Excel File may be used as a spreadsheet in Office 365 or Google Drive. Again, testing has been very limited in these alternative environments, but successful.

## ⚠ Must Knows
- The "kill switch" for the macro is the backslash ctrl + "\\" (above the "enter" key on most keyboards). Unusual behavior is always possible and this hotkey will stop the macro.
- Slower internet connections may cause the macro to malfunction and stop.

## 🔥 Hotkeys to Activate Macro
Several keyboard keys are repurposed to start and stop the macro, referred to as "Hotkeys." F1 through F12 and the numpad keys are repuprosed for quick and easy use of the macro. It is *highly* recommended that you use a keyboard with a numpad as it is much each easier to use. In case you do not have a keyboard with a numpad, the function keys will suffice.

- \\
  - Stops the Script
- ctrl + \\
  - Exits out the script entirely and closes the program. You will no longer see the green square with an "H" in the taskbar.
- Pause / F12
  - Deativates hotkeys so you can use your keyboard like normal. Press again to reactivate hotkey functionality.
- Numpad-Enter / F1
  - Copies a row of data from the spreadsheet to find a book/item in FirstSearch.
- Numpad 2 / F2
  - On a FirstSearch Record: activates the "Search for versions with same title and author" link to see other versions of the same item.
  - On a search results page in FirstSearch: opens each record in a new browser tab to compare records for importing into a spreadsheet.
- Numpad 3 / F3
  - On a FirstSearch Record: copies the record on screen and reformats it to paste to a spreadsheet.
- ctrl + Numpad 7 / F7
  - Derives an ISBN-10 from an ISBN-13 and pastes it into the ISBN-10 column.
- ctrl + Numpad 8 / F8
  - Opens a menu to fix the ISBN columns in a spreadsheet when there are multiple ISBNs in a cell.
- ctrl + Numpad 9 / F9
  - Copies a cell and if it contains a misformatted Japanese publisher name, will try to fix it.
- Numpad-Division (/)
  - For cells with multiple ISBNs: double click into a cell, place the cursor within an ISBN, then hit the hot key remove all other content. 
- Numpad-Minus (-)
   - On a browser tab with "WorldCat List of Records" OR "WorldCat Detailed Record" in the title: closes the tab.
   - On a spreadheet: Will try to translate the title of a non-English with ChatGPT and paste it into the "translated title" column.
   - On the ChatGPT website with in a "chat" labeled "Translate": Copies the translated result and pastes it in the *Translated Title* column
- Other Features
  - Includes features from the "Diacritics and Nengō" project. [Read more here](https://github.com/ahlisbon/diacriticsAndNengo#typing-vowels-with-diacritics).

## 📊 How the Spreadsheets Work
There are two spreadsheet templates to choose from:
- [Donations - Donations - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm)
- [Collection Development - Orders - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Orders%20-%20Template.xlsm)
- User Selection (pending)

These spreadsheets work in tandem with the the CJK Processing Script to bring data from a FirstSeach WorldCat record and paste it to a spreadsheet. For the script to work correctly, the order of the columns cannot be changed:
- test

Further, both spreadsheets use a great deal of data validation and conditional formatting. Much of these formatting choices were done to give real time feedback on common errors and to improve readability. However, these can be removed entirely. There are also many additional sheets, beginning with "s-", for doing data analysis on your donations/orders as well. These can also be ignored/removed/deleted.

## 📃 Using Slips

Slips help assist with processsing donations. There primary purpose is to help colleagues in other units process the donation if it is written in a language they don't read. The slips reference the spreadsheet you are using that your colleagues consult if they have any questions about the book. Each book has a slip in it with critical information:
- The name of the donation.
- An emoji, to make a visual distincition between donations. This is optional.
- A "key number" that corresponds to the key number in the spreadsheet for cross referencing.
- An identifier such as an OCLC# that colleagues will use to load a record for the book into the library catalog. 
- Volume information, especially useful for large sets of books that look the same.
- Notes, anything else you might need to add.

You have to options for creating slips. Manual and Semi-automatic, each option with it's pros and cons.

### Prepping Slips
Regardless of what method you will use. You need to do some basic preperation before printing slips.
1. In the spreadsheet, go to the "Slips" sheet.
2. In column B, name the donation and copy it into every cell in the B column.
3. Optional - in the A column, choose an emoji to distinguish slips for the donation visually.
4. Make sure that the cells are not spilling over into the page to the right.
   - If so, you need to shorten the text of something in the slip.
6. What you do next depends on if you will write the slips by hand or prepopulate them before printing.

### Manual Slips
1. Print however many pages you will need for all the items in your donation.
2. After printing, cut the slips with a paper slicer. Do not slice too many sheets of paper at once, 4 or 5 max. 

### Semi Automatic Slips
Content Pending.

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

- Pause / F12
  - Deativates hotkeys so you can use your keyboard like normal. Press again to reactivate hotkey functionality.
- Numpad-Enter / F1
  - Copies a row of data from the spreadsheet to find a book/item in FirstSearch.
- Numpad 2 / F2
  - On a FirstSearch Record: activates the "Search for versions with same title and author" link to see other versions of the same item.
  - On a search results page in FirstSearch: opens each record in a new browser tab to compare records for importing into a spreadsheet.
- Numpad 3 / F3
  - On a FirstSearch Record: copies the record on screena and reformats it to paste to a spreadsheet.
- ctrl + Numpad 7 / F7
  - Derives an ISBN-10 from an ISBN-13 and pastes it into the ISBN-10 column.
- ctrl + Numpad 8 / F8
  - *Experimental:* Opens a menu to fix the ISBN columns in a spreadsheet when there are multiple ISBNs in a cell.
- ctrl + Numpad 9 / F9
  - Copies a cell and if it contains a misformatted Japanese publisher name, will try to fix it.
- Numpad-Minus
   - Will try to translate the title of a non-English title and paste it into the "translated title" column.

## 📊 How the Spreadsheets Work
There are two spreadsheet templates to choose from:
- [Collection Development - Donations - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm)
- [Collection Development - Orders - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Orders%20-%20Template.xlsm)

These spreadsheets work in tandem with the macro to bring data from WC/FS. The most important element is that *the order of the columns CANNOT be changed*. However, there is a great deal that can be altered. Each spreadsheet has a "Guide" sheet that explains how the spreadsheet works.

Further, both spreadsheets use a great deal of data validation and conditional formatting. Much of these formatting choices were done to give real time feedback on common errors and to improve readability. However, these can be removed entirely. There are also many additional sheets, beginning with "s-", for doing data analysis on your donations/orders as well. These can also be ignored/removed/deleted.

## ℹ How the Macro Works
The macro works in tandem with the speadsheet ([see more above](https://github.com/ahlisbon/CJKmaterialProcessing#-how-the-spreadsheets-work)) to present bibliographic data in a way that makes tracking library materials, either as orders or donations, or any other number of use cases, easier to manage.

### Basic Walkthrough
_Double click_ the .exe file to start the program.
- An interface will appear, provide the following information:
  - The title of your spreadsheet and the name of your Library as it appears in WorldCat.org.
  - The URL your institution uses to access FirstSearch
  - The title of your Library as it appears in WorldCat.org.
- _Click_ "Update Settings." These will all be saved for the next time you use the macro.

Open the spreadsheet, there will be several rows of samples to try with the macro. You can delete these if you don't want to try some practice runs. Alternatively, you would enter an OCLC#, ISBN10, ISBN13, or title in the appropriate column and then run the macro to find a record in WorldCat.org or FirstSearch (WC/FS from now on) and bring all of the associated metadata into the spreadsheet. We will follow a simple example:

1. A donation has arrived. Using a scan gun, the barcode on a book is scanned to populate the ISBN-13 column.
2. Make sure a browser window is open.
   - If using FS, make sure to initially login.
4. With any cell in the same row as the ISBN13 highlighted, _press numpad1/F1_ to run the macro.
   - In WC, the macro will activate your broswer window, then open a new window to actually search WC.
   - In FS, the search will happen in the same window.
   - Expect to see some flashing blue as the macro highlights text to determine what to do next.
   - One Result vs. Multiple Results:
      - If there is 1 result in FS/WC, that record will load instantly.
      - If there are multiple results, the macro will stop at this list.
        - You should click on the record that appears to be of the highest quality/accuracy.
      - If you run a title search, the macro ALWAYS stops at the list of results.
      - If there are no results, the macro will end the search and return to the spreadsheet and enter n/a into the current row of cells.
        - At this point you will have to enter data about your item manually.
5. Assess the quality of the record for if you want to import the data.
   - If you feel the recrd is low quality, you can press _numpad7/F7_ to see what other records are available.
     - A list of records will load and then each record will attempt to be loaded individually in a new tab.
     - You can quickly browse through these tabs by pressing _numpad0_.
     - You can remove the tabs of records you have ruled out be pressing _numpad-_ (numapd minus sign).
     - When you have found a record you want to import to the spreadsheet, press _numpad6/F6_.
       - Alternatively, you can press _numpad4/F4_ to only pull Subject Heading data. This option is available becuase sometimes a good record can have poor subject headings, and you may with to opt for the subject headings of a different record. In this situation, you would restart the search process by pressing numpad1/F1, the record would appear instatntly because the macro is now using the OCLC# to load the exact record. Press numpad7/F7 to see other records and when you have found a record with subject headings you appove us, press numpad4/F4 to bring them into the spreadsheet.
      - You can press _numpad5/F5_ if you only want to import the OCLC number.
6. The bibliographic data from FS/WC will paste to the spreadsheet. The macro will move to the next row for you to run the macro again.
7. Checking the quality of the data:
   - What to check for:
     - Are there any columns you need to fill out manually? Either columns designed to be filled out manually or ones that say "n/a" that require your review.
     - Are there any errors to fix?
   - You can check as you row by row, or at the end. Checking row by row is best for donations, as you may need the item in your hands to confirm data that might be missing.
8. If you plan to use slips, and you plan to use the slips sheet in the spreadsheet ([see below](https://github.com/ahlisbon/CJKmaterialProcessing#-using-slips)). if is imperative that you keep the physical books in precise order so that the other of the slips and the order of the books is the same.

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

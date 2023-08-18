# CJK Material Processing
This is a macro script for screen scraping data form OCLC's WorldCat.org or FirstSearch database. It's primary purpose is to quickly get bibliographic data into a spreadsheet ot help with processing library material requests and orders. It works for materials in any language with a special emphasis on Japanese in particular. The macro is written in [AutoHotKey](https://www.autohotkey.com/) (AHK) and is designed to work in tandem with custom designed spreadsheets.

# 🔰 Basic Requirements for Use
1. PC with Windows OS.
   - This macro has only been tested in Windows 10. It should work with previous versions as far back as Windows 8.
2. Files to download:
   - The AHK executible file: *BibData to Spreadsheet.exe*
   - The INI file: *Bibdata to Spreadsheet.ini* (this is for saving your settings and preferences
   - Always keep the two files above together in the same folder.
   - One of three compatible Excel spreadsheets:
     - [Ordering Materials Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Orders%20-%2020xx-xx%20-%20Template.xlsm)
       - Optimized for general collection development.
     - [Donation Intake Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Donation%20-%20Template.xlsm)
       - Optimized for processing donations of recieved materials.
     - [Users Select Materials Template](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Select%20Monographs%20-%20User%20Name%20-%20Template.xlsm) (Not Yet Available)
       - Optimized for letting a Librarian share lists of items for their users to select.

This macro has been tested extensively in FireFox while using the spreadsheet as an Excel File. It has successfully worked in Google Chrome and Microsoft Edge, though has not undergone extensive testing in those browsers. Additionally, The Excel File may be used as a spreadsheet in Office 365 or Google Drive. Again, testing has been very limited in these alternative environments, but successful.

# ⚠ Must Knows
- The "kill switch" for the macro is the backslash ctrl + "\\" (above the "enter" key on most keyboards). Unusual behavior is always possible and this hotkey will stop the macro.
- Slower internet connections may cause the macro to malfunction and stop.

# 🔥 Hotkeys to Activate Macro
Several keyboard keys are repurposed to start and stop the macro, referred to as "Hotkeys." F1 through F12 and the numpad keys are repuprosed for quick and easy use of the macro. It is *highly* recommended that you use a keyboard with a numpad as it is much each easier to use. In case you do not have a keyboard with a numpad, the function keys will suffice.

## -- ⌨ Essential Hotkeys
- \\
  - Stops the script
- ctrl + \\
  - Stope the script and closes the program. You will no longer see the green square with an "H" in the taskbar.
- Pause / F12
  - Deativates hotkeys so you can use your keyboard like normal. Press again to reactivate hotkey functionality.

## --  ⌨ Hotkeys, listed in order of how to get bibliographic data on the spreadsheet.
**Important:**
  - The same key will perform different actions depending on the active window. For example, the Numpad Enter key does something different in Excel than in a browser window.
  - Emphasis should be on learning to use the hotkeys on the numpad, as it is compact and easy to rest your hand. Equivalent hotkeys are also available in the function keys row for when the numpad is not available on a keyboard.

- **Numpad Enter / F1**
  - _On a spreadsheet_: Copies a row of data from the spreadsheet to find a book/item in FirstSearch. Assumes you are storing at least OCLC#, ISBN, or Title to on the spreadsheet to look up in FirstSearch.
- **Numpad Plus / F2**
  - _On a FirstSearch Record_: activates the "Search for versions with same title and author" link to see other versions of the same item.
- **Numpad Enter / F1**
  - _On a FirstSeach record_: imports bibliographic data from the record into a spreadsheet.
  - On a search results page in FirstSearch: opens each record in a new browser tab to compare records for importing into a spreadsheet.
    - After the tabs have loaded, use **ctrl + Numpad 0** to quickly cycle through the tabs. Use **Numpad Minus** to quickly close a tab for a record you don't want to import.

## -- 🧽 Data Clean up Hotkeys
- ctrl + Numpad 7 / F7
  - On a spredsheet: Derives an ISBN-10 from an ISBN-13 and pastes it into the ISBN-10 column. Will also add an amazon URL to check price.
    - This will not work for ISBN's beginning with 979.
- ctrl + Numpad 8 / F8
  - On a spredsheet: Opens a menu to fix the ISBN columns in a spreadsheet when there are multiple ISBNs in a cell. Includes a contextual menu on how to use.
- ctrl + Numpad 9 / F9
  - On a spreadsheet: Copies a cell and if it contains a misformatted Japanese publisher name, will try to fix it.
- Numpad Division (/)
  - On a spreadsheet: For cells with multiple ISBNs: double click into a cell, place the cursor within an ISBN, then hit the hot key remove all other content.
 
## -- 💬 Chat GPT Translation Assistance HotKeys
**Important**
   - You must make an acount with Chat GPT and make sure you have a brower window with ChatGPT open.
   - You must create a "new chat" and name it "Translate" - After you've created this chat, make sure to activate it before running the script.
- Numpad Minus (-)
   - On a spreadheet: Will try to translate the title of a non-English with ChatGPT and paste it into the "translated title" column. By default will only translate the title in the row your cursor is on.
   - To translate mulitple titles, highlight however many titles in the "Title (N)" column (column U).

## -- 💴 Price Estimate HotKeys
If you are interested in tracking the general cost of the books you are selecting, this scirpt will look up materials to compare prices. Extracting the price works one of two ways once you have identified a price you believe is acceptable to pay for the item:
**Only available for Japanese at the moment**
- ctrl + Numpad Plus
  - On a spreadsheet: Looks up the item in four websites:
    - Once an amazon page has loaded, you will need to browse for an appropriate price and select highlight it. Then **press Numpad Enter or Enter** to bring that price back to the spreadsheet. 
      - [Amazon.jp](https://www.amazon.co.jp/)
      - [Amazon.com] (https://www.amazon.com/)
    - One you have loaded a record on this site, **press Numpad Enter or Enter** to bring that price back to the spreadsheet.
      - [Nihon no Furuhonya](https://www.kosho.or.jp/)
      - [JPT Book News](https://jptbooknews.jptco.co.jp/)
   - Th "URL for Price Check" column (H) will update with the relvant URL.

## -- 🍹 Quality of Life HotKeys
  - Includes features from the "Diacritics and Nengō" project. [Read more here](https://github.com/ahlisbon/diacriticsAndNengo#typing-vowels-with-diacritics).

# 📊 How the Spreadsheets Work
There are three spreadsheet templates to choose from:
- [Donations - Donations - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Donation%20-%20Template.xlsm)
- [Collection Development - Orders - Template.xslx](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/Collection%20Development%20-%20Orders%20-%20Template.xlsm)
- User Selection (pending)

## -- 👨‍🏫 Spreadsheet Fundamentals
- *Do NOT*
  - hide columns
  - change the order of columns
  - rename column headers in row 6. Exceptions explained below.
  - add your own formatting. There is an Excel macro ([see below](https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#-excel-macros)) that when activated will wipe most custom formatting you try to setup.
  - rename the various sheets within the workbook.

These spreadsheets work in tandem with the the CJK Processing Script to bring data from a FirstSeach WorldCat record and paste it to a spreadsheet. For the script to work correctly, the order of the columns cannot be changed:
- Look in row 5 of any spreadsheet.
  - if the column says "srcipt" you cannot repurpose that column.
  - if the column says "free" you can repurpose those columns for any kind of manual data entry. You can also rename the headers (row 6) in these columns.
  - In the donation template Excel file, columns B, C, D, and F are much smaller as the script never interacts with those columns, but they can be repurposed for manual data entry.
  - In the select template Excel file, columns B through O are narrow and do not show text. This is to remove visual clutter.

## -- 🎨 Formatting 
- The "Collection Development" and "Donation" Excel templates use condtional formatting to help quickly identify anomalies.
- Duplicate Checking:
  - If an ISBN or OCLC# appears twice, it will turn red.
  - If a title or series title appears twice, it turns yellow.
- Incorect formatting for an ISBN
  - If there is a space in the ISBN, the cell will turn yellow and should be addressed.
- Preferred Vendor and Collection Columns (F and I)
  - These colums will stay red until they are filled out.
- Note for Acquistions Column (G)
  - Turns yellow, emphasizes there is something to read and review.

## -- 🛠 Excel Macros
- There are two important Excel macros built into the "Collection Development" and "Donation" Excel templates. Each macro has a button you can push to activate it in the upper right of the table in either spreadsheet. If you choose to use these spreadsheets on any web platform (Microsoft 365, Google Sheets), these will not be available to you.
  - **Reset Formatting:** this resets all defaults regarding formatting in the spreadsheet. Over time, the spreadsheet can become "bloated" with formatting rules. This reset button instanly resets all formatting and improves the performance of the spreadsheet.
  - **Convert CJK Currencies:** Becuase of some limitations around how data is brought into the spreadsheet, it is necessary to run this macro to ensure Chinese Yuan, Japanese Yen, and Korean Won all display correctly in the "USD Estimate" column (AG)

## -- 📈 Statistics Sheets
There are several additional worksheets besides the "Orders" or "Donations" sheet (depending on which Excel file you are using) where you record materials for purchase. All of these have the suffix "s-" for "statistics." Each sheet provides insights about your materials. You can delete these sheets if you are interested in them.
**What to know:**
  - You can delete the sheets starting with "s-" and they will have no effect on the "Orders" or "Donations" sheet.
  - Data does not update automatically, you need to click the "Refresh Data" button on each sheet.
  - Do not try to move or rename tables, pivot tables, or charts.
  - Do not try to edit the data in a pivot table, if you see a mistake, you need to fix it at the source, which will be the "Orders" or "Donations" sheet depending on which Excel template you are using.



# 🖱 Understanding the GUI
When you run the EXE file for the script, you have 6 data fields to fill out in the GUI that affect how the script will function.

## -- 📁 File Name Prefixes
You can rename your files to whatever you like. However for data to be pasted correctly, the script needs to know which type of spreadsheet you are using. If you are preparing different books for purchase, make sure all your spreadsheets start with the same prefix, such as:
- Collection Development - 2022-23
- Collection Development - Rare Korean Books

Other things to be aware of regarding naming conventions:
- Do not use the same prefix for different types of sheets.
- Avoid having two sheets open at once. The script will not let you continue if you have more than one open.

## -- 🌐 FirstSearch URL for Your Institution
In order to pull data like an OCLC#, ISBN#, or Title from the spreadsheet to search FirstSearch, you need to provide the script with a URL that can load FirstSearch

- Do NOT log into FirstSearch and copy that URL, it will not work.
- Identify they URL your institution uses to access FirstSearch. Whatever link you are using to open FirstSearch, copy that link and paste it here.

## -- ✅ Use Check Mode
Before pasting data to your spreadhseet, a window will appear for you to review the bibliographic data.

## -- 🕗 Wait longer for Websites to Load
IF you are using a slower connection speed, you can increase the time the script will wait to load. The default is 1, which equals 3 seconds. 2 and 3 are multipliers, so the script will wait 6 or 9 seconds respectively.

# 📃 Using Slips

Slips help assist with processsing donations. There primary purpose is to help colleagues in other units process the donation if it is written in a language they don't read. The slips reference the spreadsheet you are using that your colleagues consult if they have any questions about the book. Each book has a slip in it with critical information:
- The name of the donation.
- An emoji, to make a visual distincition between donations. This is optional.
- A "key number" that corresponds to the key number in the spreadsheet for cross referencing.
- An identifier such as an OCLC# that colleagues will use to load a record for the book into the library catalog. 
- Volume information, especially useful for large sets of books that look the same.
- Notes, anything else you might need to add.

You have two options for creating slips. Manual and Semi-automatic, each option with it's pros and cons.

## -- 🖨 Prepping Slips
Regardless of what method you will use. You need to do some basic preperation before printing slips.
1. In the spreadsheet, go to the "Slips" sheet.
2. In column B, name the donation and copy it into every cell in the B column.
3. Optional - in column A, choose an emoji to visually distinguish slips for different donations.
4. Make sure that the cells are not spilling over into the page to the right.
   - If so, you need to shorten the text of something in the slip, or turn off "Word Wrap."
6. What you do next depends on if you will write the slips by hand or prepopulate them before printing.

## -- ✒ Manual Slips
1. Print however many pages you will need for all the items in your donation.
2. After printing, cut the slips with a paper slicer. Do not slice too many sheets of paper at once, 4 or 5 max.
3. Fill out the slips manually, making sure the key number and OCLC number match.

## -- ⬜ Semi Automatic Slips
Content Pending.

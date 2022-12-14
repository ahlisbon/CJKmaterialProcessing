# CJK Material Processing
This is a script for screen scraping data form OCLC's WorldCat.org or FirstSearch database. It's primary purpose is to quickly get bibliographic data into a spreadsheet ot help with processing library material requests and orders. It works for materials in any language with a special emphasis on Japanese in particular. The script is written in AutoHotKey (AHK) and is designed to work in tandem with custom designed spreadsheets.

## Basic Requirements for Use
1. PC with Windows OS.
   - This script has only been tested in Windows 10. It should work with previous versions as far back as Windows 8.
2. The AHK executible file: *BibData to Spreadsheet.exe*
   - Download it at the top of this page.
3. One of two compatiable Spreadsheets:
   - [Donation Intake Template]()
   - Ordering Materials Template (Not yet available)

## Must Knows
<ul>
  <li>
    The "kill switch" for the macro is the backslash "\" (above the "etner" key on most keyboards). Unusual behavior is always possible and this hotkey will stop the macro.
  </li>
  <li>
    You must keep the columns in the spreadsheet in the same order as they are provided. It is possible to change what data goes to which column in this macro, but it takes more than basic knowledge of the AHK code to do this.
  </li>
  <li>
    The name of your Google Sheet must start with "Collection Development - " (note that there is a final space after the hyphen. This is because the code is looking for a tab in the Google Chrome browser with this name. The macro stops it if does not find it. If you want to rename the spreadsheet, you must also edit the code on line 389 to reflect the new name you choose.
  </li>
</ul>
  </li>
## Hotkeys to Activate Macro
<ul>
  <li>
    PrintScreen: runs the macro on the Google Sheet, uses the ISBN-13 in column M to Look up book in worldcat.org and stops.
  </li>
  <li>
    ctrl + PrintScreen: same as above, but if there is only one record in worldcat.org, that record will be opened and the bibliographic data will be pasted to the Google Sheet. If there is more than one record, a link back to the search results will be pasted in column O for you to return to and check manually. This macro will loop, going to each row in the Google Sheet until there is an empty row.
  </li>
  <li>
    F6: When a row has a worldcat.org search result link in column O, use this macro to reload those results.
  </li>
  <li>
    F7: In a worldcat.org search that has multiple results, find a worldcat.org record that you want to import to the Google Sheet. Use this to extract the data for that record and paste it to the Google Sheet.
  </li>
</ul>
<h1>Editing the Code</h1>
<ul>
  <li>
    line 200: You must change the addess to how your library appears in a worldcat.org record. For example, the University of Colorado appears as "Boulder, CO 80309 United States" - This helps with checking if a requested book is already owned by your library, and a "y" (for yes) will be pasted to column L. In the code itself, you must use two comma's:"Boulder,, CO 80309 United States" because a single comma is used as code in the AHK language, but two commas are recognized as an actual comma.
  </li>
  <li>
    line 362: Same as above.
  </li>
</ul>
<h1>Known Issues</h1>
<ul>
  <li>
    When pressing the killswitch (the "\" key) sometimes the "ctrl" or "alt" or "shift" keys can stick. Tap each of them quickly to unstick them. I have experienced one time a situation it which they would not unstick at all. Implying another key was stuck. I restarted my machine in this instance.
  </li>
</ul>

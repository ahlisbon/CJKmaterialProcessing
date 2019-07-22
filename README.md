# CJK Material Processing
AutoHotKey Script for processing Chinese, Japanese, Korean and other non-Roman scripts. Used to process Library material requests and donations. For use with Google Sheets and WorldCat.org in the Google Chrome browser.
<h1>Basic Requirements for Use</h1>
<ol>
  <li>
    PC with Windows 10 Operating System (not tested on previous OS)
  </li>
  <li>
    <a href="https://www.google.com/chrome/">Google Chrome Browser</a> (not tested in other browsers)
  </li>
  <li>
    <a href="https://www.autohotkey.com/">AutoHotKey (AHK)</a> (software that can run the macro)
  </li>
</ol>
<h1>Hotkeys to Activate Macro</h1>
<ul>
  <li>
    PrintScreen: runs the macro on the Google Sheet, uses the ISBN-13 in column M to Look up book in worldcat.org and stops.
  </li>
</ul>
<ul>
  <li>
    ctrl + PrintScreen: same as above, but if there is only one record in worldcat.org, that record will be opened and the bibliographic data will be pasted to the Google Sheet. If there is more than one record, a link back to the search results will be pasted in column O for you to return to and check manually. This macro will loop, going to each row in the Google Sheet until there is an empty row.
  </li>
</ul>
<ul>
  <li>
    F6: When a row has a worldcat.org search result link in column O, use this macro to reload those results.
  </li>
</ul>
<ul>
  <li>
    F7: In a worldcat.org search that has multiple results, find a worldcat.org record that you want to import to the Google Sheet. Use this to extract the data for that record and paste it to the Google Sheet.
  </li>
</ul>
<h1>Editing the Code</h1>
<ul>
  <li>
    line 200: You must change the addess to how your library appears in a worldcat.org record. For example, the University of Colorado appears as "Boulder, CO 80309 United States" - This helps with checking if a requested book is already owned by your library, and a "y" (for yes) will be pasted to column L. In the code itself, you must use two comma's:"Boulder,, CO 80309 United States" because a single comma is used as code in the AHK language, but two commas are recognized as an actual comma.
  </li>
</ul>

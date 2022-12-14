# CJK Material Processing
This is a macro script for screen scraping data form OCLC's WorldCat.org or FirstSearch database. It's primary purpose is to quickly get bibliographic data into a spreadsheet ot help with processing library material requests and orders. It works for materials in any language with a special emphasis on Japanese in particular. The macro is written in [AutoHotKey](https://www.autohotkey.com/) (AHK) and is designed to work in tandem with custom designed spreadsheets.

## Basic Requirements for Use
1. PC with Windows OS.
   - This macro has only been tested in Windows 10. It should work with previous versions as far back as Windows 8.
2. The AHK executible file: *BibData to Spreadsheet.exe*
   - Download it at the top of this page.
3. One of two compatiable Spreadsheets:
   - [Donation Intake Template]()
   - Ordering Materials Template (Not yet available)

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


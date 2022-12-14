#noenv  ; recommended for performance and compatibility with future autohotkey releases.
#warn  UseUnsetGlobal, Off ; enable warnings to assist with detecting common errors.
;sendmode input  ; recommended for new scripts due to its superior speed and reliability.
setWorkingDir %a_scriptdir%  ; ensures a consistent starting directory.
setTitleMatchMode, 2
setKeyDelay 100

;This set of macros extracts data from spreadsheets, searches WorldCat.org to extract bibliographic data form the site, then return bibliographic data to the spreadsheet.
;The spreadsheet's title must being with "Collection Developement - " the name is case sensitive. note the last space after the dash.
;The macro has been tested with Excel, Google Sheets, and Excel in Office365; in both Firefox and Chrome.
;When macro pulls data from spreadsheet to search WorldCat it prioritizes ISBN searches, then vernacular title searches if there is no ISBN, then Romanized/English searches if there is no non-Romanized/English title.

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■ Set Preference in GUI
	;▼ Read ini file "bibdataToSpreadsheet.ini" to populate gui fields and variables.
		iniRead, sheetName, bibdataToSpreadsheet.ini, Database, sheetName, %A_Space%
		iniRead, useFS, bibdataToSpreadsheet.ini, Database, useFS, 1
		iniRead, useWC, bibdataToSpreadsheet.ini, Database, useWC, 0
		iniRead, fsURL, bibdataToSpreadsheet.ini, Database, fsURL, %A_Space%
		iniRead, libName, bibdataToSpreadsheet.ini, Database, libName, %A_Space%
		iniRead, checkMode, bibdataToSpreadsheet.ini, Database, checkMode, %A_Space%
		iniRead, loopMode, bibdataToSpreadsheet.ini, Database, loopMode, %A_Space%
	;▼ Search Engine settings.
		if (useFS= 1)
			searchEngine:= "FirstSearch"
		if (useWC= 1)
			searchEngine:= "WorldCat.org"

	;▼ Create GUI
		gui, new
		gui, font, s10
		gui, add, radio, group	vuseFS	x20 y220, Use FirstSearch to pull bibliographic data and put in a spreadsheet.
		gui, add, radio,		vuseWC	x20 y425, Use WorldCat to pull bibliographic data and put in a spreadsheet.

		gui, add, text,		x20 y20, Name of your spreadsheet, there must be text in this box 👇 👇 👇
		gui, add, text,		x40 y45, Include the title of your spreadsheet below, case sensitive.
		gui, add, edit, vsheetName w530
		gui, add, text,		x40 y95, ▶ REQUIRED: The beginning of the title of your spreadsheet must match the text above.
		gui, add, text,		x60 y120, Example: if you write "Collection Development - " above, the following will work:
		gui, add, text,		x60 y140, ⚪ Good: Collection Development - Donation
		gui, add, text,		x60 y160, ⚪ Good: Collection Development - 2021 Purchases
		gui, add, text,		x60 y180, ❌ Bad: Collections - Main Library Collection

		gui, add, text,		x20 y200, ----------------------------------------------------------------------------------------------------------------------------------------------------------------

		gui, add, text,		x40 y245,	What is your institution's URL for accessing FirstSearch:
		gui, add, edit, vfsURL w530
		gui, add, text,		x40 y295,	▶ REQUIRED: Provide the URL your institution uses to access FirstSearch.
		gui, add, text,		x40 y315,	❌ DO NOT copy the URL after you have signed into FirstSearch.
		gui, add, text,		x40 y340,	1. Locate your institution's list of databases and right click the link you would click to access FirstSearch.
		gui, add, text,		x40 y360,	2. In the popup menu, select "copy link" or a similar option.
		gui, add, text,		x40 y380,	3. 👆 Paste that link into the text box above.

		gui, add, text,		x20 y400, ----------------------------------------------------------------------------------------------------------------------------------------------------------------

		gui, add, text,		x40 y450, What is the name of your Institution on WorldCat.org?
		gui, add, edit, vlibName w530
		gui, add, text,		x40 y500,	▶ REQUIRED: Provide the name of your institution as it appears in a WorldCat item record in the "Find`r`na Copy at a Library" section.
		gui, add, text,		x50 y540,	1. Go to WorldCat.org and search for a book you know is in your institution's collection.
		gui, add, text,		x50 y560,	2. Scroll down to the "Find a Copy at a Library" section.
		gui, add, text,		x50 y580,	3. You should see your intitution in the list.*
		gui, add, text,		x50 y600,	4. Copy your institution's name exactly as it is, be sure to avoid spaces at the beginning or end.
		gui, add, text,		x50 y620,	5. 👆 Paste your institution's name into the text box above.
		gui, add, text,		x50 y640,	* If you don't see your institution, you may need to set your location manually.
		gui, add, text,		x50 y660,	* There is on option on the WorldCat.org homepage in the upper right to manually set your location.

		gui, add, text,		x20 y680, ----------------------------------------------------------------------------------------------------------------------------------------------------------------

		gui, add, checkbox, vcheckMode  x20 y700, Use check mode.
		gui, add, text,		x50 y725, Check mode lets you reivew the bibliographic data pulled from FirstSearch or WorldCat.org in an easy`r`nto read format before it's pasted to a spreadsheet.

		gui, add, text,		x20 y760, ----------------------------------------------------------------------------------------------------------------------------------------------------------------

		gui, add, checkbox, vloopMode x20 y780, Use Loop mode *EXPERIMENTAL*
		gui, add, text,		x50 y805, Loop mode will go row by row through a spreadsheet and pull data *IF* there is data in the`r`nOCLC# column. 
		
		gui, add, text,		x20 y840, ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		
		gui, add, button, w120 x20 y860, &Update Settings

		gui, show, w680 h905

	;▼ Prepopulates text fields. Solution found at https://www.autohotkey.com/board/topic/3177-gui-edit-control-filled-by-variable/
		guiControl,, sheetName, %sheetName%
		guiControl,, useFS, %useFS%
		guiControl,, useWC, %useWC%
		guiControl,, fsURL, %fsURL%
		guiControl,, libName, %libName%
return

		buttonUpdateSettings:
		gui, submit, noHide	

	;▼ Writes values to ini file.
		iniWrite, %sheetName%, bibdataToSpreadsheet.ini, Database, sheetName
		iniWrite, %useFS%, bibdataToSpreadsheet.ini, Database, useFS
		iniWrite, %useWC%, bibdataToSpreadsheet.ini, Database, useWC
		iniWrite, %fsURL%, bibdataToSpreadsheet.ini, Database, fsURL
		iniWrite, %libName%, bibdataToSpreadsheet.ini, Database, libName
		iniWrite, %checkMode%, bibdataToSpreadsheet.ini, Database, checkMode

	;▼ Spreadsheet name setting.
		allowMacro:= "yes"
		if (sheetName= ""){
			msgBox, You must have a title for the spreadsheet you will use with this macro.
			global allowMacro
			allowMacro:= "no"
			return
			}
	;▼ Search Engine settings.
		if (useFS= 1)
			searchEngine:= "FirstSearch"
		if (useWC= 1)
			searchEngine:= "WorldCat.org"
	;▼ Check FirstSearch Setting.
		if(useFS= 1 and fsURL= ""){
			msgBox, You must provide your institution's URL for accessing FirstSearch.
			global allowMacro
			allowMacro:= "no"
			return
			}
	;▼ Check WorldCat.org Setting.
		if(useWC=1 and libName= ""){
			msgBox, You must provide your institution's name as it appears in WorldCat.org.
			global allowMacro
			allowMacro:= "no"
			return
			}
	;▼ Check mode settings.
		if (checkMode= 0)
			checkMode:= "off"
		if (checkMode= 1)
			checkMode:= "on"
	;▼ Loop mode settings.
		if (loopMode= 0)
			loopMode:= "off"
		if (loopMode= 1)
			loopMode:= "on"
	;▼ Confirm Settings
	msgBox, The title of your spreadsheet is:`r`n▶　%sheetName%`r`n`r`nYou are using:`r`n▶　%searchEngine%`r`n`r`nYour Institution's Name for duplicate checking on WorldCat.org is:`r`n▶　%libName%`r`n`r`nThe URL your institution uses to access FirstSearch is:`r`n▶　%fsURL%`r`n`r`nLoop mode is set to:`r`n▶　%loopMode%`r`n`r`nCheck mode is set to:`r`n▶　%checkMode%
return

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■ Subjects Turns off hotkeys, especially to use the numpad for input instead of running macros.
pause::suspend ; Deactivates/reactivates hotkeys.
F12::suspend

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■ Search with OCLC# / ISBN-13 / ISBN-10 / non-Romanized title / Romanized title columns from spreadsheet.
numpad1::
F1::
	;▼ Confirm settings.
		;▼▲▼ Function: prevents macro from running when GUI settings have not been filled out correctly.
		allowMacroToRun()
		allowMacroToRun(){
			global allowMacro
			if inStr(allowMacro, "no"){
			msgBox, You cannot yet run the macro because your settings are not filled out correctly.
			exit
			}
		}

	;▼ For loop mode.
	doLoop:

	;▼ Copy a row in spreadsheet to run a search.
		;▼▲▼ Function: Calls a window that starts with a value equal to the "sheetName"
		confirmSpreadsheet()
		confirmSpreadsheet(){
			if winExist("Collection Development - ")
				winActivate
			else{
			msgBox, You do not have a spreadsheet beginning with "Collection Development - " in one of the following programs:`r`n`r`nExcel`r`n`r`nEdge`r`n`r`nFirefox`r`n`r`nChrome open with a window that begins with "collection development - ". If you have an online spreadsheet in an Edge, Firefox or Chrome window, make sure the tab with the speadsheet is active.
			exit
			}	
		}
		send {home}
		send +{space}
		;▼ Sometimes if there is some data in a row already, the entire row won't be selected. there is a built in check for this further down in the code. Issue in Google Sheets only.
		SelectRowAgain:= ""
		tryRowSelectAgain:
		;▲
		sleep 300
		clipboard:= ""
		send ^c
		clipwait, 2
		if errorLevel {
			msgBox, The row currently selected in your spreadsheet could not be copied.
			return
			}
		sleep 1000
		;▼　Stores each spreadsheet cell as it's own variable: bibArray[1]= data in column a, bibArray[2]= data in column b, etc.
		bibArray:= strSplit(clipboard, a_tab)

	;▼ Confirm a browser window is open.
		confirmBrowser()
		confirmBrowser(){
			if winExist("ahk_exe firefox.exe"){
				winActivate
				return
				}
			if winExist("ahk_exe chrome.exe"){
				winActivate
				return
				}
			if winExist("ahk_exe msedge.exe"){
				winActivate
				return
				}
			msgBox, You do not have one of these browsers open:`r`n`r`n▶　Firefox`r`n▶　Chrome`r`n▶　Edge`r`n`r`nPlease open one of these browsers and run the macro again.
			exit
		}

	;▼ 1 of 4 critical pieces of data is selected to run a search.
	;▼ If the OCLC# column is not empty.
		if (bibArray[16]!= ""){
			if (bibArray[16]= "n/a"){
				msgBox, The OCLC# column has a value of "n/a" - There is likely no %searchEngine% record for this item.`r`n`r`nIf you think this is an error. Clear the "n/a" value from column P and run the macro again to search with the ISBN-10, ISBN-13, Romanized/English title, or non-Romanized title.
				return
			}
		}
		if (bibArray[16]= "")
		goTo skipOCLCsearch
;🔍
	;▼ Search by OCLC# via FirstSearch
		if (searchEngine= "FirstSearch"){
			fsPrefix:= "no:"
			searchFirstSearch()
			searchFirstSearch(){
				global fsURL
				global bibArray
				global fsPrefix
				send ^n
				sleep 1000
				send !d
				sendInput % fsURL
				send {enter}
				sleep 3000
			;◯ Loop to check that page has loaded and a search can happen
				loop, 5 {
					sendInput % fsPrefix
					sendInput % bibArray[16]
					copyAllclearClip()
					if inStr(clipboard, fsPrefix)
						break
					sleep 2000
				}
			}
		}
			send {enter}
			return
;🌐
	;▼ Search by OCLC# via WorldCat.org 
		if (searchEngine= "WorldCat.org"){
			msgBox, hello from OCLC number search.
			send ^n
			sleep 1000
			send !d
			sendInput {raw}https://www.worldcat.org/OCLC/
			sendInput % bibArray[16]
			send {enter}
		;◯ Loop to check that page has loaded and then the "Show more information" drop down can be activated.
			loop, 5 {
				copyAllclearClip()
				if inStr(clipboard, "Find a Copy at a Library")
				break
				sleep 2000
				}
		sleep 1000
		send ^f
		sendInput {raw}Show more information
		sleep 300
		send {enter}
		send {esc}
		;▼　Need to tab and tab back to highlight and activate the dropdown.
		send {tab}
		send +{tab}
		send {enter}
		if(loopmode= "on")
			goTo loopPullBibData
		return
		}

		skipOCLCsearch:
	;▼　Search by ISBN-13.
		if (bibArray[14]!= ""){
			ISBN:= bibArray[14]
			if (searchEngine= "FirstSearch")
				goTo ISBNsearchFS
			if (searchEngine= "WorldCat.org")
				goTo ISBNsearchWC
			}
	;▼　Search by ISBN-10.
		if (bibArray[15]!= ""){
			ISBN:= bibArray[15]
			if (searchEngine= "FirstSearch")
				goTo ISBNsearchFS
			if (searchEngine= "WorldCat.org")
				goTo ISBNsearchWC
			}
	;▼　Search by Romanized/English title.
		if (bibArray[18]!= "" ){
			Title:= bibArray[18]
			msgBox, % Title
			if(searchEngine= "FirstSearch")
				goTo titleSearchFS
			if(searchEngine= "WorldCat.org")
				goTo titleSearchWC
			}
	;▼　Search by non-Romanized title.
		if (bibArray[17]!= ""){
			Title:= bibArray[17]
			if (searchEngine= "FirstSearch")
				goTo titleSearchFS
			if (searchEngine= "WorldCat.org")
				goTo titleSearchWC
			}
		if (SelectRowAgain= "yes")
			{
			msgBox, There is no data in the spreadsheet, the macro has stopped.
			return
			}
		send +{space}
		SelectRowAgain:= "yes"
		goTo tryRowSelectAgain

;▼　■　■　■　■　■　■　■　■　■　■　■　■　Seaching and pulling bibliographic data in FirstSearch.
	;▼ Title search in FirstSearch.
	titleSearchFS:
		confirmFirstSearchLanding()
		fsPrefix:= "ti:"
		searchFirstSearch()
		sendInput {raw}ti:
		sendInput % Title
		send {enter}
return
	
	;▼ ISBN search in FirstSearch.
	ISBNsearchFS:
		confirmFirstSearchLanding()
		fsPrefix:= "bn:"
		searchFirstSearch()
		sendInput {raw}bn:
		sendInput % ISBN
		send {enter}
return

;▼　■　■　■　■　■　■　■　■　■　■　■　■　Seaching and pulling bibliographic data in WorldCat.
;🌐
	;▼ Title search in WorldCat.
		titleSearchWC:
		send ^n
		sleep 2000
		send !d
		sendInput {raw}https://www.worldcat.org/search?q=ti:
		sleep 100
		sendInput % Title
		sleep 100
		send {enter}
		goTo worldCatResults
;🌐
	;▼ ISBN search on WorldCat.
		ISBNsearchWC:
		send ^n
		;▼ time given to ensure the browser window graphically renders.
		sleep 2000
		send !d
		sendInput {raw}https://www.worldcat.org/search?q=bn:
		sleep 100
		sendInput % ISBN
		sleep 100
		send {delete}
		sleep 100
		send {enter}
		;▼ Time given for WorldCat page to intially load.
		sleep 3000
;🌐
		worldCatResults:
	;◯　Loop to check that WorldCat search results page loads before the macro continues.
		loop, 10
			{
			sleep 1000
			;▼ The tab key is pushed once to "tab out" of the WorldCat search bar. If tab isn't pressed, only the text in the seach bar is selected. After tabbing, all the text on the page is selected.
			send {tab}
			send ^a
			clipboard:= ""
			send ^c
			clipWait, 2
			if inStr(clipboard, "Libraries")
				break
			}
			if !inStr(clipboard, "Libraries")
				{
				msgBox, WorldCat.org searh results have not loaded properly, the macro has stopped.
				exit
				}
;🌐
	;▼ Copy contents of results page to determine if there are 0, 1, or >1 results.
		send ^a
		clipboard:= ""
		send ^c
		clipWait, 2
		if errorLevel
			{
			msgBox, No data was copied from the WorldCat.org search results. The macro has stopped.
			return
			}
		if inStr(clipboard, "Try your search again")
			{
			zero:
			send !{f4}
			confirmSpreadsheet()
			send {home}
			bibArray[16]:= "n/a"
			clipboard:= bibArray[1] . "`t" . bibArray[2] . "`t" . bibArray[3] . "`t" . bibArray[4] . "`t" . bibArray[5] . "`t" . bibArray[6] . "`t" . bibArray[7] . "`t" . bibArray[8] . "`t" . bibArray[9] . "`t" . bibArray[10] . "`t" . bibArray[11] . "`t" . bibArray[12] . "`t" . bibArray[13] . "`t" . bibArray[14] . "`t" . bibArray[15] . "`t" . bibArray[16]
			send ^v
			;▼　This delay is needed for the "down" command on the next line to work correctly, not known why.
			sleep 1000
			send {down}
			sleep 1000
			}
	;▼ Open browser tab for each result.
		if inStr(clipboard, "Find a Copy at a Library")
			{
			send ^f
			sendInput {raw}find a copy
			send {esc}
			sleep 500
			send +{tab}
			send {enter}
			}
		send {tab}			
return

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■■ Load all editions of an item, each edition on a new browser tab.
numpad7::
F7::
		allowMacroToRun()
		if(searchEngine= "FirstSearch"){
			goTo firstSearchEditions
			}
		if(searchEngine= "WorldCat.org")
			goTo worldCatEditions

;▼　■　■　■　■　■　■　■　■　■　■　■　■📚 First Search Editions
;🔎
		firstSearchEditions:
		confirmFirstSearchRecord()
	;▼ Check to make sure the FirstSeach item record page has the text "View all editions and formats"
		copyAllclearClip()
		if !inStr(clipboard, "Search for versions with same title and author")
			{
			msgBox, This FirstSearch page does not have a "Search for versions with same title and author" link.
			return
			}
		send ^f
		sendInput {raw}Search for versions with same title and author
		send {esc}
		send {enter}
		sleep 2000
;🔎
	;◯ Loop to check that a FirstSearch Format and Editions page loads.
		loop, 10{
			sleep 1000
			send {tab}
			send ^a
			clipboard:= ""
			send ^c
			clipWait, 2
			if inStr(clipboard, "Records found:")
				break
			}
			if !inStr(clipboard, "Records found:"){
				msgBox, The "versions with the same title and author" page on FirstSearch has not loaded. The macro has stopped.
				return
				}
;🔎
	;▼ Parse page text to get the number of search results for the loop process below.
		getEditionsCount:= clipboard
		;▼ When there are more than 10 results, default to 9 loops.
		if inStr(getEditionsCount, "Records Found: 10 ")
			{
			getEditionsCount:= 10
			goTo loadEditionsRecords
			}
		;▼ Isolate number of results.
		getEditionsCount:= regExReplace(getEditionsCount, "`r`n|`t", "")
		getEditionsCount:= regExReplace(getEditionsCount, ".*Records found: ")	; Removes all content before results count.
		getEditionsCount:= regExReplace(getEditionsCount, " .*")				; Removes all content after results count.
		getEditionsCount:= getEditionsCount
		if (getEditionsCount > 10)
			getEditionsCount:= 10
;		msgBox, % getEditionsCount
;🔎
		loadEditionsRecords:
		recordUp:= 1
	;◯ Loop	 to check open tabs for each record on a Formats and Editions search result.
		loop, %getEditionsCount%
			{
			send ^f
			send % recordUp
			send .
			send {esc}
			sleep 300
			send {tab}
			send ^{enter}
			recordUp++
			send ^f
			sendInput {raw}advanced options ...
			}
return
	;▼ Closes search result browser tab. Unless Results are 10 or higher.
		if (getEditionsCount= 10)
			goTo skipClosingTab
		send ^w
		skipClosingTab:
return

;▼　■　■　■　■　■　■　■　■　■　■　■　■🌐 WorldCat Editions
		worldcatEditions:
		confirmWorldCatOrgRecord()
	;▼ Check to make sure the WorldCat item record page has the text "View all editions and formats"
		;▼ Cursor will be in the search box, tab one time to move out of the search box so that all the text on the page can be copied.
		send {tab}
		send ^a
		clipboard= ""
		send ^c
		clipWait, 2
		if errorLevel
			{
			msgBox, the clipboard is empty. the macro cannot run.
			return
			}
		if !inStr(clipboard, "View all formats and editions")
			{
			msgBox, This WorldCat page does not have a "View all formats and editions" link.
			return
			}
		send ^f
		sendInput {raw}view all formats and editions
		send {esc}
		send {enter}
		sleep 1000
	;◯ Loop to check that a WorldCat.org Format and Editions page loads.
		loop, 10
			{
			sleep 1000
			send {tab}
			send ^a
			clipboard:= ""
			send ^c
			clipWait, 2
			if inStr(clipboard, "Showing Editions")
				break
			}
			if !inStr(clipboard, "Showing Editions")
				{
				msgBox, The "Showing Editions" page on WorldCat.org has not loaded. The macro has stopped.
				return
				}
	;▼ Parse page text to get the number of search results for the loop process below.
		getEditionsCount:= clipboard
		;▼ When there are more than 10 results, default to 9 loops.
		if inStr(getEditionsCount, "Showing Editions 1 - 10")
			{
			getEditionsCount:= 9
			goTo firstResultForWorldCatEditions
			}
		getEditionsCount:= regExReplace(getEditionsCount, "`r`n|`t", "")
		getEditionsCount:= regExReplace(getEditionsCount, "Return to Item Details.*")	; Removes all content after results count.
		getEditionsCount:= regExReplace(getEditionsCount, ".*Showing Editions")			; Removes all content before results count.
		getEditionsCount:= regExReplace(getEditionsCount, ".* out of ")					; Isolates total count of results.
		getEditionsCount:= getEditionsCount - 1
	;▼ Open a tab for the first result of the Formats and Editions search results.
		firstResultForWorldCatEditions:
		send ^f
		sendInput {raw}Title
		send {enter}
		send {esc}
		send {tab}
		send ^{enter}
	;◯ Loop to open a tab for each result after the first resust.
		loop, %getEditionsCount%
			{
			send {tab 2}
			send ^{enter}
			}
		;▼ Closes search result browser tab.
		send ^w
		sleep 1000
return

;▼　■　■　■　■　■　■　■　■　■　■　■　■🌐 Open the "show more information" dropdown on each tab.
!F7::
!numpad7::
		allowMacroToRun()
	;◯　Loop to open "Show more information" on each browser tab window.
		getEditionsCount:= getEditionsCount + 1
		loop %getEditionsCount%
			{
			send ^f
			sendInput {raw}show more information
			send {esc}
			sleep 100
			send +{tab}
			send {enter}
			send ^{tab}
			sleep 100
			}
return

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■■ Pull bibliographic data to spreadsheet.
f4:: ;Only get subject data.
numpad4::
		allowMacroToRun()
		JustSubject:= "yes"
		JustOCLC:= ""
		if(searchEngine= "FirstSearch")
			goTo firstSearchBibData
		if(searchEngine= "WorldCat.org")
			goTo worldCatBibData

f5:: ; Only get OCLC#
numpad5::
		allowMacroToRun()
		JustOCLC:= "yes"
		JustSubject:= ""
		if(searchEngine= "FirstSearch")
			goTo firstSearchBibData
		if(searchEngine= "WorldCat.org")
			goTo worldCatBibData

f6:: ; Get all bibliographic data
numpad6::
		loopPullBibData:
		allowMacroToRun()
		JustSubject:= ""
		JustOCLC:= ""
		if(searchEngine= "FirstSearch")
			goTo firstSearchBibData
		if(searchEngine= "WorldCat.org")
			goTo worldCatBibData

f9:: ; Subject will equal "n/a".
numpad9::
		allowMacroToRun()
		Subjects:= "n/a"
		JustSubject:= "yes"
		if(searchEngine= "FirstSearch")
			goTo firstSearchBibData
		if(searchEngine= "WorldCat.org")
			goTo worldCatBibData

;▼　■　■　■　■　■　■　■　■　■　■　■　■■ Pull FirstSerach bibliographic data.
	firstSearchBibData:
		confirmFirstSearchRecord()
	;▼ Check if the item might be a duplicate. You must be on your institution's network or using a your institution's VPN if off site. Otherwise the result will always be "no" or "n".
		send {tab}
		clipboard:= ""
		send ^a
		send ^c
		clipWait, 2
		if errorLevel
			{
			msgBox, No data was copied from the WorldCat.org item record. the macro has stopped.
			return
			}
;🔎
		if inStr(clipboard, "FirstSearch indicates your institution owns the item.")
			bibArray[13]:= "y"
		if !inStr(clipboard, "FirstSearch indicates your institution owns the item.")
			bibArray[13]:= "n/a"
;🔎
	;▼ Get the source code of a FirstSearch item entry.
		send ^u
	;◯ Loop to check that a WorldCat.org item record's source code loads before the macro continues.
	loop, 10
		{
		sleep 1000
		send {tab}
		send ^a
		send ^c
		clipWait, 2
		if inStr(clipboard, "<title>FirstSearch")
			break
		}
		if !inStr(clipboard, "<title>FirstSearch")
			{
			msgBox, The source code for a WorldCat.org record has not loaded properly, the macro has stopped.
			exit
			}
;🔎
	;▼ Parse source code of a FirstSearch entry in FirstSearch.
		sourceCode:= clipboard
		sourceCode:= regExReplace(sourceCode, "`r`n|`t")									; Removes all new line carriage returns and tabs.
		sourceCode:= regExReplace(sourceCode, ".*<b>Title:", "<b>Title:")					; Removes code before bibliographic data.
		sourceCode:= regExReplace(sourceCode, ",""<b>Database:")							; Removes code after bibliographic data.
		sourceCode:= regExReplace(sourceCode, "<span class=matchterm(0|1|2|3)>","`r`n")		; Removes code FirstSearch uses to highlight text used to search for entry.
		sourceCode:= regExReplace(sourceCode, "</span>")									; Removes all instances of </span> in code to remove added </span> created by matchterm
		;▼ Removes code that highlights text as a matching term. This additional html needs to be removed because it makes the sourceCode inconsistent.
		sourceCode:= regExReplace(sourceCode, "`r`n","")
		replaceDiacritics()
;🔎
	;▼⚠ Do not run RegEx find and replace on the sourceCode variable. Populate another variable first and manipulate the new variable.

	;▼🔢 ISBN-13 in FirstSearch.
		ISBN13:= sourceCode
		ISBN13:= regExReplace(ISBN13, ".*<b>Standard No:")								; Removes code before ISBNs.
		ISBN13:= regExReplace(ISBN13, "</tr>.*")										; Removes content after ISBNs.
		if !inStr(ISBN13, "ISBN")
			ISBN13:= "n/a"
		ISBN13:= regExReplace(ISBN13, ".*ISBN:</b> ", "[")								; Removes code before ISBNs.
		ISBN13:= regExReplace(ISBN13, "</span>.*|</font>.*|; <.*| <.*", "]")			; Removes code after ISBNs.
		;▼ Remove ISBN 10s.
		ISBN13:= regExReplace(ISBN13, "; ", "][")										; Brackets make explicit when each ISBN begins and ends. Needed for next lines.
		ISBN10:= ISBN13																	
		ISBN13:= regExReplace(ISBN13, "\[..........\]|\[.......... .+?\]", "")			; Removes all ISBN-10s.
		;▲
		ISBN13:= regExReplace(ISBN13, "\]\[", " ^ ")									; Replaces seperator between ISBNs with a caret.
		ISBN13:= regExReplace(ISBN13, "\]|\[")											; Removes intial [ and final ].
		ISBN13:= regExReplace(ISBN13, "\(\(","(")										; Removes double brackets before ISBN.
		ISBN13:= regExReplace(ISBN13, "\)\)",")")										; Removes double brackets after ISBN.
		ISBN13:= regExReplace(ISBN13, "\)\)",")")										; Removes double brackets after ISBN.
		bibArray[14]:= ISBN13
;		msgBox % ISBN13
;🔎
	;▼🔟 ISBN-10 in FirstSearch.
		;▼ Text is already partially prepped from the ISBN13 section above in FirstSearch.
		ISBN10:= regExReplace(ISBN10, "\[.............\]|\[............. .+?\]", "")	; Removes all ISBN-13s.
		ISBN10:= regExReplace(ISBN10, "\]\[", " ^ ")									; Replaces seperator between ISBNs with a caret.
		ISBN10:= regExReplace(ISBN10, "\]|\[")											; Removes intial [ and final ].
		ISBN10:= regExReplace(ISBN10, "\(\(","(")										; Removes double brackets before ISBN.
		ISBN10:= regExReplace(ISBN10, "\(\(","(")										; Removes double brackets before ISBN.
		ISBN10:= regExReplace(ISBN10, "\)\)",")")										; Removes double brackets after ISBN.
		ISBN10:= regExReplace(ISBN10, "\)\)",")")										; Removes double brackets after ISBN.
		bibArray[15]:= ISBN10
;		msgBox % ISBN10
;🔎
	;▼🅾 OCLC# in FirstSearch.
		OCLC:= sourceCode
		OCLC:= regExReplace(OCLC, ".*<b>OCLC:</b> ")					; Removes code before OCLC#.
		OCLC:= regExReplace(OCLC, "<.*")								; Removes code after OCLC# data.
;		msgBox, % OCLC
		bibArray[16]:= OCLC
;🔎
	;🗨 Language in FirstSerach.
		language:= sourceCode
		;▼ Isolate relevant section of source code.
		language:= regExReplace(language, ".*<b>Language:")
		language:= regExReplace(language, "</tr>.*")
		;▼ Remove code after data.
		language:= regExReplace(language, "&nbsp.*")
		;▼ Remove code before data.
		language:= regExReplace(language, ".*>")
;		msgbox, % language
;🔎
	;▼📔🅰 Romanized title in FirstSearch.
		titleR:= sourceCode
		;▼ Isolate relevant section of source code.
		titleR:= regExReplace(titleR, ".*<b>Title:")
		titleR:= regExReplace(titleR, "</tr>.*")
		titleN:= titleR
		;▼ Remove code before data.
		titleR:= regExReplace(titleR, ".*</div><br>|.*</div>")
		;▼ Remove code between tile and subtitle
		titleR:= regExReplace(titleR, ":<br>|\.<br>", ": ")
		;▼ Remove code after data.
		titleR:= regExReplace(titleR, " /<.*|(,|\.|)</b>.*")
		;▼ Formatting touch ups.
		titleR:= regExReplace(titleR, " : ", ": ")
		titleR:= regExReplace(titleR, "=<br>","= ")
		titleR:= regExReplace(titleR, ";<br>","; ")
		titleR:= regExReplace(titleR, "--<br>","-- ")
		titleR:= regExReplace(titleR, "--</b>.*"," --")
		titleR:= regExReplace(titleR, ",<br>",", ")
		titleR:= regExReplace(titleR, "<br>"," ")
		titleR:= regExReplace(titleR, "  "," ")
		bibArray[17]:= titleR
;		msgBox, % titleR
;🔎
	;▼📔🈂 Non-Romanized title in FirstSearch.
		;▼ Text is already partially prepped from the titleR section above.
		if !inStr(titleN, "vernacular lang")
			titleN:= "n/a"
		;▼ Remove code before data.
		titleN:= regExReplace(titleN, ".*lang="".."">")
		;▼ Remove code after data.
		titleN:= regExReplace(titleN, " /.*|(.|)</div>.*")
		;▼ Formatting touch ups.
		titleN:= regExReplace(titleN, " : ", ": ")
		bibArray[18]:= titleN
;		msgBox % titleN
;🔎
	;▼✒🅰 Romanized creator in FirstSearch.
		authorR:= sourceCode
		if !inStr(authorR, "<b>Author")
			authorR:= "n/a"
		;▼ Isolate relevant section of source code.
		authorR:= regExReplace(authorR, ".*<b>Author\(s\):")
		authorR:= regExReplace(authorR, "</tr>.*")
		authorN:= authorR
		;▼ Remove code before data.
		authorR:= regExReplace(authorR, ".*</div>")
		;▼ Remove code between multiple authors.
		authorR:= regExReplace(authorR, "\&nbsp;|<br>|<a href.+?>")
		authorR:= regExReplace(authorR, "\.</a>;|\.</a>|</a>"," ^ ")
		authorR:= regExReplace(authorR, "  "," ")							; Replaces double spaces with single space created by previous line of code.
		;▼ Remove code after data.
		authorR:= regexreplace(authorR, ".*serif"">")						; Works in tandem with line of code below. Very rarely this needs to be done for the code below to work (approximately 8 times out of 650)
		authorR:= regexreplace(authorR, "\.</font>.*|</font>.*")
		authorR:= regExReplace(authorR, " \^ $|,</b>$|,$")
		;▼ Formatting touch ups.
		authorR:= regExReplace(authorR, ",;",",")							; Fixes punctuation between author and their birth/death dates.
;		msgBox, % authorR
		;▼ Dedupe identical Romanized authors.
		authorR:= regExReplace(authorR, " \^ ","`r`n")
		fileAppend, %authorR%, authorR.txt, utf-8
		authorR:= ""
		loop, read, authorR.txt
			{
			if !inStr(authorR, A_LoopReadLine)
				authorR .= A_LoopReadLine . "`r`n"
			}
		fileDelete, authorR.txt
		authorR:= regExReplace(authorR, "`r`n", " ^ ")
		authorR:= regExReplace(authorR, " \^ $")
		bibArray[19]:= authorR
;🔎
	;▼✒🏦🅰 Romanized corporate creator in FirstSearch.
		corpAuthorR:= sourceCode
		if !inStr(corpAuthorR, "<b>Corp Author")
			{
			corpAuthorR:= ""
			goTo skipCorpAuthorR
			}
		;▼ Isolate relevant section of source code.
		corpAuthorR:= regExReplace(corpAuthorR, ".*<b>Corp Author\(s\):")
		corpAuthorR:= regExReplace(corpAuthorR, "</tr>.*")
		corpAuthorN:= corpAuthorR
		if !inStr(corpAuthorR, "a href")
			{
			corpAuthorR:= "n/a"
			if !inStr(authorR, "n/a")
				{
				corpAuthorR:= ""
				goto skipCorpAuthorR
				}
			}	
		;▼ Remove code before data.
		corpAuthorR:= regExReplace(corpAuthorR, ".*<a href.+?>")
		;▼ Remove code between multiple corporate authors.
		;▼ no examples yet.
		;▼ Remove code after data.
		corpAuthorR:= regExReplace(corpAuthorR, "\.</a>.*|</a>.*")
		;▼ Formatting touch ups.
		corpAuthorR:= regExReplace(corpAuthorR, "\.; " ," ^ ")				; Replaces multiple author seperator with " ^ " delimiter.
		if inStr(authorR, "n/a")											; If there is ONLY a Romanized corporate author.
			{
			bibArray[19]:= corpAuthorR
			goTo skipCorpAuthorR
			}
		bibArray[19]:= authorR . " ^ " . corpAuthorR
		skipCorpAuthorR:
;		msgBox, % corpAuthorR
;🔎
	;▼✒🈂 Non-Romanized creator in FirstSearch.
	;▼ Text is already partially prepped from the authorR section above.
		if !inStr(authorN, "vernacular lang")
			authorN:= "n/a"
		;▼ Remove code before data.
		authorN:= regExReplace(authorN, "^.+?lang="".."">")
		;▼ Remove code between multiple authors.
		authorN:= regExReplace(authorN, " <span.+?>|<span.+?>", " ")
		authorN:= regExReplace(authorN, ", ", "zvz")						; delimits non-Romanized author names. including two lines below.
			authorN:= regExReplace(authorN, " ", " ^ ")
			authorN:= regExReplace(authorN, "zvz", ", ")
		;▼ Remove code after data.
		authorN:= regExReplace(authorN, "\.</div>.*|</div>.*")
		authorN:= regExReplace(authorN, "\,$")
		;▼ Formatting Touch ups.
		authorN:= regExReplace(authorN, "- \^ \)","- )")
		authorN:= regExReplace(authorN, "\. \^"," ^")
		authorN:= regExReplace(authorN, "\.$",",$")
;		msgBox % authorN
	;▼ Dedupe identical non-Romanized authors.
		authorN:= regExReplace(authorN, " \^ ","`r`n")						; Fixes punctuation between author and their birth/death dates.	
		fileAppend, %authorN%, authorN.txt, utf-8
		authorN:= ""
		loop, read, authorN.txt
			{
			if !inStr(authorN, A_LoopReadLine)
				authorN .= A_LoopReadLine . "`r`n"
			}
		fileDelete, authorN.txt
		authorN:= regExReplace(authorN, "`r`n", " ^ ")
		authorN:= regExReplace(authorN, " \^ $")
		bibArray[20]:= authorN
		;msgBox, % authorN
;🔎
	;▼✒🏦🈂 Non-Romanized corporate creator in FirstSearch.
	;▼ Text is already partially prepped from the corpAuthorR section above.
		if !inStr(corpAuthorN, "vernacular lang")
			{
			corpAuthorN:= ""
			goTo skipCorpAuthorN
			}
		;▼ Remove code before data.
		corpAuthorN:= regExReplace(corpAuthorN, ".*lang="".."">")
		;▼ Remove code between multiple corporate authors.
		corpAuthorN:= regExReplace(corpAuthorN, "\. ", " ^ ")
		;▼ Remove code after data.
		corpAuthorN:= regExReplace(corpAuthorN, "\.</div>.*|</div>.*")
		if inStr(authorN, "n/a")													; If there is ONLY a non-Romanized corporate author.
			{
			bibArray[20]:= corpAuthorN
			goTo skipCorpAuthorN
			}
		bibArray[20]:= authorN . " ^ " . corpAuthorN
		skipCorpAuthorN:
;		msgBox, % corpAuthorN
;🔎
	;▼📚🅰 Romanized series title in FirstSearch.
		seriesR:= sourceCode
		;▼ Isolate relevant section of source code.
		if !inStr(seriesR, "<b>Series:")
			seriesR:= "n/a"
		seriesR:= regExReplace(seriesR, ".*<b>Series:")
		seriesR:= regExReplace(seriesR, "</tr>.*")
		seriesN:= seriesR
		;▼ Remove code before data.
		seriesR:= regExReplace(seriesR, ".*</div>")
		;▼ Remove code ater data.
		seriesR:= regExReplace(seriesR, ".*serif"">")						; Works in tandem with line of code below. Very rarely this needs to be done for the code below to work.
		seriesR:= regExReplace(seriesR, "`; <b>Variation.*|</font>.*")
		seriesR:= regExReplace(seriesR, "\.$")
		;▼ Formatting touch ups.
		seriesR:= regExReplace(seriesR, " <b>Variation:</b> ")
		seriesR:= regExReplace(seriesR, " `;`; ", " `; ")
		bibArray[21]:= seriesR
;		msgBox % seriesR
;🔎
	;▼📚🈂 Non-Romanized series title in FirstSearch.
	;▼ Text is already partially prepped from the seriesR section above.
		if !inStr(seriesN, "vernacular lang")
			seriesN:= "n/a"
		;▼ Remove code before data.
		seriesN:= regExReplace(seriesN, ".*lang="".."">")
		;▼ Remove code ater data.
		seriesN:= regExReplace(seriesN, "</div>.*")
		seriesN:= regExReplace(seriesN, "\.$")
		seriesN:= regExReplace(seriesN, "`; `;", "`;")
		;▼ Formatting touch ups.
		;no examples yet
		bibArray[22]:= seriesN
;🔎
	;▼📖🅰 Romanized publisher in FirstSearch.
		publisherR:= sourceCode
		if !inStr(publisherR, "<b>Publication:")
			publisherR:= "n/a"
		;▼ Isolate relevant section of source code.
		publisherR:= regExReplace(publisherR, ".*<b>Publication:")
		publisherR:= regExReplace(publisherR, "</tr>.*")
		publisherN:= publisherR
		;▼ Remove code before data.
		publisherR:= regExReplace(publisherR, ".*</div>")					; Removes content before where a Romanized publisher name would be. Allows to check if there is no Romanization at all and mark publisher as "n/a"\\
		if !inStr(publisherR, " : ")
			publisherR:= "n/a"
		publisherR:= regExReplace(publisherR, ".* : ")
		;▼ Remove code after data.
		publisherR:= regExReplace(publisherR, ",.*")
		;▼ Formatting touch ups.
		;no examples yet.
		bibArray[23]:= publisherR
;		msgBox, % publisherR
;🔎
	;▼📖🈂 Non-Romanized publisher in FirstSearch.
	;▼ Text is already partially prepped from the publisherR section above.
		if !inStr(publisherN, "vernacular lang")
			publisherN:= "n/a"
		;▼ Remove code after data.
		publisherN:= regExReplace(publisherN, "</div>.*| <b>Edition:.*|Edition:.*")
		if !inStr(publisherN, " : ")
			publisherN:= "n/a"
		;▼ Remove code before data.
		publisherN:= regExReplace(publisherN, "^.+?lang="".."">")
		publisherN:= regExReplace(publisherN, ".* : ")
		;▼ Remove code before data again.
		publisherN:= regExReplace(publisherN, ",.*")
		;▼ Formatting touch ups.
		publisherN:= regExReplace(publisherN, ".*</b></font></b></td><td><font size=2 face=""Arial.*", "n/a")	; Some records with no Romanized publisher end produce garbled code that look like this.
		bibArray[24]:= publisherN
;		msgBox, % publisherN
;🔎	
	;▼📅 Year of publication in FirstSearch.
		pubDate:= sourceCode
		if !inStr(pubDate, "<b>Year:")
			pubDate:= "n/a"
		;▼ Isolate relevant section of source code.
		pubDate:= regExReplace(pubDate, ".*<b>Year:")
		pubDate:= regExReplace(pubDate, ".*<nobr>|</nobr>.*")
		;▼ Remove code after data.
		pubDate:= regExReplace(pubDate, "</font></td>.*")
		;▼ Remove code before data.
		pubDate:= regExReplace(pubDate, ".*>")
		;▼ Formatting touch up.
		;no example yet.
		bibArray[25]:= pubDate
;		msgBox, % pubDate
;🔎
	;▼✳ Subject heading in FirstSearch.
		subjects:= sourceCode
		if !inStr(subjects, "SUBJECT(S)")
			subjects:= "n/a"
		;▼ Isolate relevant section of source code.
		subjects:= regExReplace(subjects, ".*>SUBJECT\(S\)")
		subjects:= regExReplace(subjects, "<b>Class.*|<b>Note\(s\).*|<b>Responsibility.*|<b>Document.*")
		;▼ Remove code between multiple subjects.
		subjects:= regExReplace(subjects, "<a h.+?>","`r`n")				; Each subject put on a new line.
		subjects:= regExReplace(subjects, ".*lang="".."".*`r`n")			; Removes non-Roman subjects.
		;▼ Remove code after data.
		subjects:= regExReplace(subjects, "</a>.*")
		subjects:= regExReplace(subjects, "m)\.$")
		;▼ Remove code before data.
		subjects:= regExReplace(subjects, "<.*>")
		;▼ Formatting touch up.
		subjects:= regExReplace(subjects, "\(88.-..\) ")					; Removes 880 encoding note.
	;▼ Dedupe identical subject headings.
		fileAppend, %subjects%, subjects.txt, utf-8
		subjects:= ""
		loop, read, subjects.txt
			{
			if !inStr(subjects, A_LoopReadLine)
				subjects .= A_LoopReadLine . "`r`n"
			}
		fileDelete, subjects.txt
		subjects:= regExReplace(subjects, "---", ", ")						; returns comma with space
		subjects:= regExReplace(subjects, "--", " -- ")						; returns spaces around --
		subjects:= regExReplace(subjects, "`r`n", " ^ ")
		subjects:= regExReplace(subjects, " \^ $")
		bibArray[26]:= subjects
;		msgBox, % subjects
;🔎
	;▼ Debug message box.
	if (debug= "yes")
		{
		send ^w
		msgBox, 🔢ISBN13:`r`n%ISBN13%`r`n`r`n🔟ISBN10:`r`n%ISBN10%`r`n`r`n🅾OCLC#:`r`n%OCLC%`r`n`r`n📜(📰)Doc Type:`r`n%docType% (%matType%)`r`n`r`n📔🅰Title Romanized/English:`r`n%titleR%`r`n`r`n📔🈂Title non-Romanized`r`n%titleN%`r`n`r`n✒🅰Creator Romanized/English:`r`n%authorR%`r`n`r`n✒🏦🅰Corporate Author Non-Romanized`r`n%corpAuthorR%`r`n`r`n✒🈂Creator non-Romanized:`r`n%authorN%`r`n`r`n✒🏦🈂Corporate Author Non-Romanized`r`n%corpAuthorN%`r`n`r`n🅰📚Series Title Romanized/English:`r`n%seriesR%`r`n`r`n🈂📚Series Title non-Romanized:`r`n%seriesN%`r`n`r`n📖🅰Publisher Romanized/English:`r`n%publisherR%`r`n`r`n📖🈂Publisher non-Romanized:`r`n%publisherN%`r`n`r`n📅Year of Publication:`r`n%pubdate%`r`n`r`n✳Subject Headings:`r`n%subjects%
		send !{f4}
		goTo pasteToSpreadsheet
		}
;🔎	
	send !{f4}
	goTo pasteToSpreadsheet

;▼　■　■　■　■　■　■　■　■　■　■　■　■■ Pull WorldCat bibliographic data.
	worldCatBibData:
		;confirmWorldCatOrgRecord()

loop {
copyAllclearClip()
sleep 3000
if inStr(clipboard, "find a copy at a library")
break
}

	;▼ Check if the item might be a duplicate. You must be on your institution's network or using a your institution's VPN if off site. Otherwise the result will always be "no" or "n".
		send {tab}
		clipboard:= ""
		send ^a
		send ^c
		clipWait, 2
		if errorLevel
			{
			msgBox, No data was copied from the WorldCat.org item record. the macro has stopped.
			return
			}
;🌐

		if inStr(clipboard, libName)
			bibArray[13]:= "y"
		if !inStr(clipboard, libName)
			bibArray[13]:= "n/a"
;🌐
	;▼ Get the source code of a WorldCat.org item entry.
	send ^u	
	;◯ Loop to check that a WorldCat.org item record's source code loads before the macro continues.
	loop, 10
		{
		sleep 1000
		send {tab}
		send ^a
		send ^c
		clipWait, 2
		if inStr(clipboard, "oclcNumber")
			break
		}
		if !inStr(clipboard, "oclcNumber")
			{
			msgBox, The source code for a %searchEngine% record has not loaded properly, the macro has stopped.
			exit
			}
;🌐
	;▼ Parse source code of a WorldCat entry.
		sourceCode:= clipboard
		sourceCode:= regExReplace(sourceCode, "`r`n")							; Removes all new line carriage returns.
		sourceCode:= regExReplace(sourceCode, "`t")								; Removes all tabs.
		sourceCode:= regExReplace(sourceCode, ".*application/json"">")			; Removes content before bibliographic data.
		sourceCode:= regExReplace(sourceCode, ",""coverArtUrl.*")				; Removes content after bibliographic data.
		
		replaceDiacritics()
;🌐
	;▼⚠ Do not run RegEx find and replace on the sourceCode variable. Populate another variable first and manipulate the new variable.
	
	;▼🔢 ISBN-13.
		ISBN13:= sourceCode
		if inStr(ISBN13, """isbns"":null")
			ISBN13:= "n/a"
		ISBN13:= regExReplace(ISBN13, "sourceIsbns")							; Needs to be removed to check if ISBN metadata is null.
		ISBN13:= regExReplace(ISBN13, ".*isbns"":\[""")							; Removes content before ISBNs.
		ISBN13:= regExReplace(ISBN13, """\].*", " \]")							; Removes content after ISBNs. Keeps final ] so 2 lines down ISBN-10s can be removed.
		ISBN13:= regExReplace(ISBN13, """,""", " ^ ")							; Replaces seperator between ISBNs with a caret.
		ISBN13:= regExReplace(ISBN13, " \^ .......... .*", "")					; Removes all ISBN-10s.
		bibArray[14]:= ISBN13
;		msgBox, % ISBN13
;🌐
	;▼🔟 ISBN-10.
		ISBN10:= sourceCode
		if inStr(ISBN10, """isbns"":null")
			ISBN10:= "n/a"
		ISBN10:= regExReplace(ISBN10, "sourceIsbns")							; Needs to be removed for replace command on next line to work properly.
		ISBN10:= regExReplace(ISBN10, ".*isbns"":\[""", "\[ ")					; Removes content before ISBNs. Keeps initial [ so 3 lines down ISBN-13s can be removed.
		ISBN10:= regExReplace(ISBN10, """\].*")									; Removes content after ISBNs.
		ISBN10:= regExReplace(ISBN10, """,""", " ^ ")							; Replaces seperator between ISBNs with a caret.
		ISBN10:= regExReplace(ISBN10, ".* ............. \^ ", "")				; Removes all ISBN-13s.
		bibArray[15]:= ISBN10
;		msgBox, % ISBN10 
;🌐
	;▼🅾 OCLC#.
		OCLC:= sourceCode
		OCLC:= regExReplace(OCLC, "recordControlOclcNumbers.*")					; Needs to be removed for replace command on next line to work properly.
		OCLC:= regExReplace(OCLC, ".*""oclcNumber"":""")						; Removes content before OCLC Number.
		OCLC:= regExReplace(OCLC, """.*")
		bibArray[16]:= OCLC
;		msgBox, % OCLC		
;🌐
	;🗨 Language in FirstSerach.
		language:= sourceCode
		language:= regExReplace(language, "^.*?languageCode"":""")
		language:= regExReplace(language, """.*")
;		msgbox, % language
;🌐
	;▼📔🈂 Non-Romanized title.
		titleN:= sourceCode
		titleN:= regExReplace(titleN, ".*""titleInfo.+?"":")					; Removes content before vernacular title.
		titleN:= regExReplace(titleN, "\}.*")									; Removes content after vernacular title.
		if !inStr(titleN, "romanizedText")
			titleN:= "n/a"
		titleN:= regExReplace(titleN, ".*romanizedText"":""")					; Removes content before vernacular title.
		titleN:= regExReplace(titleN, """.*")									; Removes content after vernacular title.
		bibArray[17]:= titleN
;		msgBox, % titleN
;🌐
	;▼📔🅰 Romanized title.
		titleR:= sourceCode
		if !inStr(titleR, """titleInfo"":")
			titleR:= "n/a"
		titleR:= regExReplace(titleR, ".*titleInfo"":\{""text"":""")			; Removes content before romanized title.
		titleR:= regExReplace(titleR, """.*")									; Removes content before romanized title.
		bibArray[18]:= titleR
;		msgBox, % titleR
;🌐
	;▼✒🅰 Romanized author/creator.
		authorR:= sourceCode
		;▼ Isolate relevant section of source code.
		authorR:= regExReplace(authorR, ".*contributors"":\[")
		authorR:= regExReplace(authorR, "]},""error"".*")
		if inStr(authorR, "null")
			authorR:= "n/a"
		;▼ Author types put on seperate lines.
		authorR:= regExReplace(authorR, "\{""firstName""", "`r`nfirstName""")
		authorR:= regExReplace(authorR, "\{""nonPersonName""", "`r`nnonPersonName""")
		;▼ Code cleanup.
		authorR:= regExReplace(authorR, "^`r`n") ;Remove first CR/NL.
		authorR:= regExReplace(authorR, """},""isPrimary.*")
		authorR:= regExReplace(authorR, "\[|]")
		authorN:= authorR ;Up to this point the process is identical for non-Roman authors, see authorN section below.
		;msgBox, % authorR
		authorR:= regExReplace(authorR, "nonPersonName.*""text"":""")
		;▼ Code clean up between first and second names with romanized text.
		authorR:= regExReplace(authorR, "firstName.+?romanizedText"":""")
		authorR:= regExReplace(authorR, """,""languageCode.+?romanizedText"":""", " -- ")
		authorR:= regExReplace(authorR, ".*""romanizedText"":""") ; remove code for non-Romanized nonPerson names.
		authorR:= regExReplace(authorR, ""","".*")
		;▼ Swap name order to CJK order.
		if inStr(sourceCode, "languageCode"":""JA")
			authorR:= RegExReplace(authorR, "mU)^(.*) -- ([^ -- ]*)$", "$2 $1")
		;▼ Code cleanup: author.
		authorR:= RegExReplace(authorR, "firstName.+?"":""")
		authorR:= RegExReplace(authorR, """}.*"":""", " ")
		;▼ Formatting touch up.
		authorR:= regExReplace(authorR, " = .*") ;Remove translated authors.
		authorR:= regExReplace(authorR, "m)\.$") ;Remove trailing period.
		authorR:= regExReplace(authorR, "m) cho$| hen$| hencho$| jutsu$| kanshū$") ;Remove common CJK suffixes for author/editor
		authorR:= regExReplace(authorR, "m) and others") ;Remove common CJK suffixes for author/editor
		;▼ Formatting touch up, dates as a part of author names.
		authorR:= regExReplace(authorR, "m), ....-....$|, ....-$")
		authorR:= regExReplace(authorR, ", ....-.... \^ |, ....- \^ ", " ^ ")
		authorR:= regExReplace(authorR, "m)\(....- \)$|\(....-....\)$")
		authorR:= regExReplace(authorR, "\(....- \) \^ |\(....-....\) \^ ", " ^ ")
		;▼ Dedupes author names
		fileAppend, %authorR%, authorR.txt, utf-8
		authorR:= ""
		loop, read, authorR.txt
			{
			if !inStr(authorR, A_LoopReadLine)
				authorR .= A_LoopReadLine . "`r`n"
			}
		fileDelete, authorR.txt
		;▼ Author names delimited with " ^ ".
		authorR:= regExReplace(authorR, "`r`n", " ^ ")
		authorR:= regExReplace(authorR, " \^ $")
		bibArray[19]:= authorR
		;msgBox, % authorR
;🌐
	;▼✒🈂 Non-Romanized author/creator.
	;▼ Text is partially prepped from the AuthorR section above.
		;▼ Code clean up: non-person author.
		authorN:= regExReplace(authorN, "nonPersonName.*text"":""")
		;▼ Code clean up
		authorN:= regExReplace(authorN, ""","".*")
		authorN:= regExReplace(authorN, ".*"":""")
		;▼ Formatting touch up.
		authorN:= regExReplace(authorN, " = .*") ;Remove translated authors.
		authorN:= regExReplace(authorN, "m)\.$") ;Remove trailing period.
		authorN:= regExReplace(authorN, "編者|著|編|監修") ;Remove common CJK suffixes for author/editor.
		authorN:= regExReplace(authorN, "^henshū |^編集 ") ;Remove common CJK suffixes for author/editor.
		authorN:= regExReplace(authorN, ", author |, author")
		authorN:= regExReplace(authorN, " and others")
		;▼ Formatting touch up, dates as a part of author names.
		authorN:= regExReplace(authorN, ", ....-.... \^ |, ....- \^ ", " ^ ")
		authorN:= regExReplace(authorN, "m), ....-....$|, ....-$")
		authorN:= regExReplace(authorN, "m)\(....- \)$|\(....-....\)$")
		authorN:= regExReplace(authorN, "\(....- \) \^ |\(....-....\) \^ ", " ^ ")
		;▼ Dedupes author names
		fileAppend, %authorN%, authorN.txt, utf-8
		authorN:= ""
		loop, read, authorN.txt
			{
			if !inStr(authorN, A_LoopReadLine)
				authorN .= A_LoopReadLine . "`r`n"
			}
		fileDelete, authorN.txt
		;▼ Author names delimited with " ^ ".
		authorN:= regExReplace(authorN, "`r`n", " ^ ")
		authorN:= regExReplace(authorN, " \^ $")
		bibArray[20]:= authorN
		;msgBox, % authorN
;🌐
	;▼📚🅰 Romanized series title parsed from WorldCat source code.
		seriesR:= sourceCode
		if inStr(seriesR, """series"":null")
			seriesR:= "n/a"
		seriesR:= regExReplace(seriesR, ".*""series"":""")						; Removes content before series title.
		seriesR:= regExReplace(seriesR, ""","".*")								; Removes content after series title.
		bibArray[21]:= seriesR
;		msgBox, % seriesR
;🌐
	;▼📚🈂 Non-Romanized series title parsed from WorldCat source code.
		seriesN:= "n/a in WorldCat.org"
		bibArray[22]:= "n/a"
;		msgBox, % seriesN
;🌐
	;▼📖🅰 Romanized/English publisher parsed from WorldCat source code.
		publisherR:= sourceCode
;		if inStr(publisherR, "")
;			publisherR:= "n/a"
		publisherR:= regExReplace(publisherR, ".*publisherName.+?romanizedText"":""")
		publisherR:= regExReplace(publisherR, """.*")
		bibArray[23]:= publisherR
;		msgBox, % publisherR
;🌐
	;▼📖🔧 Fix Japanese Romanized publisher names.
		if inStr(sourceCode, "romanizedText"){
		fixPub:= publisherR
		stringLower, fixPub, fixPub
		fixRomanizedPublisherNames()
		publisherR:= fixPub
		}
		bibArray[23]:= publisherR
;		msgBox, % publisherR
;🌐
	;▼📖🈂 Non-Romanized publisher parsed from WorldCat source code.
		publisherN:= sourceCode
;		if inStr(publisherN, """publisher"":null""")
;			PubliserN:= "n/a"
		publisherN:= regExReplace(publisherN, ".*publisherName.+?text"":""")
		publisherN:= regExReplace(publisherN, """.*")
		bibArray[24]:= publisherN
;		msgBox, % publisherN
;🌐
	;▼📅 Date of publication parsed from WorldCat source code.
		pubDate:= sourceCode
		if inStr(pubDate, """publisher"":""null""")
			pubDate:= "n/a"
		pubDate:= regExReplace(pubDate, ".*publicationDate"":""")			; Removes content before publisher name.
		pubDate:= regExReplace(pubDate, """.*")								; Removes content after publisher name.
		;▼ Formatting touch up.
		pubDate:= regExReplace(pubDate, "©")
		pubDate:= regExReplace(pubDate, ".*\[")
		pubDate:= regExReplace(pubDate, "].*")
		bibArray[25]:= pubDate
;		msgBox, % pubDate
;🌐
	;▼✳ Subject heading parsed from WorldCat source code.
		subjects:= sourceCode
		if inStr(subjects, "subjectsText"":null")
			subjects:= "n/a"
		subjects:= regExReplace(subjects, ".*""subjectsText"":")				; Removes content before subjects.
		subjects:= regExReplace(subjects, ",""cartographicData.*")				; Removes content after subjects.
		subjects:= regExReplace(subjects, """,""", " ^ ")						; Replaces seperator between subject headings with a caret.
		subjects:= regExReplace(subjects, "\[""|""\]")							; Removes brackets and quotes at beginning and end of subject headings.
		;▼ Parse and append file to dedupe subject headings.
		subjects:= regExReplace(subjects, " \^ ", "`r`n")
		subjects:= regExReplace(subjects, " -- ", "--")							; temporarily removes spaces around -- to facilitate deduping identical subject headings.
		subjects:= regExReplace(subjects, ", ", "---")							; temporarily removes commas with a space to facilitate deduping identical subject headings.
		fileAppend, %subjects%, subjects.txt, utf-8
		subjects:= ""
		loop, read, subjects.txt
			{
			if !inStr(subjects, A_LoopReadLine)
				subjects .= A_LoopReadLine . "`r`n"
			}
		fileDelete, subjects.txt
		subjects:= regExReplace(subjects, "---", ", ")							; returns comma with space
		subjects:= regExReplace(subjects, "--", " -- ")							; returns spaces around --
		subjects:= regExReplace(subjects, "`r`n", " ^ ")
		subjects:= regExReplace(subjects, " \^ $")
		;▲
		bibArray[26]:= subjects
;		msgBox, % subjects
;🌐
	;▼🔀 Becuase of how WorldCat.org codes "text" vs. "romanized text", records with no romanization will appear in "non-Romanized" data variables/spreadsheet columns. This code moves roman script object data to the correct bibArray variables.
		if !inStr(sourceCode, "romanizedText"){
			;▼
			bibArray[17]:= titleR
			bibArray[18]:= "n/a"
			;▼
			bibArray[19]:= authorR
			bibArray[20]:= "n/a"
			;▼
			bibArray[21]:= seriesR
			bibArray[22]:= "n/a"
			;▼
			bibArray[23]:= publisherN
			bibArray[24]:= "n/a"
		}
			
;🌐
	;▼ Debug message box.
	if (debug= "yes")
		{
		send ^w
		msgBox, 🔢ISBN13:`r`n%ISBN13%`r`n`r`n🔟ISBN10:`r`n%ISBN10%`r`n`r`n🅾OCLC#:`r`n%OCLC%`r`n`r`n📔🅰Title Romanized/English:`r`n%titleR%`r`n`r`n📔🈂Title non-Romanized`r`n%titleN%`r`n`r`n✒🅰Creator Romanized/English:`r`n%authorR%`r`n`r`n✒🏦🅰Corporate Author Non-Romanized`r`n%corpAuthorR%`r`n`r`n✒🈂Creator non-Romanized:`r`n%authorN%`r`n`r`n✒🏦🈂Corporate Author Non-Romanized`r`n%corpAuthorN%`r`n`r`n🅰📚Series Title Romanized/English:`r`n%seriesR%`r`n`r`n🈂📚Series Title non-Romanized:`r`n%seriesN%`r`n`r`n📖🅰Publisher Romanized/English:`r`n%publisherR%`r`n`r`n📖🈂Publisher non-Romanized:`r`n%publisherN%`r`n`r`n📅Year of Publication:`r`n%pubdate%`r`n`r`n✳Subject Headings:`r`n%subjects%
		send !{f4}
		goTo pasteToSpreadsheet
		}
;🌐
	;▼ Close browser window and go back to the spreadsheet.
		send !{f4}
		sleep 100
;🌐	
		goTo pasteToSpreadsheet

;▼　■　■　■　■　■　■　■　■　■　■　■　■📊 Paste bibliographic data to spreadsheet.
		pasteToSpreadsheet:
		confirmSpreadsheet()
		;▼ Focus in Excel sometimes ends up in the "alt" menu navigations. Returns focus to the spreadsheet.
		send {esc} ; test in google sheets, office 365
		send {home}
		
	;▼ Language of the item is identified from source code.
		;▼ Detecting the lnaguage affects what kind of link is generated in column g "url for rush order".
		;▼ bibArray[15] is the column for ISBN-10s. on the Amazon site, these unique urls only work with ISBN-10s.
		if (language= "English" OR "en")
			bibArray[8]:= "https://www.amazon.com/dp/" . bibArray[15]
		if (language= "Japanese" OR "jpn")
			bibArray[8]:= "https://www.amazon.jp/dp/" . bibArray[15]
		;▼ no Korean source yet.
		if (language= "Korean")
			bibArray[8]:= "n/a"
		if bibArray[15]= "n/a"
			bibArray[8]:= "n/a"
		bibArray[8]:= regExReplace(bibArray[8], " ", "")
		bibArray[8]:= regExReplace(bibArray[8], ".*no-isbn.*","")

		;▼ If this condition is met, than the macro will only paste subject data in the 29th column (AC).
		if JustSubject= yes
			{
			clipboard:= bibArray[26]
			sendInput {right 25}
			sleep 500
			send ^v
			sleep 1000
			send {down}
			return
			}

	;▼ Move data parsed from WorldCat entry source code to the bibArray variables.
		;▼ If this condition is met, than the macro will only paste data into the first 15 columns of the spreadsheet, the 15th column (O) being for the OCLC#.
		if (JustOCLC= "yes"){
			;▼ Paste bib data columns up to OCLC only. 
			clipboard:= bibArray[1] . "`t" . bibArray[2] . "`t" . bibArray[3] . "`t" . bibArray[4] . "`t" . bibArray[5] . "`t" . bibArray[6] . "`t" . bibArray[7] . "`t" . bibArray[8] . "`t" . bibArray[9] . "`t" . bibArray[10] . "`t" . bibArray[11] . "`t" . bibArray[12] . "`t" . bibArray[13] . "`t" . bibArray[14] . "`t" . bibArray[15] . "`t" . bibArray[16]
			goTo pasteBibData
			}
		clipboard:= bibArray[1] . "`t" . bibArray[2] . "`t" . bibArray[3] . "`t" . bibArray[4] . "`t" . bibArray[5] . "`t" . bibArray[6] . "`t" . bibArray[7] . "`t" . bibArray[8] . "`t" . bibArray[9] . "`t" . bibArray[10] . "`t" . bibArray[11] . "`t" . bibArray[12] . "`t" . bibArray[13] . "`t" . bibArray[14] . "`t" . bibArray[15] . "`t" . bibArray[16] . "`t" . bibArray[17] . "`t" . bibArray[18] . "`t" . bibArray[19] . "`t" . bibArray[20] . "`t" . bibArray[21] . "`t" . bibArray[22] . "`t" . bibArray[23] . "`t" . bibArray[24] . "`t" . bibArray[25] . "`t" . bibArray[26]
		;▼ Added because the deduping process for Authors seems to add a new line carriage return that cannot be removed, but the line below removes it.
;		clipboard:= regExReplace(clipboard, "`r`n")
		
	;▼ Paste the bibliographic data to the spreadsheet.
		pasteBibData:
		sleep 1000
		send {home}
		send ^v

	;▼ Prepare next row in the spreadsheet.
		sleep 1000 ;this delay is needed for the "down" command on the next line to work correctly, not known why.
		send {down}

	;▼ Loop mode settings
		if (loopMode= "on")
			goTo, doLoop
return

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■⏩ Fast browsing navigation tools for active windows.
numpad0::
	;▼ Moves between tabs of active browser window.
		if winActive("ahk_exe firefox.exe")
			send ^{tab}
		if winActive("ahk_exe chrome.exe")
			send ^{tab}
		if winActive("ahk_exe msedge.exe")
			send ^{tab}
return

numpadSub::
	;▼ Closes tabs of active browser window.
		if winActive("ahk_exe firefox.exe")
			send ^w
		if winActive("ahk_exe chrome.exe")
			send ^w
		if winActive("ahk_exe msedge.exe")
			send ^w
return

;▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼ General Functions

;▼▲▼ Detemines if Firefox, Chrome, or Edge has an open window and activates the window.


;▼▲▼ Copy all text in a window, webpage, or text field.
	copyAllclearClip(){
		send ^a
		clipboard:= ""
		send ^c
		clipWait, 2
;		if errorLevel{
;			msgBox, The clipboard did not copy any content. The macro has stopped. 
;			exit
;			}
		}

;▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼ WorldCat Functions
;🌐
;▼▲▼ Determines if browser windows.
	confirmWorldCatOrgRecord(){
		if winExist("| WorldCat.org")
			winActivate
		else{
			msgBox, You are not on a WorldCat.org webpage. This macro has stopped.
			exit
		}
	}

;▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼ FirstSearch Functions
;🔍
;▼▲▼ Check that a browser window has an active tab with FirstSearch's Advanced Search Page Open.
	confirmFirstSearchLanding(){
		if winExist("FirstSearch: WorldCat Advanced Search")
			winActivate
		else{
			msgBox, Your browser window does not have the FirstSearch database loaded. The macro has stopped.`r`n`r`nIf your browser is open, the tab with FirstSearch may not be currently active. Click that tab and run the macro again.
			exit
		}
	}

;🔍
;▼▲▼ Determines if a browser with FirstSearch is already loaded.
	confirmFirstSearchRecord(){
		if winExist("FirstSearch: WorldCat Detailed"){
			winActivate
			return
			}
		msgBox, Your browser window does not have a FirstSearch record open. The macro has stopped.`r`n`r`nIf your browser is open, the tab with FirstSearch may not be currently active. Click that tab and run the macro again.
		exit
		}


;▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼ Diacritics and Publisher Name Fixes

;▼▲▼ Replaces composed diacritic letters with decomposed ones.
	replaceDiacritics()
		{
			global sourceCode
			sourceCode:= regExReplace(sourceCode, "ā", "ā")
			sourceCode:= regExReplace(sourceCode, "ī", "ī")
			sourceCode:= regExReplace(sourceCode, "ū", "ū")
			sourceCode:= regExReplace(sourceCode, "ē", "ē")
			sourceCode:= regExReplace(sourceCode, "ō", "ō")
			sourceCode:= regExReplace(sourceCode, "Ā", "Ā")
			sourceCode:= regExReplace(sourceCode, "Ī", "Ī")
			sourceCode:= regExReplace(sourceCode, "Ū", "Ū")
			sourceCode:= regExReplace(sourceCode, "Ē", "Ē")
			sourceCode:= regExReplace(sourceCode, "Ō", "Ō")
		;▼ Replace common "errors"
			sourceCode:= regExReplace(sourceCode, "，", ",")						; Replaces "block" comma with regular comma.
			sourceCode:= regExReplace(sourceCode, "  ", " ")
			sourceCode:= regExReplace(sourceCode, "：", ":")
		}

;▼ ■■■■■■■■■■■■■■■■■■■■■■■■■■■ Library of Japanese publishers with often misspelled names in WorldCat.
;▼ Not a function but works intimately withthe fixRomanizedPublisherNames function.
>^p::
	;▼ Copy a cell or text string to run the fixRomanizedPublisherNames on.
		clipboard:= ""
		send ^c
		clipWait, 1
		if errorLevel
			{
			msgBox, No data was copied to the clipboard to fix the publisher's name. the macro has stopped.
			return
			}
		fixPub:= clipboard
		fixRomanizedPublisherNames()
		clipboard:= fixPub
		send ^v
return

;▼ Prep and clean up text string to have no composed characters or spaces.
fixRomanizedPublisherNames()
	{
	global fixPub
	;▼ Replaces composed diacritic letters with decomposed ones.
		fixPub:= regExReplace(fixPub, "ā", "ā")
		fixPub:= regExReplace(fixPub, "ī", "ī")
		fixPub:= regExReplace(fixPub, "ū", "ū")
		fixPub:= regExReplace(fixPub, "ē", "ē")
		fixPub:= regExReplace(fixPub, "ō", "ō")
	
	;▼ Publisher names with spaces are likely correctly formatted, and can skip the name correction process.
		if fixPub contains A_space
			goTo skipfixPub
	
	;▼ Because of case matching, capital letters are replaces with lower case
		stringLower, fixPub, fixPub
		
	;a	
		fixPub:= regExReplace(fixPub, "akashishoten", "Akashi Shoten")
		fixPub:= regExReplace(fixPub, "asahishinbunshuppan", "Asahi Shinbun Shuppan")
	;b	
		fixPub:= regExReplace(fixPub, "benseishuppan", "Bensei Shuppan")
		fixPub:= regExReplace(fixPub, "bungeishunju", "Bungei Shunjū")
		fixPub:= regExReplace(fixPub, "bunkarongakkai", "Bunkaron Gakkai")
		fixPub:= regExReplace(fixPub, "bunkashobohakubunsha", "Bunka Shobō Hakubunsha")
	;c	
		fixPub:= regExReplace(fixPub, "chikumashobo", "Chikuma Shobō")
		fixPub:= regExReplace(fixPub, "chikurashobo", "Chikura Shobō")
		fixPub:= regExReplace(fixPub, "chosakukenjohosenta", "Chosakuken Jōhō Sentā")
		fixPub:= regExReplace(fixPub, "chunichieigasha", "Chūnichi Eigasha")
		fixPub:= regExReplace(fixPub, "chuodaigakushuppanbu", "Chūō Daigaku Shuppanbu")
		fixPub:= regExReplace(fixPub, "chuokeizaisha", "Chūō Keizaisha")
		fixPub:= regExReplace(fixPub, "chuokoronbijutsushuppan", "Chūō Kōron Bijutsu Shuppan")
		fixPub:= regExReplace(fixPub, "chuokoronsha", "Chūō Kōronsha")
		fixPub:= regExReplace(fixPub, "chuokoronshinsha", "Chūō Kōron Shinsha")
	;d	
		fixPub:= regExReplace(fixPub, "daiichihoki", "Daiichi Hōki")
		fixPub:= regExReplace(fixPub, "daiyamondosha", "Daiyamondosha")
	;e	
		fixPub:= regExReplace(fixPub, "enueichikeshuppan", "NHK Shuppan")
		fixPub:= regExReplace(fixPub, "esubikurieitibu", "Esubi Kurieitibu")
	;f	
		fixPub:= regExReplace(fixPub, "firumuatosha", "Firumu Āto Sha")
		fixPub:= regExReplace(fixPub, "fujiwarashoten", "Fujiwara Shoten")
		fixPub:= regExReplace(fixPub, "futamishobo", "Futami Shobo")
	;g	
		fixPub:= regExReplace(fixPub, "gakuyoshobo", "Gakuyō Shobō")
		fixPub:= regExReplace(fixPub, "gendaijinbunsha", "Gendai Jinbunsha")
		fixPub:= regExReplace(fixPub, "gendaishiryōshuppan", "Gendai Shiryō Shuppan")
		fixPub:= regExReplace(fixPub, "gendaishokan", "Gendai Shokan")
		fixPub:= regExReplace(fixPub, "genkishobo", "Genki Shobō")
		fixPub:= regExReplace(fixPub, "gentosha", "Gentōsha")
	;h	
		fixPub:= regExReplace(fixPub, "hayakawashobo", "Hayakawa Shobō")
		fixPub:= regExReplace(fixPub, "hitsujishobo", "Hitsuji Shobō")
		fixPub:= regExReplace(fixPub, "horeishuppan", "Hōrei Shuppan")
		fixPub:= regExReplace(fixPub, "horitsubunkasha", "Hōritsu Bunkasha")
		fixPub:= regExReplace(fixPub, "hojodoshuppan", "Hojodo Shuppan")
		fixPub:= regExReplace(fixPub, "honnoizumisha", "Hon no Izumisha")
		fixPub:= regExReplace(fixPub, "honnozasshisha", "Hon no Zasshisha")
		fixPub:= regExReplace(fixPub, "hoseidaigakushuppankyoku", "Hōsei Daigaku Shuppankyoku")
		fixPub:= regExReplace(fixPub, "hozokan", "Hōzōkan")

	;i	
		fixPub:= regExReplace(fixPub, "isutopuresu", "Isuto Puresu")
		fixPub:= regExReplace(fixPub, "iwanamishoten", "Iwanami Shoten")
		fixPub:= regExReplace(fixPub, "izumishoin", "Izumi Shoin")
	;j	
		fixPub:= regExReplace(fixPub, "jinbunshoin", "Jinbun Shoin")
		fixPub:= regExReplace(fixPub, "jiyukokuminsha", "Jiyū Kokuminsha")
		fixPub:= regExReplace(fixPub, "jiritsushobo", "Jiritsu Shobō")
		fixPub:= regExReplace(fixPub, "jurosha", "Jurōsha")
	;k	
		fixPub:= regExReplace(fixPub, "kadokawashoten", "Kadokawa Shoten")
		fixPub:= regExReplace(fixPub, "kaishahokankeihōmushorei", "Kaishaho Kankei Hōmu Shorei")
		fixPub:= regExReplace(fixPub, "kadokawagurupuhorudingusu", "Kadokawa Gurūpu Hōrudingusu")
		fixPub:= regExReplace(fixPub, "kadokawagurupupaburisshingu", "Kadokawa Gurūpu Paburishhingu")
		fixPub:= regExReplace(fixPub, "kanaeshobo", "Kanae Shobō")
		fixPub:= regExReplace(fixPub, "kanagawashinbunsha", "Kanagawa Shinbunsha")
		fixPub:= regExReplace(fixPub, "kasamashoin", "Kasama Shoin")
		fixPub:= regExReplace(fixPub, "kawadeshoboshinsha", "Kawade Shobō Shinsha")
		fixPub:= regExReplace(fixPub, "kawadeshobo", "Kawade Shobō")
		fixPub:= regExReplace(fixPub, "kazamashobo", "Kazama Shobō")
		fixPub:= regExReplace(fixPub, "keiogijukudaigakushuppankai", "Keiō Gijuku Daigaku Shuppankai")
		fixPub:= regExReplace(fixPub, "keizaisangyochosakai", "Keizai Sangyō Chōsakai")
		fixPub:= regExReplace(fixPub, "kenbun", "Kenbun")
		fixPub:= regExReplace(fixPub, "keisoshobo", "Keiso Shobō")
		fixPub:= regExReplace(fixPub, "kin*yūzaiseijijōkenkyūkai", "Kin'yū Zaisei Jijō Kenkyūkai")
		fixPub:= regExReplace(fixPub, "kitaojishobo", "Kitaōji Shobō")
		fixPub:= regExReplace(fixPub, "kobunsha", "Kōbunsha")
		fixPub:= regExReplace(fixPub, "koraidaigakkogurobarunihonkenkyuin", "Kōrai Daigakkō Gurōbaru Nihon Kenkyūin")
		fixPub:= regExReplace(fixPub, "koraidaigakkonihonkenkyusenta", "Kōrai Daigakkō Nihon Kenkyū Sentā")
		fixPub:= regExReplace(fixPub, "koseitorihikikyokai", "Kōsei Torihiki Kyōkai")
		fixPub:= regExReplace(fixPub, "koamagajin", "Koa Magajin")
		fixPub:= regExReplace(fixPub, "kodansha", "Kōdansha")
		fixPub:= regExReplace(fixPub, "kokusaishoin", "Kokusai Shōin")
		fixPub:= regExReplace(fixPub, "koyoshobo", "Kōyō Shobō")
		fixPub:= regExReplace(fixPub, "kyoeishobo", "Kyōei Shobo")
		fixPub:= regExReplace(fixPub, "kyotodaigakugakujutsushuppankai", "Kyōto Daigakugakujutsu Shuppankai")
		fixPub:= regExReplace(fixPub, "kyukoshoin", "Kyūko Shoin")
	;m	
		fixPub:= regExReplace(fixPub, "mainichishinbunsha", "Mainichi Shinbunsha")
		fixPub:= regExReplace(fixPub, "mainichishinbunshuppan", "Mainichi Shinbun Shuppan")
		fixPub:= regExReplace(fixPub, "matsurikasha", "Matsuri Kasha")
		fixPub:= regExReplace(fixPub, "meijitoshoshuppan", "Meiji Tosho Shuppan")
		fixPub:= regExReplace(fixPub, "meitokushuppansha", "Meitoku Shuppansha")
		fixPub:= regExReplace(fixPub, "minerubashobo", "Mineurba Shobō")
		fixPub:= regExReplace(fixPub, "minjihokenkyukai", "Minjihō Kenkyūkai")
		fixPub:= regExReplace(fixPub, "misuzushobo", "Misuzu Shobō")
		fixPub:= regExReplace(fixPub, "mitoshiyakusho", "Mitoshi Yakusho")
		fixPub:= regExReplace(fixPub, "mitsumurasuikoshoin", "Mitsumura Suiko Shoin")
		fixPub:= regExReplace(fixPub, "mitsumuratoshoshuppan", "Mitsumura Tosho Shuppan")
		fixPub:= regExReplace(fixPub, "miyaishoten", "Miyai Shoten")
		fixPub:= regExReplace(fixPub, "miyaobishuppansha", "Miyaobi Shuppansha")
		fixPub:= regExReplace(fixPub, "musashinobijutsudaigakushuppankyoku", "Musashino Bijutsu Daigaku Shuppankyoku")
	;n	
		fixPub:= regExReplace(fixPub, "nichigaiasoshietsu", "Nichigai Asoshiētsu")
		fixPub:= regExReplace(fixPub, "nihonkajoshuppan", "Nihon Kajo Shuppan")
		fixPub:= regExReplace(fixPub, "nihonhyoronsha", "Nihon Hyōronsha")
		fixPub:= regExReplace(fixPub, "nihonkindaibungakukan", "Nihon Kindai Bungagkukan")
		fixPub:= regExReplace(fixPub, "nihonkyohosha", "Nihon Kyōhōsha")
		fixPub:= regExReplace(fixPub, "nihonkyohosha", "Nihon Kyōhōsha")
		fixPub:= regExReplace(fixPub, "nikkeibipimaketingu", "Nikkei BP Māketingu")
		fixPub:= regExReplace(fixPub, "nikkeinashonarujiogurafikkusha", "Nikkei Nashonaru Jiogurafikkusha")
		fixPub:= regExReplace(fixPub, "nishidashoten", "Nishida Shoten")
	;o	
		fixPub:= regExReplace(fixPub, "okurazaimukyōkai", "Ōkura Zaimu Kyōkai")
		fixPub:= regExReplace(fixPub, "otsukishoten", "Ōtsuki Shoten")
		fixPub:= regExReplace(fixPub, "ochanomizushobo", "Ochanomizu Shobō")
		fixPub:= regExReplace(fixPub, "otowashobotsurumishoten", "Otowa Shobō Tsurumi Shoten")
	;p	
		fixPub:= regExReplace(fixPub, "paintanashonaru", "Pai Intānashonaru")
	;r	
		fixPub:= regExReplace(fixPub, "rekishishunjushuppan", "Rekishi Shunjun Shuppan")
		fixPub:= regExReplace(fixPub, "rikkashuppan", "Rikka Shuppan")
		fixPub:= regExReplace(fixPub, "rodochosakai", "Rōdō Chōsakai")
		fixPub:= regExReplace(fixPub, "romugyosei", "Rōmu Gyōsei")
	;s	
		fixPub:= regExReplace(fixPub, "san'ichishobo", "San'Ichi Shobō")
		fixPub:= regExReplace(fixPub, "san*ninsha", "San'ninsha")
		fixPub:= regExReplace(fixPub, "sanninsha", "San'ninsha")
		fixPub:= regExReplace(fixPub, "seibundoshuppan", "Seibundō Shuppan")
		fixPub:= regExReplace(fixPub, "seirinshoin", "Seirin Shoin")
		fixPub:= regExReplace(fixPub, "sekaishisosha", "Sekai Shisōsha")
		fixPub:= regExReplace(fixPub, "serikashobo", "Serika Shobō")
		fixPub:= regExReplace(fixPub, "shinchosha", "Shinchōsha")
		fixPub:= regExReplace(fixPub, "shinnihonhokishuppan", "Shin Nihon Hōki Shuppan")
		fixPub:= regExReplace(fixPub, "shinnihonshuppansha", "Shinnihon Shuppansha")
		fixPub:= regExReplace(fixPub, "shogakukan", "Shōgakukan")
		fixPub:= regExReplace(fixPub, "shojihomu", "Shoji Hōmu")
		fixPub:= regExReplace(fixPub, "shueishaintanashonaru", "Shueisha Intānashonaru")
		fixPub:= regExReplace(fixPub, "shueisha", "Shūeisha")
		fixPub:= regExReplace(fixPub, "shuwashisutemu", "Shūwa Shisutemu")
		fixPub:= regExReplace(fixPub, "sogensha", "Sōgensha")
	;t	
		fixPub:= regExReplace(fixPub, "tachibanashobo", "Tachibana Shobō")
		fixPub:= regExReplace(fixPub, "taiseishuppansha", "Taisei Shuppansha")
		fixPub:= regExReplace(fixPub, "taishukanshoten", "Taishūkan Shoten")
		fixPub:= regExReplace(fixPub, "tohoshoten", "Tōhō Shoten")
		fixPub:= regExReplace(fixPub, "tokyodaigakushuppankai", "Tōkyō Daigaku Shuppankai")
		fixPub:= regExReplace(fixPub, "tokyodoshuppan", "Tōkyōdō Shuppan")
		fixPub:= regExReplace(fixPub, "tokyohoreishuppan", "Tōkyō Hōrei Shuppan")
	;u	
		fixPub:= regExReplace(fixPub, "ueibushuppan", "Ueibu Shuppan")
	;y	
		fixPub:= regExReplace(fixPub, "yachiyo\shuppan", "Yahiyo Shuppan")
		fixPub:= regExReplace(fixPub, "yamakawashuppansha", "Yamakawa Shuppansha")
		fixPub:= regExReplace(fixPub, "yoshikawakobunkan", "Yoshikawa Kōbunkan")
		fixPub:= regExReplace(fixPub, "yumanishobo", "Yumani Shobō")
		fixPub:= regExReplace(fixPub, "yushindokobunsha", "Yūshindō Kōbunsha")
		
	;▼ adds space for "shōbo" and "shuppan".	
		fixPub:= regExReplace(fixPub, "shobo", " Shobō")
		fixPub:= regExReplace(fixPub, "shuppan", " Shuppan")
		fixPub:= regExReplace(fixPub, "  ", " ")

	;▼ sets fixPub to title case.
		stringLower, fixPub, fixPub, t
	
	skipfixPub:
	}

\::
	exitApp
return
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

setTitleMatchMode, 2
setKeyDelay, 100

;---
;Notes for using this Script
;---
;This scrip was designed to be used with Google sheets in the Google Chrome browser. It takes advantage of features unique to Chrome, you cannot use another browser.
;You must create a Google sheet that begins with the title "Collection Management - " after the final space you can name the spreadsheet whatever you like, for example:
;	Collection Management - German Books
;	Collection Management - Smith Donation
;	etc.
;The script looks for the initial characters of "Collection Management - " to know what window and tab to navigate to

;suspend hotkeys, press Ctrl+left arrow key to toggle hotkeys
>^left::
suspend, toggle
return

;---
; AutoHotKey does not require the declaration of variables, but to keep track of variables and what their purpose is I keep a list of them here to remember all the variables I use
;---
;all variables, treated as global and set with no value
	;variables for navigation
	;these variables are used as quasi if/then conditionals to change the flow of the code executed
		global vvJustOCLC:= "" ;when bibliographic data is pulled, only pastes the OCLC# to the Google sheet. Used specifcially because vendor supplied lists of books do not include OCLC#s
		global vvStopHere:="" ;various uses, makes the macro script stop running because only part of the script is needed
		global vvTitleSearch:= "" ;appends the ISBN number the WorldCat bib data that will be pasted to the Google Sheet

	;variables for storing bib data
		global vvOCLC:= "" ;holds OCLC number
		global vvAuthorR:= "" ;holds romanized author
		global vvAuthorV:= "" ;holds vernacular author
		global vvISBN:= "" ;holds ISBN number, specifically the ISBN-13
		global vvPubDate:= "" ;holds date of publication
		global vvPublisherR:= ""
		global vvPublisherV:= ""
		global vvTitleR:= "" ;holds romanized title
		global vvTitleV:= "" ;holds vernacular title
		global vvSeriesTitle:= "" ;holds series title (romanized only)
		global vvSourceCode:= "" ;holds source code from a web page to be parsed once or multiple times

	;variables for general use
		global vvCellData:= "" ;holds data from a cell copied in the Google sheet
		global vvURL:= ""
		
	;variables primarily for dupe checking in chinook
		global vvDoTitle:= "" ;for navigating chinook dupe search
		global vvCopyOrdered="" ;appends date an item was ordered if in chinook record
		
	;variables primarily for checking price on amazon.co.jp
		global vvPriceYen:= ""
		global vvCurrencyConvertor:=""
		global vvPriceUSD:= ""

;---
; Quality of life scripts
;---

insert:: ; when editing, and using sendInput, an additional line of code with "sleep 100" keeps the script from running too fast
	sendInput {raw}sleep 100
return

+esc:: ;for debugging the script, forces the script to stop and say what's in the clipboard
	sendInput {raw}msgBox, `%clipboard`%
	send {enter}
	sendInput {raw}exit
	send {enter}
return

;---
; Nagivation scripts for quickly getting to something that is needed
;--
;numpad1:: ; worldcat.org title serach in Google Chrome omnibar, then type title manually
	chromeActivate()
	send ^n
	sleep 2000
	send !d
	send {delete}
	sendInput {raw}worldcat.org
	send {tab}
	sendInput {raw}no: ;mt:bks NOT mt:elc bn:[ISBN here]
	;send ^+{left 2}
	;send +{left}
	send ^v
	send {enter}
return

numpad4:: ; worldcat.org title serach in Google Chrome omnibar, then type title manually
	chromeActivate()
	send ^n
	sleep 2000
	send !d
	send {delete}
	sendInput {raw}worldcat.org
	send {tab}
	sendInput {raw}mt:bks NOT mt:elc ti:[title here]
	send ^+{left 2}
	send +{left}
return

#c:: ;open chinook classic
	chromeActivate()
	chromeConfirm()
	run http://libraries.colorado.edu/
return

^+u:: ; get chinook record URL
	if winExist("Chinook Library Catalog - ")
		{
		winActivate
		send ^f
		send {delete}
		sendInput {raw}http://libraries.colorado.edu/record=
		send {esc}
		clipboard=
		send ^c
		clipWait 1
		if clipboard= http://libraries.colorado.edu/record=
			{
			send ^+{right 2}
			send clipboard=
			send ^c
			clipWait 1
			send ^{home}
			}
		}
return

^+L:: ; copy ISBNs from Google sheet and pastes them to WorldCat.org to search, loop
	chromeActivate()
	chromeConfirm()
	run https://docs.google.com/spreadsheets/d/1w15TVODE2I_gUkQHD5-Ghcskjc1bs9O21Moi6c9TyOA/edit#gid=999759851
return

;ctrl+tab mapped to numpad 7
numpad7::
	send ^{tab}
return

;---
; the following script is the most expansive and comprehensive, its main purpose is to pull data out of WorldCat.org. There are several different hot keys to excecute different kinds of searches
;---
; the following lines are for setting up the "navigation" variables
^+PrintScreen:: ; looping search for WorldCat Data with ISBN to only get OCLC#
^+numPadAdd:: ; same as above, but with the numpad
	vvJustOCLC:= "do"
	goTo JustOCLCsearch
PrintScreen:: ; single Search for WorldCat data with ISBN or Title Search to get bib data
numPadAdd:: ; same as above, but with the numpad
	vvJustOCLC:= ""
	vvStopHere:= "stop"
	vvTitleSearch:= ""
	goTo stopAtWorldCatResults
^PrintScreen:: ; looping search for WorldCat data with ISBN, will switch to single title search if no isbn
^numPadAdd:: ; same as above, but with the numpad
	vvJustOCLC:= ""
	JustOCLCsearch:
	vvStopHere:= ""
	stopAtWorldCatResults:
;---
; from here the macro will begin manipulating the computer
;---
	googleSheetConfirm()
	bibSearch: ; for looping scripts, this is the anchor they come back to to start the next search
	cellResetRowToISBN() ;navigates to cell in Column J to check for ISBN
	cellCopyCell()
	if vvJustOCLC = do ; the OCLC only search does not use anything other than ISBNs because vendors always supply an ISBN
		{
		if errorLevel= 1
			{
			msgBox, There is no data to search. The macro for only getting the OCLC# has stopped.
			exit
			}
		}
	if errorLevel= 1 ;tabs to column N if there is no ISBN
		{
		vvTitleSearch:= "do"
		send {right}
		send {right}
		send {right}
		send {right}
		cellCopyCell()
		if errorLevel= 1 ; tabs to column O if there is also no romanized title data
			{
			send {right}
			cellCopyCell()
			if errorLevel= 1
				{
				msgBox, There is no data to search. The macro has stopped.
				exit
				}
			}
		send ^n
		sleep 2000
		send !d
		sendInput {raw}worldcat.org
		sleep 100
		send {tab}
		sendInput {raw}mt:bks NOT mt:elc
		sleep 100
		send {space}
		sendInput {raw}ti:
		sleep 100
		send {delete}
		send ^v
		sleep 1000
		send {enter}
		sleep 3000
		exit ; title data is unreliable and needs to be checked manually, the macro stops after a title search
		}
	vvTitleSearch= ""
	send ^n ; from here, much of this code is for pulling data for WorldCat based on ISBN searches
	sleep 2000
	send !d
	sendInput {raw}worldcat.org
	sleep 100
	send {tab}
	sendInput {raw}mt:bks NOT mt:elc
	sleep 100
	send {space}
	sendInput {raw}no:
	sleep 100
	send {delete}
	send ^v
	sleep 1000
	send {enter}
	sleep 3000
	if vvStopHere= stop
			{
			vvStopHere=
			exit
			}
;If hot key for just looking at WorldCat search results was used, macro will stop here
	worldCatResultsNone()
		if errorLevel= 0
			{
			send ^w
			sleep 500
			send ^w
			sleep 500
			sendInput {raw}no results
			sleep 100
			send {down}
			goTo, bibSearch
			}
	worldCatResultsTen()
		if errorLevel= 0
			{
			send ^w
			sleep 500
			send ^w
			sleep 500
			send {right}
			send ^+v
			send {down}
			goTo, bibSearch
			}
	
	worldCatResultsOne()
		if errorLevel= 0
			{
			send {tab 9}
			send {enter}
			sleep 3000
			goTo getWorldCatSourceCode ; skips the last scenario of multiple WorldCat search results
			}
	worldCatResultsMultiple()
		send ^w
		sleep 500
		send {right}
		send ^+v
		send {down}
		goTo, bibsearch
	getWorldCatSourceCode:
		send ^u	
		sleep 3000
		sourceCodeConfirm()
		send ^f
		send {delete}
		sendInput {raw}OCLC:
		sleep 100
		send {esc}
		clipboard=
		send ^c
		clipWait, 1
		if errorLevel, 1
			{
			sleep 3000
			sourceCodeConfirm()
			send ^f
			send {delete}
			sendInput {raw}OCLC:
			sleep 100
			send {esc}
			clipboard=
			send ^c
			clipWait, 1
			if errorLevel, 1
				{
				msgBox, Error: the macro could not confirm the source code has loaded.
				exit
				}
			}
		send ^a
		sleep 500
		clipboard=
		send ^c
		clipWait, 3
		clipboard:= regExReplace(clipboard, "`r`n", "")
		vvSourceCode:= clipboard
		send !{F4}
		sleep 500
	googleSheetConfirm()
;parses WorldCat source code into key bibliographic data, puts each data in its own variable
	parseISBNfromWC()
	parseOCLCfromWC()
	parseAuthorRfromWC()
	parseAuthorVfromWC()
	parseTitleRfromWC()
	parseTitleVfromWC()
	parseSeriesTitlefromWC()
	parsePublisherRfromWC()
		;fixDiacritics()
		;fixPubR()
		;vvPublisherR:= clipboard
	parsePublisherVfromWC()
	parsePublisherDatefromWC()
	sleep 1000
	cellResetRowToISBN() ; when the script returns to the Google sheet to paste data, it needs to be in this cell to paste correctly
	if vvJustOCLC= do
		{
		clipboard:= vvOCLC
		send {right}
		send ^+v
		sleep 500
		send {down}
		goTo bibSearch
		}
	clipboard:= vvISBN . "`t" . vvOCLC . "`t" . vvAuthorR . "`t" . vvAuthorV . "`t" . vvTitleR . "`t" . vvTitleV . "`t" . vvSeriesTitle . "`t" . vvPublisherR . "`t" . vvPublisherV . "`t" . vvPubDate
	fixDiacritics() ; makes all composed characters decomposed
	send ^+v
	sleep 500
	send {down}
	sleep 100
	if vvStopHere= stop
		{
		vvStopHere=
		exit
		}
	goto bibSearch
return

;Loads the source code of a WorldCat record
F6:: ; if keyboard doesn't have numpad, use this hotkey
numpad6::
	vvStopHere:= "stop"
	If WinExist("WorldCat.org")
		WinActivate
			else
			{
			msgBox, Error: there is no window with WorldCat results open OR the tab in Chrome with the WorldCat results is not currently active.
			exit
			}
	send ^f
	sendInput {raw}add to list
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
		if errorLevel= 1
			{
			msgBox, The macro you used is only for a WorldCat record page to load the source code and extract the bliographic data.
			exit
			}
	send ^f
	sleep 100
	sendInput {raw}search results for
	send {esc}
	clipboard=
	send ^c
	if errorLevel= 1
		{
		msgBox, The macro you used is only for a WorldCat record page to load the source code and extract the bibliographic data.
		exit
		}
	sleep 100
	goTo getWorldCatSourceCode
return

; Uses Chinook Classic (http://libraries.colorado.edu/) to determine if the library already owns a book via ISBN or TITLE
^+E::
	titleDupeCheck:
	vvDoTitle:= "do"
	goTo skipISBNdupeCheck
^+I::
		vvDoTitle:= ""
	ISBNdupeCheck:
		vvCellData:= ""
		vvURL:= ""
		vvCopyOrdered:= ""
	skipISBNdupeCheck:
	googleSheetConfirm()
	cellResetRowToISBN()
		if vvDoTitle= do
			{
			send {right}
			send {right}
			send {right}
			send {right}
			}
	copyCellForDupe:
	cellCopyCell()
	vvCellData:= clipboard
	clipWait 1
	if clipboard= 
		{
		msgBox, There is no data to search. The macro has stopped.
		exit
		}
;navigate chinook search menu for ISBN or TITLE search
	chinookGoto()
	chinookConfirm()
		send {tab 3}
		sleep 100
		if vvDoTitle= do
			{
			send {tab 1}
			goTo skipSetISBNsearch
			}
		send i
		send {tab}
	skipSetISBNsearch:
	clipboard=
	clipboard:= vvCellData
	clipWait 1
	send ^v
	sleep 100
	send {tab 2}
	send {enter}
	sleep 5000
;confirm tha the chinook results page has loaded
	send ^f
	sendInput {raw}Result Page
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel= 1
		sleep 5000
;confrim chinook results have "no entires found"
	send ^f
	sendInput {raw}no matches found
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= no matches found
		{
		clipboard= no
		goTo backToGoogleSheet
		}
;confirm chinook results have multiple results, for title only. Have not yet encountered an ISBN search with multiple results
	if vvDoTitle= do
		{
		send ^f
		sleep 100
		sendInput {raw}titles (1-
		sleep 100
		send {esc}
		clipboard=
		send ^c
		clipWait, 1
		if clipboard= titles (1-
			{
			send !d
			sleep 100
			clipboard=
			send ^c
			clipWait, 1
			vvURL:= clipboard
			clipboard:= "=hyperlink(""" . vvURL . """,""" . "2+ results"")"
			goTo backToGoogleSheet
			}
		}
;confirm chinook results have one results, record will have permalink	
	send ^f
	sendInput {raw}http://libraries.colorado.edu/record=
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= http://libraries.colorado.edu/record=
		{
		send ^+{right 2}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		vvURL:= clipboard
		}
;after confirming one result, check if item is ordered or not
	send ^f
	sendInput {raw}copy ordered for
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= copy ordered for
		{
		send ^+{down}
		clipboard=
		send ^c
		clipWait, 1
		clipboard:= regExReplace(clipboard, "`r`n", "")
		clipboard:= regExReplace(clipboard, ".Description", "")
		vvCopyOrdered:= clipboard
		clipboard:= "=hyperlink(""" . vvURL . """,""" . vvCopyOrdered . """" . ")"
		goto backToGoogleSheet
		}
	if vvURL contains http://libraries.colorado.edu/record=
		clipboard:= vvURL
		goTo backToGoogleSheet
;in case of an unanticipated scenario, the clipboard will not containt any text strings that trigger a return to the Google sheet
	msgBox, An unexpected outcome has occurred while checking Chinook for an ISBN duplicate. The macro has stopped and may need new code to account for this new situation.
	exit
;data collected in Chinook for dupe check goes back to Google Sheet
		backToGoogleSheet:
		send !{F4}
		sleep 2000
		send {home}
		send {right}
		send {right}
		send {right}
		send {right}
		send {right}
		if vvDoTitle= do ; tabs one more time to move to title duple column
			{
			send {right}
			}
		send {delete}
		send ^+v
		sleep 500
		send {down}
		if vvDoTitle= do
			{
			goto titleDupeCheck
			}
		goto ISBNdupeCheck
return

;Looks up price and stops macro at the search results
^4::
	vvStopHere:= "stop"
	goTo justAmazonJPsearchResults
; Price check with Amazon, uses data from amazon.co.jp search results with only 1 result
^+4::
	vvStopHere:= ""
	justAmazonJPsearchResults:
	googleSheetConfirm()
	getAmazonJPPrice:
	cellResetRowToISBN()
;copies the ISBN, author, and title to more accurately search amazon.co.jp
	cellCopyCell()
		clipboardEmptyStop()
		vvISBN:= clipboard
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
	cellCopyCell()
		clipboardEmptyStop()
		vvAuthorV:= clipboard
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send ^c
		sleep 100
	cellCopyCell()
		clipboardEmptyStop()
		vvTitleV:=clipboard
		sleep 100
		send {esc}
		sleep 100
		clipboard=
		clipboard:= vvISBN . " "  . vvAuthorV . " " . vvTitleV
		clipboard:= regExReplace(clipboard, "n/a ", "")
	googleSheetToAmazonJP()
;check the variable "vvStopHere" when the macro for searching 1 item is run
	if vvStopHere= stop
		{
		vvStopHere=
		exit
		}
;if there are no search results, the search will end and a message to do a manual search will be in pasted in the "Price ¥ or ₩" column
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}スペルの確認を試み
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel= 0
		{
		send !{F4}
		sleep 100
		googleSheetConfirm()
		send {end}
		sleep 100
		send {left}
		sleep 100
		send {left}
		sleep 100
		send {delete}
		sleep 100
		send {raw}check manually
		sleep 100
		send {down}
		sleep 100
		goTo justAmazonJPsearchResults
		}
;URL for search results saved in "vvURL" variable
	send !d
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	vvURL:= clipboard	
;checks if there is only 1 result and takes different actions accordingly
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}1件の結果
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
;check if prime shipping is available which means direct price quote can be obtained from amazon.co.jp
	if errorLevel= 0
		{
		send ^f
		sleep 100
		send {delete}
		sleep 100
		send {raw}通常配送料無料
		sleep 100
		send {esc}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		if errorLevel= 0
		goTo forSaleDirectlyThroughAmazon.jp
		}
	if errorLevel= 1 ;opens individual record of item not being sold directly through amazon.co.jp
		{
		send ^u
		sleep 5000
		sourceCodeConfirm()
			send ^f
			sleep 100
			send {delete}
			sleep 100
;finds URL data to open the individual item record in amazon.co.jp
			send {raw}a class="a-link-normal" target="_blank" href=
			sleep 300
			send {esc}
			sleep 300
			send ^+{down}
			sleep 100
			clipboard=
			send ^c
			clipWait, 1
;trims out text that gets in the way of loading the URL
			clipboard:= regExReplace(clipboard, "`r`n", "")
			clipboard:= regExReplace(clipboard, ".*href=.", "")
			clipboard:= regExReplace(clipboard, ".\>.*", "")
			send !d
			sleep 100
;concantenates the beginning of the amazon.co.jp URL with the indivudal record URL data
			send {raw}amazon.co.jp
			sleep 100
			send ^v
			sleep 100
			send {enter}
			sleep 3000
;save the URL for use in the Google sheet
			send !d
			sleep 100
			clipboard=
			send ^c
			clipWait, 1
			vvURL:= clipboard
			send ^u
			sleep 5000
		sourceCodeConfirm()
;pull the price data, this isn't the definitive value as individual sellers on amazon.co.jp have various prices
			send ^f
			sleep 100
			send {delete}
			sleep 100
			send {raw}a-size-medium a-color-price offer-price a-text-normal
			sleep 100
			send {esc}
			sleep 100
			clipboard=
			send ^c
			clipWait, 1
;grab and trim the price data from the individual item record on amazon.co.jp
			if clipboard contains not a-size-medium
				{
				msgBox, The source code tab did not open. The macro has stopped.
				exit
				}
			send ^+{right}
			sleep 100
			send ^c
			sleep 100
			clipboard:= regExReplace(clipboard, ".*￥ ", "")
			clipboard:= regExReplace(clipboard, ",", "")
			vvPriceYen:= clipboard
			goTo amazonJPpriceToGoogleSheet
		}
forSaleDirectlyThroughAmazon.jp:
	send ^u
	sleep 5000
	sourceCodeConfirm()
		sleep 100
		send ^f
		sleep 100
		send {delete}
		sleep 100
		send {raw}a-price-whole
		sleep 300
		send {esc}
		sleep 500
		clipboard=
		send ^c
		clipWait, 3
		send ^c
		clipWait 1
;if the text string "a-price-whole" is NOT found in the source code. The macro stops. NOTE: there is more than one text string "a-price-whole" in the amazon.co.jp source code, but the FIRST instance should be for the book that came up in the search results. The other "a-price-whole" text strings are for preciously looked at or recommended books
	if clipboard!= a-price-whole
		{
		send ^f
		sleep 100
		send {delete}
		sleep 100
		send {raw}a-color-price
		sleep 300
		send {esc}
		sleep 500
		clipboard=
		send ^c
		clipWait, 1
		if clipboard!= a-color-price
			{
			msgBox, Error: an unexprect situation has occured trying to extract price data form amazon.co.jp. The macro has stopped.
			exit
			}
		}
;if the text string :a-price-whole" IS found, the price data is extracted and sent to 
		send ^+{right}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		clipboard:= regExReplace(clipboard, ".*\>", "")
		clipboard:= regExReplace(clipboard, ",", "")
		clipboard:= regExReplace(clipboard, "￥", "")
		global vvPriceYen:= clipboard
	amazonJPpriceToGoogleSheet:
		send !{F4}
		sleep 500
	googleSheetConfirm()
;add currency conversion formula to column x and then copy it and paste it as an absolute value
	;insert currency conversion formula into column X
		send {end}
		sleep 100
		send {left}
		sleep 100
		send {delete}
		sleep 100
		send {raw}=GOOGLEFINANCE("CURRENCY:JPYUSD")
		;clipboard:= "=GOOGLEFINANCE(""CURRENCY:JPYUSD"")"
		;clipWait, 1
		;send ^+v
		sleep 500
		send {enter}
		sleep 500
	;copy formula value and paste as absolute value, then store in vvCurrencyConvertor
		send {up}
		sleep 500
		clipboard=
		send ^c
		clipWait, 3
		send ^+v
		sleep 500
		send ^c
		sleep 100
		 vvCurrencyConvertor:= clipboard
;calculate price in US dollars
		clipboard:= vvPriceYen * vvCurrencyConvertor
		global vvPriceUSD:= clipboard
;paste vvPriceYen with hyperlink to amazon.co.jp results
		send {left}
		sleep 100
		send {delete}
		sleep 100
		clipboard:= "=hyperlink(""" . vvURL . """,""" . vvPriceYen """)"
		send ^+v
		sleep 300
;vvPriceUSD variable is pasted in column T
		send {left}
		sleep 100
		send {left}
		sleep 100
		send {left}
		sleep 100
		clipboard=
		clipboard:= vvPriceUSD
		clipWait, 1
		send ^+v
		sleep 300
		send {down}
		sleep 500
	goto getAmazonJPprice
exit

;=====
;=====
;FUNCTION LIBRARY
;=====
;=====

;---
; general use functions for working with spreadsheets
;---
cellCopyCell() ; copies data from a cell in a specific way to ensure no "invisible" encoding is copied with cell contents
	{
	send {f2}
	send ^+{home}
	clipboard=
	send ^c
	clipWait, 1
	send {esc}
	}

cellResetRowToISBN() ; Google sheets has problems with the {right 9} format
	{
	sendMode input
	send {home}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	}

;---
; functions for working with Chinook
;---
chinookGoto() ; loads http://libraries.colorado.edu/
	{
	send ^n
	sleep 2000	
	send !d
	sendInput {raw}http://libraries.colorado.edu/
	sleep 100
	send {enter}
	sleep 4000
	}

chinookConfirm() ; confirms Chinook web page had loaded has loaded
	{
	send ^f
	send {delete}
	sleep 100
	sendInput {raw}My Chinook
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if clipboard!= My Chinook
		{
			msgBox, The website "libraries.colorado.edu/" has failed to load, the macro has stopped.
			exit
		}
	}

chinookTitleResultsMultiple() ;copier URL in address bar and post link as "2+ results" for Google Sheet
	{
	send ^f
	sendInput {raw}titles (1-
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	send !d
	send ^c
	sendInput {raw}=hyperlink("
	sleep 100
	send ^v
	send {raw}","2+ results")
	send ^a
	send ^x
	}

;---
; functions for working with Google's Chrome browser and Google sheets
;---
chromeActivate() ; checks if Chrome is already running, opens it if not, and loads the Google sheet
	{
	if winExist("ahk_exe chrome.exe")
		winActivate
			else
			Run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
	sleep 1000
	if winExist("ahk_exe chrome.exe")
		winActivate
			else
			{
			msgBox, Google Chrome is not open, not installed on your computer, or not stored in the file path C:\Program Files (x86)\Google\Chrome\Application\chrome.exe. The macro will stop.
			exit
			}
	sleep 1000
	}


chromeConfirm() ; confirms Chrome is open, stops macro if not
	{
	if winExist("ahk_exe chrome.exe")
		WinActivate
			else
			{
			msgBox, Google Chrome is not open, The macro will stop.
			exit
			}
	sleep 1000
	}

clipboardEmptyStop() ;stops macro if there is no data in the clipboard
	{
	if clipboard=
		{
		msgBox, Error: there is no data to search. The macro has stopped.
		exit
		}
	}

googleSheetConfirm() ;confirms a Google sheet starting with "Collection Development - " is open, stops macro if not
	{
	if winExist("Collection Development - ")
		winActivate
		else
		{
		msgBox, Chrome might be open but the tab for the GoogleSheet "Collection Development - " is not active. Activate that tab and try running the macro again.
		}
	}

;opens new Chrome window and searches Amazon.com directly from address bar
googleSheetToAmazonJP()
	{
	send ^n
	sleep 2000
	send !d
	sendInput {raw}amazon.co.jp
	sleep 100
	send {tab}
	send ^v
	clipboard=
	send {enter}
	sleep 3000
	}

sourceCodeConfirm()
	{
	if winExist("view-source:")
	winActivate
		else
		{
		msgBox, The source code tab did not open. The macro has stopped.
		exit
		}
	}

;---
; The next four functions tell the script what to do based on one of four WorldCat search outcomes
;---
worldCatResultsMultiple() ; after all other possible WorldCat results have been considered, this is the last possibility, multiple results less than 10
	{
	send !d
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel, 1
		return
	vvURL:= clipboard
	clipboard:= "=hyperlink(""" . vvURL . """,""" . "2+ results"")"
	}
worldCatResultsNone() ; confirms no results in WorldCat, searches for unique text string on a "no results" page
	{
	send ^f
	send {delete}
	sendInput {raw}no results match your search
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	}
worldCatResultsOne() ; confirms 1 result in WorldCat
	{
	send ^f
	send {delete}
	sendInput {raw}Results 1-1
	sleep 100
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	}
worldCatResultsTen() ; confirms 10+  results in WorldCat. Must be done before confirming 1 result becuase the beginning of the unique text string on a WorldCat serach result for 1 result is "1-1" which looks like the beginning of "1-10"
	{
	send ^f
	send {delete}
	sendInput {raw}1-10
	send {esc}
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel, 1
		return
	send !d
	clipboard=
	send ^c
	clipWait, 1
	vvURL:= clipboard
	clipboard:= "=hyperlink(""" . vvURL . """,""" . "10+ results"")"
	}

;=====
;=====
; Trim source code variables with regular expressions
;=====
;=====

parseISBNfromWC()
	{
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*ISBN/ISSN: .......... ", "")
	clipboard:= regExReplace(clipboard, " .......... .*", "")
	clipboard:= regExReplace(clipboard, "OCLC.*", "")
	if clipboard contains <!DOCTYPE
		clipboard:= "n/a"
	vvISBN:= clipboard
	}

parseOCLCfromWC()
	{
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*OCLC:", "")
	clipboard:= regExReplace(clipboard, "</textarea>.*", "")
	vvOCLC:= clipboard
	}

parseAuthorRfromWC()
	{
	if vvSourceCode contains not bib-author-cell
		{
		clipboard:= "n/a"
		goto noAuthorRdata
		}
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*bib-author-cell.>", "<td>")
	clipboard:= regExReplace(clipboard, "</td>.*", "</td>")
	;core find and replace commands
		clipboard:= regExReplace(clipboard, ".*</span>", "<td>")
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "<td>.*?<a", "<td><a")
		clipboard:= regExReplace(clipboard, "<td> +\; </td>", "")
		clipboard:= regExReplace(clipboard, "<a.*?author'>", "")
		clipboard:= regExReplace(clipboard, "</a>", "")
		clipboard:= regExReplace(clipboard, "<a.*allauthors.*</td>", "See WC record for all authors</td>")
		clipboard:= regExReplace(clipboard, "(,|\.)</td>", "</td>")
		clipboard:= regExReplace(clipboard, " \;", "; ")
	;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, "(,|\.)\;", ";")
		clipboard:= regExReplace(clipboard, " \; ", ";")
		clipboard:= regExReplace(clipboard, "[0-9]+-", "")
	;final cleanup
		clipboard:= regExReplace(clipboard, " +", " ")
		clipboard:= regExReplace(clipboard, "<td>", "")
		clipboard:= regExReplace(clipboard, "</td>", "")
		StringLower, clipboard, clipboard, T ;I don't fully understand how this code words, but it does apply Title Case
	vvAuthorR:= clipboard
	noAuthorRdata:
	}

parseAuthorVfromWC()
	{
	if vvSourceCode contains not bib-author-cell
		{
		clipboard:= "n/a"
		goto noAuthorVdata
		}
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*bib-author-cell.>", "<td>")
	clipboard:= regExReplace(clipboard, "</td>.*", "</td>")
	;core find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, ".*lang=.en.>.*", "")
		clipboard:= regExReplace(clipboard, ".*lang=....>", "")
		clipboard:= regExReplace(clipboard, "</span.*", "</td>")
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</td>", "</td>")
		clipboard:= regExReplace(clipboard, " </td>", "</td>")
		clipboard:= regExReplace(clipboard, "<td><a href='.*", "")
	;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, " = .*", "")
		clipboard:= regExReplace(clipboard, "\(....-\)", "")
		clipboard:= regExReplace(clipboard, ", ....- author", "")
	;final cleanup
		clipboard:= regExReplace(clipboard, "</td>", "")
		clipboard:= regExReplace(clipboard, " +", " ")
	vvAuthorV:= clipboard
	noAuthorVdata:
	}

parseTitleRfromWC()
	{
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*<h1 class=.title.>", "<td>")
	clipboard:= regExReplace(clipboard, "</h1>.*", "</td>")
	;prep find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "&lt\;", "<")
		clipboard:= regExReplace(clipboard, "&gt\;", ">")
	;find and replace commands for romanized title
		clipboard:= regExReplace(clipboard, ".*</div>", "<td>")
		clipboard:= regExReplace(clipboard, " : ", ": ")
		;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, "\. / 1", "")
	;final cleanup
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</td>", "</td>")
		clipboard:= regExReplace(clipboard, "<td>", "")
		clipboard:= regExReplace(clipboard, "</td>", "")
		clipboard:= regExReplace(clipboard, " +", " ")	
	vvTitleR:= clipboard
	}

parseTitleVfromWC()
	{
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*<h1 class=.title.>", "<td>")
	clipboard:= regExReplace(clipboard, "</h1>.*", "</td>")
	;prep find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "&lt\;", "<")
		clipboard:= regExReplace(clipboard, "&gt\;", ">")
	;find and replace commands for romanized title
		clipboard:= regExReplace(clipboard, ".*lang=.en.>.*", "")
		clipboard:= regExReplace(clipboard, ".*lang=....>", "<tv>")
		clipboard:= regExReplace(clipboard, "(\.| /)</div>.*", "</tv>")
		clipboard:= regExReplace(clipboard, "<td>.*", "")
		clipboard:= regExReplace(clipboard, ".*</td>", "")
		clipboard:= regExReplace(clipboard, " : ", ": ")
	;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, " / .* \; .*", "")
	;final cleanup
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</tv>", "</tv>")
		clipboard:= regExReplace(clipboard, "<tv>", "")
		clipboard:= regExReplace(clipboard, "</tv>", "")
		clipboard:= regExReplace(clipboard, " +", " ")
	vvTitleV:= clipboard
	}

parseSeriesTitlefromWC()
	{
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*22'>", "<td>")
	clipboard:= regExReplace(clipboard, "</td>.*", "</td>")
	;prep find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "&lt\;", "<")
		clipboard:= regExReplace(clipboard, "&gt\;", ">")
	;find and replace commands for series
		clipboard:= regExReplace(clipboard, "\.</a>, ", ", ")
		clipboard:= regExReplace(clipboard, "</a>, ", ", ")
		clipboard:= regExReplace(clipboard, "</a></td>", "</td>")
	;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, " \;, ", ", ")
	;final cleanup
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</td>", "</td>")
		clipboard:= regExReplace(clipboard, "<td>", "")
		clipboard:= regExReplace(clipboard, "</td>", "")
		clipboard:= regExReplace(clipboard, " +", " ")
	;gets rid of any other textarea
		clipboard:= regExReplace(clipboard, ".*</div>.*","")
	vvSeriesTitle:= clipboard
	if clipboard= 
		vvSeriesTitle:= "no series"
	}

parsePublisherRfromWC()
	{
	if vvSourceCode contains not bib-publisher-cell
		{
		vvPublisherR:= "n/a"
		goto noPublisherRdata
		}
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*bib-publisher-cell.>", "<td>")
	clipboard:= regExReplace(clipboard, "</td>.*", "</td>")
	clipboard:= clipboard
	;prep find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "&lt\;", "<")
		clipboard:= regExReplace(clipboard, "&gt\;", ">")
	;find and replace commands for series
		clipboard:= regExReplace(clipboard, ".*</span>*", "<td>")
		clipboard:= regExReplace(clipboard, ".* : ", "<td>")
		clipboard:= regExReplace(clipboard, ", ([0-9]|\[|©).*", "</td>")
		clipboard:= regExReplace(clipboard, "<td> +", "<td>")
	;find and replace cleanup, less common situations
		clipboard:= regExReplace(clipboard, "\[.*\]", "")
		clipboard:= regExReplace(clipboard, "(([1|2])...(\.</td>|</td>))", "</td>")
		clipboard:= regExReplace(clipboard, ".* \; ", "<td>")
		clipboard:= regExReplace(clipboard, "Tōkyō ", "")
		clipboard:= regExReplace(clipboard, "(2|1).*nen .*</td>", "</td>")
		clipboard:= regExReplace(clipboard, ", (Heisei|Shōwa|Taishō|Meiji).*</td>", "</td>")
	;final cleanup
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</td>", "</td>")
		clipboard:= regExReplace(clipboard, "<td>", "")
		clipboard:= regExReplace(clipboard, "</td>", "")
		clipboard:= regExReplace(clipboard, " +", " ")	
	vvPublisherR:= clipboard
	noPublisherRdata:
	}

parsePublisherVfromWC()
{
	if vvSourceCode contains not bib-publisher-cell
		{
		vvPublisherV:= "n/a"
		goto noPublisherVdata
		}
	clipboard:= vvSourceCode
	clipboard:= regExReplace(clipboard, ".*bib-publisher-cell.>", "<td>")
	clipboard:= regExReplace(clipboard, "</td>.*", "</td>")
	clipboard:= clipboard
		;prep find and replace commands
		clipboard:= regExReplace(clipboard, "\;&nbsp", " ")
		clipboard:= regExReplace(clipboard, "&lt\;", "<")
		clipboard:= regExReplace(clipboard, "&gt\;", ">")
		;find and replace commands for series
		clipboard:= regExReplace(clipboard, ".*lang=....>", "<tv>")
		clipboard:= regExReplace(clipboard, "<td>.*", "")
		clipboard:= regExReplace(clipboard, "</span>.*", "</td>")
		;find and replace cleanup, less common situations

		;final cleanup
		clipboard:= regExReplace(clipboard, "( |,|\.|=)</td>", "</td>")
		clipboard:= regExReplace(clipboard, "<tv>", "")
		clipboard:= regExReplace(clipboard, "</td>", "")
		clipboard:= regExReplace(clipboard, " +", " ")	
	vvPublisherV:= clipboard
	noPublisherVdata:
	}

parsePublisherDatefromWC()
	{
	clipboard:=vvSourceCode
	clipboard:= regExReplace(clipboard, ".*schema:datePublished.>", "")
	clipboard:= regExReplace(clipboard, "</span>.*", "")
	clipboard:= regExReplace(clipboard, "/9999", "")
	vvPubDate:= clipboard
	if clipboard= 
	vvPubDate:= "n/a"
	}

;=====
;=====
;Library of regular expressions to fix incorrect Japanese publisher names
;=====
;=====

;replace precompoased diacrtic letters with dexomponsed ones
;DEcomposed  vowels with macron diacritic: ā ī ū ē ō Ā Ī Ū Ē Ō
;PREcomposed vowels with macron diacricic: ā ī ū ē ō Ā Ī Ū Ē Ō

^+p::
	clipboard=
	send ^c
	clipWait, 1
	fixDiacritics()
	fixPubR()
	send ^v
return
	

fixDiacritics()
	{
	clipboard:= regExReplace(clipboard, "ā","ā")
	clipboard:= regExReplace(clipboard, "ī","ī")
	clipboard:= regExReplace(clipboard, "ū","ū")
	clipboard:= regExReplace(clipboard, "ē","ē")
	clipboard:= regExReplace(clipboard, "ō","ō")
	clipboard:= regExReplace(clipboard, "Ā","Ā")
	clipboard:= regExReplace(clipboard, "Ī","Ī")
	clipboard:= regExReplace(clipboard, "Ū","Ū")
	clipboard:= regExReplace(clipboard, "Ē","Ē")
	clipboard:= regExReplace(clipboard, "Ō","Ō")
	}

fixPubR()
	{
;a
	if clipboard= akashishoten
		clipboard:= "Akashi Shoten"
	if clipboard= aokishoten
		clipboard:= "Aoki Shoten"
	if clipboard= asahishinbunShuppan
		clipboard:= "Asahi Shinbun Shuppan"
;b\
	if clipboard= benseishuppan
		clipboard:= "Bensei Shuppan"
	if clipboard= bungeishunju
		clipboard:= "Bungei Shunjū"
	if clipboard= bunkarongakkai
		clipboard:= "Bunkaron Gakkai"
	if clipboard= bunkashobōhakubunsha
		clipboard:= "Bunka Shobō Hakubunsha"
;c
	if clipboard= chikumashobo
		clipboard:= "Chikuma Shobo"
	if clipboard= chikumashobō
		clipboard:= "Chikuma Shobō"
	if clipboard= chosakukenjōhōsentā
		clipboard:= "Chosakuken Jōhō Sentā"
	if clipboard= chūnichieigasha
		clipboard:= "Chūnichi Eigasha"
	if clipboard= chūōdaigakushuppanbu
		clipboard:= "Chūō Daigaku Shuppanbu"
	if clipboard= chūōkeizaisha
		clipboard:= "Chūō Keizaisha"
	if clipboard= chūōkōronbijutsushuppan
		clipboard:= "Chūō Kōron Bijutsu Shuppan"
	if clipboard= chūōkōronshinsha
		clipboard:= "Chūō Kōron Shinsha"
	if clipboard= chuokoronshinsha
		clipboard:= "Chūō Kōron Shinsha"
;d
	if clipboard= daiichihōki
		clipboard:= "Daiichi Hōki"
;e
	if clipboard= enueichikēShuppan
		clipboard:= "NHK Shuppan"
	if clipboard= esubikurieitibu
		clipboard:= "Esubi Kurieitibu"
;f
	if clipboard= firumuātosha
		clipboard:= "Firumu Āto Sha"
	if clipboard= fujiwarashoten
		clipboard:= "Fujiwara Shoten"
	if clipboard= futamishobō
		clipboard:= "Futami Shobo"
;g
	if clipboard= gakuyōshobō
		clipboard:= "Gakuyō Shobō"
	if clipboard= gendaijinbunsha
		clipboard:= "Gendai Jinbunsha"
	if clipboard= gendaishiryōshuppan
		clipboard:= "Gendai Shiryō Shuppan"
	if clipboard= gendaishokan
		clipboard:= "Gendai Shokan"
	if clipboard= genkishobō
		clipboard:= "Genki Shobo"
	if clipboard= gentosha
		clipboard:= "Gentōsha"
;h
	if clipboard= hayakawashob
		clipboard:= "Hayakawa Shobo"
	if clipboard= hitsujishob
		clipboard:= "Hitsuji Shobo"
	if clipboard= hōreishuppan
		clipboard:= "Hōrei Shuppan"
	if clipboard= hōritsubunkasha
		clipboard:= "Hōritsu Bunkasha"
	if clipboard= hojodoshuppan
		clipboard:= "Hojodo Shuppan"
	if clipboard= hojodo shuppan
		clipboard:= "Hōjōdō Shuppan"
	if clipboard= honnoizumisha
		clipboard:= "Hon no Izumisha"
	if clipboard= honnozasshisha
		clipboard:= "Hon no Zasshisha"
;i
	if clipboard= isutopuresu
		clipboard:= "Īsuto Puresu"
	if clipboard= iwanamishoten
		clipboard:= "Iwanami Shoten"
	if clipboard= izumishoin
		clipboard:= "Izumi Shoin"
;j
	if clipboard= jinbunshoin
		clipboard:= "Jinbun Shoin"
	if clipboard= jiyū kokuminsha
		clipboard:= "Jiyū Kokuminsha"
	if clipboard= jiritsushobō
		clipboard:= "Jiritsu Shobo"

;k
	if clipboard= kadokawashoten
		clipboard:= "Kadokawa Shoten"
	if clipboard= kaishahōkankeihōmushōrei
		clipboard:= "Kaishahō Kankei Hōmu Shōrei"
	if clipboard= kanaeshobō
		clipboard:= "Kanae Shobō"
	if clipboard= kanagawashinbunsha
		clipboard:= "Kanagawa Shinbunsha"
	if clipboard= kasamashoin
		clipboard:= "Kasama Shoin"
	if clipboard= kawadeshobōshinha
		clipboard:= "Kawade Shobo Shinsha"
	if clipboard= kawade shobōshinsha
		clipboard:= "Kawade Shobo Shinsha"
	if clipboard= kawadeshobōshinsha
		clipboard:= "Kawade Shobo Shinsha"
	if clipboard= Kawade Shobōshinsha
		clipboard:= "Kawade Shobo Shinsha"
	if clipboard= kawadeshobō
		clipboard:= "Kawade Shobo"
	if clipboard= kazamashobō
		clipboard:= "Kazama Shobō"
	if clipboard= keiōgijukudaigakushuppankai
		clipboard:= "Keiō Gijuku Daigaku Shuppankai"
	if clipboard= keizaisangyōchōsakai
		clipboard:= "Keizai Sangyō Chōsakai"
	if clipboard= kenbun
		clipboard:= "Kenbun"
	if clipboard= keisoshobo
		clipboard:= "Keiso Shobō"
	if clipboard= kin'yūzaiseijijōkenkyūkai
		clipboard:= "Kin'yū Zaisei Jijō Kenkyūkai"
	if clipboard= kitaōjishobo
		clipboard:= "Kitaōji Shobo"
	if clipboard= kōraidaigakkōgurōbarunihonkenkyūin
		clipboard:= "Kōrai Daigakkō Gurōbaru Nihon Kenkyūin"
	if clipboard= kōraidaigakkōnihonkenkyūsentā
		clipboard:= "Kōrai Daigakkō Nihon Kenkyū Sentā"
	if clipboard= kōseitorihikikyōkai
		clipboard:= "Kōsei Torihiki Kyōkai"
	if clipboard= koamagajin
		clipboard:= "Koa Magajin"
	if clipboard= kodansha
		clipboard:= "Kōdansha"
	if clipboard= kokusaishoin
		clipboard:= "Kokusai Shōin"
	if clipboard= kyōeishobō
		clipboard:= "Kyōei Shobo"
	if clipboard= Kyōtodaigakugakujutsushuppankai
		clipboard= "Kyōto Daigakugakujutsu Shuppankai"
	if clipboard= kyūkoshoin
		clipboard:= "Kyūko Shoin"
;m
	if clipboard= mainichishinbunsha
		clipboard:= "Mainichi Shinbunsha"
	if clipboard= mainichishinbunShuppan
		clipboard:= "Mainichi Shinbun Shuppan"
	if clipboard= matsurikasha
		clipboard:="Matsurikasha"
	if clipboard= meijitoshoshuppan
		clipboard:= "Meiji Tosho Shuppan"
	if clipboard= meitokushuppansha
		clipboard:= "Meitoku Shuppansha"
	if clipboard= minerubashobō
		clipboard:= "Mineurba Shobō"
	if clipboard= mineruvashobō
		clipboard:= "Mineurva Shobō"
	if clipboard= minjihōkenkyūkai
		clipboard:= "Minjihō Kenkyūkai"
	if clipboard= misuzushobo
		clipboard:= "Misuzu Shobō"
	if clipboard= miyaishoten
		clipboard:= "miyai Shoten"
	if clipboard= miyaobishuppansha
		clipboard:= "Miyaobi Shuppansha"
	if clipboard= musashinobijutsudaigakuShuppankyoku
		clipboard:= "Musashino Bijutsu Daigaku Shuppankyoku"
;n
	if clipboard= nichigaiasoshiētsu
		clipboard:= "Nichigai Asoshiētsu"
	if clipboard= nihonkajoshuppan
		clipboard:= "Nihon Kajo Shuppan"
	if clipboard= nihonkeizaishinbunsha
		clipboard:= "Nihon Keizai Shinbunsha"
	if clipboard= nihonhyōronsha
		clipboard:= "Nihon Hyōronsha"
	if clipboard= nihonkyōhōsha
		clipboard:= "Nihon Kyōhōsha"
	if clipboard= nihonkyohosha
		clipboard:= "Nihon Kyōhōsha"
	if clipboard= nikkeibipimaketingu
		clipboard:= "Nikkei BP Māketingu"
	if clipboard= nikkeinashonarujiogurafikkusha
		clipboard:= "Nikkei Nashonaru Jiogurafikkusha"
;o
	if clipboard= ōkurazaimukyōkai
		clipboard:= "Ōkura Zaimu Kyōkai"
	if clipboard= ōtsukishoten
		clipboard:= "Ōtsuki Shoten"
	if clipboard= ochanomizushobō
		clipboard:= "Ochanomizu Shobō"
	if clipboard= otowashobotsurumishoten
		clipboard:= "Otowa Shobō Tsurumi Shoten"
;p
	if clipboard= paintanashonaru
		clipboard:= "Pai Intānashonaru"
;r
	if clipboard= rikkashuppan
		clipboard:= "Rikka Shuppan"
	if clipboard= rōdōchōsakai
		clipboard:= "Rōdō Chōsakai"
	if clipboard= rōmugyōsei
		clipboard:= "Rōmu Gyōsei"
;s
	if clipboard= san'ichishobō
		clipboard:= "San{U+0027}ichi Shobō"
	if clipboard= san'ninsha
		clipboard:= "San'ninsha"
	if clipboard= sanninsha
		clipboard:= "San'ninsha"
	if clipboard= seirinshoin
		clipboard:= "Seirin Shoin"
	if clipboard= sekaishisōsha
		clipboard:= "Sekai Shisōsha"
	if clipboard= serikashobō
		clipboard:= "Serika Shobo"
	if clipboard= shinchosha
		clipboard:= "Shinchōsha"
	if clipboard= shinnihonhōkishuppan
		clipboard:= "Shin Nihon Hōki Shuppan"
	if clipboard= shinnihonshuppansha
		clipboard:= "Shinnihon Shuppansha"
	if clipboard= shōjihōmu
		clipboard:= "Shōji Hōmu"
	if clipboard= shueishaintanashonaru
		clipboard:= "Shueisha Intānashonaru"
	    if clipboard= shueisha
		clipboard:= "Shūeisha"
	if clipboard= sogensha
		clipboard:= "Sōgensha"
;t
	if clipboard= tachibanashobō
		clipboard:= "Tachibana Shobō"
	if clipboard= taiseiShuppansha
		clipboard:= "Taisei Shuppansha"
	if clipboard= Taishūkanshoten
		clipboard:= "Taishūkan Shoten"
	if clipboard= tōhōshoten
		clipboard:= "Tōhō Shoten"
	if clipboard= tōkyōdaigakushuppankai
		clipboard:= "Tōkyō Daigaku Shuppankai"
	if clipboard= tōkyōhōreishuppan
		clipboard:= "Tōkyō Hōrei Shuppan"
;u
	if clipboard= ueibuShuppan
		clipboard:= "Ueibu Shuppan"
;y
	if clipboard= yachiyoshuppan
		clipboard:= "Yahiyo Shuppan"
	if clipboard= yoshikawakōbunkan
		clipboard:= "Yoshikawa Kōbunkan"
	if clipboard= yoshikawakobunkan
		clipboard:= "Yoshikawa Kōbunkan"
	if clipboard= yumanishobō
		clipboard:= "Yumani Shobo"
	if clipboard= yūshindōkōbunsha
		clipboard:= "Yūshindō Kōbunsha"
}

\::
send {ctrl}
sleep 100
send {shift}
reload

numPadSub::
send {ctrl}
sleep 100
send {shift}
reload

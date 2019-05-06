#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

setTitleMatchMode, 2
setKeyDelay, 300

;all variables, treated as global and set with no value
	;variables for navigation
	;these variables are used as quasi if/then conditionals to change the flow of the code executed
		global vVcloseTab:=""
		global vVdoTitle:= "" ; for navigating chinook dupe search
		global vVstopHere:=""
		global vvTitleSearch:= "" ;appends the ISBN number the WorldCat bib data that will be pasted to the Google Sheet

	;variables for storing bib data
		global vvOCLC:= "" ;holds OCLC number
		global vvAuthorR:= "" ;holds romanized author
		global vvAuthorV:= "" ;holds vernacular author
		global ISBN:= ""
		global vvISBN:= "" ;holds ISBN number, specifically the ISBN-13
		global vvPubDate:= "" ;holds date of publication
		global vvPublisherR:= ""
		global vvPublisherV:= ""
		global vvTitleR:= "" ;holds romanized title
		global vvTitleV:= "" ;holds vernacular title
		global vvSeriesTitle:= "" ;holds series title (romanized only)
		global vvSourceCode:= "" ;holds source code from a web page to be parsed once or multiple times
		
		
	;variables for general use
		global vVcellData:= "" ;holds data from a cell copied in the Google sheet

	;variables for WorldCat searches
		global vVwcSearchURL:= ""
		global vVjustOCLC:= ""
		
	;variables primarily for dupe checking in chinook
		global vVchinookURL:=""
		global vVcopyOrdered=""
		
	;variables primarily for checking price on amazon.co.jp
		global vVamazonJPsearchURL:= ""
		global vVpriceYen:= ""
		global vVcurrencyConvertor:=""
		global vVpriceUSD:= ""

insert::
	send sleep 100
return

+esc::
	send {raw}msgBox, `%clipboard`%
	sleep 100
	send {enter}
	sleep 100
	send {raw}exit
	sleep 100
	send {enter}
return

;Open chinook classic
#c::
	chromeActivate()
	chromeConfirm()
	run http://libraries.colorado.edu/
return

;Get chinook record URL
^+u::
	if winExist("Chinook Library Catalog - ")
		{
		winActivate
		sleep 100
		send ^f
		sleep 100
		send {delete}
		sleep 100
		send {raw}http://libraries.colorado.edu/record=
		sleep 100
		send {esc}
		sleep 100
		clipboard=
		send ^c
		clipWait 1
		if clipboard= http://libraries.colorado.edu/record=
			{
			send ^+{right 2}
			sleep 100
			send clipboard=
			send ^c
			clipWait 1
			send ^{home}
			exit
			}
		}
return

;Copy ISBNs from Google sheet and pastes them to WorldCat.org to search, loop
^+L::
	chromeActivate()
	chromeConfirm()
	run https://docs.google.com/spreadsheets/d/1w15TVODE2I_gUkQHD5-Ghcskjc1bs9O21Moi6c9TyOA/edit#gid=999759851
return

;Loop Saerch for WorldCat Data with ISBN and ONLY get OCLC#
^+PrintScreen::
	vVjustOCLC:= "do"
	goTo JustOCLCsearch
;Single Search for WorldCat data with ISBN or Title Search and get bib data
PrintScreen::
numpadEnter::
	vVstopHere:= "stop"
	goTo stopAtWorldCatResults
;Loop Search for WorldCat data with ISBN or Title Search and get bib data
^PrintScreen::
	vVjustOCLC:= ""
	JustOCLCsearch:
	vVstopHere:= ""
	stopAtWorldCatResults:
	googleSheetConfirm()
	bibSearch:
	cellResetRowToISBN()
	cellCopyCell()
		if vVjustOCLC = do
			if clipboard=
			{
			msgBox, There is no data to search. The macro for only getting the OCLC# has stopped.
			exit
			}
;tabs over to "Title (R/E)" is "ISBN" column is empty
	if clipboard=
		{
		vvTitleSearch= do
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		cellCopyCell()
;if "Title (R/E)" column is ALSO empty, moves to"Title (V)" column
		if clipboard= 
			{
			send {right}
			cellCopyCell()
				if clipboard=
					{
					msgBox, There is no data to search. The macro has stopped.
					exit
					}
					else
						{
						send {left}
						sleep 100
						}
;if there is title data, tabs back to "ISBN" column before doing a title search to paste WorldCat bib data in the correct place
			}
			send {left}
			sleep 100
			send {left}
			sleep 100
			send {left}
			sleep 100
			send {left}
			sleep 100
			send ^n
			sleep 2000
			send !d
			sleep 100
			send {raw}worldcat.org
			sleep 100
			send {tab}
			sleep 100
			send {raw}ti:
			sleep 100
			send ^v
			sleep 100
			clipboard=
			sleep 100
			send {enter}
			sleep 3000
;Title data is unreliable and needs to be checked manually, the macro stops after a title search
			exit
		}
;Opens new Google chrome window to do an ISBN search
		send ^n
		sleep 2000
		send !d
		sleep 100
		send {raw}worldcat.org
		sleep 100
		send {tab}
		sleep 100
		send {raw}bn:
		sleep 100
		send ^v
		sleep 100
		clipboard=
		sleep 100
		send {enter}
		sleep 3000
		if vVstopHere= stop
			{
			vVstopHere=
			exit
			}
;If hot key for just looking at WorldCat search results was used, macro will stop here
	worldCatResultsNone()
		if clipboard= none
			{
			send ^w
			sleep 100
			worldCatToGoogleSheetBibData()
			goTo, bibSearch
			}
	worldCatResultsTen()
		if clipboard= 1-10
			{
			send ^w
			sleep 100
			worldCatToGoogleSheetBibData()
			goTo, bibSearch
			}
	getWorldCatOneResult:
	worldCatResultsOne()
		if clipboard= 1-1
			{
			send {tab 9}
			sleep 100
			send {enter}
			sleep 3000
			goTo getWorldCatSourceCode
			}
	worldCatResultsMultiple()
		send !{F4}
		sleep 300
		worldCatToGoogleSheetBibData()
		goTo bibsearch
	getWorldCatSourceCode:
		send ^u	
		sleep 3000
		sourceCodeConfirm()
		send ^f
		sleep 100
		send {delete}
		sleep 100
		send {raw}OCLC:
		sleep 100
		send {esc}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
			if errorLevel, 1
				{
				msgBox, Error: the macro could not confirm the source code has loaded.
				exit
				}
		send ^a
		sleep 300
		clipboard=
		send ^c
		clipWait, 3
		clipboard:= regExReplace(clipboard, "`r`n", "")
		vvSourceCode:= clipboard
			if vVcloseTab= do
				{
				send ^w
				sleep 100
				send ^w
				sleep 500
				send {left} ;moves active cell back to OCLC# to properly paste data to Google sheet
				sleep 100
				goTo closedWithCtrlW
				}
		send !{F4}
		sleep 500
	closedWithCtrlW:
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
		fixDiacritics()
		fixPubR()
		vvPublisherR:= clipboard
	parsePublisherVfromWC()
	parsePublisherDatefromWC()
	if vVjustOCLC= do
		{
		clipboard:= vvOCLC
		send {tab}
		sleep 100
		send ^+v
		sleep 500
		send {down}
		sleep 100
		goTo bibSearch
		}
	clipboard:= vvOCLC . "`t" . vvAuthorR . "`t" . vvAuthorV . "`t" . vvTitleR . "`t" . vvTitleV . "`t" . vvSeriesTitle . "`t" . vvPublisherR . "`t" . vvPublisherV . "`t" . vvPubDate
		if vvTitleSearch= do
			{
			clipboard:= vvISBN . "`t" . clipboard
			sleep 100
			}
	fixDiacritics()
	send {tab}
	sleep 100
	send ^+v
	sleep 500
	send {down}
	sleep 100
	if vVstopHere= stop
		{
		vVstopHere=
		exit
		}
	goto bibSearch
return

;Pull WorldCat bib data from search with only one result in WorldCat
F3::
numpad3::
	vVstopHere:= "stop"
	vVcloseTab:= "do"
	ifWinExist Results for '
		WinActivate
			else
				{
				msgBox, This macro only works on a WorldCat Search result page for title searches. The macro has stopped.
				exit
				}
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}<< return to search results
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	sleep 100
	send ^c
	sleep 100
	if clipboard= << return to search results
		goTo wrongF3Macro
	worldCatResultsNone()
		if clipboard= no
			goto wrongF3Macro
	worldCatResultsTen()
		if clipboard contains 10+ results
			goTo wrongF3Macro
	worldCatResultsOne()
	if clipboard= 1-1
		goTo getWorldCatOneResult
	wrongF3Macro:
		msgbox, The macro you used is only for a WorldCat search result screen with only one entry. Please select the best entry and run the F4 macro.
return

;Loads the source code of WorldCat record.
F6::
numpad6::
	vVstopHere:= "stop"
	vVcloseTab:= "do"
	IfWinExist WorldCat.org
		WinActivate
			else
			{
			msgBox, There is no window with WorldCat results open OR the tab in Chrome with the WorldCat results is not currently active. The macro has stopped.
			exit
			}
	send ^f
	sleep 100
	send {raw}add to list
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	sleep 100
	send ^c
	sleep 100
		if clipboard!= add to list
			{
			msgBox, The macro you used is only for a WorldCat record page to load the source code and extract the bliographic data.
			exit
			}
	send ^f
	sleep 100
	send {raw}search results for
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	sleep 100
	send ^c
	sleep 100
		if clipboard= search results for
			{
			msgBox, The macro you used is only for a WorldCat record page to load the source code and extract the bibliographic data.
			exit
			}
	sleep 400
	goTo getWorldCatSourceCode
return

; Uses Chinook Classic (http://libraries.colorado.edu/) to determine if the library already owns a book via ISBN or TITLE
^+E::
	titleDupeCheck:
	vVdoTitle:= "do"
	goTo skipISBNdupeCheck
^+I::
		vVdoTitle:= ""
	ISBNdupeCheck:
		vVcellData:= ""
		vVchinookURL:= ""
		vVcopyOrdered:= ""
	skipISBNdupeCheck:
	googleSheetConfirm()
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
		if vVdoTitle= do
			{
			send {right}
			sleep 100
			send {right}
			sleep 100
			send {right}
			sleep 100
			send {right}
			sleep 100
			}
	copyCellForDupe:
	cellCopyCell()
	vVcellData:= clipboard
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
		if vVdoTitle= do
			{
			send {tab 1}
			goTo skipSetISBNsearch
			}
		send {down 14}
		sleep 100
		send {tab}
		sleep 100
	skipSetISBNsearch:
	clipboard=
	clipboard:= vVcellData
	clipWait 1
	send ^v
	sleep 100
	send {tab 2}
	sleep 100
	send {enter}
	sleep 3000
;confrim chinook results have "no entires found"
	send ^f
	sleep 100
	send {raw}no matches found
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= no matches found
		{
		clipboard= no
		goTo backToGoogleSheet
		}
;confirm chinook results have multiple results, for title only. Have not yet encountered an ISBN search with multiple results
	if vVdoTitle= do
		{
		send ^f
		sleep 100
		send {raw}titles (1-
		sleep 100
		send {esc}
		sleep 100
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
			vVchinookURL:= clipboard
			clipboard:= "=hyperlink(""" . vVchinookURL . """,""" . "2+ results"")"
			goTo backToGoogleSheet
			}
		}
;confirm chinook results have one results, record will have permalink	
	send ^f
	sleep 100
	send {raw}http://libraries.colorado.edu/record=
	sleep 100
	send {esc}
	sleep 100
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
		vVchinookURL:= clipboard
		}
;after confirming one result, check if item is ordered or not
	send ^f
	sleep 100
	send {raw}copy ordered for
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= copy ordered for
		{
		send ^+{down}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		clipboard:= regExReplace(clipboard, "`r`n", "")
		clipboard:= regExReplace(clipboard, ".Description", "")
		vVcopyOrdered:= clipboard
		clipboard:= "=hyperlink(""" . vVchinookURL . """,""" . vVcopyOrdered . """" . ")"
		goto backToGoogleSheet
		}
	if vVchinookURL contains http://libraries.colorado.edu/record=
		clipboard:= vVchinookURL
		goTo backToGoogleSheet
;in case of an unanticipated scenario, the clipboard will not containt any text strings that trigger a return to the Google sheet
	msgBox, An unexpected outcome has occurred while checking Chinook for an ISBN duplicate. The macro has stopped and may need new code to account for this new situation.
	exit
;data collected in Chinook for dupe check goes back to Google Sheet
		backToGoogleSheet:
		send !{F4}
		sleep 300
		send {home}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
		sleep 100
		send {tab}
;tabs one more time to move to title duple column
		if vVdoTitle= do
			{
			send {tab}
			sleep 100
			}
		sleep 100
		send {delete}
		sleep 100
		send ^+v
		sleep 100
		send {down}
		sleep 100
		if vVdoTitle= do
			{
			goto titleDupeCheck
			}
		goto ISBNdupeCheck
return

;Looks up price and stops macro at the search results
^4::
	vVstopHere:= "stop"
	goTo justAmazonJPsearchResults
; Price check with Amazon, uses data from amazon.co.jp search results with only 1 result
^+4::
	vVstopHere:= ""
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
		clipWait, 1
	googleSheetToAmazonJP()
;check the variable "vVstopHere" when the macro for searching 1 item is run
	if vVstopHere= stop
		{
		vVstopHere=
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
		send {raw}No results, check manually
		sleep 100
		send {down}
		sleep 100
		goTo justAmazonJPsearchResults
		}
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}キーワードを絞るか
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
		send {raw}No results, check manually
		sleep 100
		send {down}
		sleep 100
		goTo justAmazonJPsearchResults
		}
;URL for search results saved in "vVamazonJPsearchURL" variable
	send !d
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	vVamazonJPsearchURL:= clipboard	
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
			vVamazonJPsearchURL:= clipboard
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
			vVpriceYen:= clipboard
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
		msgBox, Something went wrong when trying to copy the price data from the source code of an amazon.co.jp search result. The macro has stopped.
		exit
		}
;if the text string :a-price-whole" IS found, the price data is extracted and sent to 
	if clipboard= a-price-whole
		send ^+{right}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		clipboard:= regExReplace(clipboard, ".*\>", "")
		clipboard:= regExReplace(clipboard, ",", "")
		global vVpriceYen:= clipboard
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
		clipboard:= "=GOOGLEFINANCE(""CURRENCY:JPYUSD"")"
		clipWait, 1
		send ^+v
		sleep 300
		send {down}
		sleep 300
	;copy formula value and paste as absolute value, then store in vVcurrencyConvertor
		send {up}
		sleep 100
		clipboard=
		send ^c
		clipWait, 1
		send ^+v
		sleep 300
		send ^c
		sleep 100
		 vVcurrencyConvertor:= clipboard
;calculate price in US dollars
		clipboard:= vVpriceYen * vVcurrencyConvertor
		global vVpriceUSD:= clipboard
;paste vVpriceYen with hyperlink to amazon.co.jp results
		send {left}
		sleep 100
		send {delete}
		sleep 100
		clipboard:= "=hyperlink(""" . vVamazonJPsearchURL . """,""" . vVpriceYen """)"
		send ^+v
		sleep 500
;vVpriceUSD variable is pasted in column T
		send {left}
		sleep 100
		send {left}
		sleep 100
		send {left}
		sleep 100
		clipboard=
		clipboard:= vVpriceUSD
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

;Copies data from a cell in a specific way to ensure no "invisible" encoding is copied with cell contents
cellCopyCell()
	{
	sendMode input
	send {f2}
	sleep 100
	send ^+{home}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	send {esc}
	sleep 100
	}

;Tabs cell cursor to the ISBN column cell cursor is in
cellResetRowToISBN()
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
	
;Tabs cell cursor to the Title (R/E) column cell cursor is in
cellResetRowToTitle()
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
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	send {right}
	sleep 100
	}

;Loads http://libraries.colorado.edu/
chinookGoto()
	{
	sendMode input
	send ^n
	sleep 2000
	send !d
	sleep 100
	send {raw}http://libraries.colorado.edu/
	sleep 100
	send {enter}
	sleep 4000
	}

;Confirms Chinook web page had loaded has loaded
chinookConfirm()
	{
	sendMode input
	send ^f
	sleep 200
	send {delete}
	sleep 100
	send {raw}My Chinook
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if clipboard!= My Chinook
		{
			msgbox, The website "libraries.colorado.edu/" has failed to load, the macro has stopped.
			exit
		}
	}

;If multiple results, will copy URL in address bar and post link as "2+ results" for Google Sheet
;chinookISBNresultsMultiple()
;	{
;	
;	}


;If multiple results, will copy URL in address bar and post link as "2+ results" for Google Sheet
chinookTitleResultsMultiple()
	{
	sendMode input
	send ^f
	sleep 100
	send {raw}titles (1-
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if clipboard= titles (1-
		{
		send !d
		sleep 100
		send ^c
		sleep 100
		send {raw}=hyperlink("
		sleep 100
		send ^v
		sleep 100
		send {raw}","2+ results")
		sleep 100
		send ^a
		sleep 100
		send ^x
		}
	}


;Checks if Chrome is already running, opens it if not, and loads the Google sheet
chromeActivate()
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
			msgbox, Google Chrome is not open, not installed on your computer, or not stored in the file path C:\Program Files (x86)\Google\Chrome\Application\chrome.exe. The macro will stop.
			exit
			}
	sleep 1000
	}

;Confirms Chrome is open, stops macro if not
chromeConfirm()
	{
	if winExist("ahk_exe chrome.exe")
		WinActivate
			else
			{
			msgbox, Google Chrome is not open, The macro will stop.
			exit
			}
	sleep 1000
	}

clipboardEmptyStop()
	{
	if clipboard=
		{
		msgBox, There is no data to search. The macro has stopped.
		exit
		}
	}

;Confirms the Google sheet "Material Selection for Japanese-Korean-Studies Materials" is open, stops macro if not
googleSheetConfirm()
	{
	if winExist("Material Selection for Japanese-Korean")
		winActivate
		else
		{
		msgBox, Chrome might be open but the tab for the GoogleSheet "Material Selection for Japanese Korean related Materials" is not active. Activate that tab and try running the macro again.
		}
	}

;opens new Chrome window and searches Amazon.com directly from address bar
googleSheetToAmazonJP()
	{
	sendMode Input
	send ^n
	sleep 2000
	send !d
	sleep 100
	send {raw}amazon.co.jp
	sleep 100
	send {tab}
	sleep 100
	send ^v
	sleep 100
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
		msgbox, The source code tab did not open. The macro has stopped.
		exit
		}
	}

;After all other possible WorldCat results have been considered, this is the last possibility, multiple results less than 10
worldCatResultsMultiple()
	{
	sendMode input
	send !d
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel= 1
		{
		msgBox, Data has not been copied to the clipboard when trying to copy the URL of WorldCat search results that resulted in 2+ records.
		exit
		}
	vVwcSearchURL:= clipboard
	clipboard:= "=hyperlink(""" . vVwcSearchURL . """,""" . "2+ results"")"
	}

;If no results in WorldCat, will copy "" to clipboard
worldCatResultsNone()
	{
	sendMode input
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}no results match your search
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	send ^c
	clipWait, 1
	if errorLevel= 1
		clipboard= none
	}

;Confirms 1 result in WorldCat
worldCatResultsOne()
	{
	sendMode input
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}1-1
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	sleep 100
	send ^c
	sleep 100
	}

;Confirms 10 or more results in WorldCat. This must be done before determining one result becuase the unique text string for one result is "1-1" which looks like the beginning of "1-10"
worldCatResultsTen()
	{
	sendMode input
	sleep 100
	send ^f
	sleep 100
	send {delete}
	sleep 100
	send {raw}1-10
	sleep 100
	send {esc}
	sleep 100
	clipboard=
	sleep 100
	send ^c
	sleep 100
	if clipboard= 1-10
		{
		send !d
		sleep 100
		send clipboard=
		send ^c
		clipWait, 1
		if errorLevel= 1
			{
			msgBox, Data has not been copied to the clipboard when trying to copy the URL of WorldCat search results that resulted in 10+ records.
			exit
			}
		vVwcSearchURL:= clipboard
		send clipboard=
		clipboard= "=hyperlink(""" . vVwcSearchURL . """,""" . "10+ results"")"
		sleep 100
		}
	}

;Puts data in columns K-S in the Google sheet "Material Selection for Japanese-Korean-Studies Materials"
worldCatToGoogleSheetBibData()
	{
	sendMode input
	sleep 100
	send {tab}
	sleep 100
	send {delete}
	sleep 100
	send ^+v	
	sleep 300
	send {down}
	sleep 100
	}

;=====
;=====
; Trim source code variables with regular expressions
;=====
;=====

;t::
;parseISBNfromWC()
;exit

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
	sleep 100
	send ^v
	sleep 100
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
	if clipboard= chikurashobō
		clipboard:= "Chikura Shobō"
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
	if clipboard= minjihōkenkyūkai
		clipboard:= "Minjihō Kenkyūkai"
	if clipboard= misuzushob
		clipboard:= "Misuzu Shobo"
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

pause::
send {ctrl up}
sleep 100
send {shift up}
reload

\::
send {ctrl up}
sleep 100
send {shift up}
reload

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay 100

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
		global vvAuthorR:= "" ;holds romanized author
		global vvAuthorV:= "" ;holds vernacular author
		global vvBibArray:= [] ;converts spreadsheet cells into an array
		global vvISBN13:= "" ;holds ISBN13 number
		global vvISBN10:= "" ;holds ISBN10 number
		global vvOCLC:= "" ;holds OCLC number
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

;==============================
;---------------
;These are the primary macros to handle taking data from the spreadsheet, using that data to search worldcat.org, and extract worldcat.org's data back to the spreadsheet.
;---------------
PrintScreen:: ;Searches worldcat.org with ISBN from spreadsheet and stops at search results. If there is no ISBN (column K) data, a search will be attempted with romanized/English title data (column O), if this title data is not available, a final attempt will be made using the vernacular title data (column N). Should be used in conjunction with other macros to pull data into spreadsheet one by one.
vvStopHere:= "stop"
;==============================
^PrintScreen:: ;Same as above, but will extract data from worldcat.org automatically.
bibSearch:
vvBibArray:= []
;---------------
;Gets data source from spreadsheet: prioritizes ISBN searches, then does vernacular title searches, than romanizes/English searches.
;---------------
googleSheetConfirm()
send {home}
send +{space}
send +{space}
send ^c
vvBibArray:= StrSplit(clipboard, A_Tab) ;stores each cell as it's own variable: vvBibArray[1] equals data in column A, vvBibArray[2] equals data in column B, etc.
if vvBibArray[2]= "" ;Deals with row selection behavior in Google sheets. If column B is empty, the send +{space} command sent twice will deactivate the row selection. The code below will not impact performance in Excel
	{
	send +{space}
	send ^c
	vvBibArray:= StrSplit(clipboard, A_Tab)
	}
;---------------
;Checks if there is an ISBN (column M) to use to search worldcat.org. If not, searches with romanized title (column Q). If not, then search with vernacular title (column P)
;---------------
if vvBibArray[15]!= ""
	{
	send ^n
	sleep 2000
	send !d
	sendInput {raw}worldcat.org
	send {tab}
	sendInput {raw}no:
	sendInput % vvBibArray[15] ;Vernacular title
	send {enter}
	exit
	}
if vvBibArray[13]= "" ;ISBN-13
	{
	if vvBibArray[17]= "" ;Romanized/English title
		{
		if vvBibArray[16]= "" ;Vernacular title
			{
			msgBox, There is no data, the macro has stopped.
			exit
			}
			else
				{
				send ^n
				sleep 2000
				send !d
				sendInput {raw}worldcat.org
				send {tab}
				sendInput {raw}mt:bks NOT mt:elc ti:
				sendInput % vvBibArray[16] ;Vernacular title
				send {enter}
				exit
				}
		}
	else
		{
		send ^n
		sleep 2000
		send !d
		sendInput {raw}worldcat.org
		send {tab}
		sendInput {raw}mt:bks NOT mt:elc ti:
		sendInput % vvBibArray[17] ;Romanized/English title
		send {enter}
		exit
		}
	}
;---------------
;Searches worldcat.org with Google Chrome browser. 
;---------------
send ^n
sleep 2000
send !d
sendInput {raw}worldcat.org
sleep 100
send {tab}
sendInput {raw}mt:bks NOT mt:elc bn:
sleep 100
sendInput % vvBibArray[13] ;ISBN-13
sleep 100
send {enter}
if vvStopHere= stop
	{
	vvStopHere= ""
	exit
	}
sleep 2000
;---------------
;Load Check
;---------------
Loop, 10
	{
	sleep 1000
	send {tab}
;---------------
;Copies all text on the page. The macro will branch out depeding on one of three outcomes: no results, one result, or multiple results.
;---------------
	send ^a
	send ^c
	if clipboard contains WorldCat is the world's largest library catalog
		break
	}
	if clipboard not contains WorldCat is the world's largest library catalog
		{
		msgBox, worldcat.org has not loaded properly, the macro has stopped.
		exit
		}
;---------------
;No results in worldcat.org.
;---------------
if clipboard contains results match your search for
	{
	send !{F4}
	googleSheetConfirm()
	vvBibArray[13]:= "no WC record"
	clipboard:= vvBibArray[1] . "`t" . vvBibArray[2] . "`t" . vvBibArray[3] . "`t" . vvBibArray[4] . "`t" . vvBibArray[5] . "`t" . vvBibArray[6] . "`t" . vvBibArray[7] . "`t" . vvBibArray[8] . "`t" . vvBibArray[9] . "`t" . vvBibArray[10] . "`t" . vvBibArray[11] . "`t" . vvBibArray[12] . "`t" . vvBibArray[13]
	send ^v
	sleep 1000 ;this delay is needed for the "down" command on the next line to work correctly, not known why.
	send {down}
	sleep 1000
	goTo bibSearch 
	}
;---------------
;One result in worldcat.org.
;---------------
if clipboard contains Results 1-1 of about 1
	{
	send {tab 11}
	send {enter}
	sleep 3000
;---------------
;Load Check
;---------------
Loop, 10
	{
	sleep 1000
	send {tab}
	send ^a
	send ^c
	if clipboard contains << Return to Search Results
		break
	}
	if clipboard not contains << Return to Search Results
		{
		msgBox, A worldcat.org record has not loaded properly, the macro has stopped.
		exit
		}
if clipboard contains Boulder,, CO 80309 United States ;The double comma is needed for AHK to see a single normal comma.
	vvBibArray[12]:= "y"
	else
		vvBibArray[12]:= "n"
}
;---------------
;Multiple results in worldcat.org. This is the only possibility after verifying the previous two possiblities (no results, one result) haven't happened. A URL will be pasted to the speadsheet in the OCLC# column and you can verify manually the best worldcat.org record. Use the two macros after this one for that process.
;---------------
	else
		{
		send !d
		clipboard=
		send ^c
		send !{F4}
		googleSheetConfirm()
		vvBibArray[15]:= clipboard
		clipboard:= vvBibArray[1] . "`t" . vvBibArray[2] . "`t" . vvBibArray[3] . "`t" . vvBibArray[4] . "`t" . vvBibArray[5] . "`t" . vvBibArray[6] . "`t" . vvBibArray[7] . "`t" . vvBibArray[8] . "`t" . vvBibArray[9] . "`t" . vvBibArray[10] . "`t" . vvBibArray[11] . "`t" . vvBibArray[12] . "`t" . vvBibArray[13] . "`t" . vvBibArray[14] . "`t" . vvBibArray[15]
		send ^v
		sleep 1000 ;this delay is needed for the "down" command on the next line to work correctly, not known why.
		send {down}
		sleep 1000
		vvBibArray= []
		goTo bibSearch
		}
;---------------
;The worldcat.org item record's source code will be opened and the bibliographic data will be parsed apart, and populated into the correct vvBibArray variable to be pasted to the spreadsheet.
;---------------
getSourceCode:
send ^u	
sleep 2000
;---------------
;Load Check
;---------------
Loop, 10
	{
	sleep 1000
	send {tab}
	send ^a
	send ^c
	if clipboard contains OCLC:
		break
	}
	if clipboard not contains OCLC:
		{
		msgBox, the sourcecode for a worldcat.org record has not loaded properly, the macro has stopped.
		exit
		}
send ^a
sleep 500
send ^c
clipboard:= regExReplace(clipboard, "`r`n", "") ;Try also replacing `t with "" - this may cause problems with parsing though.
;clipboard:= regExReplace(clipboard, "`t", "")
vvSourceCode:= clipboard
;---------------
;The Google Chrome browser window opened to search worldcat.org is closed, returning to the previously active Google Chrome browser window with the spreadsheet in it.
;---------------
send !{F4}
googleSheetConfirm()
sleep 500
parseISBN13fromWC()
vvBibArray[13]:= vvISBN13 ;column M
parseISBN10fromWC()
vvBibArray[14]:= vvISBN10 ;column N
parseOCLCfromWC()
if vvBibArray[14]!= ""
	{
	vvBibArray[7]:= "https://www.amazon.jp/dp/" . vvBibArray[14] ;column G
	if vvBibArray[7]= "https://www.amazon.jp/dp/no ISBN"
		vvBibArray[7]:= "-"
	}
vvBibArray[7]:= regExReplace(vvBibArray[7], " .*", "") ;If a record pulled from worldcat.org has multiple ISBNs, this will create an amazon.jp link using only the first ISBN-10.
vvBibArray[15]:= vvOCLC ;column O
parseTitleVfromWC()
vvBibArray[16]:= vvTitleV ;column P
parseTitleRfromWC()
vvBibArray[17]:= vvTitleR ;column Q
parseAuthorRfromWC()
vvBibArray[18]:= vvAuthorR ;column R
parseAuthorVfromWC()
vvBibArray[19]:= vvAuthorV ;column S
parseSeriesTitlefromWC()
vvBibArray[20]:= vvSeriesTitle ;column T
parsePublisherRfromWC()
if vvBibArray[9]= ""
	{
	vvBibArray[9]:= vvBibArray[20] ;This moves the series data into the volume# data if there is a volume number data in the series data, the number will appear in column I.
	vvBibArray[9]:= regExReplace(vvBibArray[9], ".*, ","")
	if vvBibArray[9]= "no series"
		vvbibArray[9]:= "-"
	}
vvBibArray[21]:= vvPublisherR ;column U
parsePublisherVfromWC()
vvBibArray[22]:= vvPublisherV ;column V
parsePublisherDatefromWC()
vvBibArray[23]:= vvPubDate ;column W
clipboard:= vvBibArray[1] . "`t" . vvBibArray[2] . "`t" . vvBibArray[3] . "`t" . vvBibArray[4] . "`t" . vvBibArray[5] . "`t" . vvBibArray[6] . "`t" . vvBibArray[7] . "`t" . vvBibArray[8] . "`t" . vvBibArray[9] . "`t" . vvBibArray[10] . "`t" . vvBibArray[11] . "`t" . vvBibArray[12] . "`t" . vvBibArray[13] . "`t" . vvBibArray[14] . "`t" . vvBibArray[15] . "`t" . vvBibArray[16] . "`t" . vvBibArray[17] . "`t" . vvBibArray[18] . "`t" . vvBibArray[19] . "`t" . vvBibArray[20] . "`t" . vvBibArray[21] . "`t" . vvBibArray[22] . "`t" . vvBibArray[23]
;***************
;The next two lines help fix a known issue with publisher names romanized by Japanese catalogers. Japanese catalogers do not add spaces to publisher names. the fixPubR function corrects this. It will fix common publisher names and is updated regularly.
;***************
sleep 2000
	fixDiacritics()
send {home}
send ^v
sleep 1000 ;this delay is needed for the "down" command on the next line to work correctly, not known why.
send {down}
sleep 1000
if vvStopHere= stop
	exit
vvStopHere:= ""
goTo bibSearch
return

;==============================
;---------------
;Pulls data from column L of spreadsheet to bring up results for searches with more than one result that must be selected manually.
;---------------
F6::
vvBibArray:= []
googleSheetConfirm()
send {home}
send +{space}
send +{space}
send ^c
vvBibArray:= StrSplit(clipboard, A_Tab) ;stores each cell as it's own variable: vvBibArray[1] equals data in column A, vvBibArray[2] equals data in column B, etc.
if vvBibArray[15]= ""
	{
	msgBox, This is no data, the macro has stopped.
	exit
	}
if clipboard contains www.worldcat.org/search
	{
	send ^n
	sleep 2000
	send !d
	clipboard:= vvBibArray[15] ; not sure why I can use "sendInput" for this variable, and instead need to put it in the clipboard and paste it.
	send ^v
	send {enter}
	exit
	}
	send {home}
	send {esc}
return

;==============================
;---------------
;Works in conjunction with previous macro to export worldcat.org data of an individual worldcat.org book record.
;---------------
F7::
vvStopHere:= "stop"
setTitleMatchMode 2
if winExist(") [WorldCat.org]") ;Makes sure the appropriate window and Google Chrome tab are open and active in order to extract data.
	winActivate
	else
		{
		setTitleMatchMode 1
		msgBox, You are not on an worldcat.org page with an individual item record open to extract data from, the macro has stopped.
		exit
		}
setTitleMatchMode 1
send {tab}
send ^a
send ^c
if clipboard contains Boulder,, CO 80309 United States ;The double comma is needed for AHK to see a single normal comma.
	vvBibArray[12]:= "y"
	else
		vvBibArray[12]:= "n"
if clipboard contains Add to list Add tags Write a review 
	goTo getSourceCode
	else
		{
		sleep 500
		send {left}
		msgBox, You activated the macro for extracting bibliographic form a worldcat.org book record. But you are not on a worldcat.org page with a record. The macro has stopped.
		exit
		}
return

;=====
;=====
; Functions
;=====
;=====

googleSheetConfirm() ;confirms a Google sheet starting with "Collection Development - " is open, stops macro if not
	{
	if winExist("Collection Development - ")
		winActivate
		else
		{
		msgBox, Chrome might be open but the tab for the GoogleSheet "Collection Development - " is not active. Activate that tab and try running the macro again.
		}
	}

;=====
;=====
; Trim source code variables with regular expressions
;=====
;=====

parseISBN13fromWC()
	{
	clipboard:= vvSourceCode
	if clipboard not contains ISBN/ISSN:
		{
		vvISBN13:= "no ISBN"
		return
		}
	clipboard:= regExReplace(clipboard, ".*ISBN/ISSN:", "ISBN/ISSN:")
	clipboard:= regExReplace(clipboard, ".*ISBN/ISSN: .......... ", "")
	clipboard:= regExReplace(clipboard, " .......... ", " ")
	clipboard:= regExReplace(clipboard, " ..........OCLC.*", "")
	clipboard:= regExReplace(clipboard, "OCLC.*", "")
	clipboard:= regExReplace(clipboard, " ..........OCLC.*","")
		clipboard:= regExReplace(clipboard, ".*ISBN/ISSN: ", "")
	if clipboard contains <!DOPUBLIC ;This text string occurs for worldcat.org records with no ISBN. This command clears the variable.
		clipboard:= "no ISBN"
	if clipboard contains PUBLICXHTML ;This text string occurs for worldcat.org records with no ISBN. This command clears the variable.
		clipboard:= "no ISBN"
	vvISBN13:= clipboard
	}

parseISBN10fromWC()
	{
	clipboard:= vvSourceCode
	if clipboard not contains ISBN/ISSN:
		{
		vvISBN10:= "no ISBN"
		return
		}
	clipboard:= regExReplace(clipboard, ".*ISBN/ISSN:", "ISBN/ISSN:")
	clipboard:= regExReplace(clipboard, ".*ISBN/ISSN: ", "")
	clipboard:= regExReplace(clipboard, " ............. ", " ")
	clipboard:= regExReplace(clipboard, " .............OCLC.*", "")
	clipboard:= regExReplace(clipboard, "OCLC.*", "")
	if clipboard contains <!DOPUBLIC ;This text string occurs for worldcat.org records with no ISBN. This command clears the variable.
		clipboard:= "no ISBN"
	if clipboard contains PUBLICXHTML ;This text string occurs for worldcat.org records with no ISBN. This command clears the variable.
		clipboard:= "no ISBN"
	vvISBN10:= clipboard
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
	fixDiacritics()
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
	fixDiacritics()
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
	fixDiacritics()
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
	fixDiacritics()
	fixPubR()
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

;==============================
^+d::
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
	if clipboard= asahishinbunShuppan
		clipboard:= "Asahi Shinbun Shuppan"
;b
	if clipboard= benseishuppan
		clipboard:= "Bensei Shuppan"
	if clipboard= bungeishunju
		clipboard:= "Bungei Shunjū"
	if clipboard= bungeishunjū
		clipboard:= "Bungei Shunjū"
	if clipboard= bunkarongakkai
		clipboard:= "Bunkaron Gakkai"
	if clipboard= bunkashobōhakubunsha
		clipboard:= "Bunka Shobō Hakubunsha"
;c
	if clipboard= chikumashobo
		clipboard:= "Chikuma Shobō"
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
		clipboard:= "Genki Shobō"
	if clipboard= gentosha
		clipboard:= "Gentōsha"
;h
	if clipboard= hayakawashobō
		clipboard:= "Hayakawa Shobō"
	if clipboard= hitsujishob
		clipboard:= "Hitsuji Shobō"
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
		clipboard:= "Jiritsu Shobō"
	if clipboard= jurosha
		clipboard:= "Jurōsha"

;k
	if clipboard= kadokawashoten
		clipboard:= "Kadokawa Shoten"
	if clipboard= kaishahōkankeihōmushōrei
		clipboard:= "Kaishahō Kankei Hōmu Shōrei"
	if clipboard= Kadokawagurupuhorudingusu
		clipboard:= "Kadokawa Gurūpu Hōrudingusu"
	if clipboard= Kadokawagurupupaburisshingu
		clipboard:= "Kadokawa Gurūpu Paburishhingu"
	if clipboard= kanaeshobō
		clipboard:= "Kanae Shobō"
	if clipboard= kanagawashinbunsha
		clipboard:= "Kanagawa Shinbunsha"
	if clipboard= kasamashoin
		clipboard:= "Kasama Shoin"
	if clipboard= kawadeshobōshinha
		clipboard:= "Kawade Shobō Shinsha"
	if clipboard= kawade shobōshinsha
		clipboard:= "Kawade Shobō Shinsha"
	if clipboard= kawadeshobōshinsha
		clipboard:= "Kawade Shobō Shinsha"
	if clipboard= Kawade Shobōshinsha
		clipboard:= "Kawade Shobō Shinsha"
	if clipboard= kawadeshobō
		clipboard:= "Kawade Shobō"
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
	if clipboard= kin*yūzaiseijijōkenkyūkai
		clipboard:= "Kin'yū Zaisei Jijō Kenkyūkai"
	if clipboard= kitaōjishobo
		clipboard:= "Kitaōji Shobō"
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
	if clipboard= Minerubashobō
		clipboard:= "Mineurba Shobō"
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
	if clipboard= nihonhyōronsha
		clipboard:= "Nihon Hyōronsha"
	if clipboard= nihonkindaibungakukan
		clipboard:= "Nihon Kindai Bungagkukan"
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
		clipboard:= "San'ichi Shobō"
	if clipboard= san*ninsha
		clipboard:= "San'ninsha"
	if clipboard= sanninsha
		clipboard:= "San'ninsha"
	if clipboard= seirinshoin
		clipboard:= "Seirin Shoin"
	if clipboard= sekaishisōsha
		clipboard:= "Sekai Shisōsha"
	if clipboard= serikashobō
		clipboard:= "Serika Shobō"
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
		clipboard:= "Yumani Shobō"
	if clipboard= yūshindōkōbunsha
		clipboard:= "Yūshindō Kōbunsha"
}

;==============================
>^left::
suspend, toggle
exit

;==============================
\::
reload
exit
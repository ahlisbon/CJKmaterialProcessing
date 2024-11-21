;Created by Adam H. Lisbon
;Associate Professor - Japanese and Korean Studies Librarian
;University Libraries
;University of Colorado Boulder
;adam.lisbon@colorado.edu

#Requires AutoHotkey v2.0
setTitleMatchMode 2

#Include "%a_scriptdir%\Functions.ahk"
#Include "%A_ScriptDir%\Diacritics And Nengo.ahk"
#Include "%A_ScriptDir%\Fix Japanese Publisher Names.ahk"
#Include "%A_ScriptDir%\CJKmP - Find and Replace Data.ahk"

;■■■■■■■■■■■■■ Global Variables
	global bibArr:= []
	global active
	global activeSearch:= 0
	global tutorialMode:= 0

;■■■■■■■■■■■■■ Read values in .ini file
	;Populate variables from INI file
		CD:=			IniRead("CJKmP - Data to Spreadsheet.ini", "Sheet Names", "CD") ;CD = Collection Development
		DI:=			IniRead("CJKmP - Data to Spreadsheet.ini", "Sheet Names", "DI") ;DI = Donation Intake
		US:=			IniRead("CJKmP - Data to Spreadsheet.ini", "Sheet Names", "US") ;US = User Selects
		fsURL:=			IniRead("CJKmP - Data to Spreadsheet.ini", "Settings", "fsURL")
		checkMode:=		IniRead("CJKmP - Data to Spreadsheet.ini", "Settings", "checkMode")



;■■■■■■■■■■■■■ Run GUI
	; GUI Interface
		bib:= Gui(, "Bibliographic Data To Spreadsheet - Options")
	;Question 1:
		bib.Add("Text",		"					x180 y20",	"▼ File Name Prefixes of Your Spreadsheets (Case Sensitive)")
		bib.Add("Link",		"					x192 y40",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----file-name-prefixes`">Read about file naming conventions</a>")
		;Q1 answer 1
			bib.Add("Text",	"					x10  y65",	"Collection Development:")
			bib.Add("Edit",	"vCD		w300	x180 y60",	CD)
		;Q1 answer 2
			bib.Add("Text",	"					x10  y95",	"Donation Intake:")
			bib.Add("Edit",	"vDI		w300	x180 y90",	DI)
		;Q1 answer 3
			bib.Add("Text",	"					x10  y125",	"Users Select Materials:")
			bib.Add("Edit",	"vUS		w300	x180 y120",	US)
	;Question 3:
		bib.Add("Text",		"					x10	 y170",	"FirstSearch URL for your institution:")
		bib.Add("Edit",		"vfsURL		w300	x180 y165",	fsURL)
	;Question 4:
		bib.Add("Text", "						x10	 y200", "Use &Check Mode:")
		bib.Add("Checkbox",	"vcheckMode			x180 y200")
		bib.Add("Text", "						x220 y200", "Review data before it's imported into your spreadsheet.")
	;Question 5:
		bib.Add("Text", "						x10	 y230", "Use &Tutorial Mode:")
		bib.Add("Checkbox",	"vtutorialMode		x180 y230")
		bib.Add("Text", "						x220 y230", "Get pop up instructions on using hotkeys.")
	;Question 6:
		bib.Add("Text",		"					x10  y260", "&Wait longer for websites to load:")
		bib.Add("DDL", 		"vloadTime	w30		x180 y255 Choose1", ["1", "2", "3"])
		bib.Add("Link",		"					x220 y260",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----wait-longer-for-websites-to-load`">What is this?</a>")
	;Process answers into variables
		bib.Add("Button",	"default			x180 y300", "&Save Settings").OnEvent("Click", settings)
	;Help Text/Link
		bib.Add("Link", 	"					x10  y310", "Read the <a href=`"https://github.com/ahlisbon/CJKmaterialProcessing#-hotkeys-to-activate-macro`">Hotkey Guide</a> on GitHub")
		bib.Show()
	;Save and error check inputs
		settings(*){
			;Post that settings are updated
				bib.Add("Text", 	"			x264 y305", "✔ Updated")
				bib.Show()
			;Set variables
				saved:= bib.Submit(0)
				;Check Mode
					global checkMode
					checkMode:= saved.checkMode
				;Tutorial Mode
					global tutorialMode
					tutorialMode:= saved.tutorialMode
				;Load Time
					global loadTime
					loadTime:= 3000*saved.loadtime
					global lt
					lt:= loadTime
			;Check inputs
				checkGUIinputs(saved.CD, saved.DI, saved.US, saved.fsURL)
			;Save values in .ini file
				IniWrite(saved.FSurl,	"CJKmP - Data to Spreadsheet.ini", "Settings", "FSurl")
				IniWrite(saved.CD,		"CJKmP - Data to Spreadsheet.ini", "Sheet Names", "CD")
				IniWrite(saved.DI,		"CJKmP - Data to Spreadsheet.ini", "Sheet Names", "DI")
				IniWrite(saved.US,		"CJKmP - Data to Spreadsheet.ini", "Sheet Names", "US")
			;Tutorial window
				if(tutorialMode= 1){
					tutorialGUI()
					tutorialContent(1,	"",
										"Your Spreadsheet",
										"Numpad Plus   OR   F1",
										"Select the row of the current active cell to look up your item on FirstSearch.")
					tutorialShow()
				}
}
			checkGUIinputs(CD, DI, US, fsURL){
				;Blank value alert
					if(CD= "") & (DI= "") & (US= ""){
						MsgBox("At least one type of spreadsheet must have a name.", stopped, 4096)
						exit
						}
					if(fsURL= ""){
						MsgBox("The field for your institution's FirstSearch URL cannot be blank.", stopped, 4096)
						exit
					}
				;Duplicate value alert
					if((CD!="") & (DI="") & (US="")) | ((CD="") & (DI!="") & (US="")) | ((CD="") & (DI="") & (US!=""))
						return
					if(CD= DI) | (CD= US) | (DI= US){
						MsgBox("You cannot have two file names that are identical.", stopped, 4096)
						exit
					}
}
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼ Error checking to stop script if criteria aren't met. Declares what spreadsheet is actively being used.
	sheetCheck(){
			spreadsheet:= ""
			if(CD= "") & (DI= "") & (US= ""){
				active.Destroy()
				MsgBox("There is no spreadsheet open with a matching prefix.`n`nCheck if:`n`n  1) your file name matches the names you submitted`n  2) The file is open.", stopped, 4096)
			}
			if((WinExist(CD) & WinExist(DI)) | (WinExist(CD) & WinExist(US)) | (WinExist(DI) & WinExist(US))){
				active.Destroy()
				MsgBox("You have at least two different types of spreadsheets open. Please close all other spreadsheets except for the one you are actively working on.", stopped, 4096)
				exit
			}
			if WinExist(CD)
				spreadsheet:= CD
			if WinExist(DI)
				spreadsheet:= DI
			if WinExist(US)
				spreadsheet:= US
			return spreadsheet
}
;▼▲▼ 
	pasteToBibSpreadsheet(){
				spreadsheet:= sheetCheck()
				dataHere(spreadsheet)
;				Send "{esc}"
				Sleep nt
				Send "{home}"
				Sleep nt
			if InStr(spreadsheet, US){
				Send "{right 3}"
				Sleep nt
			}
			Send "^v"
			Sleep nt
			Send "{down}"
			Sleep wt
}



;■■■■■■■■■■■■■ Search FirstSearch with data from spreadsheet.
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼
	getDataFromBibSpreadsheet(){
			spreadsheet:= sheetCheck()
		;Get data from spreadsheet
			dataHere(spreadsheet)
			Send "{esc}"
			Sleep nt
			Send "{home}"
			Sleep nt
			Send "+{space}"
			Sleep nt
			copy()
			global bibArr:= strSplit(A_Clipboard, A_Tab)
			for index, value in bibArr
				bibArr[index]:= StrReplace(value, "`r`n", "")
			if(bibArr[18]= "") & (bibArr[19]= "") & (bibArr[20]= "") & (bibArr[23]= "") & (bibArr[24]= ""){
				active.Destroy()
				MsgBox("There is no searchable data in your spreadsheet.", stopped, 4096)
				exit
			}
}
;▼▲▼
	searchFS(searchParameter){
			spreadsheet:= sheetCheck()
			activateBrowser()
			newWin(fsURL)
			Sleep lt
		;▼ Search priority= oclc, isbn13, isbn10, native tile, romanized title
			searchText:= searchWith(bibArr[20], bibArr[18], bibArr[19], bibArr[24], bibArr[23])
			inClip(searchText)
			
			
			Send "^v"
			Sleep nt
		;▼ Use year of publication in FirstSearch search	
			if(searchParameter= 1) & (bibArr[25]!= "n/a") & (bibArr[25]!= ""){
					findTextOnSite("Limit to:")
					Send "{tab}"
					Sleep wt
					Send bibArr[25]
					Sleep wt
			}
			Send "{enter}"
			Sleep lt/2
		;▼ Get count for how many other libraries in FirstSearch that own item
			isTi:= inStr(searchText, "ti:", , 1)
			if(isTI= 0) & ((bibArr[17]= "") | (bibArr[17]= "n/a")){
				;▼ Get count from FirstSearch list of results
					if winActive("FirstSearch: WorldCat List of Records"){
						;▼ Isolate numbers for "Libraries Worldwide:" that own item, first 10 results only.	
							libCount:= copyAll()
							Sleep wt
							libCount:= RegExReplace(libCount, "`r|`n|\t")
							libCount:= RegExReplace(libCount, ".*^.+?Libraries Worldwide: ")
							libCount:= RegExReplace(libCount, "Libraries Worldwide: ", "`n")
							libCount:= RegExReplace(libCount, " .*|<.*")
						;▼ Sum of libraries holding item
							libCount := StrSplit(libCount, "`n")
							libTotal:= 0
							for index, line in libCount
								libTotal += line  ; Add the number on this line to the total
							bibArr[17]:= libTotal
							Send "{tab}"
					}
			}
}
;▼▲▼
			searchWith(oclc, isbn13, isbn10, titleR, titleN){		
					if((oclc!= "") & (oclc!= "n/a"))
						return "no:" . oclc
					if((isbn13!= "") & (isbn13!= "n/a"))
						return "bn:" . isbn13
				;▼ Allows ISBN10 column to be used to serach titles or ISBNs. Will not work if a book title is only numbers.
					if((isbn10!= "") & (isbn10!= "n/a")){
						isText:= RegExMatch(isbn10, "\d{9}")
						if(isText= 1)
							return "bn:" . isbn10
						else
							return "ti:" . isbn10
					}
					if((titleN!= "") & (titleN!= "n/a"))
						return "ti:" . titleN
					if((titleR!= "") & (titleR!= "n/a"))
						return "ti:" . titleR
}
;▼▲▼ Tutorial GUI
	FSresultsTutorial(){
		Sleep wt*2.5
		if(tutorialMode= 1){
			if WinExist("List"){
				tutorialGUI()
				tutorialContent(1,	"",
									"FirstSearch: WorldCat List of Records",
									"Numpad Plus   OR   F1",
									"open each record from the search results in a new tab.`n● Occaisionally, an individual record will not open automatically, but you can open it manually.")
				tutorialContent(0,	"=== OR =========",
									"",
									"Instead of using a hotkey, you can click on any result.",
									"determine if you will import the data of that particular record.")
				tutorialContent(0,	"=== AND IF YOU CLICK OPEN A RECORD WITH DATA YOU WANT TO IMPORT =========",
									"",
									"Numpad Enter   OR   Enter",
									"import the data in detailed record into your spreadsheet.")
				tutorialShow()
			}
			if WinExist("Detailed"){
				tutorialGUI()
				tutorialContent(1, 	"",
									"FirstSearch: WorldCat Detailed Record",
									"Numpad Enter   OR   Enter",
									"Bring the data from this record back into the spreadsheet.")
				tutorialContent(0, 	"=== OR =========",
									"",
									"Numpad Plus   OR   F1",
									"activate the `" Search for versions with same title and author`" link. If there is no link then nothing will happen.")
				tutorialShow()
			}
		}
}
;■■■
numpadAdd::{
		confirmBrowserOpen()
		global activeSearch:= 1
		tutorialOff()
		activeGUI()
		getDataFromBibSpreadsheet()
		searchFS(searchParameter:= 0)
		active.Destroy()
		FSresultsTutorial()
}
F1::{
		confirmBrowserOpen()
		global activeSearch:= 1
		tutorialOff()
		activeGUI()
		getDataFromBibSpreadsheet()
		searchFS(searchParameter:= 0)
		active.Destroy()
		FSresultsTutorial()
}
;■■■ Include year in search
>^numpadAdd::{
		confirmBrowserOpen()
		global activeSearch:= 1
		tutorialOff()
		activeGUI()
		getDataFromBibSpreadsheet()
		searchFS(searchParameter:= 1)
		active.Destroy()
		FSresultsTutorial()
}
>^F1::{
		confirmBrowserOpen()
		global activeSearch:= 1
		tutorialOff()
		activeGUI()
		getDataFromBibSpreadsheet()
		searchFS(searchParameter:= 1)
		active.Destroy()
		FSresultsTutorial()
}



;■■■■■■■■■■■■■ Open other versions link in FirstSearch Detailed Record.
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼ Tutorial GUI
	showRecordsList(){
			siteText:= copyAll()
		;Open other versions link in FirstSearch Detailed Record.
			if InStr(siteText, "Search for versions with same title and author"){
				findTextOnSite("Search for versions with same title and author")
				Send "{enter}"
			}else
				Send "{tab}"	
}
;▼▲▼ Tutorial GUI
	openRecordsList(){
		;▼ Determine how many tabs to open.
			data:= copyAll()
			data:= RegExReplace(data, "`r|`n|`t")
			if inStr(data, "Records Found: "){
					results:= RegExReplace(data, ".*Records found: | .*")
					if(results > 10)
						results:= 10
				;▼ Loop through results and open tabs
					entry:= 1
					Loop results{
						findTextOnSite(entry . ".")
						Send "+{right}"
						Sleep nt
						tabTest:= copy()
					;▼ Loop to get past random strings with a number + period, e.g.: "1."
						Loop{
							if !InStr(tabtest, "`t"){
								Send "{tab}"
								findTextOnSite(entry . ".")
								Send "+{right}"
								Sleep nt
								tabTest:= copy()
								if InStr(tabTest, "`t")
									break
							}
							else
								break
						}
						Send "{tab}"
						Sleep nt
						Send "^{enter}"
						Sleep nt
						entry++
					}
			}
}



;■■■ Open browser tab for each search result.
#HotIf WinActive("WorldCat Detailed Record")
numpadAdd::{
		tutorialOff()
		activeGUI()
		showRecordsList()
		active.Destroy()
		FSresultsTutorial()
}
F1::{
		tutorialOff()
		activeGUI()
		showRecordsList()
		active.Destroy()
		FSresultsTutorial()
}




;■■■■■■■■■■■■■ Open browser tab for each search result.
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼ Tutorial GUI
	FSlistTutorial(){
		if(tutorialMode= 1){
			tutorialGUI()
			tutorialContent(1,	"Browser window with many FirstSearch records loaded.",
								"",
								"Numpad 0   OR   Ctrl+Tab",
								"quickly swap browser tabs and browse for a record with data you want to import to your spreadsheet.")
			tutorialContent(0,	"=== OR =========",
								"",
								"Numpad Minus   OR   Ctrl+W",
								"close browser tabs of records you are not interested in.")
			tutorialContent(0,	"=== ON A RECORD YOU WANT TO IMPORT DATA FROM =========",
								"",
								"Numpad Enter   OR   Enter",
								"import the data in detailed record into your spreadsheet.")
			tutorialShow()
		}
}
;■■■
#HotIf WinActive("WorldCat List of Records")
numpadAdd::{
		tutorialOff()
		activeGUI()
		openRecordsList()
		active.Destroy()
		FSlistTutorial()
}
F1::{
		tutorialOff()
		activeGUI()
		openRecordsList()
		active.Destroy()
		FSlistTutorial()
}
numpadEnter::{
activeGUI()
		findTextOnSite("1.")
		Send "{tab}"
		Sleep nt
		Send "{enter}"
		Sleep nt
		active.Destroy()
}


;■■■■■■■■■■■■■ When multiple ISBNs are pulled for a book, select which ISBN you want to keep/remove.
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼ GUI for fixing ISBN
	removeISBN(){
		spreadsheet:= sheetCheck()
	;GUI Interface
		fix:= Gui(, "Fix ISBN")
		fix.Add("Text", "w200", "What volume do you want to keep?")
		fix.Add("Edit", "vv")
		fix.Add("Button", "default", "Fix").OnEvent("Click", processISBN)
		fix.OnEvent("Close", processISBN)
		fix.Add("Text")
		fix.Add("Text",, "Type the volume number of the ISBN you want to keep.")
		fix.Add("Text",, "Type `"s`" to keep the ISBN for the set.")
		fix.Add("Text",, "Type `"ns`" to remove the ISBN for the set.")
		fix.Add("Text",, "Type `"h`" to keep the ISBN for hardback.")
		fix.Show()
;▼▲▼ Isolate desired ISBN / Remove undesired data
	processISBN(*){
			activeGUI()
			keep:= fix.Submit()
			keep:= keep.v
			fixIarr:= copyRowMakeArray()
			i13:= fixIarr[15]
			i10:= fixIarr[16]
		;▼ Keep an ISBN with a label like "paperback", "hardcover" etc. while removing the label.
			if(keep= "p"){
				i13:= RegExReplace(i13, " \(paperback\).*| \(pbk.\).*")
				i10:= RegExReplace(i10, " \(paperback\).*| \(pbk.\).*")
			}
			if(keep= "h"){
				i13:= RegExReplace(i13, " \(hardcover\).*")
				i10:= RegExReplace(i10, " \(hardcover\).*")
			}
			if(keep= "ns"){
				i13:= RegExReplace(i13, " \^ ............. \(set\).*")
				i10:= RegExReplace(i10, " \^ .......... \(set\).*")
			}
		;▼ Normalize text (remove space between v. and number)
			i13:= RegExReplace(i13, "\(v\. ", "(v.")
			i10:= RegExReplace(i10, "\(v\. ", "(v.")
		;▼ Isolate ISBN13 based on input
			if (fixIarr[15], "............. (v."){
					hold:= i13
					i13:= RegExReplace(i13, "i) \(v\." keep "\).*")
					i13:= RegExReplace(i13, ".* ")
					;▼ removes "(v.x)" text string to do proper RegExMatch
					i13:= RegExReplace(i13, "i)\(v\..\)| ")
					if(i13= "")
						fixIarr[15]:= hold
					if(i13!= ""){
						if(RegExMatch(hold, i13))
							fixIarr[15]:= i13
						else
							fixIarr[15]:= hold
						fixIarr[10]:= keep
					}
			}
		;▼ Isolate ISBN10 based on input
			if (fixIarr[16], ".......... (v."){
					hold:= i10
					i10:= RegExReplace(i10, "i) \(v\." keep "\).*")
					i10:= RegExReplace(i10, ".* ")
					;▼ removes "(v.x)" text string to do proper RegExMatch
					i10:= RegExReplace(i10, "i)\(v\..\)| ")
					if(i10= "")
						fixIarr[16]:= hold
					if(i10!= ""){
						if(RegExMatch(hold, i10))
							fixIarr[16]:= i10
						else
							fixIarr[16]:= hold
						fixIarr[10]:= keep
					}
			}
		;▼ Paste to spreadsheet
				inClip(fixIarr[1] . "`t" . fixIarr[2] . "`t" . fixIarr[3] . "`t" . fixIarr[4] . "`t" . fixIarr[5] . "`t" . fixIarr[6] . "`t" . fixIarr[7] . "`t" . fixIarr[8] . "`t" . fixIarr[9] . "`t" . fixIarr[10] . "`t" . fixIarr[11] . "`t" . fixIarr[12] . "`t" . fixIarr[13] . "`t" . fixIarr[14] . "`t" . fixIarr[15] . "`t" . fixIarr[16])
				pasteToBibSpreadsheet()
				active.Destroy()
	}
}
;■■■
^numpad8::{
		removeISBN()
}
^F8::{
		activeGUI()
		removeISBN()
		active.Destroy()
}



;■■■■■■■■■■■■■ When multiple ISBNs are pulled for a book, select which ISBN you want to keep/remove.
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
		splitISBNprep(){
				Send "{esc}"
				Sleep nt
				Send "{home}"
				Sleep nt
				Send "{right 17}"
				Sleep nt
}
		splitISBNdown(){
				isbn:= copy()
				isbn:= RegExReplace(isbn, " \^ ", "`r`n")
				isbn:= RegExReplace(isbn, ".*set.*")
				volumes:= isbn
				isbn:= RegExReplace(isbn, " \(v.*")
				isbn:= RegExReplace(isbn, " \(.*")
				isbn:= RegExReplace(isbn, "`r`n`r`n*")
				inClip(isbn)
				Send "^v"
				Sleep nt
				return volumes
}
		splitVolumesDown(volumes){
				Send "{left 9}"
				Sleep nt
				volumes:= RegExReplace(volumes, ".*\(v\.|.*\(")
				volumes:= RegExReplace(volumes, "\).*")
				volumes:= RegExReplace(volumes, "`r`n`r`n*")
				volumes:= RegExReplace(volumes, " ")
				inClip(volumes)
				Send "^v"
				Sleep nt
}
;■■■
^numpad9::{
		splitISBNprep()
		volumes:= splitISBNdown()
		Send "{right}"
		splitISBNdown()
		splitVolumesDown(volumes)
		
}
;■■■
F9::{
		splitISBNprep()
		volumes:= splitISBNdown()
		Send "{right}"
		splitISBNdown()
		splitVolumesDown(volumes)
}



;■■■■■■■■■■■■■ Fast ISBN Fix
#HotIf WinActive(CD) | WinActive(DI)
;Relative to the cursor in an active cell, keeps only the ISBN that the cell is in.
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼ Copies text in active cell.
	fastISBNfix(){
			Send "^+{right}"
			Sleep 25
			inCellCheck()
			rightP:= copy()
			rightP:= Trim(rightP)
			Send "^+{left}"
			Sleep 25
			;inCellCheck()
			leftP:= copy()
			leftP:= Trim(leftP)
			isbn:= leftP . rightP
			Send "{esc}"
			Sleep wt
		;▼ populate row
			arr:= copyRowMakeArray()
			len:= strLen(isbn)
			if(len= 13){
				arr[18]:= isbn
				arr[19]:= convert13to10(isbn)
			}
			if (len= 10)
				arr[19]:= isbn
			arr[8]:= getPriceURL(arr[19], arr[21])
			inClip(arr[1] . "`t" . arr[2] . "`t" . arr[3] . "`t" . arr[4] . "`t" . arr[5] . "`t" . arr[6] . "`t" . arr[7] . "`t" . arr[8] . "`t" . arr[9] . "`t" . arr[10] . "`t" . arr[11] . "`t" . arr[12] . "`t" . arr[13] . "`t" . arr[14] . "`t" . arr[15] . "`t" . arr[16] . "`t" . arr[17] . "`t" . arr[18] . "`t" . arr[19])
			pasteToBibSpreadsheet()
			
}
	;▼▲▼ Stops script if it was run outside of an active cell.
			inCellCheck(){
				copy()
				if inStr(A_Clipboard, "`t") | inStr(A_Clipboard, "`r`n"){
					Send "{home}"
					Sleep nt
					Send "{right 14}"
					Sleep nt
					Send "{esc}"
					Sleep nt
					active.Destroy()
					MsgBox("You tried to fix a cell in the ISBN13 or ISBN10 columns but weren't `"inside`" the cell. Double click a cell and have your cursor within the ISBN you want to keep in order to clean up your ISBNs.", stopped, 4096)
					exit
				}
			
}
;■■
numpadDiv::{
		activeGUI()
		fastISBNfix()
		active.Destroy()
}		
F12::{
		activeGUI()
		fastISBNfix()
		active.Destroy()
}



;■■■■■■■■■■■■■ Convert ISBN13 to ISBN10
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
	convertISBN(){
		spreadsheet:= sheetCheck()
		arr:= copyRowMakeArray()
		if(arr[18]= "n/a") | (arr[18]= ""){
			MsgBox("There is no ISBN13 to convert into an ISBN10.", stopped, 4096)
			return
		}
		if inStr(arr[18], " "){	
			MsgBox("The ISBN13 appears to not be formatted correctly.", stopped, 4096)
			return
		}
		isbn13:= arr[18] 
		arr[19]:= convert13to10(isbn13)
		if(arr[21]= "Japanese"){
			arr[8]:= "https://www.amazon.co.jp/dp/" . arr[19]
			}
		else
			arr[8]:= "https://www.amazon.com/dp/" . arr[19]
		inClip(arr[1] . "`t" . arr[2] . "`t" .  arr[3] . "`t" .  arr[4] . "`t" .  arr[5] . "`t" .  arr[6] . "`t" .  arr[7] . "`t" .  arr[8] . "`t" .  arr[9] . "`t" .  arr[10] . "`t" .  arr[11] . "`t" .  arr[12] . "`t" .  arr[13] . "`t" .  arr[14] . "`t" .  arr[15] . "`t" .  arr[16] . "`t" .  arr[17] . "`t" .  arr[18] . "`t" .  arr[19])	
		pasteToBibSpreadsheet()
}
;■■■
^numpad7::{
		activeGUI()
		convertISBN()
		active.Destroy()
}
F7::{
		activeGUI()
		convertISBN()
		active.Destroy()
}



;■■■■■■■■■■■■■ Translate title with ChatGPT.
;▼▲▼
translateWithChatGPT(){
	spreadsheet:= sheetCheck()
			global chatGPTmode
			global gptPending
			gptPending:= 0
			
			if !WinExist("Translate"){
				active.Destroy()
				msg:= "There is no browser window with an active tab open to Chat GPT.`n`n"
					. "1) Check you have ChatGPT open in your browser.`n"
					. "2) Make sure ChatGPT is your browser's the active tab.`n"
					. "3) You need to save and keep active a chat called `"Translate`" (case sensitive)."
				MsgBox(msg, stopped, 4096)
				return
			}
			
			
			Send "{esc}"
			Sleep nt
			copy()
			A_clipboard:= RegExReplace(A_clipboard, "`r`n$")
			if !InStr(A_Clipboard, "`r`n"){
					chatGPTmode:= "oneTitle"
				;▼ Copy native language and native title and verify enough information is available for translation.
					global gptArr
					gptArr:= copyRowMakeArray()
					if(gptArr[21]= "English"){
						active.Destroy()
						MsgBox("This item is already in English and doesn't need to be translated.", stopped, 4096)
						exit
					}
					if((gptArr[23]= "n/a") | (gptArr[23]="")) & ((gptArr[24]="n/a") | (gptArr[24]= "")){
						active.Destroy()
						MsgBox("There is no non-English title to translate.", stopped, 4096)
						exit
					}
					title:= copy()
					if(title= "n/a") | (title= ""){
						active.Destroy()
						MsgBox("There is no title to translate.`n`nReview Column U", stopped, 4096)
						exit
						}
				;▼ Determine which title data to use
					if(gptArr[24]= "n/a") | (gptArr[24]="")
						title:= gptArr[23]
					else
						title:= gptArr[24]
					A_clipboard:= ""
			}
			if InStr(A_Clipboard, "`r`n"){
					chatGPTmode:= "bulk"
					titles:= A_Clipboard
			}
		;▼ Go to ChatGPT window
			if WinExist("Translate")
					WinActivate
			Sleep wt
			findTextOnSite("ChatGPT can make mistakes")
			Send "+{tab}"
			Sleep nt
		;▼ Paste translation request
			if(chatGPTmode= "oneTitle"){
					prompt:= "Provide an English translation of this title:" . gptArr[24] . ". The response should only include the translated title and no other text. Do not put quotation marks around the title."
					inClip(prompt)
					Send "^v"
					Sleep nt
					Send "{enter}"
					Sleep nt
					gptPending:= 1
			}
			if(chatGPTmode= "bulk"){
					prompt:= "Provide English translations of these titles:`r`n" . titles . "`r`n`r`nThe response should only include the translated titles and no other text. Do not put quotation marks around the titles."
					inClip(prompt)
					Send "^v"
					Sleep nt
					Send "{enter}"
					Sleep nt
					gptPending:= 1
			}
}
;▼▲▼
sendTraslationBackToSpreadsheet(){
		spreadsheet:= sheetCheck()
		global gptPending
		global chatGPTmode
		if(gptPending= 0){
				active.Destroy()
				MsgBox("An error has occured while using ChatGPT to translate a title.`n`nRestart the script and try again.", stopped, 4096)
				exit
		}
	;▼ Single title translation process
		if(chatGPTmode= "oneTitle"){
			;▼ Copy response
				findTextOnSite("ChatGPT can make mistakes")
				Send "^a"
	msgbox "hello"
				Sleep nt
				title:= copyAll()
				Send "{tab}"
				Sleep nt
			;▼ Isolate translated title
				title:= RegExReplace(title, "`r`n")
				title:= RegExReplace(title, "ChatGPT can make mistakes.*")
				title:= RegExReplace(title, ".*ChatGPT")
				title:= title . " - ChatGPT translation"
			;▼ Paste translation request
				inClip(gptArr[1] . "`t" . gptArr[2] . "`t" . gptArr[3] . "`t" . gptArr[4] . "`t" . gptArr[5] . "`t" . gptArr[6] . "`t" . gptArr[7] . "`t" . gptArr[8] . "`t" . gptArr[9] . "`t" . gptArr[10] . "`t" . gptArr[11] . "`t" . gptArr[12] . "`t" . gptArr[13] . "`t" . gptArr[14] . "`t" . gptArr[15] . "`t" . gptArr[16] . "`t" . gptArr[17] . "`t" . gptArr[18] . "`t" . gptArr[19] . "`t" . gptArr[20] . "`t" . gptArr[20] . "`t" . gptArr[21] . "`t" . title)
				pasteToBibSpreadsheet()
		}
	;▼ Bulk title list transtlation process	
		if(chatGPTmode= "bulk"){
		msgbox "bulk"
			;▼ Copy Response
				Send "+{tab}"
				Sleep nt
				Send "^a"
				Sleep nt
				titles:= copyAll()
				Send "{tab}"
				Sleep nt
			;▼ Isolate translated titles
				titles:= RegExReplace(titles, "`r`n", "xtx")
				titles:= RegExReplace(titles, ".*xtxChatGPTxtxxtx|xtxChatGPT can make mistakes.*")
				titles:= RegExReplace(titles, "xtx", "xtx`r`n")
				titles:= RegExReplace(titles, "xtx", " - ChatGPT translation")
				titles:= RegExReplace(titles, "    ")
			;▼ Paste translation request
				inClip(titles)
				dataHere(spreadsheet)
				Send "{esc}"
				Sleep nt
				Send "{left 2}"
				Sleep nt
				Send "^v"
				Sleep nt
				Send "{esc}"
				gptPending:= 0
		}
		active.Destroy()
}
;■■■
#HotIf WinActive(CD) | WinActive(DI)
numpadSub::{
		activeGUI()
		translateWithChatGPT()
		active.Destroy()
	}
^-::{
		activeGUI()
		translateWithChatGPT()
		active.Destroy()
	}

#HotIf winActive("Translate")
numpadSub::{
		activeGUI()
		sendTraslationBackToSpreadsheet()
		active.Destroy()
	}
^-::{
		activeGUI()
		sendTraslationBackToSpreadsheet()
		active.Destroy()
	}


;■■■■■■■■■■■■■ Price Look up
;Applies hierarchy to seach with ISBN10 first, then ISBN13, then title in native script.
;▼▲▼
	priceSearch(isbn10, isbn13, titleN, urlPrefix:= "", urlSuffix:= "",AZisbn10prefix:= "", AZsearchPrefix:= "", isAmazon:= 0){
		global bibArr
	;Assemble search string
		;ISBN10
			if(isAmazon= 1) ;Special logic for amazon isbn10
				urlPrefix:= AZisbn10prefix
			if(isbn10!= "n/a") & (isbn10!= ""){
				if(isAmazon= 1)
					urlSuffix:= ""
				searchURL:= urlPrefix . bibArr[16] . urlSuffix
				return searchURL
			}
		;ISBN13
			if(isAmazon= 1) ;Special logic for non isbn10 search
				urlPrefix:= AZsearchPrefix
			if(isbn13!= "n/a") & (isbn13!= ""){
				searchURL:= urlPrefix . bibArr[15] . urlSuffix
				return searchURL
			}
		;Title Non-Romanized
			if(isAmazon= 1) ;Special logic for non title search
				urlPrefix:= AZsearchPrefix
			if(titleN!= "n/a") & (titleN!="n/a"){
				searchURL:= urlPrefix . bibArr[21] . urlSuffix
				return searchURL
			}
}
;▼▲▼
	searchPrice(){
		spreadsheet:= sheetCheck()
		getDataFromBibSpreadsheet()
		activateBrowser()
	;Japanese
		if(bibArr[18]= "Japanese") | (bibArr[34]= "jpy"){
			;amazonJP
				searchURL:= priceSearch(bibArr[16], bibArr[15], bibArr[21],,"&i=stripbooks","https://www.amazon.co.jp/dp/","https://www.amazon.co.jp/s?k=",1)
				newWin(searchURL)
			;amazon.com
				searchURL:= priceSearch(bibArr[16], bibArr[15], bibArr[21],,"&i=stripbooks","https://www.amazon.com/dp/","https://www.amazon.com/s?k=",1)
				newTab(searchURL)
			;kosho.or.jp / Nihon no Furuhonya
				searchURL:= priceSearch("n/a", "n/a", bibArr[21],"https://www.kosho.or.jp/products/list.php?&mode=search&search_only_has_stock=1&search_word=")
				newTab(searchURL)
			;JPT / jptbooknews.jpt.co.jp
				searchURL:= priceSearch("n/a", bibArr[15], bibArr[21],"https://jptbooknews.jptco.co.jp/product?q=")
				newTab(searchURL)
		}
		else{
			active.Destroy()
			MsgBox("At this time, price checking options are only available for Japanese materials.", stopped, 4096)
			exit
		}
}
;■■■
#HotIf WinActive(CD) | WinActive(DI)
^numpadAdd::{
		activeGUI()
		searchPrice()
		active.Destroy()
}
F4::{
		activeGUI()
		searchPrice()
		active.Destroy()
}



;■■■■■■■■■■■■■ Return price to spreadsheet
#HotIf WinActive("ahk_exe firefox.exe") | WinActive("ahk_exe msedge.exe") | WinActive("ahk_exe chrome.exe")
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
	priceToBibSpreadsheet(price, currency){
			Send "!d"
			Sleep nt
			url:= copyAll()
			Send "!{F4}"
			Sleep nt
			bibArr[8]:= url
			bibArr[34]:= currency
			bibArr[35]:= price
			inClip(bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28] . "`t" . bibArr[29] . "`t" . bibArr[30] . "`t" . bibArr[31] . "`t" . bibArr[32] . "`t" . bibArr[33] . "`t" . bibArr[34] . "`t" . bibArr[35])
			pasteToBibSpreadsheet()
			active.Destroy()
			exit
}
;▼▲▼ Furuhonya - kosho.or.jp
	furuhonya(){
			spreadsheet:= sheetCheck()
			price:= copyAll()
			if !InStr(price, "￥"){
				active.Destroy()
				MsgBox("There is no price to add from www.kohso.or.jp", stopped, 4096)
				exit
			}
			price:= RegExReplace(price, "`r|`n|`t", "xtx")
			if inStr(price, "の検索結果"){
				active.Destroy()
				MsgBox("You are on a list of search results for www.kosho.or.jp.`n`nPlease load one of the records from the results to import the price into the spreadsheet.", stopped, 4096)
				exit
			}
			price:= RegExReplace(price, "^.+?￥|xtx.*")
			price:= RegExReplace(price, ",")
			priceToBibSpreadsheet(price, "jpy")
}
;▼▲▼
	amazonUS(){
			spreadsheet:= sheetCheck()
			A_Clipboard:= ""
			Sleep nt
			Send "^c"
			Sleep nt
			if !ClipWait(2){
				active.Destroy()
				MsgBox("Nothing on Amazon.com (US site) has been highlighted to copy.`n`nWhen highlighting a price to copy make sure all the numbers and dollar sign ($) are hihglighted to import the price into the spreadsheet.", stopped, 4096)
				exit
			}
			price:= A_Clipboard
			if !inStr(price, "$"){
				active.Destroy()
				MsgBox("You may have tried to highlight a price but there was no dollar sign ($) in the text you highlighted.", stopped, 4096)
				exit
			}
			price:= RegExReplace(price, "\$")
			priceToBibSpreadsheet(price, "usd")
}

;▼▲▼
	amazonJP(){
			spreadsheet:= sheetCheck()
			A_Clipboard:= ""
			Sleep nt
			Send "^c"
			Sleep nt
			if !ClipWait(2){
				active.Destroy()
				MsgBox("Nothing on Amazon.co.jp (Japan site) has been highlighted to copy.`n`nWhen highlighting a price to copy make sure all the numbers and yen sign (¥) are hihglighted to import the price into the spreadsheet.", stopped, 4096)
				exit
			}
			price:= A_Clipboard
			if !inStr(price, "￥"){
				active.Destroy()
				MsgBox("You may have tried to highlight a price but there was no yen sign (¥) in the text you highlighted.", stopped, 4096)
				exit
			}
			price:= RegExReplace(price, "￥")
			priceToBibSpreadsheet(price, "jpy")
}
;▼▲▼ Japan Publications Trading - https://jptbooknews.jptco.co.jp
	JPT(){
			spreadsheet:= sheetCheck()
			price:= copyall()
			if !inStr(price, "Japan Publications Trading"){
				active.Destroy()
				Send "{tab}"
				exit
			}
			price:= RegExReplace(price, "`r|`n|`t")
			price:= RegExReplace(price, ".*¥\) ：|円.*")
			priceToBibSpreadsheet(price, "jpy")
}
;▼▲▼ Get Price
getPrice(){
	;Sites with unique titles
		if winActive("日本の古本屋")
			furuhonya()
		if winActive("Amazon.com:")
			amazonUS()
		if winActive("Amazon.co.jp:") | WinActive(" | Amazon")
			amazonJP()
	;Sites without unique titles
		JPT()
}
;■■■
^+numpadEnter::{
		activeGUI()
		getPrice()
		active.Destroy()
}
^+enter::{
		activeGUI()
		getPrice()
		active.Destroy()
}



;■■■■■■■■■■■■■ Conveniences
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
	moveToISBN(){
		Send "{esc}"
		Sleep nt
		Send "{home}"
		Sleep nt
		Send "{right 18}"
}
;■■■ Move active cell to ISBN column in a row
#HotIf WinActive(CD) | WinActive(DI)
^numpadEnter::moveToISBN()
^enter::moveToISBN()

;■■■ Speed up browsing through tabs to look at First Search results
#HotIf WinActive("WorldCat Detailed Record") | WinActive("WorldCat List of Records") | WinActive("")
		numpad0::Send "^{tab}"	;go to next tab
		numpadSub::Send "^w"	;close tab

;■■■ Speed up browsing through tabs to look at First Search results
#HotIf WinActive("ahk_exe firefox.exe") | WinActive("ahk_exe msedge.exe") | WinActive("ahk_exe chrome.exe")
		^numpad0::Send "^{tab}"	;go to next tab
#Hotif



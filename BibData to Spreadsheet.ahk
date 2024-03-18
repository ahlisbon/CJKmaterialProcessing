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



;■■■■■■■■■■■■■ Global Variables
	global bibArr:= []
	global active
	global activeSearch:= 0
	global tutorialMode:= 0

;■■■■■■■■■■■■■ Read values in .ini file
	;Populate variables from INI file
		CD:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "CD") ;CD = Collection Development
		DI:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "DI") ;DI = Donation Intake
		US:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "US") ;US = User Selects
		fsURL:=			IniRead("BibData to Spreadsheet.ini", "Settings", "fsURL")
		checkMode:=		IniRead("BibData to Spreadsheet.ini", "Settings", "checkMode")



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
				IniWrite(saved.FSurl,	"Bibdata to Spreadsheet.ini", "Settings", "FSurl")
				IniWrite(saved.CD,		"Bibdata to Spreadsheet.ini", "Sheet Names", "CD")
				IniWrite(saved.DI,		"Bibdata to Spreadsheet.ini", "Sheet Names", "DI")
				IniWrite(saved.US,		"Bibdata to Spreadsheet.ini", "Sheet Names", "US")
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
				Send "{esc}"
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
			Send searchText
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
			if(bibArr[17]= "") | (bibArr[17]= "n/a"){
				;▼ Get count from FirstSearch list of results
					if winActive("FirstSearch: WorldCat List of Records"){
					;▼ Isolate numbers for "Libraries Worldwide:" that own item, first 10 results only.	
						libCount:= copyAll()
						Sleep wt
						libCount:= RegExReplace(libCount, "`r|`n|`t")
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
					if (oclc!= "") & (oclc!= "n/a")
						return "no:" . oclc
					if (isbn13!= "") & (isbn13!= "n/a")
						return "bn:" . isbn13
					if (isbn10!= "") & (isbn10!= "n/a")
						return "bn:" . isbn10
					if (titleN!= "") & (titleN!= "n/a")
						return "ti:" . titleN
					if (titleR!= "") & (titleR!= "n/a")
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
		;Determine how many tabs to open.
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



;■■■■■■■■■■■■■ Open browser tab for each search result.
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



;■■■■■■■■■■■■■ Put bibliographic data from a FirstSearch detailed record and put it in a spreadsheet.
#HotIf WinActive("Detailed Record") | WinActive("Libraries that Own Item")
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼
	pullBibData(){
			global bibArr
			global activeSearch
		;▼ Error check that spreadsheet and fsURL variables are not blank.
			checkGUIinputs(CD, DI, US, fsURL)
			spreadsheet:= sheetCheck()
		;▼ FirstSearch
			if WinExist("Detailed Record"){
				WinActivate
				data:= loadSourceCode("pagename", "FirstSearch Detailed Record")
				data:= RegExReplace(data, "`r`n")							;Makes parsing easier.
				data:= RegExReplace(data, "<span class\=matchterm[0-9]>") 	;Removes code for yellow text highlighting.
				data:= RegExReplace(data, "</span>") 						;Because of highlighting above, all "</span>" are removed.
				data:= normalize(data)	
			;📚 Total volumes of multivolume sets
					volumes:= getVolumes(data)
			;Determine if active search is happening
				if(activeSearch= 0){
						bibArr:= []
						Loop 35{
							bibArr.InsertAt(1, "")
							}
						bibArr[1]:= ""
						bibArr[9]:= ""
				}
				;🏛 Check against local holdings
							if InStr(data, "FirstSearch indicates your institution owns the item.")
								FSdupe:= "y"
							else
								FSdupe:= "n"	
				;👬 Get count of libraries that also have this item
						if (bibArr[17]= "") | (bibArr[17]= "n/a"){
							;▼ Isolate number for "Libraries Worldwide that own item"
								bibArr[17]:= RegExReplace(data, ".*Libraries worldwide that own item.+? |&.*|<.*")
						}
								
				;🔢 ISBN
							if InStr(data, "<b>ISBN:"){
					;ISBN-13
							isbn13:= getISBN(data, "\[..........\]|\[.......... .+?\]")
					;ISBN-10
							isbn10:= getISBN(data, "\[.............\]|\[............. .+?\]")
							}else{
								isbn13:= "n/a"
								isbn10:= "n/a"
							}
							if inStr(isbn13, "979")
								isbn10:= "n/a"
				;🔢 OCLC#
							oclc:= RegExReplace(data, ".*<b>OCLC:</b> |<.*")
				;🗣 Language
							global language
							language:= RegExReplace(data, ".*<b>Language:.+?serif`">|&.*|<.*")
				;📔 Title
						;Romanized
							titleR:= getTitleR(data)
						;Native
							titleN:= getTitleN(data)
						;Translated
							titleT:= getTitleT(data)
				;✒ Creators
						;Creator
								if InStr(data, "<b>Author(s):"){
									;Romanized
										creatorR:= getCreatorR(data)
									;Native
										creatorN:= getCreatorN(data)
								}else{
									creatorR:= "n/a"
									creatorN:= "n/a"
								}
						;Corporate Creator
								if InStr(data, "<b>Corp Author(s):"){
									;Romanized
										corpR:= getCorpR(data)
									;Native
										corpN:= getCorpN(data)
								}else{
									corpR:= "n/a"
									corpN:= "n/a"
								}
						;Merge individual and corporate creator values.
								;Romanized
									creatorsR:= creatorR . " ^ " . corpR
									creatorsR:= RegExReplace(creatorsR, " \^ n/a|n/a \^ ")
								;Native
									creatorsN:= creatorN . " ^ " . corpN
									creatorsN:= RegExReplace(creatorsN, " \^ n/a|n/a \^ ")
				;📚 Series	
						if InStr(data, "<b>Series:"){
							;Romanized
								seriesR:= getSeriesR(data)
							;Native
								seriesN:= getSeriesN(data)
						}else{
							seriesR:= "n/a"
							seriesN:= "n/a"
						}
				;🏢 Publisher
						if InStr(data, "<b>Publication:"){
							;Romanized
								publisherR:= getPublisherR(data)
							;Native
								publisherN:= getPublisherN(data)
						}else{
							publisherR:= "n/a"
							publisherN:= "n/a"
						}
				;♎ Year of Publication
						;Romanized
							yearR:= getYearR(data)
						;Native
							if(language= "Japanese")
								yearN:= convertNengo(yearR) ;Function is in "diacriticsNengo.ahk"
							else
								yearN:= "n/a"
				;📖 Edition
						if InStr(data, "<b>Edition:"){
							;Romanized
								editionR:= getEditionR(data)
							;Native
								editionN:= getEditionN(data)
						}else{
							editionR:= "n/a"
							editionN:= "n/a"
						}	
				;💡 Subjects
						if InStr(data, "SUBJECT(S)")
							subjects:= getSubjects(data)
						else
							subjects:= "n/a"
			}
		;▼ Check Results
				checkData:=   "ISBN-13#:`n"
							. isbn13				. "`n`n"
						. "ISBN-10#:`n"
							. isbn10				. "`n`n"
						. "OCLC#:`n"
							. oclc					. "`n`n"
						. "Language:`n"
							. language				. "`n`n"
						. "Title:`n"
							. titleR				. "`n"
							. titleN				. "`n`n"
						. "Creator(s):`n"
							. creatorR				. "`n"
							. creatorN				. "`n`n"
						. "Corporate Creator(s):`n"
							. corpR					. "`n"
							. corpN					. "`n`n"
						. "Series Title:`n"
							. seriesR				. "`n"
							. seriesN				. "`n`n"
						. "Publisher:`n"
							. publisherR			. "`n"
							. publisherN			. "`n`n"
						. "Year: `n"
							. yearR					. "`n"
							. yearN					. "`n`n"
						. "Edition: `n"
							. editionR				. "`n"
							. editionN				. "`n`n"
						. "Subject(s):`n"
							. subjects
			if(checkMode= 1){
				yesNo:= MsgBox(checkData, "Review Bibliographic Data", 4100)
				if(yesNo= "No")
					exit
			}
			subjects:= regExReplace(subjects, "`n", " ^ ") ;Each subject is on a serpate line for easy reading in the :Check Data message box." This puts them on one line to be pasted into a single cell on the spreadsheet.
			subjects:= RegExReplace(subjects, " \^  \^ $| \^ $")
		;▼ Calculate other array fields based on pulled data.
				;▼ Series Number
						if InStr(seriesR, " `; ")
							vol:= RegExReplace(seriesR, ".* `; ")
						else if InStr(seriesN, " `; ")
							vol:= RegExReplace(seriesN, ".* `; ")
						else
							vol:= "n/a"
				;▼ Generate URL for Purchase
						if(language="Japanese") & (isbn10!= "") & (isbn10!= "n/a")
								priceURL:= "https://www.amazon.co.jp/dp/" . isbn10
						else
							priceURL:= "n/a"
						if inStr(isbn10, A_space)
							priceURL:= "n/a"
		
		;▼ Put parsed data into array.
				bibArr[8]:= priceURL
				bibArr[10]:= vol
				bibArr[13]:= volumes
				bibArr[16]:= FSdupe
				bibArr[18]:= isbn13
				bibArr[19]:= isbn10
				bibArr[20]:= oclc
				bibArr[21]:= language
				bibArr[22]:= titleT
				bibArr[23]:= titleR
				bibArr[24]:= titleN
				bibArr[25]:= yearR
				bibArr[26]:= yearN
				bibArr[27]:= creatorsR
				bibArr[28]:= creatorsN
				bibArr[29]:= seriesR
				bibArr[30]:= seriesN
				bibArr[31]:= publisherR
				bibArr[32]:= publisherN
				bibArr[33]:= editionR
				bibArr[34]:= editionN
				bibArr[35]:= subjects
				inClip(bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28] . "`t" . bibArr[29] . "`t" . bibArr[30] . "`t" . bibArr[31] . "`t" . bibArr[32] . "`t" . bibArr[33] . "`t" . bibArr[34] . "`t" . bibArr[35])
				
		;▼ Close FirstSearch "Detailed Record" page.
				if WinExist("Detailed Record") 
					WinActivate
				sleep wt
				if(activeSearch= 1){
					Send "!{F4}"
					sleep wt
				}
		;▼ Populate spreadsheet with data.
				if(spreadsheet= US){
					Loop 7{
						bibArr.RemoveAt(1)
					}
					inClip(bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28])
				}
				activeSearch:= 0
				pasteToBibSpreadsheet()
		;▼ Clear Variables
				
}
		;▼▲▼
			normalize(data){
				; Convert Japanese punctuation to English punctuation.
				data:= RegExReplace(data, "  ", " ")
				data:= RegExReplace(data, "：", ":")
				data:= RegExReplace(data, "、", ".")
				loop{
						if InStr(data, "  ")
							data:= RegExReplace(data, "  "," ")
						else
							Break
}
		;▼▲▼
				; Turn composed characters into decomposed characters.
					data:= regExReplace(data, "ā", "ā")
					data:= regExReplace(data, "ī", "ī")
					data:= regExReplace(data, "ū", "ū")
					data:= regExReplace(data, "ē", "ē")
					data:= regExReplace(data, "ō", "ō")
				return data
}
		;▼▲▼
			getVolumes(data){
					volumes:= RegExReplace(data, ".*<b>Description:|</font></td>.*")
					if inStr(volumes, " volumes")
						volumes:= RegExReplace(volumes, ".*serif`">| volumes.*")
					else
						volumes:= "n/a"
					volumes:= Trim(volumes)
					return volumes
}
		;▼▲▼
			getISBN(data, needle){
				;Contain beginning and end of ISBN data
					isbn:= RegExReplace(data, ".*<b>ISBN:</b> ", "[")
					isbn:= RegExReplace(isbn, "</font>.*|; <.*| <.*", "]")
				;Fix double brackes: (( and ))
					isbn:= RegExReplace(isbn, "\(\(", "(")
					isbn:= RegExReplace(isbn, "\)\)", ")")
				;▲
					isbn:= RegExReplace(isbn, "; ", "][")
					isbn:= RegExReplace(isbn, needle)
					isbn:= regExReplace(isbn, "\]\[", " ^ ")
					isbn:= regExReplace(isbn, "\]|\[")
					isbn:= regExReplace(isbn, "i)dai |-kan")
					isbn:= Trim(isbn)
					return isbn
}
		;▼▲▼
			getTitleR(data){
					titleR:= RegExReplace(data, ".*<b>Title:.+?</td>|</td>.*")
					titleR:= RegExReplace(titleR, ".* /<div><br>|.* /</div>| =<br>.*")		;Removes translated title in order to parse romanized title.
					if InStr(titleR, "Translated Title:"){									;▲
						titleR:= RegExReplace(titleR, " /<br><b>Translated Title:.*")		;▲
						titleR:= RegExReplace(titleR, ".*<br>") 							;▲
						titleR:= Trim(titleR)
						return titleR
					}
					titleR:= RegExReplace(titleR, "( :| `;|:|;)<br>", ": ")	;Subtitle
					titleR:= RegExReplace(titleR, ".*<br>| /.*")  	;▲
				;Clean Up
					titleR:= RegExReplace(titleR, ".</b>.*")
					titleR:= RegExReplace(titleR, " : ", ": ")
					titleR:= RegExReplace(titleR, ".*</div>")
					titleR:= RegExReplace(titleR, "&lt;&gt;", "n/a")	
					titleR:= Trim(titleR)
					return titleR
}
		;▼▲▼
			getTitleN(data){
					titleN:= RegExReplace(data, ".*<b>Title:.+?</td>|</td>.*")
					if !InStr(titleN, "vernacular")
						titleN:= "n/a"
					if InStr(titleN, " = ")
						titleN:= RegExReplace(titleN, " = .*")
					titleN:= RegExReplace(titleN, ".*lang=`"..`">| /.*")
					titleN:= RegExReplace(titleN, ".*lang=`"..`">|( |　|\.)/.*|(\.|,)</div>.*")
				;Clean Up
					titleN:= RegExReplace(titleN, " : ", ": ")
					titleN:= Trim(titleN)
					return titleN
}
		;▼▲▼
			getTitleT(data){
					titleT:= RegExReplace(data, ".*<b>Title:.+?</td>|</td>.*")
					if InStr(titleT, "Translated Title:"){
						titleT:= RegExReplace(titleT, ".*Translated Title:</b> |<.*|\. eng.*")
						titleT:= Trim(titleT)
						return titleT
					}
					
					if !InStr(titleT, " = ")
						titleT:= "n/a"
					titleT:= RegExReplace(titleT, ".* = | /.*")
					titleT:= Trim(titleT)
					return titleT
}
		;▼▲▼
			deDupe(txt){
					str:= ""
					loop Parse, txt, "`r"
						if !InStr(str, A_LoopField "`n")
							str .= A_LoopField "`n"
					return str
}
		;▼▲▼
			cleanCreator(creator){
					creator:= RegExReplace(creator, ", `;|\. `;|,`;", ",")
					creator:= RegExReplace(creator, "(,|;) [0-9].*|(,|-|\.) editor.*|, author.*")
					creator:= RegExReplace(creator, "\.$")
					loop{
						if InStr(creator, "`n`n")
							creator:= RegExReplace(creator, "`n`n","`n")
						else
							Break
					}
					creator:= deDupe(creator)
					creator:= RegExReplace(creator, "`n", " ^ ")
					creator:= RegExReplace(creator, "(\.|,) \^ ", " ^ ")
					creator:= RegExReplace(creator, " \^ ,| \^ $| \^$|,$| $")
					creator:= Trim(creator)
					return creator
}
		;▼▲▼
			getCreatorR(data){
					creatorR:= RegExReplace(data, ".*<b>Author\(s\):.+?</td>|</td>.*")
					creatorR:= RegExReplace(creatorR, "<br>", "`n")
					creatorR:= RegExReplace(creatorR, ".*html`" >|</a>.*")
					creatorR:= cleanCreator(creatorR)
					creatorR:= Trim(creatorR)
					return creatorR
}
		;▼▲▼
			getCreatorN(data){
					creatorN:= RegExReplace(data, ".*<b>Author\(s\):.+?</td>|</td>.*")
					if !InStr(creatorN, "vernacular")
						creatorN:= "n/a"
					creatorN:= RegExReplace(creatorN, "^.+?lang=`"..`">|</div>.*")
					;Remove dates and put each other on new line:
						creatorN:= RegExReplace(creatorN, "\([0-9]...-....\) ", "`n")
						creatorN:= RegExReplace(creatorN, "\([0-9]...-\) |\([0-9]...-\)", "`n")
						creatorN:= RegExReplace(creatorN, "(|, )[0-9]...-.... ", "`n")
						creatorN:= RegExReplace(creatorN, "(|, )[0-9]...- ", "`n")
						creatorN:= RegExReplace(creatorN, "(|, )[0-9]... ", "`n")
						creatorN:= RegExReplace(creatorN, " [0-9]...-....", "`n")
					;Remove description of authorship:
						creatorN:= RegExReplace(creatorN, "(|, )editor(. |)|(|, )author(. |)|<span class=vernacular lang=`"..`">|\. ", "`n")
					creatorN:= cleanCreator(creatorN)
					creatorN:= RegExReplace(creatorN, " \^ $| \^$")
					global language
					if(language= "Japanese"){
						creatorN:= RegExReplace(creatorN, ", ")
						}
					creatorN:= Trim(creatorN)
					return creatorN
}
		;▼▲▼
			getCorpR(data){
					corpR:= RegExReplace(data, ".*<b>Corp Author\(s\):.+?</td>|</td>.*")
					corpR:= RegExReplace(corpR, "<br>", "`n")
					corpR:= RegExReplace(corpR, ".*html`" >|</a>.*")
					corpR:= RegExReplace(corpR, ".*html`" >|</a>.*")
					corpR:= RegExReplace(corpR, " \([0-9]...- \)", "`n")
					corpR:= cleanCreator(corpR)
					corpR:= Trim(corpR)
					return corpR
}
		;▼▲▼
			getCorpN(data){
					corpN:= RegExReplace(data, ".*<b>Corp Author\(s\):.+?</td>|</td>.*")
					if !InStr(corpN, "vernacular")
						corpN:= "n/a"
					corpN:= RegExReplace(corpN, "^.+?lang=`"..`">|</div>.*")
					corpN:= RegExReplace(corpN, " <span class=vernacular lang=`"..`">| \([0-9]...- \)", "`n")
					corpN:= cleanCreator(corpN)
					corpN:= Trim(corpN)
					return corpN
}
		;▼▲▼
			getSeriesR(data){
					seriesR:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
					seriesR:= RegExReplace(seriesR, ".*</div>|.*Variation:</b> |v. |\..*")
				;Clean Up
					seriesR:= RegExReplace(seriesR, " (`;`;|`;) .*")
					seriesR:= Trim(seriesR)
					return seriesR
}
		;▼▲▼
			getSeriesN(data){
					seriesN:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
					if !InStr(seriesN, "vernacular")
						seriesN:= "n/a"
					seriesN:= RegExReplace(seriesN, ".*lang=`"..`">|(\.|,|`; |)<.*")
				;Clean Up
					seriesN:= RegExReplace(seriesN, " (`;`;|`;) .*")
					seriesN:= Trim(seriesN)
					return seriesN
}
		;▼▲▼
			getPublisherR(data){
					publisherR:= RegExReplace(data, ".*Publication.+?</td>|</td>.*")
					publisherR:= regExReplace(publisherR, ".*</div>") ;Removes native publisher data to enable check Romanized publisher.
					if !inStr(publisherR, " : ")
						publisherR:= "n/a"
					publisherR:= RegExReplace(publisherR, ".* : |,.*|\.$")
					publisherR:= fixRomanizedPublisherNames(publisherR)
					publisherR:= Trim(publisherR)
					return publisherR
}
		;▼▲▼
			getPublisherN(data){
					publisherN:= RegExReplace(data, ".*Publication.+?</td>|</td>.*")
					publisherN:= RegExReplace(publisherN, "<div><b>Edition:.*</div>") ;Removes native edition in order to parse native publisher name.
					if !InStr(publisherN, "vernacular")
						publisherN:= "n/a"
					publisherN:= RegExReplace(publisherN, "<b>Edition:.*")
					publisherN:= RegExReplace(publisherN, ".*lang=`"..`">|<.*")
					publisherN:= RegExReplace(publisherN, ".* : |,.*|\.$")
					publisherN:= publisherN
					publisherN:= Trim(publisherN)
					return publisherN
}
		;▼▲▼
			getYearR(data){
					yearR:= RegExReplace(data, ".*Year:</b>.+?serif`">")
					yearR:= RegExReplace(yearR, "<nobr>|</nobr>")
					yearR:= RegExReplace(yearR, "<.*|,.*")
					yearR:= Trim(yearR)
					return yearR
}
		;▼▲▼
			getEditionR(data){
					editionR:= RegExReplace(data, ".*<b>Publication:|(\.|)</font></td>.*")
					editionR:= regExReplace(editionR, ".*\.</div>") ;Removes native publisher data to enable check Romanized publisher.
					if !inStr(editionR, "Edition:")
						editionR:= "n/a"
					editionR:= RegExReplace(editionR, ".*Edition:</b> |(\.<|<).*")
				;Clean Up
						if(editionR= "")
							editionR:= "n/a"
					editionR:= Trim(editionR)
					return editionR
}
		;▼▲▼
			getEditionN(data){
				editionN:= RegExReplace(data, ".*<b>Publication:|(|.)</font></td>.*")
					if !inStr(editionN, "vernacular") AND (editionN, "Edition:"){
						editionN:= "n/a"
						return editionN
					}
				editionN:= RegExReplace(editionN, ".*<b>Edition:</b> <span.+?>|\.<.*|<.*")
				editionN:= Trim(editionN)
				return editionN
}
		;▼▲▼
			getSubjects(data){
				subjects:= RegExReplace(data, ".*SUBJECT\(S\).+?serif`">")
				subjects:= RegExReplace(subjects, "<b>Class.*|<b>Note\(s\).*|<b>Responsibility.*|<b>Document.*|<b>Other Titles.*")
				subjects:= regExReplace(subjects, "<br>|<div>|</a>|</div>", "`n") ;Removes content between multiple subject headings.		
				subjects:= regExReplace(subjects, ".*>|&nbsp;")
				subjects:= regExReplace(subjects, "m)^\(8.+?\) |\.$| \(Title\)")
				subjects:= regExReplace(subjects, "B\.C\.|B\.C", "BC")
				subjects:= regExReplace(subjects, "A\.D\.|A\.D", "AD")
				loop{
					if InStr(subjects, "`n`n")
						subjects:= RegExReplace(subjects, "`n`n","`n")
					else
						Break
				}
				subjects:= deDupe(subjects)
				subjects:= Trim(subjects)
				return subjects
}
;▼▲▼
importDataFinishedTutorial(){
		if(tutorialMode= 1){
			spreadsheet:= sheetCheck()
			tutorialGUI()
			tutorialContent(1,	"",
								"Your Spreadsheet",
								"Numpad Plus   OR   F1",
								"select the row of the current active cell to look up your item on FirstSearch.")
			tutorialContent(2,	"=== STOP USING TUTORIAL MODE? =========",
								"Bibliograpic Data to Spreadsheet - Options Interface",
								"Uncheck tutorial mode when you feel comfortable with the hotkeys.",
								"Look like a wizard to you colleagues 🧙‍")
			tutorialShow()
		}
}
;■■■ Pull data from FirstSearch record into spreadsheet.

numpadEnter::{
		tutorialOff()
		activeGUI()
		pullBibData()
		active.Destroy()
		importDataFinishedTutorial()
		
}
F2::{
		tutorialOff()
		activeGUI()
		pullBibData()
		active.Destroy()
		importDataFinishedTutorial()
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
F8::{
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
				Send "{right 14}"
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
				Send "{left 6}"
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
			inCellCheck()
			leftP:= copy()
			Send "{esc}"
			Sleep 25
			Send leftP . rightP
			Sleep 25
			Send "{enter}"
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
^F8::{
		activeGUI()
		fastISBNfix()
		active.Destroy()
}



;■■■■■■■■■■■■■ Convert ISBN13 to ISBN10
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
	convertISBN13toISBN10(isbn13){
		if !inStr(isbn13, "978"){
			active.Destroy()
			return "n/a"
		}
		isbn13:= RegExReplace(isbn13, "-")
		cut:= RegExReplace(isbn13, "^978|.$")
		cut:= RegExReplace(isbn13, "^978|.$")
		s:= StrSplit(cut) ; s = split
		a:= s[1]*10
		b:= s[2]*9
		c:= s[3]*8
		d:= s[4]*7
		e:= s[5]*6
		f:= s[6]*5
		g:= s[7]*4
		h:= s[8]*3
		i:= s[9]*2
		sum:= a+b+c+d+e+f+g+h+i
		remainder:= mod(sum, 11)
		checkNo:= 11-remainder
		if(checkNo= 10)
			checkNo:= "X"
		isbn10:= s[1] . s[2] . s[3] . s[4] . s[5] . s[6] . s[7] . s[8] . s[9] . checkNo
		return isbn10
}
	convertISBN(){
		spreadsheet:= sheetCheck()
		arr:= copyRowMakeArray()
		if(arr[15]= "n/a") | (arr[15]= ""){
			
			MsgBox("There is no ISBN13 to convert into an ISBN10.", stopped, 4096)
			return
		}
		if inStr(arr[15], " "){	
			MsgBox("The ISBN13 appears to not be formatted correctly.", stopped, 4096)
			return
		}
		isbn13:= arr[15] 
		arr[16]:= convertISBN13toISBN10(isbn13)
		if(arr[18]= "Japanese")
			arr[8]:= "https://www.amazon.co.jp/dp/" . arr[16]
		else
			arr[8]:= "https://www.amazon.com/dp/" . arr[16]
		data:= arr[1] . "`t" . arr[2] . "`t" .  arr[3] . "`t" .  arr[4] . "`t" .  arr[5] . "`t" .  arr[6] . "`t" .  arr[7] . "`t" .  arr[8] . "`t" .  arr[9] . "`t" .  arr[10] . "`t" .  arr[11] . "`t" .  arr[12] . "`t" .  arr[13] . "`t" .  arr[14] . "`t" .  arr[15] . "`t" .  arr[16]
		inClip(data)	
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
			Send "{esc}"
			Sleep nt
			copy()
			A_clipboard:= RegExReplace(A_clipboard, "`r`n$")
			if !InStr(A_Clipboard, "`r`n"){
					chatGPTmode:= "oneTitle"
				;▼ Copy native language and native title and verify enough information is available for translation.
					global gptArr
					gptArr:= copyRowMakeArray()
					if(gptArr[18]= "English"){
						active.Destroy()
						MsgBox("This item is already in English and doesn't need to be translated.", stopped, 4096)
						exit
					}
					
					if((gptArr[20]= "n/a") | (gptArr[20]="")) & ((gptArr[21]="n/a") | (gptArr[21]= "")){
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
					if(gptArr[21]= "n/a") | (gptArr[21]="")
						title:= gptArr[20]
					else
						title:= gptArr[21]
					A_clipboard:= ""
			}
			if InStr(A_Clipboard, "`r`n"){
					chatGPTmode:= "bulk"
					titles:= A_Clipboard
			}
		;▼ Go to ChatGPT window
			if WinExist("Translate")
					WinActivate
			else{
					active.Destroy()
					msg:= "There is no browser window with an active tab open to Chat GPT.`n`n"
						. "1) Check you have ChatGPT open in your browser.`n"
						. "2) Make sure ChatGPT is your browser's the active tab.`n"
						. "3) You need to save and keep active a chat called `"Translate`" (case sensitive)."
					MsgBox(msg, stopped, 4096)
					return
			}
			Sleep wt
			findTextOnSite("ChatGPT can make mistakes")
			Send "+{tab}"
			Sleep nt
		;▼ Paste translation request
			if(chatGPTmode= "oneTitle"){
					prompt:= "Provide an English translation of this title:" . gptArr[21] . ". The response should only include the translated title and no other text. Do not put quotation marks around the title."
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
				Send "+{tab}"
				Sleep nt
				Send "^a"
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
				inClip( gptArr[1] . "`t" . gptArr[2] . "`t" . gptArr[3] . "`t" . gptArr[4] . "`t" . gptArr[5] . "`t" . gptArr[6] . "`t" . gptArr[7] . "`t" . gptArr[8] . "`t" . gptArr[9] . "`t" . gptArr[10] . "`t" . gptArr[11] . "`t" . gptArr[12] . "`t" . gptArr[13] . "`t" . gptArr[14] . "`t" . gptArr[15] . "`t" . gptArr[16] . "`t" . gptArr[17] . "`t" . gptArr[18] . "`t" . title)
				pasteToBibSpreadsheet()
		}
	;▼ Bulk title list transtlation process	
		if(chatGPTmode= "bulk"){
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
			data:= bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28] . "`t" . bibArr[29] . "`t" . bibArr[30] . "`t" . bibArr[31] . "`t" . bibArr[32] . "`t" . bibArr[33] . "`t" . bibArr[34] . "`t" . bibArr[35]
			inClip(data)
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
		Send "{right 14}"
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



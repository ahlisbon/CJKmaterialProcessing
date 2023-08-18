﻿;Created by Adam H. Lisbon
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
	activeSearch:= 0

;■■■■■■■■■■■■■ Read values in .ini file
	;Populate variables from INI file
		CD:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "CD") ;CD = Collection Development
		DI:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "DI") ;DI = Donation Intake
		US:=			IniRead("BibData to Spreadsheet.ini", "Sheet Names", "US") ;US = User Selects
		fsURL:=			IniRead("BibData to Spreadsheet.ini", "Settings", "fsURL")
		checkMode:=		IniRead("BibData to Spreadsheet.ini", "Settings", "checkMode")



;■■■■■■■■■■■■■ Run GUI
	; GUI Interface
		bib := Gui(, "Bibliographic Data To Spreadsheet - Options")
	;Question 1:
		bib.Add("Text",		"					x180 y20",	"▼ File Name Prefixes of Your Spreadsheets (Case Sensitive)")
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
		bib.Add("Text", "						x10	 y200", "Use Check Mode:")
		bib.Add("Checkbox",	"vcheckMode			x180 y200")
	;Question 5:
		bib.Add("Text",		"					x10  y230", "Wait longer for websites to load:")
		bib.Add("DDL", 		"vloadTime	w30		x180 y225 Choose1", ["1", "2", "3"])
	;Process Answers into variables
		bib.Add("Button",	"default			x180 y270", "&Save Settings").OnEvent("Click", settings)
	;Help Text/Links
		bib.Add("Link", 	"					x10  y310", "Read the <a href=`"https://github.com/ahlisbon/CJKmaterialProcessing#-hotkeys-to-activate-macro`">Hotkey Guide</a> on GitHub")
		bib.Add("Link",		"					x192 y40",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----file-name-prefixes`">Read about file naming conventions</a>")
		bib.Add("Link",		"					x220 y229",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----wait-longer-for-websites-to-load`">What is this?</a>")
		bib.Show()
	;Save and error check inputs
		settings(*){
			;Post that settings are updated
				bib.Add("Text", 	"			x264 y275", "✔ Updated")
				bib.Show()
			;Set variables
				saved:= bib.Submit(0)
				;Check Mode
					global checkMode
					checkMode:= saved.checkMode
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
}
			checkGUIinputs(CD, DI, US, fsURL){
				;Blank value alert
					if(CD= "") & (DI= "") & (US= ""){
						MsgBox("At least one type of spreadsheet must have a name.")
						exit
						}
					if(fsURL= ""){
						MsgBox("The field for your institution's FirstSearch URL cannot be blank.")
						exit
					}
				;Duplicate value alert
					if((CD!="") & (DI="") & (US="")) | ((CD="") & (DI!="") & (US="")) | ((CD="") & (DI="") & (US!=""))
						return
					if(CD= DI) | (CD= US) | (DI= US){
						MsgBox("You cannot have two file names that are identical.")
						exit
					}
}
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
;▼▲▼ Error checking to stop script if criteria aren't met. Declares what spreadsheet is actively being used.
	sheetCheck(){
			spreadsheet:= ""
			if(CD= "") & (DI= "") & (US= "")
				MsgBox("There is no spreadsheet open with a matching prefix.`n`nCheck if:`n your file name matches the names you submitted'nThe file is open.")
			if((WinExist(CD) & WinExist(DI)) | (WinExist(CD) & WinExist(US)) | (WinExist(DI) & WinExist(US))){
				MsgBox("You have at least two different types of spreadsheets open. Please close all other spreadsheets except for the one you are actively working on.")
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
		spreadsheet:= sheetcheck()
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
		sleep wt
}



;■■■■■■■■■■■■■ Search FirstSearch/Worldcat with data from spreadsheet.
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
	getDataFromBibSpreadsheet(){
			spreadsheet:= sheetCheck()
		;Get data from spreadsheet
			dataHere(spreadsheet)
			Send "{esc}"
			Send "{home}"
			Send "+{space}"
			copy()
			global bibArr:= strSplit(A_Clipboard, A_Tab)
			if(bibArr[15]= "") & (bibArr[16]= "") & (bibArr[17]= "") & (bibArr[20]= "") & (bibArr[21]= ""){
				MsgBox("The script stopped because there is no searchable data in your spreadsheet.", stopped)
				exit
			}
}
	searchFS(){
			activateBrowser()
			newWin(fsURL)
			Sleep lt
		;▼ Search priority= oclc, isbn13, isbn10, native tile, romanized title
			bibArr[21]:= RegExReplace(bibArr[21], " : ", ": ")
			searchText:= searchWith(bibArr[17], bibArr[15], bibArr[16], bibArr[21], bibArr[20])
			Send searchText
			Sleep nt
			Send "{enter}"
}
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
;■■■
numpadAdd::{
		activeSearch:= 1
		getDataFromBibSpreadsheet()
		searchFS()
}
F1::{
		activeSearch:= 1
		getDataFromBibSpreadsheet()
		searchFS()
}



;■■■■■■■■■■■■■ Open other versions link in FirstSearch Detailed Record.
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
	showRecordsList(){
			siteText:= copyAll()
		;Open other versions link in FirstSearch Detailed Record.
			if InStr(siteText, "Search for versions with same title and author"){
				findTextOnSite("Search for versions with same title and author")
				Send "{enter}"
			}else
				Send "{tab}"
}
	openRecordsList(){
		;Determine how many tabs to open.
			data:= copyAll()
			data:= RegExReplace(data, "`r|`n|`t")
			if inStr(data, "Records Found: "){
				results:= RegExReplace(data, ".*Records found: | .*")
				if(results > 10)
					results:= 10
		;Loop through results and open tabs
				entry:= 1
				Loop results{
					findTextOnSite(entry . ".")
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
numpadAdd::		showRecordsList()
F2::			showRecordsList()
#HotIf WinActive("WorldCat List of Records")
numpadAdd::		openRecordsList()
F2::			openRecordsList()



;■■■■■■■■■■■■■ Put bibliographic data from a FirstSearch detailed record and put it in a spreadsheet.
#HotIf WinActive("Detailed Record")
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
	normalize(data){
		; Convert Japanese punctuation to American punctuation.
		data:= RegExReplace(data, "  ", " ")
		data:= RegExReplace(data, "：", ":")
		data:= RegExReplace(data, "、", ".")
		loop{
				if InStr(data, "  ")
					data:= RegExReplace(data, "  "," ")
				else
					Break
			}
		; Turn composed characters into decomposed characters.
			data:= regExReplace(data, "ā", "ā")
			data:= regExReplace(data, "ī", "ī")
			data:= regExReplace(data, "ū", "ū")
			data:= regExReplace(data, "ē", "ē")
			data:= regExReplace(data, "ō", "ō")
		return data
}
	getVolumes(data){
			volumes:= RegExReplace(data, ".*<b>Description:|</font></td>.*")
			if inStr(volumes, " volumes")
				volumes:= RegExReplace(volumes, ".*serif`">| volumes.*")
			else
				volumes:= "n/a"
			volumes:= Trim(volumes)
			return volumes
}
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
	deDupe(txt){
			str:= ""
			loop Parse, txt, "`r"
				if !InStr(str, A_LoopField "`n")
					str .= A_LoopField "`n"
			return str
}
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
	getCreatorR(data){
			creatorR:= RegExReplace(data, ".*<b>Author\(s\):.+?</td>|</td>.*")
			creatorR:= RegExReplace(creatorR, "<br>", "`n")
			creatorR:= RegExReplace(creatorR, ".*html`" >|</a>.*")
			creatorR:= cleanCreator(creatorR)
			creatorR:= Trim(creatorR)
			return creatorR
}
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
			creatorN:= RegExReplace(creatorN, " \^ $")
			global language
			if(language= "Japanese"){
				creatorN:= RegExReplace(creatorN, ", ")
				}
			creatorN:= Trim(creatorN)
			return creatorN
}
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
	getSeriesR(data){
			seriesR:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
			seriesR:= RegExReplace(seriesR, ".*</div>|.*Variation:</b> |v. |\..*")
		;Clean Up
			seriesR:= RegExReplace(seriesR, "`; `;|`;`;", ";")
			seriesR:= RegExReplace(seriesR, "; <.*|</font>.*")
			seriesR:= Trim(seriesR)
			return seriesR
}
	getSeriesN(data){
			seriesN:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
			if !InStr(seriesN, "vernacular")
				seriesN:= "n/a"
			seriesN:= RegExReplace(seriesN, ".*lang=`"..`">|(\.|,|`; |)<.*")
		;Clean Up
			seriesN:= RegExReplace(seriesN, "`; `;|`;`;", " `; ")
			seriesN:= RegExReplace(seriesN, "  ", " ")
			seriesN:= Trim(seriesN)
			return seriesN
}
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
	getYearR(data){
			yearR:= RegExReplace(data, ".*Year:</b>.+?serif`">")
			yearR:= RegExReplace(yearR, "<nobr>|</nobr>")
			yearR:= RegExReplace(yearR, "<.*|,.*")
			yearR:= Trim(yearR)
			return yearR
}
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
	getSubjects(data){
		subjects:= RegExReplace(data, ".*SUBJECT\(S\).+?serif`">")
		subjects:= RegExReplace(subjects, "<b>Class.*|<b>Note\(s\).*|<b>Responsibility.*|<b>Document.*|<b>Other Titles.*")
		subjects:= regExReplace(subjects, "<br>|<div>|</a>|</div>", "`n") ;Removes content between multiple subject headings.		
		subjects:= regExReplace(subjects, ".*>|&nbsp;")
		subjects:= regExReplace(subjects, "m)^\(8.+?\) |\.$| \(Title\)")
		subjects:= regExReplace(subjects, "B\.C\.|B\.C", "BC")
		subjects:= regExReplace(subjects, "A\.D\.|A\.D", "AD")
		;Remove French subjects.
		subjects:= RegExReplace(subjects, "i).*chine$.*|.*chine .*|.*chinoise.*|.*corée.*|.*japon.*|.*japonais.*|.*bibliographie.*|.*bouddhique.*|.*bouddhisme.*|.*confucéenne.*|.*confucianisme.*|.*dictionnaires.*|.*économique.*|.*essais.*|.*histoire.*|.*littérature.*|.*mythologie.*|.*philosophie.*|.*pratique.*|.*poésie.*|.*ï.*")
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
	pullBibData(){
		;Error check that spreadsheet and fsURL variables are not blank.
			checkGUIinputs(CD, DI, US, fsURL)
			spreadsheet:= sheetCheck()
		;FirstSearch
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
					global activeSearch
					if(activeSearch= 0){
						bibArr:= []
						Loop 7{
							bibArr.InsertAt(1, "n/a")
						}
						Loop 28{
							bibArr.InsertAt(1, "")
						}
					}
					
				;🏛 Check against local holdings
					if InStr(data, "FirstSearch indicates your institution owns the item.")
						bibArr[14]:= "y"
					else
						bibArr[14]:= "n"
						
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
		;Check Results
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
				yesNo:= MsgBox(checkData, "Review Bibliographic Data", "YesNo")
				if(yesNo= "No")
					exit
			}
			subjects:= regExReplace(subjects, "`n", " ^ ") ;Each subject is on a serpate line for easy reading in the :Check Data message box." This puts them on one line to be pasted into a single cell on the spreadsheet.
			subjects:= RegExReplace(subjects, " \^  \^ $| \^ $")
		;Calculate other array fields based on pulled data.
			;Series Number
				if InStr(seriesR, " `; ")
					vol:= RegExReplace(seriesR, ".* `; ")
				else if InStr(seriesN, " `; ")
					vol:= RegExReplace(seriesN, ".* `; ")
				else
					vol:= "n/a"
				bibArr[10]:= vol
			
			;Generate URL for Purchase
				if(language="Japanese") & (isbn10!= "") & (isbn10!= "n/a")
						bibArr[8]:= "https://www.amazon.co.jp/dp/" . isbn10
				else
					bibArr[8]:= "n/a"
				if inStr(isbn10, A_space)
					bibArr[8]:= "n/a"
		
		;Put parsed data into array.
			bibArr[11]:= volumes
			bibArr[15]:= isbn13
			bibArr[16]:= isbn10
			bibArr[17]:= oclc
			bibArr[18]:= language
			bibArr[19]:= titleT
			bibArr[20]:= titleR
			bibArr[21]:= titleN
			bibArr[22]:= creatorsR
			bibArr[23]:= creatorsN
			bibArr[24]:= seriesR
			bibArr[25]:= seriesN
			bibArr[26]:= publisherR
			bibArr[27]:= publisherN
			bibArr[28]:= yearR
			bibArr[29]:= yearN
			bibArr[30]:= editionR
			bibArr[31]:= editionN
			bibArr[32]:= subjects
			data:= bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28] . "`t" . bibArr[29] . "`t" . bibArr[30] . "`t" . bibArr[31] . "`t" . bibArr[32]
			inClip(data)
		;Close FirstSearch "Detailed Record" page.
			if WinExist("Detailed Record") 
				WinActivate
			sleep wt
			if(activeSearch= 1){
				Send "!{F4}"
				sleep 500
			}
		;Populate spreadsheet with data.
			activeSearch:= 0
			pasteToBibSpreadsheet()
}
;■■■ Pull data from FirstSearch record into spreadsheet.
numpadEnter::pullBibData()
F1::pullBibData()



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
		keep:= fix.Submit()
		keep:= keep.v
		fixIarr:= copyRowMakeArray()
		i13:= fixIarr[15]
		i10:= fixIarr[16]
		i13:= isolateISBN("...", i13, keep, 13)
		i10:= isolateISBN("", i10, keep, 10)
		data:= fixIarr[1] . "`t" . fixIarr [2] . "`t" . fixIarr [3] . "`t" . fixIarr [4] . "`t" . fixIarr [5] . "`t" . fixIarr [6] . "`t" . fixIarr [7] . "`t" . fixIarr [8] . "`t" . fixIarr [9] . "`t" . fixIarr [10] . "`t" . fixIarr [11] . "`t" . fixIarr [12] . "`t" . fixIarr [13] . "`t" . fixIarr [14] . "`t" . i13 . "`t" . i10
		inClip(data)
		pasteToBibSpreadsheet()
	}
	;▼▲▼ Find and Replace ISBN
			isolateISBN(whichI, ISBN, keep, logic){
				if(RegExMatch(keep, "[0-9].......[0-9]")){
					if(logic= 10) & inStr(ISBN, keep)
						return keep
					if(logic= 10) & !inStr(ISBN, keep)
						return ISBN	
					if(logic= 13) & inStr(ISBN, keep)
						return keep
					if(logic= 13) & !inStr(ISBN, keep)
						return ISBN 
				}
				if isInteger(keep){
					ISBN:= RegExReplace(ISBN, "\(v. ", "(")
					ISBN:= RegExReplace(ISBN, " \(" keep "\).*")
					ISBN:= RegExReplace(ISBN, ".* ")
				}
				if(keep= "s") | (keep= "p") | (keep= "h"){
					ISBN:= RegExReplace(ISBN, " \(" keep "\).*")
					ISBN:= RegExReplace(ISBN, ".* ")
				}
				if(keep= "ns") | (keep="np") | (keep= "nh"){
					if(keep="ns")
						getRid:= "set"
					if(keep="np")
						getRid:= "(paperback|pbk\.)"
					if(keep="nh")
						getRid:= "(hardcover|hbk\.)"
					ISBN:= RegExReplace(ISBN, "i)^[0-9]." whichI "........ \(" getRid "\) \^ | \^ [0-9]." whichI "........ \(" getRid "\)$| \^ [0-9]." whichI "........ \(" getRid "\)")
				}
				yesNo:= MsgBox(ISBN, "Was the correct ISBN Parsed?", "yesNo")
				if(yesNo= "Yes")
					return ISBN
				else{
					dataHere(spreadsheet)
					Send "{esc}"
					Sleep nt
					Send "{home}"
					Sleep nt
					Send "{right 14}"
					exit
				}		
			}
}
;■■
^numpad8::removeISBN()
F8::removeISBN()



;■■■■■■■■■■■■■ Fast ISBN Fix
#HotIf WinActive(CD) | WinActive(DI)
;Relative to the cursor in an active cell, keeps only the ISBN that the cell is in.
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼ Copies text in active cell.
	fastISBNfix(){
			Send "^+{right}"
			Sleep nt
			inCellCheck()
			rightP:= copy()
			rightP:= Trim(rightP)
			Send "^+{left}"
			Sleep nt
			inCellCheck()
			leftP:= copy()
			Send "{esc}"
			Sleep nt
			Send leftP . rightP
			Sleep nt
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
					MsgBox("You tried to fix a cell in the ISBN13 or ISBN10 columns but weren't `"inside`" the cell. Double click a cell and have your cursor within the ISBN you want to keep in order to clean up your ISBNs.", stopped)
					exit
				}
}
;■■
numpadDiv::fastISBNfix()



;■■■■■■■■■■■■■ Convert ISBN13 to ISBN10
#HotIf WinActive(CD) | WinActive(DI)
;▼▲▼▲▼▲▼▲▼▲▼▲▼
;▼▲▼
	convertISBN13toISBN10(data){
		if !inStr(data, "978")
			exit
		isbn13:= data
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
			MsgBox("There is no ISBN13 to convert into an ISBN10.", stopped)
			return
		}
		if inStr(arr[15], " "){
			MsgBox("The ISBN13 appears to not be formatted correctly.", stopped)
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
^numpad7::convertISBN()
F7::convertISBN()



;■■■■■■■■■■■■■ Translate title with ChatGPT.
#HotIf WinActive(CD) | WinActive(DI)
;■■
numpadSub::{
		spreadsheet:= sheetCheck()
		global chatGPTmode
		global gptPending
		gptPending:= 0
		copy()
		A_clipboard:= RegExReplace(A_clipboard, "`r`n$")
		if !InStr(A_Clipboard, "`r`n"){
			chatGPTmode:= "oneTitle"
		;Copy native language and native title and verify enough information is available for translation.
			global gptArr
			gptArr:= copyRowMakeArray()
			if(gptArr[18]= "n/a") | (gptArr[18]= ""){
				MsgBox("The language of this item can't be identified:``n`nReview Column R: Language", stopped)
				exit
			}
			if(gptArr[18]= "English"){
				MsgBox("This item is already in English and doesn't need to be translated.")
				exit
			}
			
			if((gptArr[20]= "n/a") | (gptArr[20]="")) & ((gptArr[21]="n/a") | (gptArr[21]= "")){
				MsgBox("There is no non-English title to translate.")
				exit
			}
			title:= copy()
			if(title= "n/a") | (title= ""){
				MsgBox("There is no title to translate.`n`nReview Column 21", stopped)
				exit
				}
		;Determine which title data to use
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
	;Go to ChatGPT window
		if WinExist("Translate")
			WinActivate
		else{
			msg:= "There is no browser window with an active tab open to Chat GPT.`n`n"
				. "1) Check you have ChatGPT open in your browser.`n"
				. "2) Make sure ChatGPT is your browser's the active tab.`n"
				. "3) You need to save and keep active a chat called `"Translate`" (case sensitive)."
			MsgBox(msg, stopped)
			return
		}
		Sleep wt
		findTextOnSite("ChatGPT may produce inaccurate information")
		Send "+{tab}"
		Sleep nt
	;Paste translation request
		if(chatGPTmode= "oneTitle"){
			prompt:= "Provide English translations of this " . gptArr[18] . " title. Do not include the original "	. gptArr[18] . " title:" . title
			inClip(prompt)
			Send "^v"
			Sleep nt
			Send "{enter}"
			gptPending:= 1
		}
		if(chatGPTmode= "bulk"){
			prompt:= "Provide English translations of this list of titles:" . "`r`n" . titles
			inClip(prompt)
			Send "^v"
			Sleep nt
			Send "{enter}"
			gptPending:= 1
		}
}
;■■
#HotIf winActive("Translate")
numpadSub::{
		spreadsheet:= sheetcheck()
		global gptPending
		if(gptPending= 0){
			MsgBox("An error has occured while using ChatGPT to translate a title.`n`nRestart the script and try again.", stopped)
			exit
		}	
		global chatGPTmode
		if(chatGPTmode= "oneTitle"){
		;Copy response
			Send "+{tab}"
			Sleep nt
			Send "^a"
			Sleep nt
			title:= copyAll()
			Send "{tab}"
			Sleep nt
		;Isolate translated title
			title:= RegExReplace(title, "`r`n")
			title:= RegExReplace(title, "Free Research Preview.*")
			title:= RegExReplace(title, ".*ChatGPT")
			title:= title . " - ChatGPT translation"
		;Determine if the translation is acceptable
			yesNo:= MsgBox(title, "Is this translation acceptable?", "yesNo")
			if(yesNo= "No"){
				dataHere(spreadsheet)
				Send "{esc}"
				Sleep nt
				Send "{home}"
				gptPending:= 0
				exit
			}	
			if(yesNo= "Yes"){				
			;Paste translation request
				data:= gptArr[1] . "`t" . gptArr[2] . "`t" . gptArr[3] . "`t" . gptArr[4] . "`t" . gptArr[5] . "`t" . gptArr[6] . "`t" . gptArr[7] . "`t" . gptArr[8] . "`t" . gptArr[9] . "`t" . gptArr[10] . "`t" . gptArr[11] . "`t" . gptArr[12] . "`t" . gptArr[13] . "`t" . gptArr[14] . "`t" . gptArr[15] . "`t" . gptArr[16] . "`t" . gptArr[17] . "`t" . gptArr[18] . "`t" . title
				inClip(data)
				pasteToBibSpreadsheet()
				gptPending:= 0
				exit
			}
		}
		if(chatGPTmode= "bulk"){
		;Copy Response
			Send "+{tab}"
			Sleep nt
			Send "^a"
			Sleep nt
			titles:= copyAll()
			Send "{tab}"
			Sleep nt
		;Isolate translated title
			titles:= RegExReplace(titles, "`r`n")
			titles:= RegExReplace(titles, "Free Research Preview.*")
			titles:= RegExReplace(titles, ".*ChatGPT")
			titles:= RegExReplace(titles, "^.+?`"")
			titles:= RegExReplace(titles, "`"$", " - ChatGPT translation")
			titles:= RegExReplace(titles, "`"    `"", "`r`n")
			titles:= RegExReplace(titles, "`r`n", " - ChatGPT translation`r`n")
		;Determine if the translation is acceptable
			titlesForMsg:= RegExReplace(titles, " - ChatGPT translation")
			yesNo:= MsgBox(titlesForMsg, "Is this translation acceptable?", "yesNo")
			if(yesNo= "No"){
				dataHere(spreadsheet)
				Send "{esc}"
				Sleep nt
				Send "{home}"
				gptPending:= 0
				exit
			}	
			if(yesNo= "Yes"){				
			;Paste translation request
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
				exit
			}
		}
}



;■■■■■■■■■■■■■ Price Look up
;Applies hierarchy to seach with ISBN10 first, then ISBN13, then title in native script.
;▼▲▼▲▼▲▼▲▼▲▼▲▼
#HotIf WinActive(CD) | WinActive(DI)
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
;■■■
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
		MsgBox("At this time, price checking options are only available for Japanese materials.", stopped)
			exit
		}
}
^numpadAdd::searchPrice()
F4::searchPrice()



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
			exit
}
;▼▲▼ Furuhonya - kosho.or.jp
	furuhonya(){
			spreadsheet:= sheetCheck()
			price:= copyAll()
			if !InStr(price, "￥"){
				MsgBox("There is no price to add from www.kohso.or.jp", stopped)
				exit
			}
			price:= RegExReplace(price, "`r|`n|`t", "xtx")
			if inStr(price, "の検索結果"){
				MsgBox("You are on a list of search results for www.kosho.or.jp.`n`nPlease load one of the records from the results to import the price into the spreadsheet.", stopped)
				exit
			}
			price:= RegExReplace(price, "^.+?￥|xtx.*")
			price:= RegExReplace(price, ",")
			priceToBibSpreadsheet(price, "jpy")
}
;▼▲▼
	amazonUS(){
			spreadsheet:= sheetcheck()
			A_Clipboard:= ""
			Sleep nt
			Send "^c"
			Sleep nt
			if !ClipWait(2){
				MsgBox("Nothing on Amazon.com (US site) has been highlighted to copy.`n`nWhen highlighting a price to copy make sure all the numbers and dollar sign ($) are hihglighted to import the price into the spreadsheet.", stopped)
				exit
			}
			price:= A_Clipboard
			if !inStr(price, "$"){
				MsgBox("You may have tried to highlight a price but there was no dollar sign ($) in the text you highlighted.", stopped)
				exit
			}
			price:= RegExReplace(price, "\$")
			priceToBibSpreadsheet(price, "usd")
}

;▼▲▼
	amazonJP(){
			spreadsheet:= sheetcheck()
			A_Clipboard:= ""
			Sleep nt
			Send "^c"
			Sleep nt
			if !ClipWait(2){
				MsgBox("Nothing on Amazon.co.jp (Japan site) has been highlighted to copy.`n`nWhen highlighting a price to copy make sure all the numbers and yen sign (¥) are hihglighted to import the price into the spreadsheet.", stopped)
				exit
			}
			price:= A_Clipboard
			if !inStr(price, "￥"){
				MsgBox("You may have tried to highlight a price but there was no yen sign (¥) in the text you highlighted.", stopped)
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
^numpadEnter::getPrice()
^enter::getPrice()



;■■■■■■■■■■■■■ Conveniences
#HotIf WinActive(CD) | WinActive(DI)
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
^numpadEnter::moveToISBN()
^enter::moveToISBN()

;■■■ Speed up browsing through tabs to look at First Search results
#HotIf WinActive("WorldCat Detailed Record") | WinActive("WorldCat List of Records") | WinActive("")
		numpad0::Send "^{tab}"	;go to next tab
		numpadSub::Send "^w"	;close tab

;■■■ Speed up browsing through tabs to look at First Search results
#HotIf WinActive("ahk_exe firefox.exe") | WinActive("ahk_exe msedge.exe") | WinActive("ahk_exe chrome.exe")
		^numpad0::Send "^{tab}"	;go to next tab
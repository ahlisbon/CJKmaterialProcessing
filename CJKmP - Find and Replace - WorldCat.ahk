;■■■■■■■■■■■■■ Put bibliographic data from a WorldCat.org record and put it in a spreadsheet.
;▼▲▼ Clean up creator data
creatorCleanup(clean){
	;fix lines
		clean:= fnr(clean, " `;|`;", "`n")
	;clear junk lines
		clean:= fnd(clean, ".*(\[|\]|\{|\}).*")
	;beginning and end
		clean:= fnd(clean, "m)\.$")
	;other
		clean:= fnd(clean, "^by | = .*")
	;remove role at end of names
		clean:= fnd(clean, "原作|漫画")
		clean:= fnd(clean, "m)(,) author$| cho$|著$| hen$|編$")
	;remove dates
		clean:= fnd(clean, ", \d{4}-\d{4}|, \d{4}-")
	;fix lines
		clean:= fnr(clean, "`n`n", "`n")
		clean:= fnr(clean, "`n", " ^ ")
	;beginning and end
		clean:= fnd(clean, "^ \^ | \^ $")
		clean:= fnr(clean, ", \^ ", " ^ ")
		return clean
}

;▼▲▼
pullBibDataWC(){
			global bibArr
			global activeSearch
		;▼ Error check that spreadsheet and fsURL variables are not blank.
			checkGUIinputs(CD, DI, US, fsURL)
			spreadsheet:= sheetCheck()
		;▼ WorldCat Load Check
			if !WinExist(" | WorldCat.org"){
				exit
				active.destroy()
			}
			WinActivate
		;Determine if active search is happening
			if(activeSearch= 0){
					bibArr:= []
					Loop 35{
						bibArr.InsertAt(1, "")
						}
					bibArr[1]:= ""
					bibArr[9]:= ""
			}

;▼ --------------- Bibiographic Data from WorldCat Record Page ---------------
		data:= copyAll()
		data:= singleLine(data, 1)

	;🏛 Check against local holdings
;add search string to gui to automatically link back to records
		if inStr(data, "Borrow from " . libName)
			dupe:= "yes"
		else
			dupe:= "n"
	;👬 Get count of libraries that also have this item
		libCount:= isolate(data, "Find a Copy at a Library", ".*Find a Copy at a Library|Filter.*")
		libCount:= fnd(libCount, " libraries.*")
		libCount:= fnd(libCount, ".* in ")

;▼ --------------- Bibliographic Data Cleanup - Bib Data from Source Code ---------------
		data:= loadSourceCode("pageProps", "WorldCat Record")
		data:= singleLine(data, 1)


	;🗣 Language
			language:= isolate(data, "<!-- --> <span>", ".*<!-- --> <span>|</span>.*")


	;Isolate source code	
		data:= isolate(data, "pageProps", ".*pageProps|surveymonkeyUrl.*")
		;Simplify formatting
			data:= fnr(data, "`":{`"|`":`"|`":", ": ")
			dataForSubjects:= data
			data:= fnr(data, "`"},`"|`",`"|,`"", ", ")
			data:= fnd(data, "[`"'|`"]")


	;📚 Total volumes of multivolume sets
			if inStr(data, "physicalDescription: volumes")
				volumes:= "n/a"
			else
				volumes:= isolate(data, "volumes : ", ".*physicalDescription: | .*")

		
	;🔢 ISBN
			isbn:= isolate(data, "isbns", ".*isbns: ")			;need [ to identify first isbn
			isbn:= fnr(isbn, "\].*", "]")						;need ] to identify last isbn
			if inStr(isbn, "null"){
				isbn13:= "n/a"
				isbn10:= "n/a"
			}else{
				;ISBN13
					isbn13:= fnd(isbn, " \d{10},| \d{9}X,")			;delete isbn10s in middle
					isbn13:= fnd(isbn13, ", \d{10}\]|, \d{9}X\]")	;delete last isbn10
					isbn13:= fnd(isbn13, "\[|\]")
					isbn13:= fnr(isbn13, ", ", " ^ ")
				;ISBN10
					isbn10:= fnd(isbn, " \d{13},")					;delete isbn13s in middle
					isbn10:= fnd(isbn10, "\[\d{13}, ")				;delete first isbn13
					isbn10:= fnd(isbn10, ", \d{13}\]|\d{13}\]")		;delete last isbn13
					isbn10:= fnd(isbn10, "\[|\]")
					isbn10:= fnr(isbn10, ", ", " ^ ")
					if(isbn10= "")
						isbn10:= "n/a"
				;ISBN Conversions
					;10 to 13
						if(!inStr(isbn10, " ") & (isbn10!= "n/a"))
							isbn13:= convert10to13(isbn10)
					;13 to 10
						if(!inStr(isbn13, " ") & (isbn13!= "n/a"))
							isbn10:= convert13to10(isbn13)
							
				;🔢 ISSN
					issn:= isolate(data, "issns: ", ".*issns: \[|].*")
					if(isbn= "n/a"){
						isbn10:= issn
						volumes:= "n/a"
					}
			}
			
			
	;💰 Price URL
			priceURL:= getPriceURL(isbn10, language)
			
	;🔢 OCLC#
			oclc:= isolate(data, "oclcNumber: ", ".*oclcNumber: |,.*")
	
	
	;📔 Title
			title:= isolate(data, "titleInfo: ", ".*titleInfo: text: |, languageCode.*|, creator.*")
			if inStr(title, "romanizedText"){
				titleR:= fnd(title, ".*romanizedText: | = .*")
				titleN:= fnd(title, ".*text: |, romanizedText.*| = .*")
			}else{
				titleR:= fnd(title, ".*text: ")
				titleN:= "n/a"
			}
		;Translated Title
			if inStr(title, " = ")
				titleT:= fnd(title, ".* = ")
			else
				titleT:= "n/a"
				
				
	;✒ Creators
			creator:= isolate(data, "contributors: ", ".*contributors: |, accessibilityContent.*")
					;Native Creator
						if !inStr(creator, "romanizedText")
							creatorN:= "n/a"
						else{
							;Primary Native Creator
								if !inStr(creator, "firstName")
										creatorNP:= "n/a"
								else{
										creatorNP:= fnd(creator, ", contributorNotes.*")
										creatorNP:= fnd(creatorNP, ".*fromStatementOfResponsibility:")
										creatorNP:= fnr(creatorNP, "firstName: text: ", "`n`n")
										creatorNP:= fnr(creatorNP, "nonPersonName: text:", "`n`nnonPersonName")
										creatorNP:= fnd(creatorNP, ".*nonPersonName.*")
										creatorNP:= fnd(creatorNP, ", romanizedText.*")
									;Cleanup
										;creatorNP:= fnd(creatorNP, "m)^\[.*|.*\{.*")
										creatorNP:= creatorCleanup(creatorNP)
										creatorN:= creatorNP
								}
							;Secondary Native Creator
								 creatorNS:= isolate(creator, "contributorNotes: ", "^.+?contributorNotes: ")
								 if !inStr(creatorNS, "firstName")
									creatorNS:= "n/a"
								else{
										creatorNS:= fnr(creatorNS, "firstName: text: ", "`n`n")
										creatorNS:= fnr(creatorNS, "nonPersonName: text:", "`n`nnonPersonName")
										creatorNS:= fnd(creatorNS, ".*nonPersonName.*")
										creatorNS:= fnd(creatorNS, ", .*")
									;Cleanup
										;creatorNS:= fnd(creatorNS, ".*\{.*")
										creatorNS:= creatorCleanup(creatorNS)
								}	
							;Merge primary and secondary
								if((creatorNP= "n/a") & (creatorNS= "n/a"))
									creatorN:= "n/a"
								if((creatorNP!= "n/a") & (creatorNS= "n/a"))
									creatorN:= creatorNP
								if((creatorNP= "n/a") & (creatorNS!= "n/a"))
									creatorN:= creatorNS
								if((creatorNP!= "n/a") & (creatorNS!= "n/a"))
									creatorN:= creatorNP . " ^ " . creatorNS						
							}
					;Romanized Creator ＊＊＊with＊＊＊ Romanized text in WorldCat metadata
						if inStr(creator, "romanizedText"){
							;Primary Romanized Creator
								creatorRP:= fnd(creator, "contributorNotes.*")
								if !inStr(creatorRP, "romanizedText")
									creatorRP:= "n/a"
								if !inStr(creatorRP, "firstName")
									creatorRP:= "n/a"
								if(creatorRP!= "n/a"){
										creatorRP:= fnd(creatorRP, ".*fromStatementOfResponsibility:")
										creatorRP:= fnr(creatorRP, "firstName: text: ", "`n`n")
										creatorRP:= fnr(creatorRP, "nonPersonName: text:", "`n`nnonPersonName")
										creatorRP:= fnd(creatorRP, ".*nonPersonName.*")
										creatorRP:= fnd(creatorRP, ", languageCode.+?romanizedText:")
										creatorRP:= fnd(creatorRP, ".*romanizedText: |,.*")
										creatorRP:= fnd(creatorRP, "m)^\[.*|.*\{.*")
									;cleanup
										;creatorRP:= fnd(creatorRP, ".*\}.*")
										creatorRP:= creatorCleanup(creatorRP)
								}
							;Secondary Romanized Creator
								creatorRS:= isolate(creator, "contributorNotes: ", "^.+?contributorNotes: ")
								if !inStr(creatorRS, "firstName")
									creatorRS:= "n/a"
								else{
										creatorRS:= fnr(creatorRS, "firstName: text: ", "`n`n")
										creatorRS:= fnr(creatorRS, "nonPersonName: text:", "`n`nnonPersonName")
										creatorRS:= fnd(creatorRS, ".*nonPersonName.*")
										creatorRS:= fnd(creatorRS, ", languageCode.+?romanizedText:")
										creatorRS:= fnd(creatorRS, ".*romanizedText: |,.*")
									;Cleanup
										;creatorRS:= fnd(creatorRS, ".*\}.*")
										creatorRS:= creatorCleanup(creatorRS)
								}
					;Romanized Creator ＊＊＊without＊＊＊	Romanized text in WorldCat metadata
						}else{
							;Primary Romanized Creator
									creatorRP:= fnd(creator, ".*fromStatementOfResponsibility:")
									creatorRP:= fnr(creatorRP, "firstName: text: ", "`n`n")
									creatorRP:= fnr(creatorRP, "nonPersonName: text:", "`n`nnonPersonName")
									creatorRP:= fnd(creatorRP, ".*nonPersonName.*")
									creatorRP:= fnd(creatorRP, ", secondName.*text:|, isPrimary.*")
									creatorRP:= fnd(creatorRP, "m)^\[.*|.*\{.*")
								;cleanup
									creatorRP:= creatorCleanup(creatorRP)
							;Secondary Romanized Creator
									creatorRS:= "n/a"
						}
							;Merge primary and secondary
								if((creatorRP= "n/a") & (creatorRS= "n/a"))
									creatorR:= "n/a"
								if((creatorRP!= "n/a") & (creatorRS= "n/a"))
									creatorR:= creatorRP
								if((creatorRP= "n/a") & (creatorRS!= "n/a"))
									creatorR:= creatorRS
								if((creatorRP!= "n/a") & (creatorRS!= "n/a"))
									creatorR:= creatorRP . " ^ " . creatorRS
		;Corporate Creators
				if inStr(data, "nonPersonName"){
						corp:= isolate(data, "contributors: ", ".*contributors: |, accessibilityContent.*")
					;Native Corp
						if !inStr(corp, "romanizedText")
							corpNP:= "n/a"
						else{	
							;Primary Native Corp
								if !inStr(corp, "nonPersonName")
									corpNP:= "n/a"
								else{
										corpNP:= fnd(corp, "contributorNotes.*")
										corpNP:= fnd(corpNP, ".*fromStatementOfResponsibility:")
										corpNP:= fnr(corpNP, "firstName: text: ", "`n`n")
										corpNP:= fnr(corpNP, "nonPersonName: text:", "`n`nnonPersonName")
										corpNP:= fnd(corpNP, ".*firstName.*")
										corpNP:= fnd(corpNP, ".*nonPersonName |, romanizedText.*")
									;Cleanup
										corpNP:= creatorCleanup(corpNP)
								}
							;Secondary Native Corp
								corpNS:= fnd(corp, "^.+?contributorNotes")
								if !inStr(corpNS, "nonPersonName")
									corpNS:= "n/a"
								else{
										corpNS:= fnr(corpNS, ", romanizedText: ", "`n`n")
										corpNS:= fnd(corpNS, ".*text: |.*languageCode.*")
									;Cleanup
										corpNS:= creatorCleanup(corpNS)
								}
							;Merge primary and secondary
								if((corpNP= "n/a") & (corpNS= "n/a"))
									corpN:= "n/a"
								if((corpNP!= "n/a") & (corpNS= "n/a"))
									corpN:= corpNP
								if((corpNP= "n/a") & (corpNS!= "n/a"))
									corpN:= corpNS
								if((corpNP!= "n/a") & (corpNS!= "n/a"))
									corpN:= corpNP . " ^ " . corpNS						
							}
					;Romanized Corp ＊＊＊with＊＊＊ Romanized text in WorldCat metadata
						if inStr(corp, "romanizedText"){
							;Primary Romanized Corp
								corpRP:= fnd(corp, "contributorNotes.*")
								if !inStr(corpRP, "nonPersonName")
									corpRP:= "n/a"
								else{
										corpRP:= fnd(corp, "contributorNotes.*")
										corpRP:= fnd(corpRP, ".*fromStatementOfResponsibility:")
										corpRP:= fnr(corpRP, "firstName: text: ", "`n`n")
										corpRP:= fnr(corpRP, "nonPersonName: text:", "`n`nnonPersonName")
										corpRP:= fnd(corpRP, ".*firstName.*")
										corpRP:= fnd(corpRP, ".*romanizedText: |, languageCode.*")
									;Cleanup
										corpRP:= creatorCleanup(corpRP)
								}
									
							;Secondary Romanized Corp
								corpRS:= fnd(corp, "^.+?contributorNotes")
								if !inStr(corpRS, "nonPersonName")
									corpRS:= "n/a"
								else{
										corpRS:= fnr(corpRS, "nonPersonName", "`n`nnonPersonName")
										corpRS:= fnd(corpRS, "m)nonPersonName.+?romanizedText: |, languageCode.*|^:.*")
									;Cleanup
										corpRS:= fnd(corpRS, " \(.+?\)")
										corpRS:= creatorCleanup(corpRS)
								}
					;Romanized Corp ＊＊＊without＊＊＊ Romanized text in WorldCat metadata
						}else{
							;Primary Romanized Corp
								corpRP:= fnd(corp, "contributorNotes.*")
								if !inStr(corpRP, "nonPersonName")
									corpRP:= "n/a"
								else{
									corpRP:= fnd(corpRP, ".*text: |, isPrimary.*")
;needs work, need example entry with more than one corporate name, and corporate name in secondary									
								}
							;Secondary Romanized Corp
								corpRS:= "n/a"
						}
					;Merge primary and secondary
						if((corpRP= "n/a") & (corpRS= "n/a"))
							corpR:= "n/a"
						if((corpRP!= "n/a") & (corpRS= "n/a"))
							corpR:= corpRP
						if((corpRP= "n/a") & (corpRS!= "n/a"))
							corpR:= corpRS
						if((corpRP!= "n/a") & (corpRS!= "n/a"))
							corpR:= corpRP . " ^ " . corpRS	
				}else{
					corpN:= "n/a"
					corpR:= "n/a"
				}
		;Merge creator and corporate creator.
			;Merge Native
				if((creatorN= "n/a") & (corpN= "n/a"))
					creator:= "n/a"
				if((creatorN!= "n/a") & (corpN= "n/a"))
					creator:= creatorN
				if((creatorN= "n/a") & (corpN!= "n/a"))
					creator:= corpN
				if((creatorN!= "n/a") & (corpN!= "n/a"))
					creatorN:= creatorN . " ^ " . corpN
					creatorN:= dedupe2(creatorN)
					creatorN:= fnd(creatorN, "^ \^ ")
			;Merge Romanized
				if((creatorR= "n/a") & (corpR= "n/a"))
					creator:= "n/a"
				if((creatorR!= "n/a") & (corpR= "n/a"))
					creator:= creatorR
				if((creatorR= "n/a") & (corpR!= "n/a"))
					creator:= corpR
				if((creatorR!= "n/a") & (corpR!= "n/a"))
					creatorR:= creatorR . " ^ " . corpR
					creatorR:= dedupe2(creatorR)
					creatorR:= fnd(creatorR, "^ \^ ")
					
					
	;📚 Series Title
			seriesR:= isolate(data, "series: ", ".*series: |, seriesVolumes.*")
			seriesN:= "n/a"
			
			
	;🔢 Series Number
			if inStr(data, "seriesVolumes: null")
				vol:= "n/a"
			else
			vol:= isolate(data, "seriesVolumes: ", ".*seriesVolumes: |,.*")
		;Cleanup
			vol:= fnd(vol, "\[|\]")
			vol:= fnd(vol, "^dai |-kan")
			
			
	;🏢 Publisher
			pub:= isolate(data, "publisher: ", ".*publisherName: text: |, languageCode.*")
			if inStr(pub, "romanizedText"){
				pubR:= fnd(pub, ".*romanizedText: ")
				pubN:= fnd(pub, ".*text: |, romanizedText.*")
			}else{
				pubR:= fnd(pub, ", publicationPlace.*|, publicationDate.*")
				pubN:= "n/a"
			}
			
	;♎ Year of Publication
			year:= isolate(data, "publicationDate: ", ".*publicationDate: |, .*")
			if inStr(year, "romanizedText"){
				if !inStr(year, " ["){
					yearN:= "n/a"
					yearR:= year
				}else{
					yearN:= fnd(year, ".* \[|\].*")
					yearR:= fnd(year, " \[.*")
				}
			}else{
				yearN:= "n/a"
				yearR:= year
			}
			
			
	;📖 Edition
			editionR:= isolate(data, "edition: ", ".*edition: |, .*")
			editionN:= "n/a"
			
			
	;💡 Subjects
			subjects:= isolate(dataForSubjects, "subjectsText", ".*subjectsText: \[|\].*")
		;Cleanup
			subjects:= fnr(subjects, "`",`"", " ^ ")
			subjects:= fnd(subjects, "^`"|`"$")
			
;▼ --------------- Check Results ---------------
			checkData:=		  "ISBN-13#:`n"
								. isbn13				. "`n`n"
							. "ISBN-10#:`n"
								. isbn10				. "`n`n"
							. "OCLC#:`n"
								. oclc					. "`n`n"
							. "Language:`n"
								. language				. "`n`n"
							. "Title:`n"
								. titleR				. "`n"
								. titleN				. "`n"
								. titleT				. "`n`n"
							. "Volume:`n"
								. vol					. "`n`n"
							. "Total Volumes:`n"
								. volumes				. "`n`n"
							. "Creator(s):`n"
								. creatorR				. "`n"
								. creatorN				. "`n`n"
							. "Series Title:`n"
								. seriesR				. "`n"
								. seriesN				. "`n`n"
							. "Publisher:`n"
								. pubR					. "`n"
								. pubN					. "`n`n"
							. "Year: `n"
								. yearR					. "`n"
								. yearN					. "`n`n"
							. "Edition: `n"
								. editionR				. "`n"
								. editionN				. "`n`n"
							. "Subject(s):`n"
								. subjects				. "`n`n"
							. "Price URL: `n"
							;	. priceURL
			checkData:= decompose(checkData)
			if(checkmode= 1)
				msgBox checkData

;▼ --------------- Export data back to spreadsheet ---------------
	;▼ Put parsed data into array.
			bibArr[8]:= priceURL
			bibArr[10]:= vol
			bibArr[13]:= volumes
			bibArr[16]:= dupe
			bibArr[17]:= libCount
			bibArr[18]:= isbn13
			bibArr[19]:= isbn10
			bibArr[20]:= oclc
			bibArr[21]:= language
			bibArr[22]:= titleT
			bibArr[23]:= titleR
			bibArr[24]:= titleN
			bibArr[25]:= yearR
			bibArr[26]:= yearN
			bibArr[27]:= creatorR
			bibArr[28]:= creatorN
			bibArr[29]:= seriesR
			bibArr[30]:= seriesN
			bibArr[31]:= pubR
			bibArr[32]:= pubN
			bibArr[33]:= editionR
			bibArr[34]:= editionN
			bibArr[35]:= subjects
			inClip(bibArr[1] . "`t" . bibArr[2] . "`t" . bibArr[3] . "`t" . bibArr[4] . "`t" . bibArr[5] . "`t" . bibArr[6] . "`t" . bibArr[7] . "`t" . bibArr[8] . "`t" . bibArr[9] . "`t" . bibArr[10] . "`t" . bibArr[11] . "`t" . bibArr[12] . "`t" . bibArr[13] . "`t" . bibArr[14] . "`t" . bibArr[15] . "`t" . bibArr[16] . "`t" . bibArr[17] . "`t" . bibArr[18] . "`t" . bibArr[19] . "`t" . bibArr[20] . "`t" . bibArr[21] . "`t" . bibArr[22] . "`t" . bibArr[23] . "`t" . bibArr[24] . "`t" . bibArr[25] . "`t" . bibArr[26] . "`t" . bibArr[27] . "`t" . bibArr[28] . "`t" . bibArr[29] . "`t" . bibArr[30] . "`t" . bibArr[31] . "`t" . bibArr[32] . "`t" . bibArr[33] . "`t" . bibArr[34] . "`t" . bibArr[35])
	;▼ Close FirstSearch "Detailed Record" page.
			if WinExist(" | WorldCat.org") 
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
			Sleep nt
			Send "{right 18}"
}



;■■■ Pull data from FirstSearch record into spreadsheet.
#HotIf WinActive(" | WorldCat.org")
numpadEnter::{
		activeGUI()
		pullBibDataWC()
		active.Destroy()
}
F2::{
		activeGUI()
		pullBibDataWC()
		active.Destroy()
}
#HotIf
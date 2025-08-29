;■■■■■■■■■■■■■ Put bibliographic data from a FirstSearch detailed record and put it in a spreadsheet.
;▼▲▼ Get bibliograhic data from the source code of a WorldCat detailed record in FirstSearch.
	pullBibDataFS(){
			global bibArr
			global activeSearch
		;▼ Error check that spreadsheet and fsURL variables are not blank.
			checkGUIinputs(CD, DI, US, fsURL)
			spreadsheet:= sheetCheck()
		;▼ FirstSearch
			if !WinExist("Detailed Record") & !WinExist("Libraries that Own Item"){
				exit
				active.destroy()
			}
			WinActivate
			data:= loadSourceCode("pagename", "FirstSearch Detailed Record")
			data:= singleLine(data, 1)
			data:= RegExReplace(data, "<span class\=matchterm[0-9]>") 	;Removes code for yellow text highlighting.
			data:= RegExReplace(data, "</span>") 						;Because of highlighting above, all "</span>" are removed.
		;Determine if active search is happening
			if(activeSearch= 0){
					bibArr:= []
					Loop 35{
						bibArr.InsertAt(1, "")
						}
					bibArr[1]:= ""
					bibArr[9]:= ""
			}
;▼ --------------- Data clean up ---------------
		;📚 Total volumes of multivolume sets
			volumes:= RegExReplace(data, ".*<b>Description:|</font></td>.*")
			if !inStr(volumes, "volumes")
				volumes:= "n/a"
			else if inStr(volumes, ">volumes")
				volumes:= "n/a"
			else
				volumes:= RegExReplace(volumes, ".*>| volumes.*")
			
			
		;🏛 Check against local holdings
					if InStr(data, "FirstSearch indicates your institution"){
						dupe:= RegExReplace(data, ".*</b><LI class=OPAC>|</tr>.*")
						dupe:= RegExReplace(dupe, ".*href=`"|`".*")
						dupe:= "=hyperlink(`"" . dupe . "`",`"yes`")"
					}
					else
						dupe:= "n"
						
						
		;👬 Get count of libraries that also have this item
				if (bibArr[17]= "") | (bibArr[17]= "n/a")
						bibArr[17]:= RegExReplace(data, ".*Libraries worldwide that own item.+? |&.*|<.*")
						
						
		;🔢 ISBN
				if !InStr(data, "<b>ISBN:"){
					isbn13:= "n/a"
					isbn10:= "n/a"
				}else{
					;▼ Use ISBN input into spreadsheet instead of pulling from record
						isbnL:= strLen(bibArr[19])
						isNumbers:= RegExMatch(bibArr[19], "\d{9}")
						if(((isbnL= 13) | (isbnL= 10)) & (isNumbers= 1)){
							if(isbnL= 13){
								isbn13:= bibArr[19]
								isbn10:= convert13to10(bibArr[19])
							}
							if(isbnL= 10){
								isbn10:= bibArr[19]
								isbn13:= convert10to13(bibArr[19])
							}
					;▼ Use ISBN from record
						}else{
								isbn:= RegExReplace(data, ".*<b>ISBN:</b> |</font>.*| <b>National.*| <b>LCCN.*")
								isbn:= RegExReplace(isbn, " $|;$")
							;▼ clean up
								isbn:= RegExReplace(isbn, "\(\(", "(")		;"(("
								isbn:= RegExReplace(isbn, "\)\)", ")")		;"))"
								isbn:= RegExReplace(isbn, " <b>Other.*")
								if !inStr(isbn, "v")						;removes book info in ( )
									isbn:= RegExReplace(isbn, "\(.+?\)")
								isbn:= RegExReplace(isbn, "  ", " ")
								isbn:= RegExReplace(isbn, "v\. |v ")
								len:= strLen(isbn)
							;▼ split string if one isbn10 and one isbn13
								if(len= 25){
									iPos:= inStr(isbn, "`; ")
									if(iPos= 14){
										isbn13:= RegExReplace(isbn, "`;.*")
										isbn10:= RegExReplace(isbn, ".*`; ")
									}
									if(iPos= 11){
										isbn13:= RegExReplace(isbn, ".*`; ")
										isbn10:= RegExReplace(isbn, "`;.*")
									}
								}
							;▼ split string if more than one isbn10 and one isbn13
								else{
									isbn:= RegExReplace(isbn, " `; |`; ", " ^ ")
									;isbn13
										isbn13:= RegExReplace(isbn,   "i) \^ \d{10} \^ | \^ \d{9}X ^ ", " ^ ")  ;middle isbn
										loop 2{
											isbn13:= RegExReplace(isbn13,   "i) \^ \d{10} \^ | \^ \d{9}X ^ ", " ^ ")
										}
										isbn13:= RegExReplace(isbn13, "i)^\d{10} \^ |^\d{9}X \^ ")				;starting isbn
										isbn13:= RegExReplace(isbn13, "i) \^ \d{10}$| \^ \d{9}X$")				;ending isbn
										inClip(isbn13)
									;isbn10
										isbn10:= RegExReplace(isbn,   "i) \^ \d{13} \^ ", " ^ ")				;middle isbn
										loop 2{
											isbn10:= RegExReplace(isbn10,   "i) \^ \d{13} \^ ", " ^ ")
										}
										isbn10:= RegExReplace(isbn10, "i)^\d{13} \^ ")							;starting isbn
										isbn10:= RegExReplace(isbn10, "i) \^ \d{13}$")							;ending isbn
								}
							;▼ ISBN Cleanup
								if inStr(isbn13, "979")
									isbn10:= "n/a"
								if((isbn10= "n/a") & !inStr(isbn13, " ") & !inStr(isbn13, "979"))
									isbn10:= convert13to10(isbn13)
								;ISBN 13 Cleanup
									if !InStr(isbn13, " ^ ")
										isbn13:= RegExReplace(isbn13, " ")
								;ISBN 10 Cleanup
									if !InStr(isbn10, " ^ ")
										isbn10:= RegExReplace(isbn10, " ")
						}
				}
				
				
		;🔢 ISSN
				if inStr(data, "<b>ISSN:"){
					isbn13:= "n/a"
					issnL:= strLen(bibArr[19])
					if((issnL= 9) | (issnL= 8))
						isbn10:= bibArr[19]
					else{
						isbn10:= RegExReplace(data, ".*<b>ISSN:</b> |`; <b>.*")
					}
					volumes:= "n/a"
				}
					
					
		;🔢 OCLC#
					oclc:= RegExReplace(data, ".*<b>OCLC:</b> |<.*")
					
					
		;🗣 Language
				global language
				language:= RegExReplace(data, ".*<b>Language:.+?serif`">|<.*|&.*")


		;💰 Price URL
				priceURL:= getPriceURL(isbn10, language)

				
		;📔 Title
				;Setup
						title:= RegExReplace(data, ".*<b>Title:.+?</td>|</td>.*")
						title:= RegExReplace(title, "<", "`n<")
						title:= RegExReplace(title, "m) /$|\.$")
				;Romanized
						;titleR:= RegExReplace(title, "<br>|<.*|`n")
						titleR:= RegExReplace(title, "<br>|</div>|<.*|`n")
						titleR:= RegExReplace(titleR, "=.*")
						
					;Remove Single line and clean up
				;Native
						if !InStr(title, "=vernacular")
							titleN:= "n/a"
						else{
							titleN:= RegExReplace(title, ".*class=vernacular.*>|<.*|`n")
							titleN:= RegExReplace(titleN, " = .*| =.*")
						}
				;Translated
					if !inStr(title, " = ")
						titleT:= "n/a"
					else
						titleT:= RegExReplace(title, ".*class=vernacular.*>|<.*|.* = |`n")
						
						
		;✒ Creators
				;Person/People
					if !InStr(data, "<b>Author(s):"){
						creatorR:= "n/a"
						creatorN:= "n/a"				
					}else{
						;Setup
							creator:= RegExReplace(data, ".*<b>Author.+?</td>|</td>.*")
							creator:= RegExReplace(creator, "<", "`n<")
						;Romanized
							creatorR:= RegExReplace(creator, "m)<a href.+?>|^<.*")
							creatorR:= cleanCreator(creatorR)
						;Native
							if !InStr(creator, "=vernacular")
								creatorN:= "n/a"
							else{
								creatorN:= RegExReplace(creator, "m).*class=vernacular.*>|^<.*")
								creatorN:= cleanCreator(creatorN)
							}
					}
				;🏫 Corporate Creator
					if !InStr(data, "<b>Corp Author(s):"){
						corpR:= "n/a"
						corpN:= "n/a"
					}else{
						;Setup
							corp:= RegExReplace(data, ".*<b>Corp Author\(s\):.+?</td>|</td>.*")
							corp:= RegExReplace(corp, "<", "`n<")
						;Romanized
							corpR:= RegExReplace(corp, "m)<a href.*>|^<.*")
							corpR:= cleanCreator(corpR)
						;Native
							if !InStr(corp, "=vernacular")
								corpN:= "n/a"
							else{
								corpN:= RegExReplace(corp, "<", "`n<")
								corpN:= RegExReplace(corpN, "m).*class=vernacular.*>|^<.*")
								corpN:= cleanCreator(corpN)
							}
					}
		;🈴 Merge individual and corporate creator values.
				;Romanized
					creatorsR:= creatorR . " ^ " . corpR
					creatorsR:= RegExReplace(creatorsR, " \^ n/a|n/a \^ ")
				;Native
					creatorsN:= creatorN . " ^ " . corpN
					creatorsN:= RegExReplace(creatorsN, " \^ n/a|n/a \^ ")
							
							
		;📚 Series Title
				if !InStr(data, "<b>Series:"){
					seriesR:= "n/a"
					seriesN:= "n/a"
				}else{
					;Setup
						series:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
					;Romanized
							seriesR:= RegExReplace(series, "`; <b>Variation:.*")
							seriesR:= RegExReplace(seriesR, ".*>")
							seriesR:= RegExReplace(seriesR, " `;`;.*")
					;Native
						if !InStr(series, "=vernacular")
							seriesN:= "n/a"
						else{
							seriesN:= RegExReplace(series, "`n")
							seriesN:= RegExReplace(seriesN, "</div>.*")
							seriesN:= RegExReplace(seriesN, ".*>|( |)`;.*|( |)=.*")
							seriesN:= RegExReplace(seriesN, " $|\.$")
						}
				}
				
				
		;🔢 Series Number
				if !InStr(data, "<b>Series:")
					vol:= "n/a"
				else{
					;Setup
						seriesNo:= RegExReplace(data, ".*<b>Series:.+?</td>|</td>.*")
						if(!inStr(seriesNo, ";; ") & !inStr(seriesNo, ",; "))
							vol:= "n/a"
						else{
							;Romanized
									seriesNoR:= RegExReplace(seriesNo, "`; <b>Variation:.*")
									seriesNoR:= RegExReplace(seriesNoR, ".*`;`; |.*,`; |number |no\. |v\. |\[|\]")
								;▼ Clean up
									seriesNoR:= RegExReplace(seriesNoR, "\.(`;|<).*|-`;.*|, etc.*")
								;▼ Logic test for when series data has no number
									isText:= RegExMatch(seriesNoR, "\d")
									if(isText= 0)
										seriesNoR:= "n/a"
							;Native
								if !InStr(seriesNo, "=vernacular")
									seriesNoN:= "n/a"
								else{
										seriesNoN:= RegExReplace(seriesNo, "<", "`n<")
										seriesNoN:= RegExReplace(seriesNoN, ".*class=vernacular.*>|<.*|.* `;`; |.* `; | .*")
										seriesNoN:= RegExReplace(seriesNoN, "no\.( |)") ; rare?
										seriesNoN:= RegExReplace(seriesNoN, "m)\.$")
										seriesNoN:= deDupe(seriesNoN)
									;▼ Logic test for when series data has no number
										isText:= RegExMatch(seriesNoN, "\d")
										if(isText= 0)
											seriesNoN:= "n/a"
								}
							;Logic for which series number to use
								if((SeriesNoR!= "n/a") | (seriesNoN!= "n/a")){
									if(SeriesNoR!= "n/a")
										vol:= SeriesNoR
									else
										vol:= SeriesNoN
								}else
									vol:= "n/a"
						}
				}
				
					
		;🏢 Publisher
				if !InStr(data, "<b>Publication:"){
					pubR:= "n/a"
					pubN:= "n/a"
				}else{
					;Setup
						pub:= RegExReplace(data, ".*<b>Publication.+?</td>|</td>.*")
					;Romanized
						pubR:= RegExReplace(pub, "<", "`n<")
						pubR:= RegExReplace(pubR, ".*</div>|.*serif.>|<.*|.*: |,.*")
						pubR:= cleanCreator(pubR)
						pubR:= fixRomanizedPublisherNames(pubR)
						if(pubR= "")
							pubR:= "n/a"
					;Native
						if !InStr(pub, "=vernacular")
							pubN:= "n/a"
						else{
							pubN:= RegExReplace(pub, "</b>.*")
							pubN:= RegExReplace(pubN, "<", "`n<")
							pubN:= RegExReplace(pubN, ".*class=vernacular.*>|<.*|.* : |,.*")
							pubN:= cleanCreator(pubN)
							if(pubN= "")
								pubN:= "n/a"
						}
				}
				
				
		;♎ Year of Publication
				if !inStr(data, "<b>Year:"){
					yearR:= "n/a"
					yearN:= "n/a"
				}else{
					;Romanized
						yearR:= RegExReplace(data, ".*Year:</b>.+?serif`">|.*<nobr>|( |)<.*") ;Isolate year
						yearR:= RegExReplace(yearR, "^ | $|-$|,.*") ;cleanup, trim
					;Native
						if(language= "Japanese")
							yearN:= convertNengo(yearR) ;Function is in "diacriticsNengo.ahk"
						else
							yearN:= "n/a"
				}
				
				
		;📖 Edition
				if !InStr(data, "<b>Edition:"){
					editionR:= "n/a"
					editionN:= "n/a"
				}else{
						edition:= RegExReplace(data, ".*<b>Publication:|(\.|)</font></td>.*")
						edition:= RegExReplace(edition, "<", "`n<")
					;Romanized
						editionR:= RegExReplace(edition, ".*class=vernacular.*|</b> |<.*")
						editionR:= cleanCreator(editionR)
						if(editionR= "")
							editionR:= "n/a"
					;Native
						editionN:= RegExReplace(edition, ".* : .*")
						if !inStr(editionN, "=vernacular")
							editionN:= "n/a"
						else{
							editionN:= RegExReplace(editionN, ".*class=vernacular.*>|<.*")
							editionN:= cleanCreator(editionN)							
						}
				}
				
				
		;💡 Subjects
				if !InStr(data, "SUBJECT(S)")
					subjects:= "n/a"
				else{
					subjects:= RegExReplace(data, ".*SUBJECT\(S\)|<b>Note\(s\):.*|<b>Class.*|<b>Genre.*|<b>Responsibility.*|<b>Time.*")
					subjects:= RegExReplace(subjects, "<", "`n<")
					subjects:= RegExReplace(subjects, "<a href.+?>|.*vernacular.*>|<.*|\(8.+?\)")
					subjects:= RegExReplace(subjects, "\.`; |`; ", "`n") ;put "Descriptor" subjects on individual lines
					subjects:= RegExReplace(subjects, "m)^ | $|\.$") ;trim
					subjects:= deDupe(subjects)
					subjects:= RegExReplace(subjects, "`n", " ^ ")
					if(subjects= "")
						subjects:= "n/a"
					
				}
					
				
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
								. titleN				. "`n`n"
							. "Volume:`n"
								. vol					. "`n`n"
							. "Total Volumes:`n"
								. volumes				. "`n`n"
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
								. priceURL
		if(checkMode= 1){
			yesNo:= MsgBox(checkData, "Review Bibliographic Data", 4100)
			if(yesNo= "No")
				exit
		}

;▼ --------------- Export data back to spreadsheet ---------------		
	;▼ Put parsed data into array.
			bibArr[8]:= priceURL
			bibArr[10]:= vol
			bibArr[13]:= volumes
			bibArr[16]:= dupe
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
			bibArr[31]:= pubR
			bibArr[32]:= pubN
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
			Sleep nt
			Send "{right 18}"
}

;▼ --------------- Functions ---------------

;▼▲▼
	cleanCreator(creator){					
		;Remove dates and put each creator on new line
			creator:= StrReplace(creator, ",;", ",") ;cleanup
			creator:= RegExReplace(creator, "m)\s$") ;trim trailing spaces
			creator:= RegExReplace(creator, "i)(, | |)author(, |\. | |)|(, | |)editor(, |\. | |)|(, | |)translator(, |\. | |)|(, | |)works(, |\. | |)|(, | |)selections(, |\. | |)", "`n") ;remove role
			creator:= RegExReplace(creator, "(, |)\(\d{4}-\d{4}\)")				;", (2005-2010)" OR "(2005-2010)"
			creator:= RegExReplace(creator, "(, |)\(\d{4}-\)")					;", (2005-)" OR "(2005-)"
			creator:= RegExReplace(creator, "(, |)\d{4}-\d{4}(\.|)", "`n")		;", 2005-2010" OR "2005-2010" OR "2005-2010." - might cause issues
			creator:= RegExReplace(creator, "(, |)\d{4}-", "`n")				;", 2005-" OR "2005-"
			creator:= RegExReplace(creator, "(, |)\d{4}\?-\d{4}(\.|)", "`n")	;", 2005?-2010" OR "2005?-2010" OR "2005-2010." - might cause issues
			creator:= RegExReplace(creator, "(, |)\d{3}\?-( |)", "`n")				;", 985?-" OR "985?-" OR "985." - might cause issues
		;Remove trailing punctuation & creator designation
			creator:= Trim(creator)
			creator:= RegExReplace(creator, "m),$|\.$")
		;Remove dupes
			creator:= deDupe2(creator)
		;Single line and clean up
			creator:= RegExReplace(creator, "`n", " ^ ")
			Loop{
				if InStr(creator, "  ")
					creator:= RegExReplace(creator, "  ", " ")
				else
					break
			}
			Loop{
				if InStr(creator, " ^ ^ ")
					creator:= RegExReplace(creator, " \^ \^ ", " ^ ")
				else
					break
			}
			creator:= RegExReplace(creator, "^ \^ | \^ $")
	return creator
}



;■■■ Pull data from FirstSearch record into spreadsheet.
#HotIf WinActive("Detailed Record") | WinActive("Libraries that Own Item")
numpadEnter::{
		activeGUI()
		pullBibDataFS()
		active.Destroy()
}
F2::{
		activeGUI()
		pullBibDataFS()
		active.Destroy()
}
#HotIf
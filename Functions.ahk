#requires AutoHotkey v2.0
setTitleMatchMode 2

;global variables for sleep time
	nt:= 250	;normal time
	lt:= 3000	;load time
	wt:= 500	;switching between windows time
	ft:= 250	;find text time, for using the find text (ctrl+f) feature in browers

;variables just for function library
	stopped:= "Why did the script stop?"



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Error Message
	;▼▲▼ Error message.
		errorMsg(stopped){
			MsgBox stopped, "Why did the script stop?", 4096
}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Copying
	;▼▲▼
		inClip(data){
				A_Clipboard:=""
				sleep nt
				A_Clipboard:= data
				ClipWait(2)
				if !ClipWait(2){
					MsgBox "Copying failed", stopped, 4096
					Exit
				}
				sleep nt
				return A_Clipboard
		}
	;▼▲▼
		copy(){
				A_Clipboard:= ""
				Send "^c"
				ClipWait(2)
				if !clipWait(2){
					active.Destroy()
					MsgBox "Copying failed", stopped, 4096
					Exit
				}
				sleep nt
				return A_Clipboard
		}
	;▼▲▼
		copyAll(){
				A_Clipboard:= ""
				Send "^a"
				Sleep nt
				Send "^c"
				ClipWait(2)
				if !ClipWait(2){
					MsgBox "Copying failed", stopped, 4096
					Exit
				}
				return A_Clipboard
		}
	;▼▲▼
		copyURL(){
				Send "!d"
				Sleep nt
				copyAll()
				return A_Clipboard
		}
	;▼▲▼ Copies a row from a spredsheet and puts each cell in the row in an array.
		copyRowMakeArray(){
				Send "{esc}"
				Sleep nt
				Send "{home}"
				Sleep nt
				Send "+{space}"
				Sleep nt
				A_Clipboard:= ""
				Sleep nt
				copy()
				Sleep nt
				Send "{home}"
				Sleep nt
				return arr:= strSplit(A_Clipboard, A_tab)
		}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Find and Replace
	;▼▲▼ Isolate Text to Find and Replace
		isolate(source, confirm, cut){
			if !inStr(source, confirm)
				return "n/a"
			data:= RegExReplace(source, cut)
			if data= "null"
				return "n/a"
			else
				return data
		}

	;▼▲▼ Find and Delete (fnd) Text
		fnd(source, cut){
			return RegExReplace(source, cut)
		}

	;▼▲▼ Find and Replace (fnd) Text
		fnr(source, cut, replace){
			return RegExReplace(source, cut, replace)
		}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Text Normalizing
	;▼▲▼ Normalize text for parsing: removes \r, \n, \t, double spaces
			singleLine(txt, removeRNT:= 1){
				;▼ Remove \r \n \t
					if(removeRNT= 1)
						txt:= RegExReplace(txt, "`r|`n|`t")
					if(removeRNT= 2){
						txt:= RegExReplace(txt, "`r|`n|`t", "xtx")
						loop{
							txt:= strReplace(txt, "xtxxtx", "xtx")
							if !inStr(txt, "xtxxtx")
								break
						}
					}
				;;▼ Remove double (or more) spaces, i.e. "  "
				;	loop{
				;		txt := strReplace(txt, "  ", " ")
				;		if !inStr(txt, "  ")
				;			break
				;	}
				return txt
			}
	;▼▲▼ Make composed Romaninzed Japanese decomposed, replace some Japanese punctuation with standard ones
			decompose(txt){
				;▼ Replace composed chatacers with decomposed ones
					txt:= RegExReplace(txt, "ā", "ā")
					txt:= RegExReplace(txt, "ī", "ī")
					txt:= RegExReplace(txt, "ū", "ū")
					txt:= RegExReplace(txt, "ē", "ē")
					txt:= RegExReplace(txt, "ō", "ō")
				;▼ Replace CJK Punctuation
					txt:= StrReplace(txt, "、", ",")
					txt:= StrReplace(txt, "：", ":")
					txt:= StrReplace(txt, "；", ";")
					return txt
			}
	;▼▲▼ DeDuping
			deDupe(txt){
				txt := StrReplace(txt, "`r")
				str := ''
				loop parse, txt, '`n'					;Parse by LF
					if !InStr(str, A_LoopField '`n')	;If str does not contain a matching line
						str .= A_LoopField '`n'			;Add line and LF
				str:= RTrim(str, '`n')					;Snip off extra LF and return
				return str:= LTrim(str, '`n')			;Snip off extra LF and return
			}

dedupe2(txt) {
    txt:= RegExReplace(txt, " \^ ", "^")
    str := ""								; Str to return
    loop parse, txt, "^"					; Parse by LF
        if !InStr(str, A_LoopField "`n")	; If str does not contain a matching line
            str .= A_LoopField "`n"			; Add line and LF
    txt:= RTrim(str, "`n")					; Snip off extra LF and return
    return RegExReplace(txt, "`n", " ^ ")
}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ GUI Management
	;▼▲▼ GUI alerting that script is active.
		activeGUI(){
			global active
			active:= Gui("alwaysOnTop +ToolWindow", "Script is Active")
			active.SetFont("s16")
			active.Add("Text", "x35 y85", "❌")
			active.Add("Text", "x55 y85", "⌨")
			active.SetFont("s9")
			active.Add("Text", "x85 y93", "Do not use your keyboard")
			active.SetFont("s16")
			active.Add("Text", "x35 y110", "❌")
			active.Add("Text", "x58 y110", "🖱")
			active.SetFont("s9")
			active.Add("Text", "x85 y118", "Do not use your mouse")
			active.Show("w250 h250 NoActivate")
			sleep wt
		}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Websites / Browsers / Programs existing or not.
	;▼▲▼
		confirmBrowserOpen(){
			if !winExist("ahk_exe firefox.exe") & !winExist("ahk_exe chrome.exe") & !winExist("ahk_exe msedge.exe"){
				msg:= "You have no browser open on your computer. Please open one of the following:`n"
					. "`n▶ FireFox"
					. "`n▶ Edge"
					. "`n▶ Chrome"
					. "`n`nRun the macro again after you`'ve opened a browser."
				MsgBox(msg,, 4096)
				Exit
			}
		}
	;▼▲▼
		activateBrowser(browserName:= ""){
			if winExist("ahk_exe firefox.exe"){
				browserName:= "FireFox"
				WinActivate
				Sleep wt
				return browserName
			}
			if winExist("ahk_exe chrome.exe"){
				browserName:= "Chrome"
				WinActivate
				Sleep wt
				return browserName
			}
			if winExist("ahk_exe msedge.exe"){
				browserName:= "Edge"
				WinActivate
				Sleep wt
				return browserName
			}
			msg:= "You have no browser open on your computer. Please open one of the following:`n"
			active.Destroy()
				. "`n▶ FireFox"
				. "`n▶ Edge"
				. "`n▶ Chrome"
				. "`n`nRun the macro again after you`'ve opened a browser."
			MsgBox(msg,, 4096)
			Exit
		}
	;▼▲▼
		doesTabExist(inTabName){
			if WinExist(inTabName){
				WinActivate
				Sleep 500
			}
			else{
				msg:= "The active tab in your broswer does not have `"" . inTabName . "`" in the title.`n"
				. "`nYou may have a different tab in your browser active. If so, activate that tab and run the script again.`n"
				. "`nIf you only have one tab open, then you do not have the correct website open to run this script."
				active.Destroy()
				MsgBox msg, stopped, 4096
			Exit
			}
		}
	;▼▲▼
		dataHere(title){
			if WinExist(title){
				WinActivate
				Sleep wt
			}
			else{
				msg:= "There is no program or active browser tab with `"" . title . "`" in the title.`n"
					. "`nOpen the program with the file where you are storing this data."
					. "`n***OR***"
					. "`nOpen the web based app where you are storing this data in a *NEW* browser window and make sure that it is the active tab in the browser."
				MsgBox msg, stopped, 4096
			}
		}
	;▼▲▼
		loadSiteWithSearch(searchPrefix, searchContent){
			Send "!d"
			Sleep nt
			Send searchPrefix
			Sleep nt
			Send searchContent
			Sleep nt
			Send "{enter}"
			Sleep 4000
		}
	;▼▲▼
		loadCheck(text, pageName){
			Sleep lt
			Loop 3{
				copyAll()
				if InStr(A_Clipboard, "text")
					break
				Sleep nt*2
			}
			if !InStr(A_Clipboard, text, false){
				msg:= "It appears the web page " . pageName . " didn't load, or loaded too slowly.`n`nYou can try running the script again."
				active.Destroy()
				MsgBox msg, stopped, 4096
				Exit
			}
		}
	;▼▲▼
		findTextOnSite(text){
			Send "^f"
			Sleep ft
			Send text
			sleep ft
			Send "{esc}"
			Sleep ft
		}
	;▼▲▼
		loadSourceCode(text, pageName:= ""){
			Send "^u"
			Sleep lt*.6
			Loop {
				A_Clipboard:= ""
				Sleep 250
				Send "^a"
				Sleep 100
				Send "^c"
				Sleep 100
				if InStr(A_Clipboard, text)
					break
				if A_Index= 3{
					MsgBox("It appears the source code for `"" . pageName . "`" didn't load, or loaded too slowly.`n`nYou can try running the script again.", stopped, 4096)
					A_Clipboard:= ""
					exit
				sleep 1000
				}
			}
			Send "^w"
			return A_Clipboard
		}
	;▼▲▼
		newTab(url){
			Send "^t"
			Sleep nt
			Send "!d"
			Sleep nt*2
			Send "{delete}"
			Sleep nt
			inClip(url)
			Send "^v"
			Sleep nt
			Send "{enter}"
			Sleep wt
		}
	;▼▲▼
		newWin(url){
			inClip(url)
			Send "^n"
			Sleep wt
			Send "!d"
			Sleep nt
			Send "^a"
			Sleep nt
			Send "{delete}"
			Sleep nt
			Send "^v"
			Sleep nt
			Send "{enter}"
			Sleep nt
		}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Convert ISBNs
	;▼▲▼ Convert ISBN13 to ISBN10
		convert13to10(isbn){
			;▼ Normalize ISBN13
				isbn:= RegExReplace(isbn, "-|`r|`n| ")	;remove hypens and spaces to make sure string is 13 long
			;▼ Verify string is ISBN13
				;▼ Length
					len:= strLen(isbn)
					if(len!= 13){
						errorMsg("This is not a string of 13 digits.")
						exit
					}
				;▼ All digits
					digitCount:= RegExMatch(isbn, "\d{13}")
					if(digitCount= 0){
						errorMsg("This is not a string of 13 digits.")
						exit
					}
				;▼ No 979. ISBN13s starting with 979 have no ISBN10
					if inStr(isbn, "979"){
						return "n/a"
					}
			;▼ Calculate ISBN10
				isbn:= RegExReplace(isbn, "^978")
				s:= StrSplit(isbn) ;s = split
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
				if(checkNo= 11)
					checkNo:= "0"
				if(checkNo= 10)
					checkNo:= "X"
				isbn10:= s[1] . s[2] . s[3] . s[4] . s[5] . s[6] . s[7] . s[8] . s[9] . checkNo
				return isbn10
		}

	;▼▲▼ Convert ISBN10 to ISBN13
		convert10to13(isbn){
			;▼ Normalize ISBN10
				isbn:= RegExReplace(isbn, "-| ")	;remove hypens and spaces to make sure string is 13 long
			;▼ Verify string is ISBN10
				;▼ Length
					len:= strLen(isbn)				
					if(len!= 10){
						errorMsg("This is not a string of 10 digits.")
						exit
					}
				;▼ All digits
					digitCount:= RegExMatch(isbn, "\d{10}")
					if(digitCount= 0){
						digitCount:= RegExMatch(isbn, "i)\d{9}X")
						if(digitCount:= 0){
							errorMsg("This is not a string of 10 digits.")
							exit
						}
					}
			;▼ Calculate ISBN13
				isbn:= RegExReplace(isbn, ".$")
				isbn:= "978" . isbn
				s:= StrSplit(isbn) ;s = split
				a:= s[1]*1
				b:= s[2]*3
				c:= s[3]*1
				d:= s[4]*3
				e:= s[5]*1
				f:= s[6]*3
				g:= s[7]*1
				h:= s[8]*3
				i:= s[9]*1
				j:= s[10]*3
				k:= s[11]*1
				l:= s[12]*3
				sum:= a+b+c+d+e+f+g+h+i+j+k+l
				remainder:= mod(sum, 10)
				if(remainder= 0)
					checkNo:= "0"
				else
					checkNo:= 10-remainder
				isbn13:= s[1] . s[2] . s[3] . s[4] . s[5] . s[6] . s[7] . s[8] . s[9] . s[10] . s[11] . s[12] . checkNo
				return isbn13
		}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Inputing AHK code
#HotIf WinActive("ahk_class Notepad++")
	
	;▼▲▼ Sleep Commands
	::snt::{raw}Sleep nt
	::slt::{raw}Sleep lt
	::swt::{raw}Sleep wt
	::sft::{raw}Sleep ft

	;▼▲▼ Keys with {} in AHK code
	::ssa::{raw}Send "^a"
	::sse::{raw}Send "{enter}"
	::ssb::{raw}Send "{backspace}"
	::ssp::{raw}Send "{space}"
	::sst::{raw}Send "{tab}"
	::ssesc::{raw}Send "{esc}"

	;▼▲▼ Up Down Left Right and other directions
	::ssu::{raw}Send "{up}"
	::ssd::{raw}Send "{down}"
	::ssl::{raw}Send "{left}"
	::ssr::{raw}Send "{right}"
	::ssh::{raw}Send "{home}"

	;▼▲▼ Windows Commands
	::ssc::{raw}Send "^c"
	::ssn::{raw}Send "^n"
	::ssv::{raw}Send "^v"
	::sdl::{raw}Send "{delete}"

	;▼▲▼ Browser Commands
	::ssf::{raw}Send "^f"

	;▼▲▼ AutoHotKey Commands
	::clip::{raw}A_Clipboard
	::slp::{raw}Sleep 250
	::iwe::{raw}if WinExist(" "
	
	::msx::MsgBox
	
	::reg::
		{
		Send "RegExReplace(, `"`")"
		Sleep nt
		Send "+{tab 2}"
		Sleep nt
		Send "{right}"
		}

	;▼▲▼ IniFile Commands
	::iniw::
		{
		Send "IniWrite(`"`", `"bibData.ini`", `"`", `"`")"
		Sleep nt
		Send "+{tab 6}"
		Sleep nt
		Send "{left 2}"
		}
	::inir::
		{
		Send "IniRead(`"bibData.ini`", `"`", `"`")"
		Sleep nt
		Send "+{tab 1}"
		Sleep nt
		Send "{left 3}"
		}

	;▼▲▼ activeGUI
		::acg::{raw}activeGUI()
		::acd::{raw}active.Destroy()

	;▼▲▼ GUI prompts
		::gat::{raw}.Add("Text",,)
#HotIf



;▼▲▼ Faster typing for certain charcters
	
	;Triangles
	::dtri::▼
	::utri::▲
	::ltri::◀
	::rtri::▶
	
	;Japanese
	::jo::上
	::chu::中
	::ge::下
	
	;Chinese
	::shang::上
	::zhong::中
	::xia::下

;▼▲▼▲▼▲▼▲▼▲▼▲▼ Bibliographic data functions
;▼▲▼
	getPriceURL(isbn10, language){			
		if(language= "Japanese") & (isbn10!= "") & (isbn10!= "n/a")
			url:= "https://www.amazon.co.jp/dp/" . isbn10
		else if(language="English") & (isbn10!= "") & (isbn10!= "n/a")
			url:= "http://www.amazon.com/dp/" . isbn10
		else
			url:= "n/a"
		trim(url)
		if inStr(isbn10, A_space)
			url:= "n/a"
	return url
}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Script reload, shutdown, and pause.
	^\::{
		MsgBox A_ScriptName " is reloading.", "Reloading Script", "T2"
		Reload
	}

	^+\::{
		MsgBox A_ScriptName " is shutting off", "Shutting Script Off", "T2"
		ExitApp
	}
	#SuspendExempt 
	Pause::Suspend(-1)
	F12::Suspend(-1)
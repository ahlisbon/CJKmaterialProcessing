#Requires AutoHotkey v2.0
SetTitleMatchMode 2

;global variables 
	;browserName:= ""
	;inTabName:= ""

;global variables for sleep time
	nt:= 250	;normal time
	lt:= 3000	;load time
	wt:= 500	;switching between windows time
	ft:= 250	;find text time, for using the find text (ctrl+f) feature in browers

;variables just for function library
	stopped:= "Why did the script stop?"



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



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Error Message
	;▼▲▼ Error message.
		errorMsg(stopped){
			MsgBox "Why did the script stop?", stopped, 4096
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



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Tutorial GUIs
	;▼▲▼ GUI explaining what to do next.
		;▼▲▼ Establish Tutorial GUI
			tutorialGUI(){
				if(tutorialMode= 1){
					global tutorial
					tutorial:= GUI("alwaysOnTop +ToolWindow", "Tutorial Mode")
				}
}
		;▼▲▼ Create GUI Content, Repeatable
			tutorialContent(firstS, option, onWindow, keys, action){
				global tutorial
				if(firstS= 1){
					tutorial.Add("text",, 					"On this window/program: " . onWindow)
					tutorial.Add("text",	"section y+25",	"Press Hotkey:")
					tutorial.Add("text",, 			 		"In order to:")
					tutorial.Add("text",	"ys",			keys)
					tutorial.Add("text",,					action)
				}
				if(firstS= 0){
					tutorial.Add("text",	"section xs y+25",	option)
					tutorial.Add("text",	"section xs y+25",	"Press Hotkey:")
					tutorial.Add("text",, 						"In order to:")
					tutorial.Add("text",	"ys",				keys)
					tutorial.Add("text",,						action)
				}
				if(firstS= 2){
					tutorial.Add("text",	"section xs y+25",	option)
					tutorial.Add("text",, 						"On this window/program: " . onWindow)
					tutorial.Add("text",	"section xs y+25",	"Do This:")
					tutorial.Add("text",, 						"In order to:")
					tutorial.Add("text",	"ys",				keys)
					tutorial.Add("text",,						action)
				}
}
		;▼▲▼ Render GUI
			tutorialShow(){
				global tutorial
				tutorial.Show()
				Send "!{esc}"
				Sleep wt
}	
		;▼▲▼ Turn off tutorial GUI if present.
			tutorialOff(){
				global tutorialMode
				global tutorial
				if(tutorialMode= 1)
					tutorial.Destroy
				
}



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Websites / Browsers / Programs existing or not.
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
		Sleep lt*.5
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
	
	;▼▲▼ tutorialGUI
	::tg::{raw}tutorialGUI()
	::tc::{raw}tutorialContent(
	::toff::{raw}tutorialOff()

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
#Requires AutoHotkey v2.0
setTitleMatchMode 2
#Include "%A_ScriptDir%\..\Functions.ahk"

tutorialmode:= 0

;■■■■■■■■■■■■■ Keyboard commands for typing Japanese vowels in Romaji with macrons.
	>!a::  Send "ā"
	>!i::  Send "ī"
	>!u::  Send "ū"
	>!e::  Send "ē"
	>!o::  Send "ō"
	>+!a:: Send "Ā"
	>+!i:: Send "Ī"
	>+!u:: Send "Ū"
	>+!e:: Send "Ē"
	>+!o:: Send "Ō"
	>!y::  Send "¥"
	
	::aaa::
		{
		Send "ā"
		}
	::iii::
		{
		Send "ī"
		}
	::uuu::
		{
		Send "ū"
		}
	::eee::
		{
		Send "ē"
		}
	::ooo::
		{
		Send "ō"
		}
	::AAA::
		{
		Send "Ā"
		}
	::III::
		{
		Send "Ī"
		}
	::UUU::
		{
		Send "Ū"
		}
	::EEE::
		{
		Send "Ē"
		}
	::OOO::
		{
		Send "Ō"
		}
	::yyy::
		{
		Send "¥"
		}

;■■■■■■■■■■■■■ Keyboard command for typing Korean vowels in Romanization with accent marks.
	>^o::  Send "ŏ"
	>^u::  Send "ŭ"
	>+^o:: Send "Ŏ"
	>+^u:: Send "Ŭ"
	>^w::  Send "₩"
	
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
	convertNengo(year){
		;Convert Nengō to Western years.
			RegExReplace(year, " ", "")
			;Meiji era
				if InStr(year, "m", 0) | InStr(year, "meiji", 0) | InStr(year, "明治"){
					year:= RegExReplace(year, "i)meiji|m|明治")
					if year > 45{
						MsgBox "Error. The Meiji era ends in Meiji 45.", stopped, 4096
						exit year
					}else{
						year:= year + 1867
						return year
					}
				}
			;Taishō Era
				if InStr(year, "t", 0) | InStr(year, "taisho", 0) | InStr(year, "taishō") | InStr(year, "大正"){
					year:= RegExReplace(year, "i)taishō|taisho|t|大正")
					if (year > 15){
						MsgBox "Error. The Taishō era ends in Taishō 15.", stopped, 4096
						exit year
					}else{
						year:= year + 1911
						return year
					}
				}
			;Shōwa Era
				if RegExMatch(year, "i)^s") | InStr(year, "showa", 0) | InStr(year, "shōwa", 0) | InStr(year, "昭和", 0){
					year:= RegExReplace(year, "i)shōwa|showa|s|昭和")
					if (year > 64){
						MsgBox "Error. The Shōwa era ends in Shōwa 64.", stopped, 4096
						exit
					}else{
						year:= year + 1925
						return year
					}
				}
			;Heisei Era
				if InStr(year, "h", 0) | InStr(year, "heisei", 0)  | InStr(year, "heiei", 0) | InStr(year, "平成", 0){
					year:= RegExReplace(year, "i)heisei|h|平成")
					if (year > 31){
						MsgBox "Error. The Heisei era ends in Heisei 31.", stopped, 4096
					}else{
						year:= year + 1988
						return year
					}
				}
			;Reiwa Era
				if InStr(year, "r", 0) | InStr(year, "reiwa", 0) | InStr(year, "令和", 0){
					year:= RegExReplace(year, "i)reiwa|r|令和")
					year:= year + 2018
					return year
				}

	;Check if Western year can be converted.
			if !isNumber(year){
				year:= "n/a"
				return year
				}
			if(year < 1868){
				year:= "n/a"
				return
				}
			switch year{
				case 1868:
					msgBox("Caution!`n`n1868 is a year when the imperial era changes.`n`nYou need to know the month and day of the date you are converting:`n`n--on or before September 7, 1868 is Keiō 4`n--on or after September 8, 1868 is Meiji 1`n`nThis script does not convert years when the imperial era changes.", "Dates in this year can't be converted.", 4096)
					year:= "n/a"
					return year
				case 1912:
					msgBox("Caution!`n`n1912 is a year when the imperial era changes.`n`nYou need to know the month and day of the date you are converting:`n`n--on or before June 11, 1912 is Meiji 45`n--on or after June 12, 1868 is Taishō 1`n`nThis script does not convert years when the imperial era changes.", "Dates in this year can't be converted.", 4096)
					year:= "n/a"
					return year
				case 1926:
					msgBox("Caution!`n`n1926 is a year when the imperial era changes.`n`nYou need to know the month and day of the date you are converting:`n`n--on or before December 26, 1926 is Taishō 15`n--on or after December 25, 1926 is Shōwa 1`n`nThis script does not convert years when the imperial era changes.", "Dates in this year can't be converted.", 4096)
					year:= "n/a"
					return year
				case 1989:
					msgBox("Caution!`n`n1989 is a year when the imperial era changes.`n`nYou need to know the month and day of the date you are converting:`n`n--on or before January 7, 1989 is Shōwa 64`n--on or after January 8, 1989 is Shōwa 1`n`nThis script does not convert years when the imperial era changes.", "Dates in this year can't be converted.", 4096)
					year:= "n/a"
					return year
				case 2019:
					msgBox("Caution!`n`n2019 is a year when the imperial era changes.`n`nYou need to know the month and day of the date you are converting:`n`n--on or before April 30, 2019 is Heisei 31`n--on or after May 1, 2019 is Reiwa 1`n`nThis script does not convert years when the imperial era changes.", "Dates in this year can't be converted.", 4096)
					year:= "n/a"
					return year
				}
	;Convert from Western years to Nengō.
		;Meiji Era
			if(year > 1868) & (year < 1912){
				year:= year - 1867
				return "Meiji " . year
				}
		;Taishō Era
			if(year > 1912) & (year < 1926){
				year:= year - 1911
				return "Taishō " . year
				}
		;Shōwa Era
			if(year > 1926) & (year < 1989){
				year:= year - 1925
				return "Shōwa " . year
				}
		;Heisei Era
			if(year > 1989) & (year < 2019){
				year:= year - 1988
				return "Heisei " . year
				}
		;Reiwa Era
			if(year > 2018) & (year < 2030){
				year:= year - 2018
				return "Reiwa " . year
			}
}
;■■■■■■■■■■■■■ Convert Western or Japanese years to the other.
^+y::{
	;Get year to convert.
		if WinActive("ahk_exe EXCEL.EXE") | WinActive(" - Google Sheets") | WinActive(".xls"){
			year:= copy()
			year:= RegExReplace(year, "`r`n$")
		}else{
			Send "^+{left}"
			Sleep nt
			year:= copy()
		}
		year:= RegExReplace(year, "０", "0")
		year:= RegExReplace(year, "１", "1")
		year:= RegExReplace(year, "２", "2")
		year:= RegExReplace(year, "３", "3")
		year:= RegExReplace(year, "４", "4")
		year:= RegExReplace(year, "５", "5")
		year:= RegExReplace(year, "６", "6")
		year:= RegExReplace(year, "７", "7")
		year:= RegExReplace(year, "８", "8")
		year:= RegExReplace(year, "９", "9")
	;Convert from Nengō to Western year.
		year:= convertNengo(year)
		Sleep nt
		Send year
		Sleep nt
		if WinActive("ahk_exe EXCEL.EXE") | WinActive(" - Google Sheets") | WinActive(".xls")
			Send "{enter}"
}
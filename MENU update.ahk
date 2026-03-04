;Created by Adam H. Lisbon
;Associate Professor - Japanese and Korean Studies Librarian
;University Libraries
;University of Colorado Boulder
;adam.lisbon@colorado.edu

#Requires AutoHotkey v2.0
setTitleMatchMode 2

#Include "%A_scriptDir%\..\Functions.ahk"
#Include "%A_ScriptDir%\Diacritics And Nengo.ahk"
#Include "%A_ScriptDir%\Fix Japanese Publisher Names.ahk"



;■■■■■■■■■■■■■ Global Variables
	global bibArr:= []
	global active
	global activeSearch:= 0
	global tutorialMode:= 0

;■■■■■■■■■■■■■ Read values in .ini file
		CD:=			iniRead("CJK Material Processing - Settings.ini", "Sheet Names", "CD") ;CD = Collection Development
		DI:=			iniRead("CJK Material Processing - Settings.ini", "Sheet Names", "DI") ;DI = Donation Intake
		US:=			iniRead("CJK Material Processing - Settings.ini", "Sheet Names", "US") ;US = User Selects
		useFS:=			iniRead("CJK Material Processing - Settings.ini", "Search Method", "useFS") ;FS = FirstSearch
		useWC:=			iniRead("CJK Material Processing - Settings.ini", "Search Method", "useWC") ; WC = WorldCat
		fsURL:=			iniRead("CJK Material Processing - Settings.ini", "Settings", "fsURL")
		libName:= 		iniRead("CJK Material Processing - Settings.ini", "Settings", "libName")
		checkMode:=		iniRead("CJK Material Processing - Settings.ini", "Settings", "checkMode")

;■■■■■■■■■■■■■ New GUI
	; GUI Interface
		bib:= Gui(, "MENU UPDATE")
		tab:= bib.Add("Tab3",, ["General","Chinese","Japanese","Korean","English","Prices"])
		tab.UseTab("General")
			;File Name Section
				bib.Add("GroupBox",	"	 									w495	h127"	, "File Name Prefixes")
			;File Name Prefix Clarifier
				bib.Add("Text",	"							xp183	yp+15"					,"▼ File Name Prefixes for Your Spreadsheets (Case Sensitive)")
			;File Name Prefix Labels
				bib.Add("Text", "	 vCD		Section	x33			y+15"					,"Collection Development:")
				bib.Add("Text", "	 vDI							y+15"					,"Donation Intake:")
				bib.Add("Text", "	 vUS							y+15"					,"User Selection List:")
			;File Name Text Boxes
				bib.Add("Edit", "				x205				ys-4	w300")
				bib.Add("Edit", "											w300")
				bib.Add("Edit", "											w300")
			;FirstSearch Options
				bib.Add("Text", "				Section 	x33		y+40"					,"FirstSearch URL for your institution:")
				bib.Add("Edit",	"		 					x205	ys-4	w300")
			;WorldCat Options
				bib.Add("Text", "				Section 	x33		y+40"					,"Your institution's name as it appears on WorldCat")
				bib.Add("Edit",	"		 					x205	ys-4	w300")
			;Check Mode Section
				bib.Add("Text",	"				Section		x33		y+26"					,"Use &Check Mode:")
				bib.Add("Checkbox", "vcheckMode				x205	ys")
				bib.Add("Text", "							x223 	ys"						,"Review data before it's imported into your spreadsheet.")
			;Load Time Section
				bib.Add("Text", "				Section		x33		y+20"					,"&Wait longer for websites to load:")
				bib.Add("DDL","		 vloadTime				x205 	ys-4	w30	Choose1"	,["1", "2", "3"])
			;Save Settings	
				bib.Add("Button", "	 default	Section		x205	y+20"					, "&Save Settings").OnEvent("Click", settings)				
			;Get Help
				bib.Add("Link","				Section		x350  	y+15"					,"Read the <a href=`"https://github.com/ahlisbon/CJKmaterialProcessing#-hotkeys-to-activate-macro`">Hotkey Guide</a> on GitHub")		
		tab.UseTab("Chinese")
				bib.Add("GroupBox",	"	 									w150	h200"	, "Content Pending")
		tab.UseTab("Japanese")
				bib.Add("GroupBox",	"	 		Section						w150	h200"	, "Price Checking")
				bib.Add("Checkbox", "vjpFHY					xp+10	yp-30"					, "Furuhonya")
				bib.Add("Checkbox", "vjpJPT							yp+55"					, "JPT")
				bib.Add("Checkbox", "vjpAZJ							y+13"					, "Amazon.jp")
				bib.Add("Checkbox", "vjpAZU							y+13"					, "Amazon.com")
				bib.Add("GroupBox",	"	 		Section		xp150	yp-77	w150	h200"	, "Translation")
				bib.Add("Radio", "vjTranslate				xp+10	yp+25"					, "Google")
				bib.Add("Radio", "				Section				y+13"					, "ChatGPT *")
		tab.UseTab("Korean")
				bib.Add("GroupBox",	"	 									w150	h200"	, "Content Pending")
		tab.UseTab("English")
				bib.Add("GroupBox",	"	 									w150	h200"	, "Content Pending")
		tab.UseTab("Prices")
				bib.Add("GroupBox",	"	 									w150	h200"	, "Content Pending")
			
		bib.Show




;■■■■■■■■■■■■■ Old GUI
;	; GUI Interface
;		bib:= Gui(, "MENU UPDATE")
;		
;	;Question 1:
;		bib.Add("Text",		"					x180 y20",	"▼ File Name Prefixes of Your Spreadsheets (Case Sensitive)")
;		bib.Add("Link",		"					x192 y40",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----file-name-prefixes`">Read about file naming conventions</a>")
;		;Q1 answer 1
;			bib.Add("Text",	"					x10  y65",	"Collection Development:")
;			bib.Add("Edit",	"vCD		w300	x180 y60",	CD)
;		;Q1 answer 2
;			bib.Add("Text",	"					x10  y95",	"Donation Intake:")
;			bib.Add("Edit",	"vDI		w300	x180 y90",	DI)
;		;Q1 answer 3
;			bib.Add("Text",	"					x10  y125",	"Users Select Materials:")
;			bib.Add("Edit",	"vUS		w300	x180 y120",	US)
;	;Question 3:
;		bib.Add("Text",		"					x10	 y170",	"FirstSearch URL for your institution:")
;		bib.Add("Edit",		"vfsURL		w300	x180 y165",	fsURL)
;	;Question 4:
;		bib.Add("Text", "						x10	 y200", "Use &Check Mode:")
;		bib.Add("Checkbox",	"vcheckMode			x180 y200")
;		bib.Add("Text", "						x220 y200", "Review data before it's imported into your spreadsheet.")
;	;Question 5:
;		bib.Add("Text",		"					x10  y260", "&Wait longer for websites to load:")
;		bib.Add("DDL", 		"vloadTime	w30		x180 y255 Choose1", ["1", "2", "3"])
;		bib.Add("Link",		"					x220 y260",	"<a href=`"https://github.com/ahlisbon/CJKmaterialProcessing/blob/master/README.md#----wait-longer-for-websites-to-load`">What is this?</a>")
;	;Process answers into variables
;		bib.Add("Button",	"default			x180 y300", "&Save Settings").OnEvent("Click", settings)
;	;Help Text/Link
;		bib.Add("Link", 	"					x10  y310", "Read the <a href=`"https://github.com/ahlisbon/CJKmaterialProcessing#-hotkeys-to-activate-macro`">Hotkey Guide</a> on GitHub")
;		bib.Show()
;	;Save and error check inputs
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
}
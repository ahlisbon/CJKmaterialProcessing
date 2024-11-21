;Created by Adam H. Lisbon
;Associate Professor - Japanese and Korean Studies Librarian
;University Libraries
;University of Colorado Boulder
;adam.lisbon@colorado.edu

#Requires AutoHotkey v2.0
setTitleMatchMode 2
#Include "%A_ScriptDir%\Functions.ahk"



;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions for cleaning FirstSearch bibliographic data.
fixRomanizedPublisherNames(pub){
	;List of Japanese Publishers without spaces in their names.
		;a	
			pub:= regExReplace(pub, "i)akashishoten", "Akashi Shoten")
			pub:= regExReplace(pub, "i)asahishinbunsha", "Asahi Shinbunsha")
			pub:= regExReplace(pub, "i)asahishinbunshuppan", "Asahi Shinbun Shuppan")
		;b	
			pub:= regExReplace(pub, "i)benseishuppan", "Bensei Shuppan")
			pub:= regExReplace(pub, "i)bungeishunju", "Bungei Shunjū")
			pub:= regExReplace(pub, "i)bunkarongakkai", "Bunkaron Gakkai")
			pub:= regExReplace(pub, "i)bunkashobohakubunsha", "Bunka Shobō Hakubunsha")
		;c	
			pub:= regExReplace(pub, "i)chikumashobo", "Chikuma Shobō")
			pub:= regExReplace(pub, "i)chikurashobo", "Chikura Shobō")
			pub:= regExReplace(pub, "i)chosakukenjohosenta", "Chosakuken Jōhō Sentā")
			pub:= regExReplace(pub, "i)chunichieigasha", "Chūnichi Eigasha")
			pub:= regExReplace(pub, "i)chuodaigakushuppanbu", "Chūō Daigaku Shuppanbu")
			pub:= regExReplace(pub, "i)chuokeizaisha", "Chūō Keizaisha")
			pub:= regExReplace(pub, "i)chuokoronbijutsushuppan", "Chūō Kōron Bijutsu Shuppan")
			pub:= regExReplace(pub, "i)chuokoronsha", "Chūō Kōronsha")
			pub:= regExReplace(pub, "i)chuokoronshinsha", "Chūō Kōron Shinsha")
		;d	
			pub:= regExReplace(pub, "i)daiichihoki", "Daiichi Hōki")
			pub:= regExReplace(pub, "i)daiyamondosha", "Daiyamondosha")
		;e	
			pub:= regExReplace(pub, "i)enueichikeshuppan", "NHK Shuppan")
			pub:= regExReplace(pub, "i)esubikurieitibu", "Esubi Kurieitibu")
		;f	
			pub:= regExReplace(pub, "i)firumuatosha", "Firumu Āto Sha")
			pub:= regExReplace(pub, "i)fujiwarashoten", "Fujiwara Shoten")
			pub:= regExReplace(pub, "i)futamishobo", "Futami Shobo")
		;g	
			pub:= regExReplace(pub, "i)gakuyoshobo", "Gakuyō Shobō")
			pub:= regExReplace(pub, "i)gendaijinbunsha", "Gendai Jinbunsha")
			pub:= regExReplace(pub, "i)gendaishiryōshuppan", "Gendai Shiryō Shuppan")
			pub:= regExReplace(pub, "i)gendaishokan", "Gendai Shokan")
			pub:= regExReplace(pub, "i)genkishobo", "Genki Shobō")
			pub:= regExReplace(pub, "i)gentosha", "Gentōsha")
		;h	
			pub:= regExReplace(pub, "i)hayakawashobo", "Hayakawa Shobō")
			pub:= regExReplace(pub, "i)hitsujishobo", "Hitsuji Shobō")
			pub:= regExReplace(pub, "i)horeishuppan", "Hōrei Shuppan")
			pub:= regExReplace(pub, "i)horitsubunkasha", "Hōritsu Bunkasha")
			pub:= regExReplace(pub, "i)hojodoshuppan", "Hojodo Shuppan")
			pub:= regExReplace(pub, "i)honnoizumisha", "Hon no Izumisha")
			pub:= regExReplace(pub, "i)honnozasshisha", "Hon no Zasshisha")
			pub:= regExReplace(pub, "i)hoseidaigakushuppankyoku", "Hōsei Daigaku Shuppankyoku")
			pub:= regExReplace(pub, "i)hozokan", "Hōzōkan")
		;i	
			pub:= regExReplace(pub, "i)isutopuresu", "Isuto Puresu")
			pub:= regExReplace(pub, "i)iwanamishoten", "Iwanami Shoten")
			pub:= regExReplace(pub, "i)izumishoin", "Izumi Shoin")
		;j	
			pub:= regExReplace(pub, "i)jinbunshoin", "Jinbun Shoin")
			pub:= regExReplace(pub, "i)jiyukokuminsha", "Jiyū Kokuminsha")
			pub:= regExReplace(pub, "i)jiritsushobo", "Jiritsu Shobō")
			pub:= regExReplace(pub, "i)jurosha", "Jurōsha")
		;k	
			pub:= regExReplace(pub, "i)kadokawashoten", "Kadokawa Shoten")
			pub:= regExReplace(pub, "i)kaishahokankeihōmushorei", "Kaishaho Kankei Hōmu Shorei")
			pub:= regExReplace(pub, "i)kadokawagurupuhorudingusu", "Kadokawa Gurūpu Hōrudingusu")
			pub:= regExReplace(pub, "i)kadokawagurupupaburisshingu", "Kadokawa Gurūpu Paburishhingu")
			pub:= regExReplace(pub, "i)kanaeshobo", "Kanae Shobō")
			pub:= regExReplace(pub, "i)kanagawashinbunsha", "Kanagawa Shinbunsha")
			pub:= regExReplace(pub, "i)kasamashoin", "Kasama Shoin")
			pub:= regExReplace(pub, "i)kawadeshoboshinsha", "Kawade Shobō Shinsha")
			pub:= regExReplace(pub, "i)kawadeshobo", "Kawade Shobō")
			pub:= regExReplace(pub, "i)kazamashobo", "Kazama Shobō")
			pub:= regExReplace(pub, "i)keiogijukudaigakushuppankai", "Keiō Gijuku Daigaku Shuppankai")
			pub:= regExReplace(pub, "i)keizaisangyochosakai", "Keizai Sangyō Chōsakai")
			pub:= regExReplace(pub, "i)kenbunshuppan", "Kenbun Shuppan")
			pub:= regExReplace(pub, "i)keisoshobo", "Keiso Shobō")
			pub:= regExReplace(pub, "i)kin*yūzaiseijijōkenkyūkai", "Kin'yū Zaisei Jijō Kenkyūkai")
			pub:= regExReplace(pub, "i)kitaojishobo", "Kitaōji Shobō")
			pub:= regExReplace(pub, "i)kobunsha", "Kōbunsha")
			pub:= regExReplace(pub, "i)koraidaigakkogurobarunihonkenkyuin", "Kōrai Daigakkō Gurōbaru Nihon Kenkyūin")
			pub:= regExReplace(pub, "i)koraidaigakkonihonkenkyusenta", "Kōrai Daigakkō Nihon Kenkyū Sentā")
			pub:= regExReplace(pub, "i)koseitorihikikyokai", "Kōsei Torihiki Kyōkai")
			pub:= regExReplace(pub, "i)koamagajin", "Koa Magajin")
			pub:= regExReplace(pub, "i)kodansha", "Kōdansha")
			pub:= regExReplace(pub, "i)kokusaishoin", "Kokusai Shōin")
			pub:= regExReplace(pub, "i)koyoshobo", "Kōyō Shobō")
			pub:= regExReplace(pub, "i)kyoeishobo", "Kyōei Shobo")
			pub:= regExReplace(pub, "i)kyotodaigakugakujutsushuppankai", "Kyōto Daigakugakujutsu Shuppankai")
			pub:= regExReplace(pub, "i)kyukoshoin", "Kyūko Shoin")
		;m	
			pub:= regExReplace(pub, "i)mainichishinbunsha", "Mainichi Shinbunsha")
			pub:= regExReplace(pub, "i)mainichishinbunshuppan", "Mainichi Shinbun Shuppan")
			pub:= regExReplace(pub, "i)matsurikasha", "Matsuri Kasha")
			pub:= regExReplace(pub, "i)meijitoshoshuppan", "Meiji Tosho Shuppan")
			pub:= regExReplace(pub, "i)meitokushuppansha", "Meitoku Shuppansha")
			pub:= regExReplace(pub, "i)minerubashobo", "Mineurba Shobō")
			pub:= regExReplace(pub, "i)minjihokenkyukai", "Minjihō Kenkyūkai")
			pub:= regExReplace(pub, "i)misuzushobo", "Misuzu Shobō")
			pub:= regExReplace(pub, "i)mitoshiyakusho", "Mitoshi Yakusho")
			pub:= regExReplace(pub, "i)mitsumurasuikoshoin", "Mitsumura Suiko Shoin")
			pub:= regExReplace(pub, "i)mitsumuratoshoshuppan", "Mitsumura Tosho Shuppan")
			pub:= regExReplace(pub, "i)miyaishoten", "Miyai Shoten")
			pub:= regExReplace(pub, "i)miyaobishuppansha", "Miyaobi Shuppansha")
			pub:= regExReplace(pub, "i)musashinobijutsudaigakushuppankyoku", "Musashino Bijutsu Daigaku Shuppankyoku")
		;n	
			pub:= regExReplace(pub, "i)nichigaiasoshietsu", "Nichigai Asoshiētsu")
			pub:= regExReplace(pub, "i)nihonkajoshuppan", "Nihon Kajo Shuppan")
			pub:= regExReplace(pub, "i)nihonhyoronsha", "Nihon Hyōronsha")
			pub:= regExReplace(pub, "i)nihonkindaibungakukan", "Nihon Kindai Bungagkukan")
			pub:= regExReplace(pub, "i)nihonkyohosha", "Nihon Kyōhōsha")
			pub:= regExReplace(pub, "i)nihonkyohosha", "Nihon Kyōhōsha")
			pub:= regExReplace(pub, "i)nikkeibipimaketingu", "Nikkei BP Māketingu")
			pub:= regExReplace(pub, "i)nikkeinashonarujiogurafikkusha", "Nikkei Nashonaru Jiogurafikkusha")
			pub:= regExReplace(pub, "i)nishidashoten", "Nishida Shoten")
		;o	
			pub:= regExReplace(pub, "i)okurazaimukyōkai", "Ōkura Zaimu Kyōkai")
			pub:= regExReplace(pub, "i)otsukishoten", "Ōtsuki Shoten")
			pub:= regExReplace(pub, "i)ochanomizushobo", "Ochanomizu Shobō")
			pub:= regExReplace(pub, "i)otowashobotsurumishoten", "Otowa Shobō Tsurumi Shoten")
		;p	
			pub:= regExReplace(pub, "i)paintanashonaru", "Pai Intānashonaru")
		;r	
			pub:= regExReplace(pub, "i)rekishishunjushuppan", "Rekishi Shunjun Shuppan")
			pub:= regExReplace(pub, "i)rikkashuppan", "Rikka Shuppan")
			pub:= regExReplace(pub, "i)rodochosakai", "Rōdō Chōsakai")
			pub:= regExReplace(pub, "i)romugyosei", "Rōmu Gyōsei")
		;s	
			pub:= regExReplace(pub, "i)san'ichishobo", "San'Ichi Shobō")
			pub:= regExReplace(pub, "i)san*ninsha", "San'ninsha")
			pub:= regExReplace(pub, "i)sanninsha", "San'ninsha")
			pub:= regExReplace(pub, "i)seibundoshuppan", "Seibundō Shuppan")
			pub:= regExReplace(pub, "i)seirinshoin", "Seirin Shoin")
			pub:= regExReplace(pub, "i)sekaishisosha", "Sekai Shisōsha")
			pub:= regExReplace(pub, "i)serikashobo", "Serika Shobō")
			pub:= regExReplace(pub, "i)shinchosha", "Shinchōsha")
			pub:= regExReplace(pub, "i)shinnihonhokishuppan", "Shin Nihon Hōki Shuppan")
			pub:= regExReplace(pub, "i)shinnihonshuppansha", "Shinnihon Shuppansha")
			pub:= regExReplace(pub, "i)shogakukan", "Shōgakukan")
			pub:= regExReplace(pub, "i)shojihomu", "Shoji Hōmu")
			pub:= regExReplace(pub, "i)shueishaintanashonaru", "Shueisha Intānashonaru")
			pub:= regExReplace(pub, "i)shueisha", "Shūeisha")
			pub:= regExReplace(pub, "i)shuwashisutemu", "Shūwa Shisutemu")
			pub:= regExReplace(pub, "i)sogensha", "Sōgensha")
		;t	
			pub:= regExReplace(pub, "i)tachibanashobo", "Tachibana Shobō")
			pub:= regExReplace(pub, "i)taiseishuppansha", "Taisei Shuppansha")
			pub:= regExReplace(pub, "i)taishukanshoten", "Taishūkan Shoten")
			pub:= regExReplace(pub, "i)tohoshoten", "Tōhō Shoten")
			pub:= regExReplace(pub, "i)tokyodaigakushuppankai", "Tōkyō Daigaku Shuppankai")
			pub:= regExReplace(pub, "i)tokyodoshuppan", "Tōkyōdō Shuppan")
			pub:= regExReplace(pub, "i)tokyohoreishuppan", "Tōkyō Hōrei Shuppan")
		;u	
			pub:= regExReplace(pub, "i)ueibushuppan", "Ueibu Shuppan")
		;y	
			pub:= regExReplace(pub, "i)yachiyo\shuppan", "Yahiyo Shuppan")
			pub:= regExReplace(pub, "i)yamakawashuppansha", "Yamakawa Shuppansha")
			pub:= regExReplace(pub, "i)yoshikawakobunkan", "Yoshikawa Kōbunkan")
			pub:= regExReplace(pub, "i)yumanishobo", "Yumani Shobō")
			pub:= regExReplace(pub, "i)yushindokobunsha", "Yūshindō Kōbunsha")
			
	;Adds space for "shōbo" and "shuppan".	
		pub:= regExReplace(pub, "i)shobo", " Shobō")
		pub:= regExReplace(pub, "i)shuppan", " Shuppan")
	
	;Fix formatting
		;Remove double spaces.
			loop{
				if inStr(pub, "  ")
					pub:= RegExReplace(pub, "  ", " ")
				else
					break
			}
		;Remove double macrons
			loop{
				if inStr(pub, "̄̄")
					pub:= RegExReplace(pub, "̄̄", "̄")
				else
					break
			}

	;Set publiser name to title case.
		pub:= StrTitle(pub)
		pub:= regExReplace(pub, "N/a", "n/a")
		
	return pub
}

;■■■■■■■■■■■■■ Prep and clean up text string to have no composed characters or spaces.
;▼▲▼▲▼▲▼▲▼▲▼▲▼ Functions
	fixPub(){
		if WinActive("ahk_exe EXCEL.EXE") | WinActive(" - Google Sheets") | WinActive(".xls"){
			pub:= copy()
			pub:= RegExReplace(pub, "`r`n$")
		}else{
			Send "^+{left}"
			Sleep nt
			pub:= copy()
		}
		pub:= fixRomanizedPublisherNames(pub)
		Send pub
		Sleep nt
		if WinActive("ahk_exe EXCEL.EXE") | WinActive(" - Google Sheets") | WinActive(".xls")
			Send "{enter}"
}
;■■■
^numpad9::fixPub()
!numpad9::{
		loop{
			pub:= copy()
			if(pub= "") | (pub="`r`n")
				break
			fixPub()
		}
	}
F9::fixpub()
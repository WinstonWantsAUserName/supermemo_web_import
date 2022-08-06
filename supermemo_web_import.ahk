#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

GroupAdd, Browsers, ahk_exe chrome.exe
GroupAdd, Browsers, ahk_exe firefox.exe
GroupAdd, Browsers, ahk_exe msedge.exe  ; Microsoft Edge

ReleaseKey(Key) {
  if (GetKeyState(Key))
    send {blind}{l%Key% up}{r%Key% up}
}

ClipboardGet_HTML( byref Data ) { ; www.autohotkey.com/forum/viewtopic.php?p=392624#392624
  If CBID := DllCall( "RegisterClipboardFormat", Str,"HTML Format", UInt )
  If DllCall( "IsClipboardFormatAvailable", UInt,CBID ) <> 0
    If DllCall( "OpenClipboard", UInt,0 ) <> 0
    If hData := DllCall( "GetClipboardData", UInt,CBID, UInt )
        DataL := DllCall( "GlobalSize", UInt,hData, UInt )
      , pData := DllCall( "GlobalLock", UInt,hData, UInt )
      , VarSetCapacity( data, dataL * ( A_IsUnicode ? 2 : 1 ) ), StrGet := "StrGet"
      , A_IsUnicode ? Data := %StrGet%( pData, dataL, 0 )
                    : DllCall( "lstrcpyn", Str,Data, UInt,pData, UInt,DataL )
      , DllCall( "GlobalUnlock", UInt,hData )
  DllCall( "CloseClipboard" )
  Return dataL ? dataL : 0
}

CleanHTML(Str) {
	; zzz in case you used f6 in SuperMemo to remove format before,
	; which disables the tag by adding zzz (e.g. <FONT> -> <ZZZFONT>)
	; Str := RegExReplace(Str, "is)( zzz| )style=(""|')BACKGROUND-IMAGE: url.*?(""|')")
	Str := RegExReplace(Str, "( zzz| )style=(""|').*?(""|')")
	Str := RegExReplace(Str, "ism)<\/{0,1}(zzz|)font.*?>")
	Str := RegExReplace(Str, "i)<P[^>]?+>(<BR>)+<\/P>")
	; Str := RegExReplace(Str, "is)<BR", "<P")
	Str := RegExReplace(Str, "i)<H5 dir=ltr align=left>")
	Str := RegExReplace(Str, "s)src=""file:\/\/\/.*?elements\/", "src=""file:///[PrimaryStorage]")
	Str := RegExReplace(Str, "i)\/svg\/", "/png/")
	Str := RegExReplace(Str, "i)<P[^>]?+>&nbsp;<\/P>")
	Str := RegExReplace(Str, "i)<DIV[^>]+>&nbsp;<\/DIV>")
	Return Str
}

StrReverse(String) {  ; https://www.autohotkey.com/boards/viewtopic.php?t=27215
  String .= "", DllCall("msvcrt.dll\_wcsrev", "Ptr", &String, "CDecl")
	return String
}

GetBrowserInfo(ByRef BrowserTitle, ByRef BrowserUrl, ByRef BrowserSource, ByRef BrowserDate) {
	BrowserTitle := BrowserUrl := BrowserSource := BrowserDate := ""
	ClipSaved := ClipboardAll
	Clipboard := ""
	CurrentTick := A_TickCount
	send {f6}^l  ; for moronic websites that use ctrl+L as a shortcut (I'm looking at you, paratranz)
	while (!Clipboard) {
		send ^l^c
		if (A_TickCount := CurrentTick + 500)
			Break
	}
	If (Clipboard) {
    BrowserUrl := RegExReplace(Clipboard, "#(.*)$")
    if (InStr(BrowserUrl, "youtube.com") && InStr(BrowserUrl, "v=")) {
      RegExMatch(BrowserUrl, "v=\K[\w\-]+", YTLink)
      BrowserUrl := "https://www.youtube.com/watch?v=" . YTLink
    } else if (InStr(BrowserUrl, "bilibili.com/video")) {
			BrowserUrl := RegExReplace(BrowserUrl, "(\/\?|&)vd_source=.*")
    } else if (InStr(BrowserUrl, "netflix.com/watch")) {
			BrowserUrl := RegExReplace(BrowserUrl, "\?trackId=.*")
		}
    GetBrowserTitleSourceDate(BrowserUrl, BrowserTitle, BrowserSource, BrowserDate)
	}
	if (WinActive("ahk_exe msedge.exe")) {
		send ^l{f6}
	} else {
		send ^l+{f6}
	}
	Clipboard := ClipSaved
}

GetBrowserTitleSourceDate(BrowserUrl, ByRef BrowserTitle, ByRef BrowserSource, ByRef BrowserDate) {
	WinGetActiveTitle BrowserTitle
	BrowserTitle := RegExReplace(BrowserTitle, "( - Google Chrome| — Mozilla Firefox|( and [0-9]+ more pages?)? - [^-]+ - Microsoft​ Edge)$")
	; Sites that need special attention
	if (InStr(BrowserTitle, "很帅的日报")) {
		BrowserDate := StrReplace(BrowserTitle, "很帅的日报 ")
		BrowserTitle := "很帅的日报"
	} else if (InStr(BrowserTitle, "_百度百科")) {
		BrowserSource := "百度百科"
		BrowserTitle := StrReplace(BrowserTitle, "_百度百科")
	} else if (InStr(BrowserUrl, "reddit.com")) {
		RegExMatch(BrowserUrl, "reddit\.com\/\Kr\/[^\/]+", BrowserSource)
		BrowserTitle := StrReplace(BrowserTitle, " : " . StrReplace(BrowserSource, "r/"))
	; Sites that don't include source in the title
	} else if (InStr(BrowserUrl, "dailystoic.com")) {
		BrowserSource := "Daily Stoic"
	} else if (InStr(BrowserUrl, "healthline.com")) {
		BrowserSource := "Healthline"
	} else if (InStr(BrowserUrl, "medicalnewstoday.com")) {
		BrowserSource := "Medical News Today"
	; Sites that should be skipped
	} else if (InStr(BrowserUrl, "mp.weixin.qq.com")) {
		return
	} else if (InStr(BrowserUrl, "universityhealthnews.com")) {
		return
	; Try to use - or | to find source
	} else {
		ReversedTitle := StrReverse(BrowserTitle)
		if (InStr(ReversedTitle, " | ") && (!InStr(ReversedTitle, " - ") || InStr(ReversedTitle, " | ") < InStr(ReversedTitle, " - "))) {  ; used to find source
			separator := " | "
		} else if (InStr(ReversedTitle, " - ")) {
			separator := " - "
		} else if (InStr(ReversedTitle, " – ")) {
			separator := " – "  ; websites like BetterExplained
		} else {
			separator := ""
		}
		pos := separator ? InStr(StrReverse(BrowserTitle), separator) : 0
		if (pos) {
			BrowserSource := SubStr(BrowserTitle, StrLen(BrowserTitle) - pos - 1, StrLen(BrowserTitle))
			if (InStr(BrowserSource, separator))
				BrowserSource := StrReplace(BrowserSource, separator,,, 1)
			BrowserTitle := SubStr(BrowserTitle, 1, StrLen(BrowserTitle) - pos - 2)
		}
	}
}

WaitCaretMove(OriginalX:=0, OriginalY:=0, TimeOut:=5000) {
	if (!OriginalX)
		MouseGetPos, OriginalX
	if (!OriginalY)
		MouseGetPos,, OriginalY
	StartTime := A_TickCount
	loop {
		if (A_CaretX != OriginalX || A_CaretY != OriginalY) {
			return true
		} else if (TimeOut && A_TickCount - StartTime > TimeOut) {
			return false
		}
	}
}

#If (WinActive("ahk_group Browsers") && WinExist("ahk_class TElWind"))
^+!a::  ; ctrl+shift+alt+a to import to supermemo
	ReleaseKey("ctrl")
	ReleaseKey("shift")
  KeyWait alt
  FormatTime, CurrentTime,, % "yyyy-MM-dd HH:mm:ss:" . A_msec
  ClipSaved := ClipboardAll
  LongCopy := A_TickCount, Clipboard := "", LongCopy -= A_TickCount  ; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent clipwait will need
  send ^c
  ClipWait, LongCopy ? 0.6 : 0.2, True
  if (!Clipboard) {
    send ^a^c
    ClipWait, LongCopy ? 0.6 : 0.2, True
    if (!Clipboard)
      Return
  }
  GetBrowserInfo(BrowserTitle, BrowserUrl, BrowserSource, BrowserDate)
  if (ClipboardGet_HTML(Data)) {
    HTML := CleanHTML(data)
    RegExMatch(HTML, "s)(?<=<!--StartFragment-->).*(?=<!--EndFragment-->)", HTML)
    source := BrowserSource ? "<br>#Source: " . BrowserSource : ""
    date := BrowserDate ? "<br>#Date: " . BrowserDate : "<br>#Date: Imported on " . CurrentTime
    clipboard := HTML
                . "<br>#SuperMemo Reference:"
                . "<br>#Link: " . BrowserUrl
                . source
                . date
                . "<br>#Title: " . BrowserTitle
    ClipWait 10
    WinActivate, ahk_class TElWind
    send ^{enter}h{enter}  ; clear search highlight, just in case
    WinWaitActive, ahk_class TElWind,, 0
    send ^n
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
		send ^a^+1
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
    MouseGetPos, XCoord, YCoord
    send +{home}
    WaitCaretMove(XCoord, YCoord)
    send !t  ; set title
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
    send {esc}^+{f6}
    ; WinWaitActive, ahk_class Notepad,, 5
    WinWaitNotActive, ahk_class TElWind,, 5
    WinKill, ahk_class Notepad
  }
  BrowserUrl := BrowserTitle := BrowserSource := BrowserDate := ""
  Vim.State.SetMode("Vim_Normal")
  sleep 700
  clipboard := ClipSaved
Return
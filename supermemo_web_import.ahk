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
  ; zzz in case you used f6 to remove format before,
  ; which disables the tag by adding zzz (e.g. <FONT> -> <ZZZFONT>)
  Str := RegExReplace(Str, "is)( zzz| )style=""((?!BACKGROUND-IMAGE: url).)*?""")
  Str := RegExReplace(Str, "is)( zzz| )style='((?!BACKGROUND-IMAGE: url).)*?'")
  Str := RegExReplace(Str, "ism)<\/{0,1}(zzz|)font.*?>")
  Str := RegExReplace(Str, "is)<BR", "<P")
  Str := RegExReplace(Str, "i)<H5 dir=ltr align=left>")
  Str := RegExReplace(Str, "s)src=""file:\/\/\/.*?elements\/", "src=""file:///[PrimaryStorage]")
  Str := RegExReplace(Str, "i)\/svg\/", "/png/")
  Str := RegExReplace(Str, "i)\n<P>&nbsp;<\/P>")
  Return Str
}

StrReverse(String) {  ; https://www.autohotkey.com/boards/viewtopic.php?t=27215
  String .= "", DllCall("msvcrt.dll\_wcsrev", "Ptr", &String, "CDecl")
	return String
}

GetBrowserInfo(ByRef BrowserTitle, ByRef BrowserUrl, ByRef BrowserSource) {
	ClipSaved := ClipboardAll
	Clipboard := ""
	CurrentTick := A_TickCount
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
    }
    WinGetActiveTitle, BrowserTitle
    BrowserTitle := RegExReplace(BrowserTitle, " - Google Chrome$")
    BrowserTitle := RegExReplace(BrowserTitle, " — Mozilla Firefox$")
    BrowserTitle := RegExReplace(BrowserTitle, " - .* - Microsoft​ Edge$")
    ReversedTitle := StrReverse(BrowserTitle)
    if (InStr(ReversedTitle, " | ")
        && (InStr(ReversedTitle, " | ") < InStr(ReversedTitle, " - ")
            || !InStr(ReversedTitle, " - "))) {  ; used to find source
      separator := " | "
    } else if (InStr(ReversedTitle, " - ")) {
      separator := " - "
    } else {
      separator := ""
    }
		; occurence := (InStr(ReversedTitle, separator,,, 2) > InStr(ReversedTitle, separator)) ? 2 : 1
		occurence := 1
    pos := separator ? InStr(StrReverse(BrowserTitle), separator,,, occurence) : 0
    if (pos) {
      BrowserSource := SubStr(BrowserTitle, StrLen(BrowserTitle) - pos - 1, StrLen(BrowserTitle))
      if (InStr(BrowserSource, separator))
        BrowserSource := StrReplace(BrowserSource, separator,,, 1)
      BrowserTitle := SubStr(BrowserTitle, 1, StrLen(BrowserTitle) - pos - 2)
    }
	}
	send {f6 2}
	Clipboard := ClipSaved
	Return
}

#If (WinActive("ahk_group Browsers") && WinExist("ahk_class TElWind"))
^+!a::  ; ctrl+shift+alt+a to import to supermemo
	ReleaseKey("ctrl")
	ReleaseKey("shift")
  KeyWait alt
  FormatTime, CurrentTime,, yyyy-MM-dd HH:mm:ss:%A_msec%
  ClipSaved := ClipboardAll
  clipboard := ""
  send ^c
  ClipWait 0.6
  if (ErrorLevel) {
    send ^a^c
    clipwait 0.6
    if (ErrorLevel)
      Return
  }
  GetBrowserInfo(BrowserTitle, BrowserUrl, BrowserSource)
  if (ClipboardGet_Html(Data)) {
    Html := CleanHTML(data)
    RegExMatch(Html, "s)(?<=<!--StartFragment-->).*(?=<!--EndFragment-->)", Html)
    if (BrowserSource) {
      clipboard := Html . "<br>#SuperMemo Reference:"
                  . "<br>#Date: Imported on " . CurrentTime
                  . "<br>#Source: " . BrowserSource
                  . "<br>#Link: " . BrowserUrl
                  . "<br>#Title: " . BrowserTitle
    } else {
      clipboard := Html . "<br>#SuperMemo Reference:"
                  . "<br>#Date: Imported on " . CurrentTime
                  . "<br>#Link: " . BrowserUrl
                  . "<br>#Title: " . BrowserTitle
    }
    ClipWait
    WinActivate, ahk_class TElWind
    send ^n
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
		send ^a^+1
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
    send +{home}!t  ; set title
		WinWaitNotActive, ahk_class TElWind,, 1.5  ; could appear a loading bar
		if (!ErrorLevel)
			WinWaitActive, ahk_class TElWind,, 5
    send {esc}^+{f6}
    ; WinWaitActive, ahk_class Notepad,, 5
    WinWaitNotActive, ahk_class TElWind,, 5
    WinKill, ahk_class Notepad
  }
  sleep 700
  clipboard := ClipSaved
Return
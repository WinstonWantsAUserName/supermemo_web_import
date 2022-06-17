#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

GroupAdd, Browsers, ahk_exe chrome.exe
GroupAdd, Browsers, ahk_exe firefox.exe
GroupAdd, Browsers, ahk_exe msedge.exe  ; Microsoft Edge

; Author:             Antonio Bueno <user atnbueno of Google's popular e-mail service>
; Short description:  Gets the URL of the current (active) browser tab for most modern browsers
ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"

GetActiveBrowserURL() {
	global ModernBrowsers, LegacyBrowsers
	WinGetClass, sClass, A
	If sClass In % ModernBrowsers
		Return GetBrowserURL_ACC(sClass)
	Else If sClass In % LegacyBrowsers
		Return GetBrowserURL_DDE(sClass) ; empty string if DDE not supported (or not a browser)
	Else
		Return ""
}

; "GetBrowserURL_DDE" adapted from DDE code by Sean, (AHK_L version by maraskan_user)
; Found at http://autohotkey.com/board/topic/17633-/?p=434518

GetBrowserURL_DDE(sClass) {
	WinGet, sServer, ProcessName, % "ahk_class " sClass
	StringTrimRight, sServer, sServer, 4
	iCodePage := A_IsUnicode ? 0x04B0 : 0x03EC ; 0x04B0 = CP_WINUNICODE, 0x03EC = CP_WINANSI
	DllCall("DdeInitialize", "UPtrP", idInst, "Uint", 0, "Uint", 0, "Uint", 0)
	hServer := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", sServer, "int", iCodePage)
	hTopic := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "WWW_GetWindowInfo", "int", iCodePage)
	hItem := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "0xFFFFFFFF", "int", iCodePage)
	hConv := DllCall("DdeConnect", "UPtr", idInst, "UPtr", hServer, "UPtr", hTopic, "Uint", 0)
	hData := DllCall("DdeClientTransaction", "Uint", 0, "Uint", 0, "UPtr", hConv, "UPtr", hItem, "UInt", 1, "Uint", 0x20B0, "Uint", 10000, "UPtrP", nResult) ; 0x20B0 = XTYP_REQUEST, 10000 = 10s timeout
	sData := DllCall("DdeAccessData", "Uint", hData, "Uint", 0, "Str")
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hServer)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hTopic)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hItem)
	DllCall("DdeUnaccessData", "UPtr", hData)
	DllCall("DdeFreeDataHandle", "UPtr", hData)
	DllCall("DdeDisconnect", "UPtr", hConv)
	DllCall("DdeUninitialize", "UPtr", idInst)
	csvWindowInfo := StrGet(&sData, "CP0")
	StringSplit, sWindowInfo, csvWindowInfo, `" ;"; comment to avoid a syntax highlighting issue in autohotkey.com/boards
	Return sWindowInfo2
}

GetBrowserURL_ACC(sClass) {
	global nWindow, accAddressBar
	If (nWindow != WinExist("ahk_class " sClass)) ; reuses accAddressBar if it's the same window
	{
		nWindow := WinExist("ahk_class " sClass)
		accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindow))
	}
	Try sURL := accAddressBar.accValue(0)
	If (sURL == "") {
		WinGet, nWindows, List, % "ahk_class " sClass ; In case of a nested browser window as in the old CoolNovo (TO DO: check if still needed)
		If (nWindows > 1) {
			accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindows2))
			Try sURL := accAddressBar.accValue(0)
		}
	}
	If ((sURL != "") and (SubStr(sURL, 1, 4) != "http")) ; Modern browsers omit "http://"
		sURL := "http://" sURL
	If (sURL == "")
		nWindow := -1 ; Don't remember the window if there is no URL
	Return sURL
}

; "GetAddressBar" based in code by uname
; Found at http://autohotkey.com/board/topic/103178-/?p=637687

GetAddressBar(accObj) {
	Try If ((accObj.accRole(0) == 42) and IsURL(accObj.accValue(0)))
		Return accObj
	Try If ((accObj.accRole(0) == 42) and IsURL("http://" accObj.accValue(0))) ; Modern browsers omit "http://"
		Return accObj
	For nChild, accChild in Acc_Children(accObj)
		If IsObject(accAddressBar := GetAddressBar(accChild))
			Return accAddressBar
}

IsURL(sURL) {
	Return RegExMatch(sURL, "^(?<Protocol>https?|ftp)://(?<Domain>(?:[\w-]+\.)+\w\w+)(?::(?<Port>\d+))?/?(?<Path>(?:[^:/?# ]*/?)+)(?:\?(?<Query>[^#]+)?)?(?:\#(?<Hash>.+)?)?$")
}

; The code below is part of the Acc.ahk Standard Library by Sean (updated by jethrow)
; Found at http://autohotkey.com/board/topic/77303-/?p=491516

Acc_Init()
{
	static h
	If Not h
		h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Acc_ObjectFromWindow(hWnd, idObject = 0)
{
	Acc_Init()
	If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return ComObjEnwrap(9,pacc,1)
}
Acc_Query(Acc) {
	Try Return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}
Acc_Children(Acc) {
	If ComObjType(Acc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	Else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		If DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren%
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=9?Acc_Query(child):child), NumGet(varChildren,i-8)=9?ObjRelease(child):
			Return Children.MaxIndex()?Children:
		} Else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
}

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

#If (WinActive("ahk_group Browsers"))
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
	sURL := GetActiveBrowserURL()
	WinGetClass, sClass, A
	If (sURL != "") {
    BrowserUrl := RegExReplace(sURL, "#(.*)$")
    if (InStr(BrowserUrl, "https://www.youtube.com") && InStr(BrowserUrl, "v=")) {
      RegExMatch(BrowserUrl, "v=\K[\w\-]+", YTLink)
      BrowserUrl := "https://www.youtube.com/watch?v=" . YTLink
    }
    WinGetActiveTitle, BrowserTitle
    BrowserTitle := RegExReplace(BrowserTitle, " - Google Chrome$")
    BrowserTitle := RegExReplace(BrowserTitle, " — Mozilla Firefox$")
    BrowserTitle := RegExReplace(BrowserTitle, " - .* - Microsoft​ Edge$")
    ReversedTitle := StrReverse(BrowserTitle)
    if (InStr(ReversedTitle, " | ") < InStr(ReversedTitle, " - ") && InStr(ReversedTitle, " | ")) {  ; used to find source
      separator := " | "
    } else {
      separator := " - "
    }
    pos := InStr(StrReverse(BrowserTitle), separator)
    if (pos) {
      BrowserSource := SubStr(BrowserTitle, StrLen(BrowserTitle) - pos - 1, StrLen(BrowserTitle))
      if (InStr(BrowserSource, separator))
        BrowserSource := StrReplace(BrowserSource, separator)
      BrowserTitle := SubStr(BrowserTitle, 1, StrLen(BrowserTitle) - pos - 2)
    }
  }
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
    send {esc}^+{f6}
    WinWaitNotActive, ahk_class TElWind,, 5
    WinKill, ahk_class Notepad
  }
  sleep 700
  clipboard := ClipSaved
Return
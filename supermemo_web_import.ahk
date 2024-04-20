#Requires AutoHotkey v1.1.1+  ; so that the editor would recognise this script as AHK V1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

GroupAdd, Browser, ahk_exe chrome.exe
GroupAdd, Browser, ahk_exe firefox.exe
GroupAdd, Browser, ahk_exe msedge.exe

#Include <lib>

SM := new SM()
Browser := new Browser()

#if (WinActive("ahk_group Browser"))
; Incremental web browsing
^+!b::
; Import current webpage to SuperMemo
; Incremental video: Import current YT video to SM
^+!a::
^!a::
  if (!WinExist("ahk_class TElWind")) {
    SetToolTip("Please open SuperMemo and try again.")
    return
  }
  if (WinExist("ahk_id " . SMImportGuiHwnd)) {
    WinActivate
    return
  }

  ClipSaved := ClipboardAll
  if (IWB := IfContains(A_ThisLabel, "IWB,^+!b")) {
    if (!HTMLText := Copy(false, true)) {
      SetToolTip("Text not found.")
      Clipboard := ClipSaved
      return
    }
  }

  Browser.Clear()
  if (IWB) {
    Browser.Url := Browser.ParseUrl(RetrieveUrlFromClip())
  } else {
    Browser.Url := Browser.GetUrl()
  }
  if (!Browser.Url) {
    SetToolTip("Url not found.")
    Clipboard := ClipSaved
    return
  }

  wBrowser := "ahk_id " . WinActive("A")
  Browser.FullTitle := Browser.GetFullTitle("A")
  IsVideoOrAudioSite := Browser.IsVideoOrAudioSite(Browser.FullTitle)

  SM.CloseMsgDialog()
  CollName := SM.GetCollName()
  OnlineEl := SM.IsOnline(CollName, -1)

  DupChecked := MB := false
  if (!IWB) {
    if (SM.CheckDup(Browser.Url, false))
      MB := MsgBox(3,, "Continue import?")
    DupChecked := true
  }
  WinClose, % "ahk_class TBrowser ahk_pid " . WinGet("PID", "ahk_class TElWind")
  WinActivate % wBrowser
  if (IfIn(MB, "No,Cancel"))
    Goto SMImportReturn

  Prio := Concept := CloseTab := DLHTML := ResetTimeStamp := CheckDupForIWB := ""
  Tags := RefComment := ClipBeforeGui := UseOnlineProgress := ""
  DLList := "economist.com,investopedia.com,webmd.com,britannica.com,medium.com,wired.com"
  if (IfIn(A_ThisLabel, "^+!a,IWBPriorityAndConcept,^+!b")) {
    ClipBeforeGui := Clipboard
    SetDefaultKeyboard(0x0409)  ; English-US
    Gui, SMImport:Add, Text,, % "Current collection: " . CollName
    Gui, SMImport:Add, Text,, &Priority:
    Gui, SMImport:Add, Edit, vPrio w280
    Gui, SMImport:Add, Text,, Concept &group:  ; like in default import dialog
    ConceptList := "||Online|Sources|ToDo"
    if (IfIn(CurrConcept := SM.GetDefaultConcept(), "Online,Sources,ToDo"))
      ConceptList := StrReplace(ConceptList, "|" . CurrConcept)
    list := StrLower(CurrConcept . ConceptList)
    Gui, SMImport:Add, ComboBox, vConcept gAutoComplete w280, % list
    Gui, SMImport:Add, Text,, &Tags (without # and use `; to separate):
    Gui, SMImport:Add, Edit, vTags w280
    Gui, SMImport:Add, Text,, Reference c&omment:
    Gui, SMImport:Add, Edit, vRefComment w280
    Gui, SMImport:Add, Checkbox, vCloseTab, &Close tab  ; like in default import dialog
    if (!IWB && !OnlineEl)
      Gui, SMImport:Add, Checkbox, vOnlineEl, Import as o&nline element
    if (!IWB && !IsVideoOrAudioSite && !OnlineEl) {
      check := IfContains(Browser.Url, DLList) ? "checked" : ""
      Gui, SMImport:Add, Checkbox, % "vDLHTML " . check, Import fullpage &HTML
    }
    if (IWB)
      Gui, SMImport:Add, Checkbox, vCheckDupForIWB, Check &duplication
    if (IsVideoOrAudioSite || OnlineEl) {
      Gui, SMImport:Add, Checkbox, vResetTimeStamp, &Reset time stamp
      if (IfContains(Browser.Url, "youtube.com/watch")) {
        check := (CollName = "bgm") ? "checked" : ""
        Gui, SMImport:Add, Checkbox, % "vUseOnlineProgress " . check, &Mark as use online progress
      }
    }
    Gui, SMImport:Add, Button, default, &Import
    Gui, SMImport:Show,, SuperMemo Import
    Gui, SMImport:+HwndSMImportGuiHwnd
    return
  } else {
    DLHTML := IfContains(Browser.Url, DLList)
  }

SMImportButtonImport:
  if (A_ThisLabel == "SMImportButtonImport") {
    ; Without KeyWait Enter SwitchToSameWindow() below could fail???
    KeyWait Enter
    KeyWait I
    Gui, Submit
    Gui, Destroy
    if (Clipboard != ClipBeforeGui)
      ClipSaved := ClipboardAll
  }

  if (OnlineEl != 1)
    OnlineEl := SM.IsOnline(CollName, Concept)
  if (OnlineEl)  ; just in case user checks both of them
    DLHTML := false
  if (OnlineEl && IWB) {
    ret := true
    if (MsgBox(3,, "You chosed an online concept. Choose again?") = "Yes") {
      Concept := InputBox(, "Enter a new concept:")
      if (!ErrorLevel && !SM.IsOnline(-1, Concept))
        ret := false
    }
    if (ret)
      Goto SMImportReturn
  }

  SwitchToSameWindow(wBrowser)
  if (!IWB)  ; IWB copies text before
    HTMLText := (DLHTML || OnlineEl) ? "" : Copy(false, true)  ; do not copy if download html or online element is checked

  if (CheckDupForIWB) {
    MB := ""
    if (SM.CheckDup(Browser.Url, false))
      MB := MsgBox(3,, "Continue import?")
    DupChecked := true
    WinClose, % "ahk_class TBrowser ahk_pid " . WinGet("PID", "ahk_class TElWind")
    WinActivate % wBrowser
    if (IfIn(MB, "No,Cancel"))
      Goto SMImportReturn
  }

  if (IWB)
    Browser.Highlight(CollName, Clipboard, Browser.Url)

  if (LocalFile := (Browser.Url ~= "^file:\/\/\/"))
    DLHTML := true
  SMCtrlNYT := (!OnlineEl && SM.IsCtrlNYT(Browser.Url))
  CopyAll := (!HTMLText && !OnlineEl && !DLHTML && !SMCtrlNYT)
  if (DLHTML) {
    if (LocalFile) {
      HTMLText := FileRead(EncodeDecodeURI(RegExReplace(Browser.Url, "^file:\/\/\/"), false))
      Browser.Url := RegExReplace(Browser.Url, "^file:\/\/\/", "file://")  ; SuperMemo converts file:/// to file://
    } else {
      SetToolTip("Attempting to download website...")
      if (!HTMLText := GetSiteHTML(Browser.Url)) {
        SetToolTip("Download failed."), CopyAll := true, DLHTML := false
      } else {
        ; Fixing links
        RegExMatch(Browser.Url, "^https?:\/\/.*?\/", UrlHead)
        RegExMatch(Browser.Url, "^https?:\/\/", HTTP)
        HTMLText := RegExReplace(HTMLText, "is)<([^<>]+)?\K (href|src)=""\/\/(?=([^<>]+)?>)", " $2=""" . HTTP)
        HTMLText := RegExReplace(HTMLText, "is)<([^<>]+)?\K (href|src)=""\/(?=([^<>]+)?>)", " $2=""" . UrlHead)
        HTMLText := RegExReplace(HTMLText, "is)<([^<>]+)?\K (href|src)=""(?=#([^<>]+)?>)", " $2=""" . Browser.Url)
      }
    }
  }

  if (CopyAll) {
    CopyAll()
    HTMLText := GetClipHTMLBody()
  }
  if (!OnlineEl && !HTMLText && !SMCtrlNYT) {
    SetToolTip("Text not found.")
    Goto SMImportReturn
  }

  SkipDate := (OnlineEl && !IsVideoOrAudioSite && (OnlineEl != 2))
  Browser.GetInfo(false,, (CopyAll ? Clipboard : ""),, !SkipDate, !ResetTimeStamp, (DLHTML ? HTMLText : ""))

  if (ResetTimeStamp)
    Browser.TimeStamp := "0:00"
  if (SkipDate)
    Browser.Date := ""

  SMPoundSymbHandled := SM.PoundSymbLinkToComment()
  if (Tags || RefComment) {
    TagsComment := ""
    if (Tags) {
      TagsComment := StrReplace(Trim(Tags), " ", "_")
      TagsComment := "#" . StrReplace(TagsComment, ";", " #")
    }
    if (RefComment && TagsComment)
      TagsComment := " " . TagsComment 
    if (Browser.Comment)
      Browser.Comment := " " . Browser.Comment
    Browser.Comment := Trim(RefComment) . TagsComment . Browser.Comment
  }

  WinClip.Clear()
  if (OnlineEl) {
    ScriptUrl := Browser.Url
    if (Browser.TimeStamp && (TimeStampedUrl := Browser.TimeStampToUrl(Browser.Url, Browser.TimeStamp)))
      ScriptUrl := TimeStampedUrl
    if (Browser.TimeStamp && !TimeStampedUrl) {
      Clipboard := "<SPAN class=Highlight>SMVim time stamp</SPAN>: " . Browser.TimeStamp . SM.MakeReference(true)
    } else if (UseOnlineProgress) {
      Clipboard := "<SPAN class=Highlight>SMVim: Use online video progress</SPAN>" . SM.MakeReference(true)
    } else {
      Clipboard := SM.MakeReference(true)
    }
  } else if (SMCtrlNYT) {
    Clipboard := Browser.Url
  } else {
    LineBreakList := "baike.baidu.com,m.shuaifox.com,khanacademy.org,mp.weixin.qq.com,webmd.com,proofwiki.org"
    LineBreak := IfContains(Browser.Url, LineBreakList)
    HTMLText := SM.CleanHTML(HTMLText,, LineBreak, Browser.Url)
    if (!IWB && !Browser.Date)
      Browser.Date := "Imported on " . GetDetailedTime()
    Clipboard := HTMLText . SM.MakeReference(true)
  }
  ClipWait

  InfoToolTip := "Importing:`n`n"
               . "Url: " . Browser.Url . "`n"
               . "Title: " . Browser.Title
  if (Browser.Source)
    InfoToolTip .= "`nSource: " . Browser.Source
  if (Browser.Author)
    InfoToolTip .= "`nAuthor: " . Browser.Author
  if (Browser.Date)
    InfoToolTip .= "`nDate: " . Browser.Date
  if (Browser.TimeStamp)
    InfoToolTip .= "`nTime stamp: " . Browser.TimeStamp
  if (Browser.Comment)
    InfoToolTip .= "`nComment: " . Browser.Comment
  SetToolTip(InfoToolTip)

  if (Prio ~= "^\.")
    Prio := "0" . Prio
  SM.CloseMsgDialog()

  ChangeBackConcept := ""
  if (Concept) {
    if ((OnlineEl == 1) && !SM.IsOnline(-1, Concept))
      ChangeBackConcept := Concept, Concept := "Online"
    if (!SM.SetDefaultConcept(Concept,, ChangeBackConcept))
      Goto SMImportReturn
  }

  if (SMCtrlNYT) {
    YT := (RegExMatch(Clipboard, "(?:youtube\.com).*?(?:v=)([a-zA-Z0-9_-]{11})", v) && IsUrl(Clipboard))
    ; Register browser time stamp to YT comp time stamp
    if (YT && Browser.TimeStamp) {
      WinClip.Clear()
      Clipboard := "{SuperMemoYouTube:" . v1 . "," . Browser.TimeStamp . ",0:00,0:00,3}"
      ClipWait
    }
    SM.CtrlN()
    if (YT) {
      Text := Browser.Title . SM.MakeReference()
      SM.WaitFileLoad()
      SM.EditFirstQuestion()
      SM.WaitTextFocus()
      Send ^a{BS}{Esc}
      SM.WaitTextExit()
      Clip(Text,, false)
      SM.WaitTextFocus()
      SM.WaitFileLoad()
    }
  } else {
    PrevSMTitle := WinGetTitle("ahk_class TElWind")
    SM.AltN()
    WinActivate, ahk_class TElWind
    SM.WaitTextFocus()
    TempTitle := WinWaitTitleChange(PrevSMTitle, "ahk_class TElWind")
    SM.PasteHTML()

    if (!OnlineEl) {
      SM.ExitText()
      WinWaitTitleChange(TempTitle, "A")

    } else if (OnlineEl) {
      pidSM := WinGet("PID", "ahk_class TElWind")
      Send ^t{f9}{Enter}
      WinWait, % wScript := "ahk_class TScriptEditor ahk_pid " . pidSM,, 5
      WinActivate, % wBrowser
      if (ErrorLevel) {
        SetToolTip("Script component not found.")
        Goto SMImportReturn
      }

      ; ControlSetText to "rl" first than send one "u" is needed to update the editor,
      ; thus prompting it to ask to save on exiting
      ControlSetText, TMemo1, % "rl " . ScriptUrl, % wScript
      ControlSend, TMemo1, {text}u, % wScript
      ControlSend, TMemo1, {Esc}, % wScript
      WinWait, % "ahk_class TMsgDialog ahk_pid " . pidSM
      ControlSend, ahk_parent, {Enter}
      WinWaitClose
      WinWaitClose, % wScript
    }
  }

  ; All SM operations here are handled in the background
  SM.SetElParam((IWB ? "" : Browser.Title), Prio, (SMCtrlNYT ? "YouTube" : ""), (ChangeBackConcept ? ChangeBackConcept : ""))
  if (DupChecked)
    SM.ClearHighlight()
  if (!SMPoundSymbHandled)
    SM.HandleSM19PoundSymbUrl(Browser.Url)
  SM.Reload()
  SM.WaitFileLoad()
  if (ChangeBackConcept)
    SM.SetDefaultConcept(ChangeBackConcept)
  if (Tags)
    SM.LinkConcepts(StrSplit(Tags, ";"),, wBrowser)
  SM.CloseMsgDialog()

  if (CloseTab) {
    WinActivate % wBrowser  ; apparently needed for closing tab
    ControlSend, ahk_parent, {LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}{Ctrl Down}w{Ctrl Up}, % wBrowser
  }

SMImportGuiEscape:
SMImportGuiClose:
SMImportReturn:
  EscGui := IfContains(A_ThisLabel, "SMImportGui")
  if (Esc := IfContains(A_ThisLabel, "SMImportGui,SMImportReturn")) {
    if (EscGui)
      Gui, Destroy
    if (DupChecked)
      SM.ClearHighlight()
  }
  if (OnlineEl || Esc) {
    Browser.ActivateBrowser(wBrowser)
  } else {
    SM.ActivateElWind()
  }
  Browser.Clear(), Vim.State.SetMode("Vim_Normal")
  ; If closed GUI but did not copy anything, restore clipboard
  ; If closed GUI but copied something while the GUI is open, do not restore clipboard
  if (!EscGui || (Clipboard == ClipBeforeGui))
    Clipboard := ClipSaved
  if (!Esc)
    SetToolTip("Import completed.")
  HTMLText := ""  ; empty memory
return

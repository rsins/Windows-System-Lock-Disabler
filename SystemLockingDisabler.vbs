Option Explicit

'***********************************************************************************
' Created By: Ravi Singh

  Const CREDIT_TITLE = " [(C) COPYRIGHT Ravi Singh]"

' Date Created  : May 10 2013
' Last Modified : Aug 23 2016
' Description   :
'     This VB Script is to disable system locking until user asks to cancel the 
'     script.
'***********************************************************************************

 
Dim MyUserMessage
Dim MyUserPausedMessage
Dim MsgLinesCount

Dim WSHShell
Dim IEWindowWait
Dim IEWindowCancel
Dim IEWindowPause

'MyUserMessage = "Disable System Locking Script: RUNNING," & vbNewLine & _
'                "click on the Pause button to pause script," & vbNewLine & _
'                "click on Cancel button to stop the script."
'
'MyUserPausedMessage = "Disable System Locking Script: PAUSED, " & vbNewLine & _
'                      "click on the Resume button to resume script," & vbNewLine & _
'                      "click on Cancel button to stop the script."

MyUserMessage = "Disable System Locking Script: RUNNING"
MyUserPausedMessage = "Disable System Locking Script: PAUSED"
MsgLinesCount = UBound(Split(MyUserMessage, vbNewLine)) + 1
                            
Set WSHShell = Wscript.CreateObject("WScript.Shell")

Call UserDisplayMessageHTML("Script Running", _
                            MyUserMessage, _
                            MyUserPausedMessage, _
                            MsgLinesCount, _
                            getref("PingSystemToDisableLock"), _
                            True, _
                            True)

MsgBox "Script is Cancelled/Closed.", vbInformation, "Script End"

Set WSHShell = Nothing
WScript.Quit (0)


'*****************************************************************************************
' Script ends here and Function Definitions start here.
'*****************************************************************************************

'------------------------------------------------------
' Function UserDisplayMessageHTML
' Display user a message in HTML format.
'------------------------------------------------------
Function UserDisplayMessageHTML(ByVal MyTitle, _
                                ByVal MyUserMessage, _
                                ByVal MyUserPausedMessage, _
                                ByVal MsgLinesCount, _
                                ByRef MyFunctionToRunEveryLoop, _
                                ByVal EnablePause, _
                                ByVal BringWindowToFocus)

   Const UserMessageColor1 = "brown"
   Const UserMessageColor2 = "#CCCCCC"
   Const BlinkMiliSecondInterval     = 1000    ' Text Blink Interval (ms) for User Message
   Const WaitBetweenEachRunMiliSec   = 5000    ' Send key press event after this interval (ms)
   Const WaitBetweenEachLoopMiliSec  = 100     ' Wait for this interval (ms) before checking for user input (pause/resume)
   
   Dim objIE
   Dim windowHeight
   Dim windowWidth
   
   Dim TotalWaitTime
   
   windowHeight = 175 + MsgLinesCount * 10 + (MsgLinesCount - 1) * 5
   windowWidth = (Len(MyTitle) + Len(CREDIT_TITLE) + Len("Internet Explorer")) * 6 + 130
   
   ' Start the Internet Explorer window to show the generated HTML page.
   Set objIE = CreateObject("InternetExplorer.Application")
   
   With objIE
      .FullScreen = False
      .Toolbar = False: .RegisterAsDropTarget = False
      .StatusBar = False: .Navigate ("about:blank")
      
      While .Busy: WScript.Sleep 100: Wend
      
      With .document
         With .ParentWindow
            .resizeto windowWidth, windowHeight
            .moveto (.screen.Width / 2 - windowWidth / 2), (.screen.Height / 2 - windowHeight / 2)
         End With
         
         .WriteLn("<html> " & vbNewLine & _
                  "<head> " & vbNewLine & _
                  "   <title> " & MyTitle & CREDIT_TITLE & " </title> ")

         .WriteLn("   <SCRIPT> ")
         
         .WriteLn("   function checkKey() " & vbNewLine & _
                  "   { " & vbNewLine & _
                  "     //if (window.event.keyCode == 13) { document.getElementById('usercancel').click();} " & vbNewLine & _
                  "     if (window.event.keyCode == 27) { document.getElementById('userclose').click();} " & vbNewLine & _
                  "   }  ")
                   
         .WriteLn("   function initUserMessageBlink() " & vbNewLine & _
          "   { " & vbNewLine & _
          "        var state = false; " & vbNewLine & _
          "        setInterval(function() " & vbNewLine & _
          "            { " & vbNewLine & _
          "                state = !state; " & vbNewLine & _
          "                var color = (state?'" & UserMessageColor2 & "':'" & UserMessageColor1 & "'); " & vbNewLine & _
          "                document.getElementById('usermessage').style.color = color; " & vbNewLine & _
          "            }, " & BlinkMiliSecondInterval & "); " & vbNewLine & _
                  "   } " & vbNewLine & _
                  "   initUserMessageBlink(); ")
                  
         .WriteLn("   </SCRIPT> ")
                  
         .WriteLn("   <style type=""text/css""> " & vbNewLine & _
                  "   body { background-color:#EEEEEE; font-size:12pt; font-face:verdana;} " & vbNewLine & _
                  "   table { font-size:10pt; font-face:verdana;} " & vbNewLine & _
				  "   table td { vertical-align:top; text-align:center;} " & vbNewLine & _
                  "   button { background-color:#DDDDDD; width:80px; font-size:8pt; font-face:verdana; font-weight:bold; border-style:ridge; border-width:1px; } " & vbNewLine & _
				  "   input[type=checkbox] { cursor: pointer; width: 12px; height: 12px } " & vbNewLine & _
                  "   </style> ")
                  
         .WriteLn("</head> " & vbNewLine & _
                  "<body onkeypress=""javascript:checkKey();""> " & vbNewLine & _
                  "   <table cellspacing=1 cellpadding=1 border=0 width=100%> " & vbNewLine & _
                  "   <tr><td>&nbsp;</td></tr> " & vbNewLine & _
                  "   <tr> " & vbNewLine & _
                  "     <td><b><div id=usermessage style=""color=" & UserMessageColor1 & ";"">" & MyUserMessage & "</div></b></td> " & vbNewLine & _
                  "   </tr> ")
         
         .WriteLn("   <tr> " & vbNewLine & _
              "     <td>" & _
              "        <br><label style=""color=#777777"">" & _ 
              "        <input type=""checkbox"" name=""usercheckbox"" checked=""yes"" value=""yes""/> Keep Window in Front when script is running.</label>" & _
              "     </td> " & vbNewLine & _
                  "   </tr> ")
                  
         .WriteLn("   <tr> " & vbNewLine & _
                  "     <td><br> " & vbNewLine)
                  
         If EnablePause Then
            .WriteLn("       <button id=userpauseresume>Pause</button>&nbsp;&nbsp;" & vbNewLine)
         Else
            .WriteLn("       <button id=userpauseresume style=""visibility:hidden;"">Pause</button><br>" & vbNewLine)
         End If
          
         .WriteLn("       <button id=usercancel>Cancel</button><br> " & vbNewLine & _
                  "       <button id=userclose style=""visibility:hidden;"">Close</button> " & vbNewLine & _
                  "     </td> " & vbNewLine & _
                  "   </tr> " & vbNewLine)
         
         .WriteLn("   </table> " & vbNewLine & _
                  "</body></html>")
         
         With .ParentWindow.document.body
            .scroll = "no"
            .Style.borderStyle = "outset"
            .Style.borderWidth = "3px"
         End With
         
         If BringWindowToFocus Then 
            .all.usercheckbox.checked = True
         Else
            .all.usercheckbox.checked = False
         End If
         
         .all.userpauseresume.onclick = getref("UserDisplayMessageHTML_PauseResume")
         .all.usercancel.onclick = getref("UserDisplayMessageHTML_Cancel")
         .all.userclose.onclick = getref("UserDisplayMessageHTML_Close")
         .all.usercancel.Focus
         
         objIE.Visible = True
         
         IEWindowWait = True
         IEWindowCancel = False
         IEWindowPause = False

         ' Wait for user input to be provided or window cancelled.         
         On Error Resume Next
         TotalWaitTime = 0
         While IEWindowWait
            If objIE.Visible Then 
               If Not IEWindowPause Then
                    .all.usermessage.innertext = MyUserMessage
                    .all.userpauseresume.innertext = "Pause"

                    If (TotalWaitTime = 0) Or (TotalWaitTime >= WaitBetweenEachRunMiliSec) Then
                        If .all.usercheckbox.checked Then
                            objIE.document.focus()
                        End If
                        
                        Call MyFunctionToRunEveryLoop()
                        TotalWaitTime = 0
                    End If
               Else
                  .all.usermessage.innertext = MyUserPausedMessage
                  .all.userpauseresume.innertext = "Resume"
                  TotalWaitTime = 0
               End If
            End If
            
            If Err Then 
               Call UserDisplayMessageHTML_Close()
            End If
            
            WScript.Sleep WaitBetweenEachLoopMiliSec
            TotalWaitTime = TotalWaitTime + WaitBetweenEachLoopMiliSec
         Wend
         
       End With ' document
       
       .Visible = False
  End With   ' IE
  
  objIE.Quit
  
  On Error Goto 0
  
  Set objIE = Nothing
End Function
'------------------------------------------------------

'------------------------------------------------------
' Function UserDisplayMessageHTML_Cancel
' If user clicks on Cancel button on the HTML input screen.
'------------------------------------------------------
Sub UserDisplayMessageHTML_Cancel()
    IEWindowWait = False
End Sub
'------------------------------------------------------

'------------------------------------------------------
' Function UserDisplayMessageHTML_Close
' If user closes the window.
'------------------------------------------------------
Sub UserDisplayMessageHTML_Close()
    IEWindowCancel = True
    IEWindowWait = False
End Sub
'------------------------------------------------------

'------------------------------------------------------
' Function UserDisplayMessageHTML_PauseResume
' To Pause or resume the script.
'------------------------------------------------------
Sub UserDisplayMessageHTML_PauseResume()
    If IEWindowPause Then
        IEWindowPause = False
    Else
        IEWindowPause = True
    End If
End Sub
'------------------------------------------------------

'------------------------------------------------------
' Function PingSystemToDisableLock
' Sends a keystroke so that system does not lock automatically.
'------------------------------------------------------
Sub PingSystemToDisableLock()
    WSHShell.SendKeys "{SCROLLLOCK}"
End Sub
'------------------------------------------------------

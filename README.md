<div align="center">

## Add Error log to your projects \- no changing code\!


</div>

### Description

This is a drop in replacement for the VB Message box routine (MsgBox). It will log to a file all messages that you display to the user that are marked with vbCritical. This routine also expands the standard VB Messagebox by giving you the developer the ability to log to file, even if you don't set vbCritical.
 
### More Info
 
Same as MsgBox

Insert this code into a module, will not work properly if placed in a class


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Monroe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-monroe.md)
**Level**          |Intermediate
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-monroe-add-error-log-to-your-projects-no-changing-code__1-6644/archive/master.zip)





### Source Code

```
Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional HelpFile As String, Optional Context As Single, Optional LogToFile As Boolean = False) As VbMsgBoxResult
 Dim strErrorLog As String
 Dim iFileHandle As Integer
 Dim strErrorTitle As String
 Dim iResult As Integer
 iFileHandle = FreeFile
 strErrorTitle = App.EXEName & " : " & Title
 strErrorLog = App.Path & "\" & App.EXEName & ".log"
 ' Force error loging on all critical messages
 If (Buttons And vbCritical) Then
 LogToFile = True
 End If
 ' if the user has choosen to log, or it's a critical message, log it
 If LogToFile = True Then
 Open strErrorLog For Append As #iFileHandle
 Print #iFileHandle, Now, Prompt
 Close #iFileHandle
 End If
 ' Call the real message box routine
 iResult = VBA.MsgBox(Prompt, Buttons, strErrorTitle, HelpFile, Context)
 MsgBox = iResult
End Function
```


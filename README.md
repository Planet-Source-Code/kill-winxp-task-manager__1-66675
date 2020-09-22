<div align="center">

## Kill WinXP Task Manager


</div>

### Description

After searching and searching PSC I decided I would put together a little something to help WinXP users and that Task Manager issue. This code runs every 10 milliseconds. (to my knowledge there is no side effects to running this timer every 10 milliseconds.) This timer will actually find and close the Windows Task Manager if it becomes active. Any Feedback would be nice. This is my first code submission.
 
### More Info
 
*COULD* Be a hit on resources due to the fact that its running code in a timer routine every 10 milliseconds. Although I did not encounter any.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[NULL](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kill-winxp-task-manager__1-66675/archive/master.zip)





### Source Code

```
'Simply put this goes at the top of your new form.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
Dim winHwnd As Long, RetVal As Long
'declares(experienced coders will know more about these.)
'Then simply make a timer or copy this code and create a timer with the same name on your application.
Private Sub TestTask_Timer()
winHwnd = FindWindow(vbNullString, "Windows Task Manager") 'Simply put it makes sure that task manager has been opened.
 If winHwnd <> 0 Then
  PostMessage winHwnd, WM_CLOSE, 0&, 0&
 Else
  'Doesn't Exit? Then
  'Do Nothing, Technically this is a loop I created with a timer.
 End If
frmMain.SetFocus
End Sub
'then run your app and hit "CTRL-ALT-DELETE" if everything is running correctly you shouldn't even have time to focus on the manager it will have closed to quickly.
```


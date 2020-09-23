<div align="center">

## \* Ultimate Window Handler\! \*


</div>

### Description

This is the ultimate window handler. This can

*Hide a window*

*Show a window*

*Minimize Window*

*Maximize Window*

*Close Window*
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Evans](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-evans.md)
**Level**          |Unknown
**User Rating**    |2.7 (30 globes from 11 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-evans-ultimate-window-handler__1-1655/archive/master.zip)

### API Declarations

```
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
```


### Source Code

```
Sub WindowHandle(win,cas as long)
'by storm
'Case 0 = CloseWindow
'Case 1 = Show Win
'Case 2 = Hide Win
'Case 3 = Max Win
'Case 4 = Min Win
Select Case cas
Case 0:
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
Case 1:
X = ShowWindow(win, SW_SHOW)
Case 2:
X = ShowWindow(win, SW_HIDE)
Case 3:
X = ShowWindow(win, SW_MAXIMIZE)
Case 4:
X = ShowWindow(win, SW_MINIMIZE)
End Select
'any questions e-mail me at storm@n2.com
End Sub
```


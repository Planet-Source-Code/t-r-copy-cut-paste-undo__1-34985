<div align="center">

## Copy/Cut/Paste/Undo


</div>

### Description

I was just poking around this site, and noticed that a lot of folks were making the copy/paste/cut/undo functions more difficult then they really needs to be. Rather then writing your own functions, make use of the SendMessage API and let Windows do the work for you.
 
### More Info
 
To try out the code, simply create a form with a Richtextbox named Richtextbox1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[T\.R\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/t-r.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/t-r-copy-cut-paste-undo__1-34985/archive/master.zip)

### API Declarations

```
Private Const EM_UNDO = &HC7
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Private Declare Function SendMessage Lib _
      "USER32.DLL" Alias "SendMessageA" _
      (ByVal hWnd As Long, _
      ByVal dwMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long
```


### Source Code

```
Private Sub Command1_Click()
'=============================================
'Defined strictly for the Richtextbox, though
'might work on others.
'=============================================
SendMessage RichTextBox1.hWnd, EM_UNDO, ByVal 0&, ByVal 0&
End Sub
Private Sub Command2_Click()
SendMessage RichTextBox1.hWnd, WM_COPY, ByVal 0&, ByVal 0&
End Sub
Private Sub Command3_Click()
SendMessage RichTextBox1.hWnd, WM_PASTE, ByVal 0&, ByVal 0&
End Sub
Private Sub Command4_Click()
SendMessage RichTextBox1.hWnd, WM_CUT, ByVal 0&, ByVal 0&
End Sub
'=================================================
'This can be used on textboxes and other controls
'=================================================
Private Sub Command5_Click()
SendMessage RichTextBox1.hWnd, WM_UNDO, ByVal 0&, ByVal 0&
End Sub
Private Sub Form_Load()
RichTextBox1.Text = "This is line 1" & vbCrLf & _
          "This is line 2" & vbCrLf & _
          "This is line 3" & vbCrLf & _
          "This is line 4" & vbCrLf
End Sub
```


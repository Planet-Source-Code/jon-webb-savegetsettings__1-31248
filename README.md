<div align="center">

## SaveGetSettings


</div>

### Description

Fed up with saving and loading your form's settings? Well save this module and with one line of code you can save its's size and position. One more line and you save _All_ the text and check boxes values. Come on this is a great time saver - how about a vote ??? Only missing option buttion settings... maybe next week.....
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Webb](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-webb.md)
**Level**          |Intermediate
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-webb-savegetsettings__1-31248/archive/master.zip)





### Source Code

```
Public Sub SaveValues(frmForm As Form)
'
' usage = place:
' Call SaveValues(Me)
' in the Form_Unload event
' this will save all checkbox and textbox settings
'
Dim ctlControl As Object
On Error Resume Next
For Each ctlControl In frmForm.Controls
SaveSetting App.Title, "Settings", ctlControl.Name, ctlControl.Value
'check boxes.....
SaveSetting App.Title, "Settings", ctlControl.Name, ctlControl.Text
DoEvents
Next ctlControl
End Sub
Public Sub SavePositions(frmForm As Form)
'
' usage = place:
' Call SavePositions(Me)
' in the Form_Unload event
' this will save the forms size and position
'
On Error Resume Next
If frmForm.WindowState = vbMinimized Then: Exit Sub 'don't want to come back minimized!!!
SaveSetting App.Title, "Settings", frmForm.Name & "top", frmForm.Top
SaveSetting App.Title, "Settings", frmForm.Name & "left", frmForm.Left
SaveSetting App.Title, "Settings", frmForm.Name & "width", frmForm.Width
SaveSetting App.Title, "Settings", frmForm.Name & "height", frmForm.Height
End Sub
Public Sub GetValues(frmForm As Form)
'
' usage = place:
' Call GetValues(Me)
' in the Form_Load event
' this will populate all checkbox and textbox settings
'
Dim ctlControl As Object
On Error Resume Next
For Each ctlControl In frmForm.Controls
'check boxes.....
ctlControl.Value = GetSetting(App.Title, "Settings", ctlControl.Name)
'text boxes
ctlControl.Text = GetSetting(App.Title, "Settings", ctlControl.Name)
DoEvents
Next ctlControl
End Sub
Public Sub GetPositions(frmAForm As Form)
'
' usage = place:
' Call GetPositions(Me)
' in the Form_Load event
' this will save the forms size and position
'
On Error Resume Next
frmAForm.Top = GetSetting(App.Title, "Settings", frmAForm.Name & "top", "30")
frmAForm.Left = GetSetting(App.Title, "Settings", frmAForm.Name & "left", "30")
frmAForm.Width = GetSetting(App.Title, "Settings", frmAForm.Name & "width")
frmAForm.Height = GetSetting(App.Title, "Settings", frmAForm.Name & "height")
End Sub
```


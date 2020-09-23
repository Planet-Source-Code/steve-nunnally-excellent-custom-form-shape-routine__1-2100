<div align="center">

## Excellent Custom Form Shape Routine


</div>

### Description

After putting a picture on your form, run this code and whatever background color you choose will be subtracted from the form leaving a very custom form shape.
 
### More Info
 
You must have a picture on your form, and you must have the correct color value for the transparent area of your form. Most paint programs usually tell you the red/green/blue values of a color.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Nunnally](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-nunnally.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-nunnally-excellent-custom-form-shape-routine__1-2100/archive/master.zip)

### API Declarations

```
Type POINTAPI
  x As Long
  y As Long
End Type
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
```


### Source Code

```
Dim rgn As Long 'global variable to keep track of region
Private Sub Form_Load()
 Dim maskcolor As Long
 maskcolor = RGB(0, 255, 0) '<----your color goes there
 TransBack 0, 0, Me.Width / 15, Me.Height / 15, maskcolor, Me.hdc, Me.hWnd
End Sub
' allows form to be moved by clicking anywhere on it
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
 DeleteObject rgn  'clean up before closing
End Sub
Private Sub TransBack(ByVal xstart As Long, ByVal ystart As Long, _
    ByVal xend As Long, ByVal yend As Long, ByVal bgcolor As Long, _
    ByVal thdc As Long, ByVal thWnd As Long)
 Dim rgn2 As Long, rgn3 As Long, rgn4 As Long
 Dim x1 As Long, y1 As Long, i As Long, j As Long, tj As Long
 rgn = CreateRectRgn(0, 0, 0, 0) 'create some region buffers
 rgn2 = CreateRectRgn(0, 0, 0, 0)
 rgn3 = CreateRectRgn(0, 0, 0, 0)
 ' this loop picks out the transparent colors,
 ' there MUST be three loops or Windows has a hard
 ' time handling the complex regions
 i = xstart
 x1 = (xend - xstart) + 1: y1 = (yend - ystart) + 1
 Do While i < x1
 j = ystart
 Do While j < y1
  If GetPixel(thdc, i, j) <> bgcolor Then
  tj = j
  Do While GetPixel(thdc, i, j + 1) <> bgcolor
   j = j + 1
   If j = y1 Then Exit Do
  Loop
  rgn4 = CreateRectRgn(i, tj, i + 1, j + 1)
  CombineRgn rgn3, rgn2, rgn2, 5
  CombineRgn rgn2, rgn4, rgn3, 2
  DeleteObject rgn4
  End If
  j = j + 1
 Loop
 CombineRgn rgn3, rgn, rgn, 5
 CombineRgn rgn, rgn2, rgn3, 2
 DoEvents
 i = i + 1
 Loop
 DeleteObject rgn2
 SetWindowRgn thWnd, rgn, True
End Sub
```


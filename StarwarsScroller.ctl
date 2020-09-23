VERSION 5.00
Begin VB.UserControl StarwarsScroller 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00000000&
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ToolboxBitmap   =   "StarwarsScroller.ctx":0000
   Begin VB.PictureBox PicHidden 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillColor       =   &H00404040&
      FillStyle       =   0  'Ausgefüllt
      FontTransparent =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   2625
      Left            =   0
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox PicFront 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   0
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   0
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "StarwarsScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Starwars Scroller ActiveX
' by Scythe scythe@cablenet.de

Option Explicit
Private Const Version As String = "1.0"

'Needed to Clear the Hidden Picture fastest way
'Thx to Carles for this and the GetTickCout :o)
'Blit Black was new for me
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Const BLACKNESS = &H42

'Timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Faster than Print
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Convert Picture to Array and back
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type
Private Const DIB_RGB_COLORS As Long = 0


Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
'Why Public ?
'So we only set it one time = MORE SPEED
Dim buf() As RGBQUAD            'Hold the Picture
Dim buf2() As RGBQUAD           'Hold the Picture


Dim Lines() As String 'Maximal 200 Lines of Text
Dim Ctr As Integer 'Line Counter
Dim YPos As Long 'Actual Position

Dim SpeedUp As Integer
Dim SpeedDown As Integer
Dim ColorMulti As Single

'For Control
Dim TxtFile As String
Dim TxtAsFile As Boolean
Dim TxtString As String
Dim Speed As Integer
Dim Scrolling As Boolean
Dim Trans As Boolean
Dim Txt As TextBox


'Resize the Pictures to fit
Private Sub UserControl_Resize()
 Dim i As Long
 Dim Stp As Single
 Dim Col As Single

 'Create a gradient for usercontrol
 'need this if transparent is on
 Stp = 255 / UserControl.ScaleHeight
 For i = 0 To UserControl.ScaleHeight
  UserControl.Line (0, i)-(UserControl.ScaleWidth, i), RGB(Col, Col, 0), B
  Col = Col + Stp
 Next i

 If UserControl.Height > UserControl.Width / 2 Then
  UserControl.Height = UserControl.Width / 2
 End If
 PicFront.Width = UserControl.ScaleWidth
 PicFront.Height = UserControl.ScaleHeight
 PicHidden.Height = PicFront.Height
 PicHidden.Width = PicFront.Width
 PicFront.Top = UserControl.ScaleHeight - PicFront.Height

 'if transparent & usercontrol is not in design
 If Trans And Ambient.UserMode Then
  PicFront.Left = PicFront.Width
  UserControl.MaskPicture = PicFront.Image
 End If
End Sub

Private Sub UserControl_Terminate()
 'Clean Up
 Erase buf()
 Erase buf2()
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
 Call .WriteProperty("Textfile", TxtFile)
 Call .WriteProperty("Text", TxtString, "")
 Call .WriteProperty("TextAsFile", TxtAsFile, True)
 Call .WriteProperty("TextFont", PicHidden.Font, Ambient.Font)
 Call .WriteProperty("Transparent", Trans, False)
 Call .WriteProperty("ScrollSpeed", Speed, 40)
 End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
 Set PicHidden.Font = .ReadProperty("TextFont", Ambient.Font)
 TxtString = .ReadProperty("Text", "")
 Trans = .ReadProperty("Transparent", False)
 Speed = .ReadProperty("ScrollSpeed", 40)
 TxtAsFile = .ReadProperty("TextAsFile", True)
 End With
End Sub

'Get/Set the textpath
Public Property Get Textfile() As String
Attribute Textfile.VB_ProcData.VB_Invoke_Property = "PropPageSelectFile;Text"
Textfile = TxtFile
End Property
Public Property Let Textfile(ByVal NewName As String)
TxtFile = NewName
PropertyChanged "Textfile"
End Property

'G***************************************
Public Property Get TextAsFile() As Boolean
Attribute TextAsFile.VB_ProcData.VB_Invoke_Property = ";Text"
TextAsFile = TxtAsFile
End Property
Public Property Let TextAsFile(ByVal NewBool As Boolean)
TxtAsFile = NewBool
PropertyChanged "TextAsfile"
End Property

'**************************************
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "PropertyText;Text"
Text = TxtString
End Property
Public Property Let Text(ByVal NewText As String)
TxtString = NewText
PropertyChanged "Text"
End Property





'Get/Set the textfont
Public Property Get TextFont() As Font
Set TextFont = PicHidden.Font
End Property
Public Property Set TextFont(ByVal NewFont As Font)
Set PicHidden.Font = NewFont
PropertyChanged "TextFont"
End Property

'Transparent
Public Property Get Transparent() As Boolean
Transparent = Trans
End Property
Public Property Let Transparent(ByVal NewTrans As Boolean)
Trans = NewTrans
PropertyChanged "Transparent"
End Property

'Speed
Public Property Get ScrollSpeed() As Integer
ScrollSpeed = Speed
End Property
Public Property Let ScrollSpeed(ByVal NewSpeed As Integer)
Speed = NewSpeed
PropertyChanged "ScrollSpeed"
End Property


'Needed for the About
Public Static Property Let About(sAbout As String)
Attribute About.VB_ProcData.VB_Invoke_PropertyPut = "PropPageAbout"
'
End Property

Public Property Get About() As String
About = "Starwars Scroller"
End Property



'Public Subs the User can call from his Code
Public Sub StartScroll()
 Scrolling = True
 InitScroller
End Sub
Public Sub StopScroll()
 Scrolling = False
End Sub
'Let the User check if the Scroller is running
Public Function IsScrolling() As Boolean
 IsScrolling = Scrolling
End Function


Private Sub InitScroller()
 Dim t As Long 'Timer
 Dim t1 As Long 'Timer SpeedChek


 Dim TmpX As Long
 Dim TmpZ As Long
 Dim TmpTxt As String

 'Hide Frontpicture if Transparent is on
 'Wont work if i turn it invisible cause i use invisible for quit
 If Trans Then
  PicFront.Left = PicFront.Left + PicFront.Width
 Else
  PicFront.Left = 0
 End If

 Ctr = 1

 If TextAsFile = True Then
  'Load the Text
  If Dir$(TxtFile) = "" Or TxtFile = "" Then
   MsgBox "Textfile not Found", vbCritical, "Starwars Scroller"
   Scrolling = False
   Exit Sub
  End If
  Open (TxtFile) For Input As #1
   Do Until EOF(1)
    'Redim the Lines to hold the text
    ReDim Preserve Lines(1 To Ctr)
    Line Input #1, Lines(Ctr)
    Ctr = Ctr + 1
   Loop
  Close
 Else
  'Use String Text
  TmpTxt = TxtString
  TmpX = 1
  Do Until InStr(TmpX, TmpTxt, vbCrLf) = 0
   TmpZ = TmpX
   TmpX = InStr(TmpX, TmpTxt, vbCrLf) + 2
   Ctr = Ctr + 1
   ReDim Preserve Lines(1 To Ctr)
   Lines(Ctr) = Mid$(TmpTxt, TmpZ, TmpX - TmpZ - 2)
  Loop
  'Last line without Enter ?
  If TmpX < Len(TmpTxt) Then
   Lines(Ctr) = Mid$(TmpTxt, TmpX, Len(TmpTxt) - TmpX)
   Ctr = Ctr + 1
  End If
 End If

 Ctr = Ctr - 1
 YPos = PicHidden.Height

 'Define the size of our Picture
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicHidden.ScaleWidth
 .biHeight = PicHidden.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicHidden.ScaleWidth * PicHidden.ScaleHeight
 End With

 'Resize the Arrays to hold the Pictures Data
 ReDim buf(0 To PicHidden.ScaleWidth - 1, 0 To PicHidden.ScaleHeight)
 ReDim buf2(0 To PicHidden.ScaleWidth - 1, 0 To PicHidden.ScaleHeight)

 'Create the ColorMulti-plicator
 ColorMulti = 255 / PicFront.Height

 'Speed Test
 'Try to speed Up/Down
 'The Scroller should look the same on every PC
 'If u calculate this every call the scroller wont
 'slowdown/speedup after starting/ending another programm

 If Speed = 0 Then Speed = 1
 t = GetTickCount 'Get the Time
 Scroll 'Do one Scroll
 t1 = GetTickCount 'Get the new Time

 'If the Scroller is to slow then
 'Speed it up by Increasing the Lines/Scroll
 SpeedUp = Int((t1 - t) / Speed + 0.99999)

 'Now Calculate the Speed Down
 'The Actual Speed = Time1 - Time / SpeedUp
 SpeedDown = Speed * (Speed / ((t1 - t) / SpeedUp))

Do
If GetTickCount - t > SpeedDown Then '=frame delay Increase to slow down
 t = GetTickCount
 Scroll
 DoEvents

 'User has closed the window without disabling the scroll
 'Now PicFront.Visible = False and Scrolling to
 If PicFront.Visible = False Then Scrolling = False
End If
Loop While Scrolling

 'Clear Picture
 BitBlt PicFront.hdc, 0, 0, PicFront.ScaleWidth, PicFront.ScaleHeight, 0, 0, 0, BLACKNESS
 If Trans Then
  UserControl.MaskPicture = Nothing
 End If
End Sub
Private Sub Scroll()
 Dim i As Long
 Dim x As Long
 Dim y As Single
 Dim z As Long
 Dim CurX As Long
 Dim CurY As Long

 With PicHidden
On Error Resume Next
'Clear Picture
'.Cls
If Trans Then
 'UserControl.MaskColor = &H0&
 UserControl.MaskPicture = PicFront.Image
End If
BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, 0, BLACKNESS

'end if
'Draw Credits on the hidden Picture
For i = 1 To Ctr Step 1
 'Set position for this line
 CurY = YPos + (i * .TextHeight(Lines(i)))
 If CurY > 0 Then
  CurX = (.ScaleWidth / 2) - (.TextWidth(Lines(i)) / 2)

  'Set Color to fade out the text
  If Trans Then
   .ForeColor = 100
  Else
   .ForeColor = RGB(CurY * ColorMulti, CurY * ColorMulti, 0)
  End If
  'exit if we reached the pictures end
  If CurY > .ScaleHeight Then Exit For

  'Print Text
  TextOut .hdc, CurX, CurY, Lines(i), Len(Lines(i))
 End If
Next i
'stop loop if the scroller is done
If i >= Ctr And YPos < -PicHidden.Height * 4 Then
 Scrolling = False 'To end after the first loop
 ' YPos = PicHidden.Height 'To repeat the whole thing
End If


'Get the Picture
GetDIBits .hdc, .Image.Handle, 0, Binfo.bmiHeader.biHeight, buf(0, 0), Binfo, DIB_RGB_COLORS


'Now copy the text to our visible picture
'Shrink the text from bottom to top 2 pixels/line
'DIB Arrays allways work bottom to top not like we know it top to bottom
'I dont use Scale.Height/Scale.Width cause its slower than using the real Size
For i = 0 To .ScaleHeight
 z = x 'Set our ne Startingpoint (x pixels  from the left border)
 'Z = Position on the visible Picture
 'Y = Position on the invissible Picture
 For y = 0 To .ScaleWidth Step (.ScaleWidth / (.ScaleWidth - x - x))
  'Copy only Red and Green (Blue = Allways 0)
  buf2(z, i).rgbGreen = buf(y, i).rgbGreen
  buf2(z, i).rgbRed = buf(y, i).rgbRed
  z = z + 1 'Move to the right
 Next y
 x = x + 1 'Our picture allways get smaller / Every Line
Next i

'Set the new Picture
SetDIBits PicFront.hdc, PicFront.Image.Handle, 0, Binfo.bmiHeader.biHeight, buf2(0, 0), Binfo, DIB_RGB_COLORS
'Refresh to make it vissible
PicFront.Refresh
'UserControl.Refresh

'Move the Text up one Pixel
'Increase to speed up
YPos = YPos - SpeedUp
End With
End Sub


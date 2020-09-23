VERSION 5.00
Object = "*\ASWScroller.vbp"
Begin VB.Form TestForm 
   BackColor       =   &H8000000A&
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Scroller Test"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TestForm.frx":0000
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin SWscroller.StarwarsScroller StarwarsScroller1 
      Height          =   2535
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4471
      Textfile        =   ""
      Text            =   $"TestForm.frx":3112
      TextAsFile      =   0   'False
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      ScrollSpeed     =   10
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start / Stop"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 'Check if the Scroller is Active
 If StarwarsScroller1.IsScrolling Then
  StarwarsScroller1.StopScroll
 Else
  StarwarsScroller1.StartScroll
 End If
End Sub
Private Sub Form_Load()
 StarwarsScroller1.Textfile = App.Path & "\Scroller.txt"
End Sub

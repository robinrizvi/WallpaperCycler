VERSION 5.00
Begin VB.Form tnail 
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Interval        =   1
      Left            =   2760
      Top             =   2640
   End
   Begin VB.Image imgprev 
      Height          =   330
      Left            =   240
      Picture         =   "tnail.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Browse previous image"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image imgnext 
      Height          =   330
      Left            =   2160
      Picture         =   "tnail.frx":5491
      Stretch         =   -1  'True
      ToolTipText     =   "Browse next image"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image imgmid 
      Height          =   375
      Left            =   1200
      Picture         =   "tnail.frx":A93D
      Stretch         =   -1  'True
      ToolTipText     =   "Set as wallpaper"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image tvtnail 
      Height          =   1575
      Left            =   270
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2325
   End
   Begin VB.Image tv 
      Height          =   2340
      Left            =   0
      Picture         =   "tnail.frx":E1E0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2850
   End
End
Attribute VB_Name = "tnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmrcounter As Byte

Private Sub Form_Load()
    SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
    Me.BackColor = vbCyan
    SetTrans Me, 0, Me.BackColor
    Me.Move (Screen.Width - Me.Width), (Screen.Height - Me.Height - 500)
    Me.Show
    tmrcounter = 0
    tmr.Enabled = True
End Sub

Private Sub imgmid_Click()
Dim countinfpath As String
If disable = True Then
Call tnailupdatedisplay
Exit Sub
End If
countinfpath = App.Path & "\count.inf"
Open countinfpath For Output As #1
Print #1, tnailloopcounter
Close #1
If folderpath = "" Then
Call updatesetting
Exit Sub
End If
Call clickchangewallpaper(False, folderpath)
End Sub

Private Sub imgnext_Click()
If disable = True Then
Call tnailupdatedisplay
Exit Sub
End If
tnailloopcounter = tnailloopcounter + 1
If tnailloopcounter < 0 Then tnailloopcounter = 0
Call tnailupdatedisplay
End Sub

Private Sub imgprev_Click()
If disable = True Then
Call tnailupdatedisplay
Exit Sub
End If
tnailloopcounter = tnailloopcounter - 1
If tnailloopcounter < 0 Then tnailloopcounter = 0
Call tnailupdatedisplay
End Sub

Private Sub tmr_Timer()
SetTrans Me, tmrcounter, Me.BackColor
If tmrcounter = 255 Then
tmr.Enabled = False
Exit Sub
End If
tmrcounter = tmrcounter + 1
End Sub

Private Sub tvtnail_Click()
imgmid_Click
End Sub

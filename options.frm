VERSION 5.00
Begin VB.Form opt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "options.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox tnail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1080
      TabIndex        =   5
      Top             =   3060
      Width           =   200
   End
   Begin VB.CheckBox disable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Disable Wallpaper Cycler"
      Top             =   2535
      Width           =   200
   End
   Begin VB.CheckBox rand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Randomize the wallpaper display"
      Top             =   2000
      Width           =   200
   End
   Begin VB.CheckBox strup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Auto start with windows"
      Top             =   1470
      Width           =   200
   End
   Begin VB.Label lblapply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblclose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim tmptnailloopcounter As Long
tmptnailloopcounter = tnailloopcounter
Call updatesetting
If tnailactivate = True Then
tnailloopcounter = tmptnailloopcounter
Call tnailupdatedisplay
End If
End Sub

Private Sub lblapply_Click()
Dim tmptnailloopcounter As Long
If strup.value = 0 Then Call regsetvalue("Software\WallpaperCycler", "startup", "n")
If strup.value = 1 Then Call regsetvalue("Software\WallpaperCycler", "startup", "y")
If rand.value = 0 Then Call regsetvalue("Software\WallpaperCycler", "randomize", "n")
If rand.value = 1 Then Call regsetvalue("Software\WallpaperCycler", "randomize", "y")
If disable.value = 0 Then Call regsetvalue("Software\WallpaperCycler", "disable", "n")
If disable.value = 1 Then Call regsetvalue("Software\WallpaperCycler", "disable", "y")
If tnail.value = 0 Then Call regsetvalue("Software\WallpaperCycler", "tnail", "n")
If tnail.value = 1 Then Call regsetvalue("Software\WallpaperCycler", "tnail", "y")
tmptnailloopcounter = tnailloopcounter
Call updatesetting
If tnailactivate = True Then
tnailloopcounter = tmptnailloopcounter
Call tnailupdatedisplay
End If
Unload Me
End Sub

Private Sub lblclose_Click()
Unload Me
End Sub

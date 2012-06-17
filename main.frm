VERSION 5.00
Begin VB.Form main 
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Image bmpconverter 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu rclkmenu 
      Caption         =   "rclkmenu"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu help 
         Caption         =   "Help"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu folder 
         Caption         =   "Import Folder"
      End
      Begin VB.Menu options 
         Caption         =   "Options"
      End
      Begin VB.Menu ontop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub about_Click()
abt.Show
End Sub

Private Sub exit_Click()
UnloadAllForms Me.Name
Unload Me
End Sub

Private Sub folder_Click()
Dim tmpfolderpath As String, tmptnailloopcounter As Long
tmpfolderpath = BrowseForFolder(hwnd, "Please select a folder to monitor for image files")
If tmpfolderpath <> "" Then
Call regsetvalue("Software\WallpaperCycler", "folder", tmpfolderpath)
Open (App.Path & "\count.inf") For Output As #1
Print #1, "0"
Close #1
End If
tmptnailloopcounter = tnailloopcounter
Call updatesetting
If tnailactivate = True Then
tnailloopcounter = tmptnailloopcounter
Call tnailupdatedisplay
End If
End Sub

Private Sub help_Click()
hlp.Show
End Sub

Private Sub ontop_Click()
Dim tmptnailloopcounter As Long
ontop.Checked = Not ontop.Checked
tmptnailloopcounter = tnailloopcounter
Call updatesetting
If tnailactivate = True Then
tnailloopcounter = tmptnailloopcounter
Call tnailupdatedisplay
End If
End Sub

Private Sub options_Click()
opt.Show
End Sub

Private Sub Form_Load()
Call ShellTrayAdd
Call ShellTrayModifyTip("Wallpaper Cycler", ("Click me to cycle wallpapers. Right click on me and click Options to configure me. Have FUN"))
Call updatesetting
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As Long
   
   ' The Form_MouseMove is intercepted to give systray mouse events.
   If Me.ScaleMode = vbPixels Then
      msg = x
   Else
      msg = x / Screen.TwipsPerPixelX
   End If
      
    Select Case msg
        Case WM_RBUTTONUP
            PopupMenu rclkmenu
        Case WM_LBUTTONUP
            If disable = False And folderpath <> "" Then Call clickchangewallpaper(rand, folderpath)
            Call updatesetting
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShellTrayRemove
End
End Sub


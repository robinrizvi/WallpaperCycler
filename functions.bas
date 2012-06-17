Attribute VB_Name = "functions"
Option Explicit

Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_SENDWININICHANGE = &H2
Private Const APP_SYSTRAY_ID = 999
Private Const NOTIFYICON_VERSION = &H3
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETVERSION = &H4
Private Const NIS_SHAREDICON = &H2
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504
Private Const ERROR_NONE = 0
Private Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONUP = &H202
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private NOTIFYICONDATA_SIZE As Long
Public rand As Boolean, disable As Boolean, tnailloopcounter As Long, folderpath As String, tnailactivate As Boolean

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenCurrentUser Lib "Advapi32" (ByVal samDesired As Integer, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

        ' Determine the size and type of data to be read
        lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        Else
            vValue = Empty
        End If
        
        QueryValueEx = lrc
       
End Function

Private Function QueryValue(sKeyName As String, sValueName As String) As String
       Dim lRetVal As Long         'result of the API functions
       Dim vValue As Variant      'setting of queried value
       Dim temp1, temp2, temp3, temp4

   
    temp1 = RegOpenCurrentUser(1, temp2)
    temp4 = RegCreateKey(temp2, sKeyName, temp3)
    lRetVal = QueryValueEx(temp3, sValueName, vValue)
    RegCloseKey (temp2)
    RegCloseKey (temp3)
    QueryValue = vValue
End Function
   
Public Function regsetvalue(subkey As String, field As String, value As String)
Dim temp1, temp2, temp3, temp4, temp5
temp1 = RegOpenCurrentUser(1, temp2)
temp4 = RegCreateKey(temp2, subkey, temp3)
temp5 = RegSetValueEx(temp3, field, 0, 1, value, Len(value))
RegCloseKey (temp2)
RegCloseKey (temp3)
End Function

Private Function regdelvalue(subkey As String, field As String)
Dim temp1, temp2, temp3, temp4, temp5
temp1 = RegOpenCurrentUser(1, temp2)
temp4 = RegCreateKey(temp2, subkey, temp3)
temp5 = RegDeleteValue(temp3, field)
RegCloseKey (temp2)
RegCloseKey (temp3)
End Function

Public Sub ShellTrayAdd()
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
  
   With nid
   
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = main.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = main.Icon
      .uCallbackMessage = WM_MOUSEMOVE
      .szTip = "Wallpaper Cycler" & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      
   End With
   
  Call Shell_NotifyIcon(NIM_ADD, nid)
   
  '... and inform the system of the NOTIFYICON version in use
  Call Shell_NotifyIcon(NIM_SETVERSION, nid)
       
End Sub

Public Sub ShellTrayRemove()

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
      
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = main.hwnd
      .uID = APP_SYSTRAY_ID
      .uCallbackMessage = WM_MOUSEMOVE
   End With
   
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub

Public Sub ShellTrayModifyTip(infotitle As String, infobody As String)

   Dim nid As NOTIFYICONDATA

   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = main.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = 1
      .uCallbackMessage = WM_MOUSEMOVE
      .szInfoTitle = infotitle & vbNullChar
      .szInfo = infobody & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub

Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub

Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   Const sDLLFile As String = "shell32.dll"
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   If nBufferSize > 0 Then
        ReDim bBuffer(nBufferSize - 1) As Byte
        Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
        If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
            CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
            IsShellVersion = nVerMajor >= version
        End If
   End If
  
End Function

Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0)
    Dim lStyle As Long
    lStyle = GetWindowLong(oForm.hwnd, GWL_EXSTYLE)
    If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
        SetWindowLong oForm.hwnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes oForm.hwnd, lColor, bytAlpha, LWA_COLORKEY Or LWA_ALPHA
End Sub

Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
If Topmost = True Then
    SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
    SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    SetTopMostWindow = False
End If
End Function

Public Sub UnloadAllForms(Optional FormToIgnore As String = "")
  Dim f As Form
  For Each f In Forms
    If f.Name <> FormToIgnore Then
      Unload f
      Set f = Nothing
    End If
  Next f
End Sub

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
      
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
      
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath

End Function

Public Function updatesetting()
Dim queryval As String, setupval As String, exepath As String, countinfpath As String, fileexistchk As Boolean, fso, countinfval, lr As Long
'check startup settings
queryval = QueryValue("Software\WallpaperCycler", "startup")
If queryval = "" Then
setupval = "y"
Call regsetvalue("Software\WallpaperCycler", "startup", setupval)
queryval = QueryValue("Software\WallpaperCycler", "startup")
End If
exepath = App.Path & "\" & App.EXEName & ".exe"
If queryval = "y" Then
Call regsetvalue("Software\Microsoft\Windows\CurrentVersion\Run", "WallpaperCycler", exepath)
opt.strup.value = 1
Else
Call regdelvalue("Software\Microsoft\Windows\CurrentVersion\Run", "WallpaperCycler")
opt.strup.value = 0
End If
'done startup settings

'check randomize settings
queryval = QueryValue("Software\WallpaperCycler", "randomize")
If queryval = "" Then
setupval = "n"
Call regsetvalue("Software\WallpaperCycler", "randomize", setupval)
queryval = QueryValue("Software\WallpaperCycler", "randomize")
End If
If queryval = "n" Then
opt.rand.value = 0
rand = False
Else
opt.rand.value = 1
rand = True
End If
'done randomize settings

'check disable settings
queryval = QueryValue("Software\WallpaperCycler", "disable")
If queryval = "" Then
setupval = "n"
Call regsetvalue("Software\WallpaperCycler", "disable", setupval)
queryval = QueryValue("Software\WallpaperCycler", "disable")
End If
If queryval = "n" Then
opt.disable.value = 0
disable = False
Else
opt.disable.value = 1
disable = True
Call ShellTrayModifyTip("Wallpaper Cycler", "Wallpaper Cycler is disabled. Check options")
End If
'done disable settings

'check folder settings
queryval = QueryValue("Software\WallpaperCycler", "folder")
If queryval = "" Then
Call ShellTrayModifyTip("Wallpaper Cycler", "No folder is selected to monitor. Please right-click and select import folder.")
Else
folderpath = queryval
End If
'done folder settings

'check count.inf settings
Set fso = CreateObject("Scripting.FileSystemObject")
countinfpath = App.Path & "\count.inf"
fileexistchk = fso.fileexists(countinfpath)
If fileexistchk = False Then
Open countinfpath For Output As #1
Print #1, "0"
Close #1
End If
Open countinfpath For Input As #1
    On Error GoTo reset
    Input #1, countinfval
    Close #1
If fileexistchk = False Or Val(countinfval) < 0 Or countinfval = "" Then
reset:      Close #1
            Open countinfpath For Output As #1
            Print #1, "0"
            Close #1
End If
'done checking count.inf

'check tnail settings
queryval = QueryValue("Software\WallpaperCycler", "tnail")
If queryval = "" Then
setupval = "y"
Call regsetvalue("Software\WallpaperCycler", "tnail", setupval)
queryval = QueryValue("Software\WallpaperCycler", "tnail")
End If
If queryval = "n" Then
opt.tnail.value = 0
tnailactivate = False
Unload tnail
Else
tnailactivate = True
If main.ontop.Checked = True Then lr = SetTopMostWindow(tnail.hwnd, True) Else lr = SetTopMostWindow(tnail.hwnd, False)
opt.tnail.value = 1
    countinfpath = App.Path & "\count.inf"
    Open countinfpath For Input As #1
    Input #1, countinfval
    Close #1
    tnailloopcounter = Val(countinfval) - 1
If tnailloopcounter < 0 Then tnailloopcounter = 0
tnail.Show
If folderpath <> "" And disable = False Then Call tnailupdatedisplay
If disable = True Then tnail.tvtnail.Picture = LoadPicture()
End If
'done tnail settings
End Function

Public Function tnailupdatedisplay()
Dim fso, fso1, fso2
Dim countfiles As Long, i As Long, exten_var As String, wallpapername As String, wallpaperpath As String, flag As Boolean

If folderpath = "" Or disable = True Then
tnail.tvtnail.Picture = LoadPicture()
Call updatesetting
Exit Function
End If

'count number of files in a folder
Set fso = CreateObject("Scripting.FileSystemObject")
On Error GoTo endline
Set fso1 = fso.GetFolder(folderpath)
Set fso2 = fso1.Files
countfiles = fso2.Count
countfiles = countfiles - 1
'end counting

'check for valid files to avoid infinte looping due to code below
exten_var = Right((Dir((folderpath & "\*.*"))), 3)
If exten_var = "bmp" Or exten_var = "BMP" Or exten_var = "jpg" Or exten_var = "JPG" Or exten_var = "gif" Or exten_var = "GIF" Then flag = 1
For i = 1 To countfiles
exten_var = Right((Dir()), 3)
On Error Resume Next
If exten_var = "bmp" Or exten_var = "BMP" Or exten_var = "jpg" Or exten_var = "JPG" Or exten_var = "gif" Or exten_var = "GIF" Then flag = 1
Next
If flag = 0 Then
Call ShellTrayModifyTip("Wallpaper Cycler", "The selected folder does not contain image files. Please choose another folder.")
tnail.tvtnail.Picture = LoadPicture()
Exit Function
End If
'check completed

If tnailloopcounter = countfiles Or tnailloopcounter > countfiles Then
tnailloopcounter = 0
Call tnailupdatedisplay
Exit Function
End If

wallpapername = Dir((folderpath & "\*.*"))
For i = 1 To tnailloopcounter
wallpapername = Dir()
Next
wallpaperpath = folderpath & "\" & wallpapername
exten_var = Right(wallpapername, 3)
If exten_var <> "bmp" And exten_var <> "BMP" And exten_var <> "jpg" And exten_var <> "JPG" And exten_var <> "gif" And exten_var <> "GIF" Then
        tnailloopcounter = tnailloopcounter + 1
        Call tnailupdatedisplay
        Exit Function
End If
tnail.tvtnail.Picture = LoadPicture(wallpaperpath)

Exit Function
endline: tnail.tvtnail.Picture = LoadPicture()
Call ShellTrayModifyTip("Wallpaper Cycler", "The selected folder is inaccessible. Please choose another folder")
End Function

Public Function clickchangewallpaper(randum As Boolean, folpath As String)
Dim fileexistchk As Boolean, flag As Boolean
Dim fso, fso1, fso2, countinfval
Dim countfiles As Long, loopcounter As Long, i As Long
Dim countinfpath As String, wallpapername As String, wallpaperpath As String, exten_var As String

'count number of files in a folder
Set fso = CreateObject("Scripting.FileSystemObject")
On Error GoTo endline
Set fso1 = fso.GetFolder(folpath)
Set fso2 = fso1.Files
countfiles = fso2.Count
countfiles = countfiles - 1
'end counting

'check for valid files to avoid infinte looping due to code below
exten_var = Right((Dir((folpath & "\*.*"))), 3)
If exten_var = "bmp" Or exten_var = "BMP" Or exten_var = "jpg" Or exten_var = "JPG" Or exten_var = "gif" Or exten_var = "GIF" Then flag = 1
For i = 1 To countfiles
exten_var = Right((Dir()), 3)
On Error Resume Next
If exten_var = "bmp" Or exten_var = "BMP" Or exten_var = "jpg" Or exten_var = "JPG" Or exten_var = "gif" Or exten_var = "GIF" Then flag = 1
Next
If flag = 0 Then
Call ShellTrayModifyTip("Wallpaper Cycler", "The selected folder does not contain image files. Please choose another folder.")
Exit Function
End If
'check completed

countinfpath = App.Path & "\count.inf"
fileexistchk = fso.fileexists(countinfpath)

If fileexistchk = False Then
Open countinfpath For Output As #1
Print #1, "0"
Close #1
End If


start: Open countinfpath For Input As #1
Input #1, countinfval
Close #1
loopcounter = Val(countinfval)

If randum = True Then
Randomize   ' Initialize random-number generator.
loopcounter = Int((countfiles * Rnd) + 0)   ' Generate random value between 1 and 6.
End If

If loopcounter = countfiles Or loopcounter > countfiles Or loopcounter < 0 Then
Open countinfpath For Output As #1
Print #1, "0"
Close #1
GoTo start
End If

wallpapername = Dir((folpath & "\*.*"))
For i = 1 To loopcounter
wallpapername = Dir()
Next
wallpaperpath = folpath & "\" & wallpapername
exten_var = Right(wallpapername, 3)
If exten_var <> "bmp" And exten_var <> "BMP" And exten_var <> "jpg" And exten_var <> "JPG" And exten_var <> "gif" And exten_var <> "GIF" Then
        loopcounter = loopcounter + 1
        Open countinfpath For Output As #1
        Print #1, loopcounter
        Close #1
        GoTo start
End If
main.bmpconverter.Picture = LoadPicture(wallpaperpath)
SavePicture main.bmpconverter, (App.Path & "\main.bmp")

Call ChangeWallPaper((App.Path & "\main.bmp"))

Open countinfpath For Output As #1
Print #1, Str((loopcounter + 1))
Close #1

Exit Function
endline: Call ShellTrayModifyTip("Wallpaper Cycler", "The selected folder is inaccessible. Please choose another folder")
End Function

Private Function ChangeWallPaper(ImageFile As String)
Dim lRet As Long
On Error Resume Next

lRet = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, ImageFile, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
ChangeWallPaper = lRet <> 0 And Err.LastDllError = 0
End Function

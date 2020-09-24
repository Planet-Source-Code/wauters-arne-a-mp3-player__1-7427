VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mp3Playah"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   6240
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton OptNON 
      BackColor       =   &H00404040&
      Caption         =   "None"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   4440
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.OptionButton OptRes 
      BackColor       =   &H00404040&
      Caption         =   "Resume"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1860
      TabIndex        =   28
      Top             =   4440
      Width           =   975
   End
   Begin VB.OptionButton OptRND 
      BackColor       =   &H00404040&
      Caption         =   "Random"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdOnTop 
      Caption         =   "Keep On Top"
      Height          =   315
      Left            =   5340
      TabIndex        =   26
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton CmdSysinfo 
      Caption         =   "System Info"
      Height          =   315
      Left            =   5100
      TabIndex        =   24
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Direct Volume"
      ForeColor       =   &H0000FF00&
      Height          =   1875
      Left            =   5040
      TabIndex        =   11
      Top             =   0
      Width           =   1635
      Begin VB.CommandButton CmdMute 
         Caption         =   "&Mute"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   1500
         Width           =   1530
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   1080
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Max             =   2500
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Min             =   -5000
         Max             =   5000
         TickStyle       =   3
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Center"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00404040&
         Caption         =   "Volume    "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "0%"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   780
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "PlayList"
      ForeColor       =   &H0000FF00&
      Height          =   2115
      Left            =   5040
      TabIndex        =   17
      Top             =   1860
      Width           =   1635
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   1200
         Top             =   1260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdSaveList 
         Caption         =   "Save List"
         Height          =   375
         Left            =   60
         TabIndex        =   23
         Top             =   1680
         Width           =   1515
      End
      Begin VB.CommandButton CmdLoadList 
         Caption         =   "Load Playlist"
         Height          =   375
         Left            =   60
         TabIndex        =   22
         Top             =   1320
         Width           =   1515
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Rem A&ll"
         Height          =   375
         Left            =   60
         TabIndex        =   19
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton CmdRem 
         Caption         =   "&Rem"
         Height          =   375
         Left            =   60
         TabIndex        =   20
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.TextBox TxtTime 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1140
      Width           =   555
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "A&bout"
      Height          =   435
      Left            =   3300
      TabIndex        =   5
      Top             =   720
      Width           =   1710
   End
   Begin VB.CommandButton Cmdvol 
      Caption         =   "&Volume"
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   720
      Width           =   1700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Minimi&ze"
      Height          =   435
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   1700
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   1860
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   360
      Width           =   1710
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   360
      Width           =   1700
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1700
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6180
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   540
      TabIndex        =   9
      Top             =   1140
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4995
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4995
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00404040&
      Caption         =   "Keep On Top"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5100
      TabIndex        =   25
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Menu addmenu 
      Caption         =   "Addmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuaddfile 
         Caption         =   "File"
         Index           =   1
      End
      Begin VB.Menu mnuAdddir 
         Caption         =   "Directory"
         Index           =   2
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuplayah 
      Caption         =   "Mp3Playah"
      Begin VB.Menu mnuplay 
         Caption         =   "Play"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuaddfile2 
         Caption         =   "Add File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAdddir2 
         Caption         =   "Add Dir"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRemall 
         Caption         =   "Remove All"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuVolume 
         Caption         =   "Volume"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnumini 
         Caption         =   "Minimize"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "Keep on Top"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Load List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSavelist 
         Caption         =   "Save List"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tinseconden As Integer
Dim minuten As Integer
Dim seconden As Integer

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Check1_Click()
If Check1.Value = 0 Then
AlwaysOnTop Form1, True
Else
AlwaysOnTop Form1, False
End If
End Sub



Private Sub CmdLoadList_Click()
Dim file As String
CommonDialog2.DialogTitle = "Load your list."
   CommonDialog2.MaxFileSize = 16384
   CommonDialog2.filename = ""
   CommonDialog2.Filter = "list Files|*.mlt"
   CommonDialog2.ShowOpen     ' = 1
If CommonDialog2.filename = "" Then Exit Sub
file = CommonDialog2.filename
Dim A As String
Dim X As String
On Error GoTo Error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
List1.AddItem A$
Loop
Close 1
Exit Sub
Error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub


Private Sub CmdOnTop_Click()
    If Check3.Value = 0 Then
    AlwaysOnTop Form1, True
    Check3.Value = 1
    Else
    If Check3.Value = 1 Then
    AlwaysOnTop Form1, False
    Check3.Value = 0
    End If
    End If
End Sub

Private Sub CmdSaveList_Click()
Dim naampje As String
naampje = InputBox("Name of the list?", "ListName")
naampje = naampje & ".mlt"
Open (App.Path & "\" & naampje) For Output As #1
       Dim i%
       For i = 0 To List1.ListCount - 1
       Print #1, List1.List(i)
       Next
       Close #1
    'CommonDialog3.DialogTitle = "Save your list."
    'CommonDialog3.MaxFileSize = 16384
    'CommonDialog3.filename = ""
    'CommonDialog3.Filter = "list Files|*.mlt"
    'CommonDialog3.InitDir = App.Path
    'CommonDialog3.DefaultExt = ".mlt"
    'CommonDialog3.ShowSave

End Sub

Private Sub CmdSysinfo_Click()
Call StartSysInfo
End Sub

Private Sub CmdAbout_Click()
Form2.Show vbModal
End Sub

Private Sub CmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PopupMenu addmenu, , X + 4000, Y + 2000

End Sub

Private Sub CmdClear_Click()
List1.Clear
Text1.Text = ""
End Sub

Private Sub CmdMute_Click()
If MediaPlayer1.Mute = False Then
MediaPlayer1.Mute = True
Else
MediaPlayer1.Mute = False
End If
End Sub

Private Sub CmdPause_Click()
If List1.ListCount = 0 Then Exit Sub
If Text1.Text = "" Then Exit Sub
If CmdPause.Caption = "Pause" Then
MediaPlayer1.Pause
CmdPause.Caption = "Resume"
Else
MediaPlayer1.Play
CmdPause.Caption = "Pause"
End If
End Sub

Private Sub CmdPlay_Click()
Text1 = List1.Text
On Error Resume Next
MediaPlayer1.filename = Text1.Text
If Text1.Text <> "" Then
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration
CmdPause.Enabled = True
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub CmdRem_Click()
If List1.ListIndex = -1 Then
MsgBox "No file selected", vbExclamation, "Error"
Else
List1.RemoveItem List1.ListIndex
Text1.Text = ""
End If
End Sub

Private Sub CmdStop_Click()
MediaPlayer1.Stop
Slider1.Value = 0
Text1.Text = ""
CmdPause.Enabled = False
End Sub

Private Sub Cmdvol_Click()
On Error GoTo errorhandler
       Dim lngresult As Long
       lngresult = Shell("c:\windows\Sndvol32.exe", vbNormalFocus)
       Exit Sub
errorhandler:
    lngresult = Shell("c:\winnt\system32\Sndvol32.exe", vbNormalFocus)
End Sub

Private Sub Command1_Click()
Form1.WindowState = 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
CmdPause.Caption = "Pause"
Slider3.Value = MediaPlayer1.Volume
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub List1_DblClick()
Text1 = List1.Text
On Error Resume Next
MediaPlayer1.filename = Text1.Text
If Text1.Text <> "" Then
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration
CmdPause.Enabled = True
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)

If OptRND.Value = True Then
Randomize Timer
 MyValue = Int((List1.ListCount * Rnd))
    List1.ListIndex = MyValue
    MediaPlayer1.filename = Text1.Text
    If Text1.Text <> "" Then
        MediaPlayer1.Play
        Slider1.Max = MediaPlayer1.Duration
        CmdPause.Enabled = True
        Exit Sub
    Else
        MsgBox "No file to play", vbOKOnly, "Error"
    End If
Else
    If OptRes.Value = True Then
        List1.ListIndex = List1.ListIndex + 1
        Text1.Text = List1.Text
        MediaPlayer1.filename = Text1.Text
        MediaPlayer1.Play
        CmdPause.Enabled = True
    End If
End If
If OptNON.Value = True Then
    MediaPlayer1.Stop
End If
End Sub

Private Sub mnuAbout_Click()
CmdAbout_Click
End Sub

Private Sub mnuadddir_Click(Index As Integer)
Form3.Show vbModal
End Sub

Private Sub mnuAdddir2_Click()
Form3.Show vbModal
End Sub

Private Sub mnuaddfile_Click(Index As Integer)
CommonDialog1.DialogTitle = "Load your MP3."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.filename = ""
   CommonDialog1.Filter = "MP3 Files|*.MP3"
   CommonDialog1.ShowOpen     ' = 1
If CommonDialog1.FileTitle <> "" Then
List1.AddItem CommonDialog1.filename
Text1.Text = CommonDialog1.filename
Exit Sub
Else
Exit Sub
End If
End Sub


Private Sub mnuaddfile2_Click()
CommonDialog1.DialogTitle = "Load your MP3."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.filename = ""
   CommonDialog1.Filter = "MP3 Files|*.MP3"
   CommonDialog1.ShowOpen     ' = 1
If CommonDialog1.FileTitle <> "" Then
List1.AddItem CommonDialog1.filename
Text1.Text = CommonDialog1.filename
Exit Sub
Else
Exit Sub
End If
End Sub

Private Sub mnuExit_Click()
End
Unload Me
End Sub

Private Sub mnuLoadList_Click()
CmdLoadList_Click
End Sub

Private Sub mnumini_Click()
Command1_Click
End Sub

Private Sub mnuMute_Click()
CmdMute_Click
End Sub

Private Sub mnuOnTop_Click()
CmdOnTop_Click
End Sub

Private Sub mnuplay_Click()
CmdPlay_Click
End Sub

Private Sub mnuRemall_Click()
CmdClear_Click
End Sub

Private Sub mnuRemove_Click()
CmdRem_Click
End Sub

Private Sub mnuSavelist_Click()
CmdSaveList_Click
End Sub

Private Sub mnustop_Click()
CmdStop_Click
End Sub

Private Sub mnuVolume_Click()
Cmdvol_Click
End Sub

Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
On Error GoTo DamnYou
If Slider2.Value > -500 And Slider2.Value < 500 Then
Label4.Caption = "Center"
End If
If Slider2.Value < -500 Then
Label4.Caption = "Left"
End If
If Slider2.Value > 500 Then
Label4.Caption = "Right"
End If
MediaPlayer1.Balance = Slider2.Value
Exit Sub
DamnYou:
MsgBox "Err"
Exit Sub
End Sub


Private Sub Slider3_Scroll()
Dim pim, sha
Dim foo As Integer, poo As Integer
Label2.ForeColor = RGB(0 + Slider3.Value / 10, 0, 0)
sha = Slider3.Value - 2500
MediaPlayer1.Volume = sha
On Error GoTo hell
poo = Slider3.min
foo = Slider3.Value
Label2.Caption = foo \ 25 & " %"
hell:
Exit Sub
End Sub

Private Sub Timer1_Timer()
Slider1.Value = MediaPlayer1.CurrentPosition
tinseconden = MediaPlayer1.CurrentPosition
Dim min As Integer
Dim sec As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
If sec = "-1" Then sec = "0"
TxtTime.Text = min & ":" & sec
End Sub

Public Sub AlwaysOnTop(Form1 As Form, SetOnTop As Boolean)
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If

    SetWindowPos Form1.hwnd, lFlag, Form1.Left / Screen.TwipsPerPixelX, _
    Form1.Top / Screen.TwipsPerPixelY, Form1.Width / Screen.TwipsPerPixelX, _
    Form1.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


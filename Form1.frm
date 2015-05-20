VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"Form1.frx":0000
   ClientHeight    =   2730
   ClientLeft      =   3840
   ClientTop       =   3525
   ClientWidth     =   5985
   Icon            =   "Form1.frx":008A
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5985
   Begin VB.PictureBox Picture6 
      Height          =   2700
      Left            =   7440
      ScaleHeight     =   2640
      ScaleWidth      =   5895
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   5950
      Begin VB.CommandButton Command11 
         Caption         =   "Close"
         Height          =   375
         Left            =   4200
         TabIndex        =   32
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kill process"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   31
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Top             =   120
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   2655
      End
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   4440
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   600
      Top             =   3000
   End
   Begin VB.PictureBox Picture5 
      Height          =   2700
      Left            =   7080
      ScaleHeight     =   2640
      ScaleWidth      =   5895
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   5950
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Regkeys done !"
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DisableTaskMgr"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DisableRegistryTools"
         Height          =   195
         Left            =   840
         TabIndex        =   15
         Top             =   600
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":E57F
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DisableCMD"
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Software\Policies\Microsoft\Windows\System"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   1080
         Width           =   3285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "HKEY_CURRENT_USER"
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Software\Microsoft\Windows\CurrentVersion\Policies\System"
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   2700
      Left            =   6720
      ScaleHeight     =   2640
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5950
      Begin VB.CommandButton Command8 
         Caption         =   "OK"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5655
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Double click for open a file"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1890
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox Picture2 
      Height          =   2700
      Left            =   0
      ScaleHeight     =   2576.796
      ScaleMode       =   0  'User
      ScaleWidth      =   5895
      TabIndex        =   4
      Top             =   0
      Width           =   5950
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   2595
         TabIndex        =   37
         Top             =   720
         Width           =   2655
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Top Most ? (X+Y)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   0
            Value           =   1  'Checked
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   240
         ScaleHeight     =   195
         ScaleWidth      =   2595
         TabIndex        =   35
         Top             =   720
         Width           =   2655
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Lock (CTRL+F)"
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   2625
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Enable Regedit/TaskManager/Cmd"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   5415
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Tiny process killer"
         Height          =   375
         Left            =   3840
         TabIndex        =   34
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open another file"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Open taskmgr.exe"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Open regedit.exe"
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open explorer.exe"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   375
         Left            =   2040
         TabIndex        =   33
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   5760
         X2              =   5760
         Y1              =   2342.542
         Y2              =   117.127
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5760
         Y1              =   2342.542
         Y2              =   2342.542
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5760
         Y1              =   117.127
         Y2              =   117.127
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   2342.542
         Y2              =   117.127
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   120
         Y1              =   2342.542
         Y2              =   117.127
      End
      Begin VB.Line Line17 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   2342.542
         Y2              =   117.127
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   2342.542
         Y2              =   2342.542
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   117.127
         Y2              =   117.127
      End
      Begin VB.Label lblHw 
         Caption         =   "Handle:"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblTxt 
         Caption         =   "Titlte :"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lock (CTRL+F)"
      Height          =   195
      Left            =   840
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      Caption         =   "X : "
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      Caption         =   "Y : "
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Caption         =   "Width : "
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "Height : "
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Dim Locked As Boolean
Dim handle As Long
Dim Rec As RECT
Dim CurPosWindow As POINTAPI
Dim PrecedentLocked As Boolean

Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312
Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
'
Dim bbq As New cRegistry
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command10_Click()
ProcessTerminate (List2.List(List1.ListIndex))
ProcessList
Command10.Enabled = False
 MsgBox "" & vbCrLf & "Killed", vbOKOnly + vbInformation + vbApplicationModal, "RansomHelper"
End Sub

Private Sub Command11_Click()
Picture6.Visible = False
End Sub



Private Sub Command14_Click()
Picture6.Visible = True
Picture6.Top = 0
Picture6.Left = 0
ProcessList
End Sub

Private Sub Command2_Click()
Picture4.Visible = True
Picture4.Top = 0
Picture4.Left = 0
End Sub

Private Sub Command3_Click()
With bbq
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Policies\System" 'the directory to the key
        .ValueType = REG_DWORD
        .ValueKey = "DisableRegistryTools" 'disable regedit
        .Value = 0 '1 = disabled, 0 = enabled
        .ValueKey = "DisableTaskMgr" 'disable taskmgr
        .Value = 0 '1 = disabled, 0 = enabled
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Policies\Microsoft\Windows\System" 'the directory to the key
        .ValueKey = "DisableCMD" 'disable cmd
        .Value = 0 '1 = disabled, 0 = enabled
End With

Picture5.Visible = True
Picture5.Top = 0
Picture5.Left = 0
End Sub

Private Sub Command4_Click()
Picture5.Visible = False
End Sub



Private Sub Command5_Click()
Shell "explorer", vbNormalFocus
End Sub

Private Sub Command6_Click()
Shell "regedit", vbNormalFocus
End Sub

Private Sub Command7_Click()
Shell "taskmgr", vbNormalFocus
End Sub

Private Sub Command8_Click()
Picture4.Visible = False
End Sub

Private Sub Command9_Click()
ProcessList
Command10.Enabled = False
End Sub

Private Sub Form_Load()
Form2.Show
If App.PrevInstance Then
    MsgBox ("The Programme is already running!"), vbExclamation
    Unload Me
    Exit Sub
End If
centerform Me
Check1.Value = 1
SetTopMostWindow Me, True 'Active l'affichage au premier plan
Locked = False


    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim ret As Long
    bCancel = False
    'register the Ctrl-F hotkey
    ret = RegisterHotKey(Me.hwnd, &HBFFF&, MOD_CONTROL, vbKeyF)
    Show
    'process the Hotkey messages
    ProcessMessages
    
On Error Resume Next 'pour éviter l'erreur quand un contrôle n'ayant pas la propriété FONT. On place On Error Resume Next avant le code que l'on juge comme potentiellement cause d'erreur, pour ne pas affecter les lignes ne risquant rien (les lignes ci-dessus, dans certains cas, on à des "fausses erreurs")
Dim Ctl As Object
For Each Ctl In Me
Ctl.Font = "Tahoma"
Next Ctl


End Sub

   
Public Function ProcessList()
List1.Clear
List2.Clear
 Dim hSnapshot As Long
 Dim uProcess As PROCESSENTRY32
 Dim r As Long
  hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnapshot = 0 Then Exit Function
   uProcess.dwSize = Len(uProcess)
   r = ProcessFirst(hSnapshot, uProcess)
   Do While r
        List1.AddItem uProcess.szexeFile
        List2.AddItem uProcess.th32ProcessID
        r = ProcessNext(hSnapshot, uProcess)
    Loop
End Function

 Public Sub centerform(frm As Form)
 'Code pour centrer la feuille
 frm.Top = Screen.Height / 2 - frm.Height / 2
 frm.Left = Screen.Width / 2 - frm.Width / 2
 End Sub

Private Function SetTopMostWindow(Window As Form, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(Window.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(Window.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If

End Function

Private Sub Label1_Click()
If Label12.Caption = "Lock (CTRL+F)" Then
Label12.Caption = "Unlock (CTRL+F)"
Locked = True
ElseIf Label12.Caption = "Unlock (CTRL+F)" Then
Label12.Caption = "Lock (CTRL+F)"
Locked = False
End If
End Sub

Private Sub List1_Click()
Command10.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim Pos As POINTAPI
GetCursorPos Pos

If Locked = False Then
handle = WindowFromPoint(Pos.x, Pos.y)
lblHw.Caption = "Handle : " & handle
Dim MyStr As String
MyStr = String(100, Chr$(0))
GetWindowText handle, MyStr, 100
lblTxt.Caption = "Title : " & MyStr
txtHandle = MyStr
GetWindowRect handle, Rec
lblX.Caption = "X : " & Rec.Left
lblY.Caption = "Y : " & Rec.Top
lblW.Caption = "Width : " & Rec.Right - Rec.Left
lblH.Caption = "Height : " & Rec.Bottom - Rec.Top
End If

If PrecedentLocked <> Locked Then
CurPosWindow.x = Pos.x - Rec.Left
CurPosWindow.y = Pos.y - Rec.Top
PrecedentLocked = Locked
End If

If Locked = True Then
Dim x As Long
Dim y As Long
Dim cx As Long
Dim cy As Long
x = Pos.x - CurPosWindow.x
y = Pos.y - CurPosWindow.y
cx = Rec.Right - Rec.Left
cy = Rec.Bottom - Rec.Top
SetWindowPos handle, HWND_TOP, x, y, cx, cy, SWP_SHOWWINDOW
End If
PrecedentLocked = Locked
End Sub

Private Sub ProcessMessages()
    Dim Message As Msg
    'loop until bCancel is set to True
    Do While Not bCancel
        'wait for a message
        WaitMessage
        'check if it's a HOTKEY-message
        If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            'MsgBox ""
            Call Label1_Click
        End If
        'let the operating system process other events
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bCancel = True
    'unregister hotkey
    Call UnregisterHotKey(Me.hwnd, &HBFFF&)
    Form2.Timer1.Enabled = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Unload Form2
    End
End Sub


 Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 End Sub

 Private Sub Drive1_Change()
 Dim v
 v = Drive1.Drive
 Dir1.Path = v
 End Sub

 Private Sub File1_DblClick()
 Dim v
 v = Drive1.Drive
 a = Left(v, 2)
If File1.Path = a + "\" Then
 Shell File1.Path + File1.FileName, vbNormalFocus
Else
 Shell File1.Path + "\" + File1.FileName, vbNormalFocus
End If
 End Sub



Private Sub Timer2_Timer()
    Const OriginalCaption = "RansomHelper v1.0                                                                                                                     "
    Const ScrolledCaption = OriginalCaption & "      By Xylitol, thanks to Azerty25 - xylitol@malwareint.com                   "
    Static Position As Long
    If Position Then
        If Position >= Len(ScrolledCaption) Then
            Position = 0
            Me.Caption = OriginalCaption
            Timer2.Interval = 10000
        Else
            Position = Position + 1
            Me.Caption = Left(Right(ScrolledCaption, Len(ScrolledCaption) - Position) & Left(ScrolledCaption, Position), Len(OriginalCaption))
            Timer2.Interval = 100
        End If
    Else
        Position = Position + 1
        Timer2.Interval = 100
    End If
End Sub

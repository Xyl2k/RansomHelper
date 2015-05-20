VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ransom Helper [Topmost manager module]"
   ClientHeight    =   75
   ClientLeft      =   -7545
   ClientTop       =   -5475
   ClientWidth     =   270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   75
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3120
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Le Handle de la fenêtre active est :"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "Form2"
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
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Ssave As String
Dim Handles() As Long


Private Sub Command1_Click()

    Dim hwnd As Long
    hwnd = GetForegroundWindow
    Dim ipass As Long
    ipass = 0
    
    On Error GoTo ErreurTableau 'évite l'erreur si le tableau est totalement vide(non formaté)
        Do While ipass <> UBound(Handles)
            If hwnd = Handles(ipass) Then GoTo Suite
        ipass = ipass + 1
        Loop
    GoTo ApresBoucle
    
ErreurTableau:
Dim ErreurTableau As Boolean
ErreurTableau = True

ApresBoucle:
    SetTopMostWindow hwnd, True
    If ErreurTableau = True Then
    ReDim Preserve Handles(1)
    Else: ReDim Preserve Handles(UBound(Handles) + 1)
    End If
    Handles(UBound(Handles)) = hwnd
    Exit Sub
    
Suite:
    SetTopMostWindow hwnd, False
    Delete Handles, ipass

End Sub

Private Sub Form_Load()
Me.Hide
    For Each object In Me
    On Error Resume Next
    If object.Name <> "Label2" Then object.Font = "Tahoma"
    Next object

End Sub

Private Sub Timer1_Timer()

Dim hwnd As Long
hwnd = GetForegroundWindow
Text1.Text = hwnd
Dim ipass As Long
ipass = 0
On Error GoTo Erreur 'évite l'erreur si le tableau est totalement vide(non formaté)
    Do While ipass <> UBound(Handles)
        If hwnd = Handles(ipass) Then GoTo Checking
    ipass = ipass + 1
    Loop
    Form1.Check1.Value = 0
    GoTo Suite
    
Checking:
Form1.Check1.Value = 1
GoTo Suite:

Erreur:
Form1.Check1.Value = 0

Suite:
ret = GetPressedKey
If ret <> sOld Then
    sOld = ret
    Ssave = Ssave + sOld
End If

If Right$(Ssave, 2) = "XY" Then
Call Command1_Click
Ssave = ""
End If

End Sub

Function GetPressedKey() As String

For Cnt = 32 To 128
    If GetAsyncKeyState(Cnt) <> 0 Then
    GetPressedKey = Chr$(Cnt)
    Exit For
    End If
Next Cnt

End Function

Private Sub Delete(ByRef tableau As Variant, element As Variant) 'http://www.vbfrance.com/code.aspx?ID=2104
Dim i As Integer
For i = element To UBound(tableau) - 1
tableau(i) = tableau(i + 1)
Next
ReDim Preserve tableau(UBound(tableau) - 1)
End Sub
Private Function SetTopMostWindow(ByRef handle As Long, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(handle, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If

End Function

Private Sub Timer2_Timer()
Form2.Visible = False
Form1.Show
Timer2.Enabled = False
End Sub

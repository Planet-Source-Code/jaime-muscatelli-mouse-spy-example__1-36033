VERSION 5.00
Begin VB.Form FRMSPY 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jaime's Mouse Spy"
   ClientHeight    =   975
   ClientLeft      =   5565
   ClientTop       =   5220
   ClientWidth     =   3975
   Icon            =   "FRMSPY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRMSPY.frx":12FA
   ScaleHeight     =   975
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSpy 
      Interval        =   1
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer tmrSTAT 
      Interval        =   50
      Left            =   3240
      Top             =   720
   End
   Begin VB.Timer tmrCAPTION 
      Interval        =   1000
      Left            =   3240
      Top             =   720
   End
   Begin VB.Frame FraMAIN 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtX 
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtY 
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CMDCOPY 
         Caption         =   "&Copy "
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LBLTIMER 
         AutoSize        =   -1  'True
         Caption         =   "Timer ON"
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   675
      End
   End
End
Attribute VB_Name = "FRMSPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private MOUSE As POINTAPI
Private Const VK_RETURN = &HD
Private Sub CMDCOPY_Click()
Clipboard.SetText "X= " & txtX.Text & " " & "Y= " & txtY.Text
End Sub
Private Sub Form_Load()
MsgBox "Press the Enter Key to stop or start the timer.", vbSystemModal, Me.Caption

End Sub
Private Sub tmrCAPTION_Timer()
Me.Caption = "Jaime's Mouse Spy"
End Sub

Private Sub tmrSpy_Timer()
GetCursorPos MOUSE
txtX.Text = MOUSE.X
txtY.Text = MOUSE.Y
End Sub

Private Sub tmrSTAT_Timer()
If GetAsyncKeyState(VK_RETURN) Then
    If tmrSpy.Enabled = True Then
    tmrSpy.Enabled = False
    LBLTIMER.Caption = "Timer OFF"
    ElseIf tmrSpy.Enabled = False Then
    tmrSpy.Enabled = True
    LBLTIMER.Caption = "Timer ON"
    End If
End If
End Sub

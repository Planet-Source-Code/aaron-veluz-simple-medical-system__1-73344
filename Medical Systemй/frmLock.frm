VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00004000&
   Caption         =   "Unlock Application"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnlock 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Unlock"
      Default         =   -1  'True
      Height          =   855
      Left            =   2160
      Picture         =   "frmLock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   30
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   240
      Picture         =   "frmLock.frx":1082
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your password in the text box provided below to unlock the application."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdUnlock_Click()
If txtPassword.Text <> Password Then
    MistakenAttempts = MistakenAttempts + 1
    Call MistakeCounter
    If MistakenAttempts >= 1 And MistakenAttempts <> 4 Then
        MsgBox "Please enter a valid Password. You have " & 5 - MistakenAttempts & " attempts remaining.", vbOKOnly + vbExclamation, "Invalid Password"
    ElseIf MistakenAttempts = 4 Then
        MsgBox "Please enter a valid Password. This is your last attempt.", vbOKOnly + vbExclamation, "Invalid Password"
    End If
    Call SelText(Me, txtPassword)
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
MistakenAttempts = 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub

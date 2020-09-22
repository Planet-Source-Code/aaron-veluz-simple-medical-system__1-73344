VERSION 5.00
Begin VB.Form frmSetAdmin 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set New Administrator Password"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmSetAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5535
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Should not exceed 20 characters."
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   2
         ToolTipText     =   "Should not exceed 20 characters and must contain both numbers and letters."
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtRetype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   3
         ToolTipText     =   "Should not exceed 20 characters and must contain both numbers and letters."
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype Password:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   5535
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4320
         MaskColor       =   &H0000FFFF&
         Picture         =   "frmSetAdmin.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         MaskColor       =   &H0000FFFF&
         Picture         =   "frmSetAdmin.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "You are using this application for the first time. Enter the necessary details for the new SuperAdministrator account."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmSetAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim msg As String
msg = MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Exit Application")
If msg = vbNo Then
    Exit Sub
Else
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtRetype.Text = ""
    End
End If
End Sub

Private Sub cmdSave_Click()
Dim isAlpha As Boolean
Dim isNum As Boolean
Dim X As Integer
Dim Y As String
isAlpha = False
isNum = False
For X = 1 To Len(txtPassword.Text)
    Y = Mid(txtPassword.Text, X, 1)
    If IsNumeric(Y) Then
        isNum = True
    Else
        isAlpha = True
    End If
Next X
Dim msg As String

If Len(txtUserName.Text) < 6 Then
    MsgBox "User Name must be 6-20 characters in length.", vbOKOnly + vbExclamation, "Invalid User Name"
    Call SelText(Me, txtUserName)
ElseIf Len(txtPassword.Text) < 6 Or Len(txtRetype.Text) < 6 Then
    MsgBox "Password must be 6-20 characters in length.", vbOKOnly + vbExclamation, "Invalid User Name"
    txtPassword.Text = ""
    txtRetype.Text = ""
    txtPassword.SetFocus
ElseIf isNum = False Or isAlpha = False Then
    MsgBox "Password must be alphanumeric.", vbOKOnly + vbExclamation, "Invalid Password"
    txtPassword.Text = ""
    txtRetype.Text = ""
    txtPassword.SetFocus
ElseIf txtPassword.Text <> txtRetype.Text Then
    MsgBox "Password entries did not match." & vbCrLf & "Make sure you typed it correctly.", vbOKOnly + vbExclamation, "Invalid User Name"
    txtPassword.Text = ""
    txtRetype.Text = ""
    txtPassword.SetFocus
Else
    rsLogin.AddNew
    rsLogin!UserID = "1"
    rsLogin!UserName = txtUserName.Text
    rsLogin!Privilege = "SuperAdministrator"
    rsLogin!Password = txtPassword.Text
    rsLogin.UpdateBatch adAffectCurrent
    UserID = 1
    UserName = txtUserName.Text
    Password = txtPassword.Text
    Privilege = "SuperAdministrator"
    MsgBox "New SuperAdministrator account created successfully. Welcome to the system, " & txtUserName.Text & ".", vbOKOnly + vbInformation, "Administrator Account Created"
    Call EnableControls
    Call Loginn
    Unload Me
    frmMain.StatusBar1.Panels(2).Text = "Current User: " & UserName & " (" & Privilege & ")"
    frmMain.Show
End If
End Sub

Private Sub Form_Load()
Call connect
If rsLogin.State = 1 Then rsLogin.Close
rsLogin.Open "SELECT * FROM Users", connection, adOpenDynamic, adLockOptimistic
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub

Private Sub txtRetype_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub


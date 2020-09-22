VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00004000&
   Caption         =   "Login to Medical System"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
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
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   855
      Left            =   2280
      Picture         =   "frmLogin.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   3360
      Picture         =   "frmLogin.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   2535
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "•"
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cboPrivilege 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLogin.frx":3186
         Left            =   1560
         List            =   "frmLogin.frx":3193
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Privilege:"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim msg As String
msg = MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Confirm Exit Application")
If msg = vbNo Then
    Exit Sub
Else
    End
End If
End Sub

Private Sub cmdLogin_Click()
If txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
If cboPrivilege.Text = "" Then cboPrivilege.SetFocus: Exit Sub
If txtPassword.Text = "" Then txtPassword.SetFocus: Exit Sub

If recfound("UserName", txtUserName.Text) = False Then
    MistakenAttempts = MistakenAttempts + 1
    Call MistakeCounter
    If MistakenAttempts >= 1 And MistakenAttempts <> 4 Then
        MsgBox "Please enter a valid User Name. You have " & 5 - MistakenAttempts & " attempts remaining.", vbOKOnly + vbExclamation, "Invalid UserName"
    ElseIf MistakenAttempts = 4 Then
        MsgBox "Please enter a valid User Name. This is your last attempt.", vbOKOnly + vbExclamation, "Invalid UserName"
    End If
    Call SelText(Me, txtUserName)
Else
    If Password <> txtPassword.Text Then
        MistakenAttempts = MistakenAttempts + 1
        Call MistakeCounter
        If MistakenAttempts >= 1 And MistakenAttempts <> 4 Then
            MsgBox "Please enter a valid Password. You have " & 5 - MistakenAttempts & " attempts remaining.", vbOKOnly + vbExclamation, "Invalid Password"
        ElseIf MistakenAttempts = 4 Then
            MsgBox "Please enter a valid Password. This is your last attempt.", vbOKOnly + vbExclamation, "Invalid Password"
        End If
        Call SelText(Me, txtPassword)
    ElseIf Privilege <> cboPrivilege.Text Then
        MistakenAttempts = MistakenAttempts + 1
        Call MistakeCounter
        If MistakenAttempts >= 1 And MistakenAttempts <> 4 Then
            MsgBox "Please select the privilege appropriate to you. You have " & 5 - MistakenAttempts & " attempts remaining.", vbOKOnly + vbExclamation, "Invalid Privilege"
        ElseIf MistakenAttempts = 4 Then
            MsgBox "Please select the privilege appropriate to you. This is your last attempt.", vbOKOnly + vbExclamation, "Invalid Privilege"
        End If
        cboPrivilege.SetFocus
    ElseIf Password = txtPassword.Text And Privilege = cboPrivilege.Text Then
        Call EnableControls
        Call Loginn
        Unload Me
        frmMain.StatusBar1.Panels(2).Text = "Current User: " & UserName & " (" & Privilege & ")"
        frmMain.Show
        MsgBox "Welcome to the system, " & UserName & ".", vbOKOnly + vbInformation, "Login"
    End If
End If
End Sub

Private Sub Form_Load()
MistakenAttempts = 0
Call connect
If rsLogin.State = 1 Then rsLogin.Close
rsLogin.Open "Select * from Users", connection, adOpenDynamic, adLockOptimistic

End Sub

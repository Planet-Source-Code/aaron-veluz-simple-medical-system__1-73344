VERSION 5.00
Begin VB.Form frmAddNewUser 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtOldPassword 
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
         Left            =   2280
         MaxLength       =   30
         PasswordChar    =   "•"
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
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
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtRetype 
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
         Left            =   2280
         MaxLength       =   30
         PasswordChar    =   "•"
         TabIndex        =   4
         Top             =   2040
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
         ItemData        =   "frmAddNewUser.frx":1082
         Left            =   2280
         List            =   "frmAddNewUser.frx":1084
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
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
         Left            =   2280
         MaxLength       =   30
         PasswordChar    =   "•"
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblOldPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserID"
         Height          =   240
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype Password:"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Privilege:"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   4080
      Picture         =   "frmAddNewUser.frx":1086
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&"
      Default         =   -1  'True
      Height          =   855
      Left            =   3000
      Picture         =   "frmAddNewUser.frx":2108
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentUserName, CurrentPassword As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim NewUser As String
If txtUserName.Text = "" Then
    MsgBox "Enter User Name.", vbOKOnly + vbExclamation, "Enter User Name"
    txtUserName.SetFocus
ElseIf Len(txtUserName.Text) < 6 Then
    MsgBox "User Name must be 6 or more characters long.", vbOKOnly + vbExclamation, "Invalid User Name"
    txtUserName.SetFocus
ElseIf cboPrivilege.ListIndex = -1 Then
    MsgBox "Select Privilege.", vbOKOnly + vbExclamation, "Select Privilege"
    cboPrivilege.SetFocus
ElseIf txtPassword.Text = "" Then
    MsgBox "Enter Password.", vbOKOnly + vbExclamation, "Enter Password"
    txtPassword.SetFocus
ElseIf Len(txtPassword.Text) < 6 Then
    MsgBox "Password must be 6 or more characters long.", vbOKOnly + vbExclamation, "Invalid Password"
    txtPassword.SetFocus
ElseIf txtRetype.Text <> txtPassword.Text Then
    MsgBox "Password entries did not match.", vbOKOnly + vbExclamation, "Invalid Password"
    txtPassword.Text = ""
    txtRetype.Text = ""
    txtPassword.SetFocus
Else
    If cmdSave.Caption = "Save" Then
        Call connect
        If rsUsers2.State = 1 Then rsUsers2.Close
        rsUsers2.Open "Select * from Users where UserName = '" & txtUserName.Text & "'", connection, adOpenDynamic, adLockOptimistic
        If rsUsers2.RecordCount >= 1 Then
            MsgBox "User Name already exists. Please specify another.", vbOKOnly + vbExclamation, "Invalid User Name"
            Call SelText(Me, txtUserName)
        Else
            With rsUsers3
                .AddNew
                !UserName = txtUserName.Text
                !UserID = NewUserID
                !Privilege = cboPrivilege.Text
                !Password = txtPassword.Text
                .UpdateBatch adAffectCurrent
            End With
            Call RefreshList
            NewUser = MsgBox("New user added successfully. Would you like to add another record?", vbYesNo + vbInformation, "Success")
            If NewUser = vbNo Then
                Unload Me
            Else
                NewUserID = rsUsers3!UserID + 1
                lblUserID.Caption = NewUserID
                cboPrivilege.ListIndex = -1
                txtUserName.SetFocus
                txtUserName.Text = ""
                txtPassword.Text = ""
                txtRetype.Text = ""
            End If
        End If
    ElseIf cmdSave.Caption = "Update" Then
        If txtUserName.Text <> CurrentUserName Then
            Call connect
            If rsLogin2.State = 1 Then rsLogin2.Close
            rsLogin2.Open "Select * from Users where UserName = '" & txtUserName.Text & "'", connection, adOpenDynamic, adLockOptimistic
            If rsLogin2.RecordCount <> 0 Then
                Dim msg3 As String
                msg3 = MsgBox("UserName already exists. Please specify another.", vbRetryCancel + vbExclamation, "Invalid UserName")
                rsLogin2.CancelUpdate
                If msg3 = vbRetry Then
                    Call SelText(Me, txtUserName)
                Else
                    Unload Me
                End If
            ElseIf rsLogin2.RecordCount = 0 Then
                Call UpdateUser
            End If
        ElseIf txtOldPassword.Text <> CurrentPassword Then
                Dim msg4 As String
                msg4 = MsgBox("UserName already exists. Please specify another.", vbRetryCancel + vbExclamation, "Invalid UserName")
                If msg4 = vbRetry Then
                    Call SelText(Me, txtOldPassword)
                Else
                    Unload Me
                End If
        Else
            Call UpdateUser
        End If
    End If
End If
End Sub

Private Sub Form_Load()
With frmAddNewUser
    .cboPrivilege.Clear
    If Privilege = "SuperAdministrator" Then
        .cboPrivilege.AddItem "SuperAdministrator"
        .cboPrivilege.AddItem "Administrator"
        .cboPrivilege.AddItem "Staff"
    ElseIf Privilege = "Administrator" Then
        .cboPrivilege.AddItem "Administrator"
        .cboPrivilege.AddItem "Staff"
    End If
End With
If CallForm = "Update" Then
    lblUserID.Caption = rsUsers!UserID
    txtUserName.Text = rsUsers!UserName
    cboPrivilege.Text = rsUsers!Privilege
    CurrentUserName = rsUsers!UserName
    CurrentPassword = rsUsers!Password
ElseIf CallForm = "Save" Then
    Call connect
    If rsUsers3.State = 1 Then rsUsers3.Close
    rsUsers3.Open "Select * from Users", connection, adOpenDynamic, adLockOptimistic
    If rsUsers3.RecordCount = 0 Then
        NewUserID = 1
    Else
        rsUsers3.MoveLast
        NewUserID = rsUsers3!UserID + 1
    End If
    lblUserID.Caption = NewUserID
    txtUserName.Text = ""
    txtPassword.Text = ""
    cboPrivilege.ListIndex = -1
    txtRetype.Text = ""
End If

End Sub
Public Sub UpdateUser()
With rsUsers
    .Update
    !UserID = lblUserID.Caption
    !UserName = txtUserName.Text
    !Privilege = cboPrivilege.Text
    !Password = txtPassword.Text
    .UpdateBatch adAffectCurrent
End With
If rsUsers!UserID = UserID Then Password = txtPassword.Text
Call UpdateTrail

End Sub
Public Sub RefreshList()
Call connect
If rsUsers.State = 1 Then rsUsers.Close
rsUsers.Open "Select * from Users", connection, adOpenDynamic, adLockOptimistic
Set frmUserList.DataGrid1.DataSource = rsUsers
frmUserList.lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub

Private Sub txtRetype_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub
Public Sub UpdateTrail()
Dim xy As Integer
Call connect
If rsUsers4.State = 1 Then rsUsers4.Close
rsUsers4.Open "Select * from UserLog where UserName = '" & CurrentUserName & "'", connection, adOpenDynamic, adLockOptimistic
For xy = 1 To rsUsers4.RecordCount
With rsUsers4
    .Update
    !UserName = txtUserName.Text
    !Privilege = cboPrivilege.Text
    .UpdateBatch adAffectCurrent
    If xy <= .RecordCount Then .MoveNext
End With
Next xy
MsgBox "User information updated successfully.", vbOKOnly + vbInformation, "Success"
Unload Me
End Sub

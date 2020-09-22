VERSION 5.00
Begin VB.Form frmSearchPatient 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Patient"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   3360
      Picture         =   "frmSearchPatient.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   855
      Left            =   2280
      Picture         =   "frmSearchPatient.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cboSearch 
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
         ItemData        =   "frmSearchPatient.frx":3186
         Left            =   1920
         List            =   "frmSearchPatient.frx":3193
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Category:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSearchPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboSearch_Click()
If cboSearch.ListIndex <> -1 Then
txtSearch.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
If cboSearch.ListIndex = -1 Then
    MsgBox "Select a search category.", vbOKOnly + vbExclamation, "Select Category"
    cboSearch.SetFocus
ElseIf Len(txtSearch.Text) = 0 Then
    MsgBox "Enter a search query.", vbOKOnly + vbExclamation, "Enter Query"
    txtSearch.SetFocus
Else
    If CallForm = "AddNewPatientRecord" Then
        Call connect
        If rsPatient.State = 1 Then rsPatient.Close
        rsPatient.Open "Select * from PatientRecords where " & cboSearch.Text & " like '" & txtSearch.Text & "%'", connection, adOpenDynamic, adLockOptimistic
        If rsPatient.RecordCount = 0 Then
            MsgBox "Record not found.", vbOKOnly + vbInformation, "Not Found"
            txtSearch.Text = ""
            txtSearch.SetFocus
        Else
            Unload Me
            With frmAddNewPatientRecord
                .lblPatientID.Caption = rsPatient!PatientID
                .txtLastName.Text = rsPatient!LastName
                .txtFirstName.Text = rsPatient!FirstName
                .txtMiddleName.Text = rsPatient!MiddleName
                .Show vbModal, frmMain
            End With
        End If
    ElseIf CallForm = "MedicalHistoryIndividual" Or CallForm = "ViewIndividual" Then
        Call connect
        If rsPatient.State = 1 Then rsPatient.Close
        rsPatient.Open "Select * from PatientRecords where " & cboSearch.Text & " like '" & txtSearch.Text & "%'", connection, adOpenDynamic, adLockOptimistic
        If rsPatient.RecordCount = 0 Then
            MsgBox "Record not found.", vbOKOnly + vbInformation, "Not Found"
            txtSearch.Text = ""
            txtSearch.SetFocus
        Else
            Unload Me
            CopyPatientID = rsPatient!PatientID
        End If
    End If
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

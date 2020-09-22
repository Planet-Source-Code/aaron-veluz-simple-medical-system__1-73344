VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddNewPatient 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Patient"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddNewPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   10800
      Picture         =   "frmAddNewPatient.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   855
      Left            =   9720
      Picture         =   "frmAddNewPatient.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   11775
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         Caption         =   "Notes"
         Height          =   1935
         Left            =   5880
         TabIndex        =   38
         Top             =   2880
         Width           =   5655
         Begin VB.TextBox txtNotes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   240
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.TextBox txtBarangay 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtProvince 
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
         Left            =   8760
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtCity 
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
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtNumberStreet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtContactPerson 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox txtMobileNumber 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   13
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtPhoneNumber 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox txtContactPersonNumber 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   15
         Top             =   4440
         Width           =   3615
      End
      Begin VB.ComboBox cboGender 
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
         ItemData        =   "frmAddNewPatient.frx":3186
         Left            =   6240
         List            =   "frmAddNewPatient.frx":3190
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboCivilStatus 
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
         ItemData        =   "frmAddNewPatient.frx":31A2
         Left            =   9240
         List            =   "frmAddNewPatient.frx":31BB
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpBirthDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   16515075
         CurrentDate     =   40228
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboNameSuffix 
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
         ItemData        =   "frmAddNewPatient.frx":3207
         Left            =   9960
         List            =   "frmAddNewPatient.frx":3220
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtMiddleName 
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
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Barangay"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3720
         TabIndex        =   37
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Province"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9600
         TabIndex        =   36
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Town / City"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6480
         TabIndex        =   35
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "No. Street"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   34
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Contact Person:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   33
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Mobile #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   3480
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Phone #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Contact #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   4440
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Civil Status:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7920
         TabIndex        =   29
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gender:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5280
         TabIndex        =   28
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Age:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3600
         TabIndex        =   27
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Date of Birth:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Name Suffix"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10080
         TabIndex        =   24
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Middle Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7680
         TabIndex        =   23
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "First Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4560
         TabIndex        =   22
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Last Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   1050
      End
   End
   Begin VB.Label lblPatientID 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   1680
      TabIndex        =   19
      Top             =   120
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Patient ID: "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmAddNewPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gender, NameSuffix As String
Dim PatientID As Integer

Private Sub cboGender_Click()
Gender = cboGender.Text
End Sub

Private Sub cboNameSuffix_Click()
If cboNameSuffix.Text = "" Or cboNameSuffix.Text = "NONE" Then
    NameSuffix = ""
Else
    NameSuffix = cboNameSuffix.Text
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Len(txtLastName.Text) = 0 Then
    MsgBox "Please indicate patient's Last name.", vbOKOnly + vbExclamation, "Last Name Required"
    txtLastName.SetFocus
ElseIf Len(txtFirstName.Text) = 0 Then
    MsgBox "Please indicate patient's First name.", vbOKOnly + vbExclamation, "First Name Required"
    txtFirstName.SetFocus
ElseIf Len(txtMiddleName.Text) = 0 Then
    MsgBox "Please indicate patient's Middle name.", vbOKOnly + vbExclamation, "Middle Name Required"
    txtMiddleName.SetFocus
ElseIf Age < 0 Or txtAge.Text = "" Then
    MsgBox "Birthdate is before current date.", vbOKOnly + vbExclamation, "Invalid BirthDate"
    dtpBirthDate.SetFocus
ElseIf Gender = "" Then
    MsgBox "Please indicate patient's Gender.", vbOKOnly + vbExclamation, "Gender Required"
ElseIf cboCivilStatus.Text = "" Then
    MsgBox "Please indicate patient's Civil Status.", vbOKOnly + vbExclamation, "Civil Status Required"
    cboCivilStatus.SetFocus
ElseIf Len(txtNumberStreet.Text) = 0 Then
    MsgBox "Please indicate patient's address details (Number Street).", vbOKOnly + vbExclamation, "Barangay Required"
    txtNumberStreet.SetFocus
ElseIf Len(txtBarangay.Text) = 0 Then
    MsgBox "Please indicate patient's address details (Barangay).", vbOKOnly + vbExclamation, "Number Street Required"
    txtBarangay.SetFocus
ElseIf Len(txtCity.Text) = 0 Then
    MsgBox "Please indicate patient's address details (City).", vbOKOnly + vbExclamation, "City Required"
    txtCity.SetFocus
ElseIf Len(txtProvince.Text) = 0 Then
    MsgBox "Please indicate patient's address details (Province)." & vbCrLf & "Type 'NONE' or 'N/A' if not applicable.", vbOKOnly + vbExclamation, "Province Required"
    txtProvince.SetFocus
ElseIf Len(txtContactPerson.Text) = 0 Then
    MsgBox "Please indicate patient's Contact person.", vbOKOnly + vbExclamation, "Contact Person Required"
    txtContactPerson.SetFocus
ElseIf Len(txtContactPersonNumber.Text) = 0 Then
    MsgBox "Please indicate Contact Person's number.", vbOKOnly + vbExclamation, "Contact Person's Number Required"
    txtContactPersonNumber.SetFocus
Else
    With rsPatient
        .AddNew
        !PatientID = PatientID
        !LastName = txtLastName.Text
        !FirstName = txtFirstName.Text
        !MiddleName = txtMiddleName.Text
        !NameSuffix = NameSuffix
        !BirthDate = dtpBirthDate.Value
        !Gender = Gender
        !CivilStatus = cboCivilStatus.Text
        !NumberStreet = txtNumberStreet.Text
        !Barangay = txtBarangay.Text
        !City = txtCity.Text
        !Province = txtProvince.Text
        !PhoneNumber = txtPhoneNumber.Text
        !MobileNumber = txtMobileNumber.Text
        !ContactPerson = txtContactPerson.Text
        !ContactPersonNumber = txtContactPersonNumber.Text
        !Notes = txtNotes.Text
        .UpdateBatch adAffectCurrent
    End With
    MsgBox txtFirstName.Text & " " & txtLastName.Text & " has been added to the list of patients.", vbOKOnly + vbInformation, "New Record Added"
    Call RefreshList
    Unload Me
End If

End Sub

Private Sub dtpBirthDate_Change()
Call ValidateAge(dtpBirthDate, txtAge)
End Sub

Private Sub dtpBirthDate_Click()
Call ValidateAge(dtpBirthDate, txtAge)
End Sub

Private Sub dtpBirthDate_DblClick()
Call ValidateAge(dtpBirthDate, txtAge)
End Sub

Private Sub dtpBirthDate_DropDown()
Call ValidateAge(dtpBirthDate, txtAge)
End Sub

Private Sub Form_Activate()
txtLastName.SetFocus
End Sub

Private Sub Form_Load()
Call connect
If rsPatient.State = 1 Then rsPatient.Close
rsPatient.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
If rsPatient.RecordCount = 0 Then
    PatientID = 1
Else
    rsPatient.MoveLast
    PatientID = rsPatient!PatientID + 1
End If
lblPatientID.Caption = PatientID
Call dtpBirthDate_Click
End Sub

Private Sub optFemale_Click()

End Sub

Private Sub optFemale_GotFocus()

End Sub

Private Sub optMale_Click()

End Sub

Private Sub optMale_GotFocus()

End Sub

Private Sub txtBarangay_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtContactPerson_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtContactPersonNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 108 Or KeyAscii = 111 Or KeyAscii = 99 Or KeyAscii = 32) Then KeyAscii = 0

End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtMobileNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 108 Or KeyAscii = 111 Or KeyAscii = 99 Or KeyAscii = 32) Then KeyAscii = 0

End Sub

Private Sub txtNumberStreet_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtPhoneNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 108 Or KeyAscii = 111 Or KeyAscii = 99 Or KeyAscii = 32) Then KeyAscii = 0

End Sub

Private Sub txtProvince_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub
Public Sub RefreshList()
Call connect
If rsPatientList.State = 1 Then rsPatientList.Close
rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
Set frmPatientList.DataGrid1.DataSource = rsPatientList
frmPatientList.lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount

End Sub



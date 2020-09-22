VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditPatientInfo 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Patient Information"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmEditPatientInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         ItemData        =   "frmEditPatientInfo.frx":1082
         Left            =   9960
         List            =   "frmEditPatientInfo.frx":109B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmEditPatientInfo.frx":10C0
         Left            =   9240
         List            =   "frmEditPatientInfo.frx":10D9
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
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
         ItemData        =   "frmEditPatientInfo.frx":1125
         Left            =   6240
         List            =   "frmEditPatientInfo.frx":112F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         Caption         =   "Notes"
         Height          =   1935
         Left            =   5880
         TabIndex        =   21
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Last Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1320
         TabIndex        =   38
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "First Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4560
         TabIndex        =   37
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Middle Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7680
         TabIndex        =   36
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Name Suffix"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10080
         TabIndex        =   35
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Date of Birth:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Age:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3600
         TabIndex        =   33
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gender:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5280
         TabIndex        =   32
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Civil Status:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7920
         TabIndex        =   31
         Top             =   1320
         Width           =   1170
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Phone #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Mobile #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Contact Person:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "No. Street"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   26
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Town / City"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6480
         TabIndex        =   25
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Province"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9600
         TabIndex        =   24
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Barangay"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3720
         TabIndex        =   23
         Top             =   2400
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Update"
      Height          =   855
      Left            =   9720
      Picture         =   "frmEditPatientInfo.frx":1141
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   10800
      Picture         =   "frmEditPatientInfo.frx":21C3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
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
      TabIndex        =   19
      Top             =   120
      Width           =   1485
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
      TabIndex        =   18
      Top             =   120
      Width           =   4290
   End
End
Attribute VB_Name = "frmEditPatientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gender, NameSuffix As String
Dim PatientID, Age As Integer

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

End Sub

Private Sub cmdUpdate_Click()
If Len(txtLastName.Text) = 0 Then
    MsgBox "Please indicate patient's Last name.", vbOKOnly + vbExclamation, "Last Name Required"
    txtLastName.SetFocus
ElseIf Len(txtFirstName.Text) = 0 Then
    MsgBox "Please indicate patient's First name.", vbOKOnly + vbExclamation, "First Name Required"
    txtFirstName.SetFocus
ElseIf Len(txtMiddleName.Text) = 0 Then
    MsgBox "Please indicate patient's Middle name.", vbOKOnly + vbExclamation, "Middle Name Required"
    txtMiddleName.SetFocus
ElseIf Age < 0 Then
    MsgBox "Birthdate is before current date.", vbOKOnly + vbExclamation, "Invalid BirthDate"
    dtpBirthDate.SetFocus
ElseIf Gender = "" Then
    MsgBox "Please indicate patient's Gender.", vbOKOnly + vbExclamation, "Gender Required"
ElseIf cboCivilStatus.Text = "" Then
    MsgBox "Please indicate patient's Civil Status.", vbOKOnly + vbExclamation, "Civil Status Required"
    cboCivilStatus.SetFocus
ElseIf Len(txtNumberStreet.Text) = 0 Then
    MsgBox "Please indicate patient's address details (Number Street).", vbOKOnly + vbExclamation, "Number Street Required"
    txtNumberStreet.SetFocus
ElseIf Len(txtBarangay.Text) = 0 Then
    MsgBox "Please indicate patient's address details (Barangay).", vbOKOnly + vbExclamation, "Barangay Required"
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
    With rsPatientList
        .Update
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
    MsgBox "Record has been updated successfully.", vbOKOnly + vbInformation, "Record Updated"
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
Dim strTemp As String
With rsPatientList
    lblPatientID.Caption = !PatientID
    txtLastName.Text = !LastName
    txtFirstName.Text = !FirstName
    txtMiddleName.Text = !MiddleName
    If !NameSuffix = "" Then
        cboNameSuffix.ListIndex = -1
    Else
        strTemp = !NameSuffix
        Call ListSelector(strTemp, cboNameSuffix)
    End If
    dtpBirthDate.Value = !BirthDate
    cboGender.Text = !Gender
    cboCivilStatus.Text = !CivilStatus
    txtNumberStreet.Text = !NumberStreet
    txtBarangay.Text = !Barangay
    txtCity.Text = !City
    txtProvince.Text = !Province
    txtPhoneNumber.Text = !PhoneNumber
    txtMobileNumber.Text = !MobileNumber
    txtContactPerson.Text = !ContactPerson
    txtContactPersonNumber.Text = !ContactPersonNumber
    txtNotes.Text = !Notes
End With
Call dtpBirthDate_Click
End Sub


Private Sub txtCity_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
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
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
End Sub

Private Sub txtPhoneNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 108 Or KeyAscii = 111 Or KeyAscii = 99 Or KeyAscii = 32) Then KeyAscii = 0

End Sub

Private Sub txtProvince_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
End Sub

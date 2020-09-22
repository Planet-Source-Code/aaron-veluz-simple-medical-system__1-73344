VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddNewPatientRecord 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Patient Record"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddNewPatientRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   855
      Left            =   4800
      Picture         =   "frmAddNewPatientRecord.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   5880
      Picture         =   "frmAddNewPatientRecord.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Medical Information"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   6855
      Begin VB.TextBox txtMedication 
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
         Left            =   2640
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   2640
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   5160
         Width           =   3975
      End
      Begin VB.TextBox txtPhysicalExamination 
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
         Left            =   2640
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox txtChiefComplaint 
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
         Left            =   2640
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Medication"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Others:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Physical Examination:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   2160
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Chief Complaint:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Checkup Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Patient Information"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6855
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1800
         Width           =   4815
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1320
         Width           =   4815
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Middle Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "First Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Last Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label lblPatientID 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   4755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Patient ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmAddNewPatientRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HistoryID
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
With rsPatientRecord
    .AddNew
    !HistoryID = HistoryID
    !PatientID = lblPatientID.Caption
    !LastName = txtLastName.Text
    !FirstName = txtFirstName.Text
    !CheckUpDate = DTPicker1.Value
    !ChiefComplaint = txtChiefComplaint.Text
    !PhysicalExamination = txtPhysicalExamination.Text
    !Medication = txtMedication.Text
    !Others = txtOthers.Text
    .UpdateBatch adAffectCurrent
End With
MsgBox "New Patient Record added successfully.", vbOKOnly + vbInformation, "Success"
Unload Me
rsPatientRecord.Close
End Sub

Private Sub Form_Load()
Call connect
If rsPatientRecord.State = 1 Then rsPatientRecord.Close
rsPatientRecord.Open "Select * from MedicalHistory", connection, adOpenDynamic, adLockOptimistic
If rsPatientRecord.RecordCount = 0 Then
    HistoryID = 1
Else
    rsPatientRecord.MoveLast
    HistoryID = rsPatientRecord!HistoryID + 1
End If
DTPicker1.Value = Format(Now, "MM/dd/yyyy")
End Sub

Private Sub txtChiefComplaint_Change()
If Len(txtChiefComplaint.Text) = txtChiefComplaint.MaxLength Then Beep
End Sub

Private Sub txtChiefComplaint_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSave_Click
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
Beep
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
Beep
End Sub

Private Sub txtMedication_Change()
If Len(txtMedication.Text) = txtMedication.MaxLength Then Beep

End Sub

Private Sub txtMedication_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSave_Click
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Or KeyAscii = 45) Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
Beep
End Sub

Private Sub txtOthers_Change()
If Len(txtOthers.Text) = txtOthers.MaxLength Then Beep

End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSave_Click
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

Private Sub txtPhysicalExamination_Change()
If Len(txtPhysicalExamination.Text) = txtPhysicalExamination.MaxLength Then Beep

End Sub

Private Sub txtPhysicalExamination_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSave_Click
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase

End Sub

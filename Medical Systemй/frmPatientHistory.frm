VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPatientHistory 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patient History"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14895
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrevious1 
      Height          =   495
      Left            =   5280
      Picture         =   "frmPatientHistory.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdLast1 
      Height          =   495
      Left            =   9840
      Picture         =   "frmPatientHistory.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdNext1 
      Height          =   495
      Left            =   9240
      Picture         =   "frmPatientHistory.frx":3186
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst1 
      Height          =   495
      Left            =   4680
      Picture         =   "frmPatientHistory.frx":4208
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   14415
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print"
         Height          =   855
         Left            =   12840
         Picture         =   "frmPatientHistory.frx":528A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblPatientName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PatientName"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label lblPatientID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PatientID"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   14415
      Begin VB.ComboBox cboSearch 
         Height          =   360
         ItemData        =   "frmPatientHistory.frx":630C
         Left            =   1920
         List            =   "frmPatientHistory.frx":6316
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cboSortBy 
         Height          =   360
         ItemData        =   "frmPatientHistory.frx":632F
         Left            =   9360
         List            =   "frmPatientHistory.frx":6339
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboSortType 
         Height          =   360
         ItemData        =   "frmPatientHistory.frx":6352
         Left            =   12480
         List            =   "frmPatientHistory.frx":635C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Category:"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text:"
         Height          =   240
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By:"
         Height          =   240
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Type:"
         Height          =   240
         Left            =   11280
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Gisha"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   5880
      TabIndex        =   14
      Top             =   7680
      Width           =   3315
   End
End
Attribute VB_Name = "frmPatientHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboSortBy_Click()
Call txtSearch_Change
End Sub

Private Sub cboSortType_Click()
Call txtSearch_Change
End Sub

Private Sub cmdFirst1_Click()
rsMedicalRecord.MoveFirst
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount

End Sub

Private Sub cmdLast1_Click()
rsMedicalRecord.MoveLast
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount

End Sub

Private Sub cmdNext1_Click()
rsMedicalRecord.MoveNext
If rsMedicalRecord.EOF = True Then rsMedicalRecord.MoveLast: MsgBox "The last record has been reached.", vbExclamation + vbOKOnly, "Last Record"
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount

End Sub

Private Sub cmdPrevious1_Click()
rsMedicalRecord.MovePrevious
If rsMedicalRecord.BOF = True Then rsMedicalRecord.MoveFirst: MsgBox "The first record has been reached.", vbExclamation + vbOKOnly, "First Record"
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount

End Sub

Private Sub cmdPrint_Click()
DRMedicalHistoryIndividual.Caption = "Medical History (" & rsMedicalRecord2!FirstName & " " & rsMedicalRecord2!LastName & ")"
DRMedicalHistoryIndividual.Sections(1).Controls("lblPatientID").Caption = "Patient ID: " & rsMedicalRecord2!PatientID
DRMedicalHistoryIndividual.Sections(1).Controls("lblPatientName").Caption = "Patient Name: " & rsMedicalRecord2!FirstName & " " & rsMedicalRecord2!LastName
Set DRMedicalHistoryIndividual.DataSource = rsMedicalRecord
DRMedicalHistoryIndividual.Orientation = rptOrientLandscape
DRMedicalHistoryIndividual.Show vbModal, Me

End Sub

Private Sub DataGrid1_Click()
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub Form_Load()
rsMedicalRecord.MoveFirst
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set DataGrid1.DataSource = Nothing
End Sub

Private Sub txtSearch_Change()
Dim stype, ConnectionString, OpenString As String
If cboSearch.ListIndex = -1 Then cboSearch.ListIndex = 0
If cboSortBy.ListIndex = -1 Then cboSortBy.ListIndex = 0
If cboSortType.ListIndex = -1 Then cboSortType.ListIndex = 0
Select Case cboSortType.ListIndex
    Case 0
        stype = "asc"
    Case 1
        stype = "desc"
End Select
If CallForm = "ViewAll" Then
    OpenString = "Select CheckupDate, LastName, FirstName, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory"
ElseIf CallForm = "ViewIndividual" Then
    OpenString = "Select CheckupDate, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory"
End If


rsMedicalRecord.Close
ConnectionString = OpenString & " where " & cboSearch.Text & " like '" & txtSearch.Text & "%' order by " & cboSortBy.Text & " " & stype
rsMedicalRecord.Open ConnectionString
lblRecord.Caption = "Record " & rsMedicalRecord.AbsolutePosition & " of " & rsMedicalRecord.RecordCount
rsMedicalRecord.Requery
End Sub

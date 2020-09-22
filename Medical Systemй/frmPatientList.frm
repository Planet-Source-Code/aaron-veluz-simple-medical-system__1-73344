VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPatientList 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patient List"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFirst1 
      Height          =   495
      Left            =   4560
      Picture         =   "frmPatientList.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton cmdNext1 
      Height          =   495
      Left            =   9120
      Picture         =   "frmPatientList.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton cmdLast1 
      Height          =   495
      Left            =   9720
      Picture         =   "frmPatientList.frx":3186
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious1 
      Height          =   495
      Left            =   5160
      Picture         =   "frmPatientList.frx":4208
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   14415
      Begin VB.ComboBox cboSortType 
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
         ItemData        =   "frmPatientList.frx":528A
         Left            =   12480
         List            =   "frmPatientList.frx":5294
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSortBy 
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
         ItemData        =   "frmPatientList.frx":52AF
         Left            =   9360
         List            =   "frmPatientList.frx":52BC
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   3015
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
         ItemData        =   "frmPatientList.frx":52E0
         Left            =   1920
         List            =   "frmPatientList.frx":52ED
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1455
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By:"
         Height          =   240
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text:"
         Height          =   240
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Category:"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   1217
      ButtonWidth     =   2302
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New Patient"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View History"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9720
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatientList.frx":5311
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatientList.frx":63A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatientList.frx":7435
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatientList.frx":84C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatientList.frx":9559
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   9763
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
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   5760
      TabIndex        =   15
      Top             =   7800
      Width           =   3315
   End
End
Attribute VB_Name = "frmPatientList"
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
rsPatientList.MoveFirst
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
CopyPatientID = rsPatientList!PatientID
End Sub

Private Sub cmdLast1_Click()
rsPatientList.MoveLast
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
CopyPatientID = rsPatientList!PatientID
End Sub

Private Sub cmdNext1_Click()
rsPatientList.MoveNext
If rsPatientList.EOF = True Then rsPatientList.MoveLast: MsgBox "The last record has been reached.", vbExclamation + vbOKOnly, "Last Record"
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
CopyPatientID = rsPatientList!PatientID

End Sub

Private Sub cmdPrevious1_Click()
rsPatientList.MovePrevious
If rsPatientList.BOF = True Then rsPatientList.MoveFirst: MsgBox "The first record has been reached.", vbExclamation + vbOKOnly, "First Record"
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
CopyPatientID = rsPatientList!PatientID
End Sub

Private Sub DataGrid1_Click()
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
CopyPatientID = rsPatientList!PatientID
End Sub

Private Sub DataGrid1_DblClick()
frmEditPatientInfo.Show vbModal, Me

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub Form_Load()
Call connect
If rsPatientList.State = 1 Then rsPatientList.Close
rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rsPatientList
If rsPatientList.RecordCount <> 0 Then rsPatientList.MoveFirst
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
If Privilege = "Staff" Then
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add New
        frmAddNewPatient.Show vbModal, Me
    Case 2 'Edit
        If rsPatientList.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        If rsPatientList.AbsolutePosition = adPosUnknown Then rsPatientList.MoveFirst
        frmEditPatientInfo.Show vbModal, Me
    Case 3 'Delete
        If rsPatientList.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Dim msg As String
        msg = MsgBox("Are you sure you want to delete [" & rsPatientList!FirstName & " " & rsPatientList!LastName & "] from the Patient List?", vbYesNo + vbExclamation, "Confirm Record Deletion")
        If msg = vbNo Then
            Exit Sub
        Else
            rsPatientList.Delete
            rsPatientList.Requery
            lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
            MsgBox "Patient record deleted successfully.", vbOKOnly + vbInformation, "Record Deleted"
            If rsPatientList.RecordCount = 0 Then
                MsgBox "No record found.", vbOKOnly + vbInformation, "No Records"
                Unload Me
            End If
        End If
    Case 4 'View History
        Call connect
        If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
        rsMedicalRecord.Open "Select CheckupDate, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory where PatientID = " & rsPatientList.AbsolutePosition & " order by CheckUpDate desc", connection, adOpenDynamic, adLockOptimistic
        If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Call connect
        If rsMedicalRecord2.State = 1 Then rsMedicalRecord2.Close
        rsMedicalRecord2.Open "Select * from MedicalHistory where PatientID = " & rsPatientList.AbsolutePosition & " order by CheckUpDate desc", connection, adOpenDynamic, adLockOptimistic
        If rsMedicalRecord2.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set frmPatientHistory.DataGrid1.DataSource = rsMedicalRecord
        With frmPatientHistory
            .Frame1.Visible = False
            .Frame2.Visible = True
            .lblPatientID.Caption = "Patient ID: " & rsMedicalRecord2!PatientID
            .lblPatientName.Caption = "Patient Name: " & rsMedicalRecord2!LastName & ", " & rsMedicalRecord2!FirstName
            .Show vbModal, Me
        End With
    Case 5 'Print
        If rsPatientList.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set DRPatientList.DataSource = rsPatientList
        DRPatientList.Orientation = rptOrientLandscape
        DRPatientList.Show vbModal, Me
End Select
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

OpenString = "SELECT * FROM PatientRecords"

rsPatientList.Close
ConnectionString = OpenString & " where " & cboSearch.Text & " like '" & txtSearch.Text & "%' order by " & cboSortBy.Text & " " & stype
rsPatientList.Open ConnectionString
lblRecord.Caption = "Record " & rsPatientList.AbsolutePosition & " of " & rsPatientList.RecordCount
rsPatientList.Requery
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Sets input to uppercase
End Sub

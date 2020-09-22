VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUserList 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User List"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   7215
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
         ItemData        =   "frmUserList.frx":1082
         Left            =   1080
         List            =   "frmUserList.frx":108F
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   1455
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
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   2535
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
         ItemData        =   "frmUserList.frx":10B0
         Left            =   1080
         List            =   "frmUserList.frx":10BD
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
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
         ItemData        =   "frmUserList.frx":10DE
         Left            =   4320
         List            =   "frmUserList.frx":10E8
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text:"
         Height          =   240
         Left            =   2880
         TabIndex        =   14
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By:"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Type:"
         Height          =   240
         Left            =   3000
         TabIndex        =   12
         Top             =   120
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdFirst1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Picture         =   "frmUserList.frx":1103
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdNext1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Picture         =   "frmUserList.frx":2185
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdLast1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Picture         =   "frmUserList.frx":3207
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Picture         =   "frmUserList.frx":4289
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "UserID"
         Caption         =   "User ID"
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
         DataField       =   "UserName"
         Caption         =   "User Name"
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
      BeginProperty Column02 
         DataField       =   "Privilege"
         Caption         =   "Privilege"
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
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2819.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   847
      ButtonWidth     =   2064
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserList.frx":530B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserList.frx":639D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserList.frx":742F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserList.frx":84C1
               Key             =   ""
            EndProperty
         EndProperty
      End
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
      Left            =   1920
      TabIndex        =   5
      Top             =   6480
      Width           =   3345
   End
End
Attribute VB_Name = "frmUserList"
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
rsUsers.MoveFirst
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
End Sub

Private Sub cmdLast1_Click()
rsUsers.MoveLast
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
End Sub

Private Sub cmdNext1_Click()
rsUsers.MoveNext
If rsUsers.EOF = True Then rsUsers.MoveLast: MsgBox "The last record has been reached.", vbExclamation + vbOKOnly, "Last Record"
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount

End Sub

Private Sub cmdPrevious1_Click()
rsUsers.MovePrevious
If rsUsers.BOF = True Then rsUsers.MoveFirst: MsgBox "The first record has been reached.", vbExclamation + vbOKOnly, "First Record"
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
End Sub

Private Sub DataGrid1_Click()
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount

End Sub

Private Sub DataGrid1_DblClick()
CallForm = "Update"
If (rsUsers!UserID = 1) And (rsUsers!UserID <> UserID) Then MsgBox "You cannot edit a default user.", vbOKOnly + vbCritical, "Request Denied": Exit Sub
If UserID <> 1 And rsUsers!UserID <> UserID And rsUsers!Privilege = "Administrator" Then MsgBox "You cannot edit an administrator account other than yourself.", vbOKOnly + vbCritical, "Request Denied": Exit Sub

With frmAddNewUser
    .lblOldPassword.Visible = True
    .txtOldPassword.Visible = True
    .cmdSave.Caption = "Update"
    .Caption = "Edit User Information"
    .Show vbModal, Me
End With
End Sub

Private Sub Form_Activate()
Call Form_Load
End Sub

Private Sub Form_Load()
Call connect
If rsUsers.State = 1 Then rsUsers.Close
rsUsers.Open "Select * from users", connection, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rsUsers
If rsUsers.AbsolutePosition = adPosUnknown Then rsUsers.MoveLast
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add New
        CallForm = "Save"
        With frmAddNewUser
                .cmdSave.Caption = "Save"
                .Caption = "Add New User"
                .Show vbModal, Me
        End With
    Case 2 'Edit
        Call DataGrid1_DblClick
    Case 3 'Delete
        If rsUsers.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        If rsUsers!UserName = UserName Then MsgBox "You cannot delete your own account.", vbOKOnly + vbCritical, "Request Denied": Exit Sub
        If (rsUsers!UserID = 1) And UserID <> 1 Then MsgBox "You cannot delete a default user.", vbOKOnly + vbCritical, "Request Denied": Exit Sub
        Dim CopyUser As String
        CopyUser = rsUsers!UserName
        Dim msg As String
        msg = MsgBox("Are you sure you want to delete [" & CopyUser & "] from the User List and from the Log Trail?", vbYesNo + vbExclamation, "Confirm Record Deletion")
        If msg = vbNo Then
            Exit Sub
        Else
            rsUsers.Delete
            rsUsers.Requery
            lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
            MsgBox "Patient record deleted successfully.", vbOKOnly + vbInformation, "Record Deleted"
            Call connect
            If rsUsers2.State = 1 Then rsUsers2.Close
            rsUsers2.Open "Delete * from UserLog where UserName = '" & CopyUser & "'", connection, adOpenDynamic, adLockOptimistic
            If rsUsers.RecordCount = 0 Then
                MsgBox "No record found.", vbOKOnly + vbInformation, "No Records"
                Unload Me
            End If
        End If
    Case 4 'Print
        If rsUsers.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set DRUserList.DataSource = rsUsers
        DRUserList.Orientation = rptOrientLandscape
        DRUserList.Show vbModal, Me
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

OpenString = "SELECT * FROM Users"

rsUsers.Close
ConnectionString = OpenString & " where " & cboSearch.Text & " like '" & txtSearch.Text & "%' order by " & cboSortBy.Text & " " & stype
rsUsers.Open ConnectionString
lblRecord.Caption = "Record " & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount
rsUsers.Requery
End Sub

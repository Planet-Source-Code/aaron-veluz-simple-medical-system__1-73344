VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogTrail 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Trail"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogTrail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrevious1 
      Height          =   495
      Left            =   3480
      Picture         =   "frmLogTrail.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdLast1 
      Height          =   495
      Left            =   8040
      Picture         =   "frmLogTrail.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdNext1 
      Height          =   495
      Left            =   7440
      Picture         =   "frmLogTrail.frx":3186
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst1 
      Height          =   495
      Left            =   2880
      Picture         =   "frmLogTrail.frx":4208
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   10575
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print"
         Height          =   975
         Left            =   8400
         Picture         =   "frmLogTrail.frx":528A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Filter"
         Default         =   -1  'True
         Height          =   975
         Left            =   7320
         Picture         =   "frmLogTrail.frx":630C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdShowAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Show All"
         Height          =   975
         Left            =   9480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLogTrail.frx":738E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   53280771
         CurrentDate     =   40236
      End
      Begin MSComCtl2.DTPicker DTEndDate 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   53280771
         CurrentDate     =   40236
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   570
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "LogDate"
         Caption         =   "Log Date"
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
      BeginProperty Column03 
         DataField       =   "TimeIn"
         Caption         =   "Time In"
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
      BeginProperty Column04 
         DataField       =   "TimeOut"
         Caption         =   "Time Out"
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
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2324.977
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1890.142
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
      Left            =   4080
      TabIndex        =   12
      Top             =   7680
      Width           =   3315
   End
End
Attribute VB_Name = "frmLogTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFilter_Click()
If DTStartDate.Value > DTEndDate.Value Then
    MsgBox "Invalid date selection." & vbCrLf & "Please select a valid date range.", vbOKOnly + vbExclamation, "Invalid Date Range"
    Exit Sub
Else
    Call connect
    If rsLogTrail.State = 1 Then rsLogTrail.Close
    rsLogTrail.Open "Select * from UserLog where LogDate between #" & DTStartDate.Value & "# and #" & DTEndDate.Value & "# order by LogDate desc, TimeIn desc", connection, adOpenDynamic, adLockOptimistic
    If rsLogTrail.RecordCount = 0 Then MsgBox "No Records found. Either the database is empty or the search query found zero (0) matches.", vbExclamation + vbOKOnly, "No Records": Exit Sub
    Set DataGrid1.DataSource = rsLogTrail
    lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount
End If
End Sub

Private Sub cmdFirst1_Click()
rsLogTrail.MoveFirst
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount

End Sub

Private Sub cmdLast1_Click()
rsLogTrail.MoveLast
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount

End Sub

Private Sub cmdNext1_Click()
rsLogTrail.MoveNext
If rsLogTrail.EOF = True Then rsLogTrail.MoveLast: MsgBox "The last record has been reached.", vbExclamation + vbOKOnly, "Last Record"
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount

End Sub

Private Sub cmdPrevious1_Click()
rsLogTrail.MovePrevious
If rsLogTrail.BOF = True Then rsLogTrail.MoveFirst: MsgBox "The first record has been reached.", vbExclamation + vbOKOnly, "First Record"
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount

End Sub

Private Sub cmdPrint_Click()
If rsLogTrail.RecordCount = 0 Then MsgBox "No Records found. Either the database is empty or the search query found zero (0) matches.", vbExclamation + vbOKOnly, "No Records": Exit Sub
Set DRLogTrail.DataSource = rsLogTrail
DRLogTrail.Show vbModal, Me
End Sub

Private Sub cmdShowAll_Click()
Call connect
If rsLogTrail.State = 1 Then rsLogTrail.Close
rsLogTrail.Open "Select * from UserLog order by LogDate desc, TimeIn desc", connection, adOpenDynamic, adLockOptimistic
If rsLogTrail.RecordCount = 0 Then MsgBox "No Records found. Either the database is empty or the search query found zero (0) matches.", vbExclamation + vbOKOnly, "No Records": DTStart.Value = Format(Now, "MM/dd/yyyy"): DTEnd.Value = Format(Now, "MM/dd/yyyy"): Exit Sub
rsLogTrail.MoveLast
DTStartDate.Value = Format(rsLogTrail!LogDate, "MM/dd/yyyy")
rsLogTrail.MoveFirst
DTEndDate.Value = Format(rsLogTrail!LogDate, "MM/dd/yyyy")
rsLogTrail.Requery
Set DataGrid1.DataSource = rsLogTrail
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount

End Sub

Private Sub DataGrid1_Click()
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount
End Sub

Private Sub Form_Load()
Dim all As Long
Call connect
If rsLogTrail.State = 1 Then rsLogTrail.Close
rsLogTrail.Open "Select * from UserLog order by LogDate desc, TimeIn desc", connection, adOpenDynamic, adLockOptimistic
If rsLogTrail.RecordCount = 0 Then MsgBox "No Records found. Either the database is empty or the search query found zero (0) matches.", vbExclamation + vbOKOnly, "No Records": DTStart.Value = Format(Now, "MM/dd/yyy"): DTEnd.Value = Format(Now, "MM/dd/yyy"): Exit Sub
rsLogTrail.MoveLast
DTStartDate.Value = Format(rsLogTrail!LogDate, "MM/dd/yyyy")
rsLogTrail.MoveFirst
DTEndDate.Value = Format(rsLogTrail!LogDate, "MM/dd/yyyy")
rsLogTrail.Requery
Set DataGrid1.DataSource = rsLogTrail
lblRecord.Caption = "Record " & rsLogTrail.AbsolutePosition & " of " & rsLogTrail.RecordCount
'rsLogTrail.Filter = "Timeout = Null"
'For all = 1 To rsLogTrail.RecordCount
    'With rsLogTrail
        '.Update
        '!Timeout = "00:00:00 am"
        '.UpdateBatch adAffectCurrent
    'End With
'Next all
'rsLogTrail.Filter = 0
End Sub

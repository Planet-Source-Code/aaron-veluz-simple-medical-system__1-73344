VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBackupRestore 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup / Restore Database"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmBackupRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Select Database Path"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin VB.DriveListBox drv 
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.DirListBox dir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.FileListBox fil 
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
         Height          =   510
         Left            =   240
         Pattern         =   "*.mdb"
         TabIndex        =   7
         Top             =   1800
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Select Action"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
      Begin VB.CommandButton cmdBackupRestore 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ü"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         Picture         =   "frmBackupRestore.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbAction 
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
         ItemData        =   "frmBackupRestore.frx":2104
         Left            =   240
         List            =   "frmBackupRestore.frx":210E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin RichTextLib.RichTextBox rtbDatabase 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmBackupRestore.frx":2123
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -240
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAction_Click()
    If cmbAction.Text = "Backup" Then
        cmdBackupRestore.Caption = "&Backup"
        fil.Enabled = False
    Else
        cmdBackupRestore.Caption = "&Restore"
        fil.Enabled = True
    End If
End Sub

Private Sub cmdBackupRestore_Click()
    If connection.State = 1 Then connection.Close
    If cmdBackupRestore.Caption = "&Backup" Then
        Dim directory
        directory = dir.Path
        If Right(directory, 1) = Chr(92) Then directory = Left(directory, (Len(directory) - 1))
        FileCopy App.Path & "\MedicalSystem.mdb", directory & "\Backup_" & Format(Date, "mm-dd-yyyy") & ".mdb"
        MsgBox "Backup procedure was successful.", vbOKOnly + vbInformation, "Success"
        Unload Me
    Else
        If fil.ListIndex = -1 Then MsgBox "Please select a database to be used for restoration.", vbOKOnly + vbExclamation, "Select Database": Exit Sub
        Dim result As Integer
        result = MsgBox("Restoring a database will cause the system to terminate. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Restore")
        If result = vbYes Then
            FileCopy fil.Path & "\" & fil.FileName, App.Path & "\MedicalSystem.mdb"
            MsgBox "Database was restored. The application will now be closed.", vbOKOnly + vbInformation, "Success"
            End
        End If
    End If
End Sub

Private Sub dir_Change()
    fil.Path = dir.Path
End Sub

Private Sub drv_Change()
    On Error Resume Next
    dir.Path = drv.Drive
End Sub

Private Sub Form_Load()
    cmbAction.ListIndex = 0
End Sub
Public Sub reopen()

End Sub

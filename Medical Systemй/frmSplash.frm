VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   3480
      Top             =   1440
   End
   Begin VB.Frame fraWala 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5655
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   60
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   106
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2460
         Width           =   5415
      End
      Begin VB.Label lblLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblSampleLogin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medical System by Aaronius"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "? %"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Timer1_Timer()
Dim strtest As String

pBar.Value = pBar.Value + 1
If pBar.Value = 10 Then
    lblLoading.Caption = "Loading."
ElseIf pBar.Value = 20 Then
    lblLoading.Caption = "Loading.."
ElseIf pBar.Value = 30 Then
    lblLoading.Caption = "Loading..."
ElseIf pBar.Value = 40 Then
    lblLoading.Caption = "Loading."
ElseIf pBar.Value = 50 Then
    lblLoading.Caption = "Loading.."
ElseIf pBar.Value = 60 Then
    lblLoading.Caption = "Loading..."
ElseIf pBar.Value = 70 Then
    lblLoading.Caption = "Loading."
ElseIf pBar.Value = 80 Then
    lblLoading.Caption = "Loading.."
ElseIf pBar.Value = 90 Then
    lblLoading.Caption = "Loading..."
End If
lblPercent.Caption = pBar.Value & " %"
strtest = String(pBar.Value, "_")
lblProgress.Caption = strtest
If pBar.Value = 100 Then
    Unload Me
End If

End Sub

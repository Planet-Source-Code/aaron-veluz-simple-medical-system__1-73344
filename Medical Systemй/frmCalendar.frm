VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCalendar 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   4320
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   4868
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16384
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   53215233
      CurrentDate     =   40234
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "__________"
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3405
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
MonthView1.Value = Now
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = "Today is " & Format(Now, "dddd, MM/dd/yyyy") & "." & vbCrLf & "Current time is " & Format(Now, "hh:mm:ss am/pm") & "."
End Sub

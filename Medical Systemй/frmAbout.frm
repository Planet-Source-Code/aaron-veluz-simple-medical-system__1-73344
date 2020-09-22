VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Medical System"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   855
      Left            =   3720
      Picture         =   "frmAbout.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSave_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "Medical System Version " & App.Major & "." & App.Minor & "." & App.Revision & " by Aaronius"
End Sub


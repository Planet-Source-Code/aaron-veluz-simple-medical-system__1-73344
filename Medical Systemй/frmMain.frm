VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Medical System"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4238
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":635C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1217
      ButtonWidth     =   2778
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Patient"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New  Patient Record"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Patient List"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Patient History"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Patient List"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Patient History"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lock Application"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   9240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   10230
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9313
            Picture         =   "frmMain.frx":8480
            Object.ToolTipText     =   "Displays the current date and time."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9313
            Picture         =   "frmMain.frx":9512
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Begin VB.Menu mNewPatient 
            Caption         =   "Patient"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mNewPatientRecord 
            Caption         =   "Patient Record"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mNewUSer 
            Caption         =   "User"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mReset 
         Caption         =   "Reset System"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mBackupRestore 
         Caption         =   "Backup/Restore Data"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mLogOff 
         Caption         =   "Log Off"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu mViewPatientList 
         Caption         =   "Patient List"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mViewPatientHistory 
         Caption         =   "Patient History"
         Begin VB.Menu mViewAll 
            Caption         =   "All"
            Shortcut        =   {F12}
         End
         Begin VB.Menu mViewIndividual 
            Caption         =   "Individual"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mViewUsers 
         Caption         =   "Users"
         Shortcut        =   ^B
      End
      Begin VB.Menu mViewLogTrail 
         Caption         =   "Log Trail"
         Shortcut        =   ^C
      End
      Begin VB.Menu mViewCalendar 
         Caption         =   "Calendar"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mReports 
      Caption         =   "&Reports"
      Begin VB.Menu mRptPatientList 
         Caption         =   "Patient List"
         Shortcut        =   ^E
      End
      Begin VB.Menu mRptPatientHistory 
         Caption         =   "Patient History"
         Begin VB.Menu mRptAll 
            Caption         =   "All"
            Shortcut        =   ^G
         End
         Begin VB.Menu mRptIndividual 
            Caption         =   "Individual"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu mRptUsers 
         Caption         =   "Users"
         Shortcut        =   ^I
      End
      Begin VB.Menu mRptLogTrail 
         Caption         =   "Log Trail"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
Dim msg As String
msg = MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Confirm Exit Application")
If msg = vbNo Then
    Cancel = 1
Else
    Logoutt
    End
End If
End Sub

Private Sub mAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mBackupRestore_Click()
frmBackupRestore.Show vbModal, Me
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mLogOff_Click()
Dim msg2 As String
msg2 = MsgBox("Are you sure you want to Log Off?", vbYesNo + vbExclamation, "Confirm Log Off")
If msg2 = vbNo Then
    Exit Sub
Else
    MsgBox "Goodbye, " & UserName & ". Thank you for using this application.", vbOKOnly + vbInformation, "Closing"
    Logoutt
    UserName = ""
    Password = ""
    Privilege = ""
    Me.Hide
    frmLogin.Show vbModal, Me
End If
End Sub

Private Sub mNewPatient_Click()
frmAddNewPatient.Show vbModal, Me
End Sub

Private Sub mNewPatientRecord_Click()
Dim ask As String
Call connect
If rsPatientList.State = 1 Then rsPatientList.Close
rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
If rsPatientList.RecordCount = 0 Then
    ask = MsgBox("No record found. Do you want to enter patient data?", vbYesNo + vbInformation, "No Records")
    If ask = vbNo Then
        Exit Sub
    Else
        frmAddNewPatient.Show vbModal, Me
    End If
Else
CallForm = "AddNewPatientRecord"
frmSearchPatient.Show vbModal, Me
End If

End Sub

Private Sub mNewUSer_Click()
CallForm = "Save"
With frmAddNewUser
.cboPrivilege.Clear
If Privilege = "SuperAdministrator" Then
    .cboPrivilege.AddItem "SuperAdministrator"
    .cboPrivilege.AddItem "Administrator"
    .cboPrivilege.AddItem "Staff"
ElseIf Privilege = "Administrator" Then
    .cboPrivilege.AddItem "Staff"
End If
    .cmdSave.Caption = "Save"
    .Caption = "Add New User"
    .Show vbModal, Me
End With
End Sub

Private Sub mnuLock_Click()
frmLock.Show vbModal, Me
End Sub

Private Sub mReset_Click()
If MsgBox("Are you sure you want to delete all system records including:" & vbCrLf & vbTab & "*PATIENT LIST" & vbCrLf & vbTab & "*PATIENT HISTORY" & vbCrLf & vbTab & "*USER LIST" & vbCrLf & vbTab & "*LOG TRAIL?", vbYesNo + vbExclamation, "Confirm System Reset") = vbNo Then Exit Sub
If InputBox("Enter password to reset", "RESET") <> Format(Now, "mm/dd/yyyy") & " " & UserName Then MsgBox "Invalid password.", vbOKOnly + vbExclamation, "Error": Exit Sub
    Call connect
    connection.Execute "DELETE * FROM PatientRecords"
    connection.Execute "DELETE * FROM MedicalHistory"
    connection.Execute "DELETE * FROM Userlog"
    connection.Execute "DELETE * FROM Users"
    MsgBox "All records have been deleted. The application will now be closed.", vbOKOnly + vbInformation, "Reset Complete"
    End
End Sub

Private Sub mRptAll_Click()
Call connect
If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
rsMedicalRecord.Open "Select * from MedicalHistory  order by CheckupDate desc", connection, adOpenDynamic, adLockOptimistic
If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set DRMedicalHistoryAll.DataSource = rsMedicalRecord
DRMedicalHistoryAll.Orientation = rptOrientLandscape
DRMedicalHistoryAll.Show vbModal, Me
End Sub

Private Sub mRptIndividual_Click()
CallForm = "MedicalHistoryIndividual"
frmSearchPatient.Show vbModal, Me
Call connect
If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
rsMedicalRecord.Open "Select * from MedicalHistory where PatientID = " & CopyPatientID & "  order by CheckupDate desc", connection, adOpenDynamic, adLockOptimistic
If rsMedicalRecord.RecordCount = 0 Then Exit Sub
DRMedicalHistoryIndividual.Caption = "Medical History (" & rsMedicalRecord!FirstName & " " & rsMedicalRecord!LastName & ")"
DRMedicalHistoryIndividual.Sections(1).Controls("lblPatientID").Caption = "Patient ID: " & rsMedicalRecord!PatientID
DRMedicalHistoryIndividual.Sections(1).Controls("lblPatientName").Caption = "Patient Name: " & rsMedicalRecord!FirstName & " " & rsMedicalRecord!LastName
Set DRMedicalHistoryIndividual.DataSource = rsMedicalRecord
DRMedicalHistoryIndividual.Orientation = rptOrientLandscape
DRMedicalHistoryIndividual.Show vbModal, Me

End Sub

Private Sub mRptLogTrail_Click()
Call connect
If rsLogTrail.State = 1 Then rsLogTrail.Close
rsLogTrail.Open "Select * from UserLog order by LogDate desc, TimeIn desc", connection, adOpenDynamic, adLockOptimistic
If rsLogTrail.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set DRLogTrail.DataSource = rsLogTrail
DRLogTrail.Show vbModal, Me
End Sub

Private Sub mRptPatientList_Click()
Call connect
If rsPatient.State = 1 Then rsPatient.Close
rsPatient.Open "Select * from PatientRecords order by patientid asc", connection, adOpenDynamic, adLockOptimistic
If rsPatient.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set DRPatientList.DataSource = rsPatient
DRPatientList.Orientation = rptOrientLandscape
DRPatientList.Show vbModal, Me

End Sub

Private Sub mRptUsers_Click()
Call connect
If rsUsers.State = 1 Then rsUsers.Close
rsUsers.Open "Select * from Users order by UserID", connection, adOpenDynamic, adLockOptimistic
If rsUsers.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set DRUserList.DataSource = rsUsers
DRUserList.Show vbModal, Me
End Sub

Private Sub mViewAll_Click()
CallForm = "ViewAll"
Call connect
If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
rsMedicalRecord.Open "Select CheckupDate, LastName, FirstName, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory order by CheckupDate desc", connection, adOpenDynamic, adLockOptimistic
If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set frmPatientHistory.DataGrid1.DataSource = rsMedicalRecord
With frmPatientHistory
    .Caption = "Patient History (ALL)"
    .Frame1.Visible = True
    .Frame2.Visible = False
    .Show vbModal, Me
End With

End Sub

Private Sub mViewCalendar_Click()
frmCalendar.Show vbModal, Me
End Sub

Private Sub mViewIndividual_Click()
CallForm = "ViewIndividual"
frmSearchPatient.Show vbModal, Me
Call connect
If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
rsMedicalRecord.Open "Select CheckupDate, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory where PatientID = " & CopyPatientID & " order by CheckUpDate desc", connection, adOpenDynamic, adLockOptimistic
If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Call connect
If rsMedicalRecord2.State = 1 Then rsMedicalRecord2.Close
rsMedicalRecord2.Open "Select * from MedicalHistory where PatientID = " & CopyPatientID & " order by CheckUpDate desc", connection, adOpenDynamic, adLockOptimistic
If rsMedicalRecord2.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set frmPatientHistory.DataGrid1.DataSource = rsMedicalRecord
With frmPatientHistory
    .Caption = "Patient History (" & rsMedicalRecord2!FirstName & " " & rsMedicalRecord2!LastName & ")"
    .Frame1.Visible = False
    .Frame2.Visible = True
    .lblPatientID.Caption = "Patient ID: " & rsMedicalRecord2!PatientID
    .lblPatientName.Caption = "Patient Name: " & rsMedicalRecord2!FirstName & " " & rsMedicalRecord2!LastName
    .Show vbModal, Me
End With

End Sub

Private Sub mViewLogTrail_Click()
Call connect
If rsLogTrail.State = 1 Then rsLogTrail.Close
rsLogTrail.Open "Select * from UserLog order by LogDate desc, TimeIn desc", connection, adOpenDynamic, adLockOptimistic
If rsLogTrail.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
Set frmLogTrail.DataGrid1.DataSource = rsLogTrail
frmLogTrail.Show vbModal, Me
End Sub

Private Sub mViewPatientList_Click()
Dim ask2 As String
Call connect
If rsPatientList.State = 1 Then rsPatientList.Close
rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
If rsPatientList.RecordCount = 0 Then
    ask2 = MsgBox("No record found. Do you want to enter patient data?", vbYesNo + vbInformation, "No Records")
    If ask2 = vbNo Then
        Exit Sub
    Else
        frmAddNewPatient.Show vbModal, Me
    End If
Else
Set frmPatientList.DataGrid1.DataSource = rsPatientList
frmPatientList.Show vbModal, Me
End If
End Sub

Private Sub mViewUsers_Click()
Call connect
If rsUsers.State = 1 Then rsUsers.Close
rsUsers.Open "Select * from Users", connection, adOpenDynamic, adLockOptimistic
If rsUsers.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
frmUserList.Show vbModal, Me
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1).Text = Format(Now, "dddd, mmmm dd, yyyy") & " " & Format(Time, "hh:mm:ss AM/PM")

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'New Patient
        frmAddNewPatient.Show vbModal, Me
    Case 2 'New Patient Record
        Dim ask3 As String
        Call connect
        If rsPatientList.State = 1 Then rsPatientList.Close
        rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
        If rsPatientList.RecordCount = 0 Then
            ask3 = MsgBox("No record found. Do you want to enter patient data?", vbYesNo + vbInformation, "No Records")
            If ask3 = vbNo Then
                Exit Sub
            Else
                frmAddNewPatient.Show vbModal, Me
            End If
        Else
        CallForm = "AddNewPatientRecord"
        frmSearchPatient.Show vbModal, Me
        End If
    Case 3 'View Patient List
        Dim ask4 As String
        Call connect
        If rsPatientList.State = 1 Then rsPatientList.Close
        rsPatientList.Open "Select * from PatientRecords", connection, adOpenDynamic, adLockOptimistic
        If rsPatientList.RecordCount = 0 Then
            ask4 = MsgBox("No record found. Do you want to enter patient data?", vbYesNo + vbInformation, "No Records")
            If ask4 = vbNo Then
                Exit Sub
            Else
                frmAddNewPatient.Show vbModal, Me
            End If
        Else
        Set frmPatientList.DataGrid1.DataSource = rsPatientList
        frmPatientList.Show vbModal, Me
        End If
    Case 4 'View Patient History
        CallForm = "ViewAll"
        Call connect
        If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
        rsMedicalRecord.Open "Select CheckupDate, LastName, FirstName, ChiefComplaint, PhysicalExamination, Medication, Others from MedicalHistory order by CheckupDate desc", connection, adOpenDynamic, adLockOptimistic
        If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set frmPatientHistory.DataGrid1.DataSource = rsMedicalRecord
        With frmPatientHistory
            .Caption = "Patient History (ALL)"
            .Frame1.Visible = True
            .Frame2.Visible = False
            .Show vbModal, Me
        End With
    Case 5 'Print Patient List
        Call connect
        If rsPatient.State = 1 Then rsPatient.Close
        rsPatient.Open "Select * from PatientRecords order by patientid asc", connection, adOpenDynamic, adLockOptimistic
        If rsPatient.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set DRPatientList.DataSource = rsPatient
        DRPatientList.Orientation = rptOrientLandscape
        DRPatientList.Show vbModal, Me
    Case 6 'Print Patient History
        Call connect
        If rsMedicalRecord.State = 1 Then rsMedicalRecord.Close
        rsMedicalRecord.Open "Select * from MedicalHistory  order by CheckupDate desc", connection, adOpenDynamic, adLockOptimistic
        If rsMedicalRecord.RecordCount = 0 Then MsgBox "No record found.", vbOKOnly + vbInformation, "No Records": Exit Sub
        Set DRMedicalHistoryAll.DataSource = rsMedicalRecord
        DRMedicalHistoryAll.Orientation = rptOrientLandscape
        DRMedicalHistoryAll.Show vbModal, Me
    Case 7 'Lock Application
        frmLock.Show vbModal, Me
End Select
End Sub

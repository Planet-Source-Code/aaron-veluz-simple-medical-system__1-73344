Attribute VB_Name = "SystemModule"
Option Explicit
Global NewUserID As Long
Global UserID As Long
Global UserName As String
Global Password As String
Global Privilege As String
Global MistakenAttempts As Integer
Public connection As New ADODB.connection
Public rsPatient As New ADODB.Recordset
Public rsPatientList As New ADODB.Recordset
Public rsPatientRecord As New ADODB.Recordset
Public rsMedicalRecord As New ADODB.Recordset
Public rsMedicalRecord2 As New ADODB.Recordset
Public rsUsers As New ADODB.Recordset
Public rsUsers2 As New ADODB.Recordset
Public rsUsers3 As New ADODB.Recordset
Public rsUsers4 As New ADODB.Recordset
Public rsUserList As New ADODB.Recordset
Public rsLogTrail As New ADODB.Recordset
Public rsLogin As New ADODB.Recordset
Public rsLogin2 As New ADODB.Recordset
Public CallForm As String
Public CopyPatientID, Age As Integer
Public Sub connect()
connection.CursorLocation = adUseClient
If connection.State <> 1 Then
    connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MedicalSystem.mdb;Jet OLEDB:Database Password=123!@#"
End If
End Sub

Function ListSelector(strTemp As String, cboBox As ComboBox)
Dim X As Integer

X = 1
Do While X <= cboBox.ListCount
    If cboBox.List(X) = strTemp Then
    Exit Do
    End If
    X = X + 1
Loop
If X >= cboBox.ListCount Then
    X = 0
End If

cboBox.ListIndex = X

End Function


Public Function recfound(ByVal sField As String, ByVal sfindtext As String) As Boolean

Call connect
If rsLogin.State = 1 Then rsLogin.Close
rsLogin.Open "SELECT * FROM Users where StrComp(UserName, '" & sfindtext & "', 0) = 0", connection, adOpenDynamic, adLockOptimistic

If rsLogin.EOF Then
    recfound = False
Else
    recfound = True
    UserID = rsLogin!UserID
    UserName = rsLogin!UserName
    Password = rsLogin!Password
    Privilege = rsLogin!Privilege
End If
End Function
Function MistakeCounter()
If MistakenAttempts = 5 Then
    MsgBox "Maximum number of mistaken attempts has been reached. The application will now be closed.", vbOKOnly + vbCritical, "Exit Application"
    End
End If
End Function
Sub Main()
If App.PrevInstance = True Then
    MsgBox "Application is already running.", vbOKOnly + vbExclamation, "Run"
    End
End If
Call connect
If rsLogin.State = 1 Then rsLogin.Close
rsLogin.Open "Select * from Users", connection, adOpenDynamic, adLockOptimistic
If rsLogin.EOF Then
    rsLogin.Close
    Load frmSplash
    frmSplash.Show vbModal
    frmSetAdmin.Show vbModal
Else
    rsLogin.Close
    Load frmSplash
    frmSplash.Show vbModal
    frmLogin.Show vbModal
End If
End Sub
Public Sub Loginn()
Dim Counter As Integer
Call connect
If rsLogin2.State = 1 Then rsLogin2.Close
rsLogin2.Open "Select * from UserLog order by LogID", connection, adOpenDynamic, adLockOptimistic
If rsLogin2.RecordCount = 0 Then
    Counter = 1
ElseIf rsLogin2.RecordCount <> 0 Then
    rsLogin2.MoveLast
    Counter = rsLogin2!LogID + 1
End If
With rsLogin2
    .AddNew
    !LogID = Counter
    !UserName = UserName
    !Privilege = Privilege
    !LogDate = Format(Now, "mm/dd/yyyy")
    !TimeIn = Format(Now, "hh:mm:ss am/pm")
    .UpdateBatch adAffectCurrent
    .Close
End With
End Sub
Public Sub Logoutt()
Call connect
If rsLogin2.State = 1 Then rsLogin2.Close
rsLogin2.Open "Select * from UserLog order by LogID", connection, adOpenDynamic, adLockOptimistic
If rsLogin2.RecordCount = 0 Then Exit Sub
rsLogin2.MoveLast
rsLogin2!Timeout = Format(Time, "hh:mm:ss am/pm")
rsLogin2.UpdateBatch adAffectCurrent
rsLogin2.Close

End Sub
Public Sub EnableControls()
If Privilege = "Staff" Then
    With frmMain
        .mNewPatientRecord.Visible = False
        .mNewUSer.Visible = False
        .mNewPatientRecord.Visible = False
        .mnuLock.Visible = False
        .mViewPatientHistory.Visible = False
        .mViewUsers.Visible = False
        .mRptPatientHistory.Visible = False
        .mRptLogTrail.Visible = False
        .mRptUsers.Visible = False
        .mBackupRestore.Visible = False
        .mViewLogTrail.Visible = False
        .mReset.Visible = False
    End With
    With frmPatientList
        .Toolbar1.Buttons(3).Visible = False
        .Toolbar1.Buttons(4).Visible = False
    End With
    With frmMain
        .Toolbar1.Buttons(2).Visible = False
        .Toolbar1.Buttons(4).Visible = False
        .Toolbar1.Buttons(6).Visible = False
        .Toolbar1.Buttons(7).Visible = False
    End With
ElseIf Privilege = "Administrator" Then
    With frmMain
        .mNewPatientRecord.Visible = True
        .mNewUSer.Visible = True
        .mNewPatientRecord.Visible = True
        .mnuLock.Visible = True
        .mViewPatientHistory.Visible = True
        .mViewUsers.Visible = True
        .mRptPatientHistory.Visible = True
        .mRptLogTrail.Visible = True
        .mRptUsers.Visible = True
        .mBackupRestore.Visible = True
        .mViewLogTrail.Visible = True
        .mReset.Visible = False
    End With
    With frmPatientList
        .Toolbar1.Buttons(3).Visible = True
        .Toolbar1.Buttons(4).Visible = True
    End With
    With frmMain
        .Toolbar1.Buttons(2).Visible = True
        .Toolbar1.Buttons(4).Visible = True
        .Toolbar1.Buttons(6).Visible = True
        .Toolbar1.Buttons(7).Visible = True
    End With
ElseIf Privilege = "SuperAdministrator" Then
    With frmMain
        .mNewPatientRecord.Visible = True
        .mNewUSer.Visible = True
        .mNewPatientRecord.Visible = True
        .mnuLock.Visible = True
        .mViewPatientHistory.Visible = True
        .mViewUsers.Visible = True
        .mRptPatientHistory.Visible = True
        .mRptLogTrail.Visible = True
        .mRptUsers.Visible = True
        .mBackupRestore.Visible = True
        .mViewLogTrail.Visible = True
        .mReset.Visible = True
    End With
    With frmPatientList
        .Toolbar1.Buttons(3).Visible = True
        .Toolbar1.Buttons(4).Visible = True
    End With
    With frmMain
        .Toolbar1.Buttons(2).Visible = True
        .Toolbar1.Buttons(4).Visible = True
        .Toolbar1.Buttons(6).Visible = True
        .Toolbar1.Buttons(7).Visible = True
    End With
End If

End Sub
Public Sub ValidateAge(dtp As DTPicker, txt As TextBox)
Dim MM, DD, YYYY, NowM, NowD, NowY As Integer

NowM = Val(Format(Now, "MM"))
NowD = Val(Format(Now, "dd"))
NowY = Val(Format(Now, "yyyy"))
MM = Val(Format(dtp.Value, "MM"))
DD = Val(Format(dtp.Value, "dd"))
YYYY = Val(Format(dtp.Value, "yyyy"))
Age = NowY - YYYY
txt.Text = Age
If MM > NowM Or (MM = NowM And DD > NowD) Then
    Age = Age - 1
    txt.Text = Age
End If
If Age < 0 Then
    MsgBox "Birthdate is before current date.", vbOKOnly + vbExclamation, "Invalid BirthDate"
    dtp.SetFocus
End If
End Sub
Public Sub SelText(frm As Form, txt As TextBox)
With frm
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
    txt.SetFocus
End With

End Sub

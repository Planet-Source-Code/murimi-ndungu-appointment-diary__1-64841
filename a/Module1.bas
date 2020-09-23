Attribute VB_Name = "Module1"
'====================================================================================
'
' Developed by Murimi Ndungu
' murimixp @ gmail.com
'
' Kenyan, East Africa
'
'====================================================================================
'
' *****  READ THIS BEFORE USING THIS CODE:  ******
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Appointment diary fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author.
'
' The source code for Appointment diary has been submitted
' for the purposes of education.  I find the best way to learn is to
' look at how other people do things and see if i can possibly do it
' more efficiently. Contact me for additional help/suggestions via my
' email.


'Public declarations for this module
Public objConn As ADODB.Connection
Public objCmd As ADODB.Command
Public objRS As ADODB.Recordset
Public objGetEm As ADODB.Recordset
Public Regno As String
Public currentTime As Date
Public bcurrentTime As String
Public Appstatus As String
Public AppNum As String
Public sPopulation As Integer
Dim Message As String
Public strSQL As String
Public icol As Integer
Public sName As String
Public sEmail As String
Public sNum As String
Public sView As Date
Public GridAltColor As Long 'store grid color settings
Public DelOldApp As String 'store delete settings
Public BackupOldApp As String 'store backup settings
Public ClosedSat As String 'store csat settings
Public ClosedSun As String 'store csun settings
Public ClosedHol As String 'store chol settings
Public cWeb As String 'store web settings
Public myData As cData
Public CurrentUser As String
Public UserPassword As String
Public HolText  As String 'holiday text
Public HolDate As String 'holiday date
Public sServer As String 'store mail server settings
Public sFromAddress As String 'store from address settings
Public sSubject As String 'store mail subject settings
Public sBody As String 'store mail body settings
Public bolEditReadRS As Boolean    'Variable to tell what kind of recordset (Read or Edit mode)
Public vsTime As String 'validate
Public veTime As String 'validate
Public Const msgOld = "This Feature is not available in this version"
Public Const msgSaved = "Record Saved Successfully"
Public Const msgnoRecord = "Selected date bookings are full for the duration selected."
Public Const msgnoAppoint = "There are no appointments for the selected date"
Public Const msgNumbers = "Numerals Only"

'API declaration used to ensure Splash screen stays on top.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_HWNDPARENT = (-8)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub Main()

  On Error GoTo HandleErrors
      
  frmSplash.Show
  
  'Ensure the Splash form is refreshed prior to displaying the Main form.
  DoEvents
                
  '---------------------------------------------------------------------------------------------------------------------
  'Perform other start up tasks here...
On Error GoTo ErrHandler:
   Set myData = New cData
    myData.OpenDB (App.Path & "\Appointment Diary.mdb")
If Appstatus = "Invalid" Then
Message = MsgBox("Program could not be initialized contact system admin", vbCritical)
    
    Unload frmSplash
    Exit Sub
Else
    strSQL = "SELECT color, nUsers, Registration, OldDel, OldBack, cSat, cSun, cHol, AppNew, Website FROM owner"
    GetRecordSet (strSQL)
    If objRS("Appnew").Value = "Yes" Then
        Unload frmSplash
        frmAppNew.Show
        Exit Sub
    End If
    GridAltColor = objRS("color").Value
    sPopulation = objRS("nUsers").Value
    Regno = objRS("Registration").Value
    DelOldApp = objRS("OldDel").Value
    BackupOldApp = objRS("OldBack").Value
    ClosedSat = objRS("cSat").Value
    ClosedSun = objRS("cSun").Value
    ClosedHol = objRS("cHol").Value
    cWeb = objRS("Website").Value
      If DelOldApp = 1 Then
        strSQL = "DELETE Book.AppointmentNO, Book.RegID, Book.AppDate, Book.sTime, Book.eTime, Book.Purpose, Book.Notes, Book.Venue, Book.Source, Book.Status From Book WHERE Book.AppDate <'" & SQLDate(Date) & "'"
        exCommand (strSQL)
    End If
    Unload frmSplash
    frmLogin.Show
End If
Exit Sub
ErrHandler:
Message = MsgBox("Program could not be initialized contact system administrator", vbCritical)
Unload frmSplash
Exit Sub

  'DemoDelay
 
  DoEvents
      
  Unload frmSplash

ExitHandleErrors:
  
  Exit Sub
  
HandleErrors:

  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors

End Sub

Public Sub DemoDelay()

  '    PURPOSE: To provide a delay in program execution (4 seconds) to simulate a typical applications
  '             initialisation (eg. connecting to a database)...
  '
  '      NOTES: THIS SUB-ROUTINE IS NOT REQUIRED IN A PRODUCTION APPLICATION.
  '
  '  ARGUMENTS: None.
  '    RETURNS: None.
  '    AUTHORS: T.Cummins (TCC) - 12/09/2000.
  '---------------------------------------------------------------------------------------------------------------------
  
  On Error Resume Next
      
  Dim sngStartTime As Single
  sngStartTime = Timer
  Do Until (Timer - sngStartTime) > 2
      DoEvents
  Loop

End Sub
'holiday dates
Public Function HoliDate(ConvertDate As Date) As String
    HoliDate = Format(ConvertDate, "dd/mmmm")
End Function
'convert date to dd/mm/yy format
Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "dd/mm/yyyy")
End Function
'convert time to long time format
Public Function SQLTime(ConvertTime As Date) As String
    SQLTime = FormatDateTime(ConvertTime, 3)
End Function

Public Sub Display()
strSQL = "SELECT sTime, eTime, Purpose FROM Book where AppDate = '" & SQLDate(sView) & "' ORDER BY sTime"
GetRecordSet (strSQL)

CenterCells
SizeCells
'MsgBox objRS("sTime").Value
With MDIForm1.grdGrid
         .Clear
         .Rows = 0
    While Not objRS.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .CellBackColor = GridAltColor
        .Text = SQLTime(objRS("sTime")) & " to " & SQLTime(objRS("eTime")) & ""
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Text = objRS("Purpose") & ""
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .CellForeColor = RGB(0, 0, 255)
        .Text = "Show Details"
    objRS.MoveNext
    Wend
End With
MDIForm1.lblDisplay.Caption = "Appointments for " & SQLDate(sView)
End Sub

Function ConvNumTime(pvTime)
Dim first As String
Dim last As String
Dim separator As String
separator = ":"
    Select Case Len(pvTime)
        Case 1
            first = "00"
            last = "0" & pvTime
        Case 2
            first = "00"
            last = pvTime
        Case 3
            first = Mid(pvTime, 1, 1)
            last = Mid(pvTime, 2, 2)
        Case 4
            first = Mid(pvTime, 1, 2)
            last = Mid(pvTime, 3, 2)
        Case Else
            pvTime = "00:00"
    End Select
    ConvNumTime = first & separator & last
End Function


Public Sub SizeCells()
    Dim intColumn As Integer
      
    For intColumn = 0 To 0
        MDIForm1.grdGrid.ColWidth(intColumn) = 2900
    Next intColumn
    
End Sub

Public Sub CenterCells()
    Dim intColumn As Integer
    
    For intColumn = 0 To 0
        MDIForm1.grdGrid.ColAlignment(intColumn) = flexAlignCenterCenter
    Next intColumn
    
End Sub


Public Sub GetRecordSet(strSource As String) 'Get either Readable or Editable Recordset
On Error GoTo ErrHandler 'I take care of eventual errors
If objRS.State = adStateOpen Then objRS.Close 'If the database is open close before getting a new recordset
Select Case bolEditReadRS 'Tells which kind of recordset to get
    Case False
        With objRS
            .ActiveConnection = objConn
            .CursorType = adOpenKeyset 'Move the cursor in any direction and bookmarkable
            .LockType = adLockReadOnly 'Editing is not possible
            .Source = strSource 'What Recordset to get
            .Open
        End With
    Case True
        With objRS
            .ActiveConnection = objConn
            .CursorType = adOpenKeyset 'Move the cursor in any direction and bookmarkable
            .LockType = adLockOptimistic 'Editing is possible
            .Source = strSource 'What Recordset to get
            .Open
        End With
        
End Select
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub
Public Sub showForm(frm As Form)
On Error GoTo ErrHandler
        Load frm
      'specifies form position
      frm.Top = 0
      frm.Left = MDIForm1.picLeft.Left + 5
'      frm.Width = 6735
'      frm.Height = 5565
 
Exit Sub
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub
Public Function Decrypt(ByVal Password As String) As String
'Decrypt the password, "decode" the password
    Dim CodeString As String
    Dim Passcode As String
    Dim A As Integer
   
    Passcode = Left(CStr(Password), 30)
    CodeString = ""
    For A = 1 To Len(Passcode)
        CodeString = CodeString & CStr(Chr(Asc(Mid(Passcode, A, 1)) + 19))
    Next A
    Decrypt = CodeString

End Function
Public Function Encrypt(CodeString As TextBox) As String
'Encrypt the password, "encode" the password
    Dim Password As String
    Dim Passcode As String
    Dim A As Integer
    
    Passcode = Left(CStr(CodeString), 30)
    Password = ""
    For A = 1 To Len(Passcode)
        Password = Password & CStr(Chr(Asc(Mid(Passcode, A, 1)) - 19))
    Next A
    Encrypt = Password

End Function

Public Sub exCommand(strSource As String) 'Get either Readable or Editable Recordset
On Error GoTo ErrHandler 'I take care of eventual errors
With objCmd
            .ActiveConnection = objConn
            .CommandType = adCmdText
            .CommandText = strSource
            .Execute
End With
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Public Sub GetRecSet(strSource As String)
On Error GoTo ErrHandler 'I take care of eventual errors
Set objGetEm = New ADODB.Recordset
If objGetEm.State = adStateOpen Then objGetEm.Close 'If the database is open close before getting a new recordset
        With objGetEm
            .ActiveConnection = objConn
            .CursorType = adOpenKeyset 'Move the cursor in any direction and bookmarkable
            .LockType = adLockReadOnly 'Editing is not possible
            .Source = strSource 'What Recordset to get
            .Open
        End With
   Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Schedule"
   ClientHeight    =   5235
   ClientLeft      =   2985
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5730
   Begin VB.Frame Frame7 
      Caption         =   "Details"
      Height          =   1455
      Left            =   0
      TabIndex        =   25
      Top             =   3240
      Width           =   5655
      Begin VB.TextBox txtNotes 
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtVenue 
         Height          =   285
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmnew.frx":0742
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Venue"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Notes"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   4850
      Width           =   975
   End
   Begin VB.CommandButton cmdReschedule 
      Caption         =   "&Re-Schedule"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4855
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4855
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   5040
      Picture         =   "frmnew.frx":074B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdOPen 
      Caption         =   "Open"
      Height          =   495
      Left            =   4440
      Picture         =   "frmnew.frx":0C7D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "Purpose"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   5655
      Begin VB.OptionButton optCombo 
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboPurpose 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Purpose"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Day And Time"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   5655
      Begin VB.ComboBox cboEnd 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboStart 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54919169
         CurrentDate     =   37135
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   2280
         TabIndex        =   15
         Top             =   0
         Width           =   3375
         Begin VB.Label Label7 
            Caption         =   "End Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Start Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Day:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Contacts"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   4320
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   175
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   4855
         Width           =   1455
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   160
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4850
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aTime As Date
Dim bTime As Date
Dim cTime As Date
Dim nTime As Date
Dim sTime As Variant
Dim eTime As Variant
Dim Duration As Integer
Dim Counter As Variant
Dim bCounter As Variant
Dim pTime As Integer
Dim dTime As Integer
Dim Message As String
Dim CheckTime As Integer
Dim isValid As Boolean 'hold validations
Dim isRescheduled As Boolean


Function CDate2Julian(MyDate As Date) As String
' Taken from Microsoft Access 97 Knowledgebase, Article ID: Q162745

    CDate2Julian = Format(MyDate - DateSerial(Year(MyDate) - 1, 12, 31), "000")

End Function


Private Function GetAppNumber() As String

' Appointment Number based on Julian Date.  The Julian Date is the last two
'digits of the current year
' and the current day number.  The day number is determined by adding 1 to
'each day of the year.
' For example, day 1 is January 1, day 2 is January 2, day 10 is January 10,
'day 30 is January 30,
' and day 365 is December 31, (unless this is a leap year, then day 366 is
'December 31)
'
' The format for the Appointment No is yyddd00001

    Dim strLastJulian As String
    Dim lngLastIncrement As Long
    Dim strCurrentJulian As String
    
    ' Find the last Appointment Number
    strSQL = "SELECT Max(AppointmentNO) AS Last_Appointment_No  FROM Book"

    ' Open a recordset based on the strSQL statement
    GetRecordSet (strSQL)

    ' If there is no Last Appointment Number, set the lngLastIncrement to 0
    If IsNull(objRS!Last_Appointment_No) Then
        lngLastIncrement = 0
    Else
        ' Find the current Julian Date (the first 5 characters of the Last
'Appointment No)
        strLastJulian = Mid(objRS!Last_Appointment_No, 2, 5)
        ' Find the last increment number (the last 4 characters of the Last
'Appointment No)
        lngLastIncrement = Right(objRS!Last_Appointment_No, 4)
    End If

    ' Determine today's Julian Date
    strCurrentJulian = Right(DatePart("yyyy", Date), 2) & CDate2Julian(Date)
    strCurrentJulian = Left(strCurrentJulian, 5)

    ' If today's Julian Date is the Same as the Last Appointment No's Julian
'date then add 1 to
    ' the Last Appointment No
    If strCurrentJulian = strLastJulian Then
        strCurrentJulian = "L" + strCurrentJulian + Format(lngLastIncrement + 1, "0000")
    Else
        ' Otherwise, this is the first number of today's Julian Date
        strCurrentJulian = "L" + strCurrentJulian + "0001"
        
    End If
    GetAppNumber = strCurrentJulian
    ' Close the recordset and release resources
    objRS.Close
End Function

Private Sub cboPurpose_Click()
On Error GoTo ErrHandling:
strSQL = "select Purpose, pTime from Purpose where Purpose = '" & cboPurpose.Text & "' "
GetRecordSet (strSQL)
Duration = objRS("pTime")
cboStart.Clear
cboEnd.Clear
PopulateCombos
Exit Sub
ErrHandling:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cmdADD_Click()
On Error GoTo ErrHandling:
Dim sDay As String
Dim ssTime As String
Dim seTime As String
Dim sPurpose As String
Dim sNotes As String
Dim sVenue As String
isRescheduled = False
AppNum = GetAppNumber
sDay = DTPicker1.Value
ssTime = cboStart.Text
seTime = cboEnd.Text
If optOther.Value = True Then
    sPurpose = "Other"
Else
    sPurpose = cboPurpose.Text
End If
sNotes = txtNotes.Text
sVenue = txtVenue.Text

'validations
Validations
If isValid = False Then
    Exit Sub
End If
'If SQLTime("#'" & ssTime & "#'") < SQLTime("#'" & seTime & "#'") Then
' Message = MsgBox("Invalid Time, select valid time to continue", vbInformation)
' Exit Sub
'End If
'objConn.BeginTrans
'On Error GoTo ErrHandler:
strSQL = " INSERT INTO Book (AppointmentNO, RegID, AppDate, sTime, eTime, Purpose, Notes, Venue) "
strSQL = strSQL & " VALUES ('" & AppNum & "', " & sNum & ",'" & SQLDate(DTPicker1.Value) & "','" & ssTime & "','" & seTime & "','" & sPurpose & "','" & sNotes & "','" & sVenue & "')"
exCommand (strSQL)

'strSQL = ""
'strSQL = strSQL & "SELECT uDay, usTime, uUsers FROM nusers where uDay = '" & sDay & "' and usTime = '" & ssTime & "'"
'GetRecordSet (strSQL)
'If objRS.RecordCount = 0 Then
'    strSQL = " INSERT INTO Nusers (uDay, usTime, uUsers) "
'    strSQL = strSQL & " VALUES ('" & sDay & "','" & ssTime & "'," & sPopulation & ")"
'    exCommand(strSQL)
'Else
'    strSQL = "update nusers set uUsers = " & (objRS("uUsers").Value - 1)
'    exCommand(strSQL)
'End If

sView = sDay
Display
'Exit Sub
'ErrHandler:
'    objConn.RollbackTrans
'    MsgBox "Error :" & " " & Err.Description, vbCritical
'    Exit Sub
'objConn.CommitTrans
Message = MsgBox("Your Appointment number is: " & AppNum, vbInformation)
Message = ""
Message = MsgBox("Record Succesfully Added Add Another Record", vbInformation + vbYesNo)
If Message = vbYes Then
    cboStart.Clear
    cboEnd.Clear
    PopulateCombos
    lblName.Caption = ""
    lblEmail.Caption = ""
    sNum = ""
    txtNotes.Text = ""
    
Else
    Unload Me
End If
Exit Sub
ErrHandling:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cmdDel_Click()
On Error GoTo ErrHandler:
Dim sDay As String
sDay = SQLDate(DTPicker1.Value)
'validations
If lblName.Caption = "" Then
    MsgBox "Name is required"
Else
strSQL = "Delete * from Book Where AppointmentNO = '" & AppNum & "'"
exCommand (strSQL)
sView = sDay
Display
Message = MsgBox("Record Succesfully Deleted Delete Another Record", vbInformation + vbYesNo)
End If
If Message = vbYes Then
    cboStart.Clear
    cboEnd.Clear
    PopulateCombos
    lblName.Caption = ""
    lblEmail.Caption = ""
    AppNum = ""
    txtNotes.Text = ""
    
Else
    Unload Me
End If
Exit Sub
ErrHandler:
Message = MsgBox("Error while deleting appointment", vbCritical)
Unload Me
Exit Sub
End Sub

Private Sub cmdNew_Click()
showForm frmRegistration
cmdNew.Enabled = False
End Sub

Private Sub cmdOPen_Click()
If frmNew.Caption = "New Schedule" Then
    showForm frmFind
Else
    showForm frmSearch
End If
cmdOPen.Enabled = False
End Sub


Private Sub cmdQuit_Click()
vsTime = 0
veTime = 0
Unload Me
End Sub

Private Sub cmdReschedule_Click()
On Error GoTo ErrHandler:
'reschedule appointment
Dim sDay As String

Dim sPurpose As String
Dim sNotes As String
Dim sVenue As String
Dim ssTime As String 'validate
Dim seTime As String 'validate
isRescheduled = True
sDay = DTPicker1.Value
ssTime = cboStart.Text
seTime = cboEnd.Text
If optOther.Value = True Then
    sPurpose = "Other"
Else
    sPurpose = cboPurpose.Text
End If

sNotes = txtNotes.Text
sVenue = txtVenue.Text

'validate
Validations
If isValid = False Then
    Exit Sub
End If
strSQL = "UPDATE Book SET AppDate = '" & SQLDate(DTPicker1.Value) & "', sTime = '" & ssTime & "', eTime = '" & seTime & "', Purpose = '" & sPurpose & "', Notes = '" & sNotes & "', Venue = '" & sVenue & "' Where AppointmentNO = '" & AppNum & "'"
exCommand (strSQL)
vsTime = 0
veTime = 0
sView = sDay
Display
Message = MsgBox("Record Succesfully Updated Update Another Record", vbInformation + vbYesNo)
If Message = vbYes Then
    cboStart.Clear
    cboEnd.Clear
    PopulateCombos
    lblName.Caption = ""
    lblEmail.Caption = ""
    AppNum = ""
    txtNotes.Text = ""
    
Else
    Unload Me
End If
Exit Sub
ErrHandler:
Message = MsgBox("Error while re-scheduling appointment", vbCritical)
Unload Me
Exit Sub
End Sub

Private Sub Command2_Click()
cboStart.Clear
    cboEnd.Clear
    PopulateCombos
    lblName.Caption = ""
    lblEmail.Caption = ""
    sNum = ""
    txtNotes.Text = ""
    vsTime = 0
    veTime = 0
End Sub

Private Sub DTPicker1_Change()
cboStart.Clear
cboEnd.Clear
PopulateCombos
sView = SQLDate(DTPicker1.Value)
Display
End Sub


Private Sub Form_Load()
'set focus on purpose combobox
optCombo.Value = True
'set todays date on date picker
DTPicker1.Value = Date

'populate purpose combo
strSQL = "select PurposeID, Purpose, pTime from Purpose"
GetRecordSet (strSQL)
pTime = objRS("pTime")

objRS.MoveFirst
Do While Not objRS.EOF

    cboPurpose.AddItem objRS("Purpose")
 
objRS.MoveNext
Loop
cboPurpose.ListIndex = 0
End Sub


Private Function PopulateCombos()
On Error GoTo ErrHandler:
Dim ContinueCheck As Boolean 'for checking stime
Dim ResProc As Boolean ' for checking progess
ContinueCheck = True
ResProc = False
RestartProcess:
strSQL = "select sTime, eTime from Owner"
GetRecordSet (strSQL)
If ResProc = True Then
    sTime = FormatDateTime(sTime, 4)
    sTime = Replace(sTime, ":", "")
Else
    sTime = objRS("sTime")
End If
eTime = objRS("eTime")

'convert number to time format
sTime = ConvNumTime(sTime)
sTime = FormatDateTime(sTime, 4)
'first check to see whether an appointment has been booked in order to change stime
Do While ContinueCheck = True
strSQL = "select sTime, eTime from Book Where AppDate = '" & SQLDate(DTPicker1.Value) & "' And sTime = # " & sTime & "#"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
'no record continue with stime
    sTime = Format(sTime, "h:mm")
    sTime = Replace(sTime, ":", "")
    ContinueCheck = False
Else
' set stime to new record
    sTime = objRS("eTime")
    sTime = FormatDateTime(sTime, 4)
    ContinueCheck = True
End If
Loop


bCounter = CInt(sTime)

'clear both timeslots table
If ResProc = False Then
strSQL = "Delete * From TimeSlots "
exCommand (strSQL)
strSQL = ""
strSQL = "Delete * From TimeSlotsB "
exCommand (strSQL)
End If

'convert number to time format
sTime = ConvNumTime(sTime)

sTime = FormatDateTime(sTime, 4)
        
strSQL = " INSERT INTO TimeSlots (Times) "
strSQL = strSQL & " VALUES ('" & sTime & "')"
exCommand (strSQL)
strSQL = ""
strSQL = " INSERT INTO TimeSlotsB (Times) "
strSQL = strSQL & " VALUES ('" & sTime & "')"
exCommand (strSQL)
sTime = Replace(sTime, ":", "")


For Counter = sTime To eTime

        If bCounter > Counter Then
              eTime = eTime - 2400
        End If
  
    bCounter = Counter
    Counter = ConvNumTime(Counter)
    Counter = CDate(Counter)
    Counter = FormatDateTime(DateAdd("n", Duration, Counter), 4)
    Counter = Replace(Counter, ":", "")
    
        If CInt(Counter) > eTime Then
            Exit For
        End If
        
    Counter = ConvNumTime(Counter)
    
    strSQL = " INSERT INTO TimeSlots (Times) "
    strSQL = strSQL & " VALUES ('" & Counter & "')"
    exCommand (strSQL)
    strSQL = ""
    strSQL = " INSERT INTO TimeSlotsB (Times) "
    strSQL = strSQL & " VALUES ('" & Counter & "')"
    exCommand (strSQL)
    Counter = Replace(Counter, ":", "")
    Counter = Counter - 1
Next
'delete close time from first combo
bCounter = ConvNumTime(bCounter)
bCounter = FormatDateTime(CDate(bCounter), 4)
strSQL = "DELETE * FROM TimeSlots Where Times  = " & "#" & bCounter & "#" & ""
exCommand (strSQL)
'delete old times
strSQL = "select sTime, eTime from Book Where AppDate = '" & SQLDate(DTPicker1.Value) & "'"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
'should fill combos bottom

    FillCombos
    Exit Function
Else
 objRS.MoveFirst
 Do While Not objRS.EOF
 'checking time range between major assignments
    aTime = FormatDateTime(objRS("sTime"), vbShortTime)
    bTime = FormatDateTime(objRS("eTime"), vbShortTime)
    strSQL = "SElECT * FROM TimeSlotsB WHERE Times < " & "#" & FormatDateTime(aTime, vbShortTime) & "#" & "Order By Times Desc"
    GetRecSet (strSQL)
    If objGetEm.RecordCount > 0 Then
    objGetEm.MoveFirst
        'delete time range start if any records
        cTime = FormatDateTime(objGetEm("times"), vbShortTime)
        If DateDiff("n", cTime, aTime) < Duration Then
            'delete time range start if any records
            strSQL = "DELETE * FROM TimeSlots Where Times > " & "#" & aTime & "#" & " And Times < " & "#" & bTime & "#" & ""
            exCommand (strSQL)
            strSQL = "DELETE * FROM TimeSlotsB Where Times > " & "#" & aTime & "#" & " And Times < " & "#" & bTime & "#" & ""
            exCommand (strSQL)
            'delete all other records an set stime to etime of conflicts
               strSQL = "DELETE * FROM TimeSlots Where Times > " & "#" & FormatDateTime(bTime, vbShortTime) & "#" & ""
            exCommand (strSQL)
            strSQL = "DELETE * FROM TimeSlotsB Where Times > " & "#" & FormatDateTime(bTime, vbShortTime) & "#" & ""
            exCommand (strSQL)
            'delete time range end if any records
            strSQL = "DELETE * FROM TimeSlots Where Times IN (Select Top 1 Times from TimeSlots Order By Times Desc)"
            exCommand (strSQL)
            strSQL = "DELETE * FROM TimeSlotsB Where Times IN (Select Top 1 Times from TimeSlotsB Order By Times Desc)"
            exCommand (strSQL)
            'go to the label for filling times once more with new stime
            sTime = bTime
            ContinueCheck = True
            ResProc = True
            GoTo RestartProcess:
        End If
    End If
    
 'first use the select statement to know whether there are any records
    strSQL = "SELECT * FROM TimeSlotsB Where Times > " & "#" & FormatDateTime(aTime, vbShortTime) & "#" & " And Times < " & "#" & FormatDateTime(bTime, vbShortTime) & "#" & ""

    GetRecSet (strSQL)
    If objGetEm.RecordCount > 0 Then
        'delete time range start if any records
        strSQL = "DELETE * FROM TimeSlots Where Times > " & "#" & aTime & "#" & " And Times < " & "#" & bTime & "#" & ""
        exCommand (strSQL)
        strSQL = "DELETE * FROM TimeSlotsB Where Times > " & "#" & aTime & "#" & " And Times < " & "#" & bTime & "#" & ""
        exCommand (strSQL)
        'delete all other records an set stime to etime of conflicts
        sTime = bTime
        strSQL = "DELETE * FROM TimeSlots Where Times > " & "#" & FormatDateTime(bTime, vbShortTime) & "#" & ""
        exCommand (strSQL)
        strSQL = "DELETE * FROM TimeSlotsB Where Times > " & "#" & FormatDateTime(bTime, vbShortTime) & "#" & ""
        exCommand (strSQL)
        'delete time range end if any records
        strSQL = "DELETE * FROM TimeSlots Where Times IN (Select Top 1 Times from TimeSlots Order By Times Desc)"
        exCommand (strSQL)
        strSQL = "DELETE * FROM TimeSlotsB Where Times IN (Select Top 1 Times from TimeSlotsB Order By Times Desc)"
        exCommand (strSQL)
        'go to the label for filling times once more with new stime
        ContinueCheck = True
        ResProc = True
        GoTo RestartProcess:
    End If
    
    objRS.MoveNext
 Loop
End If
strSQL = " SELECT uDay, uUsers FROM nusers Where uDay = '" & SQLDate(DTPicker1.Value) & "' AND uUsers > 0"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
    strSQL = "DELETE * FROM TimeSlots Where Times IN (Select sTime from Book Where AppDate = '" & SQLDate(DTPicker1.Value) & "'" & ")"
    exCommand (strSQL)
    strSQL = "DELETE * FROM TimeSlotsB Where Times IN (Select eTime from Book Where AppDate = '" & SQLDate(DTPicker1.Value) & "'" & ")"
    exCommand (strSQL)
End If

objRS.Close
cboStart.Clear
cboEnd.Clear
FillCombos
Exit Function
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
End Function

Private Sub Form_Unload(Cancel As Integer)
'reset buttons
Select Case frmNew.Caption
 Case "New Schedule"
    MDIForm1.optNew.Value = False
 Case "Re-Schedule"
    MDIForm1.optEdit.Value = False
 Case "Delete"
    MDIForm1.optCancel.Value = False
End Select
  
End Sub

Private Sub optCombo_Click()
On Error GoTo ErrHandling:
If cboPurpose.Text <> "" Then
strSQL = "select Purpose, pTime from Purpose where Purpose = '" & cboPurpose.Text & "' "
GetRecordSet (strSQL)
Duration = objRS("pTime")
cboStart.Clear
cboEnd.Clear
PopulateCombos
End If
Exit Sub
ErrHandling:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub optOther_Click()
On Error GoTo ErrHandler:
strSQL = "select oDuration from Owner"
GetRecordSet (strSQL)
Duration = objRS("oDuration")
cboStart.Clear
cboEnd.Clear
PopulateCombos
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub

End Sub

Public Function FillCombos()
On Error GoTo ErrHandler:
'Open the recordset getting a list of all clubs
'fill start combo box
strSQL = "Select Times from TimeSlots"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
 MsgBox msgnoRecord
 Exit Function
Else
 objRS.MoveFirst
Do While Not objRS.EOF

    cboStart.AddItem objRS("Times")
    objRS.MoveNext
    
Loop
cboStart.ListIndex = 0
End If
'fill end combo box
strSQL = "Select Times from TimeSlotsB"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
 MsgBox msgnoRecord
 Exit Function
Else
 objRS.MoveFirst
Do While Not objRS.EOF

    cboEnd.AddItem objRS("Times")
    objRS.MoveNext
    
Loop
If cboStart.List(0) = cboEnd.List(0) Then
    cboEnd.RemoveItem (0)
End If
 cboEnd.ListIndex = 0
End If
Exit Function
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
End Function

Private Sub Validations()
Dim sDay As String
Dim HolDay As String
Dim cSat As String
Dim cSun As String

sDay = DTPicker1.Value

'validations
If lblName.Caption = "" Then
    MsgBox "Name is required"
    isValid = False
    Exit Sub
End If
'validate time
'check for time difference
CheckTime = DateDiff("n", FormatDateTime(cboStart.Text, 4), FormatDateTime(cboEnd.Text, 4))
If CheckTime <= 0 Then
    Message = MsgBox("The end time is earlier than the start time", vbInformation)
    isValid = False
    Exit Sub
End If
'check time booked between
strSQL = "select sTime, eTime from Book Where AppDate = '" & SQLDate(DTPicker1.Value) & "'"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
'no problem
 isValid = True
Else
 objRS.MoveFirst
 Do While Not objRS.EOF
    aTime = FormatDateTime(objRS("sTime"), vbShortTime)
    bTime = FormatDateTime(objRS("eTime"), vbShortTime)
    cTime = FormatDateTime(cboStart.Text, vbShortTime)
    nTime = FormatDateTime(cboEnd.Text, vbShortTime)
    If aTime > cTime And aTime < nTime Then
    
    If isRescheduled = True Then
    If aTime <> vsTime And bTime <> veTime Then
        Message = MsgBox("Another appointmntment is scheduled between " & aTime & " and " & bTime & ". Thus causing conflict.", vbInformation)
        isValid = False
        Exit Sub
    End If
    Else
        Message = MsgBox("Another appointmntment is scheduled between " & aTime & " and " & bTime & ". Thus causing conflict.", vbInformation)
        isValid = False
        Exit Sub
    End If
    End If
    objRS.MoveNext
 Loop
End If

'validate date
'check for earlier date
If CDate(sDay) < CDate(Date) Then
 Message = MsgBox("Earlier Date, select valid date to continue", vbInformation)
 isValid = False
 Exit Sub
End If
If ClosedSat = 1 Then
    'remove saturdays
    cSat = Weekday(sDay)
    If cSat = 7 Then
        Message = MsgBox("Closed on Saturdays, select other day to continue", vbInformation)
        isValid = False
        Exit Sub
    End If
End If
If ClosedSun = 1 Then
    'remove sundays
    cSun = Weekday(sDay)
    If cSun = 1 Then
        Message = MsgBox("Closed on Sundays, select other day to continue", vbInformation)
        isValid = False
        Exit Sub
    End If
End If

'check for Holiday
If ClosedHol = 1 Then
HolDay = HoliDate(DTPicker1.Value)
strSQL = "SELECT hsDate From Holidays where hsDate = '" & HolDay & "'"
GetRecordSet (strSQL)
If objRS.RecordCount > 0 Then
 Message = MsgBox("The date selected is a holiday, select valid date to continue", vbInformation)
 isValid = False
 Exit Sub
End If
End If
isValid = True
End Sub

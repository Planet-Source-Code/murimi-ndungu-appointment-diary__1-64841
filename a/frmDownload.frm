VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download Data"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6870
   Begin VB.TextBox Status 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   5160
      Width           =   6855
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdReschedule 
      Caption         =   "Reschedule"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Check Status"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save All"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin MSFlexGridLib.MSFlexGrid grdRegistration 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid grdDownload 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2355
         _Version        =   393216
         RowHeightMin    =   350
         FocusRect       =   2
         AllowUserResizing=   1
         BorderStyle     =   0
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLoading 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'check whether there is data to download
Dim NotEmpty As Boolean
' Recordset definition
Dim rs As New ADODB.Recordset
'if the link is invalid or mis-spelt then the error message is:
'Operation is not allowed on an object referencing a closed or invalid connection.
'Const urlData = "http://victory/"
Dim CheckEmail As String
Dim EmailExists As Boolean
Dim sSource As String
Dim sStatus As String
Dim sDay As String
Dim ssTime As String
Dim seTime As String
Dim sPurpose As String
Dim sNotes As String
Dim sVenue As String
Dim sEmail As String
Dim sName As String
Dim numTimes As Integer
Dim Progress
Dim Green_Light As Boolean
Dim DATAFile As String

Public Sub makegrd()
grdDownload.ColWidth(0) = 500
grdDownload.Cols = 10
grdDownload.Rows = 1
grdDownload.Row = 0
grdDownload.Col = 1
grdDownload.Text = "Name"
grdDownload.Col = 2
grdDownload.Text = "Email"
grdDownload.Col = 3
grdDownload.Text = "Date"
grdDownload.Col = 4
grdDownload.Text = "Starting Time"
grdDownload.Col = 5
grdDownload.Text = "Ending Time"
grdDownload.Col = 6
grdDownload.Text = "Purpose"
grdDownload.Col = 7
grdDownload.Text = "Notes"
grdDownload.Col = 8
grdDownload.Text = "Venue"
grdDownload.Col = 9
grdDownload.Text = "Status"
'reg grid
grdRegistration.ColWidth(0) = 500
grdRegistration.Cols = 9
grdRegistration.Rows = 1
grdRegistration.Row = 0
grdRegistration.Col = 1
grdRegistration.Text = "Name"
grdRegistration.Col = 2
grdRegistration.Text = "Email"
grdRegistration.Col = 3
grdRegistration.Text = "Address"
grdRegistration.Col = 4
grdRegistration.Text = "Age"
grdRegistration.Col = 5
grdRegistration.Text = "City"
grdRegistration.Col = 6
grdRegistration.Text = "Country"
grdRegistration.Col = 7
grdRegistration.Text = "Occupation"
grdRegistration.Col = 8
grdRegistration.Text = "Status"
End Sub

Private Sub cmdDownload_Click()
On Error GoTo ErrHandler:
MsgBox msgOld
Exit Sub
'fill cweb when new
If cWeb = "" Then
    strSQL = "SELECT Website FROM owner"
    GetRecordSet (strSQL)
    cWeb = objRS("Website").Value
End If
'MsgBox msgOld
'Exit Sub
prgLoad.Value = 15
   Dim Gridline As String
    Screen.MousePointer = vbHourglass
    ' Get data
    Set rs = Nothing
    prgLoad.Value = 15
    ' Change to your URL address - I will leave my link working as long as possible
    rs.Open cWeb & "getdata.asp"
   prgLoad.Value = 35
    ' Display for edits
    'Set grdDownload.DataSource = rs
    'populate grid
If rs.EOF = True Then
MsgBox "No records"
NotEmpty = False
 Exit Sub
Else
NotEmpty = True
grdDownload.Rows = 1
rs.MoveFirst
  Do While Not rs.EOF
  'show percent on prog bar
      prgLoad.Value = 84 - prgLoad.Value + 1

    With rs

      Gridline = vbTab & !Name & vbTab & !Email & vbTab & !appdate & vbTab & !sTime & vbTab & !eTime & vbTab
      Gridline = Gridline & !purpose & vbTab & !notes & vbTab & !venue

     grdDownload.AddItem Gridline
      .MoveNext
    End With
    Loop
End If
prgLoad.Value = 85
    Screen.MousePointer = vbNormal
    grdDownload.Rows = grdDownload.Rows + 1
    cmdStatus.Enabled = True
    cmdDownload.Enabled = False
    prgLoad.Value = 100
    Exit Sub
ErrHandler:
'If obj.Status >= 400 And obj.Status <= 599 Then
'  MsgBox "Error Occurred : " & obj.Status & " - " & obj.statusText
'Else
'  MsgBox obj.responseText
'End If
Exit Sub
End Sub


Private Sub cmdReschedule_Click()
Dim SMTP_HOST As String
Dim MAIL_FROM As String
Dim RCPT_TO As String
Dim SUBJECT As String
Dim Data1 As String
Dim Data2 As String
Dim FROM As String
Dim MAIL_TO As String
'send mail to all who are rescheduled
On Error GoTo ErrHandler
'get mail settings
strSQL = "SELECT Mail.PopServer, Mail.FromAddress, Mail.Subject, Mail.Message, Mail.Message2, Mail.Name FROM Mail"
GetRecordSet (strSQL)
SMTP_HOST = objRS("PopServer").Value
MAIL_FROM = objRS("FromAddress").Value
SUBJECT = objRS("Subject").Value
Data1 = objRS("Message").Value
Data2 = objRS("Message2").Value
FROM = objRS("Name").Value
'get rescheduled addresses
With grdDownload
 .Row = 1
 Do While Not .Row >= .Rows - 1
 'check whether to be rescheduled
    .Col = 1
     sName = .Text
    .Col = 2
    sEmail = .Text
    .Col = 9
     sStatus = .Text
        
'resolve earlier date
If sStatus = "Older Date" Then
 MAIL_TO = sName
 RCPT_TO = sEmail
 'send mail
 Winsock1.Close
Winsock1.Connect SMTP_HOST, "25"
numTimes = 0
Do While Winsock1.State <> sckConnected
DoEvents
Status.Text = "Connecting to " & SMTP_HOST & ". Please wait."
numTimes = numTimes + 1
If numTimes > 8 Then
    MsgBox "Request Timed Out"
    cmdSave.Enabled = True
    Exit Sub
End If
Loop
Status.Text = "Connected to " & SMTP_HOST & "."
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "Connected to " & SMTP_HOST & "." & Chr$(13) & Chr$(10)
Do While Green_Light = False
DoEvents
Status.Text = "Waiting for reply..."
Loop
Winsock1.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10)
Do While Progress <> 1
DoEvents
Status.Text = "Sending data. (1 of 3)"
Loop
Winsock1.SendData "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)
Do While Progress <> 2
DoEvents
Status.Text = "Sending data. (2 of 3)"
Loop
Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "DATA" & Chr$(13) & Chr$(10)
Do While Progress <> 3
DoEvents
Status.Text = "Setting up body transfer..."
Loop
Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10)
Winsock1.SendData Data2 & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & Data & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
Do While Progress <> 4
DoEvents
Status.Text = "Sending data. (3 of 3)"
Loop
Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
Status.Text = "Done"
Winsock1.Close

End If

'resolve conflict
If sStatus = "Conflict" Then
  MAIL_TO = sName
 RCPT_TO = sEmail
 'send mail
 Winsock1.Close
Winsock1.Connect SMTP_HOST, "25"
numTimes = 0
Do While Winsock1.State <> sckConnected
DoEvents
Status.Text = "Connecting to " & SMTP_HOST & ". Please wait."
numTimes = numTimes + 1
If numTimes > 8 Then
    MsgBox "Request Timed Out"
    cmdSave.Enabled = True
    Exit Sub
End If
Loop
Status.Text = "Connected to " & SMTP_HOST & "."
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "Connected to " & SMTP_HOST & "." & Chr$(13) & Chr$(10)
Do While Green_Light = False
DoEvents
Status.Text = "Waiting for reply..."
Loop
Winsock1.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10)
Do While Progress <> 1
DoEvents
Status.Text = "Sending data. (1 of 3)"
Loop
Winsock1.SendData "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)
Do While Progress <> 2
DoEvents
Status.Text = "Sending data. (2 of 3)"
Loop
Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & "DATA" & Chr$(13) & Chr$(10)
Do While Progress <> 3
DoEvents
Status.Text = "Setting up body transfer..."
Loop
Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10)
Winsock1.SendData Data1 & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & Data & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
'LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
Do While Progress <> 4
DoEvents
Status.Text = "Sending data. (3 of 3)"
Loop
Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
Status.Text = "Done"
Winsock1.Close

End If
   .Row = .Row + 1
   
 Loop
End With


cmdReschedule.Enabled = False
cmdSave.Enabled = True
'ConnectToServer sServer, Winsock1
Exit Sub
ErrHandler:
MsgBox "Error while rescheduling appointments" & Err.Description
cmdReschedule.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
sSource = "Web"
sStatus = "False"
On Error GoTo ErrHandler:
With grdDownload
 .Row = 1
 Do While Not .Row > .Rows
 'check whether email exists
    .Col = 2
    CheckEmail = .Text
    strSQL = "SELECT Email, RegistrationID from Registration where Email='" & CheckEmail & "'"
    GetRecordSet (strSQL)
  If objRS.RecordCount = 0 Then

   Dim Gridline As String
    Screen.MousePointer = vbHourglass
    ' Get data
    Set rs = Nothing
    ' Change to your URL address - I will leave my link working as long as possible
    rs.Open cWeb & "getregdata.asp?ID=" & CheckEmail & ""
   
    ' Display for edits
    'Set DataGrid.DataSource = rs
    'populate grid
    If rs.EOF = True Then
    MsgBox "No records"
    Exit Sub
    Else
    grdRegistration.Rows = 1
    rs.MoveFirst
     Do While Not rs.EOF
    With rs

      Gridline = vbTab & !Name & vbTab & !Email & vbTab & !Address & vbTab & !Age & vbTab & !city & vbTab
      Gridline = Gridline & !Country & vbTab & !Occupation

     grdRegistration.AddItem Gridline
      .MoveNext
    End With
    Loop

End If
    Screen.MousePointer = vbNormal
EmailExists = False
        Else
            EmailExists = True
            'get values
            sNum = objRS("RegistrationID").Value
            .Col = 2
            sDay = .Text
            .Col = 3
            ssTime = .Text
            .Col = 4
            seTime = .Text
            .Col = 5
            sPurpose = .Text
            .Col = 6
            sNotes = .Text
            .Col = 7
            sVenue = .Text
            'insert values
            strSQL = " INSERT INTO Book (RegID, AppDate, sTime, eTime, Purpose, Notes, Venue, Source, Status) "
            strSQL = strSQL & " VALUES (" & sNum & ",'" & sDay & "','" & ssTime & "','" & seTime & "','" & sPurpose & "','" & sNotes & "','" & sVenue & "','" & sSource & "','" & sStatus & "')"
            exCommand (strSQL)
            grdDownload.RemoveItem .Row
        End If
   .Row = .Row + 1
 Loop
End With
cmdSave.Enabled = False
cmdDownload.Enabled = True
Exit Sub
ErrHandler:
'If obj.Status >= 400 And obj.Status <= 599 Then
'  MsgBox "Error Occurred : " & obj.Status & " - " & obj.statusText
'Else
'  MsgBox obj.responseText
'End If
Exit Sub
End Sub

Private Sub cmdStatus_Click()
Dim DayFine As Boolean 'check whether earlier date
On Error GoTo ErrHandler:
With grdDownload
 .Row = 1
 Do While Not .Row >= .Rows - 1
 'check whether time exists
    .Col = 2
    sEmail = .Text
    .Col = 3
     sDay = .Text
    .Col = 4
     ssTime = .Text
     
'check for earlier date
If sDay <> "" And CDate(sDay) < CDate(Date) Then
 DayFine = False
 .Col = 0
 .CellPictureAlignment = flexAlignCenterTop
  Set .CellPicture = LoadPicture(App.Path & "\wrong.jpg")
 .Col = 9
 .Text = "Older Date"
 cmdReschedule.Enabled = True
Else
 DayFine = True
End If

If ssTime <> "" And DayFine = True Then
    strSQL = "SELECT Appdate, stime from Book where AppDate = '" & sDay & "' and sTime = #" & ssTime & "#"
    GetRecordSet (strSQL)
        If objRS.RecordCount = 0 Then
            .Col = 0
            .CellPictureAlignment = flexAlignCenterTop
        Set .CellPicture = LoadPicture(App.Path & "\correct.jpg")
            .Col = 9
            .Text = "Okey"
        Else
            .Col = 0
            .CellPictureAlignment = flexAlignCenterTop
        Set .CellPicture = LoadPicture(App.Path & "\wrong.jpg")
            .Col = 9
            .Text = "Conflict"
             cmdReschedule.Enabled = True
            ' ConnectToServer sServer, Winsock1
            '.RemoveItem .Row
        End If
End If
   .Row = .Row + 1
   
 Loop
End With
cmdStatus.Enabled = False

If cmdReschedule.Enabled = True Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If

Exit Sub
ErrHandler:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
makegrd
prgLoad.Max = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set rs = Nothing
    rs.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Winsock1.GetData DATAFile
'Reply = Mid(DATAFile, 1, 3)
''LOG_FORM.LOG_TEXT.Text = LOG_FORM.LOG_TEXT.Text & DATAFile & Chr$(13) & Chr$(10)
'If Reply = 250 Or Reply = 354 Then
'Progress = Progress + 1
'End If
'If Reply = 220 Then
'Green_Light = True
'End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'        MsgBox "Error while rescheduling appointment" & vbCrLf & "Error Number: " & Number & vbCrLf & Description & vbCrLf & Source, vbCritical + vbOKOnly, Me.Caption
'    cmdReschedule.Enabled = False
'    cmdSave.Enabled = True
'    Exit Sub
End Sub

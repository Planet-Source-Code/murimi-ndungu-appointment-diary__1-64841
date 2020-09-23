VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Appointment Diary - Beta Version"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7575
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   7575
      TabIndex        =   9
      Top             =   6195
      Width           =   7575
      Begin VB.PictureBox Picture4 
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   11790
         TabIndex        =   10
         Top             =   120
         Width           =   11850
         Begin VB.TextBox txtVenue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txtNotes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   8640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblAppNumber 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1560
            TabIndex        =   32
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Appointment Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblPurpose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblEnd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   8640
            TabIndex        =   23
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblStart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   8640
            TabIndex        =   22
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   8640
            TabIndex        =   21
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label14 
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7680
            TabIndex        =   20
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "venue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Start Time:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7680
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "End Time:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7680
            TabIndex        =   15
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Day"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7680
            TabIndex        =   14
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Purpose"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   12
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblEmail 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   840
            Width           =   3015
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   6195
      Left            =   4575
      ScaleHeight     =   6135
      ScaleWidth      =   2940
      TabIndex        =   5
      Top             =   0
      Width           =   3000
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   56557569
         CurrentDate     =   37135
      End
      Begin MSFlexGridLib.MSFlexGrid grdGrid 
         Height          =   3135
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         BorderStyle     =   0
      End
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BackColor       =   &H00FFC0C0&
      Height          =   6195
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   1515
      Begin VB.OptionButton optEdit 
         Height          =   555
         Left            =   240
         Picture         =   "MDIForm1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton optSetup 
         Height          =   555
         Left            =   240
         Picture         =   "MDIForm1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.OptionButton optCancel 
         Height          =   555
         Left            =   240
         Picture         =   "MDIForm1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3600
         Width           =   975
      End
      Begin VB.OptionButton optNew 
         Height          =   555
         Left            =   240
         Picture         =   "MDIForm1.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optView 
         Height          =   555
         Left            =   240
         Picture         =   "MDIForm1.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Re-Schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblSetup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblNew 
         BackColor       =   &H00FFC0C0&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblView 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Download"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAppointment 
      Caption         =   "&Appointments"
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnureschedule 
         Caption         =   "&Re-Schedule"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnumail 
         Caption         =   "&Configure Mail"
      End
      Begin VB.Menu mnusetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Database"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHquick 
         Caption         =   "&Quick Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit
Dim ChildCount As Integer
Dim i As Integer
Dim strSQL As String


Private Sub grdGrid_Click()
On Error GoTo ErrorHandler
ClearAllLabels
grdGrid.Col = 0
If grdGrid.Text = "Show Details" Then
    grdGrid.Row = grdGrid.Row - 2
    bcurrentTime = grdGrid.Text
    bcurrentTime = Mid(bcurrentTime, 1, 10)
    currentTime = CDate(bcurrentTime)
        strSQL = "Select Registration.Name, Registration.Email, Book.AppointmentNO, Book.AppDate, Book.sTime, Book.eTime, Book.Purpose, Book.Notes, Book.Venue From Registration INNER JOIN Book  on Registration.RegistrationID = Book.RegID where Book.Appdate = '" & SQLDate(sView) & "'" & "AND sTime = " & "#" & SQLTime(currentTime) & "#"
    GetRecordSet (strSQL)
    lblAppNumber.Caption = objRS("AppointmentNO").Value
    lblName.Caption = objRS("Name").Value
    lblEmail.Caption = objRS("Email").Value
    lblPurpose.Caption = objRS("Purpose").Value
    lblDay.Caption = objRS("appdate").Value
    lblStart.Caption = objRS("stime").Value
    lblEnd.Caption = objRS("etime").Value
    txtNotes.Text = objRS("Notes").Value
    txtVenue.Text = objRS("Venue").Value
Else
    Exit Sub
End If

Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
   
End Sub

Private Sub mnuAbout_Click()
showForm frmAbout
End Sub

Private Sub mnuBackup_Click()
MsgBox msgOld
End Sub

Private Sub mnuCancel_Click()
lblCancel_Click
End Sub

Private Sub mnuCompact_Click()
Dim oJet As JRO.JetEngine
Dim mPassword As String
mPassword = "HarD24GeT$aS"
Set oJet = New JRO.JetEngine
On Error GoTo ErrorHandler
objConn.Close
Set objConn = Nothing

'First reference JRO - Select Project - References - Microsoft Jet and Replication Objects
'App.Path refers to the place that this compact program and the database is installed
'You can replace app.path with another path such as:
'"If Dir("C:\Program Files\The Name of your data base.mde") or mdb
'When compacting a database you have to make a copy
'Usually I add a number to the end of the original database
'First we'll check to make sure that the database copy does not already exist

   If Dir(App.Path & "\appdiary1.mdb") <> "" Then Kill _
      "" & App.Path & "\appdiary1.mdb"


oJet.CompactDatabase _
   "Provider=Microsoft.Jet.OLEDB.4.0;" _
   & "Data Source=" & App.Path & "\appdiary.mdb" _
   & ";Jet OLEDB:Database Password=" & mPassword, _
   "Provider=Microsoft.Jet.OLEDB.4.0;" _
   & "Data Source=" & App.Path & "\appdiary1.mdb;" _
   & "Jet OLEDB:Engine Type = 4" _
   & ";Jet OLEDB:Database Password=" & mPassword
   'References Access2000 for Access97 use number 4 instead of 5

'The procedure above creates a compacted and repaired copy of your database

' Now delete the old\original database
   Kill App.Path & "\appdiary.mdb"

   ' Rename the new compacted database back to the original name
   Name App.Path & "\appdiary1.mdb" As App.Path & "\appdiary.mdb"

'reopen the connection
 Set myData = New cData
    myData.OpenDB (App.Path & "\AppDiary.mdb")
    
    'Just a message letting the user know the database was compacted
MsgBox "Your database was successfully compacted and repaired"

Exit Sub

ErrorHandler:

   MsgBox "There was a problem compacting your database. Make sure" _
      & " this program is installed in the same path as your database." _
      & " The program will now close." _
      & Err.Description
End Sub

Private Sub mnuDownload_Click()
'strSQL = "SELECT PopServer, FromAddress, Subject, Message from Mail"
'GetRecordSet (strSQL)
'sServer = objRS("PopServer").Value
'sFromAddress = objRS("FromAddress").Value
'sSubject = objRS("Subject").Value
'sBody = objRS("Message").Value
If GetNetConnectString = "Not connected to the internet now." Then
    MsgBox "Not connected to the internet now. Try again later"
    optView.Value = False
    Exit Sub
End If
showForm frmDownload
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHquick_Click()
MsgBox msgOld
End Sub

Private Sub mnumail_Click()
showForm frmMail
End Sub

Private Sub mnuNew_Click()
lblNew_Click
End Sub

Private Sub mnuRegister_Click()
MsgBox "Visit register"
End Sub

Private Sub mnureschedule_Click()
lblEdit_Click
End Sub

Private Sub mnusetup_Click()
showForm frmAdmin
End Sub

Private Sub optCancel_Click()
frmNew.Caption = "Delete"
frmNew.cmdNew.Enabled = False
frmNew.cmdADD.Visible = False
frmNew.cmdReschedule.Visible = False
frmNew.cmdDel.Visible = True
showForm frmNew
End Sub

Private Sub optEdit_Click()
frmNew.Caption = "Re-Schedule"
frmNew.cmdNew.Enabled = False
frmNew.cmdReschedule.Visible = True
frmNew.cmdADD.Visible = False
frmNew.cmdDel.Visible = False
showForm frmNew
End Sub



Private Sub optNew_Click()
frmNew.Caption = "New Schedule"
frmNew.cmdNew.Enabled = True
frmNew.cmdADD.Visible = True
frmNew.cmdReschedule.Visible = False
frmNew.cmdDel.Visible = False
showForm frmNew
End Sub



Private Sub optSetup_Click()
showForm frmAdmin
End Sub



Private Sub optView_Click()
'strSQL = "SELECT PopServer, FromAddress, Subject, Message from Mail"
'GetRecordSet (strSQL)
'sServer = objRS("PopServer").Value
'sFromAddress = objRS("FromAddress").Value
'sSubject = objRS("Subject").Value
'sBody = objRS("Message").Value
'If GetNetConnectString = "Not connected to the internet now." Then
'    MsgBox "Not connected to the internet now. Try again later"
'    optView.Value = False
'   Exit Sub
'End If
showForm frmDownload
optView.Value = False
End Sub

Private Sub lblCancel_Click()
frmNew.Caption = "Delete"
frmNew.cmdNew.Enabled = False
frmNew.cmdReschedule.Visible = False
frmNew.cmdADD.Visible = False
frmNew.cmdDel.Visible = True
showForm frmNew
End Sub

Private Sub lblEdit_Click()
showForm frmNew
frmNew.Caption = "Re-Schedule"
frmNew.cmdNew.Enabled = False
frmNew.cmdReschedule.Visible = True
frmNew.cmdADD.Visible = False
frmNew.cmdDel.Visible = False
End Sub

Private Sub lblNew_Click()
showForm frmNew
frmNew.cmdReschedule.Visible = False
frmNew.cmdADD.Visible = True
frmNew.cmdDel.Visible = False
End Sub


Private Sub lblSetup_Click()
showForm frmAdmin
End Sub

Private Sub lblView_Click()
'strSQL = "SELECT PopServer, FromAddress, Subject, Message from Mail"
'GetRecordSet (strSQL)
'sServer = objRS("PopServer").Value
'sFromAddress = objRS("FromAddress").Value
'sSubject = objRS("Subject").Value
'sBody = objRS("Message").Value
If GetNetConnectString = "Not connected to the internet now." Then
    MsgBox "Not connected to the internet now. Try again later"
    optView.Value = False
    Exit Sub
End If
showForm frmDownload
End Sub

Private Sub MDIForm_Load()
sView = SQLDate(Date)
MonthView1.Value = SQLDate(Date)
Display
MDIForm1.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'release database object
Set myData = Nothing
'close all open child forms
i = 1
      Do While i < Forms.Count
         If Forms(i).MDIChild Then
            ' *** Do not increment i% since a form was unloaded
            Unload Forms(i)
         Else
            ' Form isn't an MDI child so go to the next form
            i = i + 1
        End If
      Loop
      ChildCount = 0
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
sView = DateClicked
Display
End Sub

Private Sub ClearAllLabels()
lblName.Caption = ""
    lblEmail.Caption = ""
    lblPurpose.Caption = ""
    lblDay.Caption = ""
    lblStart.Caption = ""
    lblEnd.Caption = ""
    txtNotes.Text = ""
    txtVenue.Text = ""
End Sub






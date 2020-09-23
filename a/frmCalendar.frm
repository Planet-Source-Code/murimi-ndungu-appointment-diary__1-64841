VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   2940
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtHoliday 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   56557569
      CurrentDate     =   37494
   End
   Begin VB.Label Label1 
      Caption         =   "Holiday"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSave_Click()
On Error GoTo ErrHandler
HolText = txtHoliday.Text
HolDate = HoliDate(MonthView1.Value)
If HolText = "" Then
    MsgBox "Enter Holiday Name", vbInformation
    Exit Sub
End If
strSQL = " INSERT INTO Holidays (Holiday, hsdate, hedate) "
strSQL = strSQL & " VALUES ('" & HolText & "','" & HolDate & "','" & HolDate & "')"
exCommand (strSQL)
Message = MsgBox("Record Succesfully Added Add Another Record", vbInformation + vbYesNo)
If Message = vbYes Then
frmAdmin.FillHolidays
    txtHoliday.Text = ""
Else
frmAdmin.FillHolidays
    Unload Me
End If
Exit Sub
ErrHandler:
    MsgBox "Error occured while adding record", vbInformation, "cmdAdd"
End Sub



Private Sub Form_Load()
MonthView1.Value = Date
End Sub

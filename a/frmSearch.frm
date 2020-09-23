VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4875
   ClientLeft      =   3240
   ClientTop       =   1965
   ClientWidth     =   6090
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame21 
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid grdSearch 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
         _Version        =   393216
         GridColor       =   16777215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cboFind 
         Height          =   315
         ItemData        =   "frmSearch.frx":0742
         Left            =   4080
         List            =   "frmSearch.frx":0752
         TabIndex        =   9
         Text            =   "Email"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cboTime 
         Height          =   315
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54984705
         CurrentDate     =   37173
      End
      Begin VB.OptionButton optEmail 
         Caption         =   "Custom Search"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Search by  Date and Time"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   6015
      Begin VB.TextBox txtSelect 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub makegrd()
On Error GoTo ErrorHandler
grdSearch.ColWidth(0) = 500
grdSearch.Cols = 9
grdSearch.Rows = 1
grdSearch.Row = 0
grdSearch.Col = 1
grdSearch.Text = "Name"
grdSearch.Col = 2
grdSearch.Text = "Email"
grdSearch.Col = 3
grdSearch.Text = "Day"
grdSearch.Col = 4
grdSearch.Text = "Starting Time"
grdSearch.Col = 5
grdSearch.Text = "Ending Time"
grdSearch.Col = 6
grdSearch.Text = "Purpose"
grdSearch.Col = 7
grdSearch.Text = "Venue"
grdSearch.Col = 8
grdSearch.Text = "Notes"

Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Public Sub FillBookedTimes()
On Error GoTo ErrorHandler
strSQL = "SELECT * FROM Book where AppDate = '" & SQLDate(DTPicker1.Value) & "'" & " ORDER BY sTime"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
    Exit Sub
Else
    objRS.MoveFirst
    Do While Not objRS.EOF
        cboTime.AddItem objRS("stime").Value
        objRS.MoveNext
    Loop
    cboTime.ListIndex = 0
End If

Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub cmdCancel_Click()
frmNew.cmdOPen.Enabled = True
vsTime = 0
veTime = 0
Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo ErrorHandler
If optDate.Value = True Then
    If cboTime.Text = "" Then
        MsgBox msgnoAppoint, vbInformation
    Else
       strSQL = "SELECT R.registrationID, R.Name, R.Email, B.* From Registration R INNER JOIN BOOK B on R.RegistrationID = B.RegID where AppDate =  '" & SQLDate(DTPicker1.Value) & "'" & " and sTime = # " & cboTime.Text & " #"
       GetRecordSet (strSQL)
       grdSearch.Rows = 1
       objRS.MoveFirst
       Do While Not objRS.EOF
        With objRS
        sNum = objRS("RegistrationID").Value
        AppNum = objRS("AppointmentNO").Value
        Gridline = vbTab & !Name & vbTab & !Email & vbTab & !appdate & vbTab & !sTime & vbTab & !eTime & vbTab
        Gridline = Gridline & !purpose & vbTab & !venue & vbTab & !notes
        grdSearch.AddItem Gridline
        .MoveNext
        End With
       Loop
       cmdSelect.Enabled = True
    End If
Else
    strSQL = "SELECT R.registrationID, R.Name, R.Email, B.* From Registration R INNER JOIN BOOK B on R.RegistrationID = B.RegID where " & cboFind.Text & " like '" & txtSearch.Text & "%'"
        GetRecordSet (strSQL)
    If objRS.EOF = True Then
            MsgBox "No records"
            Exit Sub
        Else
    grdSearch.Rows = 1
    objRS.MoveFirst
    Do While Not objRS.EOF
     With objRS
     sNum = objRS("RegistrationID").Value
     Gridline = vbTab & !Name & vbTab & !Email & vbTab & !appdate & vbTab & !sTime & vbTab & !eTime & vbTab
     Gridline = Gridline & !purpose & vbTab & !venue & vbTab & !notes
     grdSearch.AddItem Gridline
     .MoveNext
     End With
    Loop
    cmdSelect.Enabled = True
End If
End If

Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub cmdSelect_Click()
On Error GoTo ErrorHandler
grdSearch.Col = 1
frmNew.lblName.Caption = grdSearch.Text
grdSearch.Col = 2
frmNew.lblEmail.Caption = grdSearch.Text
grdSearch.Col = 3
frmNew.DTPicker1.Value = grdSearch.Text
grdSearch.Col = 4
'frmNew.cboStart.Text = grdSearch.Text
vsTime = grdSearch.Text
grdSearch.Col = 5
'frmNew.cboEnd.Text = grdSearch.Text
veTime = grdSearch.Text
grdSearch.Col = 6
frmNew.cboPurpose.Text = grdSearch.Text
grdSearch.Col = 7
frmNew.txtVenue.Text = grdSearch.Text
grdSearch.Col = 8
frmNew.txtNotes.Text = grdSearch.Text
Unload Me
'show appropriate form
Select Case frmNew.Caption
 Case "New Schedule"
    frmNew.Show
 Case "Re-Schedule"
    frmNew.Caption = "Re-Schedule"
    frmNew.cmdNew.Enabled = False
    showForm frmNew
 Case "Delete"
    frmNew.Caption = "Delete"
    frmNew.cmdNew.Enabled = False
    showForm frmNew
End Select
frmNew.cmdOPen.Enabled = True
Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub DTPicker1_Change()
'first clear time combo
cboTime.Clear
FillBookedTimes
End Sub

Private Sub Form_Load()
optDate.Value = True
DTPicker1.Value = Date
FillBookedTimes
makegrd
End Sub

Private Sub grdSearch_Click()
grdSearch.Col = 2
txtSelect.Text = grdSearch.Text
End Sub

Private Sub grdSearch_DblClick()
cmdSelect_Click
End Sub

Private Sub txtSearch_Change()
cmdSearch_Click
End Sub

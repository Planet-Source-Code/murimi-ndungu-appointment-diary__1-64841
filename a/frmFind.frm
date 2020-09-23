VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Records"
   ClientHeight    =   5640
   ClientLeft      =   2985
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6105
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   6015
      Begin VB.TextBox txtSelect 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdFind 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   16777215
      FocusRect       =   2
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   160
         Width           =   2535
      End
      Begin VB.ComboBox cboFind 
         Height          =   315
         ItemData        =   "frmFind.frx":0742
         Left            =   4080
         List            =   "frmFind.frx":074F
         TabIndex        =   8
         Text            =   "Name"
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "By"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gridline As String
Dim strSQL As String

Private Sub cboFind_Change()
'If cboFind.Text = "RegistrationID" Then
'
'    strSQL ="select * from Registration where " & cboFind.Text & " like " & txtSearch.Text & "%"
'Else
'
'    strSQL ="select * from Registration where " & cboFind.Text & " like '" & txtSearch.Text & "%'"
End Sub

Private Sub cmdCancel_Click()
frmNew.cmdOPen.Enabled = True

Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo ErrorHandler
strSQL = "select * from Registration where " & cboFind.Text & " like '" & txtSearch.Text & "%'"
GetRecordSet (strSQL)
If objRS.EOF = True Then
 MsgBox "No records, First add records by clicking on new button."
 Unload Me
Else
grdFind.Rows = 1
 objRS.MoveFirst
  Do While Not objRS.EOF
    With objRS
    
      Gridline = vbTab & !RegistrationID & vbTab & !Name & vbTab & !Email & vbTab & !Address & vbTab
      Gridline = Gridline & !city & vbTab & !Country
    
      grdFind.AddItem Gridline
      .MoveNext
    End With
    Loop
cmdSelect.Enabled = True
End If
Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Public Sub makegrd()
On Error GoTo ErrorHandler
grdFind.ColWidth(0) = 500
grdFind.Cols = 7
grdFind.Rows = 1
grdFind.Row = 0
grdFind.Col = 1
grdFind.Text = "Registration ID"
grdFind.Col = 2
grdFind.Text = "Name"
grdFind.Col = 3
grdFind.Text = "Email"
grdFind.Col = 4
grdFind.Text = "Address"
grdFind.Col = 5
grdFind.Text = "City"
grdFind.Col = 6
grdFind.Text = "Country"
Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub cmdSelect_Click()
On Error GoTo ErrorHandler
grdFind.Col = 1
sNum = grdFind.Text
grdFind.Col = 2
sName = grdFind.Text
grdFind.Col = 3
sEmail = grdFind.Text
Unload Me
frmNew.lblName.Caption = sName
frmNew.lblEmail.Caption = sEmail
frmNew.Show
Exit Sub
frmNew.cmdOPen.Enabled = True
ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub Form_Load()
makegrd
End Sub

Private Sub grdFind_Click()
On Error GoTo ErrorHandler
grdFind.Col = 2
txtSelect.Text = grdFind.Text
Exit Sub

ErrorHandler:

   MsgBox (Err.Description)
End Sub

Private Sub grdFind_DblClick()
cmdSelect_Click
End Sub

Private Sub txtSearch_Change()
cmdSearch_Click
End Sub

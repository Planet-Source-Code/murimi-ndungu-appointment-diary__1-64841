VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Schedule"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   FillStyle       =   0  'Solid
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5850
   Begin MSFlexGridLib.MSFlexGrid grdGrid 
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Details"
      Height          =   2175
      Left            =   3120
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   22806529
      CurrentDate     =   37135
   End
   Begin VB.Label Label15 
      Caption         =   "Click on a date to view appointments for that day"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblDisplay 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub Form_Load()
lblDisplay.Caption = "Appointments for " & WeekdayName(Weekday(Date)) & " :" & Date
Call SizeCells
Call CenterCells
strSQL = "SELECT * FROM Book where AppDate=#" & SQLDate(Date) & "#" & " ORDER BY sTime" & ";"
Set objRS = New ADODB.Recordset
objRS.Open strSQL, objConn, adOpenDynamic, adLockOptimistic

With grdGrid
    .Rows = 0
    
    While Not objRS.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .CellBackColor = RGB(255, 255, 100)
        .Text = SQLTime(objRS("sTime")) & " to " & SQLTime(objRS("eTime")) & ""
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Text = objRS("Ailment") & ""
                              
    objRS.MoveNext
    Wend
    
End With



End Sub



Private Sub MonthsView1_DateClick(ByVal DateClicked As Date)

strSQL = "SELECT * FROM Book where AppDate=#" & SQLDate(DateClicked) & "#" & " ORDER BY sTime" & ";"
Set objRS = New ADODB.Recordset
objRS.Open strSQL, objConn, adOpenDynamic, adLockOptimistic

With grdGrid
    '.Rows
    
    While Not objRS.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .CellBackColor = RGB(255, 255, 100)
        .Text = SQLTime(objRS("sTime")) & " to " & SQLTime(objRS("eTime")) & ""
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Text = objRS("Ailment") & ""
                              
    objRS.MoveNext
    Wend
End With

lblDisplay.Caption = "Appointments for " & DateClicked

End Sub

Private Sub SizeCells()
    Dim intColumn As Integer
    
    grdGrid.ColWidth(0) = 1100
    
    For intColumn = 0 To 0
        grdGrid.ColWidth(intColumn) = 2200
    Next intColumn
    
End Sub

Private Sub CenterCells()
    Dim intColumn As Integer
    
    For intColumn = 0 To 0
        grdGrid.ColAlignment(intColumn) = flexAlignCenterCenter
    Next intColumn
    
End Sub

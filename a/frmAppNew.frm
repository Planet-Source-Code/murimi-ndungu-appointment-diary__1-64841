VERSION 5.00
Begin VB.Form frmAppNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Form"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmAppNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Welcome To Appointment Diary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAppNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
On Error GoTo ErrHandler
Dim SerialNum As String
Dim Appstatus As String
SerialNum = "{" & txtNum.Text & "}"
strSQL = "Select Registration From Owner"
GetRecordSet (strSQL)
'If SerialNum = objRS("Registration").Value Then
If Trim(SerialNum) = "{}" Then
Appstatus = "No"
'update registration
strSQL = "UPDATE Owner SET AppNew = '" & Appstatus & "'"
exCommand (strSQL)
    Message = MsgBox("Thank You for purchasing Appointment Diary.", vbInformation)
    Call ShellExecute(Me.hwnd, "open", "http://www.refpoint.antunit.com/", 0, 0, vbNormalFocus)
    Unload Me
    frmLogin.Show
Else
    Message = MsgBox("Wrong serial number, contact vendor for details.", vbInformation)
    Call ShellExecute(Me.hwnd, "open", "http://www.refpoint.antunit.com/", 0, 0, vbNormalFocus)
    Unload Me
End If
Exit Sub
ErrHandler:
    ' Describe the error to the user.
    MsgBox "Unexpected error" & _
        vbCrLf & _
        Err.Description
    Exit Sub
End Sub

Private Sub Label2_Click()
On Error GoTo HandleErrors
Call ShellExecute(Me.hwnd, "open", "http://www.refpoint.antunit.com/", 0, 0, vbNormalFocus)
Exit Sub
  
HandleErrors:

  MsgBox Err.Description, vbCritical, App.Title & " Error"
End Sub

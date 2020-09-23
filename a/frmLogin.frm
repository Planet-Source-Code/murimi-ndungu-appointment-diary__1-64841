VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Login"
   ClientHeight    =   1950
   ClientLeft      =   4815
   ClientTop       =   7455
   ClientWidth     =   4095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1152.126
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.984
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "This program requires your display settings be set to 800 by 600 Pixels for optimum display."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsPressed As Boolean
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
      
    Unload Me
    IsPressed = Not IsPressed
    LoginSucceeded = False
End Sub
Private Sub cmdOK_Click()
    
    Dim UserExist As Boolean
    strSQL = "Select Name, Password from owner"
    GetRecordSet (strSQL)
    objRS.MoveFirst
    UserExist = False
    
    Do
    
        If (txtUserName.Text = objRS("Name").Value) Then
            CurrentUser = objRS("Name").Value
            UserPassword = Decrypt(objRS("Password").Value)
            UserExist = True
            Exit Do
        End If
        objRS.MoveNext
    Loop While Not objRS.EOF
    
    If (UserExist = False) Then
        MsgBox "Name " & txtUserName.Text & " is not a registered user. Pls. enter a valid user...", vbExclamation
        Exit Sub
    End If
    
    If (CStr(txtPassword) <> UserPassword) Then
        MsgBox "Wrong password for Name: " & txtUserName.Text, vbExclamation
        Exit Sub
    End If
    IsPressed = True
    LoginSucceeded = True
    Unload Me
    MDIForm1.Show
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer) 'Check if enter

    If KeyAscii = 13 Then cmdOK_Click

End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then txtPassword.SetFocus

End Sub

Public Function IsUserOK() As Boolean
      
    LoginSucceeded = False
    Load frmLogin
    IsPressed = False
    frmLogin.Show
    While Not IsPressed
        DoEvents
    Wend
    IsUserOK = LoginSucceeded
    
End Function


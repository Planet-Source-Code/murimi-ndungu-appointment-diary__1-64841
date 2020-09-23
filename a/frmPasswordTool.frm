VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPasswordTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Database"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmPasswordTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tbrPass 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1111
      ButtonWidth     =   2778
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "imlPass"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&New User"
            Key             =   "New"
            Object.ToolTipText     =   "Adds a new user"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Edit User"
            Key             =   "Edit"
            Object.ToolTipText     =   "Edits the selected user"
            ImageKey        =   "edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Delete User"
            Key             =   "Delete"
            Object.ToolTipText     =   "Deletes the selected user"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save &Info"
            Key             =   "Save"
            Object.ToolTipText     =   "Saves the current info"
            ImageKey        =   "save"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboUser 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame fraPass 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   5775
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtCurrent 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtVerify 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblUser 
         Caption         =   "&Username :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCurrent 
         Caption         =   "&Current Password :"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblNew 
         Caption         =   "N&ew Password :"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblVerify 
         Caption         =   "&Verify Password :"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imlPass 
      Left            =   5280
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPasswordTool.frx":030A
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPasswordTool.frx":0626
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPasswordTool.frx":0A7A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPasswordTool.frx":1356
            Key             =   "edit"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSelect 
      Caption         =   "&Select User :"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmPasswordTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveCurrent As Boolean
Dim IsEditing As Boolean

Private Sub cboUser_Click()

On Error GoTo ErrHandler
    
UserRS.MoveFirst
Do
    If (cboUser.Text = UserRS("Username").Value) Then
        Exit Do
    End If
    UserRS.MoveNext
Loop While Not UserRS.EOF
Exit Sub

ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cboUser_KeyPress(KeyAscii As Integer)

    MsgBox "Pls. choose form the specified selections only...", vbInformation
    Call Form_Load

End Sub

Private Sub cmdDone_Click()

    If IsEditing = False Then
        Unload Me
    Else
        txtUser.Text = ""
        txtCurrent.Text = ""
        txtNew.Text = ""
        txtVerify.Text = ""
        cboUser.Clear
        Call Form_Load
        tbrPass.Buttons(4).Enabled = False
    End If
    
End Sub

Private Sub Form_Load() 'Load form
    
On Error GoTo ErrHandler
    
IsEditing = False
SaveCurrent = True
If UserRS.RecordCount = 0 Then
    tbrPass.Buttons(1).Enabled = True
    tbrPass.Buttons(2).Enabled = False
    tbrPass.Buttons(3).Enabled = False
    tbrPass.Buttons(4).Enabled = False
    cboUser.Enabled = False
    fraPass.Enabled = False
    Exit Sub
End If
UserRS.MoveFirst
cboUser.Enabled = True
Do
    cboUser.AddItem UserRS("Username").Value
    UserRS.MoveNext
Loop While Not UserRS.EOF
cboUser.ListIndex = 0
tbrPass.Buttons(1).Enabled = True
tbrPass.Buttons(2).Enabled = True
tbrPass.Buttons(3).Enabled = True
fraPass.Enabled = False
frmMDIVideoDex.sbrStatus.Panels(1).Text = "Classified database. Top secret"
frmMDIVideoDex.sbrStatus.Panels(2).Text = "Classified records. Top secret"
cmdDone.Caption = "&Done"
Exit Sub

ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Unload form

    frmMDIVideoDex.sbrStatus.Panels(1).Text = "No database open"
    frmMDIVideoDex.sbrStatus.Panels(2).Text = " 0 records"

End Sub

'Toolbar events
Private Sub tbrPass_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            Call NewUser
        Case "Edit"
            Call EditUser
        Case "Delete"
            Call DeleteUser
        Case "Save"
            Call SaveUser
    End Select

End Sub
'Prepare the textboxes for input of new user
Private Sub NewUser()

    IsEditing = True
    SaveCurrent = True
    txtUser.Enabled = True
    cboUser.Enabled = False
    tbrPass.Buttons(4).Enabled = True
    tbrPass.Buttons(3).Enabled = False
    tbrPass.Buttons(2).Enabled = False
    tbrPass.Buttons(1).Enabled = False
    fraPass.Enabled = True
    txtNew.Text = ""
    txtVerify.Text = ""
    txtNew.Enabled = False
    txtVerify.Enabled = False
    txtUser.SetFocus
    cmdDone.Caption = "&Cancel"
    
End Sub
'Prepare for editing the user selected
Private Sub EditUser()

    IsEditing = True
    SaveCurrent = False
    txtUser.Enabled = False
    cboUser.Enabled = False
    txtUser.Text = cboUser.Text
    tbrPass.Buttons(4).Enabled = True
    tbrPass.Buttons(3).Enabled = False
    tbrPass.Buttons(2).Enabled = False
    tbrPass.Buttons(1).Enabled = False
    fraPass.Enabled = True
    txtNew.Text = ""
    txtVerify.Text = ""
    txtNew.Enabled = True
    txtVerify.Enabled = True
    txtCurrent.SetFocus
    cmdDone.Caption = "&Cancel"
    
End Sub
'Delete the user
Private Sub DeleteUser()

    Dim Answer
    Dim X As Integer
    
    If frmLogin.IsUserOK = False Then
        MsgBox "Access Denied !!", vbCritical
        Exit Sub
    End If
    If cboUser.Text = CurrentUser Then
        MsgBox "You can't delete the currently logged user.", vbInformation
        Exit Sub
    End If
    Answer = MsgBox("Are you sure you want to delete the user " & cboUser.Text & " ?", vbYesNo)
    If Answer = vbYes Then
        UserRS.MoveFirst
        Do
            If (cboUser.Text = UserRS("Username").Value) Then
                UserRS.Delete
                Exit Do
            End If
            UserRS.MoveNext
        Loop While Not UserRS.EOF
        cboUser.Clear
        Call Form_Load
    End If
        
End Sub
'Save info of new user
Private Sub SaveUser()

On Error GoTo ErrHandler
    
    Dim X As Integer
    If frmLogin.IsUserOK = False Then
        MsgBox "Access Denied !!", vbCritical
        Exit Sub
    End If
    If Len(txtUser.Text) = 0 Then
        MsgBox "Username field must not be empty...", vbInformation
        txtUser.SetFocus
        Exit Sub
    ElseIf Len(txtCurrent.Text) = 0 Then
        MsgBox "Password field must not be empty...", vbInformation
        txtCurrent.SetFocus
        Exit Sub
    ElseIf (Len(txtNew.Text) = 0) And (SaveCurrent = False) Then
        MsgBox "Pls. type a new password...", vbInformation
        txtNew.SetFocus
        Exit Sub
    ElseIf (Len(txtVerify.Text) = 0) And (SaveCurrent = False) Then
        MsgBox "Pls. verify your new password...", vbInformation
        txtVerfiy.SetFocus
        Exit Sub
    End If
    If UserRS.RecordCount <> 0 Then
        If ((CStr(txtCurrent) <> Decrypt(UserRS("Password").Value)) And (SaveCurrent = False)) Then
            MsgBox "Wrong password. Pls. try again...", vbInformation
            Exit Sub
        End If
    End If
    If ((txtVerify <> txtNew.Text) And (SaveCurrent = False)) Then
        MsgBox "New password and verification didn't match. Pls. try again...", vbInformation
        Exit Sub
    End If
    If SaveCurrent = True Then
        UserRS.AddNew
    Else
        UserRS.Edit
    End If
    UserRS("Username").Value = txtUser.Text
    If SaveCurrent = True Then
        UserRS("Password").Value = Encrypt(txtCurrent)
        UserRS.Update
        MsgBox "User : " & txtUser.Text & " has been successfully added to the user database...", vbInformation
    Else
        UserRS("Password").Value = Encrypt(txtVerify)
        If txtUser.Text = CurrentUser Then UserPassword = CStr(txtVerify)
        UserRS.Update
        MsgBox "User : " & txtUser.Text & " has been successfully updated in the user database...", vbInformation
    End If
    txtUser.Text = ""
    txtCurrent.Text = ""
    txtNew.Text = ""
    txtVerify.Text = ""
    cboUser.Clear
    Call Form_Load
    tbrPass.Buttons(4).Enabled = False
    Exit Sub

ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub


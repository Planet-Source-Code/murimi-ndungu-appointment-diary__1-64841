VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration"
   ClientHeight    =   5190
   ClientLeft      =   2715
   ClientTop       =   -210
   ClientWidth     =   6645
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6645
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmAdmin.frx":0742
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPurpose"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "User Information"
      TabPicture(1)   =   "frmAdmin.frx":075E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSelect"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "imlPass"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tbrPass"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDone"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraPass"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboUser"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Holidays"
      TabPicture(2)   =   "frmAdmin.frx":077A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Others"
      TabPicture(3)   =   "frmAdmin.frx":0796
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   6375
         Begin VB.CommandButton Command5 
            Caption         =   "Save"
            Height          =   495
            Left            =   4440
            TabIndex        =   47
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Delete"
            Height          =   495
            Left            =   3120
            TabIndex        =   46
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Update"
            Height          =   495
            Left            =   1800
            TabIndex        =   45
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton cmdADD 
            Caption         =   "Add"
            Height          =   495
            Left            =   480
            TabIndex        =   44
            Top             =   3960
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid grdHoliday 
            Height          =   3495
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   6165
            _Version        =   393216
            RowHeightMin    =   300
            FocusRect       =   2
            GridLines       =   3
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   3
         End
      End
      Begin VB.ComboBox cboUser 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71280
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   900
         Width           =   2175
      End
      Begin VB.Frame fraPass 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   -73200
         TabIndex        =   22
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txtUser 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   26
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtCurrent 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   27
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtNew 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtVerify 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   29
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label lblUser 
            Caption         =   "&Username :"
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblCurrent 
            Caption         =   "&Current Password :"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblNew 
            Caption         =   "N&ew Password :"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblVerify 
            Caption         =   "&Verify Password :"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "&Done"
         Height          =   495
         Left            =   -74880
         TabIndex        =   21
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Frame fraPurpose 
         Caption         =   "Setup Purpose"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   19
         Top             =   2760
         Width           =   6375
         Begin VB.ComboBox cboPurpose 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Done"
            Height          =   495
            Left            =   4800
            TabIndex        =   38
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox txtPurpose 
            Height          =   285
            Left            =   1920
            TabIndex        =   10
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtpDuration 
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin MSComctlLib.Toolbar tbrPurpose 
            Height          =   630
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   1111
            ButtonWidth     =   2514
            ButtonHeight    =   1005
            TextAlignment   =   1
            ImageList       =   "imlPass"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "&New"
                  Key             =   "New"
                  Object.ToolTipText     =   "Adds a new user"
                  ImageKey        =   "new"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "&Edit"
                  Key             =   "Edit"
                  Object.ToolTipText     =   "Edits the selected user"
                  ImageKey        =   "beta"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "&Delete"
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
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   4920
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdmin.frx":07B2
                  Key             =   "new"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdmin.frx":0ACE
                  Key             =   "save"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdmin.frx":0F22
                  Key             =   "delete"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdmin.frx":17FE
                  Key             =   "exit"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdmin.frx":1C52
                  Key             =   "beta"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Caption         =   "Existing"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Duration ( in minutes)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Purpose"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Diary Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   6375
         Begin VB.TextBox txtUsers 
            Height          =   285
            Left            =   2280
            TabIndex        =   41
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox txtoDuration 
            Height          =   285
            Left            =   2280
            TabIndex        =   7
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CommandButton cmdChangecolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Caption         =   "Change"
            Height          =   255
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   2055
         End
         Begin VB.Frame Frame4 
            Height          =   2175
            Left            =   4920
            TabIndex        =   35
            Top             =   0
            Width           =   1455
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   360
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save"
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   1680
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDuration 
            Height          =   285
            Left            =   2280
            TabIndex        =   6
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox cbobStart 
            Height          =   315
            Left            =   3360
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboaEnd 
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cbobEnd 
            Height          =   315
            Left            =   3360
            TabIndex        =   4
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cboaStart 
            Height          =   315
            Left            =   2280
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Maximum number of users"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Hrs"
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
            Index           =   1
            Left            =   4440
            TabIndex        =   37
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Hrs"
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
            Index           =   0
            Left            =   4440
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Opening Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Closing Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Grid Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Duration ( in minutes)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Other Duration(in Minutes)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   2295
         End
      End
      Begin MSComctlLib.Toolbar tbrPass 
         Height          =   2340
         Left            =   -74880
         TabIndex        =   32
         Top             =   960
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   4128
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
               ImageKey        =   "beta"
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
      Begin MSComctlLib.ImageList imlPass 
         Left            =   -74880
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdmin.frx":1F6C
               Key             =   "new"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdmin.frx":2288
               Key             =   "save"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdmin.frx":26DC
               Key             =   "delete"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdmin.frx":2FB8
               Key             =   "exit"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdmin.frx":340C
               Key             =   "beta"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Other Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   6375
         Begin VB.CheckBox chkHol 
            Caption         =   "Closed on Holidays"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   2055
         End
         Begin VB.CheckBox chkSun 
            Caption         =   "Closed on Sundays"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox chkSat 
            Caption         =   "Closed on Saturdays"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdSaveOther 
            Caption         =   "Save"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   3960
            Width           =   1695
         End
         Begin VB.CheckBox chkDel 
            Caption         =   "Delete old Appointmnens"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chkBackup 
            Caption         =   "BackUp old appointments"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Label lblSelect 
         Caption         =   "&Select User :"
         Height          =   255
         Left            =   -72720
         TabIndex        =   33
         Top             =   1020
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveCurrent As Boolean
Dim IsEditing As Boolean
Dim PID As Integer
Dim isNumber As Boolean

Private Sub cboPurpose_Click()
On Error GoTo ErrHandler
strSQL = "Select PurposeID, Purpose, pTime from Purpose where purpose = '" & cboPurpose.Text & "'"
GetRecordSet (strSQL)
PID = objRS("PurposeID").Value
txtPurpose.Text = objRS("Purpose").Value
txtpDuration.Text = objRS("pTime").Value
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cboUser_Click()

On Error GoTo ErrHandler
    
objRS.MoveFirst
Do
    If (cboUser.Text = objRS("Name").Value) Then
        Exit Do
    End If
    objRS.MoveNext
Loop While Not objRS.EOF
Exit Sub

ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cboUser_KeyPress(KeyAscii As Integer)

    MsgBox "Pls. choose form the specified selections only...", vbInformation
    Call Form_Load

End Sub

Private Sub chkBackup_Click()
MsgBox msgOld
End Sub

Private Sub chkDel_Click()
MsgBox msgOld

End Sub

Private Sub cmdADD_Click()
showForm frmCalendar
End Sub

Private Sub cmdChangecolor_Click()
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    GridAltColor = CommonDialog1.Color
    
    cmdChangecolor.BackColor = GridAltColor
    Exit Sub
    
ErrHandler:
    ' Describe the error to the user.
    MsgBox "Unexpected error" & _
       " When altering color." & _
        vbCrLf & _
        Err.Description
    Exit Sub
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
        tbrPurpose.Buttons(4).Enabled = False
    End If
    
End Sub

Private Sub cmdSave_Click()
' Install the error handler.
On Error GoTo ErrHandler

Dim sHr As Integer
Dim sMin As Integer
Dim eHr As Integer
Dim eMin As Integer
Dim Dur As Integer
Dim oDur As Integer
Dim Starting As Integer
Dim Ending As Integer

sHr = CInt(cboaStart.Text)
sMin = CInt(cbobStart.Text)
eHr = CInt(cboaEnd.Text)
eMin = CInt(cbobEnd.Text)
Dur = CInt(txtDuration.Text)
oDur = CInt(txtoDuration.Text)

If sMin = 0 Then
    Starting = sHr * 100
ElseIf sMin = 5 Then
    Starting = sHr * 100 + 5
Else
    Starting = sHr & sMin
End If

If eMin = 0 Then
    Ending = eHr * 100
ElseIf eMin = 5 Then
    Ending = eHr * 100 + 5
Else
    Ending = eHr & eMin
End If


If Starting >= Ending Then
    Ending = Ending + 2400
End If

'validate entries
If txtDuration.Text < 1 Or txtDuration.Text > 1440 Then
    Message = MsgBox("Enter valid minutes in the Duration textbox.", vbExclamation)
    txtDuration.SetFocus
    Exit Sub
End If
If txtoDuration.Text < 1 Or txtoDuration.Text > 1440 Then
    Message = MsgBox("Enter valid minutes in the Other Duration textbox.", vbExclamation)
    txtoDuration.SetFocus
    Exit Sub
End If

strSQL = "UPDATE Owner SET color = " & GridAltColor & ",sTime = " & Starting & ",eTime = " & Ending & ",hs = '" & cboaStart.Text & "',ms = '" & cbobStart.Text & "',he = '" & cboaEnd.Text & "',me = '" & cbobEnd.Text & "',Duration = " & Dur & ",oDuration = " & oDur
exCommand (strSQL)
strSQL = "UPDATE Purpose SET pTime = " & Dur & " Where PurposeID = 1"
exCommand (strSQL)
MsgBox msgSaved
cboPurpose_Click
Display
Exit Sub

ErrHandler:
    ' Describe the error to the user.
    MsgBox "Unexpected error" & _
       " When saving records." & _
        vbCrLf & _
        Err.Description
    Exit Sub

End Sub

Private Sub Command2_Click()
MsgBox msgOld
End Sub

Private Sub cmdSaveOther_Click()
On Error GoTo ErrHandler
DelOldApp = 0
BackupOldApp = 0
ClosedSat = chkSat.Value
ClosedSun = chkSun.Value
ClosedHol = chkHol.Value
strSQL = "UPDATE Owner SET OldDel = '" & DelOldApp & "',OldBack = '" & BackupOldApp & "',cSat = '" & ClosedSat & "',cSun = '" & ClosedSun & "',cHol = '" & ClosedHol & "'"
exCommand (strSQL)
MsgBox ("Record saved Successfully")
Exit Sub
ErrHandler:
    ' Describe the error to the user.
    MsgBox "Unexpected error" & _
       " When saving records." & _
        vbCrLf & _
        Err.Description
    Exit Sub
End Sub

Private Sub Command3_Click()
MsgBox msgOld
End Sub

Private Sub Command4_Click()
MsgBox msgOld
End Sub

Private Sub Command5_Click()
MsgBox msgOld
End Sub

Private Sub Form_Load() 'Load form
On Error GoTo ErrHandler
SSTab1.Tab = 0
cmdChangecolor.BackColor = GridAltColor
FillOwner
FillHours
FillMinutes
FillPurpose
FillHolidays
txtPurpose.Enabled = False
txtpDuration.Enabled = False
FillUser
Exit Sub
ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub
    

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.optSetup.Value = False
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
    MsgBox msgOld
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
MsgBox msgOld
End Sub
'Save info of new user
Private Sub SaveUser()

On Error GoTo ErrHandler

   If Len(txtUser.Text) = 0 Then
        MsgBox "Name field must not be empty...", vbInformation
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
    If objRS.RecordCount <> 0 Then
        If ((CStr(txtCurrent) <> Decrypt(objRS("Password").Value)) And (SaveCurrent = False)) Then
            MsgBox "Wrong password. Pls. try again...", vbInformation
            Exit Sub
        End If
    End If
    If ((txtVerify <> txtNew.Text) And (SaveCurrent = False)) Then
        MsgBox "New password and verification didn't match. Pls. try again...", vbInformation
        Exit Sub
    End If
    If SaveCurrent = True Then
        objRS.AddNew
    Else
        objRS.Update
    End If
    objRS("Name").Value = txtUser.Text
    If SaveCurrent = True Then
        objRS("Password").Value = Encrypt(txtCurrent)
        objRS.Update
        MsgBox "User : " & txtUser.Text & " has been successfully added to the user database...", vbInformation
    Else
        objRS("Password").Value = Encrypt(txtVerify)
        If txtUser.Text = CurrentUser Then UserPassword = CStr(txtVerify)
        objRS.Update
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

Private Sub FillHours()
Dim Hour As Integer
Dim strHour As String
For Hour = 0 To 23 Step 1
If Len(Str(Hour)) = 2 Then
   strHour = "0" & Hour
Else
    strHour = Hour
End If
 cboaStart.AddItem strHour
 cboaEnd.AddItem strHour
Next
End Sub

Public Sub FillMinutes()
Dim Minute As Integer
Dim strMinute As String
For Minute = 0 To 55 Step 5
If Len(Str(Minute)) = 2 Then
   strMinute = "0" & Minute
Else
    strMinute = Minute
End If
    cbobStart.AddItem strMinute
    cbobEnd.AddItem strMinute
Next
End Sub

Private Sub tbrPurpose_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            Call NewPurpose
        Case "Edit"
            Call EditPurpose
        Case "Delete"
            Call DeletePurpose
        Case "Save"
            Call SavePurpose
    End Select
End Sub
Private Sub NewPurpose()

    IsEditing = True
    SaveCurrent = True
    txtPurpose.Enabled = True
    txtpDuration.Enabled = True
    cboPurpose.Enabled = False
    tbrPurpose.Buttons(4).Enabled = True
    tbrPurpose.Buttons(3).Enabled = False
    tbrPurpose.Buttons(2).Enabled = False
    tbrPurpose.Buttons(1).Enabled = False
    txtPurpose.Text = ""
    txtpDuration.Text = ""
    txtPurpose.SetFocus
        
End Sub

'Prepare for editing the user selected
Private Sub EditPurpose()

    IsEditing = True
    SaveCurrent = False
    tbrPurpose.Buttons(4).Enabled = True
    tbrPurpose.Buttons(3).Enabled = False
    tbrPurpose.Buttons(2).Enabled = False
    tbrPurpose.Buttons(1).Enabled = False
    txtPurpose.Enabled = True
    txtpDuration.Enabled = True
    txtPurpose.SetFocus
       
End Sub
'Delete the user
Private Sub DeletePurpose()

    If PID = 1 Then
        MsgBox "You can't Delete the Default Record.", vbInformation
        Exit Sub
    End If
    
    Message = MsgBox("Are you sure you want to delete the Purpose " & cboPurpose.Text & " ?", vbYesNo)
    If Message = vbYes Then
        strSQL = "Delete * from purpose where purposeid = " & PID
        exCommand (strSQL)
        cboPurpose.Clear
        FillPurpose
    End If
        
End Sub
'Save info of new user
Private Sub SavePurpose()

On Error GoTo ErrHandler

    If Len(txtPurpose.Text) = 0 Then
        MsgBox "Purpose field must not be empty...", vbInformation
        txtUser.SetFocus
        Exit Sub
    ElseIf Len(txtpDuration.Text) = 0 Then
        MsgBox "Duration field must not be empty...", vbInformation
        txtCurrent.SetFocus
        Exit Sub
    End If
    'validate entries
    If txtpDuration.Text < 1 Or txtpDuration.Text > 1440 Then
        Message = MsgBox("Enter valid minutes in the Duration textbox.", vbExclamation)
        txtpDuration.SetFocus
        Exit Sub
    End If
    If SaveCurrent = True Then
        strSQL = " INSERT INTO Purpose (Purpose, pTime) "
        strSQL = strSQL & " VALUES ('" & txtPurpose.Text & "'," & CInt(txtpDuration.Text) & ")"
        exCommand (strSQL)
        MsgBox "Purpose : " & txtPurpose.Text & " has been successfully added", vbInformation
    Else
        strSQL = "UPDATE Purpose SET pTime = " & CInt(txtpDuration.Text) & " ,Purpose = '" & txtPurpose.Text & "' Where PurposeID = " & PID
        exCommand (strSQL)
        If PID = 1 Then
            strSQL = "UPDATE Owner SET Duration = " & CInt(txtpDuration.Text) & ""
            exCommand (strSQL)
        End If
        MsgBox "PURPOSE : " & txtPurpose.Text & " has been successfully updated", vbInformation
    End If
    Call FillPurpose
    txtPurpose.Enabled = False
    txtpDuration.Enabled = False
    tbrPurpose.Buttons(4).Enabled = False
    Exit Sub

ErrHandler:
    MsgBox "Error :" & " " & Err.Description, vbCritical
    Exit Sub
End Sub



Private Sub FillPurpose()

'set the purposes combo
bolEditReadRS = False
strSQL = "Select Purpose, pTime from Purpose"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
    tbrPurpose.Buttons(1).Enabled = True
    tbrPurpose.Buttons(2).Enabled = False
    tbrPurpose.Buttons(3).Enabled = False
    tbrPurpose.Buttons(4).Enabled = False
    cboPurpose.Enabled = False
    Exit Sub
End If
objRS.MoveFirst
cboPurpose.Enabled = True
cboPurpose.Clear
Do
    cboPurpose.AddItem objRS("Purpose").Value
    objRS.MoveNext
Loop While Not objRS.EOF
cboPurpose.ListIndex = 0
tbrPurpose.Buttons(1).Enabled = True
tbrPurpose.Buttons(2).Enabled = True
tbrPurpose.Buttons(3).Enabled = True
End Sub

Public Sub FillUser()
bolEditReadRS = True
strSQL = "Select Name, Password from owner"
GetRecordSet (strSQL)
IsEditing = False
SaveCurrent = True
If objRS.RecordCount = 0 Then
    tbrPass.Buttons(1).Enabled = True
    tbrPass.Buttons(2).Enabled = False
    tbrPass.Buttons(3).Enabled = False
    tbrPass.Buttons(4).Enabled = False
    cboUser.Enabled = False
    fraPass.Enabled = False
    Exit Sub
End If
objRS.MoveFirst
cboUser.Enabled = True
Do
    cboUser.AddItem objRS("Name").Value
    objRS.MoveNext
Loop While Not objRS.EOF
cboUser.ListIndex = 0
tbrPass.Buttons(1).Enabled = True
tbrPass.Buttons(2).Enabled = True
tbrPass.Buttons(3).Enabled = True
fraPass.Enabled = False
cmdDone.Caption = "&Done"

End Sub

Private Sub FillOwner()
strSQL = "Select  Owner.hs, Owner.ms, Owner.he, Owner.me, Owner.Duration, Owner.oDuration, Owner.OldDel, Owner.OldBack, Owner.cSat, Owner.cSun, Owner.cHol FROM Owner"
GetRecordSet (strSQL)
IsEditing = False
SaveCurrent = True

If objRS.RecordCount = 0 Then
    Exit Sub
End If
objRS.MoveFirst
cboaStart.Text = objRS("hs").Value
cbobStart.Text = objRS("ms").Value
cboaEnd.Text = objRS("he").Value
cbobEnd.Text = objRS("me").Value
txtDuration.Text = objRS("Duration").Value
txtoDuration.Text = objRS("oDuration").Value
chkDel.Value = objRS("OldDel").Value
chkBackup.Value = objRS("OldBack").Value
chkSat.Value = objRS("cSat").Value
chkSun.Value = objRS("cSun").Value
chkHol.Value = objRS("cHol").Value
End Sub

Private Sub txtDuration_KeyPress(KeyAscii As Integer)
'check for invalid characters
Select Case KeyAscii
 Case 48 To 57    'numbers
 Case 8, 45, 32, 13, 9 'special characters
 Case Else
 MsgBox msgNumbers
 KeyAscii = 0
 txtDuration.SetFocus
 Exit Sub
 End Select
End Sub

Public Sub FillHolidays()
Dim Gridline As String
grdHoliday.ColWidth(0) = 500
grdHoliday.Cols = 3
grdHoliday.Rows = 1
grdHoliday.Row = 0
grdHoliday.ColWidth(1) = 4000
grdHoliday.Col = 1
grdHoliday.Text = "Holiday"
grdHoliday.ColWidth(2) = 1500
grdHoliday.Col = 2
grdHoliday.Text = "Date"


strSQL = "Select Holiday, hsDate from Holidays"
GetRecordSet (strSQL)
If objRS.RecordCount = 0 Then
    Exit Sub
End If
objRS.MoveFirst
Do While Not objRS.EOF
    With objRS

      Gridline = vbTab & !Holiday & vbTab & !hsDate
     grdHoliday.AddItem Gridline
      .MoveNext
    End With
    Loop
End Sub

Private Sub ValidNum()

End Sub


Private Sub txtoDuration_KeyPress(KeyAscii As Integer)
'check for invalid characters
Select Case KeyAscii
 Case 48 To 57    'numbers
 Case 8, 45, 32, 13, 9 'special characters
 Case Else
 MsgBox msgNumbers
 KeyAscii = 0
 txtDuration.SetFocus
 Exit Sub
 End Select
End Sub


Private Sub txtpDuration_KeyPress(KeyAscii As Integer)
'check for invalid characters
Select Case KeyAscii
 Case 48 To 57    'numbers
 Case 8, 45, 32, 13, 9 'special characters
 Case Else
 MsgBox msgNumbers
 KeyAscii = 0
 txtDuration.SetFocus
 Exit Sub
 End Select
End Sub

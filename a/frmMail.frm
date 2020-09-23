VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Mail Properties"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5910
   Begin VB.TextBox txtMess 
      Height          =   1575
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "frmMail.frx":030A
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtWebsite 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Information"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtServer 
      DataField       =   "212#49#84#2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   3975
   End
   Begin VB.TextBox txtFromAddress 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtBody 
      Height          =   1575
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMail.frx":0310
      Top             =   1800
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Other Message"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sender's Name"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Website:"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      Caption         =   "SMTP Host"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   825
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      Caption         =   "From Address:"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblBody 
      AutoSize        =   -1  'True
      Caption         =   "Conflict Message:"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1260
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
MsgBox msgOld
End Sub

''----------------------------------------
''- Name: Sam Huggill
''- Email: sam@vbsquare.com
''- Web: http;//www.vbsuare.com/
''- Company: Lighthouse Internet Solutions
''- Date/Time: 05/09/99 17:15:13
''----------------------------------------
''- Notes:   Originally written by standby@dellete.com
''           but modified by Sam Huggill
''
''           Mainly cleaned up the code for easier usage
''----------------------------------------
'
'Option Explicit
'
'Private Sub cmdAbout_Click()
'    frmAbout.Show vbModal
'End Sub
'
'Private Sub cmdClose_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSelect_Click()
'
'    cmdDialog.ShowOpen
'    txtAttach = cmdDialog.FileName
'
'End Sub
'
'Private Sub cmdSend_Click()
'
'    cmdSend.Enabled = False
'
'    If ValidateEntry = False Then MsgBox "Either the server name or to address were left empty.", vbCritical + vbOKOnly, Me.Caption: cmdSend.Enabled = True: Exit Sub
'
'    If txtAttach.Text <> "" Then
'        lblStatus = "Encoding file attachment"
'        Base64EncodeFile txtAttach.Text, rtfAttach, txtOutput
'    End If
'
'    lblStatus = "Connecting to POP Server"
'    ConnectToServer txtServer.Text, Winsock1
'
'End Sub
'
Private Sub Form_Load()
On Error GoTo ErrHandler
strSQL = "SELECT Mail.PopServer, Mail.FromAddress, Mail.Subject, Mail.Message, Mail.Message2, Mail.Website, Mail.Name FROM Mail"
GetRecordSet (strSQL)
txtWebsite.Text = objRS("Website").Value
txtServer.Text = objRS("PopServer").Value
txtFromAddress.Text = objRS("FromAddress").Value
txtSubject.Text = objRS("Subject").Value
txtBody.Text = objRS("Message").Value
txtMess.Text = objRS("Message2").Value
txtName.Text = objRS("Name").Value
Exit Sub
ErrHandler:
    ' Describe the error to the user.
    MsgBox "Unexpected error" & _
               vbCrLf & _
        Err.Description
    Exit Sub
End Sub


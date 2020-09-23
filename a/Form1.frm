VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration Form"
   ClientHeight    =   4560
   ClientLeft      =   3780
   ClientTop       =   1155
   ClientWidth     =   4755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   4245
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   2070
      TabIndex        =   18
      Top             =   4245
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
      Height          =   255
      Left            =   80
      TabIndex        =   17
      Top             =   4245
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   4080
      Width           =   4695
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   3080
         TabIndex        =   16
         Top             =   122
         Width           =   1575
      End
   End
   Begin VB.TextBox txtCountry 
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtCity 
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtOccupation 
      Height          =   405
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtAddress 
      Height          =   645
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Form"
      Height          =   4095
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   195
         Left            =   2880
         TabIndex        =   22
         Top             =   1500
         Width           =   975
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   195
         Left            =   1680
         TabIndex        =   21
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Country"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "City"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblAge 
         Caption         =   "Age"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Occupation"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Email"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdADD_Click()
On Error GoTo ErrHandler
Dim sName As String
Dim sAddress As String
Dim sEmail As String
Dim sAge As String
Dim sCity As String
Dim sOccupation As String
Dim sCountry As String

sName = txtName.Text
sAddress = txtAddress.Text
sEmail = txtEmail.Text
sAge = txtAge.Text
sCity = txtCity.Text
sOccupation = txtOccupation.Text
sCountry = txtCountry.Text

strSQL = " INSERT INTO Registration (Name, Address, Age, Email, City, Occupation, Country) "
strSQL = strSQL & " VALUES ('" & sName & "','" & sAddress & "','" & sAge & "','" & sEmail & "','" & sCity & "','" & sOccupation & "','" & sCountry & "')"
exCommand (strSQL)
Message = MsgBox("Record Succesfully Added Add Another Record", vbInformation + vbYesNo)
If Message = vbYes Then
    Call ClearFields
Else
    frmNew.cmdNew.Enabled = True
    Me.Hide
    Unload Me
End If
Exit Sub
ErrHandler:
    MsgBox "Error occured while adding record", vbInformation, "cmdAdd"
End Sub

Private Sub cmdCancel_Click()
frmNew.cmdNew.Enabled = True
Unload Me
End Sub

Public Sub ClearFields()
txtName.Text = ""
txtAddress.Text = ""
txtEmail.Text = ""
txtAge.Text = ""
txtCity.Text = ""
txtOccupation.Text = ""
txtCountry.Text = ""
End Sub


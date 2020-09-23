VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credits"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "frmCredits.frx":0742
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "Appointment Diary Beta Version" & vbCrLf & vbCrLf & _
        "by Murimi Ndungu" & vbCrLf & vbCrLf & _
        "Kenya , EastAfrica" & vbCrLf & vbCrLf & _
        "Please Visit our Web Page:  www.refpoint.antunit.com" & vbCrLf & vbCrLf & _
        "You can email me at: murimixp@gmail.com"
End Sub

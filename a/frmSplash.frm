VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   1920
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register Now"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7185
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Beta Version"
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
         Left            =   5400
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Appointment Diary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRegister_Click()
On Error GoTo HandleErrors
Call ShellExecute(Me.hwnd, "open", "http://refpoint.antunit.com/", 0, 0, vbNormalFocus)
Exit Sub
  
HandleErrors:

  MsgBox Err.Description, vbCritical, App.Title & " Error"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  '    PURPOSE: To execute code when the form is about to be un-loaded.
  '      NOTES: None.
  '  ARGUMENTS: Default.
  '    RETURNS: None.
  '    AUTHORS: T.Cummins (TCC) - 12/09/2000.
  '---------------------------------------------------------------------------------------------------------------------
  
  On Error GoTo HandleErrors
    
  'Num. sec. the Splash Window is displayed.
  Const cintDisplayTimeSeconds As Integer = 3
    
  'Loop until the Display Time has elpased - if the applications loading time took longer than
  'the display time it will not enter this loop.
  Do Until (Timer - msngSplashDisplayStartTime) > cintDisplayTimeSeconds
  Loop
  
  Screen.MousePointer = vbNormal
    
ExitHandleErrors:
  
  Exit Sub
  
HandleErrors:

  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
  
  On Error GoTo ErrHandler:
   
Exit Sub
ErrHandler:
Dim Message As String
Message = MsgBox("Program could not be initialized contact system administrator", vbCritical)
Exit Sub
End Sub


Private Sub Label3_Click()
On Error GoTo HandleErrors
Call ShellExecute(Me.hwnd, "open", "http://www.refpoint.antunit.com/", 0, 0, vbNormalFocus)
Exit Sub
  
HandleErrors:

  MsgBox Err.Description, vbCritical, App.Title & " Error"
End Sub

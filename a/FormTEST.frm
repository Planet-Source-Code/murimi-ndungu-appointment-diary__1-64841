VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim A As Integer
OpeningTime = 2200
ClosingTime = 2700
Duration = 30

For Counter = OpeningTime To ClosingTime
  If A >= 2400 Then
    'Counter = CInt(Counter) - 2400
    ClosingTime = CInt(ClosingTime) - 2400
  End If
  A = OpeningTime
   Counter = ConvNumTime(Counter)
    Counter = CDate(Counter)
    Counter = FormatDateTime(DateAdd("n", Duration, Counter), 4)
    Counter = Replace(Counter, ":", "")
   A = A + Duration
        If CInt(Counter) > ClosingTime Then
            Exit For
        End If
    Combo1.AddItem Counter

Counter = Replace(Counter, ":", "")
Counter = Counter - 1

Next

End Sub

Function ConvNumTime(pvTime)
Dim first
Dim last
Dim separator
separator = ":"
    Select Case Len(pvTime)
        Case 1
            first = "00"
            last = "0" & pvTime
        Case 2
            first = "00"
            last = pvTime
        Case 3
            first = Mid(pvTime, 1, 1)
            last = Mid(pvTime, 2, 2)
        Case 4
            first = Mid(pvTime, 1, 2)
            last = Mid(pvTime, 3, 2)
        Case Else
            first = "00"
            last = "00"
    End Select
    
    ConvNumTime = first & separator & last
End Function

VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2985
   ClientTop       =   2790
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************
' Name: Send SMS message via Http for fr
'     ee
' Description:Sends an SMS message to a
'     cell phone for free. It makes use of the
'     ServerXMLHTTP object contained in msxml3
'     .dll. Uses the free German Web service w
'     ww.billiger-telefonieren.de. The cookie
'     checks of the site are circumvented by d
'     oing the cookie
'handling explicitely. Therefore this code should work even server-side!
'Please note that the site still puts some requirement on the send message. For example messages with subjects like "test" are rejected.
'And: you can't send more than a certain number of messages to the the same number.


'For the most recent updates please visit my homepage.
' By: Klemens Schmid
'
'
' Inputs:1. Message text (up to 160 char
'     s)
'2. Phone number (e.g. +49171xxxxxx)
'
' Returns:Comes back with a success or a
'     failure message depending on the HTML th
'     at the site returns.
'
'Assumes:I used it with Microsoft msxml3
'     .dll (the released version). Download it
'     from http://msdn.microsoft.com/xml/defau
'     lt.asp.
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.5746/lngWId.1/qx/
'     vb/scripts/ShowCode.htm
'for details.
'**************************************

'Author
' mailto:klemens.schmid@gmx.de, http://w
'     ww.schmidks.de
'Description
' This code fires off an SMS message to
'     the given phone number
' It makes use of the German service "ww
'     w.billiger-telefonieren.de"
' The cookie checks of the site are circ
'     umvented by doing the cookie
' handling explicitely. Therefore this c
'     ode should work even server-side!
' Please note that the site still puts s
'     ome requirement on the send
' message. For example messages with sub
'     jects like "test" are rejected.
' And: you can't send more than a certai
'     n number of messages to the
' the same number.
'Prerequisites
' The posting is done thru the ServerXML
'     HTTP object which is contained
' in the Microsoft XML object msxml3.dll
'     . Install this from
' http://msdn.microsoft.com/xml/default.
'     asp.
Option Explicit


'Public Sub SendSMS()
'    Dim strText As String
'    Dim strPhoneNo As String
'    Dim strCookie As String
'    Dim oHttp As ServerXMLHTTP
'    'make use of the XMLHTTPRequest object c
'    '     ontained in msxml.dll
'    Set oHttp = CreateObject("msxml2.serverXMLHTTP")
'    'enter your data
'    strText = InputBox("Text:", "Send Text via SMS", "vbsms:")
'    strPhoneNo = InputBox("Phone Number:", "Send Text via SMS")
'    'fire of an http request to request for
'    '     a cookie
'    oHttp.Open "GET", "http://www.billiger-telefonieren.de/sms/send.php3?action=accept", False
'    oHttp.Send
'    strCookie = oHttp.getResponseHeader("set-cookie")
'    strCookie = Left$(strCookie, InStr(strCookie, ";") - 1)
'    'better check the feedback
'    Debug.Print oHttp.responseText
'    'do the actual send
'    oHttp.Open "POST", "http://www.billiger-telefonieren.de/sms/send.php3", False
'    oHttp.setRequestHeader "Cookie", strCookie
'    'we need to do it a second time due to K
'    '     B article Q234486.
'    oHttp.setRequestHeader "Cookie", strCookie
'    oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    oHttp.Send "action=send&number=" & strPhoneNo & "&email=&message=" & strText
'    Debug.Print oHttp.responseText
'
'
'    If InStr(oHttp.responseText, "erfolgreich eine Nachricht an die") Then
'        MsgBox "Message has been sent successfully", vbInformation
'    Else
'        MsgBox "Service refused to send the message", vbCritical
'    End If
'End Sub
'
'
'
'Private Sub Form_Load()
'SendSMS
'End Sub

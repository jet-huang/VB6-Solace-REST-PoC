VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   Caption         =   "VB6 Rest Requestor"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8730
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Caption         =   "Reply Mode"
      Height          =   970
      Left            =   6600
      TabIndex        =   14
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton opTimed 
         Caption         =   "Timed"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opTopic 
         Caption         =   "Customized Topic"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnSendMsg 
      Caption         =   "SEND MSG"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtGeneralMsg 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   6
      Text            =   "ASSEMBLE OBJECTs"
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton btnSendCmd 
      Caption         =   "SEND SQL"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtSqlCmd 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   4
      Text            =   "SELECT * FROM orders WHERE ORD_AMOUNT > 1000"
      Top             =   1440
      Width           =   5415
   End
   Begin VB.ComboBox cbTestCase 
      Height          =   300
      Index           =   1
      ItemData        =   "MainForm.frx":0000
      Left            =   360
      List            =   "MainForm.frx":0010
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtMsgReceived 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2770
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   1
      Text            =   "MainForm.frx":005A
      Top             =   2880
      Width           =   8415
   End
   Begin VB.Timer tmrSendData 
      Index           =   0
      Left            =   8040
      Top             =   6240
   End
   Begin MSWinsockLib.Winsock Sck 
      Index           =   0
      Left            =   7440
      Top             =   6240
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Label lblConnectStatus 
      AutoSize        =   -1  'True
      Caption         =   "(No Connection to Solace...)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   320
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   2870
   End
   Begin VB.Label lblSocketStatus 
      AutoSize        =   -1  'True
      Caption         =   "Socket Connected: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   290
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   1900
   End
   Begin VB.Label lblTopicName 
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   380
      Left            =   1560
      TabIndex        =   13
      Top             =   840
      Width           =   4930
   End
   Begin VB.Label lblTopicCaption 
      Caption         =   "Request Topic: "
      Height          =   260
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1090
   End
   Begin VB.Label lblNumRecv 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   380
      Left            =   4920
      TabIndex        =   11
      Top             =   480
      Width           =   980
   End
   Begin VB.Label lblNumSent 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   380
      Left            =   3360
      TabIndex        =   10
      Top             =   480
      Width           =   980
   End
   Begin VB.Label lblMsgRecv 
      Alignment       =   1  '靠右對齊
      Caption         =   "Reply Received"
      Height          =   260
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Width           =   1340
   End
   Begin VB.Label lblMsgSent 
      Alignment       =   1  '靠右對齊
      Caption         =   "Request Sent"
      Height          =   260
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   980
   End
   Begin VB.Label lblComputerCaption 
      Caption         =   "Select Topic to Send:"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "ASE VB6 Requestor with Solace"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' change this to your server name
Private Const ServerName As String = "Solace VB6 REST Consumer for ASE"
Private Const REPLY_TIMEOUT As Integer = 3000

Dim msgNumSent, msgNumRecv As Integer
Dim restListenPort As Integer
Dim sComputerName As String
Dim isConnected, isAutoReply As Boolean
Dim iSocketIndex, iSocketNum As Integer
Dim sTopicName As String
Dim myMSXML As Object

' Thanks to https://knowlet3389.blogspot.com/2012/02/vb6-delaysleep.html
Private Sub WaitSecs(ByVal Sec As Single)
    Dim sgnThisTime As Single, sgnCount As Single
    sgnThisTime = Timer
    Do While sgnCount < Sec
        sgnCount = Timer - sgnThisTime
        DoEvents
    Loop
End Sub

Private Function updateMessages(strMsg As String)
    txtMsgReceived.Text = txtMsgReceived.Text & strMsg & vbCrLf
    txtMsgReceived.SelStart = Len(txtMsgReceived.Text)
End Function

Private Function decodeMessage(strMsg As String)
    Dim strTemp As String
    strTemp = LeftRange(strMsg, "***REPLY***", "***END***", , ReturnEmptyStr)
    
    If Len(strTemp) <= 0 Then
        strTemp = "*** Not a valid reply, raw data: " & vbCrLf & strMsg
    End If
    
    decodeMessage = strTemp
End Function


Private Function getReplyMode() As String
    Dim sReplyMode As String
    
    If opTimed.Value Then
        sReplyMode = "Timed"
    ElseIf opTopic.Value Then
        sReplyMode = "Topic"
    Else
        sReplyMode = "Timed"
    End If
    
    getReplyMode = sReplyMode
End Function

Private Sub btnSendCmd_Click()
    Dim iCount As Integer
    Dim sReplyMode As String
    iCount = 0
    sReplyMode = getReplyMode()
    
    myMSXML.open "POST", "http://10.10.10.51:9000/" & sTopicName, False
    myMSXML.setRequestHeader "Content-Type", "text/plain"
    myMSXML.setRequestHeader "User-Agent", "VB6 REST Requestor"
    myMSXML.setRequestHeader "Solace-Client-Name", "REST-REQUESTOR-DBACCESS"
    ' NOTE: These 2 header parameters are EXCUSIVE.
    If sReplyMode Like "Timed*" Then
        myMSXML.setRequestHeader "Solace-Reply-Wait-Time-In-ms", REPLY_TIMEOUT
    Else
        myMSXML.setRequestHeader "Solace-Reply-To-Destination", "/TOPIC/OLD/CLIENT01/REPLY/MSG"
    End If
    ' myMSXML.setRequestHeader "Content-Length",
    myMSXML.send "***SQLCMD***" & txtSqlCmd & "***END***"
    msgNumSent = msgNumSent + 1
    lblNumSent.Caption = msgNumSent
    
    While (myMSXML.ReadyState <> 4) And (iCount < 10)
        WaitSecs 0.3
        iCount = iCount + 1
    Wend
    
    ' TODO: Duplicate code.
    If myMSXML.Status = 200 Then
        updateMessages vbCrLf & "*** Solace PS+ Broker received, waiting for remote server..."
        ' In "Timed" mode, the reply will come with the HTTP POST we sent, so it will not enter "Sck_DataArrival".
        If sReplyMode Like "Timed*" Then
            updateMessages "Reponse: " & vbCrLf & decodeMessage(myMSXML.responseText)
            msgNumRecv = msgNumRecv + 1
            lblNumRecv.Caption = msgNumRecv
        End If
    Else
        updateMessages vbCrLf & "*** Something wrong... (Error:" & myMSXML.Status & ")"
        updateMessages "Error reason: " & vbCrLf & myMSXML.responseText
    End If
End Sub

Private Sub btnSendMsg_Click()
    Dim iCount As Integer
    Dim sReplyMode As String
    iCount = 0
    sReplyMode = getReplyMode()
    
    myMSXML.open "POST", "http://10.10.10.51:9000/" & sTopicName, False
    myMSXML.setRequestHeader "Content-Type", "text/plain"
    myMSXML.setRequestHeader "User-Agent", "VB6 REST Requestor"
    myMSXML.setRequestHeader "Solace-Client-Name", "REST-REQUESTOR-GENERAL"
    ' NOTE: These 2 header parameters are EXCUSIVE.
    If sReplyMode Like "Timed*" Then
        myMSXML.setRequestHeader "Solace-Reply-Wait-Time-In-ms", REPLY_TIMEOUT
    Else
        myMSXML.setRequestHeader "Solace-Reply-To-Destination", "/TOPIC/OLD/CLIENT01/REPLY/MSG"
    End If
    ' myMSXML.setRequestHeader "Content-Length",
    myMSXML.send "***BEGIN***General message from Rest Requestor:" & vbCrLf & txtGeneralMsg & "***END***"
    msgNumSent = msgNumSent + 1
    lblNumSent.Caption = msgNumSent
    
    While (myMSXML.ReadyState <> 4) And (iCount < 10)
        WaitSecs 0.3
        iCount = iCount + 1
    Wend
    
    ' TODO: Duplicate code.
    If myMSXML.Status = 200 Then
        updateMessages vbCrLf & "*** Solace PS+ Broker received, waiting for remote server..."
        ' In "Timed" mode, the reply will come with the HTTP POST we sent, so it will not enter "Sck_DataArrival".
        If sReplyMode Like "Timed*" Then
            updateMessages "Reponse: " & vbCrLf & decodeMessage(myMSXML.responseText)
            msgNumRecv = msgNumRecv + 1
            lblNumRecv.Caption = msgNumRecv
        End If
    Else
        updateMessages "*** Something wrong... (Error:" & myMSXML.Status & ")"
        updateMessages "Error reason: " & vbCrLf & myMSXML.responseText
    End If
End Sub

Private Sub cbTestCase_Click(Index As Integer)
    Dim iCaseSelected As Integer
    iCaseSelected = cbTestCase(Index).ListIndex
    
    Select Case iCaseSelected
        Case 0
            sTopicName = "ASSY/ICONND/WBI002"
            btnSendCmd.Enabled = False
            btnSendMsg.Enabled = True
        Case 1
            sTopicName = "ASSY/PROCU/MULTI"
            btnSendCmd.Enabled = False
            btnSendMsg.Enabled = True
        Case 2
            sTopicName = "ASSY/BROADCAST"
            btnSendCmd.Enabled = False
            btnSendMsg.Enabled = True
        Case 3
            sTopicName = "CMD/SQL/SELECT_REQUEST"
            btnSendCmd.Enabled = True
            btnSendMsg.Enabled = False
        Case Else
            sTopicName = "ASSY/*"
            btnSendCmd.Enabled = False
            btnSendMsg.Enabled = True
    End Select
    
    ' MsgBox "Topic name: " & sTopicName
    lblTopicName.Caption = sTopicName
End Sub

Private Sub Form_Load()
    sComputerName = "REST-REQUESTOR"
    msgNumSent = 0
    msgNumRecv = 0
    isConnected = False
    btnSendCmd.Enabled = False
    btnSendMsg.Enabled = False
    Set myMSXML = CreateObject("Microsoft.XMLHttp")
    opTimed.Caption = "Timed (" & (REPLY_TIMEOUT / 1000) & " secs)"
    
    ' This is for replys from server-side.
    restListenPort = 9999
    Sck(0).LocalPort = restListenPort ' set this to the port you want the server to listen on...
    Sck(0).Listen
    DoEvents
End Sub

Private Sub Sck_Close(Index As Integer)
    ' make sure the connection is closed
    Do
        Sck(Index).Close
        DoEvents
    Loop Until Sck(Index).State = sckClosed

    iSocketNum = iSocketNum - 1
    
    If iSocketNum <= 0 Then
        iSocketNum = 0
        lblSocketStatus.Caption = "No TCP socket used."
    Else
        lblSocketStatus.Caption = "Opened sockets: " & iSocketNum
    End If
End Sub

Private Sub Sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer
    
    ' look in the control array for a closed connection
    ' note that it's starting to search at index 1 (not index 0)
    ' since index 0 is the one listening on port 80
    For K = 1 To Sck.UBound
        If Sck(K).State = sckClosed Then Exit For
    Next K
    
    ' if all controls are connected, then create a new one
    If K > Sck.UBound Then
        K = Sck.UBound + 1
        Load Sck(K) ' create a new winsock object
    End If
    
    ' accept the connection on the closed control or the new control
    Sck(K).Accept requestID
    iSocketNum = iSocketNum + 1
    isConnected = True
    lblConnectStatus(0).Caption = sComputerName & " has connected to Solace PS+ Broker..." & vbCrLf & "Receiving messages from: " & restListenPort
    lblSocketStatus.Caption = "Opened sockets: " & iSocketNum
End Sub

Private Sub Sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String
    Dim strTopic, strPostPath As String
    Dim strMessage As String
    Dim strReplyMsg As String
    
    iSocketIndex = Index
    Sck(Index).GetData rData, vbString
    
    If rData Like "POST * HTTP/1.?*" Then
        strPostPath = LeftRange(rData, "POST ", " HTTP/1.", , ReturnEmptyStr)
        strMessage = decodeMessage(rData)
        updateMessages "Message: " & vbCrLf & strMessage
        ContentType = "Content-Type: text/html; charset=UTF-8"
        ' Build HTTP header
        sHeader = "HTTP/1.1 200 OK" & vbNewLine & _
            "Server: " & ServerName & vbNewLine & _
            ContentType & vbNewLine & _
            "Content-Length: " & Len(strReplyMsg) & vbNewLine & _
            vbNewLine
        Sck(Index).SendData sHeader
        msgNumRecv = msgNumRecv + 1
        lblNumRecv.Caption = msgNumRecv
    Else
        sHeader = "HTTP/1.1 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
        Sck(Index).SendData sHeader
    End If
End Sub

Private Sub Sck_SendComplete(Index As Integer)
    ' Since it's not a really "Web Server", we don't need to disconnect the session from Solace every message.
    ' Sck_Close Index
End Sub

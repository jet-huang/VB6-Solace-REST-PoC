VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   Caption         =   "VB6 REST Replier (Solace PoC)"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8730
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Caption         =   "Data Source"
      Height          =   975
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton opMySql 
         Caption         =   "MySQL"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton opSqlServer 
         Caption         =   "SQL Server"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnSendReply 
      Caption         =   "REPLY"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox tbReplyMsg 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Text            =   "OK"
      Top             =   4680
      Width           =   5415
   End
   Begin VB.CheckBox ckAutoReply 
      Caption         =   "Auto Reply"
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   4680
      Value           =   1  '核取
      Width           =   1335
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect!"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cbComputerName 
      Height          =   300
      Index           =   1
      ItemData        =   "MainForm.frx":0000
      Left            =   360
      List            =   "MainForm.frx":002E
      Style           =   2  '單純下拉式
      TabIndex        =   3
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
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   2
      Text            =   "MainForm.frx":006A
      Top             =   1200
      Width           =   8415
   End
   Begin VB.Timer tmrSendData 
      Index           =   0
      Left            =   8040
      Top             =   6240
   End
   Begin MSWinsockLib.Winsock Sck 
      Index           =   0
      Left            =   7560
      Top             =   6240
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
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
      TabIndex        =   13
      Top             =   6240
      Width           =   1900
   End
   Begin VB.Label lblReplyStatus 
      AutoSize        =   -1  'True
      Caption         =   "Reply Status"
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
      TabIndex        =   12
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label lblComputerCaption 
      Caption         =   "Computer Name:"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "Solace with VB6 (Server Side)"
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
      TabIndex        =   1
      Top             =   6840
      Width           =   7935
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
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   2865
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' change this to your server name
Private Const ServerName As String = "Solace-VB6-REST-Replier"

Dim msgNum As Integer
Dim restListenPort As Integer
Dim sComputerName As String
Dim isConnected, isAutoReply As Boolean
Dim iSocketIndex, iSocketNum As Integer
Dim sSqlCmd, sDbName As String

Private Function getDateTimeString() As String
    getDateTimeString = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & _
        Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Function

Private Function queryData() As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strResult As String
    
    ' MsgBox "In queryData(): " & sSqlCmd & ", DB:" & sDbName
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    '
    If sDbName Like "MySQL*" Then
        conn.ConnectionString = "dsn=vb6-mariadb"
    Else
        ' Don't know why but it's not possible to connect MSSQL with the username/password set in DSN.
        ' You must enter the credential here.
        conn.ConnectionString = "dsn=vb6-mssql;uid=sa;pwd=Solace1234"
    End If
    
    conn.Open
    rs.ActiveConnection = conn
    rs.Open sSqlCmd
    
    Do While Not rs.EOF
        strResult = strResult & rs("ORD_NUM") & ": " & rs("CUST_CODE") & vbCrLf
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    queryData = strResult
End Function

Private Function getCurrentDataSource() As String
    Dim sDataSource As String
    
    If opSqlServer.Value Then
        sDataSource = "SQLServer"
    ElseIf opMySql.Value Then
        sDataSource = "MySQL"
    Else
        sDataSource = "SQLServer"
    End If
    
    getCurrentDataSource = sDataSource
End Function

Private Function updateMessages(strMsg As String)
    txtMsgReceived.Text = txtMsgReceived.Text & strMsg & vbCrLf
    txtMsgReceived.SelStart = Len(txtMsgReceived.Text)
End Function

Private Sub btnConnect_Click()
    If isConnected = False Then
        lblConnectStatus(0).Caption = sComputerName & " is connecting to Solace..."
        Sck(0).LocalPort = restListenPort ' set this to the port you want the server to listen on...
        Sck(0).Listen
        DoEvents
        msgNum = 0
        btnConnect.Caption = "Disconnect"
        txtMsgReceived.Text = ""
        cbComputerName(1).Enabled = False
    Else
        lblConnectStatus(0).Caption = sComputerName & " is disonnecting from Solace..."
        Dim Index As Integer
    
        For Index = 0 To Sck.UBound
            If Sck(Index).State <> sckClosed Then
                Sck_Close Index
            End If
        Next Index
        isConnected = False
        btnConnect.Caption = "Connect"
        lblConnectStatus(0).Caption = "(No Connection to Solace...)"
        cbComputerName(1).Enabled = True
    End If
End Sub

Private Sub btnSendReply_Click()
    Dim strReplyMsg As String
    Dim sHeader As String, ContentType As String

    ContentType = "Content-Type: text/html; charset=UTF-8"
    strReplyMsg = "***REPLY***Reply from " & "SOMEONE" & " @ " & _
        getDateTimeString() & vbCrLf & tbReplyMsg.Text & " | Msg No. " & (msgNum + 1) & "***END***"
    ' Build HTTP header
    sHeader = "HTTP/1.1 200 OK" & vbNewLine & _
        "Server: " & ServerName & vbNewLine & _
        ContentType & vbNewLine & _
        "Content-Length: " & Len(strReplyMsg) & vbNewLine & _
        vbNewLine
    Sck(iSocketIndex).SendData sHeader
    Sck(iSocketIndex).SendData strReplyMsg
    msgNum = msgNum + 1
    lblReplyStatus.Caption = "Receiving message No. " & msgNum & ", ACKED."
End Sub

Private Sub cbComputerName_Click(Index As Integer)
     restListenPort = cbComputerName(Index).ItemData(cbComputerName(Index).ListIndex)
     sComputerName = cbComputerName(Index).Text
End Sub

Private Sub ckAutoReply_Click()
    If isAutoReply Then
        tbReplyMsg.Enabled = True
        btnSendReply.Enabled = True
        isAutoReply = False
    Else
        tbReplyMsg.Enabled = False
        btnSendReply.Enabled = False
        isAutoReply = True
    End If
End Sub

Private Sub Form_Load()
    restListenPort = 8080
    msgNum = 0
    isConnected = False
    isAutoReply = True
    ckAutoReply.Value = 1
    tbReplyMsg.Enabled = False
    btnSendReply.Enabled = False
    iSocketNum = 0
    lblSocketStatus.Caption = "No TCP socket used."
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
    ' since index 0 is the one listening on default port
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
        ' Determine which action should be taken
        '' General request/reply message
        strMessage = LeftRange(rData, "***BEGIN***", "***END***", , ReturnEmptyStr)
        If Len(strMessage) > 0 Then
            strReplyMsg = "AUTO REPLY TO: " & strMessage & " | Msg No. " & (msgNum + 1)
        '' SQL Command
        Else
            strMessage = LeftRange(rData, "***SQLCMD***", "***END***", , ReturnEmptyStr)
            If Len(strMessage) > 0 Then
                sDbName = getCurrentDataSource()
                sSqlCmd = strMessage
                ' MsgBox sSqlCmd
                strReplyMsg = "DB source: " & sDbName
                strReplyMsg = strReplyMsg & vbCrLf & queryData()
            Else
                strReplyMsg = "Unknown command..."
            End If
        End If
        
        updateMessages "Message: " & vbCrLf & strMessage
        ' txtMsgReceived.Text = txtMsgReceived.Text & vbCrLf & "Message: " & vbCrLf & strMessage
        ' txtMsgReceived.Text = txtMsgReceived.Text & vbCrLf & rData
        ' sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
        If isAutoReply Then
            ' build the header
            strReplyMsg = "***REPLY***Reply from " & sComputerName & " @ " & _
                getDateTimeString() & vbCrLf & strReplyMsg & "***END***"
            ContentType = "Content-Type: text/html; charset=UTF-8"
            ' Build HTTP header
            sHeader = "HTTP/1.1 200 OK" & vbNewLine & _
                "Server: " & ServerName & vbNewLine & _
                ContentType & vbNewLine & _
                "Content-Length: " & Len(strReplyMsg) & vbNewLine & _
                vbNewLine
                ' send the header, the Sck_SendComplete event is gonna send the file...
            Sck(Index).SendData sHeader
            Sck(Index).SendData strReplyMsg
            msgNum = msgNum + 1
            lblReplyStatus.Caption = "Receiving message No. " & msgNum & ", ACKED."
        Else
            'lblConnectStatus(Index).Caption = "Waiting for your reply..."
            lblReplyStatus.Caption = "Waiting for your reply..."
            iSocketIndex = Index
        End If
    Else
        sHeader = "HTTP/1.1 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
        Sck(Index).SendData sHeader
    End If
End Sub

Private Sub Sck_SendComplete(Index As Integer)
    ' Since it's not a really "Web Server", we don't need to disconnect the session from Solace every message.
    ' Sck_Close Index
    ' lblReplyStatus.Caption = "Send complete..."
End Sub

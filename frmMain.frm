VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{7390FFEE-1DEE-418E-88E3-CD3ED4FB6D54}#30.0#0"; "Sox.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoxOCX Demo Project"
   ClientHeight    =   8355
   ClientLeft      =   1065
   ClientTop       =   1425
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9195
   Begin SoxOCX.Sox Sox 
      Height          =   420
      Left            =   8640
      Top             =   360
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Frame frmServer 
      Caption         =   "Server"
      Height          =   5955
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   4515
      Begin MSComctlLib.ProgressBar barServerRecv 
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   5640
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame frmSend2Client 
         Caption         =   "Send 2 Client"
         Height          =   1875
         Left            =   180
         TabIndex        =   31
         Top             =   2700
         Width           =   4155
         Begin VB.TextBox txtFloodClientBytes 
            Height          =   315
            Left            =   1620
            TabIndex        =   51
            Text            =   "1000000"
            ToolTipText     =   "Type a message to send to the Server here"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton btnFloodClient 
            Caption         =   "Flood Client"
            Height          =   495
            Left            =   180
            TabIndex        =   50
            ToolTipText     =   "Send the message to the Server"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.CommandButton btnSend2Client 
            Caption         =   "Send Msg"
            Height          =   495
            Left            =   180
            TabIndex        =   34
            ToolTipText     =   "Send the message to the connected client Socket"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtSend2ClientSocket 
            Height          =   315
            Left            =   2940
            TabIndex        =   33
            Text            =   "0"
            ToolTipText     =   "Socket 0 will broadcast to all sockets"
            Top             =   300
            Width           =   1035
         End
         Begin VB.TextBox txtClientMessage 
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Text            =   "Hello"
            Top             =   720
            Width           =   1695
         End
         Begin MSComctlLib.ProgressBar barServerSend 
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   1620
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   ":Bytes"
            Height          =   195
            Left            =   3540
            TabIndex        =   52
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label lblSend2ClientSocket 
            AutoSize        =   -1  'True
            Caption         =   "To Socket:"
            Height          =   195
            Left            =   2040
            TabIndex        =   36
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lblClientMessage 
            AutoSize        =   -1  'True
            Caption         =   "Message:"
            Height          =   195
            Left            =   1500
            TabIndex        =   35
            Top             =   720
            Width           =   690
         End
      End
      Begin VB.TextBox txtServerRecv 
         Height          =   915
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   4680
         Width           =   3375
      End
      Begin VB.Frame frmServerSetup 
         Caption         =   "Setup"
         Height          =   2355
         Left            =   180
         TabIndex        =   23
         ToolTipText     =   "Binding is usually used on Multiple network card configurations"
         Top             =   300
         Width           =   4155
         Begin VB.TextBox txtServerAddress 
            Height          =   315
            Left            =   2520
            TabIndex        =   41
            Text            =   "127.0.0.1"
            ToolTipText     =   "The network card you want to dedicate to Listening (Address 255.255.255.255 is invalid)"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtServerPort 
            Height          =   315
            Left            =   2520
            TabIndex        =   40
            Text            =   "1234"
            ToolTipText     =   "The network port you want to dedicate to Listening (Valid ports R in range 0 to 5000)"
            Top             =   540
            Width           =   1455
         End
         Begin VB.CommandButton btnBind 
            Caption         =   "Bind"
            Height          =   435
            Left            =   180
            TabIndex        =   39
            ToolTipText     =   "Bind the Port and Address (Binding is done automatically in Sox, but is used to block a port, should be rename Block instead :)"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtListenSocket 
            Height          =   315
            Left            =   1740
            TabIndex        =   27
            ToolTipText     =   "The Socket that is dedicated to Listening"
            Top             =   1860
            Width           =   975
         End
         Begin VB.TextBox txtConnections 
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "0"
            ToolTipText     =   "How many Sockets are being used"
            Top             =   1500
            Width           =   975
         End
         Begin VB.CommandButton btnListen 
            Caption         =   "Listen"
            Height          =   435
            Left            =   180
            TabIndex        =   25
            ToolTipText     =   "Begin Listening on the "
            Top             =   900
            Width           =   1215
         End
         Begin VB.CommandButton btnCloseListen 
            Caption         =   "Close"
            Height          =   315
            Left            =   2820
            TabIndex        =   24
            Top             =   1860
            Width           =   1215
         End
         Begin VB.Label lblServerAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            Height          =   195
            Left            =   1740
            TabIndex        =   0
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblServerPort 
            AutoSize        =   -1  'True
            Caption         =   "Port:"
            Height          =   195
            Left            =   2025
            TabIndex        =   42
            Tag             =   "The network port you want to dedicate to Listening"
            Top             =   600
            Width           =   330
         End
         Begin VB.Label lblListenSocket 
            AutoSize        =   -1  'True
            Caption         =   "Listening On Socket:"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   1860
            Width           =   1485
         End
         Begin VB.Label lblConnections 
            AutoSize        =   -1  'True
            Caption         =   "Current Connections:"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   1500
            Width           =   1485
         End
      End
      Begin VB.Label lblServerRecv 
         Caption         =   "Received:"
         Height          =   255
         Left            =   180
         TabIndex        =   37
         Top             =   4680
         Width           =   735
      End
   End
   Begin VB.Frame frmClient 
      Caption         =   "Client"
      Height          =   4995
      Left            =   4620
      TabIndex        =   6
      Top             =   1020
      Width           =   4515
      Begin VB.Frame frmSend2Server 
         Caption         =   "Send 2 Server"
         Height          =   1815
         Left            =   180
         TabIndex        =   17
         Top             =   1860
         Width           =   4155
         Begin VB.TextBox txtFloodServerBytes 
            Height          =   315
            Left            =   1620
            TabIndex        =   48
            Text            =   "1000000"
            ToolTipText     =   "Type a message to send to the Server here"
            Top             =   960
            Width           =   1875
         End
         Begin VB.CommandButton btnFloodServer 
            Caption         =   "Flood Server"
            Height          =   495
            Left            =   180
            TabIndex        =   47
            ToolTipText     =   "Send the message to the Server"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox txtMessage2Server 
            Height          =   315
            Left            =   1620
            TabIndex        =   19
            Text            =   "Hello"
            ToolTipText     =   "Type a message to send to the Server here"
            Top             =   480
            Width           =   2355
         End
         Begin VB.CommandButton btnSend2Server 
            Caption         =   "Send Msg"
            Height          =   495
            Left            =   180
            TabIndex        =   18
            ToolTipText     =   "Send the message to the Server"
            Top             =   300
            Width           =   1215
         End
         Begin MSComctlLib.ProgressBar barClientSend 
            Height          =   195
            Left            =   180
            TabIndex        =   45
            Top             =   1500
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   ":Bytes"
            Height          =   195
            Left            =   3540
            TabIndex        =   49
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label lblMessage2Server 
            AutoSize        =   -1  'True
            Caption         =   "Message:"
            Height          =   195
            Left            =   2340
            TabIndex        =   20
            Top             =   180
            Width           =   690
         End
      End
      Begin VB.Frame frmClientSetup 
         Caption         =   "Setup"
         Height          =   1455
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   4155
         Begin VB.CommandButton btnConnect 
            Caption         =   "Connect"
            Height          =   495
            Left            =   360
            TabIndex        =   13
            ToolTipText     =   "Make a connection to the server at the Port and Address specified"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtClientAddress 
            Height          =   315
            Left            =   2700
            TabIndex        =   12
            Text            =   "127.0.0.1"
            ToolTipText     =   "The IP / Internet address you want to connect to"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtClientPort 
            Height          =   315
            Left            =   2700
            TabIndex        =   11
            Text            =   "1234"
            ToolTipText     =   "The network port you want to connect to"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtClientSocket 
            Height          =   315
            Left            =   2700
            TabIndex        =   10
            ToolTipText     =   "This is the Socket on the Server we have access to"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.CommandButton btnCloseConnect 
            Caption         =   "Close"
            Height          =   315
            Left            =   360
            TabIndex        =   9
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label lblClientAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            Height          =   195
            Left            =   2040
            TabIndex        =   16
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblClientPort 
            AutoSize        =   -1  'True
            Caption         =   "Port:"
            Height          =   195
            Left            =   2280
            TabIndex        =   15
            Tag             =   "The network port you want to dedicate to Listening"
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblClientSocket 
            AutoSize        =   -1  'True
            Caption         =   "Socket:"
            Height          =   195
            Left            =   2040
            TabIndex        =   14
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   1020
            Width           =   555
         End
      End
      Begin VB.TextBox txtClientRecv 
         Height          =   915
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3720
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar barClientRecv 
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   4680
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblClientRecv 
         Caption         =   "Received:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   735
      End
   End
   Begin VB.Frame frmSubClassing 
      Caption         =   "SubClassing"
      Height          =   915
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "This must be done before we can receive messages and MUST be undone before form unload"
      Top             =   60
      Width           =   2775
      Begin VB.CommandButton btnUnhook 
         Caption         =   "Unhook"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "UnHook messaging sub-system (MUST be done before form unloads ... done automatically there currently)"
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton btnHook 
         Caption         =   "Hook"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Hook into messaging sub-system (Must be done before ANY messages will be received from WinSock)"
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      ToolTipText     =   "Test ShutDown procedure!"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtStatus 
      Height          =   2235
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6060
      Width           =   8535
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   60
      TabIndex        =   38
      Top             =   6060
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBind_Click()
    Call Sox.Bind(txtServerAddress, Val(txtServerPort))
End Sub

Private Sub btnCloseConnect_Click()
    Call Sox.Shut(Val(txtClientSocket))
End Sub

Private Sub btnCloseListen_Click()
    Call Sox.Shut(Val(txtListenSocket))
End Sub

Private Sub btnConnect_Click()
    Let txtClientSocket = Sox.Connect(txtClientAddress, Val(txtClientPort))
    Let txtSend2ClientSocket = txtClientSocket ' Updates this text box to the latest connection number
End Sub

Private Sub btnFloodClient_Click()
    Call Sox.SendData(Val(txtSend2ClientSocket), Space$(Val(txtFloodClientBytes)))
End Sub

Private Sub btnFloodServer_Click()
    Call Sox.SendData(Val(txtClientSocket), Space$(Val(txtFloodServerBytes)))
End Sub

Private Sub btnHook_Click()
    Call Sox.Hook
End Sub

Private Sub btnListen_Click()
    Let txtListenSocket.Text = Sox.Listen(txtServerAddress, txtServerPort)
End Sub

Private Sub btnSend2Client_Click()
    Call Sox.SendData(txtSend2ClientSocket, txtClientMessage)
End Sub

Private Sub btnSend2Server_Click()
    Call Sox.SendData(Val(txtClientSocket), txtMessage2Server)
End Sub

Private Sub btnTest_Click() ' Used for general testing purposes ... usually a Sox command eg. MsgBox Sox.GetOption(0, soxSO_RCVBUF)
    Call Sox.ShutDown
End Sub

Private Sub btnUnhook_Click()
    Call Sox.UnHook
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Sox.ShutDown = soxERROR Then  ' Terminate will be True if we have ShutDown properly
        If MsgBox("ShutDown procedure has not completed!" & vbCrLf & "(Hint - Select No and Try again!)" & vbCrLf & "Execute Forced ShutDown?", vbApplicationModal + vbCritical + vbYesNo, "UNABLE TO COMPLY!") = vbNo Then
            Let Cancel = True
        Else
            Call Sox.UnHook ' Unfortunately for now, I can't get around doing this automatically for you :( VB crashes if you don't do this!
        End If
    Else
        Call Sox.UnHook ' The reason is VB closes my Mod which stores the WindowProc function used for SubClassing and VB doesn't know that! So it closes the Mod before the Control!
    End If
End Sub

Private Sub Sox_OnClose(inSox As Long) ' A client has just disconnected (Or we closed one of our Listening/Bound sockets) could implement the type of connection that was closed, eg. Listening socket, Bound socket or client socket
    Let txtConnections = Sox.Connections ' Show the number of connections to our Server
End Sub

Private Sub Sox_OnConnect(inSox As Long) ' Successfully connected to a Server (Listening Socket)
    Let txtClientSocket = inSox ' Shows our latest Sox number
    Let txtConnections = Sox.Connections ' Show the number of connections to our Server
End Sub

Private Sub Sox_OnConnection(inSox As Long) ' A new connection to our Server has been established
    Let txtConnections = Sox.Connections ' Show the number of connections to our Server
End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte) ' We will only show the first 20 characters returned
    Select Case Sox.uMsg(inSox) ' uMsg is the term I used, sorry it's not really readable as the Test for the Type of socket, but I couldn't use the word Sox.Type :( any suggestions ???
        Case soxCLIENT ' Listening & Bound Sockets are excluded here automatically 'cause they cannot send & recv, but beware the existance of them when doing a test like this
            Let txtServerRecv = txtServerRecv & "Sox " & inSox & " Received " & UBound(inData) + 1 & " bytes" & vbCrLf
            Let txtServerRecv = txtServerRecv & Left$(StrConv(inData, vbUnicode), 20) & vbCrLf
        Case soxSERVER ' Server in this case is the indication of what socket is connected on the other side, therefore normal 'Client' connections are reffered to as soxSERVER 'cause they have a server on the other side!!!
            Let txtClientRecv = txtClientRecv & "Sox " & inSox & " Received " & UBound(inData) + 1 & " bytes" & vbCrLf
            Let txtClientRecv = txtClientRecv & Left$(StrConv(inData, vbUnicode), 20) & vbCrLf
    End Select
End Sub

Private Sub Sox_OnError(inSox As Long, inError As Long, inDescription As String, inSource As String, inSnipet As String)
    With txtStatus
        Let .Text = .Text & "Error: (Sox) " & inSox & " (Error) " & inError & " (Description) " & inDescription & " (Source) " & inSource & " (Area) " & inSnipet & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

Private Sub Sox_OnRecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)
    Select Case Sox.uMsg(inSox) ' uMsg is the term I used, sorry it's not really readable as the Test for the Type of socket, but I couldn't use the word Sox.Type :( any suggestions ???
        Case soxCLIENT ' Listening & Bound Sockets are excluded here automatically 'cause they cannot send & recv, but beware the existance of them when doing a test like this
            Let barServerRecv.Value = bytesRecv / (bytesRecv + bytesRemaining) * 100
        Case soxSERVER ' Server in this case is the indication of what socket is connected on the other side, therefore normal 'Client' connections are reffered to as soxSERVER 'cause they have a server on the other side!!!
            Let barClientRecv.Value = bytesRecv / (bytesRecv + bytesRemaining) * 100
    End Select
End Sub

Private Sub Sox_OnSendComplete(inSox As Long)
    With txtStatus
        Let .Text = .Text & "Sox " & inSox & " Sent Data successfully" & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

Private Sub Sox_OnSendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)
    Select Case Sox.uMsg(inSox) ' uMsg is the term I used, sorry it's not really readable as the Test for the Type of socket, but I couldn't use the word Sox.Type :( any suggestions ???
        Case soxCLIENT ' Listening & Bound Sockets are excluded here automatically 'cause they cannot send & recv, but beware the existance of them when doing a test like this
            Let barServerSend.Value = bytesSent / (bytesSent + bytesRemaining) * 100
        Case soxSERVER ' Server in this case is the indication of what socket is connected on the other side, therefore normal 'Client' connections are reffered to as soxSERVER 'cause they have a server on the other side!!!
            Let barClientSend.Value = bytesSent / (bytesSent + bytesRemaining) * 100
    End Select
End Sub

Private Sub Sox_OnState(inSox As Long, inState As enmSoxState)
    With txtStatus
        Select Case inState
            Case soxDisconnected
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Disconnected" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxConnecting
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Connecting" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxIdle
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Idle" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxSend
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Sending" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxRecv
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Receiving" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxClosing
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Closing" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxBound
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Bound" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxListening
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Listening" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case Else
                Let .Text = .Text & "State: (Sox) " & inSox & " (State) Unknown State: " & inState & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
        End Select
    End With
End Sub

Private Sub Sox_OnStatus(inSox As Long, inSource As String, inStatus As String)
    With txtStatus
        Let .Text = .Text & "Status: (Sox) " & inSox & " (Source) " & inSource & " (Info) " & inStatus & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

Private Sub txtClientRecv_Click()
    Let txtClientRecv.SelStart = 0
    Let txtClientRecv.SelLength = Len(txtClientRecv)
End Sub

Private Sub txtClientRecv_DblClick()
    Let txtClientRecv = ""
End Sub

Private Sub txtServerRecv_Click()
    Let txtServerRecv.SelStart = 0
    Let txtServerRecv.SelLength = Len(txtServerRecv)
End Sub

Private Sub txtServerRecv_DblClick()
    Let txtServerRecv = ""
End Sub

Private Sub txtStatus_Click()
    Let txtStatus.SelStart = 0
    Let txtStatus.SelLength = Len(txtStatus)
End Sub

Private Sub txtStatus_DblClick()
    Let txtStatus = ""
End Sub

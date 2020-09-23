VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Client"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtReceived 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   2880
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Send Message"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Received From Server"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If txtUserName.Text = "" Then
        MsgBox "You need to type your username!", vbCritical, "Unable to complete"
        Exit Sub
    End If
    wsMain.Connect
    Do Until wsMain.State = 7
        ' 0 is closed, 9 is error
        If wsMain.State = 0 Or wsMain.State = 9 Then
            MsgBox "Error in connecting!", vbCritical, "Winsock Error"
            ' there was an error, so let's leave
            Exit Sub
        End If
        DoEvents  'don't freeze the system!
    Loop
    ' "log-in":
    wsMain.SendData "U" & Chr(1) & txtUserName.Text
    txtUserName.Enabled = False
    txtMessage.Enabled = True
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        wsMain.SendData "t" & Chr(1) & txtMessage.Text
        txtMessage.Text = ""
        KeyAscii = 0
    End If
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
Dim Data As String, CtrlChar As String
    wsMain.GetData Data
    CtrlChar = Left(Data, 1) ' Let's get the first char
    Data = Mid(Data, 3)      ' Then cut it off
    Select Case LCase(CtrlChar)   ' Check what it is
        Case "m"   ' Do stuff depending on it
            MsgBox Data, vbInformation, "Msg from server"
        Case "c"
            Me.Caption = "Client - " & Data
        Case Else
            txtReceived.SelStart = Len(txtReceived.Text)
            txtReceived.SelText = Data & vbCrLf
    End Select
End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: " & Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub

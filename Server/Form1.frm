VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Server"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMsgBox 
      Caption         =   "Popup Message Box"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCaption 
      Caption         =   "Set their Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtReceived 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox txtSendMessage 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtErrors 
      Height          =   1335
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wsArray 
      Index           =   0
      Left            =   4320
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2500
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   0
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2400
   End
   Begin VB.ListBox lstUsers 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "(shift-enter to broadcast)"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Received"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Send Message"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Error Log"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Users"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' We'll limit it to 101 users at a time! ;)
Dim Users(0 To 100) As String

Private Sub cmdCaption_Click()
Dim User As Integer
    ' Get Username to send to
    User = RetrieveUser(lstUsers.Text)
    If User = -1 Then
        MsgBox "Invalid User!", vbCritical, "Error"
        Exit Sub
    End If
    wsArray(User).SendData "c" & Chr(1) & InputBox("What do you want to have their caption set to?", "Alter Caption", "Hi!")
End Sub

Private Sub cmdMsgBox_Click()
Dim User As Integer
    ' Get Username to send to
    User = RetrieveUser(lstUsers.Text)
    If User = -1 Then
        MsgBox "Invalid User!", vbCritical, "Error"
        Exit Sub
    End If
    wsArray(RetrieveUser(lstUsers.Text)).SendData "m" & Chr(1) & InputBox("What do you want to have displayed on their machine?", "Popup MsgBox", "Hi!")
End Sub

Private Sub Form_Load()
    wsListen.Listen  ' make it listen
End Sub

Private Sub txtSendMessage_KeyDown(KeyCode As Integer, Shift As Integer)
Dim User As Integer
    
    'First, check to make sure someone's logged in
    If lstUsers.ListCount = 0 And KeyCode = 13 Then
    
        'Display popup
        MsgBox "Nobody to send to!", vbExclamation, "Cannot send"
        
        'Clear input
        txtSendMessage.Text = ""
        Exit Sub
    End If

    ' If it was enter and shift wasn't pressed, then...
    If KeyCode = 13 And Shift = 0 Then
        ' Get Username to send to
        User = RetrieveUser(lstUsers.Text)
        ' RetrieveUser returns -1 if the user wasn't found
        If User = -1 Then
            Exit Sub
        End If
        ' format the message
        wsArray(User).SendData "t" & Chr(1) & txtSendMessage.Text
        ' Blank the input
        txtSendMessage.Text = ""
    
    ElseIf KeyCode = 13 And Shift = 1 Then
        
        ' Loop through the users.
        ' There's better ways of doing this
        For X = 0 To 100
            
            ' If there's a username listed for them
            If Users(X) <> "" Then
                
                'Send the message
                wsArray(X).SendData "t" & Chr(1) & txtSendMessage.Text
                
                ' Don't know why this needs to be
                ' in here to work - someone tell me?
                DoEvents
            End If
        Next X
        txtSendMessage.Text = ""
    End If

End Sub

Private Function RetrieveUser(UserName As String) As Integer
Dim X As Integer

    'Check to see if nothing was selected
    If UserName = "" Then
        
        'OK, nothing selected, let's see how full
        ' the list is!
        If lstUsers.ListCount = 0 Then
            
            'Nothing in the list, so return -1
            RetrieveUser = -1
            Exit Function
        End If
        
        'If there is something in the list, send it to
        ' the first one =)
        UserName = lstUsers.List(0)
    End If
    
    ' Count through the users
    For X = 0 To 100
        
        'Check username to see if it is the right one
        If Users(X) = UserName Then
        
            'Ok, this is our man, so let's return his
            ' winsock index
            RetrieveUser = X
            Exit Function
        End If
    Next X
    RetrieveUser = -1
End Function

Private Sub txtSendMessage_KeyPress(KeyAscii As Integer)
    'Let's get rid of the annoying beep =)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub wsArray_Close(Index As Integer)
    ' Let's cycle through the list, looking for their
    ' name
    For X = 0 To lstUsers.ListCount - 1
    
        ' Check to see if it matches
        If lstUsers.List(X) = Users(Index) Then
        
            ' It matches, so let's remove it form the
            ' list and the array
            Users(Index) = ""
            lstUsers.RemoveItem X
            
            Exit For
        End If
    Next X
End Sub

Private Sub wsArray_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String, CtrlChar As String
    wsArray(Index).GetData Data
    
    ' Our format for our messages is this:
    ' CtrlChar & chr(1) & <info>
    If InStr(1, Data, Chr(1)) <> 2 Then
    ' If the 2nd char isn't chr(1), we know we have a prob
    
        MsgBox "Unknown Data Format: " & vbCrLf & _
                Data, vbCritical, "Error receiving"
        ' Make sure to leave the sub so it doesn't
        ' try to process the invalid info!
        Exit Sub
    End If
    
    'Retrieve First Character
    CtrlChar = Left(Data, 1)
  
    'Make sure to trim it, and chr(1), off
    Data = Mid(Data, 3)
    
    ' Check what it is, without regard to case
    Select Case LCase(CtrlChar)
        
        'This is to display a msgbox.
        ' I didn't enable the ability on the clients --
        '  for obvious reasons ;)
        Case "m"
            MsgBox Data, vbInformation, "Msg from client"
        
        'This is to change the caption.
        ' I didn't enable the ability on the clients --
        '  for obvious reasons ;)
        Case "c"
            Me.Caption = "Server - " & Data
        
        'This is their "login" key
        Case "u"
        
            'Add their name to the list
            lstUsers.AddItem Data
            
            'Add their name to the array
            Users(Index) = Data
            
            ' We need to remember that both
            ' the winsock index and the user array
            ' index correspond.  So you can find a
            ' users name by going "Users(<winsock index>)"
            ' or you can find the winsock index with
            ' a text name by cycling through the array.
            ' That's what the function "RetrieveUser"
            ' does - gets their winsock index from their
            ' username
            
        ' If all else fails, print it to output =)
        Case Else
            txtReceived.SelStart = Len(txtReceived.Text)
            txtReceived.SelText = Data & vbCrLf
    End Select
End Sub

Private Sub wsArray_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    ' This sets the "cursor" to the end of the textbox
    txtErrors.SelStart = Len(txtErrors.Text)
    
    ' This inserts the error message at the "cursor"
    txtErrors.SelText = "wsArray(" & Index & ") - " & Number & " - " & Description & vbCrLf
    
    ' Close it =)
    wsArray(Index).Close
    
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
    Index = FindOpenWinsock
    
    ' Accept the request using the created winsock
    wsArray(Index).Accept requestID
End Sub

Private Sub wsListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    ' This sets the "cursor" to the end of the textbox
    txtErrors.SelStart = Len(txtErrors.Text)
    
    ' This inserts the error message at the "cursor"
    txtErrors.SelText = "wsListen - " & Number & " - " & Description & vbCrLf
End Sub

Private Function FindOpenWinsock()
Static LocalPorts As Integer  ' Static keeps the
                              ' variable's state
    
    For X = 0 To wsArray.UBound
        If wsArray(X).State = 0 Then
            
            ' We found one that's state is 0, which
            '  means "closed", so let's use it
            FindOpenWinsock = X
            
            ' make sure to leave function
            Exit Function
        End If
    Next X

    '  OK, none are open so let's make one
    Load wsArray(wsArray.UBound + 1)
    
    '  Let's make sure we don't get conflicting local ports
    LocalPorts = LocalPorts + 1
    wsArray(wsArray.UBound).LocalPort = wsArray(wsArray.UBound).LocalPort + LocalPorts
    
    '  and then let's return it's index value
    FindOpenWinsock = wsArray.UBound

End Function

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormMain 
   Caption         =   "Microtik API Test"
   ClientHeight    =   5160
   ClientLeft      =   1860
   ClientTop       =   2085
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6615
   Begin VB.TextBox txtOut 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1560
      Width           =   6615
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "/system/resource/print"
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton btnDisconnect 
      Caption         =   "Disc"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Command:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Pass"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "User"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private inbuf1() As Byte ' 1st stage inbound data
Private inbuf2$ ' 2nd stage ( decoded ) inbound data
Private bErr As Boolean
Private Enum MyState_e
    CONNECTING
    WAITING_KEY
    AUTHENTICATING
    CONNECTED
End Enum
Dim MyState As MyState_e
Dim md5 As New CMD5

Private Sub Out(ByVal s$)
    txtOut.text = txtOut.text & s & vbCrLf
End Sub

Private Function GetReply$()
    Dim tmp&
    tmp = InStr(inbuf2, vbLf)
    If 0 = tmp Then Exit Function
    GetReply = Left(inbuf2, tmp - 1)
    inbuf2 = Mid(inbuf2, tmp + 1)
End Function

Private Sub btnConnect_Click()
    bErr = False
    ws.Protocol = sckTCPProtocol
    MyState = CONNECTING
    Out "(Connecting)"
    ws.Connect txtIP.text, 8728
End Sub

Private Function Hexlify$(ByVal s$)
    Dim i&, n&
    For i = 1 To Len(s)
        n = Asc(Mid(s, i, 1))
        Hexlify = Hexlify & LCase(Right("0" & Hex(n), 2))
    Next
End Function

Private Function Unhexlify$(ByVal s$)
    Dim i&, n&
    For i = 1 To Len(s) Step 2
        n = CLng("&H" & Mid(s, i, 2))
        Unhexlify = Unhexlify & Chr(n)
    Next
End Function

Private Sub btnSend_Click()
SendCommand txtCommand.text
End Sub

Private Sub Form_Resize()
    txtOut.Width = Me.ScaleWidth - txtOut.Left
    txtOut.Height = Me.ScaleHeight - txtOut.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bErr = True
    ws.Close
    End
End Sub

Private Sub ws_Connect()
    MyState = WAITING_KEY
    Out "(connected - sending /login)"
    SendCommand ("/login")
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    If bErr Then Exit Sub
    Dim ar() As Byte, i&, inbuf1_len&
    ReDim ar(0 To bytesTotal - 1)
    ws.GetData ar, vbByte, bytesTotal
    On Error Resume Next
    Err.Clear
    inbuf1_len = UBound(inbuf1)
    If Err Then
        inbuf1_len = 0
        ReDim inbuf1(0)
    Else
        inbuf1_len = inbuf1_len + 1
    End If
    On Error GoTo 0
    If inbuf1_len > 0 Then
        Dim off&
        off = inbuf1_len
        ReDim Preserve inbuf1(inbuf1_len + bytesTotal - 1)
        For i = 0 To bytesTotal - 1
            inbuf1(off + i) = ar(i)
        Next
    Else
        inbuf1 = ar
    End If
    Dim WordLen&, StartIdx&, Idx&
    StartIdx = 0
    Do While True
        Idx = StartIdx
        WordLen = CalcWordLen(inbuf1, Idx)
        If WordLen < 0 Then
            Exit Do
        End If
        If WordLen = 0 Then
            SentenceArrived (inbuf2)
            inbuf2 = ""
        Else
            If inbuf1(Idx) = Asc("=") Then
                inbuf2 = inbuf2 & " "
            End If
            For i = 0 To WordLen - 1
                inbuf2 = inbuf2 & Chr(inbuf1(Idx + i))
            Next
        End If
        StartIdx = Idx + WordLen
    Loop
End Sub

Private Sub SentenceArrived(ByVal sent$)
    Dim ar$(), tmp$, i&, re As New RegExp, chal$
    Out "I " & Replace(sent & " <<EOS>>", " ", vbCrLf & "I ")
    Select Case MyState
    Case CONNECTING
        ' this shouldn't happen!
        Out "(connected w/ packet - sending /login)"
        SendCommand "/login"
        MyState = WAITING_KEY
    Case WAITING_KEY
        Out "(got key sending credentials)"
        re.Global = False
        re.Pattern = "^!done =ret=([a-fA-F0-9]+)$"
        On Error Resume Next
        Err.Clear
        If re.Test(sent) Then
            chal = re.Replace(sent, "$1")
        Else
            bErr = True
        End If
        If Err Or bErr Then
            Out "Got error response to initial /login"
            bErr = True
            Exit Sub
        End If
        On Error GoTo 0
        Out md5.MD5Hex(Chr(0))
        Out md5.MD5Hex(Chr(0) & txtPass.text)
        tmp = md5.MD5Hex(Chr(0) & txtPass.text & Unhexlify(chal))
        tmp = "/login =name=" & txtUser.text & " =response=00" & tmp
        SendCommand tmp
        MyState = AUTHENTICATING
    Case AUTHENTICATING
        If Left(sent, 5) <> "!done" Then
            Out "Authentication failure"
            bErr = True
            Exit Sub
        End If
        MyState = CONNECTED
    Case CONNECTED
        ' do nothing - we're already sending output to text box
    End Select
End Sub

Private Function CalcWordLen&(ByRef ar() As Byte, ByRef Idx&)
    Dim tmp&
    CalcWordLen = -1 ' return error by default
    
    ' is there a single byte to begin decoding?
    If Idx > UBound(inbuf1) Then Exit Function
    
    tmp = inbuf1(Idx)
    If tmp < &H80 Then
        Idx = Idx + 1
    ElseIf tmp < &HC0& Then
        ' are there enough bytes to fully decode length?
        If (Idx + 1) > UBound(inbuf1) Then Exit Function
        tmp = tmp And Not &HC0&
        tmp = tmp * 256 + inbuf1(Idx + 1)
        Idx = Idx + 2
    ElseIf tmp < &HE0& Then
        ' are there enough bytes to fully decode length?
        If (Idx + 2) > UBound(inbuf1) Then Exit Function
        tmp = tmp And Not &HE0&
        tmp = tmp * 256 + inbuf1(Idx + 1)
        tmp = tmp * 256 + inbuf1(Idx + 2)
        Idx = Idx + 3
    ElseIf tmp < &HF0& Then
        ' are there enough bytes to fully decode length?
        If (Idx + 3) > UBound(inbuf1) Then Exit Function
        tmp = tmp And Not &HF0&
        tmp = tmp * 256 + inbuf1(Idx + 1)
        tmp = tmp * 256 + inbuf1(Idx + 2)
        tmp = tmp * 256 + inbuf1(Idx + 3)
        Idx = Idx + 4
    ElseIf tmp < &HF8& Then
        ' are there enough bytes to fully decode length?
        If (Idx + 4) > UBound(inbuf1) Then Exit Function
        tmp = inbuf1(Idx + 1)
        tmp = tmp * 256 + inbuf1(Idx + 2)
        tmp = tmp * 256 + inbuf1(Idx + 3)
        tmp = tmp * 256 + inbuf1(Idx + 4)
        Idx = Idx + 5
    Else
        bErr = True
        Out "ERROR: Received reserved control byte (0x" & Hex(inbuf1(0)) & ") - connection is indeterminate"
        Exit Function
    End If
    ' is the entire buffer here for which we are asking the length?
    If (Idx + tmp - 1) > UBound(inbuf1) Then Exit Function
    CalcWordLen = tmp
End Function

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Out "ws_Error: " & Description
    bErr = True
End Sub

Private Sub EncodeWord(ByRef buf() As Byte, ByVal sWord$)
    Dim DataLen&, HdrLen&, Idx&, tmp&, i&
    DataLen = Len(sWord)
    'If 0 = DataLen Then
    '    DataLen = 1
    '    sWord = Chr(0)
    'End If
    
    On Error Resume Next
    Err.Clear
    Idx = UBound(buf)
    If Err Then
        On Error GoTo 0
        ReDim buf(0)
        Idx = 0
    Else
        Idx = Idx + 1
    End If
    On Error GoTo 0
    If DataLen < &H80& Then
        HdrLen = 1
        ReDim Preserve buf(0 To Idx + HdrLen + DataLen - 1)
        buf(Idx) = DataLen
    ElseIf DataLen < &H4000& Then
        HdrLen = 2
        ReDim Preserve buf(0 To Idx + HdrLen + DataLen - 1)
        tmp = DataLen Or &H8000&
        buf(Idx + 1) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 0) = tmp
    ElseIf DataLen < &H200000 Then
        HdrLen = 3
        ReDim Preserve buf(0 To Idx + HdrLen + DataLen - 1)
        tmp = DataLen Or &HC00000
        buf(Idx + 2) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 1) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 0) = tmp
    ElseIf DataLen < &H10000000 Then
        HdrLen = 4
        ReDim Preserve buf(0 To Idx + HdrLen + DataLen - 1)
        tmp = DataLen Or &HE0000000
        buf(Idx + 3) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 2) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 1) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 0) = tmp
    Else
        HdrLen = 5
        ReDim Preserve buf(0 To Idx + HdrLen + DataLen - 1)
        buf(Idx) = &HF0&
        buf(Idx + 4) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 3) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 2) = tmp And &HFF&
        tmp = tmp \ 256
        buf(Idx + 1) = tmp
    End If
    Idx = Idx + HdrLen - 1 ' Idx is one-less, to make math easier below
    For i = 1 To DataLen
        buf(Idx + i) = Asc(Mid$(sWord, i, 1))
    Next
End Sub

Private Sub SendCommand(ByVal sCmd$)
    Dim ar$(), i&, buf() As Byte
    Out "O " & Replace(sCmd & " <<EOS>>", " ", vbCrLf & "O ")
    ar = Split(sCmd, " ")
    For i = 0 To UBound(ar)
        EncodeWord buf, ar(i)
    Next
    EncodeWord buf, ""
    ws.SendData buf
End Sub

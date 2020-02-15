VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_MikroTik 
   Caption         =   "Conexión a Puntos de Acceso"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   ScaleHeight     =   6165
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton GetData 
      Caption         =   "Obtener datos"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      MultiLine       =   -1  'True
      PasswordChar    =   "·"
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TextCommand 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton btnDisconnect 
      Caption         =   "Desconectar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtOut 
      Height          =   3255
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2775
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SSID"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Comando"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "Form_MikroTik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''CÓDIGO API````````````````````
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

''''''''''''''''''''''FIN CÓDIGO API''''''''''''''''''

'''''''''''''''''''''''''''SELECCIÓND DE LOS DATOS'''''''''''''''''''''''''''''
Private Sub GetData_Click()
Grid1.Rows = 10  'filas de la tabla
Grid1.Clear
Grid1.TextMatrix(0, 0) = "Parámetro"    'Títulos de las columnas
Grid1.TextMatrix(0, 1) = "Valor actual"
Grid1.Visible = True 'Que se muestre la tabla de resultados sólo cuando se haga click en el botón
Dim sString As String
Dim fileNum As Integer
fileNum = FreeFile
Open App.Path & "\mensajeAPI.txt" For Input As #fileNum  'Abre el fichero donde están los datos
Dim y As Integer
Do While Not EOF(fileNum)       'Leer fichero
Line Input #fileNum, sString
If Mid(sString, 4, 4) = "mac-" Then

y = InStrRev(sString, "=")
Grid1.TextMatrix(1, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(1, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 4) = "rx-r" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(2, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(2, 1) = Mid(sString, y + 1)


End If
If Mid(sString, 4, 4) = "tx-r" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(3, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(3, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 4) = "pack" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(4, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(4, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 4) = "upti" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(5, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(5, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 4) = "last" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(6, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(6, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 16) = "signal-strength=" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(7, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(7, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 8) = "signal-t" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(8, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(8, 1) = Mid(sString, y + 1)
End If
If Mid(sString, 4, 16) = "signal-strength-" Then
y = InStrRev(sString, "=")
Grid1.TextMatrix(9, 0) = Mid(sString, 4, y - 4)
Grid1.TextMatrix(9, 1) = Mid(sString, y + 1)
End If
Loop
Grid1.ColAlignment(1) = 2
Close fileNum
End Sub

Private Sub Form_Load()
Grid1.Cols = 2
Grid1.ColWidth(0) = 3000
Grid1.ColWidth(1) = 3000
End Sub


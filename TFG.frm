VERSION 5.00
Begin VB.Form Form_Inicio 
   BackColor       =   &H8000000D&
   Caption         =   "Inicio"
   ClientHeight    =   3210
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Seleccione una opción para comenzar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7080
      TabIndex        =   0
      Top             =   3360
      Width           =   5055
      Begin VB.CommandButton abrirAPI_Mikrotik 
         Caption         =   "Conectarse a un punto de acceso"
         Height          =   735
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton NuevoProyecto 
         Caption         =   "Comenzar un nuevo proyecto"
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu Opcion1 
         Caption         =   "Opción 1"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Herramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu CrearMapa 
         Caption         =   "Crear nuevo mapa"
      End
      Begin VB.Menu Herramienta2 
         Caption         =   "Herramienta 2"
      End
   End
End
Attribute VB_Name = "Form_Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub abrirAPI_Mikrotik_Click()
Form_MikroTik.Show
End Sub




' Private xo!, yo! ' Orig X&Y of Mouse on Drag Object
' Private Sub AccessPoint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  ' -- Save Orig X&Y of Obj..
  ' xo! = x: yo! = y
  ' If Shift = 0 Then AccessPoint.DragMode = 0: Exit Sub
  ' AccessPoint.DragMode = 1
 
' End Sub
 
' Private Sub Form1_DragDrop(Source As Control, x As Single, y As Single)
  ' If Source.Name = "AccessPoint" Then AccesPoint.Move x - xo!, y - yo!: AccesPoint.DragMode = 0
 
' End Sub



Private Sub Command1_Click()

Dim RutaMapa1, RutaMapa2, RutaMapa3 As String   'esta declaracion debes ponerla al principio del codigo de tu form (no en el evento click del command. Notaras que es una variable publica por lo que debe estar disponible en cualquier momento)

With CommonDialog1
     .DialogTitle = "SELECCIONE UN MAPA"
     .Filter = "Imagenes (*.jpg, *.bmp, *.gif, *.png, *.wmf) |*.jpg; *.bmp; *.gif; *.png; *.wmf|"
     .ShowOpen
End With

If CommonDialog1.FileName <> "" Then

    If picMap.Picture = 0 Then
        picMap.Picture = LoadPicture(CommonDialog1.FileName)
        RutaMapa1 = CommonDialog1.FileName  'FileName trae, ademas del nombre del archivo, la ruta completa de alojamiento del archivo
    ElseIf Picture2.Picture = 0 Then
        Picture2.Picture = LoadPicture(CommonDialog1.FileName)
        RutaMapa2 = CommonDialog1.FileName
    ElseIf Picture3.Picture = 0 Then
        Picture3.Picture = LoadPicture(CommonDialog1.FileName)
        RutMapa3 = CommonDialog1.FileName
    End If
     
Else
     MsgBox "No seleccionó ninguna imagen"
     RutaImagen = ""
End If

End Sub

Private Sub CrearMapa_Click()

Dim AbrirGenCDB

AbrirGenCDB = Shell("C:\Archivos de programa\GenCDB v4.0\GenCDB.EXE", 1)

End Sub




Private Sub FiltroDatos_Click()
Dim aux As String
Dim aux2 As String
Dim nombre As String
Dim poblacion As String
nombre = Text1.text
poblacion = Text2.text
Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        aux = Adodc1.Recordset.Fields("NombreVendedor")
        aux2 = Adodc1.Recordset.Fields("Poblacion")
        If nombre = aux And poblacion = aux2 Then
            MsgBox "El grupo que busca es " & nombre & " " & poblacion & " ", vbOKOnly
        Exit Sub
        Else
        Adodc1.Recordset.MoveNext
        End If
        Loop
        MsgBox ("La persona no existe")
        Adodc1.Recordset.MoveFirst
        
End Sub


Private Sub NuevoProyecto_Click()
Form_Mapas.Show
End Sub


'Private Sub Form_Load()
'SSTab1.Tab = 0 'Aparece por defecto la primera pestaña seleccionada
'SaveSizes
'PinColorIndex = -1
'NewPinIndex = -1
'End Sub



Private Sub Salir_Click()

If MsgBox("¿Está seguro de que quiere salir de la aplicación?", vbQuestion + vbYesNo, "Cerrar aplicación") = vbYes Then
    End
End If

End Sub

''''''''''''''''''''''''''''''''''''''
'API GOOGLE MAPS'
''''''''''''''''''''''''''''''''''''''

'Private Type ControlPositionType
    'Left As Single
    'Top As Single
    'Width As Single
    'Height As Single
    'FontSize As Single
'End Type

'Private m_ControlPositions() As ControlPositionType
'Private m_FormWid As Single
'Private m_FormHgt As Single

'Private Sub SaveSizes()
'Dim i As Integer
'Dim ctl As Control
' Save the controls' positions and sizes.
'ReDim m_ControlPositions(1 To Controls.Count)
'i = 1
'For Each ctl In Controls
    'With m_ControlPositions(i)
        'If TypeOf ctl Is Line Then
            '.Left = ctl.X1
            '.Top = ctl.Y1
            '.Width = ctl.X2 - ctl.X1
            '.Height = ctl.Y2 - ctl.Y1
        'Else
            '.Left = ctl.Left
            '.Top = ctl.Top
            '.Width = ctl.Width
            '.Height = ctl.Height
            'On Error Resume Next
            '.FontSize = ctl.Font.Size
            'On Error GoTo 0
        'End If
    'End With
    'i = i + 1
'Next ctl
' Save the form's size.
'm_FormWid = ScaleWidth
'm_FormHgt = ScaleHeight
'End Sub

Private Sub BuscarDireccion_Click()
Dim street As String
Dim city As String
Dim state As String
Dim zip As String
Dim queryAddress As String
queryAddress = "http://maps.google.com/maps?q="
' build street part of query string
If txtStreet.text <> "" Then
    street = txtStreet.text
    queryAddress = queryAddress & street + "," & "+"
End If
' build city part of query string
If txtCity.text <> "" Then
    city = txtCity.text
    queryAddress = queryAddress & city + "," & "+"
End If
' build state part of query string
If txtState.text <> "" Then
    state = txtState.text
    queryAddress = queryAddress & state + "," & "+"
End If
' build zip code part of query string
If txtZipCode.text <> "" Then
    zip = txtZipCode.text
    queryAddress = queryAddress & zip
End If
' pass the url with the query string to web browser control
'WebBrowser1.Navigate queryAddress
End Sub

Private Sub BuscarCoordenadas_Click()
If txtLat.text = "" Or txtLong.text = "" Then
    MsgBox "Supply a latitude and longitude value.", "Missing Data"
End If
Dim lat As String
Dim lon As String
Dim queryAddress As String
queryAddress = "http://maps.google.com/maps?q="
If txtLat.text <> "" Then
    lat = txtLat.text
    queryAddress = queryAddress & lat + "%2C"
End If
' build longitude part of query string
If txtLong.text <> "" Then
    lon = txtLong.text
    queryAddress = queryAddress & lon
End If
'WebBrowser1.Navigate queryAddress
End Sub


'Private Sub Form_Resize()
'ResizeControls
'End Sub

'Private Sub ResizeControls()
'Dim i As Integer
'Dim ctl As Control
'Dim x_scale As Single
'Dim y_scale As Single
' Don't bother if we are minimized.
'If WindowState = vbMinimized Then Exit Sub
'Get the form's current scale factors.
'x_scale = ScaleWidth / m_FormWid
'y_scale = ScaleHeight / m_FormHgt
' Position the controls.
'i = 1
'For Each ctl In Controls
    'With m_ControlPositions(i)
        'If TypeOf ctl Is Line Then
            'ctl.X1 = x_scale * .Left
            'ctl.Y1 = y_scale * .Top
            'ctl.X2 = ctl.X1 + x_scale * .Width
            'ctl.Y2 = ctl.Y1 + y_scale * .Height
        'Else
            'ctl.Left = x_scale * .Left
            'ctl.Top = y_scale * .Top
            'ctl.Width = x_scale * .Width
            'If Not (TypeOf ctl Is ComboBox) Then
                ' Cannot change height of ComboBoxes.
                'ctl.Height = y_scale * .Height
            'End If
            'On Error Resume Next
            'ctl.Font.Size = y_scale * .FontSize
            'On Error GoTo 0
        'End If
    'End With
    'i = i + 1
'Next ctl
'End Sub





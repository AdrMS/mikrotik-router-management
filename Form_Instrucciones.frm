VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form_Instrucciones 
   BackColor       =   &H8000000D&
   Caption         =   "Instrucciones"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Texto_Instrucciones 
      Height          =   6135
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10821
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form_Instrucciones.frx":0000
   End
End
Attribute VB_Name = "Form_Instrucciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Texto_Instrucciones.FileName = "C:\Documents and Settings\Administrador\Escritorio\TFG\Instrucciones.rtf"
End Sub

Private Sub Form_Resize()

'Primer y segundo parámetro es el valor Left y Top

'Parámetro 3 y 4, el ancho y alto del text _
 que en este caso es el ancho y alto del formulario

Texto_Instrucciones.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


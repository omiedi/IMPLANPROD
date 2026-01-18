VERSION 5.00
Begin VB.Form QR_CON_API_INI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR Code"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdFile 
      Caption         =   "Desde Archivo"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtFile 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4335
   End
   Begin VB.CommandButton CmdCodificTxt 
      Caption         =   "Codificar Texto"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton CmdDecodificar 
      Caption         =   "Decodificar Imagen"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdURL 
      Caption         =   "Desde URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox TxtUrl 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "http://qrcode.es/wp-content/uploads/2007/09/28w.jpg"
      Top             =   3960
      Width           =   4335
   End
   Begin VB.CommandButton CmdWebCam 
      Caption         =   "WebCam"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3810
      Left            =   6000
      Picture         =   "QR_CON_API_INI.frx":0000
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   0
      Width           =   3810
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MaxLength       =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "QR_CON_API_INI.frx":2DEA2
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Alto"
      Height          =   255
      Left            =   3195
      TabIndex        =   13
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Ancho"
      Height          =   255
      Left            =   4545
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "QR_CON_API_INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Autor:  Leandro Ascierto
'Web:    leandroascierto.com
'Date:   09/09/2011
'----------------------------------
Dim cQrCode As QR_CON_APICLS
Dim Ancho As Single
Dim Alto As Single
Dim Porcentaje As Single


'Variable de tipo IpictureDisp para a almacenar la imagen
Dim imagen As IPictureDisp


Private Sub CmdCodificTxt_Click()
    Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, Picture1.ScaleWidth, Picture1.ScaleHeight)
'    SavePicture Picture1.Picture, "C:\Imagen.BMP"
    
    If Picture1.Picture Is Nothing Then MsgBox "Error!"
    'Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, 200, 200, "UTF-8", "L", vbRed, vbBlue, 3)
End Sub

Private Sub CmdDecodificar_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromPicture(Picture1.Picture, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdFile_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromFile(TxtFile.Text, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdURL_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromUrl(TxtUrl.Text, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub

Private Sub CmdWebCam_Click()
    QR_CON_API1.Show , Me
End Sub

Private Sub Command1_Click()
imprimir
End Sub

Private Sub Form_Load()
    Set cQrCode = New QR_CON_APICLS
    TxtFile.Text = App.Path & "\casquillo_gorra.jpg"
    Text2 = 20
    Text3 = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cQrCode = Nothing
End Sub



Private Sub Establecer_Imagen(Obj As Object, RutaImg As String)
'************************************************* **********************
'* Este procedimiento carga la imagen en la variable imagen
'* y calcula las dimensiones y el porcentaje del ancho y alto de la imagen
'* que despues mediante el procedimiento CentrarPicture se dibuja _
* con el metodo PaintPicture
'************************************************* **********************
With Obj
'Borramos el contenido del picture
.Cls
'Para que al repintar no sepierda el gráfico
.AutoRedraw = True
End With

'Cargamos la imagen seleccionada del FileListBox
Set imagen = LoadPicture(RutaImg)

'Almacenamos el ancho y alto de la imagen
Ancho = imagen.Width
Alto = imagen.Height

'Sacamos el porcentaje para reescalar la imàgen
If Ancho > Obj.Width Or Alto > Obj.Height Then
If Ancho > Alto Then
Porcentaje = (Obj.Width * 100) / Ancho
Else
Porcentaje = (Obj.Height * 100) / Alto
End If

Call CentrarPicture(Obj)

Exit Sub

End If

If Ancho <= Obj.Width Or Alto <= Obj.Height Then
If Ancho > Alto Then
Porcentaje = (Obj.Width * 100) / Ancho
Else
Porcentaje = (Obj.Width * 100) / Alto
End If

End If

Call CentrarPicture(Obj)

End Sub

Public Sub CentrarPicture(Obj As Object)

Dim centro1 As Single
Dim centro2 As Single

'Sacamos el porcentaje del ancho y el alto
Ancho = ((Ancho * Porcentaje) / 100) / 0.99
Alto = ((Alto * Porcentaje) / 100) / 0.99

'La posicion X e y de la imagen dentro del Picture
centro1 = (Obj.Width - Ancho) / 2
centro2 = (Obj.Height - Alto) / 2

'Lo dibujamos
Obj.PaintPicture imagen, centro1, centro2, Ancho, Alto

'eliminamos la variable imagen
Set imagen = Nothing

End Sub
Sub imprimir()
Printer.Print
VB.Printer.Scale (0, 0)-(200, 500)
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = 50: Printer.CurrentY = 100
'for the position of the picture on paper
With Picture1
.AutoRedraw = True
.FontSize = 14
Picture1.Print ""

.Picture = .Image
Dim anch As Double
Dim alt As Double
anch = Val(Text2): If anch <= 0 Then anch = 20
alt = Val(Text3): If alt <= 0 Then alt = 10
Printer.PaintPicture Picture1.Picture, 2, 2, anch, alt


'En esta instruccion puedes indicarle tambien el tamaño
Printer.EndDoc
End With
End Sub

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlImagen 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdInsertarImagen 
      Caption         =   "Insertar Imagen"
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtfTexto 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12726
      _Version        =   393217
      TextRTF         =   $"FRMRICHTEXTBOX.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const WM_PASTE = &H302

Private Sub cmdInsertarImagen_Click()
On Error GoTo ErrorImagen
Dim strRutaImagen As String
Dim imgLaImagen As Picture
Dim strContenidoPortapapeles As String
With cdlImagen
    .DialogTitle = "Insertar imagen"
    .CancelError = True
    .Filter = "Todos los archivos de imágenes|*.gif;*.jpg;*.jpe;*.bmp;*.png|Imágenes GIF (*.gif)|*.gif|Imágenes JPG (*.jpg, *.jpe)|*.jpg;*.jpe|Imágenes de mapas de bits (*.bmp)|*.bmp|Imágenes PNG (*.png)|*.png|Todos los archivos (*.*)|*.*"
    .FilterIndex = 1
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .ShowOpen
    strRutaImagen = .FileName
End With
Set imgLaImagen = LoadPicture(strRutaImagen)
strContenidoPortapapeles = Clipboard.GetText
Clipboard.Clear
Clipboard.SetData imgLaImagen
SendMessage rtfTexto.hWnd, WM_PASTE, 0, 0
Clipboard.Clear
Clipboard.SetText strContenidoPortapapeles
Exit Sub
ErrorImagen:
If Err.Number <> 32755 Then
    MsgBox "Error " & Err.Number & " " & Err.Description
    Exit Sub
End If
End Sub

Private Sub cmdInsertarImagen_Click_Click()

End Sub

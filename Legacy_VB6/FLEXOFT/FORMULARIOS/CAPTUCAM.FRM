VERSION 5.00
Begin VB.Form CAPTUCAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura Imagen desde WebCam"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3180
      Left            =   4440
      ScaleHeight     =   3120
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command10 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CAPTUCAM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Refresca Imagen de Camara"
         Top             =   1520
         Width           =   640
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "CAPTUCAM.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ayuda en Línea"
         Top             =   810
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "CAPTUCAM.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit"
         Top             =   110
         Width           =   650
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CAPTUCAM.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   2400
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "CAPTUCAM.frx":08A8
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Captura Imagen"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Genera BMP"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox chkProg 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Genera JPG"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vista Previa"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
End
Attribute VB_Name = "CAPTUCAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal c As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Long
Dim retval As Long
Dim loadStr As String
Dim hwndc As Long

'---->CAPTURAR LA IMAGEN DE LA CAM A JPEG<----
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
'----> CÓDIGO CAPTURA DE FOTOS CON WEBCAM <----
'Declaraciones:
Const ws_visible = &H10000000
Const ws_child = &H40000000
Const WM_USER = 1024
Const WM_CAP_EDIT_COPY = WM_USER + 30
Const wm_cap_driver_connect = WM_USER + 10
Const wm_cap_set_preview = WM_USER + 50
Const wm_cap_set_overlay = WM_USER + 51
Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Const WM_CAP_SEQUENCE = WM_USER + 62
Const WM_CAP_SINGLE_FRAME_OPEN = WM_USER + 70
Const WM_CAP_SINGLE_FRAME_CLOSE = WM_USER + 71
Const WM_CAP_SINGLE_FRAME = WM_USER + 72
Const DRV_USER = &H4000
Const DVM_DIALOG = DRV_USER + 100
Const PREVIEWRATE = 30

Dim RUTA As String
Dim FORMATOIMAGEN As String



Private Sub Command10_Click()

Picture3.Visible = True
Picture4.Visible = False

End Sub

Private Sub Command2_Click()
'Código que realiza la captura de la imagen:
temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
Set Picture4.Picture = Clipboard.GetData

'Initialize input path
loadStr = RUTA$

'Required by DIjpg.dll
SavePicture Picture4.Image, "C:\tmp.bmp"

'Save to JPEG
On Error Resume Next
retval = DIWriteJpg(loadStr, 10000, chkProg.Value) 'AQUI ES EL ERROR... dice QUE SE REQUIERE UN OBJETO
If Err > 0 Then
    MsgBox "Imposible Generar Archivo en Formato JPG. Falta Libreria Dijpg.dll"
    On Error GoTo 0
    Exit Sub
End If
On Error GoTo 0

If retval = 1 Then 'Success
Picture4.Picture = LoadPicture(loadStr)
Else 'An error occured
MsgBox "DIWriteJpg did not succeed"
End If

'Remove temporary file
Kill "C:\tmp.bmp"
End Sub


Private Sub hshQual_Change()
'txtQual.Text = Str(hshQual.Value)
End Sub
' Fin de la Captura
Private Sub Command1_Click()
'Código que realiza la captura de la imagen:

temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
Set Picture2.Picture = Clipboard.GetData

Picture4.Visible = False

End Sub

Private Sub Command3_Click()
'Código que realiza la captura de la imagen:
temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
Set Picture4.Picture = Clipboard.GetData

SavePicture Picture4.Image, RUTA$

End Sub

Private Sub Command4_Click()

Picture3.Visible = False
Select Case FORMATOIMAGEN
    Case "JPG": Call Command2_Click
    Case "BMP": Call Command3_Click
    Case Else: Call Command3_Click
End Select

End Sub

Private Sub Command5_Click()

Unload Me

End Sub

Private Sub Form_Load()
'Código que activa la captura de imágenes, se supone un formulario con 2 picture llamados "picture2" y "picture3")

CMD$ = Trim$(Command$)

hwndc = capCreateCaptureWindow("Ventana de Captura", ws_child Or ws_visible, 0, 0, Picture3.Width, Picture3.Height, Picture3.hwnd, 0)
If (hwndc <> 0) Then
temp = SendMessage(hwndc, wm_cap_driver_connect, 0, 0)
temp = SendMessage(hwndc, wm_cap_set_preview, 1, 0)
temp = SendMessage(hwndc, WM_CAP_SET_PREVIEWRATE, PREVIEWRATE, 0)
End If

If CMD$ <> "" Then
    RUTA = Trim(CMD$)
    POS = InStr(1, CMD$, Chr$(124))
    If POS > 0 Then
        RUTA = UCase(Mid(CMD$, 1, POS - 1))
        FORMATOIMAGEN = UCase(Mid(CMD$, POS + 1))
    End If
Else
    MsgBox "Faltan Datos de Ubicacion y Formato"
    Unload Me
End If

If FORMATOIMAGEN = "JPG" Then
    On Error Resume Next
    retval = DIWriteJpg("", 10000, chkProg.Value) 'AQUI ES EL ERROR... dice QUE SE REQUIERE UN OBJETO
    If Err > 0 Then
        MsgBox "Imposible Generar Archivo en Formato JPG. Falta Libreria Dijpg.dll"
        On Error GoTo 0
        Unload Me
    End If
    On Error GoTo 0
End If

End Sub



VERSION 5.00
Begin VB.Form QR_CON_API1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enfoca con la camara un QRCODE"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4560
      TabIndex        =   1
      Top             =   2565
      Width           =   4560
      Begin VB.CommandButton CmdCapture 
         Caption         =   "Capturar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "QR_CON_API1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CreateCaptureWindow (Picture1.hwnd)
    If ConnectDriver = False Then
        MsgBox "No se pudo iniciar la WebCam"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisconnectDriver
    DestroyCaptureWindow
End Sub

Private Sub CmdCapture_Click()
    Set QR_CON_API_INI.Picture1.Picture = capGetImage
    Unload Me
End Sub

Private Sub Picture1_Resize()
    Me.Width = Picture1.ScaleWidth + (Me.Width - Me.ScaleWidth)
    Me.Height = Picture1.Height + 1000
End Sub

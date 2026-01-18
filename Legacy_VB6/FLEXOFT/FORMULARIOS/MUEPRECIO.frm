VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MUEPRECIO 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Precios de Venta"
   ClientHeight    =   1620
   ClientLeft      =   7245
   ClientTop       =   4620
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4600
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "MUEPRECIO.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "MUEPRECIO.frx":0234
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "MUEPRECIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)

   Call APLICACIONSKINS(Me)

   '
End Sub

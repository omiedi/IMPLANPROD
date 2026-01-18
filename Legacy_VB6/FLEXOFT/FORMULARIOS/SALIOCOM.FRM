VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SALIOCOM 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Salida - Orden de Compra "
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPCIONES DE SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2460
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Ajuntar PDF's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   645
         TabIndex        =   5
         Top             =   1200
         Width           =   1545
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Envio por Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   285
         TabIndex        =   4
         Top             =   840
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   285
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   1620
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   3960
      Picture         =   "SALIOCOM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   2640
      Picture         =   "SALIOCOM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "SALIOCOM.frx":0614
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "SALIOCOM.frx":0848
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "SALIOCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check2_Click()
    If Check2.Value < 1 Then
       Check3.Value = 0
       Check3.Enabled = False
     Else
       Check3.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
  '
  Hide
  '
End Sub
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  Call APLICACIONSKINS(Me)
  '
End Sub

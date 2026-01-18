VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STD_FORM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox ENTRATEX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1785
      Width           =   1275
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "STD_FORM.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STD_FORM.frx":0234
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label PROCONTROL 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label MUESTRATEX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label LEYENDA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LEYENDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   2010
   End
End
Attribute VB_Name = "STD_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
End Sub

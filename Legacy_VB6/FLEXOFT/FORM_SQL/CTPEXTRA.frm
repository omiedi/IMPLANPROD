VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CTPEXTRA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Datos Extras Pedidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8115
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10185
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   30
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "OBSERVACIONES DE ENTREGA "
         Top             =   4200
         Width           =   2565
      End
      Begin VB.TextBox TXTOBS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Index           =   30
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "CTPEXTRA.frx":0000
         Top             =   4560
         Width           =   5085
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   32
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "CÓDIGO DE CLIENTE"
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox TXTOBS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   32
         Left            =   1920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   2205
      End
      Begin VB.Frame FRAME17 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Adjuntar Archivos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Index           =   7
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5175
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2550
            ItemData        =   "CTPEXTRA.frx":0007
            Left            =   120
            List            =   "CTPEXTRA.frx":0009
            TabIndex        =   7
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   3780
            TabIndex        =   6
            Top             =   2925
            Width           =   1125
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2520
            TabIndex        =   5
            Top             =   2925
            Width           =   1185
         End
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9360
         Picture         =   "CTPEXTRA.frx":000B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9360
         Picture         =   "CTPEXTRA.frx":0155
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Emite Listado"
         Top             =   840
         Width           =   600
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CTPEXTRA.frx":045F
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CTPEXTRA.frx":0693
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4680
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "CTPEXTRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    OPIMPRE = ""
    Hide
End Sub

Private Sub Command2_Click()
   '
   If Option3.Value = True Then
        OPIMPRE = "0"
     ElseIf Option4.Value = True Then
        OPIMPRE = "1"
     ElseIf Option5.Value = True Then
        OPIMPRE = "2"
     ElseIf Option6.Value = True Then
        OPIMPRE = "3"
   End If
   '
   If Option1.Value = True Then
        AGRUPA = "1"
     ElseIf Option2.Value = True Then
        AGRUPA = "2"
   End If
   '
   Hide
   '
End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

   Height = 4200
End Sub

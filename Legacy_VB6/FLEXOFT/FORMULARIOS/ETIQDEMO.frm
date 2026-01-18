VERSION 5.00
Begin VB.Form ETIQDEM1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7785
      Left            =   9770
      TabIndex        =   15
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7785
      Left            =   9870
      TabIndex        =   14
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4005
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   525
      Width           =   5055
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Distribuye:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   2360
         Width           =   1275
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disco 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3150
         TabIndex        =   12
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefax: 54-11-4343-1462   /   54-11-4473-1634"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   2775
         Width           =   3900
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " EMPRENDIMIENTOS - Maipú 26 - P.11 - 'C' - Bs.As."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   2600
         Width           =   5055
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   -105
         X2              =   4515
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "                         FLEXDEMO"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   220
         TabIndex        =   9
         Top             =   2050
         Width           =   2430
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "                         ARCHIVOS DE PROGRAMAS"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   220
         TabIndex        =   8
         Top             =   1860
         Width           =   4740
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "  2.- Ejecución:   Botón INICIO"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   1680
         Width           =   3585
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "                      ACEPTAR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   6
         Top             =   1485
         Width           =   2430
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "                         A:\SETUP.EXE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   250
         TabIndex        =   5
         Top             =   1305
         Width           =   2430
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "                         EJECUTAR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   250
         TabIndex        =   4
         Top             =   1125
         Width           =   3585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1.- Instalación: Botón INICIO"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   945
         Width           =   3585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FLEXDEMO - Sistema de Demostración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -105
         TabIndex        =   2
         Top             =   660
         Width           =   4425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones Informaticas Integrales  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   105
         Width           =   2100
      End
      Begin VB.Image Image1 
         Height          =   600
         Index           =   0
         Left            =   0
         Picture         =   "ETIQDEMO.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4305
      End
   End
End
Attribute VB_Name = "ETIQDEM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


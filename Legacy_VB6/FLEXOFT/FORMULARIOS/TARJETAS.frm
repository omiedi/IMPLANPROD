VERSION 5.00
Begin VB.Form TARJETAS 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   6825
   ClientLeft      =   -5190
   ClientTop       =   -1890
   ClientWidth     =   11490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2745
      Index           =   3
      Left            =   6300
      TabIndex        =   15
      Top             =   3650
      Width           =   4950
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   4635
         Begin VB.Label Label6 
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
            Height          =   750
            Index           =   3
            Left            =   2520
            TabIndex        =   17
            Top             =   105
            Width           =   2100
         End
         Begin VB.Image Image2 
            Height          =   570
            Index           =   3
            Left            =   0
            Picture         =   "TARJETAS.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4830
         End
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   0
         X2              =   4620
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Ernesto Pablo Bauer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2000
         TabIndex        =   19
         Top             =   960
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diaz Colodrero 3519 - (1431) Buenos Aires - Rep. Argentina  Telefax: 54-11-4541-1515     Phone: 15-4473-1634 "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   300
         TabIndex        =   18
         Top             =   2060
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2745
      Index           =   2
      Left            =   6300
      TabIndex        =   10
      Top             =   750
      Width           =   4950
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   4635
         Begin VB.Label Label6 
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
            Height          =   750
            Index           =   2
            Left            =   2520
            TabIndex        =   12
            Top             =   105
            Width           =   2100
         End
         Begin VB.Image Image2 
            Height          =   570
            Index           =   2
            Left            =   0
            Picture         =   "TARJETAS.frx":13166
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4830
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diaz Colodrero 3519 - (1431) Buenos Aires - Rep. Argentina  Telefax: 54-11-4541-1515     Phone: 15-4473-1634 "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   2060
         Width           =   4635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Ernesto Pablo Bauer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2000
         TabIndex        =   13
         Top             =   960
         Width           =   2745
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   0
         X2              =   4620
         Y1              =   2000
         Y2              =   2000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2745
      Index           =   1
      Left            =   1150
      TabIndex        =   5
      Top             =   3650
      Width           =   4950
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4635
         Begin VB.Label Label6 
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
            Height          =   750
            Index           =   1
            Left            =   2520
            TabIndex        =   7
            Top             =   105
            Width           =   2100
         End
         Begin VB.Image Image2 
            Height          =   570
            Index           =   1
            Left            =   0
            Picture         =   "TARJETAS.frx":262CC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4830
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diaz Colodrero 3519 - (1431) Buenos Aires - Rep. Argentina  Telefax: 54-11-4541-1515     Phone: 15-4473-1634 "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   2060
         Width           =   4635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Ernesto Pablo Bauer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2000
         TabIndex        =   8
         Top             =   960
         Width           =   2745
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   0
         X2              =   4620
         Y1              =   2000
         Y2              =   2000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2745
      Index           =   0
      Left            =   1150
      TabIndex        =   0
      Top             =   750
      Width           =   4950
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4635
         Begin VB.Label Label6 
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
            Height          =   750
            Index           =   0
            Left            =   2520
            TabIndex        =   4
            Top             =   105
            Width           =   2100
         End
         Begin VB.Image Image2 
            Height          =   570
            Index           =   0
            Left            =   0
            Picture         =   "TARJETAS.frx":39432
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4830
         End
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   0
         X2              =   4620
         Y1              =   2000
         Y2              =   2000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Ernesto Pablo Bauer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2000
         TabIndex        =   2
         Top             =   960
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diaz Colodrero 3519 - (1431) Buenos Aires - Rep. Argentina  Telefax: 54-11-4541-1515     Phone: 15-4473-1634 "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   2060
         Width           =   4635
      End
   End
End
Attribute VB_Name = "TARJETAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

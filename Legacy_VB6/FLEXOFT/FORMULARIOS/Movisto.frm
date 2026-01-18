VERSION 5.00
Begin VB.Form MOVISTO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Stock"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7245
      TabIndex        =   4
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ajustes y Transferencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   420
      TabIndex        =   3
      Top             =   4725
      Width           =   4635
      Begin VB.CommandButton Command11 
         Caption         =   "Transferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2625
         TabIndex        =   18
         Top             =   735
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Recuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   315
         TabIndex        =   17
         Top             =   735
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salidas de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   5775
      TabIndex        =   2
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   12
         Top             =   2205
         Width           =   4005
         Begin VB.CommandButton Command9 
            Caption         =   "Descartes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command8 
            Caption         =   "  Vale de Materiales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   15
            Top             =   420
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Expedición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   11
         Top             =   630
         Width           =   4005
         Begin VB.CommandButton Command7 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   14
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   13
            Top             =   420
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Entradas a Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   420
      TabIndex        =   1
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fabricación Propia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   6
         Top             =   2205
         Width           =   4005
         Begin VB.CommandButton Command5 
            Caption         =   "Materiales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   10
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Producción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Proveedores Externos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   5
         Top             =   630
         Width           =   4005
         Begin VB.CommandButton Command3 
            Caption         =   "Aprobación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Recepción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3465
      TabIndex        =   0
      Top             =   6825
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "Movisto.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   10920
   End
End
Attribute VB_Name = "MOVISTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Close
    End
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Shell LUBAS$ + "WINTABU BRSTOCK", vbNormalFocus
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    COMMAND3.Enabled = False
    Shell LUBAS$ + "WINTABU RECDEF", vbNormalFocus
    COMMAND3.Enabled = True
End Sub

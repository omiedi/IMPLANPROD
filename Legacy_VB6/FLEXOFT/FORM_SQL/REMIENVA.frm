VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form REMIENVA 
   BackColor       =   &H00808080&
   ClientHeight    =   8835
   ClientLeft      =   1920
   ClientTop       =   3285
   ClientWidth     =   14145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850.024
   ScaleMode       =   0  'User
   ScaleWidth      =   13344.34
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000C&
      Height          =   389
      Left            =   4680
      TabIndex        =   66
      Top             =   0
      Width           =   5295
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000C&
         Caption         =   "Agrupar por código y lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   68
         Top             =   120
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000C&
         Caption         =   "Agrupar por código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Filtrar Selecciòn Solo Pendientes de Pasar a Waldbot"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   65
      Top             =   480
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00808080&
      Caption         =   "IMPRESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12600
      TabIndex        =   62
      Top             =   5760
      Width           =   1455
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   63
         Text            =   " "
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   64
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
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
      Height          =   315
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos de Entrega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   53
      TabIndex        =   40
      Top             =   4800
      Width           =   13992
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   10320
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "Còd.Postal:"
         Top             =   600
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox TXTCODPOSTAL 
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
         Height          =   285
         Left            =   11520
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox LENTREGA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   44
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox LENTREGA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   43
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   600
         Width           =   1770
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   240
         Width           =   1770
      End
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "Facturas Asociadas:"
      Top             =   840
      Visible         =   0   'False
      Width           =   1802
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1965
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   480
      Width           =   5325
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12000
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "Fecha:"
      Top             =   480
      Width           =   810
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12930
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   480
      Width           =   1113
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12000
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10200
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "Numero de Remito:"
      Top             =   120
      Width           =   1770
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "Numero de Cliente:"
      Top             =   120
      Width           =   1770
   End
   Begin VB.CommandButton Command5 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13875
      TabIndex        =   28
      Top             =   120
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3020
      Left            =   0
      TabIndex        =   12
      Top             =   5760
      Width           =   6285
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Caption         =   "Transportista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2520
         TabIndex        =   21
         Top             =   210
         Width           =   3690
         Begin VB.CommandButton Command18 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox Text45 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   750
         End
         Begin VB.CommandButton Command19 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            TabIndex        =   47
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox LENTREGA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   27
            Top             =   675
            Width           =   3510
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00808080&
            Caption         =   "C.U.I.T."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Left            =   105
            TabIndex        =   25
            Top             =   975
            Width           =   1800
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   105
               TabIndex        =   26
               Top             =   210
               Width           =   1590
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00808080&
            Caption         =   "Patente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Left            =   1995
            TabIndex        =   23
            Top             =   975
            Width           =   1590
            Begin VB.TextBox Text18 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   105
               TabIndex        =   24
               Top             =   210
               Width           =   1380
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00808080&
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   105
            TabIndex        =   22
            Top             =   1560
            Width           =   3480
            Begin VB.CommandButton Command3 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   3120
               TabIndex        =   52
               ToolTipText     =   "Aplicar Domicilio de Transporte Como Domicilio de Entrega"
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox DOMITRAN 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Top             =   600
               Width           =   2980
            End
            Begin VB.TextBox DOMITRAN 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   2980
            End
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00808080&
         Caption         =   "Numero de Bultos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         TabIndex        =   19
         Top             =   945
         Width           =   2325
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   525
            TabIndex        =   20
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00808080&
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         TabIndex        =   17
         Top             =   210
         Width           =   2325
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   18
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00808080&
         Caption         =   "Conductor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2325
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   630
            TabIndex        =   15
            Top             =   735
            Width           =   1590
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   14
            Top             =   340
            Width           =   2115
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DNI:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   16
            Top             =   800
            Width           =   435
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   8520
      TabIndex        =   8
      Top             =   7140
      Width           =   5475
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Height          =   580
         Left            =   600
         TabIndex        =   55
         Top             =   980
         Visible         =   0   'False
         Width           =   350
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808080&
            Caption         =   "Option2"
            Height          =   200
            Left            =   60
            TabIndex        =   57
            Top             =   360
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808080&
            Caption         =   "Option1"
            Height          =   195
            Left            =   60
            TabIndex        =   56
            Top             =   140
            Width           =   255
         End
      End
      Begin VB.TextBox TXTFLETE 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   960
         TabIndex        =   53
         Text            =   " "
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   4995
         TabIndex        =   51
         ToolTipText     =   "Borra información de Caja de Texto"
         Top             =   1080
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   5145
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   5145
      End
      Begin VB.TextBox Text16 
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
         Index           =   4
         Left            =   210
         TabIndex        =   9
         Top             =   2500
         Width           =   3375
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Flete:"
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
         Height          =   225
         Index           =   7
         Left            =   0
         TabIndex        =   54
         Top             =   1020
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808080&
      Caption         =   "Valor Declarado de los Items Seleccionados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8550
      TabIndex        =   6
      Top             =   6495
      Width           =   4155
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3930
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00808080&
      Caption         =   "Orden  de Suministro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8550
      TabIndex        =   4
      Top             =   5760
      Width           =   4035
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   24
         TabIndex        =   5
         Top             =   300
         Width           =   3585
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00808080&
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12930
      TabIndex        =   2
      Top             =   6495
      Width           =   1095
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   885
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "REMIENVA.frx":0000
      Top             =   3000
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3765
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Doble Click Seleccióna Items a Imprimir"
      Top             =   960
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   6641
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuRevertirSeleccion 
         Caption         =   "Revertir Selección"
      End
   End
   Begin VB.Menu mnuConsultasyListados 
      Caption         =   "Consultas y Listados"
   End
End
Attribute VB_Name = "REMIENVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Option3_Click()
Call Carga_de_Pantalla(Text3, 2)
End Sub

Private Sub option4_Click()
  Call Carga_de_Pantalla(Text3, 1)
End Sub

Private Sub Check2_Click()
    Call Text3_Change
End Sub

Private Sub Check3_Click()
  
End Sub

Private Sub Command1_Click()
     Command1.Enabled = False
     If TRIM$(Text3) = "" Then
       Call COMUNI("Imposible Agregar Sin Remito Seleccionado Previamente")
       GoTo 99
     End If
     
     Call CARSELMESA
     
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     CMO1$ = "RP"
     
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND (substring(DOPRILAR,1,2) = 'RT' OR substring(DOPRILAR,1,2) = 'RP') "
    
     CONDI$ = CONDI$ + " AND Nume_Clie = " & XVALO(Text2)
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     FECHACBTE$ = Mid$(VALACT2$("SELREMIT"), 1, 8)
     Me.Text4 = FECHACBTE$

     MEXAGE$ = "Remito Existente en Grilla"
     If VALIDA_COD_GRID%(NUDOCU$, Me.GRID, 9, MEXAGE$) < 1 Then
        Text3.TEXT = NUDOCU$
        
     End If
    
    
99   Command1.Enabled = True

End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   VDEF$ = REGBLAN$("TRANSPOR")
   CONDI$ = "CODI_TRA > 0"
10 PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
   Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
   '
   OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
     If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
       Call MENSERR(24, "Falta Nombre del Transporte")
       VDEF$ = OBX$
       GoTo 10
     End If
     '
     Call BLOQARCH("TRANSPOR")
     PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
     ' Graba en Tabla SQL Transportes.
     COD1% = PROXI%
     DENO$ = TRIM$(Mid$(OBX$, 3, 32))
     RASO$ = TRIM$(Mid$(OBX$, 35, 48))
     DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
     LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
     Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
     PIVA% = Asc(Mid$(OBX$, 194, 1))
     TELE$ = TRIM$(Mid$(OBX$, 195, 48))
     FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
     mail$ = TRIM$(Mid$(OBX$, 291, 48))
     CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
     Dim TRANSPOR As ADODB.Recordset
     Set TRANSPOR = New ADODB.Recordset
     SQLX$ = "Select * from TRANSPORTES"
25   On Error Resume Next
     TRANSPOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     With TRANSPOR
          If COD1% > 0 Then
                 .AddNew
                 !CODI_TRA = COD1%
                 !DENO_TRA = DENO$
                 !RASO_TRA = RASO$
                 !DOMI_TRA = DOMI$
                 !LOCA_TRA = LOCA$
                 !CUIT_TRA = Cuit$
                 !PIVA_TRA = PIVA%
                 !TELE_TRA = TELE$
                 !FAXI_TRA = FAXI$
                 !MAIL_TRA = mail$
                 !CTAC_TRA = CTAC$
                 .Update
          End If
     End With
     '
   End If
99 Command16.Enabled = True

End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   NTRA = XVALO(Text45)
   If NTRA > 0 Then
     If NTRA < 32767 Then
       NTRAI% = NTRA
       'If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
       
       'MODIFICADO POR QUE EN GABAL SI GRABA EN UNA EMPRESA
       'NO LE DEJA EDITAR POR QUE NO ENCUENTRA EL ECOARCH
       CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
         'ASIGNO CON LA FUNICION SI NO ENCUENTRA EN SQL TRAE DE DAT
         '*********************************************************
         CUITRA$ = CUITTRANS$(NTRAI%)
         DENOTRA$ = DENOTRANS$(NTRAI%)
         RASOTRA$ = RASOTRANS$(NTRAI%)
         DOMITRA$ = DOMITRANS$(NTRAI%)
         LOCATRA$ = LOCATRANS$(NTRAI%)
         TELETRA$ = TELETRANS$(NTRAI%)
         FAXITRA$ = FAXITRANS$(NTRAI%)
         MAILTRA$ = MAILTRANS$(NTRAI%)
         CTACTRA$ = CTACTRANS$(NTRAI%)
         PIVATRA% = PIVATRANS%(NTRAI%)
         '
         'ARMO EL STRING
         '**************
         Call REPLA(VDEF$, MKI$(NTRAI%), 1, 2)
         Call REPLA(VDEF$, DENOTRA$, 3, 32)
         Call REPLA(VDEF$, RASOTRA$, 35, 48)
         Call REPLA(VDEF$, DOMITRA$, 83, 48)
         Call REPLA(VDEF$, LOCATRA$, 131, 48)
         Call REPLA(VDEF$, CUITRA$, 179, 15)
         Call REPLA(VDEF$, Chr$(PIVATRA%), 194, 1)
         Call REPLA(VDEF$, TELETRA$, 195, 48)
         Call REPLA(VDEF$, FAXITRA$, 243, 48)
         Call REPLA(VDEF$, MAILTRA$, 291, 48)
         Call REPLA(VDEF$, CTACTRA$, 339, 48)

         'VDEF$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
10       OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
         If OBX$ <> "" Then
           If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
             Call MENSERR(24, "Falta Nombre del Transporte")
             VDEF$ = OBX$
             GoTo 10
           End If
           '
           Call BLOQARCH("TRANSPOR")
           Call FILTERPUTRECORD("TRANSPOR", 1, MKI$(NTRAI%), OBX$)
           Call CIERRARCH("TRANSPOR")
           ' Edita y Graba en Tabla SQL Transportes.
           COD1% = NTRAI%
           DENO$ = TRIM$(Mid$(OBX$, 3, 32))
           RASO$ = TRIM$(Mid$(OBX$, 35, 48))
           DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
           LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
           Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
           PIVA% = Asc(Mid$(OBX$, 194, 1))
           TELE$ = TRIM$(Mid$(OBX$, 195, 48))
           FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
           mail$ = TRIM$(Mid$(OBX$, 291, 48))
           CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
'           SQLX$ = "Codi_Tra = " & COD1% & ""
           
          Dim RS1 As ADODB.Recordset
          Set RS1 = New ADODB.Recordset
          SQLX$ = "Select * from TRANSPORTES"
          SQLX$ = SQLX$ + " WHERE Codi_Tra = " & COD1%
30        On Error Resume Next
          RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 30
          End If
          On Error GoTo 0
        
          '
          With RS1
           Do While Not .EOF
              !DENO_TRA = Left$(TRIM$(DENO$), 128)
              !RASO_TRA = Left$(TRIM$(RASO$), 128)
              !DOMI_TRA = Left$(TRIM$(DOMI$), 128)
              !LOCA_TRA = Left$(TRIM$(LOCA$), 128)
              !CUIT_TRA = Left$(TRIM$(Cuit$), 32)
              !PIVA_TRA = PIVA%
              
              
              !TELE_TRA = Left$(TRIM$(TELE$), 64)
              !FAXI_TRA = Left$(TRIM$(FAXI$), 64)
              !MAIL_TRA = Left$(TRIM$(mail$), 64)
'              !MAIL_TRA = Left$(TRIM$(Text10(3)), 128)
              !CTAC_TRA = Left$(TRIM$(CTAC$), 128)
              .Update
              .MoveNext
           Loop
          End With
          RS1.Close
          Set RS1 = Nothing
           '
           Call Text45_Change
         End If
         GoTo 99
       End If
     End If
   End If
   '
   Call MENSERR(24, "Imposible Editar Datos Transportista")
   '
99 Command18.Enabled = True

End Sub


Private Sub Command19_Click()
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY CODI_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
        Text45 = TRIM$(Str$(NTRAN%))
   End If

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   TXTFLETE = ""
   Option1.Value = False
   Option2.Value = False
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Me.LENTREGA(0) = Me.DOMITRAN(3)
   Me.LENTREGA(1) = Me.DOMITRAN(4)

End Sub

Private Sub Command4_Click()
   
End Sub

Private Sub Command5_Click()
     Command5.Enabled = False
     Screen.MousePointer = 11
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     NPFIJOCTP& = XVALO(CONTROL("TRCTPAWL", "PFIJOCTP"))  'NUMERO DE PREFIJO
     NPFIJOCAMBACERES& = XVALO(CONTROL("TRCAPAWL", "PFIJOCAM"))  'NUMERO DE PREFIJO CAMBACERES
                        
     If NPFIJOCTP& < 1 Or NPFIJOCAMBACERES& < 1 Then
        Call COMUNI("Falta Configurar Prefijo de Remito de Waldbott para los Realizados en Conteplast\Realizar ésta Configuración Desde Módulo de Pedidos\Menú: Configuración > Ruta de conexión a Base de Datos Waldbott ")
        Exit Sub
     End If
     '
     

10  TituloLista$ = "Lista de Remitos Pendientes de Transferir a Waldbott"
    Campos$ = "Fecha/D8;Nro.Remito/A20;Cta./f8;Razón Social/A40;Pedido/f8;DOPRINUM/A0;VendContep/A16;O_COMPRA/A16;NroVendEnv/f10"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    Dim RS1 As New ADODB.Recordset
    
12  SQLX$ = " SELECT FechMov,DOPRILAR,Nume_Clie,RASO_CLI,NuPedCli,DOPRINUM,NAMEVEND,ORDENCOMPRA,NroVend_Envap"
    SQLX$ = SQLX$ + " FROM MOVISTOCTP  WITH(NOLOCK)"
    If Check2.Value > 0 Then
       SQLX$ = SQLX$ + " WHERE (PasoaWb < 1 OR PasoaWb IS NULL)"
    End If
    SQLX$ = SQLX$ + " GROUP BY FechMov,DOPRILAR,Nume_Clie,RASO_CLI,NuPedCli,DOPRINUM,NAMEVEND,ORDENCOMPRA,NroVend_Envap"
    SQLX$ = SQLX$ + " ORDER BY FechMov DESC,DOPRILAR DESC"
    
    
    Set RS1 = READSET(SQLX$)
              
    
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    Screen.MousePointer = 1
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(RESP_ZELE_VECT(2))
     If NUDOCU$ = "" Then GoTo 99
     FECHACBTE$ = TRIM$(RESP_ZELE_VECT(1))
     If NUDOCU$ = "" Then GoTo 10
     
     Me.Text4 = FECHACBTE$
     Me.Text3 = NUDOCU$
99   Command5.Enabled = True
     Screen.MousePointer = 1
      
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
       mnuImprimir.Visible = False
    End If
    Text23.TEXT = CONTROL$("", "COPIAREM")
    
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
       Command4.Visible = True
    End If

    Campos$ = "CI/F0;Ord./F4;Descripción/A32;Salida/F9.1;Ingreso/f9.1;Nro.Ped./F7;P.Unit./F12.2;N° Lote/A16;Remito/A16;Còdigo/A16;Descrip2/A0;ObservaItem/A0;ORDENCOMPRA/A0"
    Call CARGA_ENCABEZADO(Me.GRID, Campos$, Me)
End Sub

Private Sub GRID_DblClick()
    i& = GRID.MouseRow
    j& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = j&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, REG&)
    Call CalcutotalSeleccionado
End Sub

Private Sub mnAnulRemOficial_Click()
10 Campos$ = "Fecha/D8;Remito Oficial/A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES GENERADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(1)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   Call ACTUREGISTRO("MOVIOFI", "STATUS", "REMIOFI = '" & NROREMI$ & "'", 9, REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVIOFI", "REFERENCIA", "REMIOFI = '" & NROREMI$ & "'", "(Anul)" & RESP_ZELE_VECT(3), REGAF&, "NOMESS")

   GoTo 10
   
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
        i& = GRID.MouseRow
        j& = GRID.MouseCol
        If i& = 0 Or i& > GRID.Rows Then Exit Sub
        '
        On Error Resume Next
        GRID.Row = i&
        GRID.COL = j&
        On Error Resume Next
    
        REG& = i&
        If REG& < 1 Then
            Exit Sub
        End If
        '_____________________________


        NC = XVALO(Text2)
        CI1% = XVALO(GRID.TextMatrix(REG&, 1))
        Descrip$ = GRID.TextMatrix(REG&, 3)
        NUPEDCLIX& = XVALO(GRID.TextMatrix(REG&, 6))
        REMI$ = TRIM$(GRID.TextMatrix(REG&, 9))
        CODDX$ = TRIM$(GRID.TextMatrix(i&, 10))
        
        If NC < 1 Or CI1% < 1 Or REMI$ = "" Or NUPEDCLIX& < 1 Then Exit Sub
        
        CONDI$ = "NUPEDCLI = " & NUPEDCLIX&
        CONDI$ = CONDI$ + " AND DOPRILAR = '" & REMI$ & "'"
        CONDI$ = CONDI$ + " AND NUME_CLIE = " & NC
        CONDI$ = CONDI$ + " AND COD_INTE = " & CI1%
        OPX% = COMALTER("Opciones Disponibles\\Cancelar\Modificar Cantidad\Modificar Observaciòn")
        If OPX% < 2 Then Exit Sub
        
        If OPX% = 2 Then
             CANTI = XVALO(GRID.TextMatrix(REG&, 4))
             CantidadModificada$ = TRIM$(InputBox$("Cantidad: " & CODDX$, "Cantidad Actual", CANTI))
             If (CantidadModificada$) = StringEmpty Then Exit Sub
             If XVALO(CantidadModificada$) < 1 Then Exit Sub
             If XVALO(CantidadModificada$) > CANTI Then
               Call COMUNI("La Cantidad Modificada No Puede ser Mayor a la Original\Si ya Habia Modificado Reinicie la Aplicaciòn")
               Exit Sub
             End If
             'si paso todo modifica la cantidad
             GRID.TextMatrix(REG&, 4) = XVALO(CantidadModificada$)
             Exit Sub
        End If
        'si paso aca es por que selecciono la de la observacion
        OBSERITEM$ = VALOBADA("MOVISTOCTP", "OBSERVAITEM", CONDI$, "NOMESS")
        OBSERVAMODIFICADA$ = TRIM$(InputBox$("Obs.del Item: " & CODDX$, "Observaciòn Actual", OBSERITEM$))
        If (OBSERVAMODIFICADA$) = StringEmpty Then Exit Sub
        If TRIM$(UCase$(OBSERITEM$)) = TRIM$(UCase$(OBSERVAMODIFICADA$)) Then
           Exit Sub
        End If
        
        OPX% = COMALTER%("Atenciòn!!! Confirmar Ediciòn de Observaciòn\\Cancelar\Confirmar Ediciòn")
        If OPX% < 2 Then Exit Sub
        Call ACTUREGISTRO("MOVISTOCTP", "OBSERVAITEM", CONDI$, OBSERVAMODIFICADA$, REGAF&, "NOMESS")
        GRID.TextMatrix(REG&, 12) = OBSERVAMODIFICADA$
   End If

End Sub

Private Sub mnuimprimir_Click()
     Call ReimpresionDeRemito(Text3, OK%)
End Sub

Sub ReimpresionDeRemito(rto$, OK%, Optional TIPODOC1$, Optional NOPRESENTAMENSA%)
    If Me.GRID.Rows < 2 Then Exit Sub
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Dim ENV As New Envap
    PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PRIREEDI"))
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    
    
    REPLAORI% = 0
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If

    OK% = 0
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
   PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
 'MsgBox "printerport = *" + PRINTERPORT$ + "*"
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    'ASIGNO EL FORMULARIO CONFIGURADO - SI NO TIENE CONFIGURADO FRM LE VUELVO A ASIGNAR EL GENERAL
    
    If REMITOSINCARGO% > 0 Then
        FORIPRE$ = TRIM$(CONTROL$("", "FORREMSC"))
        If FORIPRE$ = "" Then
          FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
        End If
    End If
    
 'MsgBox "foripre = *" + FORIPRE$ + "*"
    If TRIM$(FORIPRE$) <> "" Then
      '
'      PREFIREM$ = TRIM(Label10)
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
'      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
'      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      Call LIMPIATETAB
      
      II& = 0
      
      Dim OBSERX$()
        'SE REALIZA ESTE CAMBIO PARA CONSIDERAR
        'SI LAS CAJAS DE TEXTO DE LAS OBSERVACIONES TIENEN MAS TEXTO QUE LA CAPACIDAD VISIBLE
        'REALIZA UN FRATEXTO POR CADA CAJA CONSIDERANDO LA LONG. DEL PARAMETRO OBSERVA DEL FRM
        'ANTES SI EL TEXTO SUPERABA LA LONGITUD DEL PARAMETRO TRUNCABA
      FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "OBSERVA" + "/FORMATO")))
      ATELIB% = XVALO(Mid$(FORTELIB$, 2))

      For i& = 0 To 1
          Call FRATEXTO(Text16(i&), ATELIB%)
          For JJ& = 1 To CARETEX%
            II& = II& + 1
            ReDim Preserve OBSERX$(II&)
            OBSERX$(II&) = RETEX$(JJ&) 'CARGA EN UN VECTOR LAS OBSERVACIONES CON LA LONGITUD PARAMETRIZADA DEL FRM
          Next JJ&
      Next i&
      For U& = 1 To II&
           TTXXX$ = OBSERX$(U&)
           Call GRAREG("!OBSERVOF", TTXXX$, U&)
      Next U&
      'FIN CAMBIO
      FF% = FECHANUM(Text4)
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      
      NPFIJOCAMBACERES& = XVALO(CONTROL("TRCAPAWL", "PFIJOCAM"))
      NPFIJOCTP& = TRIM$(CONTROL("TRCTPAWL", "PFIJOCTP"))
      
      REMITOORIGINAL$ = Text3.TEXT
      PREFIJOORIGINAL% = XVALO(Mid$(REMITOORIGINAL$, 6, 4))
    
      If PREFIJOORIGINAL% <> 9 Then  'SI NO ES 9 ES CAMBACERES
        PrefijoConfiguradoParaRemitos = FORMATOCON0$(NPFIJOCAMBACERES&, 4)
      Else
        PrefijoConfiguradoParaRemitos = FORMATOCON0$(NPFIJOCTP&, 4) 'SI NO ES CONTEPLAST
      End If

      
      
      
      'cambia el prefijo que viene de conteplast por el que esta configurado en el pedido en Envaplast
      'de RT-X-0009-00001234  a RT-X-0002-00001234
      
      RemitoConPrefijoConfigurado$ = Left$(REMITOORIGINAL$, 5) & FORMATOCON0$(XVALO(PrefijoConfiguradoParaRemitos), 4) & "-" & Right(REMITOORIGINAL$, 8)
      NUMEDOC$ = RemitoConPrefijoConfigurado$
      NROREMIT$ = RemitoConPrefijoConfigurado$
      
      'ACA PUEDE MODIFICAR NUMERO DE REMITO (SOLO )
      
30    NROREMITModificado$ = TRIM$(InputBox$("Prefijo de Remito:", "Ingrese Nª de Remito", NROREMIT$))
      If TRIM$(NROREMITModificado$) = "" Then
        Call COMUNI("Debe Ingresar un Nùmero de Remito")
        GoTo 30
      End If
      '
      NroRemitoOriginaldeSucursal$ = REMITOORIGINAL$ 'es el que se realizo en la suscursal sin el cambio del prefijo
      NROREMIT$ = NROREMITModificado$
      
      TIPODOC$ = "Remito"
      NCLI = XVALO(Text2)
      If InStr(EPAC$, "ENVAPLAST") > 0 Then NCLI = NCLI: NROPROV_PROPIO = ENV.NROPROV_ENVAPLAST_WALDOBTT(NCLI1&)
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Me.LENTREGA(2))
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
        If TRIM$(LENTREGA(2).TEXT) <> "" Then TEXTOENTREGA$ = TRIM$(Me.LENTREGA(2).TEXT)
        If TRIM$(DOMITRAN(3).TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Me.DOMITRAN(3))
        If TRIM$(DOMITRAN(4).TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Me.DOMITRAN(4))
        If TRIM$(TXTCODPOSTAL.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(TXTCODPOSTAL)
        CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TEXTOENTREGA$
      End If
      
      
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMITModificado$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(DOMITRAN(4))
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = TRIM$(FORMATNUM$(ACUMLAIMPOX#, "F14.2"))
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = TRIM$(STRNPEDI$)
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(rasocli$(NCLI)) + " - Almacén " + Label11
      
      CODPARAM$(19) = "PROV-CLI": DOCPARAM$(19) = NROPROV_PROPIO
      CODPARAM$(20) = "N-REMITO": DOCPARAM$(20) = NroRemitoOriginaldeSucursal$ 'ACA IMPRIME EL REMITO ORIGINAL COMO DOC.INT.:
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 20
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      pasox% = 0
      For i& = 1 To GRID.Rows - 1
         If TRIM$(GRID.TextMatrix(i&, 0)) <> "*" Then GoTo 50
         REMI$ = TRIM$(GRID.TextMatrix(i&, 9))
         CI% = XVALO(GRID.TextMatrix(i&, 1))
         ORD& = XVALO(GRID.TextMatrix(i&, 2))
         DESKRIP$ = GRID.TextMatrix(i&, 3)
         CAN# = XVALO(GRID.TextMatrix(i&, 4))
         CANING = XVALO(GRID.TextMatrix(i&, 5))
         NPEDI& = XVALO(GRID.TextMatrix(i&, 6))
         LOTE$ = TRIM$(GRID.TextMatrix(i&, 8))
         CODDX$ = TRIM$(GRID.TextMatrix(i&, 10))
         DESCRIDOS$ = GRID.TextMatrix(i&, 11)
         ObservaItem$ = GRID.TextMatrix(i&, 12)
         ORDENCOMOPRA$ = GRID.TextMatrix(i&, 13)
         
         CAITAB1% = CAITAB1% + 1
         '
         TETABU1$(1, CAITAB1%) = CODDX$
         TETABU1$(2, CAITAB1%) = DESKRIP$
         '
         TETABU1$(3, CAITAB1%) = TRIM$(FORMATNUM$(CAN#, "F10.1"))
         TETABU1$(4, CAITAB1%) = LOTE$
         TETABU1$(5, CAITAB1%) = "" 'UNID$
         TETABU1$(6, CAITAB1%) = ORDENCOMOPRA$
         TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
         TETABU1$(8, CAITAB1%) = ORDENCOMOPRA$
         TETABU1$(9, CAITAB1%) = "" '
              '

         TETABU1$(10, CAITAB1%) = ""
         TETABU1$(11, CAITAB1%) = ""
         If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
           DES2$ = DESCRIDOS$
           If DES2$ <> "" Then
              Call PRINTEXT(DES2$, 2, "DES-MAT", FORIPRE$)
           End If
         End If

         TETABU1$(12, CAITAB1%) = ""
         '
         TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
         '
         If CONTROL("REMSTAN", "PRIOBSPE") = "SI" Then 'OBSERVACIONES DE PEDIDOS
            Call PRINTEXT(ObservaItem$, 2, "DES-MAT", FORIPRE$)
          End If
          If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then  'IMPRIME EL LOTE DEBAJO DE LA DESCRIPCION Y/O OBSERVACION DEL PEDIDO
             Call PRINTEXT(LOTE$, 2, "DES-MAT", FORIPRE$)
             Call CARLINSEP("INVISIBLE")         'Carga linea de separacion entre codigo y codigo EN ESTE CASO SIN LINEA
          Else
             Call CARLINSEP
          End If
      
50    Next i&

      CantidadCopiasRemito% = XVALO(Text23)
      If CantidadCopiasRemito% > 0 Then
        CACOPDOC% = CantidadCopiasRemito%
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
        CACOPDOC% = 0
      Else
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
      End If
      End If
'      For YX& = 1 To ULTREG&("!BOLREDET")
'        YY$ = REGLEIDO$("!BOLREDET", YX&)
'        CI% = CVI(Mid$(YY$, 83, 2))
'        CAN# = CVD(Mid$(YY$, 89, 8))
'        NPEDI& = CVS(Mid$(YY$, 151, 4))
'        CAC# = 0
'        If CI% > 0 Or Frame15.Visible = True Then
'          If CI% <= NUMAS% Or Frame15.Visible = True Then
'            If CAN# + CAC# > 0.005 Or Frame15.Visible = True Then
'              '
'              UNID$ = Unimed$(CI%)
'              If UNID$ = "" Then UNID$ = Unimed$(CI%)
'              CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
'              If CAC# >= 0.05 Then
'                CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
'                Call REPLA(CANS$, "(*)", 1, 3)
'              End If
'              NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
'              NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'              ORD& = XVALO(Mid$(YY$, 241, 5)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
'              '
'              SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
'              If SERINU$ <> "" Then
'                If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
'                  SERINU$ = SERINU$ + " - "
'                End If
'              End If
'              SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
'              '
'              CAITAB1% = CAITAB1% + 1
'              CODDX$ = CODEXT$(CI%)
'              If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
'                CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
'              End If
'              If REPLAORI% = 1 Then
'                CODDX$ = CODORIG$(CI%, NCLI)
'              End If
'              '
'              TETABU1$(1, CAITAB1%) = CODDX$
'              TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
'              If Frame15.Visible = True Then
'                TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
'              End If
'              '
'              TETABU1$(3, CAITAB1%) = CANS$
'              TETABU1$(4, CAITAB1%) = NULOTE$
'              TETABU1$(5, CAITAB1%) = UNID$
'              TETABU1$(6, CAITAB1%) = NUOCO$
'              TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
'              TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
'              TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
'              '
'              'VALORIZACION ALEJPA SEGUN LISTA VIGENTE
'              IMPOUNI# = 0: IMPORTEX# = 0
'              If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 Then
'                IMPOUNI# = PRELISTACT#(CI%, TIMONE%)
'                IMPOUNI# = IMPOUNI# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
'                IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
'                ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
'              End If
'
'              TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
'              TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
'              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
'                 DES2$ = TRIM$(DESCRIT2$(CI%))
'                 If DES2$ <> "" Then
'                    Call PRINTEXT(DESCRIT2$(CI%), 2, "DES-MAT", FORIPRE$)
'                 End If
'              End If
'
'                CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
'                'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
'                If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
'                'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
'              If InStr(EPAC$, "ENVAPLAST") > 0 Then
''                 If CORIG$ <> "" Then 'al final dijo Enrique que no quiere que se imprima
''                   Call PRINTEXT(CORIG$, 1, "COD-MAT", FORIPRE$)
''                 End If
'              Else
'                 TETABU1$(12, CAITAB1%) = CORIG$
'              End If
'              'FIN ALEJPA
'              '
'              TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
'
'
'
'              'IMPRESION DE ESPECIFICACION
'              If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
'                XX$ = ESPECIFIT$(CI%)
'                Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
'              End If
'              '
'              If CONTROL("REMSTAN", "PRIOBSPE") = "SI" Then 'OBSERVACIONES DE PEDIDOS
'                NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
'                NUORITPE% = CVS(Mid$(YY$, 189, 4))
'                OBSPE$ = OBSERVAPEDI$(CI%, NCLI, NUPEDIORI&, NUORITPE%)
'                Call PRINTEXT(OBSPE$, 2, "DES-MAT", FORIPRE$)
'              End If
'              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then  'IMPRIME EL LOTE DEBAJO DE LA DESCRIPCION Y/O OBSERVACION DEL PEDIDO
'                 Call PRINTEXT(NULOTE$, 2, "DES-MAT", FORIPRE$)
'                 Call CARLINSEP("INVISIBLE")         'Carga linea de separacion entre codigo y codigo EN ESTE CASO SIN LINEA
'              Else
'                 Call CARLINSEP
'              End If
'
'              If SERINU$ <> "" Then
'                CAITAB1% = CAITAB1% + 1
'                For KI1% = 1 To 20
'                  TETABU1$(KI1%, CAITAB1%) = ""
'                Next KI1%
'                TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
'              End If
'              '
'              Call DETLOTREM(CI%, YX&) 'SI NO LO AGREGA A LA DESCRIPCION
'              '
'              'NCIY% = NCLI
'              CIIY% = CI%
'              'Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
'            End If
'          End If
'        End If
'      Next YX&
'      '
' 'MsgBox "antes de PRINTDOC(FOREMIFA)"
'      CantidadCopiasRemito% = XVALO(Text23)
'      If CantidadCopiasRemito% > 0 Then
'        CACOPDOC% = CantidadCopiasRemito%
'        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
'        CACOPDOC% = 0
'      Else
'        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
'      End If
'      '
'    End If
    '

    '99
99
End Sub


Private Sub Text14_Change()

End Sub

Private Sub Text14_DblClick()

End Sub


Private Sub mnuRemOficial_Click()
      FE% = HOY(HO$)
      User% = USNBR% Mod 100
      
      HaySeleccion% = 0
      For i& = 1 To GRID.Rows - 1
         If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
            HaySeleccion% = HaySeleccion% + 1
            Exit For
         End If
      Next i&
      '
      If HaySeleccion% < 1 Then
       Call COMUNI("Atención No se Han Seleccionado Items Para Imprimir\Puede Seleccionar Todos desde el Menú > Edición > Revertir Selección\O Realizando Doble Click Sobre los Items a Imprimir")
       Exit Sub
      End If
      '
      If HaySeleccion% > 0 And HaySeleccion% < GRID.Rows - 1 Then
       OPX% = COMALTER%("Atención Solo se Imprimiran los Items Seleccionados\\Cancelar\Imprimir Items Seleccionados")
       If OPX% < 2 Then Exit Sub
      End If
      
      PREFI$ = TRIM$(InputBox$("Prefijo de Remito:", "Ingrese el Prefijo", PREFI$))
      If PREFI$ = "" Then Exit Sub
      '
      If Len(PREFI$) > 4 Then
         Call COMUNI("Prefijo no Válido (Máximo 4 Dígitos)")
         Exit Sub
      End If
      '
      While Len(PREFI$) < 4: PREFI$ = "0" + PREFI$: Wend
      '
      For U& = 1 To 4
           UI% = Asc(Mid$(PREFI$, U&))
           If UI% < 48 Or UI% > 57 Then
               Call COMUNI("Prefijo no Válido")
               Exit Sub
               Exit Sub
           End If
      Next U&
      '
20    NROREMIX$ = TRIM$(InputBox$("Nro de Remito: (Máximo 8 Caracteres)", "Prefijo " & PREFI$, NROREMIX$))
      If NROREMIX$ = "" Then Exit Sub
      
      If Len(NROREMIX$) > 8 Then
         Call COMUNI("Longitud Ingresada Supera la Cantidad Permitida")
         Exit Sub
      End If
      If XVALO(NROREMIX$) < 1 Then
         Call COMUNI("Debe Ingresar Números")
         Exit Sub
      End If
      REMITOFICIAL$ = "RT-X-" & PREFI$ & "-" & FORMATOCON0$(XVALO(NROREMIX$), 8)
      If CantRegistros&("MOVIOFI", "REMIOFI = '" & REMITOFICIAL$ & "' AND STATUS < 9", "NOMESS") > 0 Then
         Call COMUNI("Imposible Continuar !! Número de Remito Existente\Deberá Anular el Mismo e Intentar Nuevamente")
         GoTo 20
      End If
      
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      
      On Error GoTo ERROR_GUARDAR
      FLEXCONN.BeginTrans
      SQLX$ = "Select * from MOVIOFI WHERE MOVIOFI_ID < 1"
      rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      
      With rs
        For i& = 1 To GRID.Rows - 1
            If TRIM$(GRID.TextMatrix(i&, 0)) <> "*" Then GoTo 50
            .AddNew
            REMI$ = TRIM$(GRID.TextMatrix(i&, 9))
            NUREM& = XVALO(Mid$(REMI$, 6))
            NROPED& = XVALO(GRID.TextMatrix(i&, 6))
            CI1% = XVALO(GRID.TextMatrix(i&, 1))
            CODDX$ = CODEXT$(XVALO(GRID.TextMatrix(i&, 1)))
            If REPLAORI% = 1 Then
               CODDX$ = CODORIG$(CI1%, NCLI)
            End If
            NPDX& = XVALO(GRID.TextMatrix(i&, 6))
            CANSAL = XVALO(GRID.TextMatrix(i&, 4))
            CANING = XVALO(GRID.TextMatrix(i&, 5))
            NSERIE$ = TRIM$(GRID.TextMatrix(i&, 8))
            !COD_INTE = CI1%
            !Cod_Exte = CODDX$
            !FeReMovi = Now
            !FECHMOV = CVDAT(FE%)
            !Descrip = Left$(DESCRIT0$(CI1%), 64)
            !idenlote = NSERIE$
            !DOPRILAR = REMI$
            !NUPEDCLI = NROPED&
            !Nume_Clie = Text2
            !CANTINGRE = CANING
            !CANTSALID = CANSAL
            !DOPRINUM = NUREM&
            !User_Tran = User%
            !REMIOFI = REMITOFICIAL$
            !REFERENCIA = Left$(rasocli$(XVALO(Text2)), 18) & "-" & FECHATEX$(FE%)
            !Status = 0
            If NPDX& <> NPDX_ANT& And NPDX& > 0 Then
               TTXX$ = Space$(76)
               TXTX$ = "Generación de Remito Oficial N°: " & REMITOFICIAL$ & vbCrLf
               Call AGREGA_ANOTAPED_NEW(NPDX&, TXTX$)
            End If
            NPDX_ANT& = NPDX&
            .Update
50      Next i&
      End With
ERROR_GUARDAR:

      If Err.Number <> 0 Then
        FLEXCONN.RollbackTrans
        Call COMUNI("Se Produjo un Error al Grabar.\No se Ha Realizado Ninguna Transacción.")
      Else
        FLEXCONN.CommitTrans
      End If
      
99
      On Error GoTo 0
      rs.Close
      Set rs = Nothing

      Call ReimpresionDeRemito(REMITOFICIAL$, OK%, "REMITOFI", 1)
      
End Sub

Private Sub mnuReportesVarios_Click()
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuRevAnul_Click()

10 Campos$ = "Fecha/D8;Remito Oficial/A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES ANULADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(2)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   Call ACTUREGISTRO("MOVIOFI", "STATUS", "REMIOFI = '" & NROREMI$ & "'", 0, REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVIOFI", "REFERENCIA", "REMIOFI = '" & NROREMI$ & "'", Mid$(RESP_ZELE_VECT(3), 7), REGAF&, "NOMESS")

   GoTo 10

End Sub

Sub mnuRevertirSeleccion_Click()
   mnuRevertirSeleccion.Enabled = False
   Screen.MousePointer = 11
   FIN& = GRID.Rows - 1
   GRID.Redraw = False
   For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, i&)
   Next i&
   Call CalcutotalSeleccionado
   GRID.Redraw = True
   Screen.MousePointer = 1
   mnuRevertirSeleccion.Enabled = True
End Sub


Private Sub mnuVerRemiof_Click()
10 Campos$ = "Fecha/D8;Remito Of./A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES GENERADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(0)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   
   DOCUNU& = 0
   TIDOCUM$ = "REMITOFI"

   NUDOCU$ = TRIM$(Str$(NPX&))
   Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, SAFETEXT$(XVALO(Mid$(NROREMI$, 11))), NROREMI$)
   GoTo 10

End Sub

Private Sub mnuVincularReportes_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     TXTFLETE = "Flete A Cargo del Cliente"
   Else
    TXTFLETE = ""
   End If
End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
     EMPRESAX$ = EMPREAC$
     If Len(EMPREAC$) > 6 Then
       POS1% = InStr(EMPREAC$, " ")
       EMPRESAX$ = Mid$(EMPREAC$, 1, POS1%)
     End If
     TXTFLETE = "Flete A Cargo de " & EMPRESAX$
   Else
     TXTFLETE = ""
   End If

End Sub


Private Sub Text17_Change()
A = Text17
End Sub

Sub Text3_Change()
    '
    If Option4.Value = True Then valor% = 1
    If Option3.Value = True Then valor% = 2
    Call Carga_de_Pantalla(Text3, valor%)

End Sub


Sub Carga_de_Pantalla(DOCUREM$, Optional AGRUPADA%)
    Text2 = ""
'    Text3 = ""
    'Text4 = ""
    Text6 = ""
    'Text8 = ""
    Text12 = ""
    Text15 = ""
    Text18 = ""
    Text19 = ""
    Text14 = ""
    Text13 = ""
    Text21 = ""
    Text20 = ""
    Text23 = ""
    Me.LENTREGA(0) = ""
    Me.LENTREGA(1) = ""
    Me.LENTREGA(2) = ""
    Text12 = ""
    Text18 = ""
    Me.DOMITRAN(3) = ""
    Me.DOMITRAN(4) = ""
    Me.Text45 = ""
    Text7 = ""
    If Command1.Enabled = True Then
      GRID.Rows = 1
      Text17 = ""
      Label18 = ""
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    If Text8 = "RP" Then CMO$ = "RP"
    '
    
    ACUMULATOTAL = 0
    If Me.Text3 <> "" Then
        FF% = FECHANUM(Text4)
        CONDI$ = "CodiMovi = '" & CMO$ & "'  "
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & TRIM$(Text3) & "'"
        CONDI$ = CONDI$ + " AND CANTSALID > 0 AND FECHMOV='" & FETEXCOR1$(FF%) & "'"
        VUELTA% = 0: j& = 0
        If AGRUPADA% = 1 Then
            SQLX$ = "SELECT COD_INTE, SUM(CantSalid) AS CANTSALID, SUM(cantingre) AS CANTINGRE, NUPEDCLI"
            SQLX$ = SQLX$ + ", SUM(ValoUnit * CANTSALID)AS IMPORTE ,DOMIENT,ORDENCOMPRA,COD_EXTE"
            SQLX$ = SQLX$ + " , Nume_Clie, FECHMOV,IDENLOTE,RASO_CLI FROM MOVISTOCTP WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
            
            SQLX$ = SQLX$ + " GROUP BY COD_INTE,NUPEDCLI,DOMIENT,ORDENCOMPRA,COD_EXTE"
            SQLX$ = SQLX$ + " , Nume_Clie, FECHMOV,IDENLOTE,RASO_CLI"
            SQLX$ = SQLX$ + " ORDER BY NUPEDCLI,FECHMOV"
        ElseIf AGRUPADA% = 2 Then
            SQLX$ = "SELECT COD_INTE, SUM(CantSalid) AS CANTSALID, SUM(cantingre) AS CANTINGRE, NUPEDCLI"
            SQLX$ = SQLX$ + ", SUM(ValoUnit * CANTSALID)AS IMPORTE ,DOMIENT,ORDENCOMPRA,COD_EXTE"
            SQLX$ = SQLX$ + " , Nume_Clie, FECHMOV,RASO_CLI FROM MOVISTOCTP WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
            
            SQLX$ = SQLX$ + " GROUP BY COD_INTE,NUPEDCLI,DOMIENT,ORDENCOMPRA,COD_EXTE"
            SQLX$ = SQLX$ + " , Nume_Clie, FECHMOV,RASO_CLI"
            SQLX$ = SQLX$ + " ORDER BY NUPEDCLI,FECHMOV"
        
        Else
            SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, ValoUnit,DOMIENT,ORDENCOMPRA,COD_EXTE"
            SQLX$ = SQLX$ + " , Nume_Clie, FECHMOV,IDENLOTE,RASO_CLI,DESCRIDOS,OBSERVAITEM FROM MOVISTOCTP WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
            SQLX$ = SQLX$ + " ORDER BY NuOrItem"

        End If
        
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             If VUELTA% = 0 Then
               VUELTA% = 1
               Text2 = XVALO(!Nume_Clie)
               Text6 = SAFETEXT$(!RASO_CLI)
               Text4 = FECHATEX$(CVINT(!FECHMOV))
               LENTREGA(0) = SAFETEXT(!DOMIENT)
               DOMITRAN(3) = SAFETEXT(!DOMIENT)
               Text20 = SAFETEXT(!ORDENCOMPRA)
             End If
             j& = GRID.Rows
             GRID.Rows = j& + 1
             GRID.TextMatrix(j&, 1) = XVALO(!COD_INTE)
             GRID.TextMatrix(j&, 2) = j&
             '
             GRID.TextMatrix(j&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F9.1"))
             GRID.TextMatrix(j&, 5) = TRIM$(FORMATNUM$(XVALO(!CANTINGRE), "F9.1"))
             GRID.TextMatrix(j&, 6) = XVALO(!NUPEDCLI)
             
             
             IMPORTEX = 0
             
             If AGRUPADA% < 2 Then
                GRID.TextMatrix(j&, 8) = SAFETEXT$(!idenlote)
             End If
             
             GRID.TextMatrix(j&, 9) = Text3
             GRID.TextMatrix(j&, 10) = SAFETEXT$(!Cod_Exte)
             If Option4.Value = True Or Option3.Value = True Then
                'ESTOS CAMPOS NO SE PUEDEN AGRUPAR EN CONSULTA SQL
                CONDIAGRU$ = CONDI$ + "AND NUPEDCLI = " & XVALO(!NUPEDCLI)
                CONDIAGRU$ = CONDIAGRU$ + " And COD_INTE = " & XVALO(!COD_INTE)
                DESCRIPX$ = VALOBADA("MOVISTOCTP", "DESCRI_LAR", CONDIAGRU$, "NOMESS")
                DESCRIDOSX$ = VALOBADA("MOVISTOCTP", "DESCRIDOS", CONDIAGRU$, "NOMESS")
                ObservaItemx$ = VALOBADA("MOVISTOCTP", "ObservaItem", CONDIAGRU$, "NOMESS")
                GRID.TextMatrix(j&, 3) = DESCRIPX$
                GRID.TextMatrix(j&, 7) = TRIM$(FORMATNUM$(XVALO(!IMPORTE), "F12.2"))
                GRID.TextMatrix(j&, 11) = DESCRIDOSX$
                GRID.TextMatrix(j&, 12) = ObservaItemx$
                ACUMULATOTAL = ACUMULATOTAL + XVALO(GRID.TextMatrix(j&, 7))
             Else
                GRID.TextMatrix(j&, 3) = SAFETEXT$(!DESCRI_LAR)
                GRID.TextMatrix(j&, 7) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
                GRID.TextMatrix(j&, 11) = SAFETEXT$(!DESCRIDOS)
                GRID.TextMatrix(j&, 12) = SAFETEXT$(!ObservaItem)
                IMPORTEX = (XVALO(!CANTSALID) - XVALO(!CANTINGRE)) * XVALO(!VALOUNIT)
                ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
             End If
             GRID.TextMatrix(j&, 13) = SAFETEXT$(!ORDENCOMPRA)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        Text17 = TRIM$(FORMATNUM$(XVALO(ACUMULATOTAL), "F12.2"))
        Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)
        '

'        'MOSTRAR DATOS DE ENTREGA
'        For i& = 1 To GRID.Rows - 1
'           NPDX& = XVALO(GRID.TextMatrix(i&, 6))
'           If NPDX& > 0 Then
'             SQLX$ = "SELECT DOMIENT,LOCAENT,TranspEnt,NROOCOM,FLETE_TEX, NTranspo FROM PEDENCA WITH(NOLOCK)"
'             SQLX$ = SQLX$ + " WHERE NROPED = " & NPDX&
'             Set rs = READSET(SQLX$)
'             With rs
'             Do While Not .EOF
'               If TRIM$(SAFETEXT$(!DOMIENT)) <> "" Then
'                 LENTREGA(0) = SAFETEXT(!DOMIENT)
'                 LENTREGA(1) = SAFETEXT(!LOCAENT)
'                 NTRAN% = XVALO(!NTranspo)
'                 If NTRAN% > 0 Then ' SI VIENE CON CODIGO DE TRANSPORTE COMPLETA AUTOMATICAMENTE
'                   Text45 = NTRAN%
'                 Else
'                   LENTREGA(2) = SAFETEXT(!TranspEnt)
'                 End If
'                 Text20 = SAFETEXT(!NROOCOM)
'                 TXTFLETE = SAFETEXT(!FLETE_TEX)
'                 GoTo 20
'               End If
'               .MoveNext
'             Loop
'             End With
'20           rs.Close
'             Set rs = Nothing
'           End If
'        Next i&
    End If
    Call mnuRevertirSeleccion_Click
'    Call CalcutotalSeleccionado
End Sub
Sub CalcutotalSeleccionado()

        Text17 = "": ACUMULATOTAL = 0
        For i& = 1 To GRID.Rows - 1
            IMPORTEX = 0

            If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
              If Option4.Value = True Or Option3.Value = True Then
                 IMPORTEX = XVALO(GRID.TextMatrix(i&, 7))
              Else
                 IMPORTEX = (XVALO(GRID.TextMatrix(i&, 4)) - XVALO(GRID.TextMatrix(i&, 5))) * XVALO(GRID.TextMatrix(i&, 7))
              End If
            End If
            ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
        Next i&
        Text17 = FORMATNUM$(ACUMULATOTAL, "F12.2")
End Sub
Sub LIMPIARTEXT(FORMULARIO As Form, LIMPIOTEXT%)
  Dim CONTROLX As CONTROL
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
           CONTROLX.TEXT = ""
        End If
  Next
End Sub

Private Sub Text45_Change()
  Me.LENTREGA(2) = ""
  Text12 = ""
  Text18 = ""
  Me.DOMITRAN(3) = ""
  Me.DOMITRAN(4) = ""
  '
  NTRA = XVALO(Text45)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
        Text12 = CUITTRANS$(NTRAI%)
        LENTREGA(2) = DENOTRANS$(NTRAI%)
        DOMITRAN(3) = DOMITRANS$(NTRAI%)
        DOMITRAN(4) = LOCATRANS$(NTRAI%)
      End If
    End If
  End If

End Sub



VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPFACTU00 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Facturación"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11610
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
   ScaleHeight     =   7215
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   115
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Device"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   9400
      TabIndex        =   101
      Top             =   135
      Width           =   1100
      Begin VB.OptionButton Option26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fac.El."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   80
         TabIndex        =   114
         ToolTipText     =   "Doble Click para Configurar Punto de Venta (Normal - Bienes de Capital)"
         Top             =   1000
         Width           =   950
      End
      Begin VB.OptionButton Option23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   80
         TabIndex        =   103
         Top             =   285
         Value           =   -1  'True
         Width           =   950
      End
      Begin VB.OptionButton Option22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "C.Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   80
         TabIndex        =   102
         Top             =   580
         Width           =   950
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   8520
      Left            =   10710
      ScaleHeight     =   8460
      ScaleWidth      =   1845
      TabIndex        =   75
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox Picture16 
         Height          =   1200
         Left            =   0
         ScaleHeight     =   1140
         ScaleWidth      =   1020
         TabIndex        =   117
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   7200
            Left            =   0
            Picture         =   "OPFACTU00.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Imput"
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
         Picture         =   "OPFACTU00.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Tabla de Motivos de Imputación"
         Top             =   5480
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Equiv"
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
         Picture         =   "OPFACTU00.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   4790
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU00.frx":262C
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Configuraciones Especiales"
         Top             =   4100
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Picture         =   "OPFACTU00.frx":2776
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   3400
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Picture         =   "OPFACTU00.frx":2A80
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Parámetros varios: Tasas IVA, Letra Exportación, Vencimientos de CAI"
         Top             =   2710
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
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
         Picture         =   "OPFACTU00.frx":2BCA
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1665
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Picture         =   "OPFACTU00.frx":2ED4
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   975
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "OPFACTU00.frx":31DE
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   285
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   76
         Top             =   6300
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   77
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "OPFACTU00.frx":3328
         Top             =   6630
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Puertos de Impresión -Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5350
      TabIndex        =   37
      Top             =   105
      Width           =   4005
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   680
         Width           =   2955
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   300
         Width           =   2955
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Suprimir  Re-Direccionamiento On Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   105
         TabIndex        =   81
         Top             =   1020
         Width           =   3650
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3570
         TabIndex        =   52
         Top             =   680
         Width           =   330
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3570
         TabIndex        =   51
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rem:"
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
         Index           =   51
         Left            =   105
         TabIndex        =   39
         Top             =   735
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fac:"
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
         Index           =   41
         Left            =   50
         TabIndex        =   38
         Top             =   375
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formularios de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7040
      Left            =   210
      TabIndex        =   12
      Top             =   105
      Width           =   5000
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sistema Mantenimiento "
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
         Left            =   240
         TabIndex        =   118
         Top             =   6720
         Width           =   2475
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación Tipo ""B"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Index           =   1
         Left            =   210
         TabIndex        =   13
         Top             =   2400
         Width           =   4600
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
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
            Index           =   5
            Left            =   165
            TabIndex        =   110
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   1185
            Width           =   630
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
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
            Index           =   1
            Left            =   165
            TabIndex        =   107
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   270
            Width           =   630
         End
         Begin VB.TextBox Text1 
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
            Index           =   5
            Left            =   2415
            TabIndex        =   86
            Top             =   1155
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   3465
            TabIndex        =   85
            Top             =   1155
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   5
            Left            =   3780
            Picture         =   "OPFACTU00.frx":524A
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   1155
            Width           =   650
         End
         Begin VB.TextBox Text1 
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
            Index           =   1
            Left            =   2415
            TabIndex        =   16
            Top             =   260
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   3465
            TabIndex        =   15
            Top             =   260
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   1
            Left            =   3780
            Picture         =   "OPFACTU00.frx":5554
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   260
            Width           =   650
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4725
            Y1              =   1000
            Y2              =   1000
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mat.Texto Libre:"
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
            Index           =   5
            Left            =   15
            TabIndex        =   88
            Top             =   1230
            Width           =   2325
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   5
            Left            =   150
            TabIndex        =   87
            Top             =   1530
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. Codificado:"
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
            Index           =   1
            Left            =   0
            TabIndex        =   18
            Top             =   335
            Width           =   2325
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   630
            Width           =   3480
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formulario de Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Index           =   3
         Left            =   210
         TabIndex        =   31
         Top             =   5620
         Width           =   4600
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
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
            Index           =   3
            Left            =   170
            TabIndex        =   109
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   235
            Width           =   630
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   3
            Left            =   3780
            Picture         =   "OPFACTU00.frx":585E
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   260
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   3465
            TabIndex        =   33
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Index           =   3
            Left            =   2415
            TabIndex        =   32
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   150
            TabIndex        =   36
            Top             =   630
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            Index           =   3
            Left            =   0
            TabIndex        =   35
            Top             =   335
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación de Exportación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1070
         Index           =   2
         Left            =   210
         TabIndex        =   25
         Top             =   4450
         Width           =   4600
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
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
            Index           =   2
            Left            =   170
            TabIndex        =   108
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   235
            Width           =   630
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   2
            Left            =   3780
            Picture         =   "OPFACTU00.frx":5B68
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   260
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3465
            TabIndex        =   27
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Index           =   2
            Left            =   2415
            TabIndex        =   26
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   150
            TabIndex        =   30
            Top             =   630
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            Index           =   2
            Left            =   0
            TabIndex        =   29
            Top             =   335
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación Tipo ""A"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   315
         Width           =   4600
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   111
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   1185
            Width           =   615
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
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
            Index           =   0
            Left            =   170
            TabIndex        =   106
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   235
            Width           =   630
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   4
            Left            =   3780
            Picture         =   "OPFACTU00.frx":5E72
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1155
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   3465
            TabIndex        =   71
            Top             =   1155
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Index           =   4
            Left            =   2415
            TabIndex        =   70
            Top             =   1155
            Width           =   1065
         End
         Begin VB.TextBox Text1 
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
            Index           =   0
            Left            =   2415
            TabIndex        =   22
            Top             =   260
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   3465
            TabIndex        =   21
            Top             =   260
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   0
            Left            =   3780
            Picture         =   "OPFACTU00.frx":617C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   260
            Width           =   650
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   150
            TabIndex        =   74
            Top             =   1530
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. Texto Libre:"
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
            Index           =   4
            Left            =   45
            TabIndex        =   73
            Top             =   1230
            Width           =   2325
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4725
            Y1              =   1000
            Y2              =   1000
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. Codificado:"
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
            Left            =   360
            TabIndex        =   24
            Top             =   315
            Width           =   2010
         End
         Begin VB.Label DEFOR1 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   150
            TabIndex        =   23
            Top             =   630
            Width           =   3480
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles Operativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      Left            =   5350
      TabIndex        =   0
      Top             =   1440
      Width           =   5160
      Begin VB.Frame Frame20 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Contado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4095
         TabIndex        =   112
         ToolTipText     =   "Tilde Este Casillero para Habilitar Opción de Cancelación de Contado"
         Top             =   1360
         Width           =   880
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   285
            TabIndex        =   113
            ToolTipText     =   "Tilde Este Casillero para Habilitar Opción de Cancelación de Contado"
            Top             =   190
            Width           =   500
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Agrupación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   98
         Top             =   3950
         Width           =   1335
         Begin VB.OptionButton Option21 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Por Suc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   80
            TabIndex        =   100
            Top             =   480
            Width           =   1170
         End
         Begin VB.OptionButton Option20 
            BackColor       =   &H00FFFFC0&
            Caption         =   "No Agrup."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   80
            TabIndex        =   99
            Top             =   240
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2160
         TabIndex        =   94
         Top             =   3950
         Width           =   1400
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agrega Desp.Imp."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "%VDec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4095
         TabIndex        =   82
         ToolTipText     =   "Coeficiente de Valor Declarado"
         Top             =   820
         Width           =   880
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   190
            MaxLength       =   2
            TabIndex        =   83
            ToolTipText     =   "Coeficiente de Valor Declarado"
            Top             =   210
            Width           =   400
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   1730
         TabIndex        =   67
         Top             =   3000
         Width           =   1680
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pre-Impresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   69
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Auto-Numera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   68
            Top             =   520
            Width           =   1420
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formulario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   210
         TabIndex        =   64
         Top             =   3000
         Width           =   1470
         Begin VB.OptionButton Option11 
            BackColor       =   &H00FFFFC0&
            Caption         =   "H.Suelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   66
            Top             =   520
            Width           =   1100
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   65
            Top             =   300
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   3460
         TabIndex        =   61
         Top             =   3000
         Width           =   1470
         Begin VB.OptionButton Option13 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pr.Neto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   63
            Top             =   300
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Option14 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Porciento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   62
            Top             =   520
            Width           =   1170
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   3480
         TabIndex        =   53
         Top             =   1950
         Width           =   1455
         Begin VB.OptionButton Option19 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Multi-Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   110
            TabIndex        =   97
            Top             =   690
            Width           =   1210
         End
         Begin VB.OptionButton Option16 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   110
            TabIndex        =   55
            Top             =   240
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton Option15 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Bi-Monetario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   110
            TabIndex        =   54
            Top             =   470
            Width           =   1200
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pre-Facturación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   1510
         TabIndex        =   48
         Top             =   1950
         Width           =   1920
         Begin VB.OptionButton Option18 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   60
            ToolTipText     =   "Por Remitos Pendientes de Facturación"
            Top             =   240
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pedido / F.Lim."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Por Pedidos Pendientes de Entrega Hasta Fecha de Corte"
            Top             =   480
            Width           =   1425
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Selecc O/ Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Por Selección de Ordenes de Venta Pendientes de Entrega"
            Top             =   720
            Width           =   1740
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Expedición - Depósito de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   210
         TabIndex        =   43
         Top             =   4900
         Width           =   4725
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1155
            TabIndex        =   45
            Text            =   " "
            Top             =   300
            Width           =   540
         End
         Begin VB.CommandButton Command9 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   44
            Top             =   300
            Width           =   175
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1995
            TabIndex        =   47
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito:"
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
            Left            =   0
            TabIndex        =   46
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1720
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Width           =   3750
         Begin VB.CommandButton Command3 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3465
            TabIndex        =   58
            Top             =   1125
            Width           =   150
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   700
            Left            =   1575
            TabIndex        =   9
            Top             =   210
            Width           =   2115
            Begin VB.OptionButton Option17 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Unif.Debit.+ Credit."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   105
               TabIndex        =   59
               Top             =   440
               Width           =   2010
            End
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Independiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   105
               TabIndex        =   11
               Top             =   220
               Width           =   1590
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Correlativa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   105
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   8
            Top             =   290
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   7
            Top             =   510
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2520
            TabIndex        =   57
            Top             =   1125
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración:"
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
            Left            =   -105
            TabIndex        =   56
            Top             =   1155
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3780
            Y1              =   950
            Y2              =   950
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
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
         Left            =   4095
         TabIndex        =   4
         Top             =   220
         Width           =   880
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   190
            TabIndex        =   5
            Text            =   "FC"
            Top             =   210
            Width           =   400
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   210
         TabIndex        =   1
         Top             =   1950
         Width           =   1245
         Begin VB.OptionButton Option24 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sin Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   104
            Top             =   690
            Width           =   1100
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Itemizado "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   470
            Width           =   1100
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   250
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Verificaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   40
         Top             =   3950
         Width           =   1905
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Exist. de Piezas"
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
            TabIndex        =   41
            Top             =   480
            Width           =   1755
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Lím. de Crédito"
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
            TabIndex        =   42
            Top             =   240
            Width           =   1635
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPFACTU00.frx":6486
      TabIndex        =   116
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPFACTU00.frx":6500
      Top             =   1200
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu parametrosvarios 
         Caption         =   "Parámetros varios: Tasas IVA, Letra Exportación, Vencimientos de CAI"
      End
      Begin VB.Menu cotimoneda 
         Caption         =   "Cotización Dolar y Tipo de Cambio"
      End
      Begin VB.Menu configuracionesespeciales 
         Caption         =   "Configuraciones Especiales"
      End
      Begin VB.Menu codigosequivalentesdeclientes 
         Caption         =   "Códigos Equivalentes de Clientes"
      End
      Begin VB.Menu motivosdeimputacion 
         Caption         =   "Tabla de Motivos de Imputación"
      End
   End
End
Attribute VB_Name = "OPFACTU00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub codigosequivalentesdeclientes_Click()
   Call Command10_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

'\\\OT-06-0339-MH-24/08/06///
Private Sub Command10_Click()
Command10.Enabled = False
Call FINAL("*CODORCL")
Command10.Enabled = True
End Sub
'\\\OT-06-0339-MH-FIN///

Private Sub Command11_Click()
   '
   Command11.Enabled = False
   OPFACES00.Show 1
   Command11.Enabled = True
   '
End Sub

Private Sub Command12_Click()
    Call BLANARCH("!IMPDECRE")
    For i& = 1 To ULTREG&("IMPDECRE")
        x$ = REGLEIDO$("IMPDECRE", i&)
        If CVI(Left$(x$, 2)) > 0 Then
            Call REGAPP("!IMPDECRE", x$)
        End If
    Next i&
    Call CIERRARCH("IMPDECRE")
    Call CIERRARCH("!IMPDECRE")
    '
10  VTB% = VENTABU%("IMPDECRE/NC/TITUPAN=Importe Débitos y Créditos")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For i& = 1 To ULTREG&("!IMPDECRE")
        x$ = REGLEIDO$("!IMPDECRE", i&)
        If CVI(Left$(x$, 2)) > 0 Then
           MOTIVO$ = TRIM(Mid$(x$, 3, 40))
           If MOTIVO$ = "" Then
              Call MENSERR(24, "Falta Motivo de la Imputacion")
              GoTo 10
           End If
           CUECON% = CVI(Mid$(x$, 43, 2))
           If CODCUE$(CUECON%) = "" Then
              Call MENSERR(24, "Falta Cuenta Contable en la Imputacion")
              GoTo 10
           End If
           J& = J& + 1
           Call GRAREG("IMPDECRE", x$, J&)
        End If
    Next i&
    Call SETULTREG("IMPDECRE", J&)
    Call CIERRARCH("IMPDECRE")
    '


End Sub

Private Sub Command13_Click(index As Integer)
  Command13(index).Enabled = False
  Call TRALOFIL("LEYFAC", 1, MKI$(index))
    
  OPX% = ALTERNA("Leyenda Predefinida\Alarma Fin Talonario")
  Select Case OPX%
   Case Is = 1  'LEYENDA PREDEFINIDA
    VTB% = VENTABU%("LEYFAC/TITUPAN=" + DEFOR1(index))
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("LEYFAC")
        XX$ = REGLEIDO$("LEYFAC", U&)
        If CVI(Left$(XX$, 2)) <> index Then
          JJ& = JJ& + 1
          Call GRAREG("LEYFAC", XX$, JJ&)
        End If
      Next U&
      For U& = 1 To ULTREG&("!LEYFAC")
        XX$ = REGLEIDO$("!LEYFAC", U&)
        Call REPLA(XX$, MKI$(index), 1, 2)
        JJ& = JJ& + 1
        Call GRAREG("LEYFAC", XX$, JJ&)
      Next U&
      Call SETULTREG("LEYFAC", JJ&)
    End If
    '
  Case Is = 2     'Grabar Numero Fin de Talonario
    FORMU$ = Text1(index)
    NUMERO$ = CONTROL("ALATALO", FORMU$)
    RESP$ = InputBox("Talonario " + FORMU$ + " avisar al Numero:", , NUMERO$)
    Call GRACONTROL("ALATALO", FORMU$, RESP$)
    '
  End Select
  Command13(index).Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORFAC-A", Text1(0).TEXT)
    Call GRACONTROL("", "FACTEXLI", Text1(4).TEXT)
    Call GRACONTROL("", "FORFAC-B", Text1(1).TEXT)
    Call GRACONTROL("", "FORFAC-E", Text1(2).TEXT)
    Call GRACONTROL("", "FACTELIB", Text1(5).TEXT)
    Call GRACONTROL("", "FOREMIFA", Text1(3).TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text7.TEXT)
    Call GRACONTROL("", "CODFACTU", Text9.TEXT)
    '
    MONUME$ = "NO"
    If Check6.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("FACTURA", "SISTMANT", MONUME$)
    '
    Call GRACONTROL(IDENTER$, "PORTFAC", Combo1.TEXT)
    Call GRACONTROL("*", "PORTFAC", Combo1.TEXT)
    Call GRACONTROL(IDENTER$, "PORTREM", Combo2.TEXT)
    Call GRACONTROL("*", "PORTREM", Combo2.TEXT)
    Call GRACONTROL("", "COPIAFAC", Text5.TEXT)
    Call GRACONTROL("", "COPIAREM", Text4.TEXT)
    '\\\OT-05-0492-WAR-11/07/05///
     Call GRACONTROL("", "COEVADEC", Text6.TEXT)
    '\\\OT-05-0492-WAR-FIN
    MONUMERA$ = "AUTO"
    If Option3.Value = True Then
        MONUMERA$ = "MANUAL"
    End If
    Call GRACONTROL("", "NUMFACTU", MONUMERA$)
    '
    MONUME$ = "NORMAL": FORCAFAC% = 1
    If Option2.Value = True Then
        MONUME$ = "ITEMIZADO": FORCAFAC% = 2
      ElseIf Option24.Value = True Then
        MONUME$ = "SINREMITO": FORCAFAC% = 3
    End If
    Call GRACONTROL("FACTUMAI", "FORMACAR", MONUME$)
    '
    MONUME$ = "CORRELA": NUMECORR% = 1
    If Option6.Value = True Then
         MONUME$ = "INDEPEN": NUMECORR% = (-1)
       ElseIf Option17.Value = True Then
         MONUME$ = "UNIFICA": NUMECORR% = (-2)
    End If
    Call GRACONTROL("NUMEFACT", "NUMECORR", MONUME$)
    '
    PREFACTU$ = "REMITO"
    If Option7.Value = True Then
        PREFACTU$ = "ORDESPA"
      ElseIf Option8.Value = True Then
        PREFACTU$ = "PEDIDO"
    End If
    Call GRACONTROL("FACTURA", "PREFACTU", PREFACTU$)
    '
    NUMERA$ = "PREIMPRE"
    If Option10.Value = True Then NUMERA$ = "AUTONUME"
    Call GRACONTROL("FACTUREM", "IMPRENUM", NUMERA$)
    '
    TIFORMU$ = "CONTINUO"
    If Option11.Value = True Then TIFORMU$ = "HOJASUEL"
    Call GRACONTROL("FACTUREM", "FORMULAR", TIFORMU$)
    '
    MODESCUE$ = "PRENETO"
    If Option14.Value = True Then MODESCUE$ = "PORCIENTO"
    Call GRACONTROL("FACTUREM", "MODESCUE", MODESCUE$)
    '
    MODOMONE$ = "PESOS"
    If Option15.Value = True Then MODOMONE$ = "BIMONETA"
    If Option19.Value = True Then MODOMONE$ = "MULTIMON"
    Call GRACONTROL("GESVENTA", "MODOMONE", MODOMONE$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("FACTURA", "VERLICRE", MONUME$)
    '
    '\\\OT-08-735-OD-09/09/08///
    MONUME$ = "NO"
    If Check1.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("FACTURA", "VERISTOK", MONUME$)
    '\\\OT-08-735-OD-FIN///
    '
    '\\\OT-05-419-WAR-23/06/05///
    SUPCONF$ = "NO": CANCELCONF% = 0
    If Check3.Value = 1 Then SUPCONF$ = "SI": CANCELCONF% = 1
    Call GRACONTROL("FACTURA", "SUPCONF", SUPCONF$)
    '\\\OT-05-419-WAR-FIN///
    ' AGREGA DESPACHO DE IMPORTACION
    AGREDESP$ = "NO"
    If Check4.Value = 1 Then AGREDESP$ = "SI"
    Call GRACONTROL("FACTURA", "AGREDESP", AGREDESP$)
    '
    CANCECON$ = "NO"
    If Check5.Value = 1 Then CANCECON$ = "SI"
    Call GRACONTROL("FACTURA", "CANCECON", CANCECON$)
    '
    AGRUP$ = "NOAGRUPA"
    If Option21.Value = True Then AGRUP$ = "TRANSUCU"
    Call GRACONTROL("FACTUREM", "AGRUPEDI", AGRUP$)
    '
     '\\\OT-08-0076-FL-22/02/08/// '\\\OT-08-0076-FL-FIN///
    DEVICE$ = "PRINTER"
    If Option22.Value = True Then DEVICE$ = "CFISCAL"
    If Option26.Value = True Then DEVICE$ = "FELECT" ' FACTURA ELECTRONICA
    Call GRACONTROL("FACTURA", "SALIDA", DEVICE$)
    '\\\OT-08-0076-FL-FIN///
    '
    FLINUMERA% = FECHANUM(Label5)
    FFS$ = REGLEIDO$("FORMFAC", 2&)
    Call REPLA(FFS$, MKI$(FLINUMERA%), 19, 2)
    Call GRAREG("FORMFAC", FFS$, 2)
    Call CIERRARCH("FORMFAC")
    '
   
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada !\Re-Inicie Ahora el Módulo de Facturación.")
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label5 = VDEVU$
   End If
   Command3.Enabled = True
End Sub

'
Private Sub Command4_Click(index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(index).TEXT = VDEVU$
        DEFOR1(index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click(index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(Text1(index).TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CARDATCENCALC
    Command8.Enabled = True
    '
'    Call TRALOCAL("VENCICAI", "")
'    VTB% = VENTABU%("VENCICAI/NC/TITUPAN=Tabla de Vencimientos de CAI")
'    If VTB% >= 0 Then
'      JJ& = 0
'      For U& = 1 To ULTREG&("!VENCICAI")
'        X$ = REGLEIDO$("!VENCICAI", U&)
'        If TRIM$(Left$(X$, 32)) <> "" Then
'          JJ& = JJ& + 1
'          Call GRAREG("VENCICAI", X$, JJ&)
'        End If
'      Next U&
'      Call SETULTREG("VENCICAI", JJ&)
'    End If
'    Call CIERRARCH("VENCICAI")
'    Call CIERRARCH("!VENCICAI")
'    '
End Sub

Private Sub configuracionesespeciales_Click()
   Call Command11_Click
End Sub

Private Sub cotimoneda_Click()
   Call Command5_Click
End Sub

Private Sub Form_Load()
    
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2(0).BackColor = &H8000000F
      Me.Frame2(1).BackColor = &H8000000F
      Me.Frame2(2).BackColor = &H8000000F
      Me.Frame2(3).BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      Me.Frame13.BackColor = &H8000000F
      Me.Frame14.BackColor = &H8000000F
      Me.Frame15.BackColor = &H8000000F
      Me.Frame16.BackColor = &H8000000F
      Me.Frame17.BackColor = &H8000000F
      Me.Frame18.BackColor = &H8000000F
      Me.Frame19.BackColor = &H8000000F
      Me.Frame20.BackColor = &H8000000F
      
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      Me.Option9.BackColor = &H8000000F
      Me.Option10.BackColor = &H8000000F
      Me.Option11.BackColor = &H8000000F
      Me.Option12.BackColor = &H8000000F
      Me.Option13.BackColor = &H8000000F
      Me.Option14.BackColor = &H8000000F
      Me.Option15.BackColor = &H8000000F
      Me.Option16.BackColor = &H8000000F
      Me.Option17.BackColor = &H8000000F
      Me.Option18.BackColor = &H8000000F
      Me.Option19.BackColor = &H8000000F
      Me.Option20.BackColor = &H8000000F
      Me.Option21.BackColor = &H8000000F
      Me.Option22.BackColor = &H8000000F
      Me.Option23.BackColor = &H8000000F
      Me.Option24.BackColor = &H8000000F
      Me.Option26.BackColor = &H8000000F
     
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      
      MARCOBARRA.Top = 6800
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0

    End If
    '
    Call APLICACIONSKINS(Me)
    '
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
      x$ = REGLEIDO$("SUCURSAL", U&)
      CODEPO% = Asc(Mid$(x$, 64, 1))
      If CODEPO% > 0 Then
        If CODEPO% <> 99 Then
          If J& < 62 Then
            J& = J& + 1
            Y$ = REGBLAN$("TADEPOSI")
            Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
            Call REPLA(Y$, Left$(x$, 63), 2, 63)
            Call GRAREG("TADEPOSI", Y$, J&)
          End If
        End If
      End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text1(0).TEXT = CONTROL$("", "FORFAC-A")
    DEFOR1(0).Caption = ECOARCH$("INDIFRM", Text1(0).TEXT)
    Text1(4).TEXT = CONTROL$("", "FACTEXLI")
    DEFOR1(4).Caption = ECOARCH$("INDIFRM", Text1(4).TEXT)
    Text1(5).TEXT = CONTROL$("", "FACTELIB")
    DEFOR1(5).Caption = ECOARCH$("INDIFRM", Text1(5).TEXT)
    Text1(1).TEXT = CONTROL$("", "FORFAC-B")
    DEFOR1(1).Caption = ECOARCH$("INDIFRM", Text1(1).TEXT)
    Text1(2).TEXT = CONTROL$("", "FORFAC-E")
    DEFOR1(2).Caption = ECOARCH$("INDIFRM", Text1(2).TEXT)
    Text1(3).TEXT = CONTROL$("", "FOREMIFA")
    DEFOR1(3).Caption = ECOARCH$("INDIFRM", Text1(3).TEXT)
    '
    PFAC$ = CONTROL$(IDENTER$, "PORTFAC")
      If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTFAC")
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo1.TEXT = PFAC$
      End If
    Next
    '
    PREM$ = CONTROL$(IDENTER$, "PORTREM")
      If PREM = "" Then PREM$ = CONTROL$("*", "PORTREM")
    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PREM$ Then
        Combo2.TEXT = PREM$
      End If
    Next
    
    Text5.TEXT = CONTROL$("", "COPIAFAC")
    Text4.TEXT = CONTROL$("", "COPIAREM")
    Text7.TEXT = CONTROL$("", "DEPOEXPE")
    '\\\OT-05-0492-WAR-11/07/05///
     Text6.TEXT = CONTROL("", "COEVADEC")
    '\\\OT-05-0492-WAR-FIN
    Label1.Caption = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then Option3.Value = True
    If CONTROL$("NUMEFACT", "NUMECORR") = "INDEPEN" Then Option6.Value = True
    If CONTROL$("NUMEFACT", "NUMECORR") = "UNIFICA" Then Option17.Value = True
    If CONTROL$("FACTUMAI", "FORMACAR") = "ITEMIZADO" Then Option2.Value = True
    If CONTROL$("FACTUMAI", "FORMACAR") = "SINREMITO" Then Option24.Value = True
    If CONTROL$("FACTUREM", "IMPRENUM") = "AUTONUME" Then Option10.Value = True
    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then Option11.Value = True
    If CONTROL$("FACTUREM", "MODESCUE") = "PORCIENTO" Then Option14.Value = True
    If CONTROL$("GESVENTA", "MODOMONE") = "BIMONETA" Then Option15.Value = True
    If CONTROL$("GESVENTA", "MODOMONE") = "MULTIMON" Then Option19.Value = True
    '
    '
    PREFACTU$ = TRIM$(CONTROL("FACTURA", "PREFACTU"))
    If PREFACTU$ = "ORDESPA" Then Option7 = True
    If PREFACTU$ = "PEDIDO" Then Option8 = True
    '
    CMO$ = TRIM$(CONTROL$("", "CODFACTU")): If CMO$ <> "" Then Text9.TEXT = CMO$
    '
    '\\\OT-08-0765-OD-09/09/08///
    If CONTROL$("FACTURA", "VERISTOK") = "SI" Then Check1.Value = 1
    '\\\OT-08-0765-OD-FIN///
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then Check2.Value = 1
    If CONTROL$("FACTURA", "SUPCONF") = "SI" Then Check3.Value = 1
    If CONTROL$("FACTURA", "AGREDESP") = "SI" Then Check4.Value = 1
    If CONTROL$("FACTURA", "CANCECON") = "SI" Then Check5.Value = 1
    '
    If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then Option21 = True
    '

    '\\\OT-08-0076-FL-22/02/08///
    If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then Option22 = True
    If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then Option26 = True ' Factura Electronica
    If CONTROL("FACTURA", "SISTMANT") = "SI" Then Check6.Value = 1
    '\\\OT-08-0076-FL-FIN///
    '
    
    FFS$ = REGLEIDO$("FORMFAC", 2&)
    FEX = CVI(Mid$(FFS$, 19, 2))
    Label5 = FECHATEX$(FEX)
    '
End Sub

Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub Label5_DblClick()
   Call Command3_Click
End Sub

Private Sub motivosdeimputacion_Click()
     Call Command12_Click
End Sub

Private Sub Option26_DblClick()

      PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
      PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
      PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
      '
      Call REPLA(VDEF$, PVTANORMAL$, 1, 4)
      Call REPLA(VDEF$, PVTABCPTAL$, 5, 4)
      Call REPLA(VDEF$, PVTAFAEXP$, 9, 4)
      '
      'ABRE EL OBJETO
30    OBX$ = OBJPLA$("PUVEFAEL", VDEF$)
      '
      If OBX$ = "" Then Exit Sub
      '
      PVNAL$ = Mid$(OBX$, 1, 4) ' leo el objeto
      PVCPTL$ = Mid$(OBX$, 5, 4)
      PVTFXP$ = Mid$(OBX$, 9, 4)
      '
      ' valida que al menos uno tenga punto de venta
      If XVALO(PVNAL$) < 1 And XVALO(PVCPTL$) < 1 And XVALO(PVTFXP$) < 1 Then ' valido que tenga el punto de venta normal
         'Call MENSERR(24, "Punto de Venta Normal Mandatario")
         Call MENSERR(24, "Debe Ingresar Al Menos Un Punto de Venta")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      If Len(TRIM$(PVNAL$)) < 4 Then ' si es menor a 4 len doy formato
        While Len(TRIM$(PVNAL$)) < 4: PVNAL$ = "0" + PVNAL$: Wend
      End If
      '
      If XVALO(PVCPTL$) > 0 Then
        If Len(TRIM$(PVCPTL$)) < 4 Then '
          While Len(TRIM$(PVCPTL$)) < 4: PVCPTL$ = "0" + PVCPTL$: Wend
        End If
      End If
      '
      'FORMATEO PUNTO DE VENTA DE F.EXPORT
      If XVALO(PVTFXP$) > 0 Then
        If Len(TRIM$(PVTFXP$)) < 4 Then '
          While Len(TRIM$(PVTFXP$)) < 4: PVTFXP$ = "0" + PVTFXP$: Wend
        End If
      End If
      
      'VALIDACION QUE NO HAYA REPETICIONES DE PUNTO DE VENTA COMUN Y B.CAP
      If PVNAL$ = PVCPTL$ Then
         Call MENSERR(24, "Puntos de Venta Normal y de Bienes de Capital\No Pueden ser Iguales")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      'VALIDACION QUE NO HAYA REPETICIONES DE PUNTO DE VENTA FACT.EXP
      If XVALO(PVTFXP$) > 0 Then
        If PVNAL$ = PVTFXP$ Or PVCPTL$ = PVTFXP$ Then
           Call MENSERR(24, "Puntos de Venta Factura de Exportación\Debe Ser Diferente al P.Vta. Normal y al de Bienes de Capitales")
           VDEF$ = OBX$
           GoTo 30
        End If
      End If
      '
      'MENSAJE DE CONFIRMACION DE GRABACION
      TTXX$ = "Punto de Venta Normal: " & PVNAL$ & " \Punto de Venta Bienes de Capital: " & PVCPTL$ & " \Punto de Venta Factura de Exportación: " & PVTFXP$
      OPX% = COMALTER%(TTXX$ + "\\Aceptar \Modificar") ' Confirmacion
      If OPX% < 1 Or OPX% > 1 Then
        VDEF$ = OBX$
        GoTo 30
      End If
      '
      'GRABO
      Call GRACONTROL("FACTURA", "PUVENELE", PVNAL$)
      Call GRACONTROL("FACTURA", "PUVENCAP", PVCPTL$)
      Call GRACONTROL("FACTURA", "PUVEFAEX", PVTFXP$)
      '
End Sub

Private Sub parametrosvarios_Click()
   Call Command8_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_DblClick(index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(index).TEXT = VDEVU$
        DEFOR1(index).Caption = VALACT2$("INDIFRM")
    End If
End Sub



Private Sub Text8_Click()
    Text7.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub



Private Sub Text7_Change()
    '\\\OT-06-0570-RG-28/12/06///
    DEPO% = XVALO(Text7)
    Label1.Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1.Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1.Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'


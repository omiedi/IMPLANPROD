VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPFACTU07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Facturación"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11445
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
   ScaleHeight     =   7440
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   115
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame19 
      BackColor       =   &H00E2D3C0&
      Caption         =   "DEVICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9300
      TabIndex        =   99
      Top             =   0
      Width           =   1200
      Begin VB.CheckBox chkExportaSimple 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Ex.Simple"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   70
         TabIndex        =   135
         Top             =   660
         Width           =   1020
      End
      Begin VB.OptionButton Option26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Factura. Electrónica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   50
         TabIndex        =   109
         ToolTipText     =   "Doble Click para Configurar Punto de Venta (Normal - Bienes de Capital)"
         Top             =   1150
         Width           =   1100
      End
      Begin VB.OptionButton Option25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Imp. Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   50
         TabIndex        =   108
         Top             =   900
         Width           =   1100
      End
      Begin VB.OptionButton Option23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   50
         TabIndex        =   101
         Top             =   450
         Value           =   -1  'True
         Width           =   1100
      End
      Begin VB.OptionButton Option22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "C.Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   50
         TabIndex        =   100
         Top             =   200
         Width           =   1100
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   10560
      ScaleHeight     =   8490
      ScaleWidth      =   1875
      TabIndex        =   75
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   123
         Top             =   6720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   124
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command90 
         Caption         =   "Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   118
         ToolTipText     =   "Observaciones Mail Factura"
         Top             =   2415
         Width           =   630
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Imput"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Tabla de Motivos de Imputación"
         Top             =   5480
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Equiv"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   4790
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   91
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
         Picture         =   "OPFACTU07.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   3400
         Visible         =   0   'False
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
         Picture         =   "OPFACTU07.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Parámetros varios: Tasas IVA, Letra Exportación, Vencimientos de CAI"
         Top             =   2710
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":0832
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1665
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   980
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACTU07.frx":0E46
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   285
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PUERTOS DE IMPRESIÓN - COPIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5100
      TabIndex        =   37
      Top             =   0
      Width           =   4125
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Replica en Expedición"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2210
         TabIndex        =   114
         Top             =   1250
         Width           =   1890
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Fact. Electrónica por Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   80
         TabIndex        =   117
         Top             =   1250
         Width           =   2250
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   600
         Width           =   2715
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   260
         Width           =   2715
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Suprimir Redirección On Line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   80
         TabIndex        =   79
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Text4 
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
         Height          =   315
         Left            =   3720
         TabIndex        =   52
         Top             =   640
         Width           =   330
      End
      Begin VB.TextBox Text5 
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
         Height          =   315
         Left            =   3720
         TabIndex        =   51
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   0
         TabIndex        =   39
         Top             =   675
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   0
         TabIndex        =   38
         Top             =   345
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FORMULARIOS DE IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7425
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   4920
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FACTURACIÓN TIPO B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Index           =   1
         Left            =   155
         TabIndex        =   13
         Top             =   2445
         Width           =   4600
         Begin VB.CommandButton Command13 
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
            Height          =   255
            Index           =   1
            Left            =   170
            TabIndex        =   105
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   1750
            Width           =   2295
         End
         Begin VB.TextBox Text1 
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
            Height          =   300
            Index           =   5
            Left            =   2415
            TabIndex        =   84
            Top             =   1035
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
            TabIndex        =   83
            Top             =   1035
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   5
            Left            =   3780
            Picture         =   "OPFACTU07.frx":0F90
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   1035
            Width           =   650
         End
         Begin VB.TextBox Text1 
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
            Height          =   300
            Index           =   1
            Left            =   2415
            TabIndex        =   16
            Top             =   240
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   1
            Left            =   3780
            Picture         =   "OPFACTU07.frx":129A
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   260
            Width           =   650
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Material de Texto Libre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   0
            TabIndex        =   86
            Top             =   1125
            Width           =   2325
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   5
            Left            =   150
            TabIndex        =   85
            Top             =   1410
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Material Codificado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   630
            Width           =   3480
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FORMULARIO DE REMITO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Index           =   3
         Left            =   155
         TabIndex        =   31
         Top             =   5895
         Width           =   4600
         Begin VB.CommandButton Command13 
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
            Height          =   255
            Index           =   3
            Left            =   170
            TabIndex        =   107
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   960
            Width           =   2310
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   3
            Left            =   3780
            Picture         =   "OPFACTU07.frx":15A4
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
            Height          =   300
            Index           =   3
            Left            =   2415
            TabIndex        =   32
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
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
            Caption         =   "Formulario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   35
            Top             =   350
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FACTURACIÓN DE EXPORTACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Index           =   2
         Left            =   155
         TabIndex        =   25
         Top             =   4560
         Width           =   4600
         Begin VB.CommandButton Command13 
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
            Height          =   255
            Index           =   2
            Left            =   170
            TabIndex        =   106
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   960
            Width           =   2310
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   2
            Left            =   3780
            Picture         =   "OPFACTU07.frx":18AE
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
            Height          =   300
            Index           =   2
            Left            =   2415
            TabIndex        =   26
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
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
            Caption         =   "Formulario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   29
            Top             =   350
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FACTURACIÓN TIPO A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   0
         Left            =   155
         TabIndex        =   19
         Top             =   240
         Width           =   4600
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
            Index           =   6
            Left            =   3465
            TabIndex        =   138
            Top             =   1815
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Height          =   315
            Index           =   6
            Left            =   2415
            TabIndex        =   137
            Top             =   1800
            Width           =   1065
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
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
            Index           =   6
            Left            =   3780
            Picture         =   "OPFACTU07.frx":1BB8
            TabIndex        =   136
            Top             =   1800
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   104
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   1870
            Width           =   750
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   4
            Left            =   3780
            Picture         =   "OPFACTU07.frx":1EC2
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1035
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
            Top             =   1035
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Height          =   300
            Index           =   4
            Left            =   2415
            TabIndex        =   70
            Top             =   1035
            Width           =   1065
         End
         Begin VB.TextBox Text1 
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   0
            Left            =   3780
            Picture         =   "OPFACTU07.frx":21CC
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   260
            Width           =   650
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fact. Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   960
            TabIndex        =   140
            Top             =   1920
            Width           =   1365
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   2280
            TabIndex        =   139
            Top             =   1830
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   4
            Left            =   150
            TabIndex        =   74
            Top             =   1410
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Material de Texto Libre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   0
            TabIndex        =   73
            Top             =   1140
            Width           =   2325
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Material Codificado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   24
            Top             =   350
            Width           =   2010
         End
         Begin VB.Label DEFOR1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
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
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTROLES OPERATIVOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5880
      Left            =   5100
      TabIndex        =   0
      Top             =   1560
      Width           =   5400
      Begin VB.CommandButton Command16 
         Caption         =   "Combinación de Iva (Equipo y Accesorio)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   133
         Top             =   5520
         Width           =   3255
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Combinación de Iva(Equipo y Accesorio)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         ToolTipText     =   "Permite asociar accesorios a un equipo para que al generarar o.venta y/o factura el accesorio tome el iva del Equipo"
         Top             =   5550
         Width           =   3375
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Impr.Lote (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4030
         TabIndex        =   131
         Top             =   3480
         Width           =   1280
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PRECIO PEDIDO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   3915
         TabIndex        =   125
         Top             =   2940
         Width           =   1335
         Begin VB.CheckBox Check10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pesifica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   950
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONTADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   4350
         TabIndex        =   110
         ToolTipText     =   "Tilde Este Casillero para Habilitar Opción de Cancelación de Contado"
         Top             =   1400
         Width           =   900
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   285
            TabIndex        =   111
            ToolTipText     =   "Tilde Este Casillero para Habilitar Opción de Cancelación de Contado"
            Top             =   230
            Width           =   500
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E2D3C0&
         Caption         =   "AGRUPACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3910
         TabIndex        =   96
         Top             =   3840
         Width           =   1335
         Begin VB.OptionButton Option21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   98
            Top             =   520
            Width           =   1170
         End
         Begin VB.OptionButton Option20 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "No Agrupa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   97
            Top             =   240
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DESCRIPCIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2450
         TabIndex        =   92
         Top             =   3840
         Width           =   1400
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Especifica- ción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   90
            TabIndex        =   113
            Top             =   435
            Width           =   1215
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "+ Despacho"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   320
            Left            =   90
            TabIndex        =   93
            Top             =   160
            Width           =   1215
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E2D3C0&
         Caption         =   "%V.DEC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4350
         TabIndex        =   80
         ToolTipText     =   "Coeficiente de Valor Declarado"
         Top             =   830
         Width           =   900
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   2
            TabIndex        =   81
            ToolTipText     =   "Coeficiente de Valor Declarado"
            Top             =   210
            Width           =   400
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "NUMERACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   1470
         TabIndex        =   67
         Top             =   2940
         Width           =   1320
         Begin VB.OptionButton Option9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pre Impresa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   69
            Top             =   240
            Value           =   -1  'True
            Width           =   1150
         End
         Begin VB.OptionButton Option10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Autonumera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   68
            Top             =   520
            Width           =   1180
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FORMULARIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   120
         TabIndex        =   64
         Top             =   2940
         Width           =   1335
         Begin VB.OptionButton Option11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Hoja Suelta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   66
            Top             =   520
            Width           =   1120
         End
         Begin VB.OptionButton Option12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Continuo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   65
            Top             =   240
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DESC."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   2810
         TabIndex        =   61
         Top             =   2940
         Width           =   1095
         Begin VB.OptionButton Option13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pr. Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            MaskColor       =   &H00E2D3C0&
            TabIndex        =   63
            Top             =   240
            Value           =   -1  'True
            Width           =   980
         End
         Begin VB.OptionButton Option14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Porciento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            MaskColor       =   &H00E2D3C0&
            TabIndex        =   62
            Top             =   520
            Width           =   1000
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E2D3C0&
         Caption         =   "MONEDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   3720
         TabIndex        =   53
         Top             =   1875
         Width           =   1530
         Begin VB.OptionButton Option19 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Multi-Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   95
            Top             =   720
            Width           =   1320
         End
         Begin VB.OptionButton Option16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   55
            Top             =   240
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Bi-Monetario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   54
            Top             =   480
            Width           =   1320
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PRE-FACTURACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   1520
         TabIndex        =   48
         Top             =   1875
         Width           =   2160
         Begin VB.OptionButton Option18 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   60
            ToolTipText     =   "Por Remitos Pendientes de Facturación"
            Top             =   240
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton Option8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pedidos s/fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   50
            ToolTipText     =   "Por Pedidos Pendientes de Entrega Hasta Fecha de Corte"
            Top             =   480
            Width           =   1500
         End
         Begin VB.OptionButton Option7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Por Pedido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   80
            TabIndex        =   49
            ToolTipText     =   "Por Selección de Ordenes de Venta Pendientes de Entrega"
            Top             =   720
            Width           =   1110
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2D3C0&
         Caption         =   "EXPEDICIÓN - DEPÓSITO DE SALIDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   845
         Left            =   120
         TabIndex        =   43
         Top             =   4680
         Width           =   5130
         Begin VB.CommandButton Command15 
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
            TabIndex        =   128
            Top             =   480
            Width           =   175
         End
         Begin VB.TextBox Text2 
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
            Left            =   1155
            TabIndex        =   127
            Text            =   " "
            ToolTipText     =   "Utilizado como Depósito de Reposición de Facturación (Consume de Dep.Expedición si no hay stock toma de este."
            Top             =   480
            Width           =   540
         End
         Begin VB.TextBox Text7 
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
            Left            =   1155
            TabIndex        =   45
            Text            =   " "
            Top             =   180
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
            Top             =   180
            Width           =   175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pañol/Factur"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   130
            ToolTipText     =   "Utilizado como Depósito de Reposición de Facturación (Consume de Dep.Expedición si no hay stock toma de este."
            Top             =   555
            Width           =   1065
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1995
            TabIndex        =   129
            ToolTipText     =   "Utilizado como Depósito de Reposición de Facturación (Consume de Dep.Expedición si no hay stock toma de este."
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1995
            TabIndex        =   47
            Top             =   180
            Width           =   2775
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   46
            Top             =   255
            Width           =   1065
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "NUMERACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3750
         Begin VB.CommandButton Command14 
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
            Left            =   3520
            TabIndex        =   119
            Top             =   1280
            Width           =   150
         End
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
            Left            =   3520
            TabIndex        =   58
            Top             =   930
            Width           =   150
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2D3C0&
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   1575
            TabIndex        =   9
            Top             =   190
            Width           =   2115
            Begin VB.OptionButton Option17 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2D3C0&
               Caption         =   "Unif. Débito y Credito"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   440
               Width           =   2010
            End
            Begin VB.OptionButton Option6 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2D3C0&
               Caption         =   "Independiente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   120
               TabIndex        =   11
               Top             =   225
               Width           =   1590
            End
            Begin VB.OptionButton Option5 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2D3C0&
               Caption         =   "Correlativa"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   120
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   315
            TabIndex        =   8
            Top             =   255
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Automática"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   315
            TabIndex        =   7
            Top             =   510
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            TabIndex        =   122
            ToolTipText     =   "Doble Click para Limpiar Casillero"
            Top             =   1280
            Width           =   360
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Graba Cotización en Mo- neda Extranjera (sólo $)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   150
            TabIndex        =   121
            Top             =   1155
            Width           =   1820
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2445
            TabIndex        =   120
            ToolTipText     =   "Selección de Moneda a Utilizar en Texto en Factura de Valorización en Dolares"
            Top             =   1280
            Width           =   1080
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2445
            TabIndex        =   57
            Top             =   950
            Width           =   1080
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   150
            TabIndex        =   56
            Top             =   960
            Width           =   2145
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3730
            Y1              =   870
            Y2              =   870
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CÓDIGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4350
         TabIndex        =   4
         Top             =   240
         Width           =   900
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   190
            TabIndex        =   5
            Text            =   "FC"
            Top             =   210
            Width           =   400
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FORMATO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   120
         TabIndex        =   1
         Top             =   1875
         Width           =   1365
         Begin VB.OptionButton Option24 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Base Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   102
            Top             =   720
            Width           =   1200
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Itemizado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1100
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "VERIFICACIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   2265
         Begin VB.CheckBox Check13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pos. Iva c/Afip"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            TabIndex        =   134
            Top             =   405
            Width           =   1035
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Vend."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1005
            TabIndex        =   112
            Top             =   200
            Width           =   795
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Considera Reservas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   50
            TabIndex        =   41
            ToolTipText     =   "Considerar Reservas al Verificar Stocks de Expedición"
            Top             =   405
            Width           =   1050
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   50
            TabIndex        =   42
            Top             =   200
            Width           =   950
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPFACTU07.frx":24D6
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPFACTU07.frx":270A
      TabIndex        =   116
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transacciones"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu mnuCotizacionesMoneda 
         Caption         =   "Tabla de Monedas y Cotizaciones"
      End
      Begin VB.Menu mnuPuntosVentaExportaSimple 
         Caption         =   "Puntos de Venta Exporta Simple"
      End
      Begin VB.Menu mnuObsConfiguradas 
         Caption         =   "Observaciones Preconfiguradas"
      End
      Begin VB.Menu mnuFacturaCredito 
         Caption         =   "Factura de Crédito"
         Begin VB.Menu mnuMontoMinimoFCCredito 
            Caption         =   "Monto Mínimo de Facturación"
         End
         Begin VB.Menu mnuCBU 
            Caption         =   "CBU"
         End
      End
   End
End
Attribute VB_Name = "OPFACTU07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CargandoFormulario%



Private Sub chkExportaSimple_Click()

    If chkExportaSimple.Value = 0 Or CargandoFormulario% = 1 Then Exit Sub
    Resultado% = CLACONTRA%()
    If Resultado% = 0 Then
        chkExportaSimple.Value = 0
    Else
        mnuPuntosVentaExportaSimple.Visible = True
    End If
    
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

'\\\OT-06-0339-MH-24/08/06///
Sub Command10_Click()
Command10.Enabled = False
Call FINAL("*CODORCL")
Command10.Enabled = True
End Sub
'\\\OT-06-0339-MH-FIN///

Sub Command11_Click()
   '
   Command11.Enabled = False
   OPFACES00.Show 1
   Command11.Enabled = True
   '
End Sub

Sub Command12_Click()
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

Private Sub Command13_Click(Index As Integer)
    Command13(Index).Enabled = False
    Call TRALOFIL("LEYFAC", 1, MKI$(Index))
    VTB% = VENTABU%("LEYFAC/TITUPAN=" + DEFOR1(Index))
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("LEYFAC")
        XX$ = REGLEIDO$("LEYFAC", U&)
        If CVI(Left$(XX$, 2)) <> Index Then
          JJ& = JJ& + 1
          Call GRAREG("LEYFAC", XX$, JJ&)
        End If
      Next U&
      For U& = 1 To ULTREG&("!LEYFAC")
        XX$ = REGLEIDO$("!LEYFAC", U&)
        Call REPLA(XX$, MKI$(Index), 1, 2)
        JJ& = JJ& + 1
        Call GRAREG("LEYFAC", XX$, JJ&)
      Next U&
      Call SETULTREG("LEYFAC", JJ&)
    End If
    Command13(Index).Enabled = True

End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call SELECHO("TAMONED")
   If XVALO(VALACT1$("TAMONED")) > 0 Then
        MONEMX% = XVALO(VALACT1$("TAMONED"))
        Label8 = MONEMX%
   End If
   '
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text2.TEXT = VDEVU$
        Label1(1).Caption = VALACT2$("TADEPOSI")
    End If

End Sub

Private Sub Command16_Click()
   If CLACONTRA% > 0 Then
      Command16.Value = False
   End If
End Sub

Private Sub Command2_Click()
    '
    If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then 'si estaba configurada a modo electronica
      If Option26.Value = False Then   ' y se cambio a otro modo avisa y pide confirmacion
       OPX% = COMALTER%("Confiramación de Cambio en Modalidad de Facturación\Se ha Cambiado De Factura Electrónica a Modo Printer\\Cancelar\Cambiar a Modo Printer")
       If OPX% < 2 Then Exit Sub
      End If
    End If
    '
    Call GRACONTROL("", "FORFAC-A", Text1(0).TEXT)
    Call GRACONTROL("", "FACTEXLI", Text1(4).TEXT)
    Call GRACONTROL("", "FACCRE-A", Text1(6).TEXT)
    Call GRACONTROL("", "FORFAC-B", Text1(1).TEXT)
    Call GRACONTROL("", "FORFAC-E", Text1(2).TEXT)
    Call GRACONTROL("", "FACTELIB", Text1(5).TEXT)
    Call GRACONTROL("", "FOREMIFA", Text1(3).TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text7.TEXT)
    Call GRACONTROL("", "DPREPPAÑ", Text2.TEXT)
    
    Call GRACONTROL("", "CODFACTU", Text9.TEXT)
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
    Call GRACONTROL("FACTURA", "VERIRESE", MONUME$) ' CONSIDERA RESERVAS
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
    If Option25.Value = True Then DEVICE$ = "IFISCAL" ' Impresora Fiscal
    If Option26.Value = True Then DEVICE$ = "FELECT" ' FACTURA ELECTRONICA
    Call GRACONTROL("FACTURA", "SALIDA", DEVICE$)
    '\\\OT-08-0076-FL-FIN///
    '
    VerVend$ = "NO" ' validacion vendedor
    If Check6.Value = 1 Then VerVend$ = "SI"
    Call GRACONTROL("FACTURA", "VERIVEN", VerVend$)
    '
    Especi$ = "" ' CONTROL DE SI IMPRIME ESPECIFICACIONES
    If Check7.Value = 1 Then Especi$ = "SI"
    Call GRACONTROL("FACTUREM", "INCLUSPE", Especi$)
    '
    IMPREMDES$ = ""
    If Check8.Value = 1 Then IMPREMDES$ = "SI" ' CONTROL PARA IMPRIMIR COPIA DE REMITO EN IMPRESORA DE DESPACHO
    Call GRACONTROL("FACTURA", "REPLIREM", IMPREMDES$)
    '
    MONUME$ = "NO"
    If Check12.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("FACTURA", "VALITIVA", MONUME$)
    
    ENVIAMAILFACT$ = ""
    If Check9.Value = 1 Then ENVIAMAILFACT$ = "SI" ' CONTROL PARA ENVIO POR MAIL DE FACTURA
    Call GRACONTROL("FACTURA", "ENVFAMAIL", ENVIAMAILFACT$)
    '
    MONUME$ = 0
    If XVALO(Label8) > 0 Then MONUME$ = TRIM$(Label8)
    Call GRACONTROL("FACTURA", "MONTEXFA", MONUME$)
    '
    MONUME$ = ""
    If Check10.Value = 1 Then MONUME$ = "SI" ' CONTROL PARA ENVIO POR MAIL DE FACTURA
    Call GRACONTROL("FACTURA", "PESPREPE", MONUME$)
    
    MONUME$ = ""
    If Check13.Value = 1 Then MONUME$ = "SI" ' VALIDACION POSICION DE IVA CON AFIP
    Call GRACONTROL("FACTURA", "VAPOSIVA", MONUME$)
    
    Call GRACONTROL$("FACTURA", "IMPRLOTE", SAFETEXT$(Check11.Value))
    
    FLINUMERA% = FECHANUM(Label5)
    FFS$ = REGLEIDO$("FORMFAC", 2&)
    Call REPLA(FFS$, MKI$(FLINUMERA%), 19, 2)
    Call GRAREG("FORMFAC", FFS$, 2)
    Call CIERRARCH("FORMFAC")
    '
    VALO$ = ""
    If chkExportaSimple.Value = 1 Then VALO$ = "SI"
    Call GRACONTROL$("FACTURA", "EXSIMPLE", VALO$)
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
Private Sub Command4_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        DEFOR1(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Sub Command5_Click()
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

Private Sub Command6_Click(Index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(Text1(Index).TEXT), "Formularios de Impresion")
End Sub

Sub Command8_Click()
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

Private Sub Command90_Click()
  Call TEXTO_MAIL_FACTURA
End Sub
Sub TEXTO_MAIL_FACTURA()
      '
      Call COPYSTRU("OBSERVOF", "TXMAIFAC")
      Call TRACENTRAL("TXMAIFAC", "")           ' DE LOCAL A CENTRAL
      Call TRALOCAL("TXMAIFAC", "OBSERVOF")     ' DE CENTRAL A LOCAL
      VTB% = VENTABU%("OBSERVOF/TITUPAN=Texto de Mail para Facturas")
      If VTB% < 0 Then Exit Sub
      '
      For U& = 1 To ULTREG&("!OBSERVOF")
           TX$ = REGLEIDO$("!OBSERVOF", U&)
           Call REPLA(x$, TX$, 1, 64)
           Call GRAREG("!TXMAIFAC", x$, U&)
      Next U&
      '
      Call SETULTREG("!TXMAIFAC", U& - 1)
      Call TRACENTRAL("TXMAIFAC", "")           ' DE LOCAL A CENTRAL
      Call TRALOCAL("TXMAIFAC", "")             ' DE CENTRAL A LOCAL
      Call CIERRARCH("!TXMAIFAC")
      Call CIERRARCH("OBSERVOF")
      '
End Sub



Private Sub Form_Load()
    '
    CargandoFormulario% = 1
    Call APLICACIONSKINS(Me, Picture1)
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
    If CONTROL("FACTURA", "VALITIVA") = "SI" Then
       Command16.Visible = False
    End If
    

    Text1(0).TEXT = CONTROL$("", "FORFAC-A")
    DEFOR1(0).Caption = ECOARCH$("INDIFRM", Text1(0).TEXT)
    Text1(4).TEXT = CONTROL$("", "FACTEXLI")
    DEFOR1(4).Caption = ECOARCH$("INDIFRM", Text1(4).TEXT)
    Text1(6).TEXT = CONTROL$("", "FACCRE-A")
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
    Text2.TEXT = CONTROL$("", "DPREPPAÑ")
    
    '\\\OT-05-0492-WAR-11/07/05///
     Text6.TEXT = CONTROL("", "COEVADEC")
    '\\\OT-05-0492-WAR-FIN
    Label1(0).Caption = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    Label1(1).Caption = ECOARCH$("TADEPOSI", Chr$(Val(Text2.TEXT)))
    
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
    If CONTROL$("FACTURA", "VERIRESE") = "SI" Then Check1.Value = 1
    Check11.Value = XVALO(CONTROL$("FACTURA", "IMPRLOTE"))
    
    '\\\OT-08-0765-OD-FIN///
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then Check2.Value = 1
    If CONTROL$("FACTURA", "SUPCONF") = "SI" Then Check3.Value = 1
    If CONTROL$("FACTURA", "AGREDESP") = "SI" Then Check4.Value = 1
    If CONTROL$("FACTURA", "CANCECON") = "SI" Then Check5.Value = 1
    If CONTROL("FACTURA", "VERIVEN") = "SI" Then Check6.Value = 1
    If CONTROL("FACTURA", "ENVFAMAIL") = "SI" Then Check9.Value = 1
    If CONTROL("FACTURA", "VALITIVA") = "SI" Then Check12.Value = 1
    '
    If CONTROL$("FACTUREM", "AGRUPEDI") = "TRANSUCU" Then Option21 = True
    '
    '\\\OT-08-0076-FL-22/02/08///
    If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then Option22 = True
    '\\\OT-08-0076-FL-FIN///
    '
    Check7.Value = 0
    If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then Check7.Value = 1
    '
    If CONTROL$("FACTURA", "SALIDA") = "IFISCAL" Then Option25 = True ' Impresora Fiscal
    If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then Option26 = True ' Factura Electronica
    '
    Check8.Value = 0 ' CONTROL PARA IMPRIMIR COPIA DE REMITO EN IMPRESORA DE DESPACHO
    If CONTROL("FACTURA", "REPLIREM") = "SI" Then Check8.Value = 1
    '
    Check10.Value = 0
    If CONTROL("FACTURA", "PESPREPE") = "SI" Then Check10.Value = 1
    '
    Check1.Value = 0
    If CONTROL("FACTURA", "VAPOSIVA") = "SI" Then Check13.Value = 1
    
    FFS$ = REGLEIDO$("FORMFAC", 2&)
    FEX = CVI(Mid$(FFS$, 19, 2))
    Label5 = FECHATEX$(FEX)
    'LABEL REGISTRA LA MONEDA A UTILIZAR EN EL TEXTO DE LAS FACTURAS QUE SE AGREGA INDICANDO A CUANTO EQUIVALEN EN U$S EL TOTAL DE LA FC EN $
    Label8 = SAFETEXT$(XVALO(CONTROL("FACTURA", "MONTEXFA")))
    '
    If CONTROL$("FACTURA", "EXSIMPLE") = "SI" Then
        chkExportaSimple.Value = 1
        mnuPuntosVentaExportaSimple.Visible = True
    End If
    '
    CargandoFormulario% = 0
    '
End Sub

Private Sub Frame20_DragDrop(Source As CONTROL, x As Single, Y As Single)
14865344
End Sub

Private Sub Label5_DblClick()
   Call Command3_Click
End Sub

Private Sub Label8_Change()
   Label6 = ""
   If XVALO(Label8) > 0 Then
      Label6 = ECOARCH$("TAMONED", MKI$(XVALO(Label8)))
   End If
End Sub

Private Sub Label8_DblClick()
   Label8 = ""
End Sub


Private Sub mnuCBU_Click()

    CBUActual$ = TRIM$(CONTROL$("FACTURA", "CBUCREDI"))
    CBU$ = InputBox("CBU", "CBU", CBUActual$)
    Call GRACONTROL("FACTURA", "CBUCREDI", TRIM$(CBU$))
    
End Sub

Private Sub mnuCotizacionesMoneda_Click()
   Call Me.Command5_Click
End Sub

Private Sub mnuMontoMinimoFCCredito_Click()
    
    MontoMinimoActual$ = TRIM$(CONTROL$("FACTURA", "MOMINCRE"))
    MontoMinimo# = Val(InputBox("Mínimo", "Monto Mínimo de Facturación", MontoMinimoActual$))
    Call GRACONTROL("FACTURA", "MOMINCRE", TRIM$(FORMATNUM$(MontoMinimo#, "F10.2")))

End Sub

Private Sub mnuObsConfiguradas_Click()
   CargaObservacionesPredeterminadas
End Sub

Private Sub mnuPuntosVentaExportaSimple_Click()

  Call GENERAR_ESTRUCTURAS_SQL
  IDLIST& = ObtenerIdTablaAuxiliar("PTOVTAEXSIM")
  
  Call ABM_TABLA(IDLIST&)
  
End Sub



Private Sub Option10_Click()
14865344
End Sub

Private Sub Option2_Click()
14865344
End Sub

Private Sub Option22_Click()
    Call COMUNI("Opcion No Disponible en la Presente Version")
    Option23.Value = True
End Sub

Private Sub Option23_DblClick()
    
      PVTANORMALPRINT$ = CONTROL("FACTURA", "PUVENPRI") 'P VTA. NORMAL MODO PRINTER
      PVTAFAEXPRINT$ = CONTROL("FACTURA", "PUVEXPRI")  'P.VTA. FACT.EXP. MODO PRINTER
      '
      Call REPLA(VDEF$, PVTANORMALPRINT$, 1, 4)
      Call REPLA(VDEF$, PVTAFAEXPRINT$, 5, 4)
      '
      'ABRE EL OBJETO
30    OBX$ = OBJPLA$("PUVENPRI", VDEF$)
      '
      If OBX$ = "" Then Exit Sub
      '
      PVTANORMALPRINT$ = Mid$(OBX$, 1, 4) ' leo el objeto
      PVTAFAEXPRINT$ = Mid$(OBX$, 5, 4)
      '
      ' valida que al menos uno tenga punto de venta
      If XVALO(PVTANORMALPRINT$) < 1 And XVALO(PVTAFAEXPRINT$) < 1 Then ' valido que tenga el punto de venta normal
         'Call MENSERR(24, "Punto de Venta Normal Mandatario")
         Call MENSERR(24, "Debe Ingresar Al Menos Un Punto de Venta")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      If Len(TRIM$(PVTANORMALPRINT$)) < 4 Then ' si es menor a 4 len doy formato
        While Len(TRIM$(PVTANORMALPRINT$)) < 4: PVTANORMALPRINT$ = "0" + PVTANORMALPRINT$: Wend
      End If
      '
      If XVALO(PVTAFAEXPRINT$) > 0 Then
        If Len(TRIM$(PVTAFAEXPRINT$)) < 4 Then '
          While Len(TRIM$(PVTAFAEXPRINT$)) < 4: PVTAFAEXPRINT$ = "0" + PVTAFAEXPRINT$: Wend
        End If
      End If
      '
      
      'VALIDACION QUE NO HAYA REPETICIONES DE PUNTO DE VENTA COMUN Y B.CAP
      If PVTANORMALPRINT$ = PVTAFAEXPRINT$ Then
         Call MENSERR(24, "Puntos de Venta Normal y Factura de Exportación\No Pueden ser Iguales")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      'VALIDO QUE NO SE CHOQUEN CON LOS NUMEROS DE FACTURA ELECTRONICA
      PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
      PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
      PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
      If XVALO(PVTANORMALPRINT$) = XVALO(PVTANORMAL$) Or XVALO(PVTANORMALPRINT$) = XVALO(PVTABCPTAL$) Or XVALO(PVTANORMALPRINT$) = XVALO(PVTAFAEXP$) Then
         Call MENSERR(24, "Puntosde Venta Normal Ingresado se Estan Repitiendo con los de Factura Electrónica")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      If XVALO(PVTAFAEXPRINT$) = XVALO(PVTANORMAL$) Or XVALO(PVTAFAEXPRINT$) = XVALO(PVTABCPTAL$) Or XVALO(PVTAFAEXPRINT$) = XVALO(PVTAFAEXP$) Then
         Call MENSERR(24, "Punto de Venta Exportación Ingresado se Estan Repitiendo con los de Factura Electrónica")
         VDEF$ = OBX$
         GoTo 30
      End If
     
      '
      'MENSAJE DE CONFIRMACION DE GRABACION
      TTXX$ = "Punto de Venta Normal (Printer): " & PVTANORMALPRINT$ & " \Punto de Venta Exportación (Printer): " & PVTAFAEXPRINT$
      OPX% = COMALTER%(TTXX$ + "\\Aceptar \Modificar") ' Confirmacion
      If OPX% < 1 Or OPX% > 1 Then
        VDEF$ = OBX$
        GoTo 30
      End If
      '
      'GRABO
      Call GRACONTROL("FACTURA", "PUVENPRI", TRIM$(PVTANORMALPRINT$))
      Call GRACONTROL("FACTURA", "PUVEXPRI", TRIM$(PVTAFAEXPRINT$))
      '

End Sub


Private Sub Option24_Click()
14865344
End Sub

Private Sub Option25_Click()
    Call COMUNI("Opcion No Disponible en la Presente Version")
    Option23.Value = True
End Sub

Private Sub Option26_DblClick()

      PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
      PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
      PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
      PVTMODOCREDITO$ = CONTROL("FACTURA", "PUVECRED") ' P.VTA. MODO CREDITO
      '
      Call REPLA(VDEF$, PVTANORMAL$, 1, 4)
      Call REPLA(VDEF$, PVTABCPTAL$, 5, 4)
      Call REPLA(VDEF$, PVTAFAEXP$, 9, 4)
      Call REPLA(VDEF$, PVTMODOCREDITO$, 13, 4)
      '
      'ABRE EL OBJETO
30    OBX$ = OBJPLA$("PUVEFAEL", VDEF$)
      '
      If OBX$ = "" Then Exit Sub
      '
      PVNAL$ = Mid$(OBX$, 1, 4) ' leo el objeto
      PVCPTL$ = Mid$(OBX$, 5, 4)
      PVTFXP$ = Mid$(OBX$, 9, 4)
      PVTMODOCREDITO$ = Mid$(OBX$, 13, 4)
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
      
      If Len(TRIM$(PVTMODOCREDITO$)) < 4 Then ' si es menor a 4 len doy formato
        While Len(TRIM$(PVTMODOCREDITO$)) < 4: PVTMODOCREDITO$ = "0" + PVTMODOCREDITO$: Wend
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
      Call GRACONTROL("FACTURA", "PUVECRED", PVTMODOCREDITO$)
      '
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        DEFOR1(Index).Caption = VALACT2$("INDIFRM")
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

Private Sub Text2_Change()
    DEPO% = XVALO(Text2)
    Label1(1).Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))

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
    Label1(0).Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1(0).Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1(0).Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPORFA07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Orden de Fabricación"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11505
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
   ScaleHeight     =   6600
   ScaleWidth      =   11505
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   6735
      Left            =   10605
      ScaleHeight     =   6675
      ScaleWidth      =   1320
      TabIndex        =   39
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command13 
         Height          =   650
         Left            =   120
         Picture         =   "OPORFA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Tabla de Operarios"
         Top             =   4560
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Tabla de Secciones Productivas"
         Top             =   3855
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   45
         Top             =   6255
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   300
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1785
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   990
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprime Formulario de Configuración"
         Top             =   3195
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORFA07.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Otras Opciones de Configuración"
         Top             =   2475
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   20
      TabIndex        =   0
      Top             =   90
      Width           =   10560
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   73
         Top             =   5415
         Width           =   4950
         Begin VB.CommandButton Command14 
            Height          =   240
            Left            =   4320
            Picture         =   "OPORFA07.frx":1526
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Setup de Trazabilidad con Códigos Qr"
            Top             =   180
            Width           =   525
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Imprime Etiqueta de O.Fabricación ??"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   460
            Left            =   30
            TabIndex        =   74
            Top             =   360
            Width           =   1545
         End
         Begin VB.CheckBox Check20 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Imprime Etiqueta de Recepción"
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
            Height          =   360
            Left            =   75
            TabIndex        =   77
            Top             =   405
            Width           =   1575
         End
         Begin VB.CheckBox Check19 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Solo Items Trazables "
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
            Height          =   360
            Left            =   1680
            TabIndex        =   76
            Top             =   405
            Width           =   1095
         End
         Begin VB.TextBox Text15 
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
            Height          =   300
            Left            =   4320
            TabIndex        =   75
            ToolTipText     =   "Ingresar Cantidad de Columnas que posee el rollo donde se imprime la etiqueta."
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Columnas de Etiquetas en Rollo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   2760
            TabIndex        =   78
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5295
         TabIndex        =   65
         Top             =   5415
         Width           =   5190
         Begin VB.TextBox TEXT13 
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
            Left            =   1680
            TabIndex        =   71
            Text            =   " "
            Top             =   240
            Width           =   1380
         End
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
            Height          =   300
            Left            =   3060
            TabIndex        =   70
            Top             =   240
            Width           =   180
         End
         Begin VB.TextBox Text14 
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
            Left            =   90
            TabIndex        =   69
            TabStop         =   0   'False
            Text            =   " "
            Top             =   600
            Width           =   3150
         End
         Begin VB.CommandButton Command16 
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
            Height          =   650
            Left            =   4455
            Picture         =   "OPORFA07.frx":1670
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   240
            Width           =   645
         End
         Begin VB.OptionButton Option14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Unidad"
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
            Left            =   3540
            TabIndex        =   67
            ToolTipText     =   "Imprime una Etiqueta por cada Unidad"
            Top             =   300
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton Option13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Bulto"
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
            Left            =   3540
            TabIndex        =   66
            ToolTipText     =   "Imprime una Etiqueta Por Bulto"
            Top             =   640
            Width           =   900
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Etiqueta de Recepción"
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
            Left            =   -195
            TabIndex        =   72
            Top             =   360
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   3190
         TabIndex        =   51
         Top             =   1230
         Width           =   2445
         Begin VB.OptionButton Option11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Trazable+Ubicación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   50
            TabIndex        =   56
            Top             =   655
            Value           =   -1  'True
            Width           =   1950
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ubicaciones"
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
            Left            =   50
            TabIndex        =   53
            Top             =   420
            Width           =   1380
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Estructura"
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
            Left            =   50
            TabIndex        =   52
            Top             =   195
            Width           =   1380
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Scrap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   6765
         TabIndex        =   48
         Top             =   2265
         Width           =   1575
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "M.Primas"
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
            Left            =   180
            TabIndex        =   50
            Top             =   315
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Producto"
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
            Left            =   180
            TabIndex        =   49
            Top             =   580
            Width           =   1200
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "M.Prima Unica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   8670
         TabIndex        =   36
         Top             =   2265
         Width           =   1815
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "V.Materiales"
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
            Left            =   270
            TabIndex        =   38
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Form.Ppal"
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
            Left            =   270
            TabIndex        =   37
            Top             =   580
            Width           =   1380
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
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
         Height          =   920
         Left            =   1750
         TabIndex        =   34
         Top             =   1230
         Width           =   1365
         Begin VB.CheckBox Check18 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Datos MRP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "Imprime Necesidad de MRP en O.F. (Se deberá configurar también plantilla de Formulario)"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Impresion Fantasma"
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
            Left            =   120
            TabIndex        =   35
            Top             =   200
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documentos Anexos (Requiere Sistema de Calidad)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   28
         Top             =   2265
         Width           =   6345
         Begin VB.CheckBox Check12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hoja de Inspección"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4860
            TabIndex        =   32
            Top             =   360
            Width           =   1305
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hojas de Procesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3420
            TabIndex        =   31
            Top             =   360
            Width           =   1485
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especificaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1350
            TabIndex        =   30
            Top             =   315
            Width           =   1905
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Plano"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   270
            TabIndex        =   29
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command6 
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
         Left            =   9660
         Picture         =   "OPORFA07.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   285
         Width           =   765
      End
      Begin VB.TextBox Text4 
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
         Left            =   5505
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   4080
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
         Height          =   360
         Left            =   9285
         TabIndex        =   21
         Top             =   285
         Width           =   300
      End
      Begin VB.TextBox Text3 
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
         Left            =   7920
         TabIndex        =   20
         Text            =   " "
         Top             =   285
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fórmula de Consumo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   6780
         TabIndex        =   14
         Top             =   1230
         Width           =   3690
         Begin VB.CommandButton Command12 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2040
            Picture         =   "OPORFA07.frx":1C84
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Nivel Máximo Modo Lineal y Comienzo de Linealizada"
            Top             =   620
            Width           =   405
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lineal + Linealizada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   59
            Top             =   660
            Width           =   2325
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Multinivel Linealizada"
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
            Left            =   120
            TabIndex        =   16
            Top             =   420
            Width           =   2325
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lineal (Un Nivel de Explosión)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   3120
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   240
         TabIndex        =   11
         Top             =   3230
         Width           =   10260
         Begin VB.CheckBox Check17 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cierre de Fabricación con Control de Calidad"
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
            Left            =   525
            TabIndex        =   62
            Top             =   1720
            Width           =   4350
         End
         Begin VB.CheckBox Check16 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Linealizada - Consulta Impresión Subconjunto"
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
            Left            =   5205
            TabIndex        =   61
            Top             =   1530
            Width           =   4110
         End
         Begin VB.CheckBox Check15 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Obviar Control de Repetición de Componentes"
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
            Left            =   5205
            TabIndex        =   58
            Top             =   1000
            Width           =   4150
         End
         Begin VB.CommandButton Command11 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            Picture         =   "OPORFA07.frx":1DCE
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Tabla de Relaciones Código - Sector"
            Top             =   240
            Width           =   405
         End
         Begin VB.CheckBox Check13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exige Declarar Sector Productivo"
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
            Left            =   5205
            TabIndex        =   33
            Top             =   1280
            Width           =   4110
         End
         Begin VB.CheckBox Check8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Múltiplos de Lotes Económicos (Tandas)"
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
            Left            =   5205
            TabIndex        =   27
            Top             =   740
            Width           =   4110
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Vale Individual por Operación"
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
            Left            =   5205
            TabIndex        =   26
            Top             =   460
            Width           =   4110
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fracciona por Lotes Máximos"
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
            Left            =   5205
            TabIndex        =   25
            Top             =   215
            Width           =   4110
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes Consumidos al Cierre"
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
            Left            =   525
            TabIndex        =   19
            Top             =   1455
            Width           =   4350
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes al Lanzamiento"
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
            Left            =   525
            TabIndex        =   18
            Top             =   1170
            Width           =   4110
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Verifica Componentes al Lanzamiento"
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
            Left            =   525
            TabIndex        =   17
            Top             =   885
            Width           =   4110
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Observaciones - Leyendas - Comentarios"
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
            Left            =   525
            TabIndex        =   13
            Top             =   600
            Width           =   4110
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Generar O/F's por Subconjuntos"
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
            Left            =   525
            TabIndex        =   12
            Top             =   315
            Width           =   3150
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2625
         TabIndex        =   10
         Text            =   " "
         Top             =   285
         Width           =   1380
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
         Height          =   285
         Left            =   3990
         TabIndex        =   9
         Top             =   285
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3855
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
         Left            =   4305
         Picture         =   "OPORFA07.frx":1F18
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   285
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   210
         TabIndex        =   3
         Top             =   1230
         Width           =   1485
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Manual"
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
            TabIndex        =   5
            Top             =   315
            Width           =   1300
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Automatica"
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
            TabIndex        =   4
            Top             =   580
            Value           =   -1  'True
            Width           =   1300
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   5730
         TabIndex        =   1
         Top             =   1230
         Width           =   825
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            TabIndex        =   2
            Text            =   "OF"
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Auxiliar:"
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
         Left            =   5505
         TabIndex        =   24
         Top             =   390
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Principal:"
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
         Left            =   210
         TabIndex        =   6
         Top             =   390
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPORFA07.frx":2222
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPORFA07.frx":2456
      TabIndex        =   55
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "OPORFA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

    If Check1.Value = 1 Then
        Command11.Enabled = True
    Else
        Command11.Enabled = False
    End If
    
End Sub

Private Sub Check16_Click()

    If Option4.Value = False And Option12.Value = False Then
        Check16.Value = False
    End If

End Sub



Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    'OT-05-0728-WAR-08/11/05///
    Call CREATABLA_ARTICULOS
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    
    Call TRALOCAL("SECPROD", "SECPROD")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(str$(Left + 200))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(str$(Top + 850))
    VTB% = VENTABU%("SECPROD1/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!SECPROD")
         X$ = REGLEIDO$("!SECPROD", U&)
         JJ& = JJ& + 1
         Call GRAREG("SECPROD", X$, JJ&)
         
      Next U&
      Call SETULTREG("SECPROD", JJ&)
      Call CIERRARCH("SECPROD")
    End If
    'OT-05-0728-WAR-FIN///
    
    SQLX$ = "UPDATE secprod_sql SET MARBORRA = 1 "
    FLEXCONN.Execute (SQLX$)

    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        EcoSecprod$ = ECOARCH$("SECPROD", MKS$(COSECC))
        SQLX$ = "INSERT INTO secprod_sql (NumeOper,DescOper,CODIEMPR,MARBORRA) "
        SQLX$ = SQLX$ + " VALUES('" & TRIM$(COSECC) & "'"
        SQLX$ = SQLX$ + " ,'" & EcoSecprod$ & "'"
        SQLX$ = SQLX$ + " ," & CodiEmp%
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
    Next U&
    Call CIERRARCH("SECPROD")
    SQLX$ = "DELETE FROM secprod_sql WHERE MARBORRA = 1"
    FLEXCONN.Execute (SQLX$)

    
ERROR_GUARDAR:
    If Err <> 0 Then
       Screen.MousePointer = 1
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Call CIERRARCH("SECPROD")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call CIERRARCH("SECPROD")
   End If
End Sub

Private Sub Command11_Click()

    Call TRALOCAL("CODSEPRO", "")
    
    VTB% = VENTABU%("CODSEPRO")
    If VTB% >= 0 Then
        Call TRACENTRAL("CODSEPRO", "")
    End If

End Sub



Private Sub Command12_Click()

    NIVELMAXIMO% = XVALO(Control("ORDEFABR", "NIVMAXEX"))
    
    Respuesta$ = InputBox("Último Nivel Lineal", "Modalidad Lineal + Linealizada", CStr(NIVELMAXIMO%))
    
    Call GRACONTROL("ORDEFABR", "NIVMAXEX", TRIM$(Respuesta$))

End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   EMP_GES05.Show 1
   Command13.Enabled = True

End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CONETIEQEYM.Show 1
   
99 Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command16_Click()
 Command16.Enabled = False
 Call CONECRUN("FLFORMS/" + TRIM$(Text13.TEXT), "Formularios de Impresion")
 Command16.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   If CLACONTRA% > 0 Then
      Command17.Visible = False
   End If
   Command17.Enabled = True
   
End Sub


Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMORFA", Text1.TEXT)
    Call GRACONTROL("", "FORMORFB", Text3.TEXT)
    Call GRACONTROL("", "CODMORFA", Text5.TEXT)
    '
    MONUME$ = "NO" ' PRESENTA VENTANA AHP VALORIZADA
    If Check20.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ORDEFABR", "IMETIQOF", MONUME$)
    '
    If Check19.Value = 1 Then MONUME$ = "1"
    Call GRACONTROL("ORDEFABR", "ETIQTRAZ", MONUME$)

    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEORFA", MONUME$)
    '
    MONUME$ = "LINEAL": FORMUCON% = 1
    If Option4.Value = True Then
       MONUME$ = "MULTILIN"
       FORMUCON% = 2
    End If
    Call GRACONTROL("", "FORMUCON", MONUME$)
    '
    Call GRACONTROL("ORDEFABR", "ETIQUEOF", Text13.TEXT)
    '
    MONUME$ = SAFETEXT$(XVALO(Text15))
    Call GRACONTROL("ORDEFABR", "CANTCOLU", MONUME$)
    '
    If Option12.Value = True Then
       If Check1.Value = False Then
          Call MENSERR(24, "Opción Lineal + Linealizada\Requiere Generación por Subconjuntos.")
          Exit Sub
       End If
       MONUME$ = "LINLINEA"
    End If
    Call GRACONTROL("", "FORMUCON", MONUME$)
    '
    MONUME$ = "NO": If Option14.Value = True Then MONUME$ = "SI"
    Call GRACONTROL("ORDEFABR", "PORUNIDA", MONUME$)

    MONUME$ = "NO": If Option13.Value = True Then MONUME$ = "SI"
    Call GRACONTROL("ORDEFABR", "PORBULTO", MONUME$)

    MONUME$ = "NO"
    If Check1.Value = 1 Then
       MONUME$ = "SI"
       If FORMUCON% = 2 Then
          Call MENSERR(24, "Opciones no Compatibles")
          Check1.Value = 0
          Exit Sub
       End If
    End If
    Call GRACONTROL("", "GESUBOR", MONUME$)
    '
    MONUME$ = "NO": If Check2.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("", "OBSERVOF", MONUME$)
    '
    MONUME$ = "NO": CHECOMPO% = (-1)
    If Check3.Value = 1 Then
        MONUME$ = "SI"
        CHECOMPO% = 1
    End If
    Call GRACONTROL("", "CHECOMPO", MONUME$)
    '
    MONUME$ = "NO": PRECOLAN% = (-1)
    If Check4.Value = 1 Then
        MONUME$ = "SI"
        PRECOLAN% = 1
    End If
    Call GRACONTROL("", "PRECOLAN", MONUME$)
    '
    MONUME$ = "NO": PRECOCIE% = (-1)
    If Check5.Value = 1 Then
        MONUME$ = "SI"
        PRECOCIE% = 1
    End If
    Call GRACONTROL("", "PRECOCIE", MONUME$)
    '
    MONUME$ = "NO": FRALOMAX% = (-1)
    If Check6.Value = 1 Then
        MONUME$ = "SI"
        FRALOMAX% = 1
    End If
    Call GRACONTROL("", "FRALOMAX", MONUME$)
    '
    MONUME$ = "NO": VALINDOP% = (-1)
    If Check7.Value = 1 Then
        MONUME$ = "SI"
        VALINDOP% = 1
    End If
    Call GRACONTROL("", "VALINDOP", MONUME$)
    '
    MONUME$ = "NO": VERITAND% = (-1)
    If Check8.Value = 1 Then
        MONUME$ = "SI"
        VERITAND% = 1
    End If
    Call GRACONTROL("", "VERITAND", MONUME$)
    '
    MONUME$ = "NO": EXISEPRO% = (-1)
    If Check13.Value = 1 Then
        MONUME$ = "SI"
        EXISEPRO% = 1
    End If
    Call GRACONTROL("", "EXISEPRO", MONUME$)
    '
    MONUME$ = "NO"
    If Check9.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIPLANO", MONUME$)
    '
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIESPEC", MONUME$)
    '
    MONUME$ = "NO"
    If Check11.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIHOPRO", MONUME$)
    '
    MONUME$ = "NO"
    If Check12.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIHOINS", MONUME$)
    '
    MONUME$ = "NO"
    If Check14.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL$("ORDEFABR", "SUPRINT", MONUME$)
    '
    MONUME$ = "SI"
    If Check15.Value = 1 Then
        MONUME$ = "NO"
    End If
    Call GRACONTROL("ORDEFABR", "CONTRORE", MONUME$)
    '
    MONUME$ = "NO"
    If Check16.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "MULTIMP", MONUME$)
    '
    MONUME$ = "NO"
    If Check17.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "CIERCALI", MONUME$)
    '
    MONUME$ = "VALEMAT": If Option6.Value = True Then MONUME$ = "INCLUPRI"
    Call GRACONTROL$("ORFABRI", "MPRIUNIQ", MONUME$)
    '
    ' \\\OT-5-739-FG-02/12/05///
    SCRAP1$ = "MAPRIMA"
    If Option7.Value = True Then
      SCRAP1$ = "PRODUCTO"
    End If
    Call GRACONTROL$("ORDEFABR", "MODOSCRA", SCRAP1$)
    ' \\\OT-5-739-FG-FIN///
    
    ORDEN$ = "ESTRUC":
    If Option10.Value = True Then ORDEN$ = "UBICA"
    If Option11.Value = True Then ORDEN$ = "TRA+UBI"
    Call GRACONTROL("ORDEFABR", "ORDENA", ORDEN$)
    
    MONUME$ = ""
    If Check18.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("ORDEFABR", "IMPOFMRP", MONUME$)
    '
    
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   PRINTFORM
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPVALEMAT.Show 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

    Text1.TEXT = Control$("", "FORMORFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FORMORFB")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    If Control("ORDEFABR", "IMETIQOF") = "SI" Then
       Check20.Value = 1
       Command17.Visible = False
    Else
       Command16.Visible = True
    End If
    
    Check19.Value = XVALO(Control("ORDEFABR", "ETIQTRAZ"))
    '
    Text15 = XVALO(Control("ORDEFABR", "CANTCOLU"))
    '
    Text13 = Control("ORDEFABR", "ETIQUEOF")
    Text14.TEXT = ECOARCH$("INDIFRM", Text13.TEXT)
    '
    MONUME$ = Control("ORDEFABR", "PORUNIDA")
    If MONUME$ = "SI" Then Option14.Value = True
    '
    MONUME$ = Control("ORDEFABR", "PORBULTO")
    If MONUME$ = "SI" Then Option13.Value = True
    '
    
    If Control$("", "NUMEORFA") = "MANUAL" Then Option1.Value = True
    If Control$("", "GESUBOR") = "SI" Then Check1.Value = 1: If Check1.Value = 1 Then Command11.Enabled = True
    If Control$("", "OBSERVOF") = "SI" Then Check2.Value = 1
    If Control$("", "CHECOMPO") = "SI" Then Check3.Value = 1
    If Control$("", "PRECOLAN") = "SI" Then Check4.Value = 1
    If Control$("", "PRECOCIE") = "SI" Then Check5.Value = 1
    If Control$("", "FRALOMAX") = "SI" Then Check6.Value = 1
    If Control$("", "VALINDOP") = "SI" Then Check7.Value = 1
    If Control$("", "VERITAND") = "SI" Then Check8.Value = 1
    If Control$("", "EXISEPRO") = "SI" Then Check13.Value = 1
    If Control("ORDEFABR", "CONTRORE") = "NO" Then Check15.Value = 1
    If Control$("ORDEFABR", "PRIPLANO") = "SI" Then Check9.Value = 1
    If Control$("ORDEFABR", "PRIESPEC") = "SI" Then Check10.Value = 1
    If Control$("ORDEFABR", "PRIHOPRO") = "SI" Then Check11.Value = 1
    If Control$("ORDEFABR", "PRIHOINS") = "SI" Then Check12.Value = 1
    If Control$("", "FORMUCON") = "MULTILIN" Then Option4.Value = True
    If Control$("", "FORMUCON") = "LINLINEA" Then Option12.Value = True
    If Control$("ORDEFABR", "SUPRINT") = "SI" Then Check14.Value = 1
    If Control$("ORFABRI", "MPRIUNIQ") = "INCLUPRI" Then Option6.Value = 1
    ' \\\OT-5-739-FG-02/12/05///
    If Control$("ORDEFABR", "MODOSCRA") = "PRODUCTO" Then Option7.Value = True
    ' \\\OT-5-739-FG-FIN///
    '
    Option9.Value = True ' CONTROL DE ORDENAMIENTO DE IMPRESION DE ORDEN DE FABRICACION
    If Control$("ORDEFABR", "ORDENA") = "UBICA" Then Option10.Value = True
    If Control$("ORDEFABR", "ORDENA") = "TRA+UBI" Then Option11.Value = True
    '
    CMO$ = TRIM$(Control$("", "CODMORFA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMANN") > 0 Then
       Command9.Visible = True
    End If
    '
    If Control("ORDEFABR", "MULTIMP") = "SI" Then Check16.Value = 1
    '
    
    Check18.Value = 0
    If Control("ORDEFABR", "IMPOFMRP") = "SI" Then Check18.Value = 1

End Sub

Private Sub Option12_Click()
    
    Command12.Enabled = True
    
End Sub

Private Sub Option3_Click()
    
    Command12.Enabled = False
    Check16.Value = False
    
End Sub

Private Sub Option4_Click()

    Command12.Enabled = False
    
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
Private Sub Text13_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text14.TEXT = VALACT2$("INDIFRM")
    End If

End Sub


Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SETUPCOM 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras y Expedición - Configuración y Formularios"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11640
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
   ScaleHeight     =   6390
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   8520
      Left            =   10635
      ScaleHeight     =   8460
      ScaleWidth      =   1965
      TabIndex        =   55
      Top             =   0
      Width           =   2025
      Begin VB.PictureBox Picture16 
         Height          =   6345
         Left            =   0
         ScaleHeight     =   6285
         ScaleWidth      =   1020
         TabIndex        =   98
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   7185
            Left            =   0
            Picture         =   "SETUPCOM.frx":0000
            Stretch         =   -1  'True
            Top             =   -960
            Width           =   1095
         End
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
         Height          =   640
         Left            =   175
         Picture         =   "SETUPCOM.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1380
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
         Height          =   640
         Left            =   175
         Picture         =   "SETUPCOM.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   750
         Width           =   650
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   3375
         Left            =   70
         TabIndex        =   61
         Top             =   2040
         Width           =   825
         Begin VB.CommandButton Command17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Picture         =   "SETUPCOM.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Texto de Mail de Compras"
            Top             =   240
            Width           =   600
         End
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   125
            Picture         =   "SETUPCOM.frx":2CB6
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Clasificación de Proveedores por Moneda de Compra y Valuación"
            Top             =   2290
            Width           =   595
         End
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   125
            Picture         =   "SETUPCOM.frx":2FC0
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Reglamento de Compras"
            Top             =   2800
            Width           =   595
         End
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   125
            Picture         =   "SETUPCOM.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Tabla de Tipo de Cambio"
            Top             =   1260
            Width           =   595
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
            Height          =   500
            Left            =   125
            Picture         =   "SETUPCOM.frx":3414
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Tasas de I.V.A."
            Top             =   740
            Width           =   595
         End
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   125
            Picture         =   "SETUPCOM.frx":355E
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Tabla de Embalajes Codificados"
            Top             =   1770
            Width           =   595
         End
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
         Height          =   640
         Left            =   175
         Picture         =   "SETUPCOM.frx":3868
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   56
         Top             =   5520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   57
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "SETUPCOM.frx":39B2
         Top             =   5700
         Width           =   1515
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lugar de Entrega Habitual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5500
      TabIndex        =   52
      Top             =   105
      Width           =   4980
      Begin VB.CheckBox Check15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Presentar Contacto de O.Compra"
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
         Left            =   1135
         TabIndex        =   85
         Top             =   735
         Width           =   3495
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "SETUPCOM.frx":58D4
         Top             =   240
         Width           =   3720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leyenda:"
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
         Left            =   -1260
         TabIndex        =   54
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00C0FFFF&
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
      Height          =   750
      Left            =   3840
      TabIndex        =   49
      Top             =   4485
      Width           =   1335
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Directa"
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
         TabIndex        =   51
         Top             =   210
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Posterior"
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
         Top             =   450
         Width           =   1125
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reglamento de Compras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1070
      Left            =   5500
      TabIndex        =   48
      Top             =   1140
      Width           =   4980
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
         Height          =   730
         Left            =   3975
         Picture         =   "SETUPCOM.frx":58D6
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   660
         Width           =   3645
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2295
         TabIndex        =   68
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
         Left            =   3660
         TabIndex        =   67
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   -120
         TabIndex        =   71
         Top             =   300
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   5160
      Begin VB.Frame Frame23 
         BackColor       =   &H00C0FFFF&
         Height          =   400
         Left            =   240
         TabIndex        =   105
         Top             =   5800
         Width           =   4700
         Begin VB.CheckBox Check20 
            BackColor       =   &H00C0FFFF&
            Caption         =   "No Considerar Desc. al Actulizar Costo de Repos."
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
            TabIndex        =   106
            Top             =   160
            Width           =   4500
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H00C0FFFF&
         Height          =   400
         Left            =   210
         TabIndex        =   103
         Top             =   5355
         Width           =   1785
         Begin VB.OptionButton Option14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Imprime MRP"
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
            Left            =   210
            TabIndex        =   104
            ToolTipText     =   "Imprime en la orden de Compra la necesidad segun MRP segun Cálculo Inmediato Anterior"
            Top             =   160
            Width           =   1440
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Depurar O.Compras Vencidas"
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
         Left            =   2150
         TabIndex        =   86
         Top             =   5130
         Width           =   2805
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   89
            Text            =   " "
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tolerancia Entrega:"
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
            TabIndex        =   87
            Top             =   255
            Width           =   1695
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Impres."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2100
         TabIndex        =   76
         Top             =   1080
         Width           =   1080
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aprob."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2140
         TabIndex        =   72
         Top             =   4380
         Width           =   1335
         Begin VB.OptionButton Option13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Destino"
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
            TabIndex        =   74
            Top             =   450
            Width           =   1125
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Unidades"
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
            TabIndex        =   73
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documentos Anexos  (Requiere Sist.Calidad)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   210
         TabIndex        =   28
         Top             =   3675
         Width           =   4745
         Begin VB.CheckBox Check8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "L.Materiales"
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
            Left            =   3300
            TabIndex        =   81
            Top             =   300
            Width           =   1365
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "H.R.I.'s"
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
            Left            =   2330
            TabIndex        =   31
            Top             =   310
            Width           =   1065
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especific."
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
            Left            =   1080
            TabIndex        =   30
            Top             =   300
            Width           =   1905
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Planos"
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
            TabIndex        =   29
            Top             =   300
            Width           =   945
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualiza C.Rep."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   210
         TabIndex        =   44
         Top             =   4380
         Width           =   1785
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nunca"
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
            Left            =   210
            TabIndex        =   47
            Top             =   700
            Width           =   1440
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Prov.Habit"
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
            Left            =   210
            TabIndex        =   46
            Top             =   495
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Siempre"
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
            Left            =   210
            TabIndex        =   45
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Precio Present. \ Val. Precio"
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
         Left            =   2100
         TabIndex        =   40
         Top             =   2925
         Width           =   2865
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
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
            Left            =   2150
            TabIndex        =   79
            Text            =   " "
            Top             =   300
            Width           =   500
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Costo Repos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   430
            Left            =   1250
            TabIndex        =   41
            Top             =   220
            Width           =   820
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ult.Comp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   430
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Última Compra"
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   2640
            TabIndex        =   80
            Top             =   330
            Width           =   165
         End
      End
      Begin VB.Frame Frame11 
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
         Height          =   1300
         Left            =   210
         TabIndex        =   37
         Top             =   2325
         Width           =   1800
         Begin VB.CheckBox Check7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Val. Cod-Prov"
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
            Left            =   210
            TabIndex        =   78
            Top             =   1040
            Width           =   1500
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Repite Propio"
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
            Left            =   210
            TabIndex        =   43
            Top             =   720
            Width           =   1485
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "App.Revision"
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
            Left            =   210
            TabIndex        =   39
            Top             =   480
            Width           =   1485
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nro. Original"
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
            Left            =   210
            TabIndex        =   38
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1800
         Begin VB.CheckBox Check14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Des.Complem."
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
            Left            =   120
            TabIndex        =   84
            Top             =   450
            Width           =   1605
         End
         Begin VB.CheckBox Check13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Desc. Orig."
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
            Left            =   120
            TabIndex        =   83
            Top             =   200
            Width           =   1605
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especificación"
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
            Left            =   120
            TabIndex        =   36
            Top             =   700
            Width           =   1605
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Moneda Costeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   2100
         TabIndex        =   22
         Top             =   1845
         Width           =   2850
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Restauracion Automatica"
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
            TabIndex        =   75
            Top             =   640
            Width           =   2565
         End
         Begin VB.CommandButton Command10 
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
            Left            =   525
            TabIndex        =   33
            Top             =   285
            Width           =   175
         End
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
            Left            =   120
            TabIndex        =   32
            Text            =   " "
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S/Tabla Proveedores"
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
            Left            =   840
            TabIndex        =   34
            Top             =   285
            Width           =   1905
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
         Height          =   315
         Left            =   2415
         TabIndex        =   21
         Text            =   " "
         Top             =   330
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
         Height          =   360
         Left            =   3780
         TabIndex        =   20
         Top             =   330
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
         Height          =   300
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
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
         Height          =   730
         Left            =   4095
         Picture         =   "SETUPCOM.frx":5BE0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   285
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3240
         TabIndex        =   14
         Top             =   1080
         Width           =   885
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Auto"
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
            Left            =   100
            TabIndex        =   15
            Top             =   480
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Man"
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
            Left            =   100
            TabIndex        =   16
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cód."
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
         Left            =   4150
         TabIndex        =   12
         ToolTipText     =   "Código"
         Top             =   1080
         Width           =   800
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   190
            TabIndex        =   13
            Text            =   "OC"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   17
         Top             =   375
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remito en Consignación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4070
      Left            =   5500
      TabIndex        =   0
      Top             =   2325
      Width           =   4980
      Begin VB.CommandButton Command19 
         Caption         =   "Configuración   -   Valida Stock/Pedidos P/O.Compra     "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   170
         Picture         =   "SETUPCOM.frx":5EEA
         TabIndex        =   102
         Top             =   2085
         Width           =   4650
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generar Arch. Transferencia de O.C. a Central"
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
         Left            =   200
         TabIndex        =   101
         ToolTipText     =   "Generar Archivo de Transferencia de O.Compra a Central"
         Top             =   2310
         Width           =   4590
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valida Stock Mínimo"
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
         Left            =   2450
         TabIndex        =   100
         ToolTipText     =   "Si el Stock esta es Superior al Stock Mínimo No Permitie Compar/ Caso Contraro Valida Contra el Stock Maximo"
         Top             =   2070
         Width           =   2430
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valida Stock / Pedidos "
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
         Left            =   200
         TabIndex        =   99
         ToolTipText     =   "Valida Stocks y Pedidos Pendientes al momento de Comprar Evitando la Sobre - Compra "
         Top             =   2070
         Width           =   2190
      End
      Begin VB.Frame Frame21 
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
         Height          =   650
         Left            =   165
         TabIndex        =   93
         Top             =   3285
         Width           =   4635
         Begin VB.CommandButton Command18 
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
            Height          =   330
            Left            =   2265
            TabIndex        =   95
            Top             =   240
            Width           =   200
         End
         Begin VB.Label Label11 
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
            Height          =   330
            Left            =   2460
            TabIndex        =   97
            ToolTipText     =   "Selección de Moneda a Utilizar en Texto en la Orden de Compra"
            Top             =   240
            Width           =   2040
         End
         Begin VB.Label Label10 
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
            Height          =   330
            Left            =   1665
            TabIndex        =   96
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda x Defecto:"
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
            Left            =   -120
            TabIndex        =   94
            Top             =   315
            Width           =   1665
         End
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Utiliza Formulario de Remito de Ventas"
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
         Left            =   200
         TabIndex        =   92
         Top             =   1820
         Width           =   4365
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agrupa Consignaciones"
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
         Left            =   200
         TabIndex        =   82
         Top             =   1605
         Width           =   2325
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   165
         TabIndex        =   23
         Top             =   2530
         Width           =   4635
         Begin VB.TextBox Text8 
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
            Left            =   2205
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   " "
            Top             =   215
            Width           =   2280
         End
         Begin VB.TextBox Text7 
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
            Left            =   1155
            TabIndex        =   25
            Text            =   " "
            Top             =   215
            Width           =   750
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
            Height          =   300
            Left            =   1890
            TabIndex        =   24
            Top             =   215
            Width           =   175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito:"
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
            Left            =   -1260
            TabIndex        =   27
            Top             =   320
            Width           =   2325
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cód."
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
         Left            =   3990
         TabIndex        =   9
         ToolTipText     =   "Código"
         Top             =   915
         Width           =   855
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   190
            TabIndex        =   10
            Text            =   "RC"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   210
         TabIndex        =   6
         Top             =   915
         Width           =   3645
         Begin VB.OptionButton Option4 
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
            Left            =   2100
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
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
            Left            =   525
            TabIndex        =   7
            Top             =   300
            Width           =   1380
         End
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
         Left            =   3675
         TabIndex        =   4
         Top             =   180
         Width           =   175
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
         Height          =   300
         Left            =   2310
         TabIndex        =   3
         Text            =   " "
         Top             =   180
         Width           =   1380
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
         Height          =   300
         Left            =   210
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3645
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
         Height          =   730
         Left            =   3990
         Picture         =   "SETUPCOM.frx":61F4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   -105
         TabIndex        =   5
         Top             =   240
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SETUPCOM.frx":64FE
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "SETUPCOM.frx":6578
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
      Begin VB.Menu tasasdeiva 
         Caption         =   "Tasas de I.V.A."
      End
      Begin VB.Menu tabladetipodecambio 
         Caption         =   "Tabla de Tipo de Cambio"
      End
      Begin VB.Menu embalajes 
         Caption         =   "Tabla de Embalajes Codificados"
      End
      Begin VB.Menu monedayproveedor 
         Caption         =   "Clasificación de Proveedores por Moneda de Compra y Valuación"
      End
      Begin VB.Menu ReglamentodeCompras 
         Caption         =   "Reglamento de Compras "
      End
      Begin VB.Menu mnuconfigmail 
         Caption         =   "Configuración de Cta. de Mail"
      End
      Begin VB.Menu ConfigMineparts 
         Caption         =   "Configuración Mineparts"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "SETUPCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("TAMONED")
    VDEVU$ = VALACT1$("TAMONED")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Label5 = VALACT2$("TAMONED")
    End If
End Sub

Sub Command11_Click()
   Command11.Enabled = False
   Call CARTAMONED
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command11.Enabled = False
   Call CARTAEMBAL
   Command11.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   '
   Screen.MousePointer = 11
   '
   UREVI& = Val(Control$("OCOMPRA", "RMONEPRO")) - 128
   If UREVI& < 1 Then UREVI& = 1
   '
   FIN& = ULTREG&("OCOMENCA")
   AGREPRO% = 0
   PUNPROAC$ = ""
   For U& = UREVI& To FIN&
     XX$ = REGLEIDO$("OCOMENCA", U&)
     NPRO% = CVI(Mid$(XX$, 51, 2))
     If NPRO% > 0 Then
       RFF$ = RECFILT$("CLAPROMO", 1, MKI$(NPRO%))
       If Len(RFF$) < 4 Then
         For KKI% = 1 To Len(PUNPROAC$) Step 2
           If CVI(Mid$(PUNPROAC$, KKI%, 2)) = NPRO% Then GoTo 19
         Next KKI%
         YY$ = REGBLAN$("CLAPROMO")
         Call REPLA(YY$, MKI$(NPRO%), 1, 2)
         Call REPLA(YY$, rasocli$((-1) * NPRO%), 3, 28)
         Call REPLA(YY$, Chr$(1), 31, 1)
         Call REPLA(YY$, Chr$(1), 32, 1)
         Call REGAPP("CLAPROMO", YY$)
         Call CIERRARCH("CLAPROMO")
         PUNPROAC$ = PUNPROAC$ + MKI$(NPRO%)
         AGREPRO% = 1
       End If
     End If
19 Next U&
   '
   If AGREPRO% > 0 Then
     Call FILESORT("CLAPROMO", "", 2, 2, "ASC")
   End If
   '
   Screen.MousePointer = 1
   Call TRALOCAL("CLAPROMO", "")
   VTB% = VENTABU%("CLAPROMO")
   If VTB% >= 0 Then
     Call TRACENTRAL("CLAPROMO", "")
     Call GRACONTROL("OCOMPRA", "RMONEPRO", Str$(ULTREG&("OCOMENCA")))
   End If
   '
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
  '
  RTA% = ALTERNA%("Tabla de Reglamentos\Editar Reglamento Existente")
  If RTA% = 1 Then
      Call CARTAREGLA
    ElseIf RTA% = 2 Then
      Call SELECHO("INDIRCOM")
      VDEVU$ = VALACT1$("INDIRCOM")
      If TRIM(VDEVU$) <> "" Then
        With REGLACOM04
          .Label37 = VDEVU$
          .Show 1
        End With
      End If
  End If
  '
End Sub

Private Sub Command15_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text12TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text11.TEXT), "Formularios de Impresion")
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
      '
      Call COPYSTRU("OBSERVOF", "TXMAICOM")
      Call TRALOCAL("TXMAICOM", "OBSERVOF")
      VTB% = VENTABU%("OBSERVOF/TITUPAN=Texto de Mail para Orden de Compra")
      If VTB% < 0 Then Exit Sub
      '
      For U& = 1 To ULTREG&("!OBSERVOF")
           TX$ = REGLEIDO$("!OBSERVOF", U&)
           Call REPLA(X$, TX$, 1, 64)
           Call GRAREG("!TXMAICOM", X$, U&)
      Next U&
      '
      Call SETULTREG("!TXMAICOM", U& - 1)
      Call CIERRARCH("!TXMAICOM")
      Call TRACENTRAL("TXMAICOM", "")
      Call CIERRARCH("OBSERVOF")
      '
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call SELECHO("TAMONED")
   If XVALO(VALACT1$("TAMONED")) > 0 Then
        MONEMX% = XVALO(VALACT1$("TAMONED"))
        Label10 = MONEMX%
   End If
   '
   Command18.Enabled = True

End Sub

Private Sub Command19_Click()
   If CLACONTRA% < 1 Then
     Exit Sub
   End If
   '
   Command19.Visible = False
   
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    If XVALO(Text14) > 365 Then
       Call COMUNI("Tolerancia de Entrega Supera Máximo de '365'")
       On Error Resume Next
       Text14.SelStart = 0
       Text14.SelLength = Len(Text14.TEXT)
       Text14.SetFocus
       On Error GoTo 0
       GoTo 99
    End If
    '
    Call GRACONTROL("", "FORMOCOM", Text1.TEXT)
    Call GRACONTROL("", "FOREMICO", Text3.TEXT)
    Call GRACONTROL("", "CODMOCOM", Text5.TEXT)
    Call GRACONTROL("", "CODRECON", Text6.TEXT)
    Call GRACONTROL("", "DEPOCONS", Text7.TEXT)
    Call GRACONTROL("OCOMPRA", "MONECOS", Text9.TEXT)
    '\\\OT-05-0628-WAR-19/09/05///
    Call GRACONTROL("OCOMPRA", "LENTREGA", Text10.TEXT)
    '\\\OT-05-0628-WAR-FIN///
    '\\\OT-05-0676-WAR-14/10/05///
    Call GRACONTROL("OCOMPRA", "REGLACOM", Text11.TEXT)
    '\\\OT-05-0676-WAR-FIN///
    Call GRACONTROL("", "SOBREPRE", Text13.TEXT)

    On Error Resume Next
      FORCOMPRA.Text15 = Text9
    On Error GoTo 0
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEOCOM", MONUME$)
    MONUME$ = "AUTO": If Option3.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMERECO", MONUME$)
    '
    MONUME$ = "NO"
    If Check9.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIPLANO", MONUME$)
    '
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIESPEC", MONUME$)
    '
    MONUME$ = "NO"
    If Check11.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "PRIHRINS", MONUME$)
    '
    'GRABACION DE OPCION DE LISTADO DE MATERIALES
    MONUME$ = "NO"
    If Check8.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "LISMATE", MONUME$)
    '
    'GRABACION DE AGRUPA CONSIGNACIONES
    MONUME$ = "NO"
    If Check12.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "AGRUCONS", MONUME$)
    
    'GRABACION DE USA FORMULARIO DE REMITO DE VENTAS
    MONUME$ = "NO"
    If Check16.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "UTIRTVTA", MONUME$)
    '
    MONUME$ = "NO"
    If Check1.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "INCLUSPE", MONUME$)
    '
    MONUME$ = "NO"
    If Check13.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "IMPDESPR", MONUME$)
    '
    MONUME$ = "NO"
    If Check14.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "IMPDESCO", MONUME$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "REPLAORI", MONUME$)
    '
    MONUME$ = "NO"
    If Check3.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "REVICOD", MONUME$)
    '
    MONUME$ = "NO"
    If Check4.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "RECOPROP", MONUME$)
    '
    MONUME$ = "ULTICOM"
    If Option5.Value = True Then
        MONUME$ = "COSTOREP"
    End If
    Call GRACONTROL("OCOMPRA", "PRECIDEF", MONUME$)
    '
    MONUME$ = "DIRECTO"
    If Option10.Value = True Then
        MONUME$ = "POSTERIOR"
    End If
    Call GRACONTROL("OCOMPRA", "MODOAPRO", MONUME$)
    '
    MONUME$ = "PROHABIT"
    If Option7.Value = True Then MONUME$ = "SIEMPRE"
    If Option9.Value = True Then MONUME$ = "NUNCA"
    Call GRACONTROL("OCOMPRA", "ACOSREP", MONUME$)
    '
    '\\\OT-05-0688-WAR-20/10/05///
    MONUME$ = "UNI" 'UNIDADES
    If Option13.Value = True Then
        MONUME$ = "DES" 'DESTINO
    End If
    Call GRACONTROL("OCOMPRA", "MODOIMPU", MONUME$)
    '\\\OT-05-0688-WAR-FIN///
    '\\\OT-08-00884-OD-26/09/08///
    MONUME$ = "MANUAL"
    If Option2.Value = True Then
     MONUME$ = "AUTOMA"
    End If
    Call GRACONTROL("OCOMPRA", "NUMERA", MONUME$)
    '\\\OT-08-00884-OD-FIN///
    '
    '\\\OT-08-00980-FL-28/10/08///
    MONUME$ = "NO"
    If Check5.Value = 1 Then
     MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "RESAUTO", MONUME$)
    '\\\OT-08-00980-FL-FIN///
    '
    '\\\OT-08-01047-FL-21/11/08///Imprime O/C luego de la aprobación
    If Check6.Value = 1 Then
     Call GRACONTROL("OCOMPRA", "IMPRIME", "ALAPROBA")
    End If
    '\\\OT-08-001047-FL-FIN///
    '
    'Valida Codigo-Proveedor
    MONUME$ = "NO"
    If Check7.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("OCOMPRA", "VALCOPRO", MONUME$)
    '
    MONUME$ = ""
    If Check15.Value = 1 Then MONUME$ = "NOPRESCON"
    Call GRACONTROL("OCOMPRA", "PRESENCO", MONUME$)
    '
    Call GRACONTROL("ORCOMPRA", "DIOCVEN", FORMATNUM$(XVALO(Text14), "F4.0"))
    '
    MONUME$ = "1"
    If XVALO(Label10) > 1 Then MONUME$ = TRIM$(Label10)
    Call GRACONTROL("ORCOMPRA", "MOCODEFE", MONUME$)

    If Check1.Value = 0 And Check13.Value = 0 And Check14.Value = 0 Then
        Call COMUNI("Debe Elegir Uno de los Valores Para Descripcion")
        GoTo 99
    End If
    '
    'SI ESTA MARCADO VALIDA QUE LAS COMPRAS NO SUPEREN EL MAXIMO PERMITDO
    'EL MAXIMO PERMITIDO SE OBTIENE QUE NO SUPERE EL STOCK MAXIMO O BIEN QUE DE POSITIVO (PEDIDOS PENDIENTES DE ENTREGA - ORDENES DE COMPRA PENDIENTES)
    MONUME$ = ""
    If Check17.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("OCOMPRA", "VALISTPE", MONUME$)
    '
    'SI ESTA MARCADO EL CHECK18 SOLO SE PODRA PEDIR SI EL STOCK ESTA POR DEBAJO DEL MINIMO, O BIEN QUE DE POSITIVO (PEDIDOS PENDIENTES DE ENTREGA - ORDENES DE COMPRA PENDIENTES)
    MONUME$ = ""
    If Check18.Value = 1 And Check17.Value = 0 Then
      OPX% = COMALTER%("Atención !!!! \No se Grabará la Opción Valida Stock Mínimo\Debido a Que Es Derivada de la Validación 'Valida Stock / Pedidos'\\Cancelar\Continuar")
      If OPX% < 2 Then GoTo 99
    Else
      If Check18.Value = 1 Then MONUME$ = "SI"
      Call GRACONTROL("OCOMPRA", "CONSSTMI", MONUME$)
    End If
    '
    MONUME$ = ""
    If Check19.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("OCOMPRA", "ENVARSUC", MONUME$)
    
    '
    MONUME$ = ""
    If Option14.Value = True Then MONUME$ = "SI"
    Call GRACONTROL("OCOMPRA", "IMPRCMRP", MONUME$)
    
    '
    MONUME$ = ""
    If Check20.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("OCOMPRA", "NODESCUE", MONUME$)
    
    Call CIERRARCH("*.*")
    Call Command1_Click
    '
99  Command2.Enabled = True
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
    Command5.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command3.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
    Command3.Enabled = True
End Sub

Private Sub Command8_Click()
    If DERACCE%("TASAIVA", "tasas de IVA") = 2 Then
      Call CARDATCENCALC
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub ConfigMineparts_Click()
'   If CLACONTRA% < 1 Then Exit Sub
   CfgMineparts.Show 1
    
End Sub

Private Sub embalajes_Click()
   Call Command12_Click
End Sub

Private Sub Form_Load()
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      Me.Frame12.BackColor = &HFCD7B6
      Me.Frame13.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      Me.Frame15.BackColor = &HFCD7B6
      Me.Frame16.BackColor = &HFCD7B6
      Me.Frame17.BackColor = &HFCD7B6
      Me.Frame18.BackColor = &HFCD7B6
      Me.Frame19.BackColor = &HFCD7B6
      Me.Frame20.BackColor = &HFCD7B6
      Me.Frame21.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6
      Me.Option6.BackColor = &HFCD7B6
      Me.Option7.BackColor = &HFCD7B6
      Me.Option8.BackColor = &HFCD7B6
      Me.Option9.BackColor = &HFCD7B6
      Me.Option10.BackColor = &HFCD7B6
      Me.Option11.BackColor = &HFCD7B6
      Me.Option12.BackColor = &HFCD7B6
      Me.Option13.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      Me.Check3.BackColor = &HFCD7B6
      Me.Check4.BackColor = &HFCD7B6
      Me.Check5.BackColor = &HFCD7B6
      Me.Check6.BackColor = &HFCD7B6
      Me.Check7.BackColor = &HFCD7B6
      Me.Check8.BackColor = &HFCD7B6
      Me.Check9.BackColor = &HFCD7B6
      Me.Check10.BackColor = &HFCD7B6
      Me.Check11.BackColor = &HFCD7B6
      Me.Check12.BackColor = &HFCD7B6
      Me.Check13.BackColor = &HFCD7B6
      Me.Check14.BackColor = &HFCD7B6
      Me.Check15.BackColor = &HFCD7B6
      Me.Check16.BackColor = &HFCD7B6
      '
      
      MARCOBARRA.Top = 5700
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") Then
      ConfigMineparts.Visible = True
    End If
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
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
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    Text7.TEXT = Control$("", "DEPOCONS")
    If Val(Text7) < 0 Or Val(Text7) > 255 Then
        Call MENSERR(24, "Mal Número de Depósito de Consumo")
      Else
        Text8.TEXT = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    End If
    '
    Text9.TEXT = TRIM$(Left$(Control$("OCOMPRA", "MONECOS"), 3))
    If Val(Text9) >= 0 And Val(Text9) < 256 Then
         Label5 = ECOARCH$("TAMONED", Chr$(Val(Text9.TEXT)))
       Else
         Call MENSERR(24, "Mal Moneda de Costeo")
    End If
    '
    Text1.TEXT = Control$("", "FORMOCOM")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FOREMICO")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '\\\OT-05-0676-WAR-14/10/05///
    Text11.TEXT = Control$("", "REGLACOM")
    Text12.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    '\\\OT-05-0676-WAR-FIN///
    '
    Text13.TEXT = "0"
    If XVALO(Control$("", "SOBREPRE")) > 0 Then
       Text13.TEXT = Control$("", "SOBREPRE") 'PORCENTAJE DE MARGEN DE PRECIO DE COMPRA
    End If
    '

    If Control$("", "NUMEOCOM") = "MANUAL" Then Option1.Value = True
    If Control$("", "NUMERECO") = "MANUAL" Then Option3.Value = True
    '
    CMO$ = TRIM$(Control$("", "CODMOCOM")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(Control$("", "CODRECON")): If CMO$ <> "" Then Text6.TEXT = CMO$
    '
    If Control$("OCOMPRA", "PRIPLANO") = "SI" Then Check9.Value = 1
    If Control$("OCOMPRA", "PRIESPEC") = "SI" Then Check10.Value = 1
    If Control$("OCOMPRA", "PRIHRINS") = "SI" Then Check11.Value = 1
    
    If Control$("OCOMPRA", "IMPDESPR") = "SI" Then Check13.Value = 1
    If Control$("OCOMPRA", "IMPDESCO") = "SI" Then Check14.Value = 1
    
    If Control$("OCOMPRA", "INCLUSPE") = "SI" Then Check1.Value = 1
    If Control$("OCOMPRA", "REPLAORI") = "SI" Then Check2.Value = 1
    If Control$("OCOMPRA", "REVICOD") = "SI" Then Check3.Value = 1
    If Control$("OCOMPRA", "RECOPROP") = "SI" Then Check4.Value = 1
    If Control$("OCOMPRA", "VALCOPRO") = "SI" Then Check7.Value = 1        'Valida Codigo-Proveedor
    If Control("OCOMPRA", "LISMATE") = "SI" Then Check8.Value = 1 ' LISTADO DE MATERIALES DE ARMADORES
    If Control("OCOMPRA", "AGRUCONS") = "SI" Then Check12.Value = 1  'AGRUPA CONSIGNACIONES
    If Control("OCOMPRA", "UTIRTVTA") = "SI" Then Check16.Value = 1  'UTILIZA FORMULARIO DE REMITO DE VENTAS
    '
    If Control("OCOMPRA", "PRECIDEF") = "COSTOREP" Then Option5.Value = True
    If Control("OCOMPRA", "MODOAPRO") = "POSTERIOR" Then Option10.Value = True
    '\\\OT-05-0688-WAR-20/10/05///
    If Control("OCOMPRA", "MODOIMPU") = "DES" Then Option13.Value = True
    '\\\OT-05-0688-WAR-FIN///
    '
    ACOSREP$ = UCase$(Control$("OCOMPRA", "ACOSREP"))
    Option8.Value = True
    If ACOSREP$ = "SIEMPRE" Then Option7.Value = True
    If ACOSREP$ = "NUNCA" Then Option9.Value = True
    '
    '\\\OT-05-0628-WAR-19/09/05///
    Text10.TEXT = Control$("OCOMPRA", "LENTREGA")
    '\\\OT-05-0628-WAR-FIN///
    '
    '\\\OT-08-00884-OD-26/09/08///
    If Control$("ORCOMPRA", "NUMERA") = "AUTOMA" Then Option2.Value = True
    '\\\OT-08-00884-OD-FIN///
    '
    '\\\OT-08-00980-FL-28/10/08///
    If Control$("OCOMPRA", "RESAUTO") = "SI" Then Check5.Value = 1
    '\\\OT-08-00980-FL-FIN///
    '
    '\\\OT-08-01047-FL-21/11/08/// Levanta la configuracion de impresion: Al aprobar (checked = SI)/Siempre(checked = NO)
    If Control$("OCOMPRA", "IMPRIME") = "ALAPROBA" Then Check6.Value = 1
    '\\\OT-08-01047-FL-FIN///
    
    If Control("OCOMPRA", "PRESENCO") = "NOPRESCON" Then Check15.Value = 1
    
    Text14 = FORMATNUM$(XVALO(Control("ORCOMPRA", "DIOCVEN")), "F4.0")
    '
    If XVALO(Control("ORCOMPRA", "MOCODEFE")) < 1 Then ' MONEDA POR DEFECTO
      Label10 = "1"
    Else
      Label10 = TRIM$(XVALO(Control("ORCOMPRA", "MOCODEFE")))
    End If
    '
    'VALIDA CONTROL DE LAS COMPRAS SEGUN STOCK Y PEDIDOS PENDIENTES. ' le agrego uno dos y tres para saber sin que el usuario descubra el boton saber que esta configurado
    Check17.Value = 0
    If Control("OCOMPRA", "VALISTPE") = "SI" Then Check17.Value = 1: Command19.Caption = Command19.Caption + "  1"
    Check18.Value = 0
    If Control("OCOMPRA", "CONSSTMI") = "SI" Then Check18.Value = 1: Command19.Caption = Command19.Caption + "  2"
    Check19.Value = 0
    If Control("OCOMPRA", "ENVARSUC") = "SI" Then Check19.Value = 1: Command19.Caption = Command19.Caption + "  3"    'ENVIAR ARCHIVO DE TRANSFERENCIA DE O.C. A CENTRAL
    Option14.Value = False
    If Control("OCOMPRA", "IMPRCMRP") = "SI" Then Option14.Value = True
    '
    Check20.Value = 0
    If Control("OCOMPRA", "NODESCUE") = "SI" Then Check20.Value = 1
End Sub




Private Sub grabar_Click()
  Call Command2_Click
End Sub

Private Sub Label10_Change()
   Label11 = ""
   If XVALO(Label10) > 0 Then
      Label11 = ECOARCH$("TAMONED", MKI$(XVALO(Label10)))
   End If

End Sub

Private Sub mnuCarpComprasInt_Click()
  CfgContep.Show 1
End Sub

Private Sub mnuconfigmail_Click()
    CONFIMAIL.Show 1
End Sub

Private Sub mnuTablasAuxiliares_Click()
   Call GENERAR_ESTRUCTURAS_SQL
   Call PRESENTA_TABLAS_AUXILIARES

End Sub

Private Sub monedayproveedor_Click()
   Call Command13_Click
End Sub

Private Sub Option15_Click()

End Sub

Private Sub ReglamentodeCompras_Click()
   Call Command14_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub tabladetipodecambio_Click()
   Call Command11_Click
End Sub

Private Sub tasasdeiva_Click()
   Call Command8_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text12.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text13_LostFocus()
   Text13 = FORMATNUM(Text13, "F4.1")
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Text9_Change()
   If Val(Text9) < 1 Then
       Label5 = "S/Tabla Proveedores"
     Else
       Label5 = ECOARCH$("TAMONED", Chr$(Val(Text9)))
   End If
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("TAMONED")
    VDEVU$ = VALACT1$("TAMONED")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
    End If
End Sub

Sub CARTAMONED()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      JU& = 0
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         If TRIM$(Mid$(X$, 2, 27)) <> "" Then
           JU& = JU& + 1
           Call REPLA(X$, Chr$(JU&), 1, 1)
           Call GRAREG("TAMONED", X$, JU&)
         End If
      Next U&
      Call SETULTREG("TAMONED", JU&)
   End If
   Call CIERRARCH("TAMONED")
End Sub

Sub CARTAEMBAL()
   Call TRALOCAL("TAEMBALA", "")
   VTB% = VENTABU%("TAEMBALA/NC/TITUPAN=Embalajes Tipificados")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAEMBALA")
         X$ = REGLEIDO$("!TAEMBALA", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAEMBALA", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAEMBALA")
      Call SETULTREG("TAEMBALA", FIN&)
   End If
   Call CIERRARCH("TAEMBALA")
End Sub

Sub CARTAREGLA()
  '
  Call TRALOCAL("INDIRCOM", "INDIRCOM")
  VTB% = VENTABU%("INDIRCOM/NC/TITUPAN=Reglamentos de Compra")
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!INDIRCOM")
       X$ = REGLEIDO$("!INDIRCOM", U&)
       If CVI(Mid$(X$, 1, 2)) > 0 Then
         If TRIM(Mid$(X$, 3, 46)) <> "" Then
           JJ& = JJ& + 1
           Call GRAREG("INDIRCOM", X$, JJ&)
         End If
       End If
10   Next U&
     Call SETULTREG("INDIRCOM", JJ&)
     Call CIERRARCH("INDIRCOM")
     Call FILESORT("INDIRCOM", "", 1, 1, "ASC")
  End If
  '
End Sub

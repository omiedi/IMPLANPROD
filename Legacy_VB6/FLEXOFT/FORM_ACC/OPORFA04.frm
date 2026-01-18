VERSION 5.00
Begin VB.Form OPORFA04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Orden de Fabricación"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   585
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
   ScaleHeight     =   5850
   ScaleWidth      =   11505
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture16 
      Height          =   5865
      Left            =   10500
      ScaleHeight     =   5805
      ScaleWidth      =   1035
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6000
         Left            =   0
         Picture         =   "OPORFA04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
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
      Begin VB.CommandButton Command10 
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
         Picture         =   "OPORFA04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Tabla de Secciones Productivas"
         Top             =   4100
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   45
         Top             =   4935
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
         Picture         =   "OPORFA04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   300
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
         Picture         =   "OPORFA04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1785
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
         Picture         =   "OPORFA04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   990
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
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
         Picture         =   "OPORFA04.frx":2E00
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprime Formulario de Configuración"
         Top             =   3320
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "More"
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
         Picture         =   "OPORFA04.frx":346A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Otras Opciones de Configuración"
         Top             =   2475
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPORFA04.frx":3774
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5250
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5500
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   10200
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Scrap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   6640
         TabIndex        =   48
         Top             =   2500
         Width           =   1455
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "M.Primas"
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
               Name            =   "MS Sans Serif"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   8190
         TabIndex        =   36
         Top             =   2500
         Width           =   1815
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "V.Materiales"
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
               Name            =   "MS Sans Serif"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   2835
         TabIndex        =   34
         Top             =   1470
         Width           =   2130
         Begin VB.CheckBox Check14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Suprime Impre-  sión (Oculta)"
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
            Left            =   315
            TabIndex        =   35
            Top             =   280
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documentos Anexos (Requiere Sistema de Calidad)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   2500
         Width           =   6345
         Begin VB.CheckBox Check12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hoja de Inspección"
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
            Left            =   4860
            TabIndex        =   32
            Top             =   360
            Width           =   1305
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hojas de Procesos"
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
            Left            =   3420
            TabIndex        =   31
            Top             =   360
            Width           =   1485
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especificaciones"
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
            Left            =   1350
            TabIndex        =   30
            Top             =   315
            Width           =   1905
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Plano"
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
         Left            =   9300
         Picture         =   "OPORFA04.frx":5696
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   525
         Width           =   645
      End
      Begin VB.TextBox Text4 
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
         Left            =   5145
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3960
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
         Height          =   285
         Left            =   8925
         TabIndex        =   21
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   7560
         TabIndex        =   20
         Text            =   " "
         Top             =   525
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fórmula de Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   6195
         TabIndex        =   14
         Top             =   1470
         Width           =   3795
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Multinivel Linealizada"
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
            Left            =   2205
            TabIndex        =   16
            Top             =   280
            Width           =   1485
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lineal (Un Nivel de Explosión"
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
            Left            =   210
            TabIndex        =   15
            Top             =   280
            Value           =   -1  'True
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   210
         TabIndex        =   11
         Top             =   3500
         Width           =   9780
         Begin VB.CheckBox Check13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exige Declarar Sector Productivo"
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
            Left            =   5565
            TabIndex        =   33
            Top             =   1455
            Width           =   4110
         End
         Begin VB.CheckBox Check8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Múltiplos de Lotes Económicos (Tandas)"
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
            Left            =   5565
            TabIndex        =   27
            Top             =   885
            Width           =   4110
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Vale Individual por Operación"
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
            Left            =   5565
            TabIndex        =   26
            Top             =   600
            Width           =   4110
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fracciona por Lotes Máximos"
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
            Left            =   5565
            TabIndex        =   25
            Top             =   315
            Width           =   4110
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes Consumidos al Cierre"
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
            Left            =   525
            TabIndex        =   19
            Top             =   1455
            Width           =   4110
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes al Lanzamiento"
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
            Left            =   525
            TabIndex        =   18
            Top             =   1170
            Width           =   4110
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Verifica Componentes al Lanzamiento"
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
            Left            =   525
            TabIndex        =   17
            Top             =   885
            Width           =   4110
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Observaciones - Leyendas - Comentarios"
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
            Left            =   525
            TabIndex        =   13
            Top             =   600
            Width           =   4110
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Generar O/F's por Subconjuntos"
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
            Left            =   525
            TabIndex        =   12
            Top             =   315
            Width           =   4110
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2625
         TabIndex        =   10
         Text            =   " "
         Top             =   525
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
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   315
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         Picture         =   "OPORFA04.frx":59A0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   525
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
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
         Height          =   920
         Left            =   210
         TabIndex        =   3
         Top             =   1470
         Width           =   2450
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
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
            Left            =   630
            TabIndex        =   5
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
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
            Left            =   630
            TabIndex        =   4
            Top             =   580
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
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
         Height          =   920
         Left            =   5145
         TabIndex        =   1
         Top             =   1470
         Width           =   855
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   190
            TabIndex        =   2
            Text            =   "OF"
            Top             =   370
            Width           =   450
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Auxiliar:"
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
         Left            =   5145
         TabIndex        =   24
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Principal:"
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
         Left            =   210
         TabIndex        =   6
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tabla 
      Caption         =   "Tablas"
      Begin VB.Menu seccionesproductivas 
         Caption         =   "Tabla de Secciones Productivas"
      End
   End
   Begin VB.Menu Configuraciones 
      Caption         =   "Configuraciones Extras"
      Begin VB.Menu confextra 
         Caption         =   "Configuarciones Extras"
      End
   End
End
Attribute VB_Name = "OPORFA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
  'OT-05-0728-WAR-08/11/05///
  Call TRALOCAL("SECPROD", "SECPROD")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left + 200))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
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
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMORFA", Text1.TEXT)
    Call GRACONTROL("", "FORMORFB", Text3.TEXT)
    Call GRACONTROL("", "CODMORFA", Text5.TEXT)
    '
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
    MONUME$ = "VALEMAT": If Option6.Value = True Then MONUME$ = "INCLUPRI"
    Call GRACONTROL$("ORFABRI", "MPRIUNIQ", MONUME$)
    '
    ' \\\OT-5-739-FG-02/12/05///
    SCRAP$ = "MAPRIMA"
    If Option7.Value = True Then
      SCRAP$ = "PRODUCTO"
    End If
    Call GRACONTROL$("ORDEFABR", "MODOSCRA", SCRAP$)
    ' \\\OT-5-739-FG-FIN///
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

Private Sub confextra_Click()
    Call Command9_Click
End Sub

Private Sub Form_Load()
    '
If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      '
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      Me.Check6.BackColor = &H8000000F
      Me.Check7.BackColor = &H8000000F
      Me.Check8.BackColor = &H8000000F
      Me.Check9.BackColor = &H8000000F
      Me.Check10.BackColor = &H8000000F
      Me.Check11.BackColor = &H8000000F
      Me.Check12.BackColor = &H8000000F
      Me.Check13.BackColor = &H8000000F
      Me.Check14.BackColor = &H8000000F
      '
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      '
      MARCOBARRA.Top = 5500
      Set MARCOBARRA.Container = Me.Picture16
      
    End If
    Text1.TEXT = CONTROL$("", "FORMORFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FORMORFB")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    If CONTROL$("", "NUMEORFA") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "GESUBOR") = "SI" Then Check1.Value = 1
    If CONTROL$("", "OBSERVOF") = "SI" Then Check2.Value = 1
    If CONTROL$("", "CHECOMPO") = "SI" Then Check3.Value = 1
    If CONTROL$("", "PRECOLAN") = "SI" Then Check4.Value = 1
    If CONTROL$("", "PRECOCIE") = "SI" Then Check5.Value = 1
    If CONTROL$("", "FRALOMAX") = "SI" Then Check6.Value = 1
    If CONTROL$("", "VALINDOP") = "SI" Then Check7.Value = 1
    If CONTROL$("", "VERITAND") = "SI" Then Check8.Value = 1
    If CONTROL$("", "EXISEPRO") = "SI" Then Check13.Value = 1
    If CONTROL$("ORDEFABR", "PRIPLANO") = "SI" Then Check9.Value = 1
    If CONTROL$("ORDEFABR", "PRIESPEC") = "SI" Then Check10.Value = 1
    If CONTROL$("ORDEFABR", "PRIHOPRO") = "SI" Then Check11.Value = 1
    If CONTROL$("ORDEFABR", "PRIHOINS") = "SI" Then Check12.Value = 1
    If CONTROL$("", "FORMUCON") = "MULTILIN" Then Option4.Value = True
    If CONTROL$("ORDEFABR", "SUPRINT") = "SI" Then Check14.Value = 1
    If CONTROL$("ORFABRI", "MPRIUNIQ") = "INCLUPRI" Then Option6.Value = 1
    ' \\\OT-5-739-FG-02/12/05///
    If CONTROL$("ORDEFABR", "MODOSCRA") = "PRODUCTO" Then Option7.Value = True
    ' \\\OT-5-739-FG-FIN///
    '
    CMO$ = TRIM$(CONTROL$("", "CODMORFA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMANN") > 0 Then
       Command9.Visible = True
    End If
    '
End Sub

Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub imprimir_Click()
   Call Command8_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub seccionesproductivas_Click()
    Call Command10_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
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

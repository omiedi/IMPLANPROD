VERSION 5.00
Begin VB.Form OPREMIT04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Remito de Expedición"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
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
   ScaleHeight     =   4350
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   4845
      Left            =   11025
      ScaleHeight     =   4785
      ScaleWidth      =   1320
      TabIndex        =   35
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox Picture16 
         Height          =   4320
         Left            =   0
         ScaleHeight     =   4260
         ScaleWidth      =   1020
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   4350
            Left            =   0
            Picture         =   "OPREMIT04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command19 
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
         Left            =   120
         Picture         =   "OPREMIT04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   2640
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
         Picture         =   "OPREMIT04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1800
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
         Picture         =   "OPREMIT04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   990
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
         Picture         =   "OPREMIT04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   300
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   3465
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPREMIT04.frx":2C40
         Top             =   3780
         Width           =   1515
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones Avanzadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   5565
      TabIndex        =   20
      Top             =   210
      Width           =   5265
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   46
         Top             =   315
         Width           =   1485
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Número de Serie"
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
            TabIndex        =   47
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Puertos de Impresión - Cant. Copias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   180
         TabIndex        =   41
         Top             =   3045
         Width           =   4905
         Begin VB.TextBox Text10 
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
            Left            =   4305
            TabIndex        =   43
            Top             =   315
            Width           =   330
         End
         Begin VB.TextBox Text12 
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
            Left            =   1155
            TabIndex        =   42
            Top             =   315
            Width           =   2955
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remitos:"
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
            Left            =   -105
            TabIndex        =   44
            Top             =   390
            Width           =   1170
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1890
         TabIndex        =   33
         Top             =   315
         Width           =   1600
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Reemplazo Original"
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
            TabIndex        =   34
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   180
         TabIndex        =   26
         Top             =   315
         Width           =   1600
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Verificar Existencia "
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
            TabIndex        =   27
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Expedición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   180
         TabIndex        =   21
         Top             =   1260
         Width           =   4905
         Begin VB.TextBox Text9 
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
            Left            =   1470
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1260
            Width           =   3330
         End
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
            Left            =   630
            TabIndex        =   29
            Text            =   " "
            Top             =   1260
            Width           =   540
         End
         Begin VB.CommandButton Command6 
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
            Left            =   1155
            TabIndex        =   28
            Top             =   1260
            Width           =   175
         End
         Begin VB.TextBox Text8 
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
            Left            =   1470
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   560
            Width           =   3330
         End
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
            Left            =   630
            TabIndex        =   23
            Text            =   " "
            Top             =   560
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
            Left            =   1155
            TabIndex        =   22
            Top             =   560
            Width           =   175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Precios:"
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
            Left            =   105
            TabIndex        =   31
            Top             =   1000
            Width           =   2325
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito de Salida:"
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
            Left            =   105
            TabIndex        =   25
            Top             =   315
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Remito de Expedición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prefijo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3675
         TabIndex        =   18
         Top             =   2100
         Width           =   1275
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
            Height          =   360
            Left            =   210
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   15
         Top             =   2100
         Width           =   3330
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hoja Suelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1890
            TabIndex        =   17
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Formulario Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   315
            TabIndex        =   16
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Talonario en Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   210
         TabIndex        =   11
         Top             =   3045
         Width           =   4740
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
            Left            =   4370
            TabIndex        =   13
            Top             =   330
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
            Height          =   300
            Left            =   3215
            TabIndex        =   12
            Text            =   " "
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Left            =   735
            TabIndex        =   14
            Top             =   380
            Width           =   2325
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
         Height          =   300
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   420
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
         Height          =   300
         Left            =   3780
         TabIndex        =   9
         Top             =   420
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
         Height          =   300
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
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
         Height          =   720
         Left            =   4095
         Picture         =   "OPREMIT04.frx":4B62
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   750
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   3
         Top             =   1260
         Width           =   3750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
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
            Left            =   315
            TabIndex        =   5
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Auto-Numeración"
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
            Left            =   1785
            TabIndex        =   4
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cód"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4095
         TabIndex        =   1
         Top             =   1260
         Width           =   750
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
            Left            =   140
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "RT"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Label Label2 
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
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   1275
      Left            =   4410
      TabIndex        =   32
      Top             =   105
      Width           =   1140
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
End
Attribute VB_Name = "OPREMIT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\  \Re-Ingrese Ahora a la Aplicación\de Despachos por Remito.")
    Call FINAL("")
End Sub

'\\\OT-05-0241-WAR-02/05/05///
Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORCL")
   Command19.Enabled = True
End Sub
'\\\OT-05-0241-WAR-FIN///

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
       FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
       Call MENSERR(24, "Fecha no Válida")
       Text3.SetFocus
       Exit Sub
    End If
    '
    CMO$ = UCase$(TRIM$(Text5.TEXT))
    If Len(CMO$) <> 2 Or CMO$ = "RF" Then
       Call MENSERR(24, "Código de Movimiento no Válido")
       Text5.TEXT = "RT"
       Text5.SetFocus
       Exit Sub
    End If
    '
    Call GRACONTROL("", "FOREMIFA", Text1.TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text7.TEXT)
    LIPREXPE$ = ""
    If Val(Text6.TEXT) > 0 Then
      If Val(Text6.TEXT) < 256 Then
        If ECOARCH$("LISTAS", Chr$(Val(Text6.TEXT))) <> "" Then
          LIPREXPE$ = TRIM$(Str$(Val(Text6.TEXT)))
        End If
      End If
    End If
    Call GRACONTROL("", "LIPREXPE", LIPREXPE$)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NUREMIFA", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPAREMI", TIPAPEL$)
    '
    MODOEXPE$ = "REMIFAC": TEXPE$ = "Remito de Ventas"
    FOREMIT04.Caption = "Despacho de Materiales - " + TEXPE$
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        FOREMIT04.Caption = FOREMIT04.Caption + "  -  " + EPAC$
    End If
    Call GRACONTROL("", "MODOEXPE", MODOEXPE$)
    '
    Call GRACONTROL("", "FENUREFA", FENUE$)
    '
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "*" Then
      If Len(PREFIREM$) <> 4 Then
        Call COMUNI("Prefijo no Válido")
        Text4.SetFocus
        Exit Sub
      End If
      For U& = 1 To 4
        UI% = Asc(Mid$(PREFIREM$, U&))
        If UI% < 48 Or UI% > 57 Then
            Call COMUNI("Prefijo no Válido")
            Text4.SetFocus
            Exit Sub
        End If
      Next U&
    End If
    Call GRACONTROL("", "PREFIREM", PREFIREM$)
    '
    CMO$ = UCase$(TRIM$(Text5.TEXT))
    Call GRACONTROL("", "CODREFAC", CMO$)
    '
    VERISTOK$ = "NOVERIF"
    If Check1.Value > 0 Then
        VERISTOK$ = "VERISTOK"
    End If
    Call GRACONTROL("REMSTAN", "VERISTOK", VERISTOK$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("REMSTAN", "REPLAORI", MONUME$)
    '
    MONUME$ = "NO"
    If Check3.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("REMSTAN", "NUSERIE", MONUME$)
    '
    Call GRACONTROL(IDENTER$, "PORTREM", Text12.TEXT)
    Call GRACONTROL("", "COPIAREM", Text10.TEXT)
    
    NROPRO& = 0
    Call NUREMIFA(NROPRO&)
    FOREMIT04.Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("*.*")
    'Hide
    '\\\OT-05-0241-WAR-02/05/05///
    Unload Me 'DESCARGUA EL FORMULARIO DESPUES DEL SETUP
    '\\\OT-05-0241-WAR-FIN///
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
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("LISTAS")
    End If
End Sub

Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text9_Click()
   Text6.SetFocus
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("LISTAS")
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
'
Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      
      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      Me.Check3.BackColor = &HFCD7B6
      
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
      Me.Frame13.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 3950
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If
    '
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        x$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(x$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(y$, Left$(x$, 63), 2, 63)
                 Call GRAREG("TADEPOSI", y$, J&)
              End If
           End If
        End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text1.TEXT = CONTROL$("", "FOREMIFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FENUREFA")
    Text7.TEXT = CONTROL$("", "DEPOEXPE")
    Text8.TEXT = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    Text6.TEXT = CONTROL$("", "LIPREXPE")
    Text9.TEXT = ECOARCH$("LISTAS", Chr$(Val(Text6.TEXT)))
    '
    Text12.TEXT = CONTROL$(IDENTER$, "PORTREM")
    Text10.TEXT = CONTROL$("", "COPIAREM")

    If CONTROL$("", "NUREMIFA") = "AUTONUME" Then Option2.Value = True
    If CONTROL$("", "TIPAREMI") = "HOJASUEL" Then Option4.Value = True
    'If CONTROL$("", "MODOEXPE") = "ODESPACH" Then Option6.Value = True
    '
    PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Text4.TEXT = PREFIREM$
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       Check1.Value = 1
    End If
    '\\\OT-05-0241-WAR-02/05/05///
    Command19.Enabled = False
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       Check2.Value = 1
       Command19.Enabled = True
    End If
    If CONTROL$("REMSTAN", "NUSERIE") = "SI" Then
       Check3.Value = 1
    End If
    '\\\OT-05-0241-WAR-FIN///
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


Private Sub Text3_DblClick()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text8_Click()
   Text7.SetFocus
End Sub

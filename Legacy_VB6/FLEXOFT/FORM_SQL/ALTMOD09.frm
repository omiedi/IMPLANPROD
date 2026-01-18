VERSION 5.00
Begin VB.Form ALTMOD09 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de Datos de Modelos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8730
      Left            =   10800
      ScaleHeight     =   8670
      ScaleWidth      =   1395
      TabIndex        =   52
      Top             =   -45
      Width           =   1455
      Begin VB.Frame Frame8 
         BackColor       =   &H00400000&
         Caption         =   "Opcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2265
         Left            =   110
         TabIndex        =   62
         Top             =   5505
         Width           =   875
         Begin VB.CommandButton AGenRep 
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
            Left            =   105
            Picture         =   "ALTMOD09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   1560
            Width           =   700
         End
         Begin VB.CommandButton Command24 
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
            Left            =   105
            Picture         =   "ALTMOD09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Movimientos - Entrega y Recepción"
            Top             =   240
            Width           =   700
         End
         Begin VB.CommandButton Command7 
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
            Left            =   105
            Picture         =   "ALTMOD09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Transacciones de Inspección"
            Top             =   900
            Width           =   700
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00400000&
         Caption         =   "Config"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1635
         Left            =   110
         TabIndex        =   59
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   3825
         Width           =   875
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
            Height          =   660
            Left            =   105
            Picture         =   "ALTMOD09.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   300
            Width           =   700
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
            Height          =   660
            Left            =   105
            Picture         =   "ALTMOD09.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
            Top             =   900
            Width           =   700
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00400000&
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3720
         Left            =   110
         TabIndex        =   53
         Top             =   90
         Width           =   875
         Begin VB.CommandButton Command3 
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
            Height          =   660
            Left            =   105
            Picture         =   "ALTMOD09.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Guarda los Cambios Realizados"
            Top             =   900
            Width           =   700
         End
         Begin VB.CommandButton Command1 
            Caption         =   "New"
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
            Left            =   105
            Picture         =   "ALTMOD09.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Alta de Nuevo Modelo"
            Top             =   2340
            Width           =   700
         End
         Begin VB.CommandButton Command2 
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
            Left            =   105
            Picture         =   "ALTMOD09.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   1560
            Width           =   700
         End
         Begin VB.CommandButton Command9 
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
            Height          =   660
            Left            =   105
            Picture         =   "ALTMOD09.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Cierra y Sale de la Aplicación"
            Top             =   240
            Width           =   700
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
            Height          =   660
            Left            =   105
            Picture         =   "ALTMOD09.frx":1912
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Baja del Modelo Activo"
            Top             =   3000
            Width           =   700
         End
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   -435
         Picture         =   "ALTMOD09.frx":1C1C
         Top             =   7860
         Width           =   2010
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1500
      Left            =   105
      TabIndex        =   39
      Top             =   6720
      Width           =   10620
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   42
         Top             =   315
         Width           =   10580
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   44
            Top             =   75
            Width           =   10200
         End
         Begin VB.Label Label10 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   43
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "O.K"
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
         Left            =   8280
         Picture         =   "ALTMOD09.frx":5206
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   760
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   735
         Width           =   7155
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
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
         Height          =   750
         Left            =   8295
         TabIndex        =   47
         Top             =   735
         Width           =   2265
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
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
         Height          =   750
         Left            =   0
         TabIndex        =   46
         Top             =   735
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   45
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1480
         Y2              =   1480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1335
      Left            =   5775
      TabIndex        =   24
      Top             =   2100
      Width           =   4950
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   875
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   315
         Width           =   4530
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Modelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1905
      Left            =   105
      TabIndex        =   12
      Top             =   105
      Width           =   10620
      Begin VB.CommandButton Command12 
         Height          =   520
         Left            =   9870
         Picture         =   "ALTMOD09.frx":5510
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Acceso a B.Datos de Fotos y Planos"
         Top             =   250
         Width           =   540
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
         Left            =   9450
         TabIndex        =   49
         Text            =   " "
         Top             =   840
         Width           =   960
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   645
         Left            =   105
         TabIndex        =   26
         Top             =   1200
         Width           =   10410
         Begin VB.OptionButton Option8 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Anual"
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
            Left            =   7035
            TabIndex        =   27
            Top             =   45
            Width           =   960
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Trim."
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
            Left            =   6090
            TabIndex        =   29
            Top             =   45
            Width           =   1380
         End
         Begin VB.TextBox Text7 
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
            Left            =   1155
            TabIndex        =   32
            Text            =   " "
            Top             =   210
            Width           =   2220
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Seman."
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
            Left            =   4935
            TabIndex        =   31
            Top             =   45
            Width           =   1380
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Mens."
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
            Left            =   5460
            TabIndex        =   30
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Semes"
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
            Left            =   6510
            TabIndex        =   28
            Top             =   270
            Width           =   1380
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Ubicación:"
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
            TabIndex        =   38
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Frecuencia de Verificación:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3570
            TabIndex        =   37
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "U.Control:"
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
            Left            =   8295
            TabIndex        =   36
            Top             =   45
            Width           =   960
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Prox.Control:"
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
            Left            =   7980
            TabIndex        =   35
            Top             =   345
            Width           =   1275
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   9345
            TabIndex        =   34
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Label9 
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
            Height          =   255
            Left            =   9345
            TabIndex        =   33
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   600
         Left            =   3780
         TabIndex        =   16
         Top             =   180
         Width           =   6000
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Otros:"
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
            Left            =   2730
            TabIndex        =   21
            Top             =   270
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Madera"
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
            Left            =   315
            TabIndex        =   19
            Top             =   270
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Aluminio"
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
            Left            =   1470
            TabIndex        =   18
            Top             =   270
            Width           =   1275
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
            Height          =   285
            Left            =   3570
            TabIndex        =   17
            Text            =   " "
            Top             =   210
            Width           =   2115
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
         Height          =   280
         Left            =   3465
         TabIndex        =   15
         Top             =   420
         Width           =   180
      End
      Begin VB.TextBox Text2 
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
         Left            =   1260
         TabIndex        =   14
         Text            =   " "
         Top             =   840
         Width           =   6840
      End
      Begin VB.TextBox Text3 
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
         Left            =   1275
         TabIndex        =   13
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFD0&
         Caption         =   "Cant.Piezas:"
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
         Left            =   7455
         TabIndex        =   50
         Top             =   900
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Index           =   2
         Left            =   -735
         TabIndex        =   48
         Top             =   900
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Index           =   0
         Left            =   -735
         TabIndex        =   20
         Top             =   480
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Pieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1335
      Left            =   105
      TabIndex        =   5
      Top             =   2100
      Width           =   5580
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   9
         Top             =   315
         Width           =   180
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
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Index           =   1
         Left            =   -735
         TabIndex        =   23
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label5 
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
         Height          =   435
         Left            =   1260
         TabIndex        =   22
         Top             =   735
         Width           =   4110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFD0&
         Caption         =   "U.Med:"
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
         Left            =   3780
         TabIndex        =   11
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   4830
         TabIndex        =   10
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3090
      Left            =   105
      TabIndex        =   0
      Top             =   3570
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   4005
         Left            =   10290
         TabIndex        =   7
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   4
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   3
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   0
         TabIndex        =   2
         Top             =   735
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   1
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   6
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
End
Attribute VB_Name = "ALTMOD09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIDES%

Private Sub AGenRep_Click()
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub


'
Private Sub Command1_Click()
    Command1.Enabled = False
    '
10  VDEF$ = String$(66, 0)
    '
15  OKX% = 0
    OBX$ = OBJPLA$("ALTAMODELO", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    MODENUE$ = TRIM$(Mid$(OBX$, 1, 16))
    DESCNUE$ = TRIM$(Mid$(OBX$, 17, 48))
    '
    If MODENUE$ = "" Then
      Call COMUNI("Falta Codigo de Modelo")
      GoTo 15
    End If
    '
    If DESCNUE$ = "" Then
      Call COMUNI("Falta Descripción del Modelo")
      GoTo 15
    End If
    '
    FULINSP% = CVI(Mid$(OBX$, 65, 2))
    If FULINSP% < 1 Or FULINSP% > HOY(HOS$) Then
      Call MENSERR(24, "Mandatorio Indicar Fecha de Ultima Inspección")
      GoTo 15
    End If
    '
'    If ECOARCH$("BADAMODE", MODENUE$) = "" Then
'      X$ = REGBLAN$("BADAMODE")
'      Call REPLA(X$, MODENUE$, 1, 16)
'      Call REPLA(X$, DESCNUE$, 17, 48)
'      ' Valores por omision para Frecuencia, Ult.Control y Prox.Control
'      Call REPLA(X$, "1 - Semanal", 97, 16)
'        FEX = FULINSP%
'      Call REPLA(X$, MKI$(FULINSP%), 113, 2)
'      Call REPLA(X$, MKI$(0), 115, 2)
'      Call REGAPP("BADAMODE", X$)
'      Call FRESHECHO("BADAMODE")
'    Else
'      Call COMUNI("Código de Modelo Existente")
'      GoTo 10
'    End If
    '
'    Text3 = Mid$(X$, 1, 16)
    '
    CONDI$ = "Cod_Mod ='" & MODENUE$ & "'"
    CODIMOD& = CantRegistros("BADAMODE", CONDI$)
    If CODIMOD& = 0 Then
       Call CreaRegistro("BADAMODE", "COD_MOD", MODENUE$)
       Call ACTUREGISTRO("BADAMODE", "DESCRIPCION", CONDI$, Left$(DESCNUE$, 64), REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "UBIC", CONDI$, "", REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "CANTPIEZAS", CONDI$, 0, REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "FRECUENCIA", CONDI$, "1 - Semanal", REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "FECHULT_CON", CONDI$, FETEXCOR1$(FULINSP%), REGAF&, "NOMESS")
            FF% = FECHANUM("01/01/80")
       Call ACTUREGISTRO("BADAMODE", "FECHPRO_CON", CONDI$, FETEXCOR1$(FF%), REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "MATERIAL", CONDI$, 0, REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "DESC_MAT", CONDI$, "", REGAF&, "NOMESS")
       Call ACTUREGISTRO("BADAMODE", "COD_INTE", CONDI$, 0, REGAF&, "NOMESS")
       '
'       SQLX$ = "SELECT * FROM BADAMODE"
'       SQLX$ = SQLX$ + " WHERE Cod_Mod > 0"
'       Dim NEWMOD As ADODB.Recordset
'       Set NEWMOD = New ADODB.Recordset
'       NEWMOD.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'       With NEWMOD
'            .AddNew
'            !Cod_Mod = Left$(MODENUE$, 16)
'            !Descripcion = Left$(DESCNUE$, 64)
'            !Ubic = ""
'            !CantPiezas = 0
'            !Frecuen = "1 - Semanal"
'            !FechUlt_Con = CVDAT(FULINSP%)
'             FF% = FECHANUM("01/01/80")
'            !FechPro_Con = CVDAT(FF%)
'            !MATERIAL = 0
'            !DESC_MAT = ""
'            !COD_INTE = 0
'            .Update
'       End With
'       NEWMOD.Close
'       Set NEWMOD = Nothing
    Else
      Call COMUNI("Código de Modelo Existente")
      GoTo 10
    End If
    Text3 = MODENUE$
    Text2 = DESCNUE$
   '
99  Command1.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   '
   VERIMODE.Show 1
'   Call COPYSTRU("INDIPEDI", "PEPENDES")
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
'   For U& = 1 To FIN& Step 4
'     Call TRACE(20, U&, FIN&)
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     NPX& = CVS(Left$(XX$, 4))
'     If NPX& <> NPXA& Then
'       NPXA& = NPX&
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
'       FEPE = CVI(Mid$(XX$, 5, 2))
'       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
'       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
'       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
'       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
'       JJ& = JJ& + 1
'       Call GRAREG("PEPENDES", YY$, JJ&)
'     End If
'   Next U&
'   '
'   Call SETULTREG("PEPENDES", JJ&)
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
'   NPX& = Val(VALACT1$("PEPENDES"))
'   If NPX& < 1 Then GoTo 99
'   '
'   Label7 = TRIM$(Str$(NPX&))
'   '
99 Command21.Enabled = True
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   CONDI$ = "Cod_Mod = '" & Text3 & "'"
   If SAFETEXT$(VALOBADA("BADAMODE", "Cod_Mod", CONDI$)) <> "" Then
        CODMODE$ = AJUSTI$(Text3, 16)
        '
     If TRIM$(Text1) <> "" Then
        Call FRATEXTO(Text1, 56)
        For KKI% = 1 To CARETEX%
          TTXX$ = Space$(76)
          Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
          If KKI% = 1 Then
            Call REPLA(TTXX$, HOS$, 1, 8)
            Call REPLA(TTXX$, USERNAME$, 69, 24)
          End If
          Call REPLA(CODMODE$, ALTMOD09.Text3, 1, 16)
          'Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
          Call AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
        Next KKI%
        Call AGREGA_ANOTAMODELO(CODMODE$, String$(88, "="))
        'Call REGAPP("ANOTAMOD", CODMODE$ + String$(88, "="))
        'Call CIERRARCH("ANOTAMOD")
        Text1 = ""
        '
        List1.Clear
        SQLX$ = "SELECT * FROM BADAMODE"
        SQLX$ = SQLX$ + " WHERE Cod_Mod = '" & TRIM$(CODMODE$) & "'"
        '
        Dim ANOTMOD As New ADODB.Recordset
        Set ANOTMOD = New ADODB.Recordset
25      On Error Resume Next
        ANOTMOD.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With ANOTMOD
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             On Error GoTo 0
             XX$ = SAFETEXT$(!ANOTAMOD): LONGI& = Len(XX$)
             INICIO& = 1
             Do While INICIO& < LONGI&
                 FIN& = InStr(INICIO&, XX$, Chr$(10))
                 If FIN& > 0 Then
                    TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
                    List1.AddItem TTXX$
                    INICIO& = FIN& + 2
                 Else
                    TTXX$ = Mid$(XX$, INICIO&)
                    List1.AddItem TTXX$
                    Exit Do
                 End If
             Loop
          End If
          On Error GoTo 0
        End With
        ANOTMOD.Close
        Set ANOTMOD = Nothing
'       For U& = 1 To Len(RFF$) Step 4
'         RENOTA& = CVS(Mid$(RFF$, U&, 4))
'         XX$ = REGLEIDO$("ANOTAMOD", RENOTA&)
'         ALTAMODE.List1.AddItem Mid$(XX$, 17, 92)
'       Next U&
'       '
     End If
   End If
   '
   Text1 = ""
   '
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   OPX% = COMALTER%("Base de Datos Maestra a Actualizar:\\Maestro de Items de Stock (Piezas)\Proveedores y Contratistas")
   If OPX% = 1 Then
       Call CONECRUN("AMASTO07", "Archivos Maestros")
     ElseIf OPX% = 2 Then
       Call CONECRUN("ARCHIG07/AMAPRO", "Archivos Maestros")
   End If
   Command11.Enabled = True
End Sub


Private Sub Command12_Click()
   Command12.Enabled = False
   CONDI$ = "Cod_Mod = '" & Text3 & "'"
   CODIMOD& = CantRegistros("BADAMODE", CONDI$)
   If CODIMOD& = 1 Then
     Call FORPLA09.CARPLANOS
     FORPLA09.Show 1
   End If
   Command12.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
'   OPX% = COMALTER%("Opciones de Listados:\\B.Datos Completa\Pendientes de Inspección\Ficha de Inspección")
'   CORRESPO$ = "Corresponde a: "
'   Select Case OPX%
'     Case 1    ' Base de Datos Completa
'       Call COPYSTRU("BADAMODE", "BADAMODX")
'       '
'       CORRESPO$ = CORRESPO$ + "B.Datos Completa - "
'       FTX% = COMALTER%("Filtrado por Tipo de Material:\\Todos\Madera\Aluminio\Otros") - 1
'         If FTX% < 0 Or FTX% > 3 Then GoTo 99
'       If FTX% < 1 Then
'           CORRESPO$ = CORRESPO$ + "Todos -"
'           Call CIERRARCH("BADAMODE")
'           Call CIERRARCH("BADAMODX")
'           Call BAJALDISCO
'           Call COPYFILE(LUDAT$ + "BADAMODE.DAT", LUDAT$ + "BADAMODX.DAT")
'         Else
'           If FTX% = 1 Then CORRESPO$ = CORRESPO$ + "Madera - "
'           If FTX% = 2 Then CORRESPO$ = CORRESPO$ + "Aluminio - "
'           If FTX% = 3 Then CORRESPO$ = CORRESPO$ + "Otros - "
'           Call FILTERFILE("BADAMODE", "BADAMODX", 7, TRIM$(Str$(FTX%)))
'       End If
'       '
'       If ULTREG&("BADAMODX") < 1 Then
'         Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
'         GoTo 99
'       End If
'       '
'       OPY% = COMALTER%("Opciones Alternativas de Ordenamiento:\\Por Modelo\Por Pieza\Por Fecha Próximo Control")
'         If OPY% < 1 Or OPY% > 3 Then GoTo 99
'       ASORT% = 1: CORRESPO$ = CORRESPO$ + "Por Modelo"
'       If OPY% = 2 Then ASORT% = 9: CORRESPO$ = CORRESPO$ + "Por Pieza"
'       If OPY% = 3 Then ASORT% = 6: CORRESPO$ = CORRESPO$ + "Por Fecha Proximo Control"
'       Call FILESORT("BADAMODX", "", ASORT%, ASORT%, "ASC")
'       '
'       Call HEADERS("BADAMODX", "")
'       Call HEADERS("BADAMODX", CORRESPO$)
'       Call FINAL("*LIBAMOD")
'       '
'     Case 2    ' Pendientes de Inspeccion
'       Call COPYSTRU("BADAMODE", "BADAMODX")
'       '
'       Call SELECHO("SELFECHA/Prox.Insp.Hasta")
'       VDEVU$ = VALACT1$("SELFECHA")
'       If VDEVU$ <> "" Then
'         '
'         CORRESPO$ = CORRESPO$ + "Proximo Control Hasta " + VDEVU$ + " - "
'         ' FILTRADO POR FECHA
'         FLIM% = FECHANUM(VDEVU$)
'         FIN& = ULTREG&("BADAMODE")
'         Call COPYSTRU("BADAMODE", "BADAMODY")
'         JJ& = 0
'         For U& = 1 To FIN&
'           XX$ = REGLEIDO$("BADAMODE", U&)
'           If CVI(Mid$(XX$, 115, 2)) <= FLIM% Then
'             JJ& = JJ& + 1
'             Call GRAREG("BADAMODY", XX$, JJ&)
'           End If
'         Next U&
'         Call SETULTREG("BADAMODY", JJ&)
'         Call CIERRARCH("BADAMODY")
'         '
'         If ULTREG&("BADAMODY") < 1 Then
'           Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
'           GoTo 99
'         End If
'         '
'         FTX% = COMALTER%("Filtrado por Tipo de Material:\\Todos\Madera\Aluminio\Otros") - 1
'           If FTX% < 0 Or FTX% > 3 Then GoTo 99
'         If FTX% < 1 Then
'             CORRESPO$ = CORRESPO$ + "Todos -"
'             Call CIERRARCH("BADAMODE")
'             Call CIERRARCH("BADAMODX")
'             Call BAJALDISCO
'             Call COPYFILE(LUDAT$ + "BADAMODY.DAT", LUDAT$ + "BADAMODX.DAT")
'           Else
'             If FTX% = 1 Then CORRESPO$ = CORRESPO$ + "Madera - "
'             If FTX% = 2 Then CORRESPO$ = CORRESPO$ + "Aluminio - "
'             If FTX% = 3 Then CORRESPO$ = CORRESPO$ + "Otros - "
'             Call FILTERFILE("BADAMODY", "BADAMODX", 7, TRIM$(Str$(FTX%)))
'         End If
'         '
'         If ULTREG&("BADAMODX") < 1 Then
'           Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
'           GoTo 99
'         End If
'         '
'         OPY% = COMALTER%("Opciones Alternativas de Ordenamiento:\\Por Modelo\Por Pieza\Por Fecha Próximo Control")
'           If OPY% < 1 Or OPY% > 3 Then GoTo 99
'         ASORT% = 1: CORRESPO$ = CORRESPO$ + "Por Modelo"
'         If OPY% = 2 Then ASORT% = 9: CORRESPO$ = CORRESPO$ + "Por Pieza"
'         If OPY% = 3 Then ASORT% = 6: CORRESPO$ = CORRESPO$ + "Por Fecha Proximo Control"
'         Call FILESORT("BADAMODX", "", ASORT%, ASORT%, "ASC")
'         '
'         Call HEADERS("BADAMODX", "")
'         Call HEADERS("BADAMODX", CORRESPO$)
'         Call FINAL("*LIBAMOD")
'         '
'       End If
'       '
'     Case 3
'        '
'        FORIPRE$ = TRIM$(CONTROL$("", "FICHINMO"))
'        If FORIPRE$ = "" Then Exit Sub
'        '
'        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
'          Call MENSERR(24, "Imposible Imprimir Solicitud Interna de Materiales.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
'          Exit Sub
'        End If
'        '
'        Screen.MousePointer = 11
'        '
'        If TRIM$(Text3) = "" Then
'           Call MENSERR(24, "Imposible Imprimir Ficha de Inspección. Falta Código de Modelo")
'           Screen.MousePointer = 0
'           Exit Sub
'        End If
'        '
'        CODTABU1$(1) = "FECH-LAN"
'        CODTABU1$(2) = "REF-MAT3"
'        CODTABU1$(3) = "NRO-LOTE"
'        CODTABU1$(4) = "ZONA-VEN"
'        CACOLTAB1% = 4
'        '
'        CAITAB1% = 0
'        For I = 0 To (List1.ListCount - 1)
'          CAITAB1% = CAITAB1% + 1
'          If Left$(List1.List(I), 1) <> "=" Then
'            FECHA$ = Mid$(List1.List(I), 1, 8)
'            DETALLE$ = Mid$(List1.List(I), 11, 56)
'            Usuario$ = Mid$(List1.List(I), 69, 17)
'            TETABU1$(1, CAITAB1%) = FECHA$
'            TETABU1$(2, CAITAB1%) = DETALLE$
'            TETABU1$(3, CAITAB1%) = Usuario$
'            TETABU1$(4, CAITAB1%) = ""
'          Else
'            FECHA$ = String$(8, "=")
'            DETALLE$ = String$(60, "=")
'            Usuario$ = String$(13, "=")
'            FIRMA$ = String$(13, "=")
'            TETABU1$(1, CAITAB1%) = FECHA$
'            TETABU1$(2, CAITAB1%) = DETALLE$
'            TETABU1$(3, CAITAB1%) = Usuario$
'            TETABU1$(4, CAITAB1%) = FIRMA$
'          End If
'        Next I
'        '
'        'CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = "001"
'        CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = TRIM$(Text5)
'        CODPARAM$(3) = "DES-MAT"
'        If Option1.Value Then
'          DOCPARAM$(3) = "Madera"
'        ElseIf Option2.Value Then
'          DOCPARAM$(3) = "Alumnio"
'        Else
'          DOCPARAM$(3) = TRIM$(Text4)
'        End If
'        CODPARAM$(4) = "COD-MAT": DOCPARAM$(4) = TRIM$(Text3)
'        CODPARAM$(5) = "UBI-MPRI": DOCPARAM$(5) = TRIM$(Text7)
'        CODPARAM$(6) = "CANTIDAD": DOCPARAM$(6) = TRIM$(Text8)
'        CODPARAM$(7) = "REF-MAT2"
'        If Option4.Value Then
'          DOCPARAM$(7) = "Semanal"
'        ElseIf Option5.Value Then
'          DOCPARAM$(7) = "Mensual"
'        ElseIf Option6.Value Then
'          DOCPARAM$(7) = "Trimestral"
'        ElseIf Option7.Value Then
'          DOCPARAM$(7) = "Semestral"
'        Else
'          DOCPARAM$(7) = "Anual"
'        End If
'        CAPARDOC% = 7
'        '
'        Call BLANARCH("!OBSERVOF")
'        Call FRATEXTO(Text6.TEXT, 35)
'        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
'        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
'        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
'        If CARETEX% >= 4 Then Call GRAREG("!OBSERVOF", RETEX$(4), 4)
'        URETE& = CARETEX%
'        Call SETULTREG("!OBSERVOF", URETE&)
'        Call CIERRARCH("!OBSERVOF")
'        '
'        FORIPRE$ = TRIM$(CONTROL$("", "FICHINMO"))
'        If FORIPRE$ <> "" Then
'          Call PRINTDOC("FICHINMO")
'        End If
'        '
'     Case Else
'        GoTo 99
'        '
'   End Select
'   '
99 Command23.Enabled = True
'   '
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
'   If ECOARCH("BADAMODE", Text3) <> "" Then
     OPX% = ALTERNA%("Despachar Modelo A Proveedor\Recibir Modelo de Proveedor")
     If OPX% = 1 Then
         Call DESPAMODE
       ElseIf OPX% = 2 Then
         Call RECIMODE
     End If
'   End If
   Command24.Enabled = True
End Sub

Private Sub Command26_Click()
'   Command26.Enabled = False
'   CFGDESPA.Show 1
'   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    CONDI$ = "Cod_Mod = '" & Text3 & "'"
    CODIMOD& = CantRegistros("BADAMODE", CONDI$)
    If CODIMOD& = 1 Then
        CONDI$ = "Cod_Mod = '" & Text3 & "'"
        Call ACTUREGISTRO("BADAMODE", "DESCRIPCION", CONDI$, Text2, REG&, "NOMESS")
        ' Material
        If Option1.Value Then Call ACTUREGISTRO("BADAMODE", "MATERIAL", CONDI$, 1, REG&, "NOMESS")
        If Option2.Value Then Call ACTUREGISTRO("BADAMODE", "MATERIAL", CONDI$, 2, REG&, "NOMESS")
        If Option3.Value Then ' Selecciono Otros
          Call ACTUREGISTRO("BADAMODE", "MATERIAL", CONDI$, 3, REG&, "NOMESS")
          Call ACTUREGISTRO("BADAMODE", "DESC_MAT", CONDI$, Text4, REG&, "NOMESS")
        End If
        Call ACTUREGISTRO("BADAMODE", "UBIC", CONDI$, Text7, REG&, "NOMESS") ' Ubicacion
        '    ' Frecuencia
        If Option4.Value Then Call ACTUREGISTRO("BADAMODE", "FRECUEN", CONDI$, "1 - Semanal", REG&, "NOMESS")
        If Option5.Value Then Call ACTUREGISTRO("BADAMODE", "FRECUEN", CONDI$, "2 - Mensual", REG&, "NOMESS")
        If Option6.Value Then Call ACTUREGISTRO("BADAMODE", "FRECUEN", CONDI$, "3 - Trimestral", REG&, "NOMESS")
        If Option7.Value Then Call ACTUREGISTRO("BADAMODE", "FRECUEN", CONDI$, "4 - Semestral", REG&, "NOMESS")
        If Option8.Value Then Call ACTUREGISTRO("BADAMODE", "FRECUEN", CONDI$, "5 - Anual", REG&, "NOMESS")
        ' Valida que el codigo sea valido
        CIXI% = CODINT%(Text5)
        If Text5 <> "" Then
          If CIXI% < 1 Then
            Call COMUNI("Código de Pieza No Válido")
            GoTo 99
          End If
        End If
        FFCONT% = FECHANUM(Label6)           'SE AGREGA GRABACION DE FECHA DE ÚLTIMO CONTROL
        Call ACTUREGISTRO("BADAMODE", "FECHULT_CON", CONDI$, CVDAT(FFCONT%), REG&, "NOMESS")
        '
        PROCON% = FECHANUM(Label9)
        Call ACTUREGISTRO("BADAMODE", "FECHPRO_CON", CONDI$, CVDAT(PROCON%), REG&, "NOMESS")
        Call ACTUREGISTRO("BADAMODE", "COD_INTE", CONDI$, CIXI%, REG&, "NOMESS")
        '
        TTYZ$ = Text6.TEXT
        Call ACTUREGISTRO("BADAMODE", "ObservaMod", CONDI$, TTYZ$, REG&, "NOMESS")
        CANPIEZAS = XVALO(Text8)
        Call ACTUREGISTRO("BADAMODE", "Cantpiezas", CONDI$, CANPIEZAS, REG&, "NOMESS")
    End If
   '
99 Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   Command1.Enabled = False
   '
   CONDI$ = "Cod_Inte >= 0 order by Cod_Mod Asc"
   Call CARGALISEL("BADASQL", "BADAMODE", "Cod_Mod/A16;Descripcion/A30", CONDI$)
   Call FRESHECHO("BADASQL")
   Call SELECHO("BADASQL/Archivo de Modelos")
  
   NCXI$ = TRIM$(VALACT1$("BADASQL"))
   If NCXI$ <> "" Then
      Text3 = NCXI$
   End If
99 Command1.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   OPX% = ALTERNA%("Configuración General\Formularios")
   If OPX% = 1 Then
       Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
       Call FINAL("")
     ElseIf OPX% = 2 Then
       OPALMOD09.Show 1
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
'   If ECOARCH("BADAMODE", Text3) <> "" Then
    CONDI$ = "Cod_Mod = '" & Text3 & "'"
    If SAFETEXT$(VALOBADA("BADAMODE", "Cod_Mod", CONDI$)) <> "" Then
     If DERACCE%("BAJAMODE", "Dar de Baja Modelo de Almacén") >= 2 Then
       If COMALTER%("Confirme la Baja Definitiva del Modelo\'" + TRIM$(Text3) + "'.\\Cancelar Baja\Confirmar") = 2 Then
'         Call FILTERDELETE("BADAMODE", 1, AJUSTI$(Text3, 16))
         Call BORRAREGISTROS("BADAMODE", CONDI$, REG&, "NOMESS")
         Call COMUNI("Baja Definitiva Procesada")
         Call BLANFORMU
         Text2 = ""
         Text3 = ""
         CONDI$ = "Cod_Inte >= 0 order by Cod_Mod Asc"
         Call CARGALISEL("BADASQL", "BADAMODE", "Cod_Mod/A16;Descripcion/A30", CONDI$)
         Call FRESHECHO("BADASQL")
       End If
     End If
   End If
99 Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("REGINSPE", "Registro de Inspección de Modelos") >= 2 Then
       With NOVMOD09
         .Text1 = ""
         .Label3 = Text3
         .Label6 = Text2
         .Label7 = Text5
         .Label8 = Label5
         .Text10 = FECHATEX$(HOY(HOS$))
         .Label9 = USERNAME
         .Option1.Value = True
       End With
       '
       NOVMOD09.Show 1
       '
       Call CARGANOTA(Text3)
   End If
   '
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text5.TEXT = COD$
   End If
   
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    DoEvents
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    'Label9 = " " + HOS$
    'Label12 = Left$(USERNAME, 17)
    '
    Call ABRECONEXION
    Call ABREMASTER
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label2_Click()
    '
'    If ECOARCH$("BADAMODE", ALTAMODE.Text3) <> "" Then
    If Text3 <> "" Then
      NOVMOD09.Show 1
    Else
      Call COMUNI("Error en Código de Modelo")
    End If
    '
End Sub

Private Sub Option1_GotFocus()
    '
    Text4.Enabled = False
    '
End Sub

Private Sub Option2_GotFocus()
    '
    Text4.Enabled = False
    '
End Sub

Private Sub Option3_GotFocus()
    '
    Text4.Enabled = True
    '
End Sub

Private Sub Option4_Click()
   Call CALPROCON(7)
End Sub

Private Sub Option5_Click()
   Call CALPROCON(30)
End Sub

Private Sub Option6_Click()
   Call CALPROCON(90)
End Sub

Private Sub Option7_Click()
   Call CALPROCON(180)
End Sub

Private Sub Option8_Click()
   Call CALPROCON(365)
End Sub

Private Sub Text2_Change()
   MODIDES% = 1
End Sub

Private Sub Text3_Change()
   '
   Call BLANFORMU
   '
   CONDI$ = "Cod_Mod = '" & Text3 & "'"
   Text2 = VALOBADA("BADAMODE", "Descripcion", CONDI$, "NOMESS")
   MODIDES% = 0
   Text7 = VALOBADA("BADAMODE", "Ubic", CONDI$, "NOMESS")
   NROMAT% = XVALO(VALOBADA("BADAMODE", "Material", CONDI$, "NOMESS"))
   '
   Select Case NROMAT%
         Case 1: Option1.Value = True
         Case 2: Option2.Value = True
         Case 3: Option3.Value = True
         '
         Text4 = VALOBADA("BADAMODE", "Desc_Mat", CONDI$, "NOMESS")
   End Select
   '
   FRECUENCIA% = XVALO(VALOBADA("BADAMODE", "Frecuen", CONDI$, "NOMESS"))
   FF% = CVINT(VALOBADA("BADAMODE", "FechUlt_Con", CONDI$, "NOMESS"))
   Label6 = FECHATEX(FF%)
   Select Case FRECUENCIA%
         Case 1: Option4.Value = True: Call CALPROCON(7)
         Case 2: Option5.Value = True: Call CALPROCON(30)
         Case 3: Option6.Value = True: Call CALPROCON(90)
         Case 4: Option7.Value = True: Call CALPROCON(180)
         Case 5: Option8.Value = True: Call CALPROCON(365)
   End Select
   '
   CI% = XVALO(VALOBADA("BADAMODE", "Cod_Inte", CONDI$, "NOMESS"))
   If CI% > 0 Then
       Label5 = DESCRIT0$(CI%)
       Text5 = CODEXT$(CI%)
       Label3 = UNIMED$(CI%)
   End If
   '
   CANPIEZAS = VALOBADA("BADAMODE", "CantPiezas", CONDI$, "NOMESS")
   Text8 = FORMATNUM$(CANPIEZAS, "F8.1")
   '
   Call CARGANOTA(Text3)
   '
   Text6 = VALOBADA("BADAMODE", "ObservaMod", CONDI$, "NOMESS")
End Sub

Sub BLANFORMU()
    '
    Option1.Value = True
    Text4 = ""
    Text1 = ""
    Text2 = ""
    Text7 = ""
    Option4.Value = True
    Label6 = ""
    Label9 = ""
    Text5 = ""
    Label5 = ""
    Label3 = ""
    Text6 = ""
    List1.Clear
    Text8 = ""
    '
End Sub

Sub CARGANOTA(COMODE$)
    '
    List1.Clear
    If COMODE$ <> "" Then
        XX$ = VALOBADA("BADAMODE", "ANOTAMOD", "Cod_Mod = '" & COMODE$ & "'")
        LONGI& = Len(XX$): INICIO& = 1
        '
        Do While INICIO& < LONGI&
           FIN& = InStr(INICIO&, XX$, Chr$(10))
           If FIN& > 0 Then
               TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
               List1.AddItem TTXX$
               INICIO& = FIN& + 2
             Else
               TTXX$ = Mid$(XX$, INICIO&)
               List1.AddItem TTXX$
               Exit Do
           End If
        Loop
    End If
    '
End Sub
'
Private Sub Text3_DBLCLICK()
    '
    Call Command4_Click
    '
End Sub

Private Sub Text5_Change()
    '
    CI% = CODINT%(Text5)
    If CI% > 0 Then
      Label5 = DESCRIT0$(CI%)
      Label3 = UNIMED$(CI%)
    Else
      Label5 = "": Label3 = ""
    End If
    '
End Sub

Private Sub Text5_DblClick()
    '
    Call Command8_Click
    '
End Sub

Sub CALPROCON(CADIAS%)
    '
    FEBASE = FECHANUM(Label6)
    If FEBASE < 1 Then
      Label9 = ""
      Exit Sub
    End If
    '
    FEPROCO = FEBASE
    While DIENTRE%(FEBASE, FEPROCO) < CADIAS%
      FEPROCO = DIPOST(FEPROCO)
    Wend
    '
    Label9.ForeColor = RGB(0, 0, 0)
    If FEPROCO <= HOY(HOS$) Then
       Label9.ForeColor = RGB(255, 0, 0)
    End If
    '
    Label9 = FECHATEX$(FEPROCO)
    '
End Sub

Sub DESPAMODE()
    '
    If DERACCE%("REMIMODE", "Remitir Modelo a Proveedor") >= 2 Then
       REMIMO09.Text6 = "Modelo: " + TRIM$(Text3) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Text2 + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
       REMIMO09.Text6 = REMIMO09.Text6 + "Cantidad de Piezas: " + TRIM$(Text8) + Chr$(13) + Chr$(10) + "Próximo Control: " + TRIM$(Label9) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
       REMIMO09.Text6 = REMIMO09.Text6 + TRIM$(Text6)
       REMIMO09.Show 1
       '
       Call CARGANOTA(Text3)
    End If
    '
End Sub

Sub RECIMODE()
    '
    If DERACCE%("RECIMODE", "Recibir Modelo de Proveedor") >= 2 Then
      RECEMO09.Label4 = "Modelo: " + TRIM$(Text3) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Text2
      RECEMO09.Show 1
      '
      Call CARGANOTA(Text3)
    End If
    '
End Sub

Sub AGREGA_ANOTAMODELO(NORXX$, TXTX$)
    '
    If NORXX$ <> "" Then
        ANOTANTE$ = VALOBADA("BADAMODE", "ANOTAMOD", "Cod_Mod = '" & NORXX$ & "'")
        ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
        Call ACTUREGISTRO("BADAMODE", "ANOTAMOD", "Cod_Mod = '" & NORXX$ & "'", ANOTA$, RAFE&)
    End If
    '
End Sub

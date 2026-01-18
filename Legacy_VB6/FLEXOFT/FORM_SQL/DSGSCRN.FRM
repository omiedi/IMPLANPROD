VERSION 5.00
Begin VB.Form DSGSCRN 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Formularios de Impresion"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   15315
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9920
      ScaleHeight     =   480
      ScaleWidth      =   5475
      TabIndex        =   56
      Top             =   0
      Width           =   5475
      Begin VB.CommandButton Command9 
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
         Left            =   1800
         Picture         =   "DSGSCRN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Muestra por Pantalla"
         Top             =   0
         Width           =   600
      End
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
         Height          =   480
         Left            =   2400
         Picture         =   "DSGSCRN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Refresca y Re-Dibuja Todo"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command3 
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
         Left            =   1200
         Picture         =   "DSGSCRN.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Genera Nuevo Formulario de Impresión"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   4800
         Picture         =   "DSGSCRN.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Termina - Cierra la Aplicación"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   4200
         Picture         =   "DSGSCRN.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   0
         Width           =   600
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
         Height          =   480
         Left            =   3600
         Picture         =   "DSGSCRN.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Guarda Cambios en Formulario Activo"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3000
         Picture         =   "DSGSCRN.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Imprime Muestra del Formulario Activo"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   9780
      Left            =   735
      ScaleHeight     =   9720
      ScaleWidth      =   10980
      TabIndex        =   42
      Top             =   1155
      Width           =   11040
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   100
         Left            =   0
         TabIndex        =   49
         Top             =   9450
         Width           =   10380
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00E0E0E0&
         Height          =   6105
         Left            =   0
         ScaleHeight     =   6045
         ScaleWidth      =   480
         TabIndex        =   51
         Top             =   0
         Width           =   540
         Begin VB.Line Line2 
            X1              =   510
            X2              =   510
            Y1              =   0
            Y2              =   25000
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6960
         LargeChange     =   100
         Left            =   10710
         TabIndex        =   43
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00E0E0E0&
         Height          =   540
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   11085
         TabIndex        =   50
         Top             =   0
         Width           =   11145
         Begin VB.Line Line1 
            X1              =   0
            X2              =   15000
            Y1              =   450
            Y2              =   450
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   6630
         Left            =   -30
         ScaleHeight     =   6570
         ScaleWidth      =   11085
         TabIndex        =   44
         Top             =   -30
         Width           =   11145
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   6960
            Left            =   525
            ScaleHeight     =   6900
            ScaleWidth      =   10320
            TabIndex        =   45
            Top             =   470
            Width           =   10380
            Begin VB.TextBox Text3 
               BackColor       =   &H00FFFFFF&
               Height          =   750
               Left            =   420
               TabIndex        =   47
               Top             =   210
               Visible         =   0   'False
               Width           =   2220
            End
            Begin VB.ListBox List1 
               Height          =   2205
               Left            =   420
               TabIndex        =   46
               Top             =   1470
               Visible         =   0   'False
               Width           =   3165
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000000C0&
               BorderStyle     =   5  'Dash-Dot-Dot
               Visible         =   0   'False
               X1              =   5145
               X2              =   5145
               Y1              =   0
               Y2              =   20000
            End
            Begin VB.Line Line3 
               BorderColor     =   &H000000FF&
               BorderStyle     =   5  'Dash-Dot-Dot
               Visible         =   0   'False
               X1              =   0
               X2              =   20000
               Y1              =   4200
               Y2              =   4200
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Label17"
               Height          =   195
               Left            =   525
               TabIndex        =   48
               Top             =   1050
               Visible         =   0   'False
               Width           =   570
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Frame1"
      Height          =   11040
      Left            =   11865
      TabIndex        =   8
      Top             =   -100
      Width           =   75
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   5685
      Left            =   11910
      ScaleHeight     =   5625
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   4095
      Width           =   3375
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1050
         TabIndex        =   65
         Text            =   " "
         Top             =   2520
         Visible         =   0   'False
         Width           =   1550
      End
      Begin VB.CommandButton Command13 
         Height          =   285
         Left            =   2400
         Picture         =   "DSGSCRN.frx":1526
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2310
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subrayado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   55
         Top             =   3650
         Width           =   1170
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Italica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   54
         Top             =   3400
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Negrita"
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
         Left            =   1785
         TabIndex        =   53
         Top             =   3150
         Width           =   1065
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         Left            =   1365
         Max             =   1
         Min             =   1
         TabIndex        =   52
         Top             =   405
         Value           =   1
         Width           =   150
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Delete"
         Height          =   250
         Left            =   1995
         TabIndex        =   34
         Top             =   405
         Width           =   960
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Text            =   " "
         Top             =   2310
         Width           =   2115
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   840
         TabIndex        =   31
         Top             =   1000
         Width           =   2115
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1155
         TabIndex        =   28
         Text            =   "8"
         Top             =   3400
         Width           =   540
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   2835
         Width           =   2115
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2190
         TabIndex        =   26
         Text            =   "0.0"
         Top             =   1815
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2190
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   1515
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Text            =   "0.0"
         Top             =   1815
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Text            =   "0.0"
         Top             =   1515
         Width           =   750
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Texto:"
         Height          =   225
         Left            =   -945
         TabIndex        =   33
         Top             =   2355
         Width           =   1695
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   30
         Top             =   705
         Width           =   2115
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   29
         Top             =   405
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño:"
         Height          =   225
         Left            =   -210
         TabIndex        =   22
         Top             =   3195
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fuente:"
         Height          =   225
         Left            =   -1050
         TabIndex        =   21
         Top             =   2865
         Width           =   1800
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   225
         Left            =   1050
         TabIndex        =   20
         Top             =   1845
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   225
         Left            =   1050
         TabIndex        =   19
         Top             =   1545
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   225
         Left            =   -945
         TabIndex        =   18
         Top             =   1845
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.- X1:"
         Height          =   225
         Left            =   -945
         TabIndex        =   17
         Top             =   1545
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   225
         Left            =   -945
         TabIndex        =   16
         Top             =   1035
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   225
         Left            =   -945
         TabIndex        =   15
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden:"
         Height          =   225
         Left            =   -945
         TabIndex        =   14
         Top             =   435
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Propiedades"
         Height          =   285
         Left            =   -10
         TabIndex        =   13
         Top             =   -30
         Width           =   16000
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   75
      Left            =   11865
      TabIndex        =   9
      Top             =   4935
      Width           =   15000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   5685
      Left            =   11910
      ScaleHeight     =   5625
      ScaleWidth      =   3315
      TabIndex        =   10
      Top             =   470
      Width           =   3375
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
         Height          =   450
         Left            =   1560
         Picture         =   "DSGSCRN.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Re-Imprimir Documentos Emitidos"
         Top             =   3600
         Width           =   645
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
         Height          =   450
         Left            =   735
         Picture         =   "DSGSCRN.frx":1CDA
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ver Documentos Emitidos"
         Top             =   4305
         Width           =   645
      End
      Begin VB.CommandButton Command99 
         Caption         =   "Recupera"
         Height          =   225
         Left            =   2625
         TabIndex        =   39
         Top             =   4515
         Visible         =   0   'False
         Width           =   645
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
         Height          =   450
         Left            =   1995
         Picture         =   "DSGSCRN.frx":1FE4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Descarga de Documentos Antiguos / Control y Recuperación de Datos"
         Top             =   4305
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2625
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   4305
         Visible         =   0   'False
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   -105
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command14 
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
         Left            =   105
         Picture         =   "DSGSCRN.frx":212E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Buscar Documento por Texto Embebido"
         Top             =   4305
         Width           =   645
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Configuracion General"
         Height          =   285
         Left            =   -10
         TabIndex        =   12
         Top             =   -30
         Width           =   16000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formulario Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8730
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
         Height          =   360
         Left            =   3465
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "DSGSCRN.frx":3494
         Top             =   370
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
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
         Left            =   3150
         TabIndex        =   3
         Top             =   370
         Width           =   170
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
         Height          =   360
         Left            =   1260
         TabIndex        =   2
         Text            =   " "
         Top             =   370
         Width           =   1905
      End
      Begin VB.Label Label2 
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
         Left            =   525
         TabIndex        =   1
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   4725
      Pattern         =   "*.FRM"
      TabIndex        =   5
      Top             =   6930
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   120
      Picture         =   "DSGSCRN.frx":3498
      Stretch         =   -1  'True
      ToolTipText     =   "Parametro"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   105
      Picture         =   "DSGSCRN.frx":35E2
      ToolTipText     =   "Cuadro de Imagen"
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   4
      Left            =   75
      Picture         =   "DSGSCRN.frx":38EC
      ToolTipText     =   "Caja"
      Top             =   3855
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   135
      Picture         =   "DSGSCRN.frx":4642
      ToolTipText     =   "Linea"
      Top             =   3345
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   2
      Left            =   85
      Picture         =   "DSGSCRN.frx":5184
      ToolTipText     =   "Parametro"
      Top             =   2200
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   1
      Left            =   150
      Picture         =   "DSGSCRN.frx":5E5E
      ToolTipText     =   "Leyenda Fija"
      Top             =   1700
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   105
      Picture         =   "DSGSCRN.frx":677C
      ToolTipText     =   "Diseño General"
      Top             =   1150
      Width           =   480
   End
   Begin VB.Label TEPRUORI 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label TEPRURED 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "DSGSCRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARRABOR%
Dim PORCOORDE%
'
Dim XMOUSE As Single, YMOUSE As Single
Dim ESTAPRETADO%, MUEVETODO%
Dim XORI, YORI
Dim MODO$
Dim UTAFUE
'
Dim XREGLA, YREGLA

Private Sub Check1_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     NEGRI(IOBJ%) = Check1.Value
   End If
End Sub

Private Sub Check2_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     ITALI(IOBJ%) = Check2.Value
   End If
End Sub

Private Sub Check3_Click()
   Call PREDITEXTO
   If IOBJ% <= CANOBJ% + 1 Then
     SUBRA(IOBJ%) = Check3.Value
   End If
End Sub

Private Sub Combo1_Change()
    Call PREDITEXTO
End Sub

Private Sub Combo1_Click()
    Call PREDITEXTO
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
        If EXISTE%(LUDAT$ + CODFOR$ + ".SQR") > 0 Then
             Call Text1_Change
          Else
             Call CARGAFORM
             Call REDIBUJA
        End If
    End If
    '
End Sub

Sub CARGAFORM()
    '
    CANOBJ% = 0
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    '
    ' CARGA LAS CAJAS
    MAXCA% = 0
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
            '
            CANOBJ% = CANOBJ% + 1
            TIOBJ$(CANOBJ%) = "BOX"
               MAXCA% = MAXCA% + 1
            ONAME$(CANOBJ%) = "BOX-" & MAXCA%
            TAFUEN(CANOBJ%) = 1
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                TAFUEN(CANOBJ%) = Asc(Mid$(RF$, 73, 1))
            End If
            '
            H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
            V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
            H2(CANOBJ%) = CVS(Mid$(RF$, 65, 4))
            V2(CANOBJ%) = CVS(Mid$(RF$, 69, 4))
        End If
    Next I%
    '
    ' dibuja lineas *************************************
    MAXLI% = 0
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
            '
            CANOBJ% = CANOBJ% + 1
            TIOBJ$(CANOBJ%) = "LINE"
               MAXLI% = MAXLI% + 1
            ONAME$(CANOBJ%) = "LINE-" & MAXLI%
            TAFUEN(CANOBJ%) = 1
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                TAFUEN(CANOBJ%) = Asc(Mid$(RF$, 73, 1))
            End If
            '
            H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
            V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
            H2(CANOBJ%) = CVS(Mid$(RF$, 65, 4))
            V2(CANOBJ%) = CVS(Mid$(RF$, 69, 4))
        End If
    Next I%
    '
    ' escribe leyendas fijas ***************************
    MAXLA% = 0
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
           If TRIM$(Mid$(RF$, 9, 48)) <> "" Then
               FUENX$ = TRIM$(Mid$(RF$, 87, 32))
               If FUENX$ = "" Then FUENX$ = "Courier"
               TAFU = CVS(Mid$(RF$, 119, 4))
               '
               CANOBJ% = CANOBJ% + 1
               TIOBJ$(CANOBJ%) = "LABEL"
                 MAXLA% = MAXLA% + 1
               ONAME$(CANOBJ%) = "LABEL-" & MAXLA%
               FUENTE$(CANOBJ%) = FUENX$
               TAFUEN(CANOBJ%) = TAFU
               '
               H1(CANOBJ%) = CVS(Mid$(RF$, 57, 4))
               V1(CANOBJ%) = CVS(Mid$(RF$, 61, 4))
               TEXTO(CANOBJ%) = TRIM$(Mid$(RF$, 9, 48))
           End If
        End If
109 Next I%
    '
End Sub
'
Private Sub Command10_Click()
   If Text3.Visible Then Exit Sub
   Command10.Enabled = False
   Call REDIBUJA
   Command10.Enabled = True
End Sub

Sub REDIBUJA()
   Text3.Visible = False
   Text3 = ""
   Picture1.Cls
   For I% = 1 To CANOBJ%
      Call DIBUOBJ(I%)
   Next I%
End Sub
'
Sub DIBUOBJ(I%)
   '
   MODOX$ = TIOBJ$(I%)
   On Error Resume Next
   Picture1.ForeColor = RGB(0, 0, 0)
   Select Case MODOX$
      Case "LABEL"
         Picture1.CurrentY = 56.7 * V1(I%)
         Picture1.CurrentX = 56.7 * H1(I%)
         Picture1.Font = FUENTE$(I%)
         Picture1.FontSize = TAFUEN(I%)
         Picture1.FontBold = NEGRI(I%)
         Picture1.FontItalic = ITALI(I%)
         Picture1.FontUnderline = SUBRA(I%)
         Picture1.Print TEXTO(I%)
         '
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         Label17 = TEXTO(I%) & "*M"
         H2(I%) = H1(I%) + Label17.Width / 56.7
         V2(I%) = V1(I%) + Label17.Height / 56.7
      Case "PARAM"
         Picture1.CurrentY = 56.7 * V1(I%)
         Picture1.CurrentX = 56.7 * H1(I%)
         Picture1.Font = FUENTE$(I%)
         Picture1.FontSize = TAFUEN(I%)
         Picture1.FontBold = NEGRI(I%)
         Picture1.FontItalic = ITALI(I%)
         Picture1.FontUnderline = SUBRA(I%)
         Picture1.ForeColor = RGB(255, 0, 0)
         Picture1.Print TEXTO(I%)
         Picture1.ForeColor = RGB(0, 0, 0)
         '
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         Label17 = TEXTO(I%) + "*M"
         H2(I%) = H1(I%) + Label17.Width / 56.7
         V2(I%) = V1(I%) + Label17.Height / 56.7
      Case "TEXT"
         Picture1.CurrentY = 56.7 * V1(I%)
         Picture1.CurrentX = 56.7 * H1(I%)
         Picture1.Font = FUENTE$(I%)
         Picture1.FontSize = TAFUEN(I%)
         Picture1.FontBold = NEGRI(I%)
         Picture1.FontItalic = ITALI(I%)
         Picture1.FontUnderline = SUBRA(I%)
         Picture1.ForeColor = RGB(0, 255, 0)
         Picture1.Print TEXTO(I%)
         Picture1.ForeColor = RGB(0, 0, 0)
         '
         Label17.Font = Picture1.Font
         Label17.FontSize = Picture1.FontSize
         Label17.FontBold = Picture1.FontBold
         Label17.FontItalic = Picture1.FontItalic
         Label17 = TEXTO(I%) + "*M"
         H2(I%) = H1(I%) + Label17.Width / 56.7
         V2(I%) = V1(I%) + Label17.Height / 56.7
      Case "LINE"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = XVALO(TAFUEN$(I%))
         Picture1.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%))
      Case "BOX"
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = XVALO(TAFUEN$(I%))
         Picture1.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%)), , B
   End Select
   Call FRESHALL
   On Error GoTo 0
   '
End Sub
'
Private Sub Command11_Click()
   '
   Command11.Enabled = False
   OPPX% = ALTERNA%("Descarga Documentos Viejos\Control y Recuperación de Datos\Control de Documentos")
   If OPPX% = 1 Then
        Call DESCADOC
      ElseIf OPPX% = 2 Then
        Call CHECKDOC
      ElseIf OPPX% = 3 Then
        Call VERIDOC
   End If
   '
99 Screen.MousePointer = 1
   Command11.Enabled = True
   MARCOBARRA.Visible = False
   End Sub

Sub DESCADOC()
   '
   If DERACCE%("DESCADOC", "Descarga de Base de Datos de Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Descargar la Base de Datos\de Documentos Impresos ?\\No, Conservar\Si, Descargar.") = 2 Then
       MARCOBARRA.Visible = True
       Call BLOQARCH("PDOCSPL")
       On Error Resume Next
       Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Proceso Imposible.\   \Cierre FLEXOFT en TODAS las\Estaciones de Trabajo, y\Repita la Petición.")
         Call CIERRARCH("MENSAJES")
         GoTo 99
       End If
       On Error GoTo 0
       '
       Screen.MousePointer = 11
       Call COPYSTRU("PDOCSPL", "PDOCSPL0")
       '
       FIN& = ULTREG&("PDOCLST")
       For U& = FIN& To 1 Step -1
         XX$ = REGLEIDO$("PDOCLST", U&)
         UDOC& = CVS(Left$(XX$, 4))
         If UDOC& > 0 Then GoTo 20
       Next U&
       UDOC& = 1
       '
20     FINDESCA& = UDOC& / 2   ' descarga la mitad del archivo
       Dim TD0 As String * 512
       NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
       URE0& = LOF(NX0%) / 512
       FIN& = ULTREG&("PDOCSPL0")
       NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
       URE1& = 0
       '
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("PDOCSPL0", U&)
         TD0$ = XX$
         If CVS(Left$(XX$, 4)) <= FINDESCA& Then
              URE0& = URE0& + 1
              Put #NX0%, URE0&, TD0$
            Else
              URE1& = URE1& + 1
              Put #NX1%, URE1&, TD0$
         End If
       Next U&
       '
       Call CIERRARCH("PDOCSPL0")
       Close #NX0%
       Close #NX1%
       '
       Screen.MousePointer = 1
       On Error Resume Next
       Kill LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
           On Error GoTo 0
           Call COMUNI("No se Pudo Completar el Proceso - Consulte.")
         Else
           '\\\OT-07-0111-RG-15-03-07/// 'MENSAJE QUE MUESTRA CUANDO FINALIZA EL PROCESO
           On Error GoTo 0
           Call COMUNI("Fin de Proceso de Descarga:\   \Operación Exitosa\")
           '\\\OT-07-0111-RG-FIN///
       End If
       On Error GoTo 0
       '
     End If
   End If
   '
99 End Sub

Sub CHECKDOC()
   '
   If DERACCE%("CHECKDOC", "Control y Recuperación de B.D. Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Recuperar la Base de Datos\de Documentos Impresos ?\\No, Cancelar\Si, Recuperar.") = 2 Then
       MARCOBARRA.Visible = True
       OPX% = COMALTER%("Para este Proceso debe Cerrarse FLEXOFT\en TODAS las Estaciones de Trabajo.\   \Pulse 'Continuar' para Iniciar el Proceso\\Continuar\Cancelar")
       If OPX% <> 1 Then Exit Sub
       '
       EPAC$ = UCase$(EMPREAC$)
       If InStr(EPAC$, "DOSIVAC") > 0 Then
          Call CHECKDOSI
          GoTo 99
       End If
       '
       TTXX$ = "Para el Proceso de Recuperación,\Tome Archivo de Backup 'PDOCSPL.DAT'\y Re-Nómbrelo como 'BKSPOOL.DAT'.\   \Luego Coloque este Último en la \Carpeta de Datos '" + REPLACAR$(LUDAT$, "\", "/")
       If COMALTER%(TTXX$ + "\\Cancelar\Recuperar") = 2 Then
          Screen.MousePointer = 11
          Call COPYSTRU("PDOCSPL", "BKSPOOL")
          FIN& = ULTREG&("BKSPOOL")
          For U& = 2100000 To FIN&
             Call TRACE(20, U&, FIN&)
             X$ = REGLEIDO$("PDOCSPL", U&)
             Y$ = REGLEIDO$("BKSPOOL", U&)
             If Y$ <> X$ Then
                If COMALTER%("Recupera Registro Numero " & U& & "\\No, Continuar\Si,Recuperar") = 2 Then
                   Call GRAREG("PDOCSPL", Y$, U&)
                End If
             End If
          Next U&
          Call CIERRARCH("*.*")
          Screen.MousePointer = 1
          Call COMUNI("Fin de Recuperación")
       End If
       '
     End If
   End If
   '
99 List2.Visible = False
   Picture3.Visible = False
   Call LIBARCH("*.*")
   '
End Sub

Sub VERIDOC()
    Dim ISDOC%(), DESRE&(), HASRE&()
    '
    MARCOBARRA.Visible = True
    FIN& = ULTREG&("PDOCLST")
    For U& = FIN& To 1 Step -1
       XX$ = REGLEIDO$("PDOCLST", U&)
       UDOC& = CVS(Left$(XX$, 4))
       If UDOC& > 0 Then GoTo 20
    Next U&
    UDOC& = 1
    '
20  UDOC& = UDOC& + 256
    ReDim ISDOC%(UDOC&), DESRE&(UDOC&), HASRE&(UDOC&)
    '
    PRIDOC& = 0: ULTDOC& = 0
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCLST", U&)
       NDOC& = CVS(Left$(XX$, 4))
' VERIFICAR AQUI ORDEN CORRELATIVO
       If NDOC& > 0 Then
         If PRIDOC& = 0 Then PRIDOC& = NDOC&
         If NDOC& <= UDOC& Then
           ISDOC%(NDOC&) = 1
           If NDOC& > ULTDOC& Then ULTDOC& = NDOC&
           If NDOC& <= DOCANT& Then
              YY$ = TRIM$(Mid$(XX$, 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Desorden en Indice."
           End If
           DOCANT& = NDOC&
           GoTo 29
         End If
       End If
29   Next U&
     '
     'NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
     'URE0& = LOF(NX0%) / 512
     'FIN& = ULTREG&("PDOCSPL0")
     'NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
     'URE1& = 0
     '
     FIN& = ULTREG&("PDOCSPL")
     DOCANT& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCSPL", U&)
       NDOC& = CVS(Left$(XX$, 4))
       If NDOC& < DOCANT& Then
           YY$ = TRIM$(Mid$(XX$, 5, 60))
           MsgBox "Registro " & U& & " de PDOCSPL - Desorden en Detalle. (" & DOCANT& & " / " & NDOC& & ")."
       End If
       DOCANT& = NDOC&
       If NDOC& > 0 Then
         If NDOC& <= UDOC& Then
           If DESRE&(NDOC&) < 1 Then DESRE&(NDOC&) = U&
           If HASRE&(NDOC&) < U& Then HASRE&(NDOC&) = U&
         End If
       End If
     Next U&
     '
     For U& = 1 To UDOC&
        If ISDOC%(U&) = 1 Then
           If DESRE&(U&) < 1 Or HASRE&(U&) < DESRE&(U&) Then
              YY$ = TRIM$(Mid$(REGLEIDO$("PDOCLST", U&), 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Falta Detalle "
           End If
           If DESRE&(U&) > 0 Then
             If DESRE&(U&) <= HASRE&(U& - 1) Then
               MsgBox "Punteros Cruzados"
             End If
           End If
        End If
     Next U&
     '
End Sub

Sub CHECKDOSI()
   '
   Call BLOQARCH("PDOCSPL")
   FIDATI# = FIDATIM#(LUDAT$ + "PDOCSPL.DAT")
   '
   Screen.MousePointer = 11
   Call COPYSTRU("PDOCSPL", "PDOCSPL0")
   On Error Resume Next
   Kill LUDAT$ + "PDOCSPL0.DAT"
   On Error GoTo 0
   '
   If EXISTE%(LUDAT$ + "PDOCSPL0.DAT") > 0 Then
      Call MENSERR(24, "Proceso Imposible. No fue Posible Borrar\Archivo de Transferencia 'PDOCSPL0.DAT'")
      GoTo 99
   End If
   '
   List2.Visible = True
   Picture3.Visible = True
   '
   YY$ = REGLEIDO$("PDOCLST", 1)
   MINI = CVS(Left$(YY$, 4))
   URELI& = ULTREG&("PDOCLST")
   YY$ = REGLEIDO$("PDOCLST", URELI&)
   MAXI = CVS(Left$(YY$, 4))
   '
   FIN& = ULTREG&("PDOCSPL")
   JJ& = 0: NDOCA& = 0: CAREDEL& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      VALOPRO& = 1 + (U& / FIN&) * (MAXI - MINI + 1)
      XX$ = REGLEIDO$("PDOCSPL", U&)
      VALEIDO = CVS(Left$(XX$, 4))
      If (Abs(VALOPRO& - VALEIDO)) < 10 Or ((Abs(VALOPRO& - VALEIDO)) / VALOPRO&) < 0.1 Then
         JJ& = JJ& + 1
         Call GRAREG("PDOCSPL0", XX$, JJ&)
         NDOC1& = VALEIDO
         If NDOC1& <> NDOCA& Then
           YY$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(NDOC1&))
           ZZ$ = FORMATNUM$(NDOC1&, "F7.0") + "  " + Mid$(YY$, 5, 60)
           If List2.ListCount > 30000 Then
              List2.RemoveItem 0
           End If
           List2.AddItem ZZ$
           List2.Visible = True
           List2.ListIndex = List2.ListCount - 1
           List2.Refresh
           NDOCA& = NDOC1&
         End If
         CAREDEL& = 0
         GoTo 29
      End If
      If CAREDEL& = 0 Then
        List2.AddItem "   **********************  Documentos Perdidos  *************************"
      End If
      CAREDEL& = CAREDEL& + 1
      '
29  Next U&
   '
   Call CIERRARCH("PDOCSPL")
   Call CIERRARCH("PDOCSPL0")
   '
   If ULTREG&("PDOCSPL") <> FIN& Then
      Call MENSERR(24, "Imposible Continuar. Usuarios Conectados\Activos Durante el Proceso.")
      GoTo 99
   End If
   '
   Close
   Reset
   Call BLOQARCH("PDOCSPL")
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL.BAD"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL.DAT' como 'PDOCSPL.BAD'.\   \Repita el Proceso.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL0.DAT" As LUDAT$ + "PDOCSPL.DAT"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL0.DAT' como 'PDOCSPL.DAT'.\   \Consulte Urgente a Soporte FLEXOFT.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   Call COMUNI("Fin de Proceso de Recuperación:\   \Operación Exitosa\")
99 On Error GoTo 0
   '
End Sub
'
Private Sub Command13_Click()
    DEVALOBA.Label14 = Label14
    DEVALOBA.Show 1
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
10  TEBUS$ = InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", TEBUS$)
    If TEBUS$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    MARCOBARRA.Visible = True
    Call COPYSTRU("PDOCLST", "PDOCLIS")
    URE& = ULTREG&("PDOCSPL")
    For I& = 1 To ULTREG&("PDOCSPL")
      Call TRACE(20, I&, URE&)
      RELE$ = REGLEIDO$("PDOCSPL", I&)
      X$ = Mid$(RELE$, 5)
      
      If InStr(1, X$, TEBUS$) > 0 Then
         DOCUNU& = CVS(Left$(RELE$, 4))
         '
         RECORD$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
         '
         Z& = Z& + 1
         Call GRAREG("PDOCLIS", RECORD$, Z&)
         '
         Call CIERRARCH("PDOCLIS")
      End If
    Next I&
    Call SETULTREG("PDOCLIS", Z&)
    MARCOBARRA.Visible = False
    '
    If Z& < 1 Then
       Call MENSERR(24, "No se Encontraron Documentos\para el Argumento de Búsqueda.")
       GoTo 10
    End If
    '
25  Screen.MousePointer = 1
    Call SELECHO("PDOCLIS/Búsqueda en Documentos por Texto: '" + TEBUS$ + "'")
    DOCUNU& = Val(VALACT1$("PDOCLIS"))
    If DOCUNU& > 0 Then
 '      Call MUESTRA_DOCU(DOCUNU&)
       GoTo 25
    End If
    GoTo 10
    '
99  Command14.Enabled = True
    '
End Sub

Private Sub Command15_Click()
   IOBJ% = XVALO(Label14)
   If IOBJ% <= CANOBJ Then
      For KKI% = IOBJ% + 1 To CANOBJ%
         TIOBJ$(KKI% - 1) = TIOBJ$(KKI%)
         ONAME$(KKI% - 1) = ONAME$(KKI%)
         H1(KKI% - 1) = H1(KKI%)
         H2(KKI% - 1) = H2(KKI%)
         V1(KKI% - 1) = V1(KKI%)
         V2(KKI% - 1) = V2(KKI%)
         TEXTO(KKI% - 1) = TEXTO(KKI%)
         FUENTE$(KKI% - 1) = FUENTE$(KKI%)
         TAFUEN$(KKI% - 1) = TAFUEN$(KKI%)
         NEGRI(KKI% - 1) = NEGRI(KKI%)
         ITALI(KKI% - 1) = ITALI(KKI%)
         SUBRA(KKI% - 1) = SUBRA(KKI%)
         FORPA$(KKI% - 1) = FORPA$(KKI%)
         FOREPLA$(KKI% - 1) = FOREPLA$(KKI%)
         TASEL$(KKI% - 1) = TASEL$(KKI%)
         CASEL$(KKI% - 1) = CASEL$(KKI%)
         CONDSEL$(KKI% - 1) = CONDSEL$(KKI%)
      Next KKI%
      CANOBJ% = CANOBJ% - 1
      Label14 = CANOBJ%
      Call REDIBUJA
   End If
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Close
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Call MENSERR(24, "No Disponible por Ahora")
    ' Call ALTANUEFORM
End Sub

Private Sub Command4_Click()
   '
   For I% = 1 To CANOBJ%
      Call PRINOBJ(I%)
   Next I%
   Printer.EndDoc
   '
End Sub

Sub PRINOBJ(I%)
   '
   MODOX$ = TIOBJ$(I%)
   On Error Resume Next
   Select Case MODOX$
      Case "LABEL"
         Printer.CurrentY = 56.7 * V1(I%)
         Printer.CurrentX = 56.7 * H1(I%)
         Printer.Font = FUENTE$(I%)
         Printer.FontSize = TAFUEN(I%)
         Printer.FontBold = NEGRI(I%)
         Printer.FontItalic = ITALI(I%)
         Printer.FontUnderline = SUBRA(I%)
         Printer.Print TEXTO(I%)
      Case "PARAM"
         Printer.CurrentY = 56.7 * V1(I%)
         Printer.CurrentX = 56.7 * H1(I%)
         Printer.Font = FUENTE$(I%)
         Printer.FontSize = TAFUEN(I%)
         Printer.FontBold = NEGRI(I%)
         Printer.FontItalic = ITALI(I%)
         Printer.FontUnderline = SUBRA(I%)
         Printer.Print FORMATNUM$(TEXTO(I%), FORPA$(I%))
      Case "LINE"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = 3 * XVALO(TAFUEN$(I%))
         Printer.DrawStyle = 0
         Printer.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%)), QBColor(0)
      Case "BOX"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = 3 * XVALO(TAFUEN$(I%))
         Printer.DrawStyle = 0
         Printer.FillStyle = 1
         Printer.Line (56.7 * H1(I%), 56.7 * V1(I%))-(56.7 * H2(I%), 56.7 * V2(I%)), QBColor(0), B
   End Select
   Call FRESHALL
   On Error GoTo 0
   '
End Sub

Private Sub Command5_Click()
    '
    CODFOR$ = TRIM$(Text1)
    If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       Exit Sub
    End If
    '
    ' VERIFICA REPETICION DE NOMBRES
    For KKI% = 1 To CANOBJ%
       If TRIM$(ONAME$(KKI%)) = "" Then
          Call MENSERR(24, "Falta Nombre de Objeto")
          Label14 = KKI%
          Text9.SetFocus
          Exit Sub
       End If
       '
       For KKJ% = 1 To KKI% - 1
          If ONAME$(KKI%) = ONAME$(KKJ%) Then
             Call MENSERR(24, "Nombre de Objeto Repetido")
             Label14 = KKI%
             Text9.SetFocus
             Exit Sub
          End If
       Next KKJ%
    Next KKI%
    '
    N1% = FILEOPEN%(LUDAT$ + CODFOR$ + ".SQR", 2, 128)
    For I% = 1 To CANOBJ%
      TTXX$ = "TIPO = " + TIOBJ$(I%)
      TTXX$ = TTXX$ + Chr$(124) + "NAME = " + ONAME$(I%)
      TTXX$ = TTXX$ + Chr$(124) + " H1 = " + FORMATNUM$(H1(I%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " V1 = " + FORMATNUM$(V1(I%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " H2 = " + FORMATNUM$(H2(I%), "F10.1")
      TTXX$ = TTXX$ + Chr$(124) + " V2 = " + FORMATNUM$(V2(I%), "F10.1")
        '
        TX123$ = TEXTO(I%)
      If TRIM$(TX123$) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TEXT = " + TX123$
      If TRIM$(FUENTE$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FONT = " + FUENTE$(I%)
      If TRIM$(TAFUEN$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " SIZE = " + TAFUEN$(I%)
      If TRIM$(FORPA$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FORM = " + FORPA$(I%)
      '
      If NEGRI(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " NEGRI = 1"
      If ITALI(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " ITALI = 1"
      If SUBRA(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " SUBRA = 1"
      '
      If TRIM$(FOREPLA$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FOREPLA = " + FOREPLA$(I%)
      If TRIM$(TASEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TASEL = " + TASEL$(I%)
      If TRIM$(CASEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CASEL = " + CASEL$(I%)
      If TRIM$(CONDSEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CONDSEL = " + CONDSEL$(I%)
      Print #N1%, TTXX$
    Next I%
    '
    Close #N1%
    '
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("PRIFORMS")
End Sub

Private Sub Command7_Click()
   '
   On Error Resume Next
   Printer.DrawWidth = 3
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   If Err Then
     Call MENSERR(24, "Imposible Imprimir.\Error de Comunicacion con la Impresora.")
     Resume
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   '
   TPSP$ = ""
   Screen.MousePointer = 11
   '
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  For KKL& = L& To URE&          'SE CAMBIO EL ARCHIVO DE LECTURA
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       ElseIf CVS(Left$(TBUF$, 4)) > DOCUNU& Then
         GoTo 91
     End If
   Next KKL&
   '
91 Close #NX0%
   Screen.MousePointer = 11
   On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   FORMALTER.List1.Clear
   On Error GoTo 0
   '
   Printer.EndDoc        ' RESETEA IMPRESORA
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
    '    Call PRINTCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Printer.EndDoc
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub
'
Private Sub Command8_Click()
   '
   For KKI% = 0 To 5
     EDIFORMU.Picture1(KKI%).Cls
     EDIFORMU.Picture1(KKI%).Top = 0
     EDIFORMU.Picture1(KKI%).Refresh
     EDIFORMU.Picture1(KKI%).Height = 7070
     EDIFORMU.Picture1(KKI%).Width = 11700
     EDIFORMU.Picture1(KKI%).DrawWidth = 3
     EDIFORMU.Picture1(KKI%).DrawStyle = 0
     EDIFORMU.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   '
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
   '     Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   '
   EDIFORMU.Show
   '
End Sub

Private Sub Command99_Click()
   Screen.MousePointer = 11
   FIN& = ULTREG&("PDOCLST")
   For U& = 73740 To FIN&
     XX$ = REGLEIDO$("PDOCLST", U&)
     Call REPLA(XX$, MKS$(U&), 1, 4)
     If U& >= 73740 Then
       If U& <= 73742 Then
         Call REPLA(XX$, "Documento Numero " + TRIM$(Str$(U&)), 5, 124)
       End If
     End If
     Call GRAREG("PDOCLST", XX$, U&)
   Next U&
   Call CIERRARCH("PDOCLST")
   Call FRESHECHO("PDOCLST")
   '
   FIN& = ULTREG&("PDOCSPL")
   TTXX$ = ""
   For U& = 1470000 To FIN&
     XX$ = REGLEIDO$("PDOCSPL", U&)
     IU& = CVS(Left$(XX$, 4))
     If IU& < 10000 Then
       IU& = IU& + 73742
       Call REPLA(XX$, MKS$(IU&), 1, 4)
       Call GRAREG("PDOCSPL", XX$, U&)
     End If
   Next U&
   Call CIERRARCH("PDOCSPL")
   Screen.MousePointer = 1
   Call COMUNI("Recuperacion Completada")
End Sub

'
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Show
    Call ACOMODATODO
    Refresh
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    HOII% = HOY(HOS$)
    If HOII% = FECHANUM("21/11/05") Then
      Command99.Visible = True
    End If
    '
    Screen.MousePointer = 11
    On Error Resume Next
    Combo1.Clear
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        Combo1.AddItem Printer.FontName
    Next I%
    On Error GoTo 0
    Combo1.TEXT = "Courier New"
    '
    Call DIBUREGHOR
    Call DIBUREGVER
    
    Call CARINDIFORM
    '
    CODFOR$ = APLINVO$
    Text1.TEXT = CODFOR$
    Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
    '
    If TRIM$(Text1) <> "" Then
      Call SHOWFORM
         'Else
         'Screen.MousePointer = 1
         'Call Command1_Click
    End If
    Screen.MousePointer = 1
    '
End Sub

Sub DIBUREGHOR()
    'DIBUJA REGLA HORIZONTAL
    Picture6.Line (Picture1.Left, 200)-(15000, 200)
    For IX = 0 To 200 Step 5
       INI = 150
       Picture6.CurrentX = Picture1.Left + IX * 56.7 - 90
       Picture6.CurrentY = 250
       Picture6.Font = TAHOMA
       Picture6.FontSize = 7
       IXXX$ = TRIM$(Str$(IX)): While Len(IXXX$) < 3: IXXX$ = " " + IXXX$ + " ": Wend
       If IX Mod 10 = 0 Then
          INI = 110
          If IX >= 0 Then Picture6.Print TRIM$(Str$(IX))
       End If
       If IX Mod 50 = 0 Then
          INI = 50
       End If
       Picture6.Line (Picture1.Left + IX * 56.7, INI)-(Picture1.Left + IX * 56.7, 200)
    Next IX
    '
End Sub
'
Sub DIBUREGVER()
    'DIBUJA REGLA VERTICAL
    Picture7.Line (200, Picture1.Top)-(200, 25000)
    For IX = 0 To 300 Step 5
       INI = 150
       Picture7.CurrentX = 250
       Picture7.CurrentY = Picture1.Top + IX * 56.7 - 60
       Picture7.Font = TAHOMA
       Picture7.FontSize = 7
       IXXX$ = TRIM$(Str$(IX)): While Len(IXXX$) < 3: IXXX$ = " " + IXXX$ + " ": Wend
       If IX Mod 10 = 0 Then
          INI = 110
          If IX >= 0 Then Picture7.Print TRIM$(Str$(IX))
       End If
       If IX Mod 50 = 0 Then
          INI = 50
       End If
       Picture7.Line (INI, Picture1.Top + IX * 56.7)-(200, Picture1.Top + IX * 56.7)
    Next IX
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame3.MousePointer = 9
   If ARRABOR% = 1 Then
      If Frame3.Left + X > 20 Then
          Frame3.Left = Frame3.Left + X
          Call ACOMODATODO
          Call DIBUREGHOR
      End If
   End If
End Sub

Private Sub frame3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Sub ACOMODATODO()
    '
    On Error Resume Next
    Picture2.Left = Frame3.Left + Frame3.Width - 30
    Picture2.Width = Width - Picture2.Left
    Picture3.Left = Frame3.Left + Frame3.Width - 30
    Picture3.Width = Width - Picture3.Left
    Frame5.Left = Frame3.Left + Frame3.Width - 30
    Frame5.Width = Width - Frame5.Left
    '
    Picture2.Height = Frame5.Top - Picture2.Top
    Picture3.Top = Frame5.Top + Frame5.Height - 50
    Picture3.Height = ScaleHeight - Picture3.Top
    '
    Picture5.Width = Frame3.Left - Picture5.Left - 100
    VScroll1.Top = -20
    VScroll1.Left = Picture5.Width - VScroll1.Width - 50
    VScroll1.Height = Picture5.Height - HScroll1.Height - 550
    '
    Picture5.Height = ScaleHeight - Picture5.Top - Top - 100
    HScroll1.Left = -20
    HScroll1.Top = Picture5.Height - HScroll1.Height - 50
    HScroll1.Width = Picture5.Width - VScroll1.Width
    '
    Picture1.Width = Printer.ScaleWidth
    Picture1.Height = Printer.ScaleHeight
    Picture4.Width = Picture1.Left + Picture1.Width + VScroll1.Width + 150
    Picture4.Height = Picture1.Top + Picture1.Height
    '
    Picture6.Top = -30
    Picture6.Height = Picture1.Top + 60
    Picture6.Left = Picture4.Left
    Picture6.Width = Picture4.Width
    '
    Picture7.Left = -30
    Picture7.Width = Picture1.Left + 60
    Picture7.Top = Picture4.Top
    Picture7.Height = Picture4.Height
    '
    VScroll1.Value = 0
    HScroll1.Value = 0
    DIFHOR = Picture4.Width - Picture5.Width: If DIFHOR < 0 Then DIFHOR = 0
    HScroll1.Max = DIFHOR
    DIFVER = Picture4.Height - Picture5.Height: If DIFVER < 0 Then DIFVER = 0
    VScroll1.Max = DIFVER
    On Error GoTo 0
    '
End Sub
'
Private Sub frame5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame5.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame5.Top = Frame5.Top + Y
      Call ACOMODATODO
   End If
End Sub

Private Sub frame5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Private Sub HScroll1_Scroll()
   Picture4.Left = -30 - HScroll1.Value
   Picture6.Left = Picture4.Left
   Call DIBUREGHOR
End Sub

Private Sub Image1_Click(Index As Integer)
    If Index > 0 Then
       IOBJ% = CANOBJ% + 1
       Label14 = IOBJ%
       Select Case Index
          Case 1
            MODO$ = "LABEL": Label15 = "LABEL (Leyenda Fija)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text11.Visible = False
            Label16 = "Texto:"
            Command13.Visible = False
            Text8 = UTAFUE
          Case 2
            MODO$ = "PARAM": Label15 = "PARAM (Parametro)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text8 = UTAFUE
            Text11.Top = Text10.Top
            Text11.Left = Text10.Left
            Text11.Visible = True
            Text11 = "F10.2"
            Label16 = "Formato:"
            Command13.Visible = True
            If FORPA$(IOBJ%) <> "" Then Text11.TEXT = FORPA$(IOBJ%)
          Case 3
            MODO$ = "LINE": Label15 = "LINE (Linea)"
            ALIN = XVALO(Text8)
            If ALIN < 1 Or ALIN > 3 Then ALIN = 1
            Text8 = CStr(ALIN)                    ' espesor
          Case 4
            MODO$ = "BOX": Label15 = "BOX (Rectangulo)"
            ALIN = XVALO(Text8)
            If ALIN < 1 Or ALIN > 3 Then ALIN = 1
            Text8 = CStr(ALIN)                    ' espesor
          Case 6
            MODO$ = "TEXT": Label15 = "TEXT (Caja de Texto)"
            If Combo1.TEXT = "" Then Combo1.TEXT = "Courier New"
            If UTAFUE = "" Then UTAFUE = "8"
            Text8 = UTAFUE
            Text11.Top = Text10.Top
            Text11.Left = Text10.Left
            Text11.Visible = True
            Text11 = "F10.2"
            Label16 = "Formato:"
            Command13.Visible = True
            If FORPA$(IOBJ%) <> "" Then Text11.TEXT = FORPA$(IOBJ%)
       End Select
       '
       MAXCAI% = 1
25     ONAM$ = MODO$ + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ%
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
       Text9 = ONAM$
       Text4 = "0.0"
       Text5 = "0.0"
       Text6 = "0.0"
       Text7 = "0.0"
    End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Left$(Label15, 4) = "LINE" Then
      XORI = 56.7 * XVALO(Text4)
      YORI = 56.7 * XVALO(Text6)
      XMOUSE = 56.7 * XVALO(Text5)
      YMOUSE = 56.7 * XVALO(Text7)
      ESTAPRETADO% = 1
      Call Picture1_MouseMove(1, 0, XMOUSE, YMOUSE)
      Picture1.SetFocus
'      Call Picture1_MouseMove(1, 0, XMOUSE, YMOUSE)
   End If
      
End Sub

Private Sub Label14_Change()
    If VScroll2.Max < CANOBJ% Then
       If VScroll2.Value < CANOBJ% Then
          VScroll2.Max = CANOBJ%
       End If
    End If
    '
    IOBJ% = Label14
    Label15 = TIOBJ$(IOBJ%)
    MODO$ = TIOBJ$(IOBJ%)
    '
    Text4 = FORMATNUM$(H1(IOBJ%), "F7.1")
    Text5 = FORMATNUM$(H2(IOBJ%), "F7.1")
    Text6 = FORMATNUM$(V1(IOBJ%), "F7.1")
    Text7 = FORMATNUM$(V2(IOBJ%), "F7.1")
    '
    If FUENTE$(IOBJ%) <> "" Then Combo1 = FUENTE$(IOBJ%)
    If XVALO(TAFUEN$(IOBJ)) >= 1 Then Text8 = TAFUEN$(IOBJ%)
    Check1.Value = NEGRI(IOBJ%)
    Check2.Value = ITALI(IOBJ%)
    Check3.Value = SUBRA(IOBJ%)
    '
    Text9 = ONAME$(IOBJ%)
    Text10 = TEXTO(IOBJ%)
    '
    Command13.Visible = False
    Text11.Visible = False
    Label16 = "Texto:"
    If MODO$ = "PARAM" Then
       Label16 = "Formato:"
       Command13.Visible = True
       Text11.Top = Text10.Top
       Text11.Left = Text10.Left
       Text11.Visible = True
       Text11.TEXT = FORPA$(IOBJ%)
    End If
    '
End Sub


Private Sub Picture1_DblClick()
  '
  OBUS$ = "LABEL"
  '
2 For I% = CANOBJ% To 1 Step -1
    If TIOBJ$(I%) = OBUS$ Then
      If XMOUSE >= 56.7 * H1(I%) Then
        If XMOUSE <= 56.7 * H2(I%) Then
          If (YMOUSE >= 56.7 * V1(I%) And YMOUSE <= 56.7 * V2(I%)) Or (YMOUSE >= 56.7 * V2(I%) And YMOUSE <= 56.7 * V1(I%)) Then
               GoTo 6
          End If
        End If
      End If
    End If
  Next I%
  '
  If OBUS$ = "LABEL" Then OBUS$ = "PARAM": GoTo 2
  If OBUS$ = "PARAM" Then OBUS$ = "LINE": GoTo 2
  If OBUS$ = "LINE" Then OBUS$ = "BOX": GoTo 2
  Exit Sub
  '
6 ' PRESENTA DATOS DEL OBJETO
  IOBJ% = I%
  Label14 = IOBJ%
  Label15 = TIOBJ$(I%)
  MODO$ = TIOBJ$(I%)
  '
  Text4 = FORMATNUM$(H1(I%), "F7.1")
  Text5 = FORMATNUM$(H2(I%), "F7.1")
  Text6 = FORMATNUM$(V1(I%), "F7.1")
  Text7 = FORMATNUM$(V2(I%), "F7.1")
  '
  If TIOBJ$(I%) = "LABEL" Or TIOBJ$(I%) = "PARAM" Then
        '
        Combo1.TEXT = FUENTE$(I%)
        Text8.TEXT = TAFUEN(I%)
        Check1.Value = NEGRI(I%):     Text3.FontBold = Check1.Value
        Check2.Value = ITALI(I%):     Text3.FontItalic = Check2.Value
        Check3.Value = SUBRA(I%):     Text3.FontUnderline = Check3.Value
        '
        Text3.Top = 56.7 * V1(I%) - 50: Text3.Height = 56.7 * (V2(I%) - V1(I%))
        Text3.Left = 56.7 * H1(I%) - 50: Text3.Width = 56.7 * (H2(I%) - H1(I%))
        '
        Text3.ForeColor = RGB(0, 0, 0)
        If TIOBJ$(I%) = "PARAM" Then Text3.ForeColor = RGB(255, 0, 0)
        Text3 = TEXTO(I%)
        Call Text3_Change
        Text3.Visible = True
        Text3.SetFocus
        Exit Sub
        '
     ElseIf TIOBJ$(I%) = "BOX" Then
        MODO$ = "BOX"
        ESTAPRETADO% = 2
        
  End If
  '
End Sub

Private Sub Picture1_GotFocus()
   Text3.Visible = False
   Text3 = ""
   Text3.Width = 0
   Text3.Height = 0
   Command13.Visible = False
   Text11.Visible = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    XMOUSE = X: YMOUSE = Y
    If MODO$ = "" Then Exit Sub
    If Text3.Visible = True Then Exit Sub
    '
    If ESTAPRETADO = 0 Then
       ESTAPRETADO% = 1
       IOBJ% = CANOBJ% + 1
       Label14 = IOBJ%
       '
       MAXCAI% = 1
25     ONAM$ = MODO$ + "-" + TRIM$(Str$(MAXCAI%))
       For KKI% = 1 To CANOBJ%
          If ONAME$(KKI%) = ONAM$ Then
             MAXCAI% = MAXCAI% + 1
             GoTo 25
          End If
       Next KKI%
       Text9 = ONAM$
       '
       XORI = X: XMOUSE = X
         Text4 = FORMATNUM$(X / 56.7, "F7.1")
         Text5 = FORMATNUM$(X / 56.7, "F7.1")
       YORI = Y: YMOUSE = Y
         Text6 = FORMATNUM$(Y / 56.7, "F7.1")
         Text7 = FORMATNUM$(Y / 56.7, "F7.1")
    End If
    '
    If ESTAPRETADO = 2 Then
       ESTAPRETADO% = 1
       '
       MUEVETODO% = 1
       If Abs(XMOUSE - 56.7 * XVALO(Text5)) < 100 Then
          If Abs(YMOUSE - 56.7 * XVALO(Text7)) < 100 Then
             MUEVETODO% = 0
          End If
       End If
       Screen.MousePointer = 1
       If MUEVETODO% = 1 Then
          Screen.MousePointer = 5
          Exit Sub
       End If
       Screen.MousePointer = 8
       '
       IOBJ% = XVALO(Label14)
       XORI = 56.7 * XVALO(Text4): XMOUSE = X
         Text5 = FORMATNUM$(X / 56.7, "F7.1")
       YORI = 56.7 * XVALO(Text6): YMOUSE = Y
         Text7 = FORMATNUM$(Y / 56.7, "F7.1")
    End If
End Sub
'
Private Sub Picture1_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If ESTAPRETADO% <> 1 Then Exit Sub
    '
    If MUEVETODO% = 1 Then
       MUEVETODO% = 0
       XORI = H1(IOBJ%) * 56.7
       YORI = V1(IOBJ%) * 56.7
       XMOUSE = H2(IOBJ%) * 56.7
       YMOUSE = V2(IOBJ%) * 56.7
    End If
    Screen.MousePointer = 1
    '
    If MODO$ <> "LINE" Then
       If XORI > XMOUSE Then SWAP XORI, XMOUSE
          Text4 = FORMATNUM$(XORI / 56.7, "F7.1")
          Text5 = FORMATNUM$(XMOUSE / 56.7, "F7.1")
       If YORI > YMOUSE Then SWAP YORI, YMOUSE
          Text6 = FORMATNUM$(YORI / 56.7, "F7.1")
          Text7 = FORMATNUM$(YMOUSE / 56.7, "F7.1")
    End If
    '
    Call GUARDAELOBJETO
    '
    ESTAPRETADO% = 0
    If MODO$ = "LABEL" Or MODO$ = "PARAM" Or MODO$ = "TEXT" Then
       Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B ' BORRA RECUADRO
       If XMOUSE <> XORI Or YMOUSE <> YORI Then
          Text3.Visible = True
          Text3.Left = XORI + 10
          Text3.Top = YORI + 10
          On Error Resume Next
            Text3.Width = X - XORI - 20
            If MODO$ = "LABEL" Then
                 Text3.ForeColor = RGB(0, 0, 0)
                 Call Text3_Change
                 Text3.SetFocus
               Else
                 Text3.ForeColor = RGB(255, 0, 0)
                 Text3 = Text10
                 Call Text3_Change
                 Text11.SetFocus
            End If
          On Error GoTo 0
       End If
    End If
    '
End Sub

Sub GUARDAELOBJETO()
    ' GUARDA EL OBJETO
    If Text4 <> Text5 Or Text6 <> Text7 Then
       IOBJ% = XVALO(Label14)
       If CANOBJ% < IOBJ% Then CANOBJ% = IOBJ%
       '
       PPXX% = InStr(Label15, "("): If PPXX% < 1 Then PPXX% = 1 + Len(Label15)
       MODO$ = TRIM$(Left$(Label15, PPXX% - 1))
       TIOBJ$(IOBJ%) = MODO$
       ONAME$(IOBJ%) = Text9.TEXT
       H1(IOBJ%) = XVALO(Text4)
       H2(IOBJ%) = XVALO(Text5)
       V1(IOBJ%) = XVALO(Text6)
       V2(IOBJ%) = XVALO(Text7)
       '
       TEXTO(IOBJ%) = Text10
       FUENTE$(IOBJ%) = Combo1.TEXT
       TAFUEN(IOBJ%) = Text8
       '
    End If
End Sub
'
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ESTAPRETADO = 1 Then
      Call BORRAELOBJETOACTIVO
      If MUEVETODO = 1 Then
         DIFEX = X - XMOUSE: DIFEY = Y - YMOUSE
         XMOUSE = X: YMOUSE = Y
         H1(IOBJ%) = H1(IOBJ%) + DIFEX / 56.7
         H2(IOBJ%) = H2(IOBJ%) + DIFEX / 56.7
         V1(IOBJ%) = V1(IOBJ%) + DIFEY / 56.7
         V2(IOBJ%) = V2(IOBJ%) + DIFEY / 56.7
         Call DIBUOBJ(IOBJ%)
         Exit Sub
      End If
      '
      If MODO$ = "LINE" Then
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255)
         Else
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B
      End If
      XMOUSE = X: YMOUSE = Y
        Text5 = FORMATNUM$(X / 56.7, "F7.1")
        Text7 = FORMATNUM$(Y / 56.7, "F7.1")
      Picture1.DrawStyle = 2
      Picture1.DrawWidth = 1
      If MODO$ = "LINE" Or MODO$ = "BOX" Then
         Picture1.DrawStyle = 0
         Picture1.DrawWidth = 1
           ESPESOR = XVALO(Text8): If ESPESOR < 1 Then ESPESOR = 1
           Picture1.DrawWidth = ESPESOR
      End If
      If MODO$ = "LINE" Then
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE)
        Else
           Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), , B
      End If
   End If
End Sub

Sub BORRAELOBJETOACTIVO()
   IOBJ% = XVALO(Label14)
   If TIOBJ$(IOBJ%) = "LINE" Then
         Picture1.Line (56.7 * H1(IOBJ%), 56.7 * V1(IOBJ%))-(56.7 * H2(IOBJ%), 56.7 * V2(IOBJ%)), RGB(255, 255, 255)
      ElseIf TIOBJ$(IOBJ%) = "BOX" Then
         Picture1.Line (56.7 * H1(IOBJ%), 56.7 * V1(IOBJ%))-(56.7 * H2(IOBJ%), 56.7 * V2(IOBJ%)), RGB(255, 255, 255), B
   End If
End Sub

Private Sub Picture2_DblClick()
   '
   Picture2.Visible = False
   Frame1.Visible = True
   VScroll1.Visible = True
   HScroll1.Visible = True
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
   Command8.Enabled = True
   Line2.Visible = True
   '
   Picture1.Visible = True
   Picture1.Refresh
   '
End Sub

Private Sub Picture5_Click()
AAA = BBB
End Sub

Private Sub Picture6_Click()
    Line4.Visible = False
End Sub

Private Sub Picture6_DblClick()
    Line4.X1 = XREGLA
    Line4.X2 = XREGLA
    Line4.Visible = True
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XREGLA = X - Picture1.Left
End Sub

Private Sub Picture7_Click()
    Line3.Visible = False
End Sub

Private Sub Picture7_DBLClick()
    Line3.Y1 = YREGLA
    Line3.Y2 = YREGLA
    Line3.Visible = True
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    YREGLA = Y - Picture1.Top
    'MsgBox YREGLA
End Sub

Private Sub Text1_Change()
    '
    CODFOR$ = TRIM$(Text1)
    If TRIM$(CODFOR$) <> "" Then
       CANOBJ% = 0
       Call CARGADATFOR(CODFOR$)
       Call REDIBUJA
       VScroll2.Max = CANOBJ%
       Label14 = CStr(CANOBJ%)
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = VALACT2$("INDIFRM")
        Call SHOWFORM
    End If
    '
End Sub

Private Sub Text10_Change()
    IKX% = XVALO(Label14)
    TEXTO(IKX%) = Text10
End Sub

Private Sub Text11_Change()
    IKX% = XVALO(Label14)
    FORPA$(IKX%) = UCase$(Text11)
    If Left$(FORPA$(IKX%), 1) = "F" Then
         VALOR = "1.23456789"
         VATEX$ = FORMATNUM$(VALOR, FORPA$(IKX%))
         LOREF% = Len(VATEX$)
         VATEX$ = TRIM$(VATEX$)
         VINI% = 1
         While Len(VATEX$) < LOREF%
            VINI% = VINI% + 1 Mod 10
            VATEX$ = TRIM$(Str$(VINI%)) + VATEX$
         Wend
      ElseIf Left$(FORPA$(IKX%), 1) = "D" Then
         If FORPA$(IKX%) = "D" Then FORPA$(IKX%) = "D8"
         VALOR = FECHANUM("09/09/09")
         VATEX$ = FORMATNUM$(VALOR, FORPA$(IKX%))
         LOREF% = XVALO(Mid$(FORPA$(IKX%), 2))
         If LOREF% < 1 Then LOREF% = 8
         VATEX$ = TRIM$(Left$(VATEX$, LOREF%))
      ElseIf Left$(FORPA$(IKX%), 1) = "A" Then
         VALOR = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         LOREF% = XVALO(Mid$(FORPA$(IKX%), 2))
         While Len(VALOR) < LOREF%: VALOR = VALOR + "ABCDEFGHIJKLMNOPQRSTUVWXYZ": Wend
         VATEX$ = FORMATNUM$(VALOR, FORPA$(IKX%))
      ElseIf Left$(FORPA$(IKX%), 1) = "C" Then
         VALOR = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         LOREF% = XVALO(Mid$(FORPA$(IKX%), 2))
         VATEX$ = FORMATNUM$(VALOR, "A" + Mid$(FORPA$(IKX%), 2))
    End If
    If Text3.Visible = True Then
       Text3 = VATEX$
    End If
    Text10 = VATEX$
    Call Text8_LostFocus
End Sub

Private Sub Text2_LOSTFOCUS()
    Call GRATITUFORM
End Sub

Sub CARLISFUEN()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
         Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
         Call FINAL("")
    End If
    '
    'Call BLANARCH("FUENPRIN")
    Call SETULTREG("FUENPRIN", 0)
    On Error Resume Next
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        List1.AddItem Printer.FontName
    Next I%
    On Error GoTo 0
    '
    JJ& = 0
    For I% = 1 To List1.ListCount
      YY$ = TRIM$(List1.List(I% - 1))
      If YY$ <> "" Then
        If InStr(UCase$(YY$), "BENGUIAT") < 1 Then
          XX$ = REGBLAN$("FUENPRIN")
          Call REPLA(XX$, YY$, 1, 42)
          Printer.FontName = YY$
          A1 = 0: A2 = 0
          On Error Resume Next
          A1 = Printer.TextWidth(String$(80, "l"))
          A2 = Printer.TextWidth(String$(80, "M"))
          On Error GoTo 0
          If A1 > 0 And A2 > 0 Then
            If A1 = A2 Then
              Call REPLA(XX$, "(fija)", 43, 6)
            End If
            'Call REGAPP("FUENPRIN", XX$)
            JJ& = JJ& + 1
            Call GRAREG("FUENPRIN", XX$, JJ&)
          End If
        End If
      End If
    Next I%
    '
    Call SETULTREG("FUENPRIN", JJ&)
    Call CIERRARCH("FUENPRIN")
    '
End Sub

Sub CARINDIFORM()
    '
    Screen.MousePointer = 11
    Dim RF As String * 128
    File1.Path = LUDAT$
    File1.Pattern = "*.FRM"
    JJ& = 0
    '
    For U& = 0 To File1.ListCount - 1
      File1.ListIndex = U&
      ARXYZ$ = UCase$(File1.filename)
      PPXX% = InStr(ARXYZ$, ".FRM")
      If PPXX% > 1 Then
        ARXYZ$ = TRIM$(Left$(ARXYZ$, PPXX% - 1))
        If ARXYZ$ <> "" Then
          TITUFOR$ = ""
          RF11$ = LOADFILE(LUDAT$ + ARXYZ$ + ".FRM")
          If Len(RF11$) >= 128 Then
            '
            For I1& = 1 To Len(RF11$) Step 128
              RF$ = Mid$(RF11$, I1&, 128)
              If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                TITUFOR$ = Mid$(RF$, 9, 48)
                GoTo 10
              End If
            Next I1&
            '
10          INF$ = REGBLAN$("INDIFRM")
            Call REPLA(INF$, ARXYZ$, 1, 8)
            Call REPLA(INF$, TITUFOR$, 9, 48)
            JJ& = JJ& + 1
            Call GRAREG("INDIFRM", INF$, JJ&)
            '
          End If
        End If
      End If
19  Next U&
    '
    Call SETULTREG("INDIFRM", JJ&)
    Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'
Sub SHOWFORM()
   '
   Static CACOMM%, PRICOMM$(128)
   '
   FORMALTER.PSPOOL.Clear
 '  MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
 '  MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   MUESTRAFORBLA% = 1
'   Call CARFORWIN
   MUESTRAFORBLA% = 0
'   Call SETSPOOL("ENDDOC", "")
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
   End If
   '
10 TPSP$ = ""
   For KKII% = 1 To FORMALTER.PSPOOL.ListCount
      TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKII%
   '
   Picture1.Cls
   Picture1.Top = 0
   Picture1.Refresh
   Picture1.Height = 7070
   Picture1.Width = 11700
   Picture1.DrawWidth = 3
   Picture1.DrawStyle = 0
   Picture1.FillStyle = 1
   '
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   '
   SPOOLSTRING = TPSP$
   '
20 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        If COPRI% = 2 Or COPRI% = 3 Then GoTo 30
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMM(COPRI%, ATRIP$)
      End If
      GoTo 20
      '
   End If
   '
30 Screen.MousePointer = 1
   ANCHPAGINA = Picture1.Width
   Picture1.Refresh
   VScroll1.Min = 0
   VScroll1.Value = 0
   VScroll1.Max = 1
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
   HScroll1.Min = 0
   HScroll1.Value = 0
   HScroll1.Max = 1
   If ANCHPAGINA > Frame1.Width Then
      HScroll1.Max = (ANCHPAGINA - Frame1.Width) / 100
   End If
   '
   Picture1.Refresh
   '
'   MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
'   MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
'   DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
'   DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   Picture1.ForeColor = RGB(0, 100, 0)
   Picture1.Line (56.7 * (MARIZW + 1), 56.7 * (DETDESMM + MARSUW))-(EXTREDET - 56.7, 56.7 * (DETHASMM + MARSUW)), RGB(200, 255, 255), BF
'   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   INTERLI = Val(INTERX$)
   If INTERLI < 1 Then INTERLI = 4
   '
   AFORMX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRM")
   LNXP& = 1 + Len(AFORMX$) / 128
   For KU& = 1 To LNXP&
     RFP$ = Mid$(AFORMX$, 128 * KU& - 127, 128)
200  FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
     FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
     FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
     If TRIM$(FORMA1$ + FORMA2$) <> "" Then
       On Error Resume Next
       Picture1.FontName = FUENX$
       If Err Then
         Picture1.FontName = "Courier"
       End If
       On Error GoTo 0
       On Error Resume Next
       Picture1.FontSize = Str$(CVS(Mid$(RFP$, 119, 4)))
       On Error GoTo 0
       Picture1.FontBold = False
       If Asc(Mid$(RFP$, 123, 1)) > 32 Then
         Picture1.FontBold = True
       End If
       Picture1.FontItalic = False
       If Asc(Mid$(RFP$, 124, 1)) > 32 Then
         Picture1.FontItalic = True
       End If
       Picture1.FontUnderline = False
       If Asc(Mid$(RFP$, 125, 1)) > 32 Then
         Picture1.FontUnderline = True
       End If
       '
       PH1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
       PV1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
       FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
       '
       PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
       PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
       FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
       '
       If FORMA1$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA1$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA1$, 1) = "A" Or Left$(FORMA1$, 1) = "C" Or Left$(FORMA1$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA1$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA1$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA1$)
           CAREMA% = Val(Mid$(FORMA1$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH1
               Picture1.CurrentY = PV1
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV1 <= (MARSUW + 4) * 56.7 Then
                   PV1 = PV1 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV1 = PV1 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 880
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV1 <= (MARSUW + 4) * 56.7 Then
             PV1 = PV1 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH1
           Picture1.CurrentY = PV1
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
880    If FORMA2$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA2$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA2$, 1) = "A" Or Left$(FORMA2$, 1) = "C" Or Left$(FORMA2$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA2$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA2$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA2$)
           CAREMA% = Val(Mid$(FORMA2$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH2
               Picture1.CurrentY = PV2
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV2 <= (MARSUW + 4) * 56.7 Then
                   PV2 = PV2 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV2 = PV2 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 890
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV2 <= (MARSUW + 4) * 56.7 Then
             PV2 = PV2 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH2
           Picture1.CurrentY = PV2
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
890 End If
     '
   Next KU&
   '
   If ALTUPAGINA < Picture1.Height Then
     ALTUPAGINA = Picture1.Height
   End If
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
End Sub

Sub SHOWCOMM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     '
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     '
   Case 5
     Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture1.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture1.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture1.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture1.DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     Picture1.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture1.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16))
     BB11 = Val(Mid$(ATRIP$, 17, 16))
     AA22 = Val(Mid$(ATRIP$, 33, 16))
     BB22 = Val(Mid$(ATRIP$, 49, 16))
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
     If Picture1.Height < ALTUPAGINA - 1 Then
       Picture1.Height = ALTUPAGINA
     End If
     If AA11 > Picture1.Width - 300 Then
        Picture1.Width = AA11 + 300
     End If
     If AA22 > Picture1.Width - 300 Then
        Picture1.Width = AA22 + 300
     End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     Picture1.FontName = TRIM$(ATRIP$)
   Case 14
     Picture1.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     Picture1.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     Picture1.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     If Val(TRIM$(ATRIP$)) > Picture1.Width - 300 Then
        Picture1.Width = Val(TRIM$(ATRIP$)) + 300
     End If
     Picture1.CurrentX = Val(TRIM$(ATRIP$))
   Case 18
     If Val(TRIM$(ATRIP$)) + 200 > ALTUPAGINA Then ALTUPAGINA = Val(TRIM$(ATRIP$)) + 1000
       If Picture1.Height < ALTUPAGINA - 1 Then
         Picture1.Height = ALTUPAGINA
       End If
     Picture1.CurrentY = Val(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture1.Print ATRIP$;
        Else
         Picture1.Print ATRIP$
     End If
   Case 20
     Picture1.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     REDUFORMU = 1
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub

Private Sub Text3_Change()
   '
   Text3.Font = Combo1.TEXT
   On Error Resume Next
   Text3.FontSize = XVALO(Text8)
   On Error GoTo 0
   If Text3.Visible = True Then
       Text10 = Text3
       Call BORRATRASTEX
   End If
   '
   Picture1.CurrentY = Text3.Top + 50
   Picture1.CurrentX = Text3.Left + 50
   Picture1.Font = Text3.Font
   Picture1.FontSize = Text3.FontSize
   Picture1.FontBold = Text3.FontBold
   Picture1.FontItalic = Text3.FontItalic
   Picture1.FontUnderline = Text3.FontUnderline
      FCOL = Picture1.ForeColor
   Picture1.ForeColor = Text3.ForeColor
   Picture1.Print Text3.TEXT
   Picture1.ForeColor = FCOL     ' PARA RECUPERAR EL COLOR
   '
   Label17.Font = Text3.Font
   Label17.FontSize = Text3.FontSize
   Label17.FontBold = Text3.FontBold
   Label17.FontItalic = Text3.FontItalic
   Label17.FontUnderline = Text3.FontUnderline
   Label17 = Text3.TEXT + "*M"
   If Label17.Width > Text3.Width Then
      If Text3 <> "" Then
         Text3.Width = Label17.Width
         Text5 = FORMATNUM$((Text3.Left + Text3.Width) / 56.7, "F7.1")
      End If
   End If
   Text3.Height = Label17.Height - 31
   '
   Call GUARDAELOBJETO
   '
End Sub

Sub BORRATRASTEX()
   If Text3.Visible = True Then
      Picture1.Line (Text3.Left, Text3.Top)-(Text3.Left + Text3.Width, Text3.Top + Text3.Height), RGB(255, 255, 255), BF
   End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
   If TIOBJ$(IOBJ%) = "PARAM" Then
      KeyCode = 0
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If TIOBJ$(IOBJ%) = "PARAM" Then
      KeyAscii = 0
   End If
End Sub

'
Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture1.Line (XORI, YORI)-(XMOUSE, YMOUSE), RGB(255, 255, 255), B
   Picture1.Line (Text3.Left, Text3.Top)-(Text3.Left + Text3.Width, Text3.Top + Text3.Height), RGB(255, 255, 255), BF
   '
   ESTAPRETADO% = 2
   XORI = X: YORI = Y
   If XORI >= Text3.Width - 100 Then
      ESTAPRETADO% = 3
   End If
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If ESTAPRETADO% = 2 Then
      Text3.Top = Text3.Top + Y - YORI
      Text3.Left = Text3.Left + X - XORI
     ElseIf ESTAPRETADO% = 3 Then
      Text3.Width = Text3.Width + X - XORI
      XORI = X
   End If
   '
   If ESTAPRETADO% > 1 Then
      Text4 = FORMATNUM$((Text3.Left + 50) / 56.7, "F7.1")
      Text5 = FORMATNUM$((Text3.Left + 50 + Text3.Width) / 56.7, "F7.1")
      Text6 = FORMATNUM$((Text3.Top + 50) / 56.7, "F7.1")
      Text7 = FORMATNUM$((Text3.Top + 50 + Text3.Height) / 56.7, "F7.1")
   End If
   '
End Sub

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Text3_Change
   ESTAPRETADO = 0
End Sub

Private Sub Text4_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      H1(IOBJ%) = XVALO(Text4)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Left = 56.7 * XVALO(Text4)
          Call Text3_Change
        Else
          Call REDIBUJA
      End If
   End If
End Sub

Private Sub Text5_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      H2(IOBJ%) = XVALO(Text5)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Left = 56.7 * XVALO(Text5) - Text3.Width
          Call Text3_Change
        Else
          Call REDIBUJA
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text5_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text6_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      V1(IOBJ%) = XVALO(Text6)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Top = 56.7 * XVALO(Text6)
          Call Text3_Change
        Else
          Call REDIBUJA
      End If
   End If
End Sub

Private Sub Text6_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text7_Change()
   If PORCOORDE% = 1 Then
      IOBJ% = XVALO(Label14)
      V2(IOBJ%) = XVALO(Text7)
      If Text3.Visible = True Then
          Call BORRATRASTEX
          Text3.Top = 56.7 * XVALO(Text7) - Text3.Height
          Call Text3_Change
        Else
          Call REDIBUJA
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
   PORCOORDE% = 1
End Sub

Private Sub Text4_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text5_lostfocus()
   PORCOORDE% = 0
End Sub
'
Private Sub Text6_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text7_lostfocus()
   PORCOORDE% = 0
End Sub

Private Sub Text8_Change()
   Call PREDITEXTO
   If XVALO(Text8) >= 1 Then
     If IOBJ% <= CANOBJ% + 1 Then
       TAFUEN$(IOBJ%) = TRIM$(Text8)
       MODOI$ = UCase$(Label15)
       If MODOI$ = "LABEL" Or MODOI$ = "PARAM" Then
           UTAFUE = TRIM$(Text8)
         ElseIf MODOI$ = "BOX" Or MODOI$ = "LINE" Then
           Call DIBUOBJ(IOBJ%)
       End If
     End If
   End If
End Sub

Sub PREDITEXTO()
   If Text3.Visible = True Then
     Call BORRATRASTEX
     If Combo1.TEXT <> "" Then
         Text3.Font = Combo1.TEXT
       Else
         Combo1.TEXT = Text3.Font
     End If
     On Error Resume Next
     Text3.FontSize = Text8.TEXT
     Text3.FontBold = Check1.Value
     Text3.FontItalic = Check2.Value
     Text3.FontUnderline = Check3.Value
     On Error GoTo 0
     Call Text3_Change
   End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Call Text8_LostFocus
   End If
End Sub

Private Sub Text8_LostFocus()
   If Text3.Visible = False Then
      Call REDIBUJA
   End If
End Sub

Private Sub Text9_Change()
   IOBJ% = XVALO(Label14)
   If IOBJ% <= CANOBJ% + 1 Then
      ONAME$(IOBJ%) = UCase$(Text9)
   End If
End Sub

'
Private Sub VScroll1_Scroll()
   Picture4.Top = -30 - VScroll1.Value
   Picture7.Top = Picture4.Top
   Call DIBUREGVER
End Sub

Private Sub Command9_Click()
   '
   VScroll1.Visible = False
   HScroll1.Visible = False
   Command1.Enabled = False
   Command2.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
   Command8.Enabled = False
   Line2.Visible = False
   '
   ALPAMAX = 0: ANPAMAX = 0
     If Picture1.Width > ANPAMAX Then
       ANPAMAX = Picture1.Width + 200
     End If
     If Picture1.Height > ALPAMAX Then
       ALPAMAX = Picture1.Height + 200
     End If
   REDUV = ScaleHeight / ALPAMAX
   REDUH = ScaleWidth / ANPAMAX
   REDUC = REDUV: If REDUH < REDUC Then REDUC = REDUH
   REDUFORMU = REDUC
   '
   Picture2.Cls
   Picture2.Top = (ScaleHeight - ALPAMAX * REDUC) / 2
   Picture2.Left = (ScaleWidth - ANPAMAX * REDUC) / 2
   Picture2.Height = ALPAMAX * REDUC
   Picture2.Width = ANPAMAX * REDUC
   Picture2.Refresh
   Picture2.DrawWidth = 2
   Picture2.DrawStyle = 0
   Picture2.FillStyle = 1
   '
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
   Screen.MousePointer = 11
   TPSP$ = SPOOLSTRING$
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMZOOM(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Frame1.Visible = False
   Screen.MousePointer = 1
   Picture2.Refresh
      Picture1.Visible = False
      Picture2.Visible = False
   Picture2.Visible = True
   Picture2.Refresh
   '
End Sub

Sub SHOWCOMMZOOM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'ALTUPAGINA = Val(TRIM$(ATRIP$))
   Case 5
     'Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture2.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture2.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture2.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture2.DrawWidth = 0.66 * Val(TRIM$(ATRIP$))
   Case 10
     Picture2.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture2.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16)) * REDUFORMU
     BB11 = Val(Mid$(ATRIP$, 17, 16)) * REDUFORMU
     AA22 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
     BB22 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     'If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     'If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     'For KKI% = 0 To CAPAGINAS%
     '  If Picture1(KKI%).Height < ALTUPAGINA Then
     '    Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
     'If AA11 > Picture1.Width Then
     '   Picture1.Width = AA11
     'End If
     'If AA22 > Picture1.Width Then
     '   Picture1.Width = AA22
     'End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     TEPRUORI.FontName = TRIM$(ATRIP$)
   Case 14
     TEPRUORI.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     TEPRUORI.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     TEPRUORI.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     Picture2.CurrentX = Val(TRIM$(ATRIP$)) * REDUFORMU
   Case 18
     Picture2.CurrentY = (Val(TRIM$(ATRIP$)) + TOPEPAGINA) * REDUFORMU
   Case 19
     TEPRUORI.Caption = ATRIP$
     AMATEX = TEPRUORI.Width * REDUFORMU
     TEPRURED.Font = "Arial"
     If InStr(UCase$(TEPRUORI.FontName), "COURIER") > 0 Then
        TEPRURED.FontName = "Courier New"
     End If
     TEPRURED.FontBold = False
     TEPRURED.Caption = ATRIP$
     For KKKK = 24 To 1 Step -1
       TEPRURED.FontSize = KKKK / 2
       If TEPRURED.Width <= AMATEX Then GoTo 17
     Next KKKK
     '
17   Picture2.FontName = TEPRURED.FontName '"Arial"
     Picture2.FontSize = KKKK / 2
     Picture2.FontBold = False
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture2.Print ATRIP$;
        Else
         Picture2.Print ATRIP$
     End If
   Case 20
     Picture2.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'
Sub INCLUDEIMG(ATRIP$, MODOPRE$)
   '
   BBUS% = Val(Mid$(ATRIP$, 1, 16))
   IMAGENUMBER& = Val(Mid$(ATRIP$, 17, 16))
   PHX1 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
   PVX1 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
   ACHX = Val(Mid$(ATRIP$, 65, 16)) * REDUFORMU
   ALTX = Val(Mid$(ATRIP$, 81, 16)) * REDUFORMU
   '
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELI& = CVS(Mid$(RFF$, KKI%, 4))
     If RELI& > 0 Then
       ZZ$ = REGLEIDO$("PIMGLST", RELI&)
       If CVS(Mid$(ZZ$, 53, 4)) = IMAGENUMBER& Then GoTo 10
     End If
   Next KKI%
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
10 NOMBIMAG$ = TRIM$(Mid$(ZZ$, 3, 38))
   LARGIMAG& = CVS(Mid$(ZZ$, 41, 4))
   FECHIMAG# = CVD(Mid$(ZZ$, 45, 8))
   FIRSTRECORD& = CVS(Mid$(ZZ$, 57, 4))
   LASTRECORD& = CVS(Mid$(ZZ$, 61, 4))
   '
   If FIRSTRECORD& < 1 Or LASTRECORD& < FIRSTRECORD& Or LASTRECORD& > ULTREG&("PIMGSPL") Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
   If FIRSTRECORD& > 1 Then
      XX$ = REGLEIDO$("PIMGSPL", FIRSTRECORD& - 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   If LASTRECORD& < ULTREG&("PIMGSPL") Then
      XX$ = REGLEIDO$("PIMGSPL", LASTRECORD& + 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   IMAGSTRI$ = String$(508 * (1 + LASTRECORD& - FIRSTRECORD&), 0)
   PINSERT& = 1
   For U& = FIRSTRECORD& To LASTRECORD&
      XX$ = REGLEIDO$("PIMGSPL", U&)
      If CVS(Left$(XX$, 4)) <> IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
      Mid$(IMAGSTRI$, PINSERT&, 508) = Mid$(XX$, 5, 508)
      PINSERT& = PINSERT& + 508
   Next U&
   '
   If Len(IMAGSTRI$) < LARGIMAG& Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   IMAGSTRI$ = Left$(IMAGSTRI$, LARGIMAG&)
   '
   NX% = FreeFile
   Open LUCOM$ + NOMBIMAG$ For Binary As #NX%
   Put #NX%, , IMAGSTRI$
   Close #NX%
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(LUCOM$ + NOMBIMAG$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   AA11 = PHX1
   BB11 = PVX1 + TOPEPAGINA * REDUFORMU
   AA22 = AA11 + ACHX
   BB22 = BB11 + ALTX
   If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
   If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
   '
   If Picture1.Height < ALTUPAGINA - 1 Then
     Picture1.Height = ALTUPAGINA
   End If
   '
   If AA11 > Picture1.Width Then
     Picture1.Width = AA11
   End If
   '
   If AA22 > Picture1.Width Then
     Picture1.Width = AA22
   End If
   '
   If REDUFORMU = 1 Then
       Picture1.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA, ACHX, ALTX
     Else
       Picture2.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA * REDUFORMU, ACHX, ALTX
   End If
   '
End Sub

Sub GRATITUFORM()
   '
   Screen.MousePointer = 11
   If TRIM$(Text1.TEXT) = TRIM$(CODFOR$) Then
     If TRIM$(CODFOR$) <> "" Then
       ARXYZ$ = TRIM$(CODFOR$)
       TITUFOR$ = Text2.TEXT
       NX% = FILEOPEN%(LUDAT$ + ARXYZ$ + ".FRM", 0, 128)
       Dim RF As String * 128
       '
       For I% = 1 To LOF(NX%) / 128
         Get #NX%, I%, RF$
         If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
           RFY$ = RF$
           Call REPLA(RFY$, TITUFOR$, 9, 48)
           LSet RF$ = RFY$
           Put #NX%, I%, RF$
           GoTo 10
         End If
       Next I%
       '
       I% = 1 + LOF(NX%) / 128
       RFY$ = Space$(56) + String$(72, 0)
       Call REPLA(RFY$, "TITUFORM", 1, 8)
       Call REPLA(RFY$, TITUFOR$, 9, 48)
       LSet RF$ = RFY$
       Put #NX%, I%, RF$
       '
10     Close #NX%: NX% = 0
       '
       For V& = 1 To ULTREG&("INDIFRM")
         INF$ = REGLEIDO$("INDIFRM", V&)
         If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
           If TITUFOR$ <> "" Then
             Call REPLA(INF$, TITUFOR$, 9, 48)
             Call GRAREG("INDIFRM", INF$, V&)
           End If
           GoTo 19
         End If
       Next V&
       '
       INF$ = REGBLAN$("INDIFRM")
       Call REPLA(INF$, ARXYZ$, 1, 8)
       Call REPLA(INF$, TITUFOR$, 9, 48)
       Call REGAPP("INDIFRM", INF$)
       '
     End If
   End If
   '
19 Call CIERRARCH("INDIFRM")
   Call FRESHECHO("INDIFRM")
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub VScroll2_Change()
    Text3.Visible = False
    Label14 = CANOBJ% + 1 - VScroll2.Value
End Sub

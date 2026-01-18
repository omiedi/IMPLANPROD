VERSION 5.00
Begin VB.Form CALCAP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planeamiento de Capacidad (Equipos y M.Obra)"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "CALCAP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   Begin VB.CommandButton Command19 
      Caption         =   "Query"
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
      Left            =   10920
      Picture         =   "CALCAP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Acceso Directo a Consultas y Listados"
      Top             =   3015
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Setup"
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
      Left            =   10920
      Picture         =   "CALCAP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   2115
      Width           =   855
   End
   Begin VB.CommandButton Command17 
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
      Height          =   855
      Left            =   10920
      Picture         =   "CALCAP.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Ayuda en Línea (si Disponible)"
      Top             =   1215
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10920
      TabIndex        =   49
      Top             =   5520
      Width           =   855
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   51
         Top             =   840
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   12000
         End
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
         Height          =   735
         Left            =   120
         Picture         =   "CALCAP.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Re-Cálculo de Capacidad"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   165
         TabIndex        =   53
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Master"
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
      Left            =   10920
      Picture         =   "CALCAP.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Maestro de Items de Stock"
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Resumen de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   6510
      TabIndex        =   13
      Top             =   180
      Width           =   4260
      Begin VB.ListBox List5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "CALCAP.frx":1374
         Left            =   90
         List            =   "CALCAP.frx":1376
         TabIndex        =   21
         Top             =   3150
         Width           =   4080
      End
      Begin VB.PictureBox Picture5 
         Height          =   435
         Left            =   90
         ScaleHeight     =   375
         ScaleWidth      =   4020
         TabIndex        =   19
         Top             =   2850
         Width           =   4080
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   100
            TabIndex        =   20
            Top             =   300
            Width           =   4425
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Cat"
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
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Categoría"
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
            Left            =   550
            TabIndex        =   32
            Top             =   0
            Width           =   2070
         End
         Begin VB.CommandButton Command8 
            Caption         =   "hh:mm:ss"
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
            Left            =   2640
            TabIndex        =   31
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.ListBox List4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "CALCAP.frx":1378
         Left            =   90
         List            =   "CALCAP.frx":137A
         TabIndex        =   16
         Top             =   660
         Width           =   4080
      End
      Begin VB.PictureBox Picture4 
         Height          =   435
         Left            =   90
         ScaleHeight     =   375
         ScaleWidth      =   4020
         TabIndex        =   14
         Top             =   360
         Width           =   4080
         Begin VB.CommandButton Command6 
            Caption         =   "hh:mm:ss"
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
            Left            =   2640
            TabIndex        =   30
            Top             =   0
            Width           =   1150
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Máquina / Equipo"
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
            Left            =   800
            TabIndex        =   29
            Top             =   0
            Width           =   1830
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Codigo"
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
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   785
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   15
            Top             =   550
            Width           =   4425
         End
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total M.O.D:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   23
         Top             =   4760
         Width           =   2030
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   22
         Top             =   4700
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total T.Maquina:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   17
         Top             =   2205
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   6510
      TabIndex        =   4
      Top             =   5400
      Width           =   4260
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1200
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   43
         Top             =   1490
         Width           =   2640
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
            Height          =   600
            Left            =   0
            Picture         =   "CALCAP.frx":137C
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Resumen Total por Máquina / Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen Total por Máquina / Equipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   45
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   795
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   37
         Top             =   980
         Width           =   2640
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
            Height          =   600
            Left            =   0
            Picture         =   "CALCAP.frx":1686
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Resumen por Máquina y Pieza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen por Má- quina y Pieza"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   39
            Top             =   30
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   40
         Top             =   480
         Width           =   2640
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
            Height          =   600
            Left            =   0
            Picture         =   "CALCAP.frx":1990
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Detalle por Máquina, Pieza y Producto"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle por Máquina, Pieza y Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   42
            Top             =   30
            Width           =   1905
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   360
         TabIndex        =   34
         Top             =   2340
         Width           =   3465
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Pantalla"
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
            Left            =   360
            TabIndex        =   36
            Top             =   0
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFFF&
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
            Height          =   225
            Left            =   1920
            TabIndex        =   35
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4230
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4230
         Y1              =   2240
         Y2              =   2240
      End
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
      Height          =   855
      Left            =   10920
      Picture         =   "CALCAP.frx":1C9A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Seleccion de Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   105
      TabIndex        =   1
      Top             =   1365
      Width           =   6350
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
         Height          =   570
         Left            =   5355
         Picture         =   "CALCAP.frx":1DE4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   1020
         Width           =   870
      End
      Begin VB.CommandButton Command21 
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
         Left            =   5355
         Picture         =   "CALCAP.frx":20EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nuevo Programa de Fabricacion"
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton Command22 
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
         Left            =   5775
         Picture         =   "CALCAP.frx":2238
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nuevo Programa por Copia"
         Top             =   420
         Width           =   450
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Programa de Fabricacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   105
      TabIndex        =   0
      Top             =   3265
      Width           =   6350
      Begin VB.CommandButton Command16 
         Height          =   580
         Left            =   5280
         Picture         =   "CALCAP.frx":2382
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Emite Listado del Programa de Fabricación Activo"
         Top             =   4100
         Width           =   855
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Ultimo Cálculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   54
         Top             =   4020
         Width           =   5010
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   330
            Left            =   360
            TabIndex        =   55
            Top             =   2340
            Width           =   3465
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   59
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   58
            Top             =   315
            Width           =   2295
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1770
            TabIndex        =   57
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -600
            TabIndex        =   56
            Top             =   315
            Width           =   2295
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   105
         TabIndex        =   2
         Top             =   760
         Width           =   6105
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   915
         Left            =   105
         ScaleHeight     =   855
         ScaleWidth      =   6045
         TabIndex        =   5
         Top             =   420
         Width           =   6105
         Begin VB.CommandButton Command15 
            Caption         =   "Codigo                   "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   1920
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Descripcion de Producto       "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1920
            TabIndex        =   47
            Top             =   0
            Width           =   2740
         End
         Begin VB.CommandButton Command13 
            Caption         =   "   Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4670
            TabIndex        =   46
            Top             =   0
            Width           =   1150
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   7
            Top             =   550
            Width           =   4425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo           Descripción              Cant."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   6
            Top             =   360
            Width           =   11460
         End
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricacion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   4095
      TabIndex        =   27
      Top             =   840
      Width           =   3585
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacidad de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   3150
      TabIndex        =   26
      Top             =   525
      Width           =   3585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Analisis de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   2415
      TabIndex        =   25
      Top             =   210
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "CALCAP.frx":268C
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Cap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   210
      TabIndex        =   8
      Top             =   630
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      FillColor       =   &H00004000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   750
      Left            =   315
      Shape           =   2  'Oval
      Top             =   270
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10370
      Picture         =   "CALCAP.frx":2996
      Top             =   7570
      Width           =   2010
   End
End
Attribute VB_Name = "CALCAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IPROGRA&

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command10_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List5.ListCount - 1
     AAA$ = Mid$(List5.List(KIL& - 1), 1, 4)
     BBB$ = Mid$(List5.List(KIL&), 1, 4)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command11_Click()
  Command11.Enabled = False
  Call CIERRARCH("*.*")
  If Option1.Value = True Then
      Call MENSERR(24, "No Disponible por Pantalla.\  \Imprima o Exporte a Archivo Texto.")
    Else
      Call CONECRUN("*LICAPAC", "Listado Analisis de Capacidad")
  End If
  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
  Command12.Enabled = False
  Call CIERRARCH("*.*")
  If Option1.Value = True Then
      Call CONECRUN("*COCAPAC", "Consulta Resumen de Capacidad")
    Else
      Call CONECRUN("*RECAPAC", "Listado Resumen de Capacidad")
  End If
  Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 39, 10)
     BBB$ = Mid$(List1.List(KIL&), 39, 10)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command14_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 17, 22)
     BBB$ = Mid$(List1.List(KIL&), 17, 22)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub COMMAND15_CLICK()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 1, 15)
     BBB$ = Mid$(List1.List(KIL&), 1, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command16_Click()
   Command6.Enabled = False
   '
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Call COPYSTRU("DEPROFAB", "LPROFAB")
     RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
     JJ& = 0
     For U& = 1 To Len(RFF$) Step 4
       REFI& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("DEPROFAB", REFI&)
       JJ& = JJ& + 1
       Call GRAREG("LPROFAB", XX$, JJ&)
     Next U&
     Call SETULTREG("LPROFAB", JJ&)
     Call CIERRARCH("LPROFAB")
     '
     Call HEADERS("LPROFAB", "")
     Call HEADERS("LPROFAB", "Programa: " + List3.TEXT)
     Call FINAL("*LPROFAB")
     Screen.MousePointer = 1
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   Call COMUNI("No Disponible")
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command19.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command19.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call CONECRUN("WINTABU", "Consultas y Listados")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
  Call CIERRARCH("*.*")
  If Option1.Value = True Then
      Call MENSERR(24, "No Disponible por Pantalla.\  \Imprima o Exporte a Archivo Texto.")
    Else
      Call CONECRUN("*DECAPAC", "Listado Detalle de Capacidad")
  End If
  Command2.Enabled = True
End Sub

Private Sub CALCAPPED()
'
' \\\OT-5-537-FG-22/07/05///
Dim PEDIPEN(32766) As Single
'
' INICIALIZA EL VECTOR
For I% = 1 To 32766
  PEDIPEN(I%) = 0
Next I%
'
' SELECCIONA EL RANGO DE FECHAS A CONSIDERAR
Call DESHASFECHA(DES%, HAS%, PERIO$)
'
FIN& = ULTREG&("PEDDETA")
MAXCOD% = 0
'
' BUSCA LOS PEDIDOS PENDIENTES
Screen.MousePointer = 11
For U& = 1 To FIN&
  Call TRACE(20, U&, FIN&)
  X$ = REGLEIDO$("PEDDETA", U&)
    FE% = CVI(Mid$(X$, 25, 2)) ' fecha de entrega solicitada
    If FE% >= DES% Then
      If FE% <= HAS% Then
        If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
          CI% = CVI(Mid$(X$, 33, 2))
          If CI% > 0 And CI% <= NUMAS% Then
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            If CAPEN > 0.05 Then
                '
                CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                If CAPEN < 0 Then CAPEN = 0
                CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                If CANOASI < 0 Then CANOASI = 0
                '
                PEDIPEN(CI%) = PEDIPEN(CI%) + CAPEN
                If MAXCOD% < CI% Then MAXCOD% = CI%
                '
            End If
          End If
         End If
        End If
      End If
    End If
Next U&
'
' GENERA EL PROGRAMA DE FABRICACION
If MAXCOD% > 0 Then
  ' GRABA EN EL INDICE DE PROGRAMAS
  FIN1& = ULTREG&("PROFABRI")
  XX$ = REGLEIDO$("PROFABRI", FIN1&)
  NUMPROG& = CVS(Mid$(XX$, 1, 4)) + 1
  '
  Y$ = REGBLAN$("PROFABRI")
  Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
    DESCRIP$ = "Pedidos Pendientes del " & FECHATEX$(DES%) & " al " & FECHATEX$(HAS%)
  Call REPLA(Y$, DESCRIP$, 5, 48)
  Call REPLA(Y$, "1", 63, 1)
  Call REPLA(Y$, "1", 64, 1)
    FIN1& = FIN1& + 1
  Call GRAREG("PROFABRI", Y$, FIN1&)
  Call SETULTREG("PROFABRI", FIN1&)
  '
  ' GRABA EL DETALLE DEL PROGRAMA
  JJ& = ULTREG&("DEPROFAB")
  For I% = 1 To MAXCOD%
    If PEDIPEN(I%) > 0.005 Then
      Y$ = REGBLAN$("DEPROFAB")
      Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
        CI% = I%
      Call REPLA(Y$, MKI$(CI%), 5, 2)
        CAPEN = PEDIPEN(CI%)
      Call REPLA(Y$, MKS$(CAPEN), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("DEPROFAB", Y$, JJ&)
    End If
  Next I%
  Call SETULTREG("DEPROFAB", JJ&)
  '
  ' ACTUALIZA LA LISTA DE PROGRAMAS
  List3.Clear
  For U& = 1 To ULTREG&("PROFABRI")
    XX$ = REGLEIDO$("PROFABRI", U&)
    DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
    MODOCA$ = Mid$(XX$, 63, 2)
    If DENOPRO$ <> "" Then
      TTXX$ = Space$(130)
      Call REPLA(TTXX$, DENOPRO$, 1, 48)
      Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
      Call REPLA(TTXX$, MODOCA$, 129, 2)
      List3.AddItem TTXX$
    End If
  Next U&
  If List3.ListCount > 0 Then
    List3.ListIndex = List3.ListCount - 1
    Call LISPROFAB
    DoEvents
  End If
  '
Else
  Call COMUNI("No Existen Pedidos Pendientes dentro del Rango Seleccionado")
End If
Screen.MousePointer = 1
' \\\OT-5-537-FG-FIN///
'
End Sub

Private Sub CALCAPOF()
'
' \\\OT-5-537-FG-22/07/05///
Dim OFCERRA(32766) As Single
'
' INICIALIZA EL VECTOR
For I% = 1 To 32766
  OFCERRA(I%) = 0
Next I%
'
' SELECCIONA EL RANGO DE FECHAS A CONSIDERAR
Call DESHASFECHA(DES%, HAS%, PERIO$)
'
FIN& = ULTREG&("MOVISTO")
MAXCOD% = 0
'
' BUSCA LAS O/F CERRADAS
Screen.MousePointer = 11
For U& = 1 To FIN&
  Call TRACE(20, U&, FIN&)
  X$ = REGLEIDO$("MOVISTO", U&)
    FE% = CVI(Mid$(X$, 1, 2)) ' fecha del movimiento
    If FE% >= DES% Then
      If FE% <= HAS% Then
        CODMOV$ = Mid$(X$, 21, 2)
        If CODMOV$ = "OF" Then
          CANENT = CVD(Mid$(X$, 88, 8))
          If CANENT > 0.005 Then
            CI% = CVI(Mid$(X$, 3, 2))
            If CI% > 0 And CI% <= NUMAS% Then
              OFCERRA(CI%) = OFCERRA(CI%) + CANENT
              If MAXCOD% < CI% Then MAXCOD% = CI%
            End If
          End If
        End If
      End If
    End If
Next U&
'
' GENERA EL PROGRAMA DE FABRICACION
If MAXCOD% > 0 Then
  ' GRABA EN EL INDICE DE PROGRAMAS
  FIN1& = ULTREG&("PROFABRI")
  XX$ = REGLEIDO$("PROFABRI", FIN1&)
  NUMPROG& = CVS(Mid$(XX$, 1, 4)) + 1
  '
  Y$ = REGBLAN$("PROFABRI")
  Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
    DESCRIP$ = "Por O/F's Cerradas del " & FECHATEX$(DES%) & " al " & FECHATEX$(HAS%)
  Call REPLA(Y$, DESCRIP$, 5, 48)
  Call REPLA(Y$, "2", 63, 1)
  Call REPLA(Y$, "1", 64, 1)
    FIN1& = FIN1& + 1
  Call GRAREG("PROFABRI", Y$, FIN1&)
  Call SETULTREG("PROFABRI", FIN1&)
  '
  ' GRABA EL DETALLE DEL PROGRAMA
  JJ& = ULTREG&("DEPROFAB")
  For I% = 1 To MAXCOD%
    If OFCERRA(I%) > 0.005 Then
      Y$ = REGBLAN$("DEPROFAB")
      Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
        CI% = I%
      Call REPLA(Y$, MKI$(CI%), 5, 2)
        CANENT = OFCERRA(CI%)
      Call REPLA(Y$, MKS$(CANENT), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("DEPROFAB", Y$, JJ&)
    End If
  Next I%
  Call SETULTREG("DEPROFAB", JJ&)
  '
  ' ACTUALIZA LA LISTA DE PROGRAMAS
  List3.Clear
  For U& = 1 To ULTREG&("PROFABRI")
    XX$ = REGLEIDO$("PROFABRI", U&)
    DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
    MODOCA$ = Mid$(XX$, 63, 2)
    If DENOPRO$ <> "" Then
      TTXX$ = Space$(130)
      Call REPLA(TTXX$, DENOPRO$, 1, 48)
      Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
      Call REPLA(TTXX$, MODOCA$, 129, 2)
      List3.AddItem TTXX$
    End If
  Next U&
  If List3.ListCount > 0 Then
    List3.ListIndex = List3.ListCount - 1
    Call LISPROFAB
    DoEvents
  End If
  '
Else
  Call COMUNI("No Existen Entradas por Orden de Fabricación dentro del Rango Seleccionado")
End If
Screen.MousePointer = 1
' \\\OT-5-537-FG-FIN///
'
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   '
   X% = COMALTER%("Opciones de Generación de Programa\\Manual\\Por Pedidos Pendientes\\Por O/F's Cerradas")
   '
   If X% = 1 Then
     URE& = ULTREG&("PROFABRI")
     VDEF$ = MKS$(URE& + 1) + Space$(48)
     OBX$ = OBJPLA$("NUEPROFAB", VDEF$)
     If OBX$ <> "" Then
        DENOPRO$ = TRIM$(Mid$(OBX$, 5, 48))
        If DENOPRO$ <> "" Then
           Call BLOQARCH("PROFABRI")
           REBLA$ = REGBLAN$("PROFABRI")
           URE& = ULTREG&("PROFABRI")
           Call REPLA(REBLA$, MKS$(URE& + 1), 1, 4)
           Call REPLA(REBLA$, DENOPRO$, 5, 48)
           Call REPLA(Y$, "  ", 63, 2)
           Call REGAPP("PROFABRI", REBLA$)
           Call CIERRARCH("PROFABRI")
           '
           ILISTA& = List3.ListCount
           TTXX$ = Space$(130)
           Call REPLA(TTXX$, DENOPRO$, 1, 48)
           Call REPLA(TTXX$, TRIM$(Str$(URE& + 1)), 121, 8)
           Call REPLA(TTXX$, MODOCA$, 129, 2)
           List3.AddItem TTXX$
           List3.ListIndex = ILISTA&
        End If
     End If
   ElseIf X% = 2 Then
     Call CALCAPPED
   Else
     Call CALCAPOF
   End If
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If List3.ListIndex >= 0 Then
     REGA& = Val(Mid$(List3.TEXT, 121, 8))
     DENOA$ = TRIM$(Left$(List3.TEXT, 48))
     URE& = ULTREG&("PROFABRI")
10   VDEF$ = MKS$(URE& + 1) + DENOA$
     OBX$ = OBJPLA$("NUEPROFAB", VDEF$)
     If OBX$ <> "" Then
       DENOPRO$ = TRIM$(Mid$(OBX$, 5, 48))
       If DENOPRO$ <> "" Then
         For KU& = 1 To List3.ListCount
           If DENOPRO$ = TRIM$(Left$(List3.List(KU& - 1), 48)) Then
             Call MENSERR(24, "Denominacion ya Utilizada.\Debe Cambiar Denominacion del Plan")
             GoTo 10
           End If
         Next KU&
         '
         SCMA = Screen.MousePointer
         Screen.MousePointer = 11
         Call BLOQARCH("PROFABRI")
         Call BLOQARCH("DEPROFAB")
         REBLA$ = REGBLAN$("PROFABRI")
         URE& = ULTREG&("PROFABRI")
         Call REPLA(REBLA$, MKS$(URE& + 1), 1, 4)
         Call REPLA(REBLA$, DENOPRO$, 5, 48)
         Call REPLA(REBLA$, "  ", 63, 2)
         Call REGAPP("PROFABRI", REBLA$)
         Call CIERRARCH("PROFABRI")
         '
         ILISTA& = List3.ListCount
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, DENOPRO$, 1, 48)
         Call REPLA(TTXX$, TRIM$(Str$(URE& + 1)), 121, 8)
         MODOCA$ = "  "
         Call REPLA(TTXX$, MODOCA$, 129, 2)
         List3.AddItem TTXX$
         List3.ListIndex = ILISTA&
         '
         RFF$ = RECFILT$("DEPROFAB", 1, MKS$(REGA&))
         List1.Clear
         For U& = 1 To Len(RFF$) Step 4
           REFI& = CVS(Mid$(RFF$, U&, 4))
           XX$ = REGLEIDO$("DEPROFAB", REFI&)
           Call REPLA(XX$, MKS$(URE& + 1), 1, 4)
           Call REGAPP("DEPROFAB", XX$)
           '
           CIXI% = CVI(Mid$(XX$, 5, 2))
           CAXX = CVS(Mid$(XX$, 7, 4))
           TTXX$ = Space$(64)
           Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 22)
           Call REPLA(TTXX$, FORMATNUM$(CAXX, "F10.1"), 39, 10)
           List1.AddItem TTXX$
         Next U&
         Call CIERRARCH("DEPROFAB")
         Screen.MousePointer = SCMA
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If List3.ListIndex >= 0 Then
      If COMALTER%("Está Seguro de Anular el Plan Activo ?\\No, Conservarlo\Si, Anularlo") = 2 Then
        REGA& = Val(Mid$(List3.TEXT, 121, 8))
        XX$ = REGLEIDO$("PROFABRI", REGA&)
        Call REPLA(XX$, Space$(48), 5, 48)
        Call GRAREG("PROFABRI", XX$, REGA&)
        '
        '\\\OT-5-0537-FG-12/08/05///
        ' ANULACION DEL DETALLE DEL PROGRAMA
        NUMPROG& = CVS(Mid$(XX$, 1, 4))
        FIN& = ULTREG&("DEPROFAB")
        JJ& = 0
        For U& = 1 To FIN&
          YY$ = REGLEIDO$("DEPROFAB", U&)
          NPRO& = CVS(Mid$(YY$, 1, 4))
          If NPRO& <> NUMPROG& Then
            JJ& = JJ& + 1
            Call GRAREG("DEPROFAB", YY$, JJ&)
          End If
        Next U&
        '
        Call SETULTREG("DEPROFAB", JJ&)
        '\\\OT-5-0537-FG-FIN///
        '
        List3.Clear
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("PROFABRI")
          XX$ = REGLEIDO$("PROFABRI", U&)
          DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
          MODOCA$ = Mid$(XX$, 63, 2)
          If DENOPRO$ <> "" Then
            TTXX$ = Space$(128)
            Call REPLA(TTXX$, DENOPRO$, 1, 48)
            Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
            Call REPLA(TTXX$, MODOCA$, 129, 2)
            List3.AddItem TTXX$
          End If
        Next U&
        If List3.ListCount > 0 Then
          List3.ListIndex = List3.ListCount - 1
        End If
        Screen.MousePointer = 1
      End If
   End If
   Command24.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   AMASTO01.Show 1
   Command3.Enabled = True
End Sub


Private Sub Command4_Click()
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 1, 6)
     BBB$ = Mid$(List4.List(KIL&), 1, 6)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command5_Click()
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 9, 15)
     BBB$ = Mid$(List4.List(KIL&), 9, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command6_Click()
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 24, 12)
     BBB$ = Mid$(List4.List(KIL&), 24, 12)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CALCUCAP
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List5.ListCount - 1
     AAA$ = Mid$(List5.List(KIL& - 1), 24, 12)
     BBB$ = Mid$(List5.List(KIL&), 24, 12)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Command9_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List5.ListCount - 1
     AAA$ = Mid$(List5.List(KIL& - 1), 6, 18)
     BBB$ = Mid$(List5.List(KIL&), 6, 18)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List5.List(KIL& - 1)
          List5.RemoveItem (KIL& - 1)
          List5.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
End Sub

Private Sub Form_Load()
   '
   'Call CENTRAFORM(Me)
   'Show
   'Refresh
   ' VERANTER$ = "PROMRP01"
   Call VERJOBID(OKEY%)

   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   OPENFORMS = DoEvents
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("PROFABRI")
     XX$ = REGLEIDO$("PROFABRI", U&)
     DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
     MODOCA$ = Mid$(XX$, 63, 2)
     FECA% = CVI(Mid$(XX$, 65, 2))
     HOCA% = CVI(Mid$(XX$, 67, 2))
     If DENOPRO$ <> "" Then
       TTXX$ = Space$(130)
       Call REPLA(TTXX$, DENOPRO$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
       Call REPLA(TTXX$, MODOCA$, 129, 2)
       Call REPLA(TTXX$, FECHATEX$(FECA%), 131, 8)
       Call REPLA(TTXX$, HORATEX$(HOCA%), 139, 8)
       List3.AddItem TTXX$
     End If
   Next U&
   If List3.ListCount > 0 Then
     List3.ListIndex = List3.ListCount - 1
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Sub CALCUCAP()
   '
   ' calculo de capacidad
   '
   Screen.MousePointer = 11
   ILIS3& = List3.ListIndex
   CASUBCON% = 1
   If Val(Mid$(List3.TEXT, 129, 1)) = 2 Then ' POR ORDENES DE FACRICACION CERRADAS
     CASUBCON% = 0
   End If
   '
   Dim COEQUI$(1024), TIEQUI(1024)
   CANEQ% = 0
   Dim COCATO%(1024), TICATO(1024)
   CANCA% = 0
   '
   REBLA$ = REGBLAN$("CAPANECE")
   JJ& = 0
   FIP& = List1.ListCount
   For I& = 1 To FIP&
     Screen.MousePointer = 11
     Call TRACE(20, I&, FIP&)
     Label13 = TRIM$(Str$(I&))
     Label13.Refresh
     List1.ListIndex = I& - 1
     P00$ = List1.List(I& - 1)
     P1$ = TRIM$(Left$(P00$, 16))
     CIKI% = CODINT(P1$)
     '
     If CIKI% > 0 Then
       If CIKI% <= NUMAS% Then
         CANPROG = Val(Mid$(P00$, 39, 10))
         Call EXPLOMUL(CIKI%)
         Screen.MousePointer = 11
         '
         FIN& = CONTAMUL%
         If CASUBCON% < 1 Then
           FIN& = 1
         End If
         '
         URE& = ULTREG&("SECOPER")
         SUTIE = 0: SUMOD = 0: SUMAQ = 0
         ICX% = 0
         '
         For UI& = 1 To FIN&
           '
           CIXI% = CIK%(UI&)
           CODEX$ = CODEXT$(CIXI%)
           CACOM = CANT(UI&)
           '
           RFF$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
           IOP% = 0
           For UJ& = 1 To Len(RFF$) Step 4
             RXI& = CVS(Mid$(RFF$, UJ&, 4))
             If RXI& > 0 Then
               If RXI& <= URE& Then
                 SCP$ = REGLEIDO$("SECOPER", RXI&)
                 NOPY% = CVI(Mid$(SCP$, 3, 2))
                 DOPE$ = Mid$(SCP$, 5, 48)
                 LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
                 CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
                 MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
                 MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
                 MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
                 MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
                 MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
                 MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
                 HORAPROD = CANPROG * CACOM * MINLOTE / 60 / LOTESTA
                 '
                 CAOPS = CVS(Mid$(SCP$, 83, 4))
                 CATEGO% = Asc(Mid$(SCP$, 87, 1))
                 If CATEGO% < 1 Then CATEGO% = 1
                 '
                 If HORAPROD > 0.01 Then
                   EQUIPO$ = Mid$(SCP$, 53, 6) + Mid$(SCP$, 103, 18)
                   For KU% = 1 To 24 Step 6
                     EQUI$ = TRIM$(Mid$(EQUIPO$, KU%, 6))
                     If EQUI$ <> "" Then
                       RECAP$ = REBLA$
                       Call REPLA(RECAP$, EQUI$, 1, 6)
                       Call REPLA(RECAP$, MKI$(CIXI%), 7, 2)
                       Call REPLA(RECAP$, MKI$(NOPY%), 9, 2)
                       Call REPLA(RECAP$, DOPE$, 11, 48)
                       Call REPLA(RECAP$, MKS$(CANPROG * CACOM), 59, 4)
                       Call REPLA(RECAP$, MKI$(CIKI%), 67, 2)
                       Call REPLA(RECAP$, MKS$(HORAPROD), 69, 4)
                       Call REPLA(RECAP$, MKS$(HORAPROD), 77, 4)
                       JJ& = JJ& + 1
                       Call GRAREG("CAPANECE", RECAP$, JJ&)
                       '
                       For TUI% = 1 To CANEQ%
                         If COEQUI$(TUI%) = EQUI$ Then GoTo 20
                       Next TUI%
                       CANEQ% = CANEQ% + 1
                       COEQUI$(CANEQ%) = EQUI$
                       TIEQUI(CANEQ%) = 0
                       TUI% = CANEQ%
                       '
20                     TIEQUI(TUI%) = TIEQUI(TUI%) + HORAPROD
                       '
                     End If
                   Next KU%
                   '
                   For TUI% = 1 To CANCA%
                     If COCATO%(TUI%) = CATEGO% Then GoTo 30
                   Next TUI%
                   CANCA% = CANCA% + 1
                   COCATO%(CANCA%) = CATEGO%
                   TICATO(CANCA%) = 0
                   TUI% = CANCA%
                   '
30                 TICATO(TUI%) = TICATO(TUI%) + HORAPROD * CAOPS
                   '
                 End If
               End If
             End If
           Next UJ&
         Next UI&
       End If
     End If
   Next I&
   List1.ListIndex = (-1)
   '
   Call SETULTREG("CAPANECE", JJ&)
   Call FILESORT("CAPANECE", "", 1, 3, "ASC")
   '
   'JJ& = 0
   'For TU& = 1 To CANEQ%
   '  If TRIM$(COEQUI$(TU&)) <> "" Then
   '    If TIEQUI(TU&) > 0 Then
   '      REBLA$ = REGBLAN$("RCAPEQUI")
   '      Call REPLA(REBLA$, COEQUI$(TU&), 5, 6)
   '      Call REPLA(REBLA$, MKS$(TIEQUI(TU&)), 11, 4)
   '      JJ& = JJ& + 1
   '      Call GRAREG("!RCAPEQUI", REBLA$, JJ&)
   '    End If
   '  End If
   'Next TU&
   'Call SETULTREG("!RCAPEQUI", JJ&)
   '
   List4.Clear: SUTIE = 0: J4& = 0
   For TUI% = 1 To CANEQ%
     If TRIM$(COEQUI$(TUI%)) <> "" Then
       If TIEQUI(TUI%) > 0 Then
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, COEQUI$(TUI%), 1, 6)
         Call REPLA(TTXX$, ECOARCH$("PADMAQ", COEQUI$(TUI%)), 9, 15)
           SUTIE = SUTIE + TIEQUI(TUI%)
           HMS$ = HORMINSE$(3600 * TIEQUI(TUI%))
           While Len(HMS$) < 12: HMS$ = " " + HMS$: Wend
         Call REPLA(TTXX$, HMS, 24, 12)
         List4.AddItem (TTXX$)
         J4& = J4& + 1
         Call GRAREG("!RECALMAQ", MKS$(IPROGRA&) + TTXX$, J4&)
       End If
     End If
   Next TUI%
   List4.Refresh
   Label5 = HORMINSE$(3600 * SUTIE)
   Call SETULTREG("!RECALMAQ", J4&)
   '
   List5.Clear: SUTIE = 0: J5& = 0
   For TUI% = 1 To CANCA%
     If TICATO(TUI%) > 0 Then
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, AJUSTD$(Str$(COCATO%(TUI%)), 3), 1, 3)
       Call REPLA(TTXX$, ECOARCH$("CATOPER", Chr$(COCATO%(TUI%))), 6, 18)
       SUTIE = SUTIE + TICATO(TUI%)
           HMS$ = HORMINSE$(3600 * TICATO(TUI%))
           While Len(HMS$) < 12: HMS$ = " " + HMS$: Wend
       Call REPLA(TTXX$, HMS, 24, 12)
       List5.AddItem (TTXX$)
       J5& = J5& + 1
       Call GRAREG("!RECALMOD", MKS$(IPROGRA&) + TTXX$, J5&)
     End If
   Next TUI%
   List5.Refresh
   Label10 = HORMINSE$(3600 * SUTIE)
   Call SETULTREG("!RECALMOD", J5&)
   '
   Call TRACENFIL("RECALMAQ", 1, MKS$(IPROGRA&))
   Call TRACENFIL("RECALMOD", 1, MKS$(IPROGRA&))
   '
   Call COPYSTRU("CAPANECE", "CAPAEQPI")
   Call COPYSTRU("CAPANECE", "CAPAEQUI")
   J1& = 0: EQPIA$ = "": SUCAN1 = 0: SUHOR1 = 0
   J2& = 0: EQUIPA$ = "": SUCAN2 = 0: SUHOR2 = 0
   '
   FIN& = ULTREG&("CAPANECE")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("CAPANECE", U&)
     CANX = CVS(Mid$(XX$, 59, 4))
     HORX = CVS(Mid$(XX$, 69, 4))
     If Left$(XX$, 8) <> EQPIA$ Then
       If EQPIA$ <> "" Then
         RECAP$ = URELE$
         Call REPLA(RECAP$, MKI$(0), 9, 2)
         Call REPLA(RECAP$, Space$(48), 11, 48)
         Call REPLA(RECAP$, MKS$(SUCAN1), 59, 4)
         Call REPLA(RECAP$, MKI$(0), 67, 2)
         Call REPLA(RECAP$, MKS$(SUHOR1), 69, 4)
         Call REPLA(RECAP$, MKS$(SUHOR1), 77, 4)
         J1& = J1& + 1
         Call GRAREG("CAPAEQPI", RECAP$, J1&)
         SUCAN1 = 0: SUHOR1 = 0
       End If
     End If
     '
     If Left$(XX$, 6) <> EQUIPA$ Then
       If EQUIPA$ <> "" Then
         RECAP$ = URELE$
         Call REPLA(RECAP$, MKI$(0), 7, 2)
         Call REPLA(RECAP$, MKI$(0), 9, 2)
         Call REPLA(RECAP$, Space$(48), 11, 48)
         Call REPLA(RECAP$, MKS$(SUCAN2), 59, 4)
         Call REPLA(RECAP$, MKI$(0), 67, 2)
         Call REPLA(RECAP$, MKS$(SUHOR2), 69, 4)
         Call REPLA(RECAP$, MKS$(SUHOR2), 77, 4)
         J2& = J2& + 1
         Call GRAREG("CAPAEQUI", RECAP$, J2&)
         SUCAN2 = 0: SUHOR2 = 0
       End If
     End If
     URELE$ = XX$
     EQPIA$ = Left$(URELE$, 8)
     EQUIPA$ = Left$(URELE$, 6)
     SUCAN1 = SUCAN1 + CANX: SUCAN2 = SUCAN2 + CANX
     SUHOR1 = SUHOR1 + HORX: SUHOR2 = SUHOR2 + HORX
     '
   Next U&
   '
   RECAP$ = URELE$
   Call REPLA(RECAP$, MKI$(0), 9, 2)
   Call REPLA(RECAP$, Space$(48), 11, 48)
   Call REPLA(RECAP$, MKS$(SUCAN1), 59, 4)
   Call REPLA(RECAP$, MKI$(0), 67, 2)
   Call REPLA(RECAP$, MKS$(SUHOR1), 69, 4)
   Call REPLA(RECAP$, MKS$(SUHOR1), 77, 4)
   J1& = J1& + 1
   Call GRAREG("CAPAEQPI", RECAP$, J1&)
   '
   RECAP$ = URELE$
   Call REPLA(RECAP$, MKI$(0), 7, 2)
   Call REPLA(RECAP$, MKI$(0), 9, 2)
   Call REPLA(RECAP$, Space$(48), 11, 48)
   Call REPLA(RECAP$, MKS$(SUCAN2), 59, 4)
   Call REPLA(RECAP$, MKI$(0), 67, 2)
   Call REPLA(RECAP$, MKS$(SUHOR2), 69, 4)
   Call REPLA(RECAP$, MKS$(SUHOR2), 77, 4)
   J2& = J2& + 1
   Call GRAREG("CAPAEQUI", RECAP$, J2&)
   '
   Call SETULTREG("CAPAEQPI", J1&)
   Call SETULTREG("CAPAEQUI", J2&)
   '
   TXKX$ = FILTERGETRECORD$("PROFABRI", 1, MKS$(IPROGRA&))
   HOII% = HOY(HOS$)
   Call REPLA(TXKX$, MKI$(HOII%), 65, 5)
   HORI% = HORANUM%(Format$(Time, "hh:mm"))
   Call REPLA(TXKX$, MKI$(HORI%), 67, 5)
   Call FILTERPUTRECORD("PROFABRI", 1, MKS$(IPROGRA&), TXKX$)
   '
   ILIS3& = List3.ListIndex
   TTXX$ = Left$(List3.TEXT + Space$(16), 146)
   List3.RemoveItem ILIS3&
   Call REPLA(TTXX$, FECHATEX$(HOII%), 131, 8)
   Call REPLA(TTXX$, HORATEX$(HORI%), 139, 8)
   List3.AddItem TTXX$, ILIS3&
   '
   Label15 = FECHATEX$(HOII%)
   Label17 = HORATEX$(HORI%)
   On Error Resume Next
   List3.ListIndex = ILIS3&
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label1_DblClick()
   Call CARPROFAB
End Sub

Sub CARPROFAB()
   DEPROGRA$ = List3.TEXT
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   TIPROGRA% = Val(Mid$(DEPROGRA, 129, 1))
   If TIPROGRA% > 0 Then
     Call MENSERR(24, "Tipo de Programa Importado\no Puede ser Editado")
     Exit Sub
   End If
   '
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        Call TRALOFIL("DEPROFAB", 1, MKS$(IPROGRA&))
        VTB% = VENTABU%("CARPROFA")
        If VTB% >= 0 Then
          Call TRACENFIL("DEPROFAB", 1, MKS$(IPROGRA&))
          Call LISPROFAB
        End If
   End If
End Sub
'
Sub LISPROFAB()
   DEPROGRA$ = List3.TEXT
   Label15 = TRIM$(Mid$(DEPROGRA$, 131, 8))
   Label17 = TRIM$(Mid$(DEPROGRA$, 139, 8))
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        SCMA = Screen.MousePointer
        Screen.MousePointer = 11
        RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
        List1.Clear
        For U& = 1 To Len(RFF$) Step 4
          REFI& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("DEPROFAB", REFI&)
          CIXI% = CVI(Mid$(XX$, 5, 2))
          CAXX = CVS(Mid$(XX$, 7, 4))
          TTXX$ = Space$(64)
          Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
          Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 22)
          Call REPLA(TTXX$, FORMATNUM$(CAXX, "F10.1"), 39, 10)
          List1.AddItem TTXX$
        Next U&
        '
        Call TRALOFIL("RECALMAQ", 1, MKS$(IPROGRA&))
        Call TRALOFIL("RECALMOD", 1, MKS$(IPROGRA&))
        '
        List4.Clear: Label5 = "": TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMAQ")
          TTXX$ = Mid$(REGLEIDO$("!RECALMAQ", U&), 5)
          List4.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        List4.Refresh
        If TOTIX >= 1 Then
          Label5 = HORMINSE$(TOTIX)
        End If
        '
        List5.Clear: Label10 = "": TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMOD")
          TTXX$ = Mid$(REGLEIDO$("!RECALMOD", U&), 5)
          List5.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        List5.Refresh
        If TOTIX >= 1 Then
          Label10 = HORMINSE$(TOTIX)
        End If
        '
        Screen.MousePointer = SCMA
   End If
End Sub
'

Private Sub List1_DblClick()
   Call CARPROFAB
End Sub

Private Sub List3_Click()
   Call LISPROFAB
End Sub

Private Sub LIST3_DBLCLICK()
   Call CARPROFAB
End Sub



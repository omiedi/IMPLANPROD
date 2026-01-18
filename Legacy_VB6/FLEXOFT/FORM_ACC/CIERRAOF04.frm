VERSION 5.00
Begin VB.Form CIERRAOF04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Fabricación con Trazabilidad"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "# "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6615
      TabIndex        =   49
      Top             =   590
      Width           =   3060
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
         Left            =   630
         TabIndex        =   51
         Top             =   210
         Width           =   810
      End
      Begin VB.TextBox Text6 
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
         Left            =   2100
         TabIndex        =   50
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL:"
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
         Left            =   -750
         TabIndex        =   53
         Top             =   260
         Width           =   1335
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLC:"
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
         Left            =   700
         TabIndex        =   52
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   945
      TabIndex        =   43
      Top             =   630
      Width           =   435
   End
   Begin VB.CommandButton Command7 
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
      Left            =   1365
      TabIndex        =   42
      Top             =   630
      Width           =   175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   4095
      TabIndex        =   41
      Top             =   630
      Width           =   435
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
      Left            =   4515
      TabIndex        =   40
      Top             =   630
      Width           =   175
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
      Left            =   8715
      TabIndex        =   37
      Top             =   210
      Width           =   960
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
      Left            =   6195
      TabIndex        =   34
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de la O/F Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5950
      Left            =   105
      TabIndex        =   8
      Top             =   1365
      Width           =   9640
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Materiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4200
         Left            =   105
         TabIndex        =   26
         Top             =   1640
         Width           =   9360
         Begin VB.Frame Frame9 
            Caption         =   "Frame4"
            Height          =   4050
            Left            =   8700
            TabIndex        =   54
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame3"
            Height          =   4050
            Left            =   6510
            TabIndex        =   31
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame3"
            Height          =   4050
            Left            =   1680
            TabIndex        =   30
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   4050
            Left            =   7665
            TabIndex        =   29
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   4050
            Left            =   4620
            TabIndex        =   28
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   4050
            Left            =   5565
            TabIndex        =   27
            Top             =   120
            Width           =   15
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3840
            ItemData        =   "CIERRAOF04.frx":0000
            Left            =   0
            List            =   "CIERRAOF04.frx":0002
            TabIndex        =   32
            Top             =   350
            Width           =   9400
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Codigo                  Descripción                                  Uso Un.      Neces.       Trazado          Falta    Dp    "
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
            Left            =   0
            TabIndex        =   33
            Top             =   105
            Width           =   9400
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1410
         Left            =   50
         TabIndex        =   9
         Top             =   210
         Width           =   9465
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. O/F:"
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
            TabIndex        =   24
            Top             =   250
            Width           =   960
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1155
            TabIndex        =   23
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
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
            Left            =   2415
            TabIndex        =   22
            Top             =   250
            Width           =   960
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3465
            TabIndex        =   21
            Top             =   210
            Width           =   330
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Atraso:"
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
            Left            =   7560
            TabIndex        =   20
            Top             =   885
            Width           =   960
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   465
            Left            =   8610
            TabIndex        =   19
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitada:"
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
            Left            =   6405
            TabIndex        =   18
            Top             =   255
            Width           =   1905
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   8400
            TabIndex        =   17
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
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
            TabIndex        =   16
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1155
            TabIndex        =   15
            Top             =   525
            Width           =   2010
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1155
            TabIndex        =   14
            Top             =   840
            Width           =   5160
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
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
            Left            =   3675
            TabIndex        =   13
            Top             =   255
            Width           =   1485
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5250
            TabIndex        =   12
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo:"
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
            Left            =   3675
            TabIndex        =   11
            Top             =   525
            Width           =   1485
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5250
            TabIndex        =   10
            Top             =   525
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   5
      Top             =   210
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004000&
      Height          =   8460
      Left            =   9870
      ScaleHeight     =   8400
      ScaleWidth      =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
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
         Picture         =   "CIERRAOF04.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2320
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
         Height          =   640
         Left            =   105
         Picture         =   "CIERRAOF04.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Acceso a Trazabilidad de Reservas"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   105
         Picture         =   "CIERRAOF04.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Aprueba Cierre de Orden de Fabricación"
         Top             =   5355
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   6510
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CIERRAOF04.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CIERRAOF04.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CIERRAOF04.frx":0BB6
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6885
         Width           =   1515
      End
   End
   Begin VB.Label Label22 
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
      Left            =   945
      TabIndex        =   47
      Top             =   945
      Width           =   2445
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada:"
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
      Left            =   0
      TabIndex        =   46
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label20 
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
      Left            =   4095
      TabIndex        =   45
      Top             =   945
      Width           =   2310
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Consumo:"
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
      Left            =   2940
      TabIndex        =   44
      Top             =   690
      Width           =   1125
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   7560
      TabIndex        =   38
      Top             =   285
      Width           =   1125
   End
   Begin VB.Label Label24 
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
      Height          =   285
      Left            =   5250
      TabIndex        =   36
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Cierre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3570
      TabIndex        =   35
      Top             =   285
      Width           =   1650
   End
   Begin VB.Label Label117 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/F Nro:"
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
      Left            =   -420
      TabIndex        =   7
      Top             =   285
      Width           =   1335
   End
   Begin VB.Label Label144 
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
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   210
      Width           =   1770
   End
End
Attribute VB_Name = "CIERRAOF04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call TRAZORF04.Command10_Click
    Command10.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    If XVALO(Text3) = 99 Then
       If COMALTER%("Atención: Verifique Depósitos Predeterminados de Consumo !\\Continuar\Cancelar") <> 1 Then
          GoTo 99
       End If
    End If
    '
    Dim CODIN%(4096), LOTRE$(4096), FEALT%(4096), SALFIS(4096), PP1$(4096), PP2$(4096)
    Dim CANECE(32767)
    '
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% = 0 Or FECHA% > HOII% Then
       Call MENSERR(24, "Fecha No Válida")
       GoTo 99
    End If
     '
    If XVALO(Text1) < 1 Then
       Call MENSERR(24, "Imposible Cerrar - Falta Cantidad")
       GoTo 99
    End If
    '
    DEPO% = XVALO(Text2)
    If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Entrada")
       GoTo 99
    End If
    '
    DP2% = XVALO(Text3)
    If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Consumo")
       GoTo 99
    End If
    '
    For KKU% = 1 To List1.ListCount
      If XVALO(Mid$(List1.List(KKU% - 1), 72, 10)) >= 0.05 Then
        Call MENSERR(24, "Cierre de O/F Imposible.\   \Stock Trazado Insuficiente\en Código '" + TRIM$(Left$(List1.List(KKU% - 1), 16)) + "'.")
        GoTo 99
      End If
      DPXX1% = XVALO(Mid$(List1.List(KKU% - 1), 82, 5))
      If DPXX1% < 1 Or DPXX1% > 255 Then
        Call MENSERR(24, "Depósito de Consumo Inexistente o No Válido\en Código '" + TRIM$(Left$(List1.List(KKU% - 1), 16)) + "'.")
        GoTo 99
      End If
    Next KKU%
    '
    Call BLOQARCH("RESERVOF")
    ' VERIFICAR NEGATIVOS
    CAITMO% = 0
    For KKI% = 1 To NUMAS%
       CANECE(KKI%) = 0
    Next KKI%
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       CODE$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(CODE$)
       CACONSU = XVALO(Mid$(TTXX$, 52, 10))
       If CACONSU > 0.005 Then
          CAITMO% = CAITMO% + 1
          CIMOVIM%(CAITMO%) = CIXI%
          CAMOVIM(CAITMO) = (-1) * CACONSU
          CANECE(CIXI%) = CANECE(CIXI%) + CACONSU
       End If
    Next KKI%
    '
    If PUEDEMOVI%(DP2%) < 0 Then
      Call COMUNI("Transacción Cancelada por Control de Negativos.")
      GoTo 99
    End If
    '
    ' CARGAR UN VECTOR CON LOS LOTES DISPONIBLES
    NORFA$ = Label2
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CantRes - CantCons > 0 "
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
       On Error Resume Next
       .MoveLast
       If Err Then
          OPXX% = COMALTER%("No Existen Reservas Vigentes\para la O/F Elegida.\  \¿ Realmente Confirma el Cierre\Sin Consumo de Materiales Componentes ?\\No, Cancelar Cierre\Si Confirmarlo")
          If OPXX% = 2 Then GoTo 20   ' CONFIRMA CIERRE
          GoTo 99
       End If
       FIN& = .RecordCount
       .MoveFirst
       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
          K& = .AbsolutePosition
          Call TRACE(20, K&, FIN&)
          '
          CIXI% = !Cod_Inte
          IDLO$ = TRIM$(!IdenLote)
          If IDLO$ <> "" Then
            CANRES = !CantRes - !CantCons
            JJ& = JJ& + 1
            CODIN%(JJ&) = CIXI%
            LOTRE$(JJ&) = IDLO$
              Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
            FEALT%(JJ&) = FECREALOT%(IDLO$)
              SALNOCO = SALNOCONS(IDLO$, CIXI%)
              If CANRES > SALNOCO Then
                 Call MENSERR(24, "Error de Trazado en Código '" + CODEXT$(CIXI%) + "'.\Imposible Cerrar esta O/F.\   \Vuelva a Trazar las Reservas y Re-Intente.")
                 GoTo 99
              End If
            SALFIS(JJ&) = CANRES
            PP1$(JJ&) = !PNL_DES
            PP2$(JJ&) = !PLC_HAS
            CANECE(CIXI%) = CANECE(CIXI%) - SALNOCO
          End If
19     .MoveNext
       Loop
    End With
    '
20  Screen.MousePointer = 1
    '
    For KKI% = 1 To NUMAS%
       If CANECE(KKI%) >= 0.05 Then
          Call MENSERR(24, "Cierre de O/F Imposible.\Error de Lotes Trazados en\Código '" + CODEXT$(KKI%) + "'.\   \Revise Lotes y Situación de Stocks.")
          GoTo 99
       End If
    Next KKI%
    '
    ' ordenar por fecha de alta
    For U1& = 1 To JJ&
      For U2& = U1& + 1 To JJ&
        If FEALT%(U1&) > FEALT%(U2&) Then
          SW = SWAP(FEALT%(U1&), FEALT%(U2&))
          SW = SWAP(CODIN%(U1&), CODIN%(U2&))
          SW = SWAP(LOTRE$(U1&), LOTRE$(U2&))
          SW = SWAP(SALFIS(U1&), SALFIS(U2&))
          SW = SWAP(PP1$(U1&), PP1$(U2&))
          SW = SWAP(PP2$(U1&), PP2$(U2&))
        End If
      Next U2&
    Next U1&
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("RESERVA")
    '
    CMOV$ = "CF"
    CODICON$ = Label10
    DOPRI$ = TRIM$(Label2)
    DOSEC$ = "PF-" + TRIM$(Mid$(Label2, 4))
    REFE$ = "CONSUMO " + CODICON$
    NORIT% = 0
    '
    ' GRABAR MOVIMIENTO DE ENTRADA
    CICON% = CODINT%(CODICON$)
    CODCON$ = CODEXT$(CICON%)
    ILOTE$ = Label2
       PDES$ = Text5
       PHAS$ = Text6
    CMOV$ = "OF"
    DOPRI$ = Label2
    DOSEC$ = "IF-" + TRIM$(Mid$(DOPRI$, 4))
    REFE$ = "Fabricación Propia " + CODICON$
    VUNI# = 0
    MONEX% = 1
    CANPRO = XVALO(Text1)
    GoSub GRAPRODUC
    '
    ' GRABAR CONSUMO Y CANCELA RESERVA
    SQLT$ = "SELECT * FROM MOVISTO "
    SQLT$ = SQLT$ + " WHERE CodiMovi = '" & CMOV$ & "' "
    SQLT$ = SQLT$ + " AND DoPriCor = '" & DOPRI$ & "' "
    'SQLT$ = SQLT$ + " AND IdenMovi = '" & IDMO$ & "' "
    Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenDynaset)
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       CODE$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(CODE$)
       CACONSU = XVALO(Mid$(TTXX$, 52, 10))
       DPXX1% = XVALO(Mid$(TTXX$, 82, 5))
       '
       Do While CACONSU >= 0.05
          HUBOCON% = 0
          For KKJ& = 1 To JJ&
            If CODIN%(KKJ&) = CIXI% Then
               CACON = CACONSU
               If CACON > SALFIS(KKJ&) Then CACON = SALFIS(KKJ&)
               If CACON > 0 Then
                  LOTE$ = LOTRE$(KKJ&)
                  CANTIX = CACON
                      PCOM$ = PP1$(KKJ&)
                      PFIN$ = PP2$(KKJ&)
                  NORIT% = NORIT% + 1
                  CMOV$ = "CF"
                  GoSub GRACONSUM
                  GoSub CANCELARES
                  Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
                  SALFIS(KKJ&) = SALFIS(KKJ&) - CACON
                  CACONSU = CACONSU - CACON
                  HUBOCON% = 1
               End If
            End If
          Next KKJ&
          If HUBOCON% < 1 Then
             Call MENSERR(24, "Imposible Completar Cierre Parcial\de O/F '" + NORFA$ + "'")
             GoTo 99
          End If
       Loop
       '
    Next KKI%
    '
    ' ACTUALIZAR PENDIENTE DE O/F
    GoSub ACUPENOF
    If CIERRA% = 1 Then Call ANULARES(NORFA$)
    Label144 = ""
    GoTo 99

'*********************************************************************************
    '
ACUPENOF:
    ACUPRODU = 0
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF' "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CICON% & " "
    Set ORYACIE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With ORYACIE
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            ACUPRODU = ACUPRODU + !CantIngre
         .MoveNext
         Loop
      End If
    End With
    '
    CIERRA% = 0
    If ACUPRODU >= XVALO(Label13) - 0.05 Then
        CIERRA% = 1
      Else
        MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
        If MDXX% = 2 Then CIERRA% = 1
    End If
    '
    Call APERBASDAT("ORFAB")
    With OrdeFabr
       .FindFirst "IdSubOr = '" & NORFA$ & "'"
       If .NoMatch = False Then
         .Edit
         !CanApro = ACUPRODU
         !CanRech = !CanRech + SCRAP
         If CIERRA% = 1 Then
             !Statuorf = 3         ' MARCA COMO CERRADA
           Else
             If !Statuorf > 2 Then !Statuorf = 1
         End If
         !FechCierr = CVDAT(FECHANUM(Label24))
         .Update
       End If
    End With
    '
    '25/04/07 (OT-07-0161)'PARA GUARDAR EL DEPOSITO DE ENTRADA Y EL DEPOSITO
    If CICON% > 0 Then    'DE CONSUMO EN EL MAESTRO DE ARTICULOS
        Master.FindFirst ("CODINT = " & CICON%)
        If Master.NoMatch = False Then
           Master.Edit
           Master!DepoEntra = XVALO(Text2)
           Master!DepoCon = XVALO(Text3)
           Master.Update
        End If
    End If
    '
Return
    '
GRAPRODUC:
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = 'OF') "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    'SQLX$ = SQLX$ + " AND IdenMovi = '" & IDMO$ & "'"
    Set PRODORFA = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With PRODORFA
      .AddNew
      !Fechmov = CVDAT(FECHA%)
      !Cod_Inte = CICON%
      !COD_EXTE = CODCON$
      !DESCRIP = DESCRIT0$(CICON%)
      !IdenLote = TRIM$(ILOTE$)
      !PNL_DES = PDES$
      !PLC_HAS = PHAS$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
      !NuOrItem = 0
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = VUNI#
      !MoneVal = MONEX%
      !TipoCam = TICAMONE(MONEX%)
         DEPOI% = DEPO%: If DEPO% = 99 Then DEPOI% = DEPOPRE%(CICON%)
      !DepoMovi = DEPOI%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = CANPRO
      !Cantsalid = 0
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
    End With
    '
    Call CREALOTIN(FECHA%, CICON%, ILOTE$, CDbl(CANPRO))
    '
Return
'
'
GRACONSUM:
    With MACONSUM
      .AddNew
      !Fechmov = CVDAT(FECHA%)
      !Cod_Inte = CIXI%
      !COD_EXTE = CODE$
      !DESCRIP = DESCRIT0$(CIXI%)
      !IdenLote = TRIM$(LOTE$)
      !PNL_DES = PCOM$
      !PLC_HAS = PFIN$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
         LOTECON$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(CANTIX, "F8.1")
      !LotConsum = LOTECON$
      !NuOrItem = NORIT%
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = 0
      !MoneVal = 1
      !TipoCam = TICAMONE(1)
         DEPOI% = DPXX1%: If DPXX1% = 99 Then DEPOI% = DEPOPRE%(CIXI%)
      !DepoMovi = DEPOI%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = 0
      !Cantsalid = Abs(CANTIX)
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
      '
    End With
Return
'
CANCELARES:
    CANTLIM = CANTIX - 0.05
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + "AND IdenLote = '" & LOTE$ & "' "
    SQLX$ = SQLX$ + " AND CantRes - CantCons >= " & TRIM$(FORMATNUM$(CANTLIM, "F12.2"))
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
          Call MENSERR(24, "No Existen Reservas Vigentes\para la O/F Elegida. ")
          GoTo 99
       End If
       .Edit
       !CantRes = !CantRes - CANTIX
       .Update
       If !CantRes < 0.05 Then
          .Delete
       End If
    End With
    '
Return
'
'
99 Call LIBARCH("RESERVOF")
   Command2.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
   '
   Call ACTECOR(2)
   NROOF$ = ""
   If ULTREG&("ECORDEP") < 1 Then
     Call MENSERR(24, "No Hay Ordenes de Fabricación en Curso\para el Código Seleccionado")
     GoTo 90
   End If
   '
   Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
   NROOF$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
   '
90 Label144 = NROOF$
   Screen.MousePointer = 1
   '
End Sub


Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = TRIM$(VALACT1$("SELFECHA"))
       If VDEV$ <> "" Then
       Label24 = VDEV$
    End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   TRAZORF04.Label32 = Label144
   TRAZORF04.Show 1
   Call Text1_Change
   Command5.Enabled = True
End Sub

Private Sub Command7_Click()
    Call CommSel("TADEPOSI", Command7, Text2)
End Sub

Private Sub Command9_Click()
    Call CommSel("ECODEP", Command9, Text3)
End Sub

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  'AGREGUE15/09/06MH
  '
  Screen.MousePointer = 1
  '
End Sub

Private Sub Label144_Change()
    '
    Label2 = Label144
    '
End Sub


Private Sub Label15_Change()
   Text1 = Label15
End Sub

Private Sub Label2_Change()
   Screen.MousePointer = 11
   NORFA$ = Label2
   Label4.Caption = ""
   Label6.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label11.Caption = ""
   Label13.Caption = ""
   Label15.Caption = ""
   List1.Clear
   '
   Call ABREORFAB
   '
   SQLX$ = " SELECT * FROM Ordefabr "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With ORFATEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      '
      CIIX% = !Cod_Inte
      Label10.Caption = CODEXT$(CIIX%)
      Label11.Caption = DESCRIT$(CIIX%)
      Label13.Caption = FORMATNUM$(!CanProy, "F9.1")
        FESOLENT = CVINT(!FechEnSol)
      Label8.Caption = FECHATEX$(FESOLENT)
        DIATRA& = DIENTRE(FESOLENT, FECALC): If DIATRA < 1 Then DIATRA = 0
      Label6 = TRIM$(Str$(DIATRA&))
      Label4.Caption = !Statuorf
        SALDORF = !CanProy - !CanApro: If SALDORF < 0 Then SALDORF = 0
      Label15.Caption = FORMATNUM$(SALDORF, "F9.1")
   End With
   '
   HOII% = HOY(HOS$)
   Label24 = HOS$
   Text2 = Str(DEPENART%(CIIX%))  'PARA MOSTRAR EL DEPOSITO DE ENTRADA Y DE CONSUMO
   Text3 = Str(DEPCONART%(CIIX%)) 'ASIGNADO AL ARTICULO
   '
   Call CARGACOTRA
   '
End Sub

Sub CARGACOTRA()
    '
    Static ULTIORFA$
    '
    Dim CODIN%(2048), CATRAZ(2048), USOUNI(2048)
    '
    CANFA = XVALO(Text1)
    NORFA$ = TRIM$(Label2)
    If NORFA$ = "" Then Exit Sub
    '
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CantRes - CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY NuOrItem ASC"
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveLast
      If Err Then
         If Label144 <> ULTIORFA$ Then     ' PARA QUE NO REPITA EL MENSAJE
            Call MENSERR(24, "No Existen Reservas Vigentes\para la O/F Elegida. ")
            ULTIORFA$ = Label144
         End If
         GoTo 99
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        K& = .AbsolutePosition
        Call TRACE(20, K&, FIN&)
        '
        CIXI% = !Cod_Inte
        IDLO$ = TRIM$(!IdenLote)
        CANRES = !CantRes - !CantCons
        CANTRA = 0
        If IDLO$ <> "" Then
            CANTRA = CANRES
        End If
        '
        For JJ& = 1 To CANMO&
           If CODIN%(JJ&) = CIXI% Then
              CATRAZ(JJ&) = CATRAZ(JJ&) + CANTRA
              GoTo 19
           End If
        Next JJ&
        CANMO& = CANMO& + 1
        CODIN%(CANMO&) = CIXI%
        CATRAZ(JJ&) = CANTRA
        USOUNI(JJ&) = !UsoUnit
        '
19    .MoveNext
      Loop
    End With
    '
    List1.Clear
    For JJ& = 1 To CANMO&
       TTXX$ = Space$(128)
       CIXI% = CODIN%(JJ&)
       Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
       Call REPLA(TTXX$, DESCRIT$(CIXI%), 17, 25)
       Call REPLA(TTXX$, FORMATNUM(USOUNI(JJ&), "F10.4"), 42, 10)
         CANCO = USOUNI(JJ&) * CANFA
       Call REPLA(TTXX$, FORMATNUM(CANCO, "F10.1"), 52, 10)
         CANTRA = CATRAZ(JJ)
       Call REPLA(TTXX$, FORMATNUM(CANTRA, "F10.1"), 62, 10)
         FALTA = CANCO - CATRAZ(JJ): If FALTA < 0 Then FALTA = 0
       Call REPLA(TTXX$, FORMATNUM(FALTA, "F10.1"), 72, 10)
       DPXX% = XVALO(Text3)
       If XVALO(Text3) = 99 Then DPXX% = DEPOPRE%(CIXI%)
       Call REPLA(TTXX$, FORMATNUM$(DPXX%, "F5.0"), 82, 5)
       List1.AddItem TTXX$
    Next JJ&
99 End Sub
'
Private Sub Label24_Change()
   FEX1 = FECHANUM(Label8)
   FEX2 = FECHANUM(Label24)
   ATRASO = DIENTRE(FEX1, FEX2)
   Label6 = FORMATNUM$(ATRASO, "I3")
End Sub

Private Sub Label24_DblClick()
   Call Command4_Click
End Sub

Private Sub Label8_Change()
   Call Label24_Change
End Sub
'
Private Sub Text1_Change()
   Call CARGACOTRA
End Sub

Private Sub Text2_Change()
  DEPO% = XVALO(Text2)
  Label22.Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
End Sub

Private Sub Text2_DblClick()
  Call Command7_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Private Sub Text3_Change()
  DEPO% = XVALO(Text3)
  Label20.Caption = TRIM$(ECOARCH$("ECODEP", Chr$(DEPO%)))
  Call CARGACOTRA
End Sub

Private Sub Text3_DblClick()
  Call Command9_Click
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       List1.SetFocus
    End If
End Sub



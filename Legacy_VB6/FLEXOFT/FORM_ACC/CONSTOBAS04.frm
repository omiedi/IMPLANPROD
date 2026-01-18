VERSION 5.00
Begin VB.Form CONSTOBAS04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   6270
   ClientLeft      =   1050
   ClientTop       =   1500
   ClientWidth     =   9210
   ClipControls    =   0   'False
   Icon            =   "CONSTOBAS04.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Comprometido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2835
      TabIndex        =   36
      Top             =   1470
      Width           =   2535
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   38
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label203 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1150
         TabIndex        =   37
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "A Ingresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5565
      TabIndex        =   33
      Top             =   1470
      Width           =   2535
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   35
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label303 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   34
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   30
      Top             =   1470
      Width           =   2430
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   195
         TabIndex        =   32
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   31
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   6210
      Left            =   8300
      ScaleHeight     =   6150
      ScaleWidth      =   1530
      TabIndex        =   21
      Top             =   0
      Width           =   1590
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   5400
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "List"
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
         Picture         =   "CONSTOBAS04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Acceso a Consultas y Listados"
         Top             =   3420
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move"
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
         Picture         =   "CONSTOBAS04.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Realizar Movimientos de Stock"
         Top             =   4110
         Width           =   650
      End
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Kardx"
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
         Picture         =   "CONSTOBAS04.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2730
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
         Picture         =   "CONSTOBAS04.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Picture         =   "CONSTOBAS04.frx":0832
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Setup"
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
         Picture         =   "CONSTOBAS04.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1680
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONSTOBAS04.frx":0F7E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5700
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Situación de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   210
      TabIndex        =   10
      Top             =   2625
      Width           =   7890
      Begin VB.Label Label401 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   16
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Neto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2625
         TabIndex        =   15
         Top             =   375
         Width           =   1380
      End
      Begin VB.Label Label402 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   14
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   13
         Top             =   375
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faltantes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5565
         TabIndex        =   12
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label403 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6510
         TabIndex        =   11
         Top             =   660
         Width           =   1170
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
      Height          =   1860
      Left            =   210
      TabIndex        =   5
      Top             =   4260
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2730
      TabIndex        =   0
      Top             =   210
      Width           =   2430
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
      Height          =   360
      Left            =   7455
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5160
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5145
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.ListBox LIUBIC 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   4200
      TabIndex        =   2
      Top             =   4260
      Width           =   3900
   End
   Begin VB.Label Label301 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6615
      TabIndex        =   50
      Top             =   6510
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Compras:"
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
      Left            =   5670
      TabIndex        =   49
      Top             =   6615
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label302 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6615
      TabIndex        =   48
      Top             =   6930
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricación:"
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
      TabIndex        =   47
      Top             =   7035
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asignados y Reservas:"
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
      Left            =   420
      TabIndex        =   46
      Top             =   7035
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label202 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   45
      Top             =   7140
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos:"
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
      Left            =   630
      TabIndex        =   44
      Top             =   6825
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label201 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   43
      Top             =   6720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Line Line1 
      X1              =   8100
      X2              =   8100
      Y1              =   3960
      Y2              =   6120
   End
   Begin VB.Line Line2 
      X1              =   4005
      X2              =   4005
      Y1              =   3960
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   210
      X2              =   3990
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   8100
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Economico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5670
      TabIndex        =   20
      Top             =   1075
      Width           =   1275
   End
   Begin VB.Label LECONOMI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   7035
      TabIndex        =   19
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label STOMI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   7035
      TabIndex        =   18
      Top             =   735
      Width           =   1065
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mínimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5670
      TabIndex        =   17
      Top             =   760
      Width           =   1275
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ubicaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   3960
      Width           =   3900
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   8
      Top             =   3960
      Width           =   3795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   420
      TabIndex        =   7
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6510
      TabIndex        =   6
      Top             =   315
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remoto:"
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
      Left            =   2940
      TabIndex        =   42
      Top             =   7140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label102 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3885
      TabIndex        =   41
      Top             =   7035
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Propio:"
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
      Left            =   2940
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label101 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3885
      TabIndex        =   39
      Top             =   6615
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "CONSTOBAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%
Dim MOCONPED%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    ASEL$ = "MASTER"
    '
    Call SELECHO(ASEL$)
    CODD$ = VALACT1$(ASEL$)
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub

Private Sub Command27_Click()
   Command27.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    COMMAND3.Enabled = False
    FIKARDEX.Text1 = Text1
    FIKARDEX.Show 1
    COMMAND3.Enabled = True
End Sub



Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   MOVISTBAS04.Show 1
   If CODINT%(Text1.TEXT) > 0 Then
      COD$ = TRIM$(Text1.TEXT)
      CI% = CODINT%(COD$)
      Text2.TEXT = UNIMED$(CI%)
      Text3.TEXT = DESCRIT$(CI%)
      Call MUEDIS(CI%)                ' recalcular existencias
   End If
   Command6.Enabled = True
End Sub


Private Sub Command9_Click()
    Command9.Enabled = False
    LISTOKBAS04.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_Activate()
   Call GRATITFOR(Caption)
   WindowState = 0
   OPENFORMS = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    '
    MOCONPED% = 1
    '
End Sub
'


Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Label13_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
End Sub

Private Sub Label202_DblClick()
   CONTROCLIKK% = 1
   TOPIN = List2.TopIndex
   For U& = 1 To List2.ListCount
     List2.Selected(U& - 1) = False
     If Left$(List2.List(U& - 1), 2) = "OD" Then
       List2.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   List2.TopIndex = TOPIN
   On Error GoTo 0
   CONTROCLIKK% = 0
End Sub


Private Sub LIUBIC_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    Label101.Caption = ""
    Label103.Caption = ""
    Label203.Caption = ""
    Label303.Caption = ""
    Label301.Caption = ""
    Label401.Caption = ""
    Label402.Caption = ""
    Label403.Caption = ""
    List1.Clear
    LIUBIC.Clear
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        OPENFORMS = DoEvents
        Call MUEDIS(CI%)                ' recalcular existencias
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    ASEL$ = "MASTER"
    '
    Text1.SelLength = 0
    Call SELECHO(ASEL$)
    Text1.SelLength = 0
    CODD$ = VALACT1$(ASEL$)
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_Click()
    Text1.SetFocus
End Sub

Sub SHOWPEDIPEN()
     '
     Static UFEPEDET#
     If CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT")) <> UFEPEDET# Then
        UFEPEDET# = CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT"))
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        For REPRESU& = 1 To UREPRE&
          Call TRACE(20, REPRESU&, UREPRE&)
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
        Next REPRESU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Then
         PPXX% = InStr(XX$, "Nro.")
         If PPXX% > 0 Then
           XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
         End If
         If InStr(XX$, NUDOCU$) > 4 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
         End If
       End If
     Next UI&
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).Top = 0
       EDITFORM.Picture1(KKI%).Refresh
       EDITFORM.Picture1(KKI%).Height = 7170
       EDITFORM.Picture1(KKI%).Width = 11700
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     ALTUPAGINA = 7170
     TOPEPAGINA = 0
     PAGINACTIVA% = 0
     CAPAGINAS% = 0
     '
10   PPXX% = InStr(TPSP$, Chr$(255))
     If PPXX% > 0 Then
       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
       TPSP$ = Mid$(TPSP$, PPXX% + 1)
       Call TRACE(20, FIN& - Len(TPSP$), FIN&)
       '
       If Len(COPRIAT$) > 0 Then
         COPRI% = Asc(Left$(COPRIAT$, 1))
         ATRIP$ = Mid$(COPRIAT$, 2)
         Call SHOWCOMMAND(COPRI%, ATRIP$)
       End If
       GoTo 10
       '
     End If
     '
     Screen.MousePointer = 1
     PAGINACTIVA% = 0
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).Visible = True
     EDITFORM.Picture1(PAGINACTIVA%).Refresh
     EDITFORM.VScroll1.Min = 0
     EDITFORM.VScroll1.Value = 0
     EDITFORM.VScroll1.Max = 1
     If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     End If
     EDITFORM.COMMAND3.Enabled = False
     EDITFORM.Command1.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$, HOYY
     '
     Screen.MousePointer = 11
     If CTX% <> 1 Then
        IDENSEDE$ = CONTROL$("", "IDENSEDE")
        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        HOYY = HOY(HOS$)
        CTX% = 1
     End If
     '
     List1.Clear
     STOPROP = 0
     STOREMO = 0
     TOPEDID = 0
     TOSUMIN = 0
     TOFABRI = 0
     TORESER = 0
     STOMINI = STOMIN(CI%)
     LECONOM = LOTESTAN(CI%)
     GoSub PRESEDIS
     '
     Call CARDEPOS
     '
     Screen.MousePointer = 11
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     STOPROP = 0
     If CI% > 0 Then
       For U& = 1 To NUSU%
         CAPAR = STODEPOS(CI%, COSU%(U&)) '        XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
         If Abs(CAPAR) >= 0.05 Then
           XXX$ = ADEPO$(U&) + CSTRING$(MKS$(CAPAR), 4, 8, 1, 2)
           If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
           List1.AddItem XXX$
           If DISTO%(U&) >= 0 Then
             STOPROP = STOPROP + CAPAR
           End If
         End If
       Next U&
       GoSub PRESEDIS
     End If
     '
10   'List2.Clear
     HOII% = HOY(HO$)
     LI2TEX$ = Space$(96)
     Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 29)
     Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, 1, 2), 44, 9)
     Call REPLA(LI2TEX$, HO$, 74, 9)
     OBSER$ = ""
     If STOPROP < 0 Then
          OBSER$ = "!!! Negativo !!!"
        ElseIf STOPROP < STOMINI Then
          OBSER$ = "Sub-Mínimo"
     End If
     Call REPLA(LI2TEX$, OBSER$, 84, 24)
     'List2.AddItem LI2TEX$
     SALSTO = STOTOT
     '
12   TORECEP = 0
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         K& = 0
         Do While (.EOF = False)
            CARECEP = !CANTNOM - !CanToApro
            If CARECEP > 0 Then
              NCI% = !NumProv
              REMIPRO$ = !NumRem_Prov
              NROBOL& = !NumeBore
              FEX = CVINT(!FECHRECEP)
              LI2TEX$ = Space$(96)
              Call REPLA(LI2TEX$, "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CARECEP), 4, 9, 1, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
              Call REPLA(LI2TEX$, "RT." + REMIPRO$, 84, 24)
              List2.AddItem LI2TEX$
              TOSUMIN = TOSUMIN + CARECEP
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     '
     If TORECEP > 0.05 Then
        XXX$ = "RECEPC. " + CSTRING$(MKS$(TORECEP), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     GoSub PRESEDIS
     '
     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
     Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NROX& = !NroPed
           NCI% = !NROCLIE
           CAPEN = !CanPed - !CantEnt
           FEX = CVINT(!FeSolEnt)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
             LI2TEX$ = Space$(96)
             Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               OBSER$ = ""
               If FEX < HOYY Then
                 ATRAS% = DIENTRE%(FEX, HOYY)
                 OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
               End If
             Call REPLA(LI2TEX$, OBSER$, 84, 24)
             List2.AddItem LI2TEX$
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     GoTo 15
     '
     '
15   Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !Cant_Ocom * COEFCAN
             CAREC = !Cant_Rec * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = !Npro_Ocom
               NROX& = !Nume_Ocom
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   FEX = HOYY
                 End If
               LI2TEX$ = Space$(96)
               Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
               Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               Call REPLA(LI2TEX$, OBSER$, 84, 24)
               List2.AddItem LI2TEX$
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '
     '\\\OT-06-0149-WAR-25/04/06///
22      Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM AUTOENT "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Aten < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
     '
     Set AtenTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = 0 '!N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = 1 '!Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !Cant_Aten * COEFCAN
             CAREC = !Acuen_Aten * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = !Npro_Aten
               NROX& = !Nume_Aten
               FEX = CVINT(!Fentre_Sol)
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   FEX = HOYY
                 End If
               LI2TEX$ = Space$(96)
               Call REPLA(LI2TEX$, "AE." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
               Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 29)
               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               Call REPLA(LI2TEX$, OBSER$, 84, 24)
               List2.AddItem LI2TEX$
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '\\\OT-06-0149-WAR-FIN///
     '
25   Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CACOM = !CanProy
           CAREC = !CanApro
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             NroPed& = !NUMEORFA
             FEX = CVINT(!FechEnSol)
             REFE$ = !Referen
             LI2TEX$ = Space$(96)
             FPE% = CVINT(!FechProyEntre)
             Call REPLA(LI2TEX$, "OF." + TRIM$(CSTRING$(MKS$(NroPed&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, REFE$, 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
             'Call REPLA(LI2TEX$, IDENSEDE$, 84, 8)
             '
             If FPE% > 0 Then
               FEX = FPE%
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               FEX = CVINT(!FechRecal)
               Call REPLA(LI2TEX$, "Calc." + FECHATEX$(FEX), 84, 24)
             End If
             '
             List2.AddItem LI2TEX$
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '
35   Call ABRESERVA
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
     '
     Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           TORE# = !CantRes
           CACO# = !CantCons
           '
           If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
              TORE# = TORE# / 1000
              CACO# = CACO# / 1000
           End If
           '
           SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
           SARE = TORE# - CACO#
           If SARE * SIGNO% < 0 Then
             SARE = 0
           End If
           CAPEN = SARE
           If Abs(CAPEN) > 0 Then
             NORFA$ = TRIM$(!IdSubOr)
             FEX = CVINT(!FechRes)
             If Left$(NORFA$, 2) = "OR" Then
               REFE$ = "Reserva p/Reparación"
               GoTo 38
             End If
             '
             CICON% = 0
             Call ABREORFAB
             OrdeFabr.FindFirst "IdSubOr = '" & NORFA$ & "'"
             If OrdeFabr.NoMatch = False Then
               CICON% = OrdeFabr!Cod_Inte
             End If
             DESTIN$ = NORFA$
             If CICON% > 0 Then
               DESTIN$ = TRIM$(CODEXT$(CICON%)) + " - " + TRIM$(DESCRIT0$(CICON%))
             End If
             '
             REFE$ = "Reserva p/" + DESTIN$
             '
38            LI2TEX$ = Space$(96)
             Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
             Call REPLA(LI2TEX$, REFE$, 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
             'Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
             List2.AddItem LI2TEX$
             TORESER = TORESER + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     GoSub PRESEDIS
     '
     STOTER = 0
     Call ABRESTOCKS
     SQLX$ = "SELECT * FROM MACONSIG "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + " ORDER BY Femi_Ocom "
     Set LIMACONTMP = Stocks.OpenRecordset(SQLX$, 4)
     '
     With LIMACONTMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TORE# = !Cant_Entreg
           CACO# = !Cant_Rendid
           CAPEN = TORE# - CACO#
           CAPENENT# = !Cant_Asig - !Cant_Entreg
           NUOC& = XVALO(Mid$(!Nro_Ocom, 4, 6))
           '
           If Abs(CAPEN) > 0 Then
              NUREX$ = "RC-" + TRIM$(!NRemi_Consig)
              NORFA$ = !Nro_Ocom
              NPRO% = !Npro_Ocom
              REFE$ = RASOCLI$((-1) * NPRO%)
              FEX = CVINT(!Fecha_Entreg)
              SDOR$ = ""
              LI2TEX$ = Space$(96)
              '
              Call REPLA(LI2TEX$, Left$(NUREX$, 12), 1, 12)
              Call REPLA(LI2TEX$, REFE$, 15, 30)
              If CAPEN > 0 Then
                   Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
                 Else
                   Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPEN)), 4, 9, 1, 2), 53, 9)
              End If
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
              Call REPLA(LI2TEX$, NORFA$, 84, 24)
              List2.AddItem LI2TEX$
              STOREMO = STOREMO + CAPEN
              STOTER = STOTER + CAPEN
              GoTo 47
           End If
           '
47         If CAPENENT# >= 0.05 Then
             If (STATOCOM%(NUOC&) < 8 And STATOCOM%(NUOC&) >= 0) Then
                NORFA$ = !Nro_Ocom
                NPRO% = !Npro_Ocom
                REFE$ = RASOCLI$((-1) * NPRO%)
                FEX = CVINT(!Femi_Ocom)
                SDOR$ = ""
                LI2TEX$ = Space$(96)
                '
                Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
                Call REPLA(LI2TEX$, REFE$, 15, 30)
                If CAPENENT# > 0 Then
                     Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPENENT#), 4, 9, 1, 2), 53, 9)
                   Else
                     Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPENENT#)), 4, 9, 1, 2), 44, 9)
                End If
                Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
                List2.AddItem LI2TEX$
                TORESER = TORESER + CAPENENT
             End If
           End If
           '
         .MoveNext
         Loop
       End If
     End With
     '
     If Abs(STOTER) >= 0.005 Then
        XXX$ = "TERCEROS" + CSTRING$(MKS$(STOTER), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     '
     GoSub PRESEDIS
     '
     ' ORDENAMIENTO POR FECHA
     For U& = 1 To List2.ListCount
45     TTXX$ = List2.List(U& - 1)
       FE1% = FECHANUM(Mid$(TTXX$, 74, 8))
       If Mid$(UCase$(TTXX$), 15, 12) = "STOCK PROPIO" Then FE1% = 0
       For U2& = U& + 1 To List2.ListCount
         TTYY$ = List2.List(U2& - 1)
         FE2% = FECHANUM(Mid$(TTYY$, 74, 8))
         If FE2% < FE1% Then
           List2.RemoveItem U2& - 1
           List2.AddItem TTYY$, U& - 1
           GoTo 45
         End If
       Next U2&
     Next U&
     List2.Refresh
     '
     ' STOCK PROYECTADO
     STOPRO = 0
     For U& = 1 To List2.ListCount
       TTXX$ = List2.List(U& - 1)
       STOPRO = STOPRO + XVALO(Mid$(TTXX$, 44, 9)) - XVALO(Mid$(TTXX$, 53, 9))
       STOPROX$ = FORMATNUM$(STOPRO, "F10.1")
       Call REPLA(TTXX$, STOPROX$, 62, 10)
       List2.RemoveItem U& - 1
       List2.AddItem TTXX$, U& - 1
     Next U&
     List2.Refresh
     '
20   LIUBIC.Clear
     RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
     For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           Y$ = Mid$(X$, 3, 24)
           LIUBIC.AddItem Y$
        End If
     Next U&
     '
     Screen.MousePointer = 1
     Call GRAPHDIS
     '
     Exit Sub
     '
PRESEDIS:
     '
     STOMI.Caption = TRIM$(CSTRING$(MKS$(STOMINI), 4, 12, 1, 2))
     LECONOMI.Caption = TRIM$(CSTRING$(MKS$(LECONOM), 4, 12, 1, 2))
     '
     Label101.Caption = TRIM$(CSTRING$(MKS$(STOPROP), 4, 12, 1, 2))
     Label102.Caption = TRIM$(CSTRING$(MKS$(STOREMO), 4, 12, 1, 2))
     STOTOT = STOPROP + STOREMO
     Label103.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, 1, 2))
     '
     Label201.Caption = TRIM$(CSTRING$(MKS$(TOPEDID), 4, 12, 1, 2))
     Label202.Caption = TRIM$(CSTRING$(MKS$(TORESER), 4, 12, 1, 2))
     TOCOMPR = TOPEDID + TORESER
     Label203.Caption = TRIM$(CSTRING$(MKS$(TOCOMPR), 4, 12, 1, 2))
     '
     Label301.Caption = TRIM$(CSTRING$(MKS$(TOSUMIN), 4, 12, 1, 2))
     Label302.Caption = TRIM$(CSTRING$(MKS$(TOFABRI), 4, 12, 1, 2))
     TOINPRE = TOSUMIN + TOFABRI
     Label303.Caption = TRIM$(CSTRING$(MKS$(TOINPRE), 4, 12, 1, 2))
     '
     STONET = STOTOT - TOCOMPR
     STODIS = STONET - STOMINI + TOINPRE
     STOFAL = 0
     If STODIS < 0 Then
        STOFAL = Abs(STODIS)
        STODIS = 0
     End If
     '
     Label401.Caption = TRIM$(CSTRING$(MKS$(STONET), 4, 12, 1, 2))
     Label402.Caption = TRIM$(CSTRING$(MKS$(STODIS), 4, 12, 1, 2))
     Label403.Caption = TRIM$(CSTRING$(MKS$(STOFAL), 4, 12, 1, 2))
     '
     DoEvents
     '
Return
'
End Sub
'
'
Sub GRAPHDIS()
   '
   If Picture1.Visible = False Then
     Exit Sub
   End If
   '
   ALTUMA = 0.78 * Picture1.Height
   STOK = XVALO(Label103): If STOK < 0 Then STOK = 0
   COMP = XVALO(Label301): If COMP < 0 Then COMP = 0
   FABR = XVALO(Label302): If FABR < 0 Then FABR = 0
   '
   MINI = XVALO(STOMI): If MINI < 0 Then MINI = 0
   RESE = XVALO(Label203): If RESE < 0 Then RESE = 0
   '
   Screen.MousePointer = 11
   CIIX% = CODINT%(Text1)
   TENT = 0
   RFF$ = RECFILT$("COBERFA", 2, MKI$(CIIX%))
   For KU& = 1 To Len(RFF$) Step 4
     KL& = CVS(Mid$(RFF$, KU&, 4))
     XZ$ = REGLEIDO$("COBERFA", KL&)
     If CVI(Mid$(XZ$, 51, 2)) = 22 Then
       TENT = TENT + CVS(Mid$(XZ$, 57, 4))
     End If
   Next KU&
   If TENT < 0 Then TENT = 0
   '
   MAXI = STOK + COMP + FABR
   If MINI + RESE + TENT > MAXI Then
     MAXI = MINI + RESE + TENT
   End If
   Screen.MousePointer = 1
   '
   If MAXI < 1 Then
     Shape1.Visible = False
     Shape2.Visible = False
     Shape3.Visible = False
     Shape4.Visible = False
     Shape5.Visible = False
     Shape6.Visible = False
     Exit Sub
   End If
   '
   COEF = ALTUMA / MAXI
   
   Shape1.Height = STOK * COEF
   Shape1.Top = 0.94 * Picture1.Height - Shape1.Height
   Shape2.Height = COMP * COEF
   Shape2.Top = Shape1.Top - Shape2.Height
   Shape3.Height = FABR * COEF
   Shape3.Top = Shape2.Top - Shape3.Height
   '
   Shape4.Height = MINI * COEF
   Shape4.Top = 0.94 * Picture1.Height - Shape4.Height
   Shape5.Height = RESE * COEF
   Shape5.Top = Shape4.Top - Shape5.Height
   Shape6.Height = TENT * COEF
   Shape6.Top = Shape5.Top - Shape6.Height
   '
   If STOK > 0 Then Shape1.Visible = True
   If COMP > 0 Then Shape2.Visible = True
   If FABR > 0 Then Shape3.Visible = True
   If MINI > 0 Then Shape4.Visible = True
   If RESE > 0 Then Shape5.Visible = True
   If TENT > 0 Then Shape6.Visible = True

End Sub


'
Sub RECIBETRA()
    '
    Screen.MousePointer = 11
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then Exit Sub
    '
10  LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       Exit Sub
    End If
    On Error GoTo 20
    FileCopy ORIGARCH$ + "\ATRASUC.DAT", LU$ + "ATRASUC.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Unidad de Transferencia '" + ORIGARCH$ + "'\Leída Exitosamente.\  \Pulse 'Aceptar Mensaje' para\Iniciar Importación de Datos.")
    GoTo 30
    '
20  Resume 21
21  On Error GoTo 0
    Call MENSERR(24, "Imposible Leer Unidad de Transferencia.")
    Exit Sub
    '
30  Screen.MousePointer = 11
    Dim FINUAR%(128), ARTRAS$(128), UREAR&(128)
    Dim BUFGRA As String * 128
    CANAR% = 0
    LU$ = LUDAT$
    FIN& = ULTREG&("ATRASUC")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ATRASUC", U&)
      ANOM$ = Left$(X$, 8)
      TTXX$ = Mid$(X$, 9, 128)
      '
      For KKI% = 1 To CANAR%
        If ARTRAS$(KKI%) = ANOM$ Then GoTo 35
      Next KKI%
      CANAR% = CANAR% + 1
      KKI% = CANAR%
      FINUAR%(KKI%) = FILEOPEN%(LU$ + TRIM$(ANOM$) + ".DAT", 0, 128)
      ARTRAS$(KKI%) = ANOM$
      UREAR&(KKI%) = 0
      '
35    UREAR&(KKI%) = UREAR&(KKI%) + 1
      BUFGRA$ = TTXX$
      Put #FINUAR%(KKI%), UREAR&(KKI%), BUFGRA$
    Next U&
    Screen.MousePointer = 1
    '
    For KKI% = 1 To CANAR%
      FIN& = LOF(FINUAR%(KKI%)) / 128
      If LOF(FINUAR%(KKI%)) > 128 * FIN& Then FIN& = FIN& + 1
      For U& = UREAR&(KKI%) + 1 To FIN&
        BUFGRA$ = String$(128, 0)
        Put #FINUAR%(KKI%), UREAR&(KKI%), BUFGRA$
      Next U&
      Close #FINUAR%(KKI%)
      FINUAR%(KKI%) = 0
    Next KKI%
    '
    For KKI% = 1 To CANAR%
      X1$ = REGLEIDO$(ARTRAS$(KKI%), 1)
      Call GRAREG(ARTRAS$(KKI%), X1$, 1)
      Call CIERRARCH(ARTRAS$(KKI%))
    Next KKI%
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    URET& = ULTREG&("RTRASLA")
    For U& = 1 To ULTREG("RTRASLA")
      XX$ = REGLEIDO$("RTRASLA", U&)
      IDENDO$ = Mid$(XX$, 19, 4)
      RFF$ = RECFILT$("ATRASLA", 4, IDENDO$)
      For UJ& = 1 To Len(RFF$) Step 4
        RELE& = CVS(Mid$(RFF$, UJ&, 4))
        If RELE& > 0 Then
          If RELE& <= ULTREG&("ATRASLA") Then
            YY$ = REGLEIDO$("ATRASLA", RELE&)
            If Left$(YY$, 112) = Left$(XX$, 112) Then GoTo 45 ' YA RECIBIDO
          End If
        End If
      Next UJ&
      '
      YY$ = Left$(XX$, 112) + String$(256, 0)
      Call REGAPP("ATRASLA", YY$)
45  Next U&
    '
    Call CIERRARCH("ATRASLA")
    Call CIERRARCH("RTRASLA")
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    Call COMUNI("Importación Completa.")
    '
End Sub

'
Sub GENARTRA()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ENTRECOM", "TRAENTRE")
    Call BLOQARCH("ATRACEN")
    Call BLANARCH("ATRACEN")
    '
    JJ& = 0
    For U& = 1 To ULTREG("ENTRECOM")
      XX$ = REGLEIDO$("ENTRECOM", U&)
      STATU% = CVI(Mid$(X$, 191, 2))
      If STATU% <> 9 Then        ' NO ESTA ANULADO
        VENVI% = CVI(Mid$(XX$, 193, 2))     ' veces enviado
        If VENVI% < 5 Then
          Call REPLA(XX$, MKI$(VENVI% + 1), 193, 2)
          Call GRAREG("ENTRECOM", XX$, U&)
          JJ& = JJ& + 1
          Call GRAREG("TRAENTRE", XX$, JJ&)
        End If
      End If
    Next U&
    Call SETULTREG("TRAENTRE", JJ&)
    Call CIERRARCH("ENTRECOM")
    Call CIERRARCH("TRAENTRE")
    '
    Dim BUFLE As String * 128
    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + "TRAENTRE.DAT", 0, 128)
    For KU& = 1 To 1 + LOF(N1%) / 128
      Get #N1%, KU&, BUFLE$
      TTXX$ = Space$(8) + String$(128, 0)
      Call REPLA(TTXX$, "TRAENTRE", 1, 8)
      Call REPLA(TTXX$, BUFLE$, 9, 128)
      Call REGAPP("ATRACEN", TTXX$)
    Next KU&
    Close #N1%
    '
    Call CIERRARCH("ATRACEN")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    FILEORIG.Caption = "Destino del Archivo"
40  FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then GoTo 99
    '
    LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       Exit Sub
    End If
    On Error GoTo 50
    FileCopy LU$ + "ATRACEN.DAT", ORIGARCH$ + "\ATRACEN.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Archivo de Transferencia Electrónica\para Central Avellaneda\Generado Exitosamente en '" + ORIGARCH$ + "'.")
    GoTo 99
    '
50  Resume 51
51  On Error GoTo 0
    Call MENSERR(24, "Error de Copiado a Unidad.\   \ Intente de Nuevo.")
    GoTo 99
    '
99  Screen.MousePointer = 1
End Sub



VERSION 5.00
Begin VB.Form SITSTOK04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   ClipControls    =   0   'False
   Icon            =   "SITSTOK04.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture6 
      Height          =   8100
      Left            =   10950
      ScaleHeight     =   8040
      ScaleWidth      =   915
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image3 
         Height          =   8150
         Left            =   0
         Picture         =   "SITSTOK04.frx":014A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6615
      Top             =   105
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame11"
      Height          =   3450
      Left            =   8715
      TabIndex        =   71
      Top             =   4410
      Width           =   15
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   3450
      Left            =   7740
      TabIndex        =   70
      Top             =   4410
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   3450
      Left            =   6615
      TabIndex        =   69
      Top             =   4410
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   3450
      Left            =   5670
      TabIndex        =   68
      Top             =   4410
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   3450
      Left            =   4725
      TabIndex        =   67
      Top             =   4410
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   3450
      Left            =   1575
      TabIndex        =   66
      Top             =   4410
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00004040&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1140
      TabIndex        =   55
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   64
         Top             =   6300
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "SITSTOK04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Imprime Situación de Stocks"
         Top             =   6615
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
         Picture         =   "SITSTOK04.frx":2B4C
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Realizar Movimientos de Stock"
         Top             =   5100
         Width           =   650
      End
      Begin VB.CommandButton COMMAND3 
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
         Height          =   690
         Left            =   105
         Picture         =   "SITSTOK04.frx":2C96
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2625
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Price"
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
         Picture         =   "SITSTOK04.frx":2FA0
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Consulta de Precios de Venta"
         Top             =   3315
         Width           =   650
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
         Picture         =   "SITSTOK04.frx":32AA
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Acceso a Consultas y Listados"
         Top             =   4410
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
         Picture         =   "SITSTOK04.frx":33F4
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   180
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
         Picture         =   "SITSTOK04.frx":353E
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   870
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
         Picture         =   "SITSTOK04.frx":3848
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1560
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "SITSTOK04.frx":3C8A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7450
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   2680
      Left            =   8580
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   53
      Top             =   1575
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   2325
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   2310
         Y1              =   270
         Y2              =   270
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2680
      Left            =   8580
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   50
      Top             =   1575
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2310
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cobertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   2325
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   525
         Width           =   950
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1155
         Width           =   950
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1785
         Width           =   950
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   750
         Left            =   105
         Top             =   525
         Width           =   950
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   105
         Top             =   1260
         Width           =   950
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   105
         Top             =   1785
         Width           =   950
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   105
      MultiSelect     =   1  'Simple
      TabIndex        =   22
      Top             =   4750
      Width           =   10800
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
      Height          =   600
      Left            =   5670
      Picture         =   "SITSTOK04.frx":5BAC
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Visualizar Ordenes de Venta Pendientes de Entrega"
      Top             =   210
      Visible         =   0   'False
      Width           =   560
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   435
      Left            =   4200
      TabIndex        =   42
      Top             =   210
      Width           =   1275
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Maestro"
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
         TabIndex        =   44
         Top             =   210
         Width           =   1240
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0F0F0&
         Caption         =   "L.Precios"
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
         TabIndex        =   43
         Top             =   0
         Value           =   -1  'True
         Width           =   1275
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
      Height          =   1120
      Left            =   105
      TabIndex        =   33
      Top             =   3150
      Width           =   8310
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
         Height          =   345
         Left            =   105
         Picture         =   "SITSTOK04.frx":5EB6
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Muestra / Oculta Situación de Cobertura"
         Top             =   630
         Width           =   560
      End
      Begin VB.Image Image1 
         Height          =   250
         Left            =   105
         Picture         =   "SITSTOK04.frx":61C0
         Stretch         =   -1  'True
         ToolTipText     =   "Histórico de Existencias"
         Top             =   315
         Width           =   560
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Reserv:"
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
         Left            =   2205
         TabIndex        =   49
         Top             =   375
         Width           =   1380
      End
      Begin VB.Label Label404 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   3105
         TabIndex        =   48
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label401 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         TabIndex        =   39
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St.Neto:"
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
         Left            =   210
         TabIndex        =   38
         Top             =   375
         Width           =   1380
      End
      Begin VB.Label Label402 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   5010
         TabIndex        =   37
         Top             =   675
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
         Left            =   4200
         TabIndex        =   36
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
         Left            =   6090
         TabIndex        =   35
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label403 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   6930
         TabIndex        =   34
         Top             =   675
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
      Height          =   1590
      Left            =   5670
      TabIndex        =   15
      Top             =   1470
      Width           =   2745
      Begin VB.Label Label303 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   32
         Top             =   1155
         Width           =   1170
      End
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
         Left            =   420
         TabIndex        =   31
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label301 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   19
         Top             =   315
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
         Left            =   420
         TabIndex        =   18
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label302 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   17
         Top             =   735
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
         Left            =   210
         TabIndex        =   16
         Top             =   840
         Width           =   1065
      End
   End
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
      Height          =   1590
      Left            =   2835
      TabIndex        =   10
      Top             =   1470
      Width           =   2745
      Begin VB.Label Label203 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   30
         Top             =   1155
         Width           =   1170
      End
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
         Left            =   420
         TabIndex        =   29
         Top             =   1260
         Width           =   855
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
         Left            =   210
         TabIndex        =   14
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label202 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   13
         Top             =   735
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
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label201 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1365
         TabIndex        =   11
         Top             =   315
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
      Height          =   1590
      Left            =   105
      TabIndex        =   8
      Top             =   1470
      Width           =   2640
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1260
         TabIndex        =   28
         Top             =   1155
         Width           =   1170
      End
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
         Left            =   315
         TabIndex        =   27
         Top             =   1260
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
         Left            =   315
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1260
         TabIndex        =   25
         Top             =   735
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
         Left            =   315
         TabIndex        =   24
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   1260
         TabIndex        =   9
         Top             =   315
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
      Height          =   2085
      Left            =   8580
      TabIndex        =   5
      Top             =   480
      Width           =   2325
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
      Left            =   1050
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
      Left            =   7770
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
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5370
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
      Left            =   3465
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
      Height          =   1185
      Left            =   8580
      TabIndex        =   2
      Top             =   3070
      Width           =   2325
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10920
      X2              =   10920
      Y1              =   0
      Y2              =   8400
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   100
      X2              =   12105
      Y1              =   7870
      Y2              =   7870
   End
   Begin VB.Line Line4 
      X1              =   8580
      X2              =   10895
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line Line3 
      X1              =   8580
      X2              =   10895
      Y1              =   4260
      Y2              =   4260
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
      Left            =   5985
      TabIndex        =   46
      Top             =   1080
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
      Left            =   7350
      TabIndex        =   45
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
      Left            =   7350
      TabIndex        =   41
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
      Left            =   5985
      TabIndex        =   40
      Top             =   765
      Width           =   1275
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comprob.         Referencia                                        Entra         Sale      St.Proy.    Fecha      Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   23
      Top             =   4470
      Width           =   10800
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
      Height          =   300
      Left            =   8580
      TabIndex        =   21
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   2805
      Width           =   2325
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
      Height          =   300
      Left            =   8580
      TabIndex        =   20
      Top             =   210
      Width           =   2325
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   315
      Width           =   960
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
      Left            =   6825
      TabIndex        =   6
      Top             =   315
      Width           =   855
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu FichKar 
         Caption         =   "Ficha Tipo Kardex"
      End
      Begin VB.Menu SitCob 
         Caption         =   "Mostrar / Ocultar Situación de Cobertura"
      End
      Begin VB.Menu CoPreVe 
         Caption         =   "Consulta de Precios de Venta"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConsyList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu Gen 
         Caption         =   "Generales"
         Begin VB.Menu ExistyUbi 
            Caption         =   "Existencias y Ubicaciones"
         End
         Begin VB.Menu FichaKardex 
            Caption         =   "Ficha Kardex -Historico de Movimientos-"
         End
      End
      Begin VB.Menu DeCtrl 
         Caption         =   "De Control"
         Begin VB.Menu ExFiAct 
            Caption         =   "Existencias Físicas Actuales"
            Begin VB.Menu ExFiPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ExFiImp 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu ExiNeg 
            Caption         =   "Existencias Negativas"
            Begin VB.Menu ExNegPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ExNegImpr 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu ItStSubMi 
            Caption         =   "Items con Stock Sub-Mínimo"
            Begin VB.Menu ItStPant 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu ItStImp 
               Caption         =   "Por Impresión"
            End
         End
         Begin VB.Menu DiGenSt 
            Caption         =   "Disponibilidad General de Stock"
            Begin VB.Menu DiStPan 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu DiStImp 
               Caption         =   "Por Impresión"
            End
         End
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu GesReser 
         Caption         =   "Gestión de Reservas"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu ParFabr 
         Caption         =   "Ordenes de Fabricación Pendientes por Artículo"
      End
   End
   Begin VB.Menu AcceDire 
      Caption         =   "Complementos"
      Begin VB.Menu MovStock 
         Caption         =   "Material en Depósitos"
      End
      Begin VB.Menu php25 
         Caption         =   "-"
      End
      Begin VB.Menu movimientodestocks 
         Caption         =   "Movimientos de Stocks"
         Begin VB.Menu recepciondeproveedores 
            Caption         =   "Recepción de Proveedores"
         End
         Begin VB.Menu s5 
            Caption         =   "-"
         End
         Begin VB.Menu consumosproduccion 
            Caption         =   " Informar Consumos por Propia Producción"
         End
         Begin VB.Menu S 
            Caption         =   "-"
         End
         Begin VB.Menu remitirmercaderias 
            Caption         =   "Remitir Mercaderías"
         End
         Begin VB.Menu s6 
            Caption         =   "-"
         End
         Begin VB.Menu remitoconsignacion 
            Caption         =   "Remito en Consignacion - Proveedores de Servicios"
         End
         Begin VB.Menu s7 
            Caption         =   "-"
         End
         Begin VB.Menu ajusteexistencia 
            Caption         =   "Ajuste de Existencias"
         End
         Begin VB.Menu s8 
            Caption         =   "-"
         End
         Begin VB.Menu transferencias 
            Caption         =   "Transferencias de Stock"
         End
         Begin VB.Menu s9 
            Caption         =   "-"
         End
         Begin VB.Menu valeconsumo 
            Caption         =   "Vale de Consumo de Materiales"
         End
         Begin VB.Menu s10 
            Caption         =   "-"
         End
         Begin VB.Menu descartesstock 
            Caption         =   "Descartes de Stock"
         End
      End
      Begin VB.Menu estadisticasdestock2 
         Caption         =   "Estadisticas de Stocks"
         Begin VB.Menu estadisticasdestock 
            Caption         =   "Estadisticas de Stock"
         End
      End
      Begin VB.Menu transferenciastock 
         Caption         =   "Informar Transferencias de Stock"
      End
      Begin VB.Menu ss5 
         Caption         =   "-"
      End
      Begin VB.Menu controldepositos 
         Caption         =   "Control de Depósitos"
      End
      Begin VB.Menu ss6 
         Caption         =   "-"
      End
      Begin VB.Menu OtrAccDirec 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu confi 
      Caption         =   "Configuración"
      Begin VB.Menu configdispo 
         Caption         =   "Configuración - Disponibilidad de Stocks"
      End
      Begin VB.Menu mnuConfiguraciongeneral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu util 
      Caption         =   "Utilidades"
      Begin VB.Menu Stockminimorecomendado 
         Caption         =   "Stock Minimo Recomendado"
      End
   End
   Begin VB.Menu Ayud 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "SITSTOK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%

Private Sub Abrir_Click()
Call Command2_Click
End Sub

Private Sub ajusteexistencia_Click()
   Call MOVISTK04.Command25_Click
End Sub

Private Sub Ayud_Click()
Call Command5_Click
End Sub

Private Sub Command1_Click()
    IDTX$ = IDENTER$
    AVALSTOK$ = "LISTAPRE"
    If Option2.Value = True Then
       AVALSTOK$ = "MAESTRO"
    End If
    Call GRACONTROL(IDTX$, "AVALSTOK", AVALSTOK$)
    Call CIERRARCH("*.*")
    If Left$(App.EXEName, 6) = "ESTOCK" Then
        Hide
      Else
        Call FINAL("")
    End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call SHOWPEDIPEN
   Command10.Enabled = True
End Sub



Private Sub Command2_Click()
    '
    ASEL$ = "MASTER"
    If Option1.Value = True Then
       ASEL$ = "LIPRE001"
    End If
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
    If (USNBR% Mod 100) = 1 Then
         OPDISTK04.Show 1
       Else
         Call MENSERR(24, "La Opción Elegida Requiere Privilegios de Supervisor")
    End If
    Command27.Enabled = True
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    FIKARDEX.Text1 = Text1
    FIKARDEX.Show 1
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   If Picture1.Visible = True Then
       Picture1.Visible = False
     Else
       Picture1.Visible = True
   End If
   Call GRAPHDIS
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   MOVISTK04.Show 1
   If CODINT%(Text1.TEXT) > 0 Then
      COD$ = TRIM$(Text1.TEXT)
      CI% = CODINT%(COD$)
      Text2.TEXT = UNIMED$(CI%)
      Text3.TEXT = DESCRIT$(CI%)
      Call MUEDIS(CI%)                ' recalcular existencias
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Static CTXX%, MODOCORE%, FECACOS$
   '
   If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
      Exit Sub
   End If
   '
   CODD$ = TRIM$(Text1.TEXT)
   CIXX% = CODINT%(CODD$)
   If CIXX% <= 0 Then
      Exit Sub
   End If
   '
   If MODOCORE% < 1 Then
      MODOCORE% = 1
      MCDR$ = TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE")))
      If MCDR$ = "PREBASE" Then
           MODOCORE% = 2
         ElseIf MCDR$ = "COSCALCU" Then
           MODOCORE% = 3
           If FECACOS$ = "" Then
             FECACOS$ = TRIM$(CONTROL("ANACOSTO", "FECALCOS"))
           End If
      End If
   End If
   '
   MUEPRECIO.List1.Clear
   ZZ$ = "   C.Reposición" + Space$(30): COEFI = 0
   If MODOCORE% = 2 Then
        ZZ$ = "Precio Base" + Space$(30)
        COEFI = 15
     ElseIf MODOCORE% = 3 Then
        ZZ$ = "Cos.Calcul.(" + FECACOS$ + ")" + Space$(30)
   End If
   '
   COPRES = COSUNI(CIXX%) * (1 + COEFI / 100)
   If MODOCORE% = 3 Then COPRES = COSCALCU(CIXX%, MODOCO%)
   Call REPLA(ZZ$, CSTRING$(MKS$(COPRES), 4, 11, 4, 2), 25, 11)
   MUEPRECIO.List1.AddItem ZZ$
   MUEPRECIO.List1.AddItem Space$(80)    'String$(80, 45)
   '
   For U& = 1 To ULTREG&("LISTAS")
      ZZ$ = Mid$(REGLEIDO$("LISTAS", U&), 2, 32)
      If TRIM$(ZZ$) <> "" Then
         ZZ$ = AJUSTD$(TRIM$(Str$(U&)), 2) + "-" + ZZ$
         CLIS$ = TRIM$(Str$(U&))
         While Len(CLIS$) < 3: CLIS$ = "0" + CLIS$: Wend
         APREC$ = "LIPRE" + CLIS$
         PRLST# = PRELISTA#(APREC$, CIXX%)
         If PRLST# >= 0.00005 Then
           Call REPLA(ZZ$, CSTRING$(MKD$(PRLST#), 4, 11, 4, 2), 25, 11)
           MUEPRECIO.List1.AddItem ZZ$
         End If
      End If
   Next U&
   '
   MUEPRECIO.Show 1
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   FORIPRE$ = TRIM$(CONTROL$("", "FORDISTO"))
   If FORIPRE$ = "" Then
        PRINTFORM
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Situacion de Stocks"
        '
        CODPARAM$(1) = "COD-MAT"
        CODPARAM$(2) = "DES-MAT"
        CODPARAM$(3) = "CANTIDAD"
        CODPARAM$(4) = "CANT-MIN"
        CODPARAM$(5) = "CANT-RES"
        CODPARAM$(6) = "CANT-ING"
        CODPARAM$(7) = "CANT-DIS"
        CODPARAM$(8) = "CANT-FAL"
        CAPARDOC% = 8
        '
        DOCPARAM$(1) = TRIM$(Text1.TEXT)
        DOCPARAM$(2) = TRIM$(Text3.TEXT)
        DOCPARAM$(3) = TRIM$(Label103.Caption)
        DOCPARAM$(4) = TRIM$(STOMI.Caption)
        DOCPARAM$(5) = TRIM$(Label203.Caption)
        DOCPARAM$(6) = TRIM$(Label303.Caption)
        DOCPARAM$(7) = TRIM$(Label402.Caption)
        DOCPARAM$(8) = TRIM$(Label403.Caption)
        '
        CODTABU1$(1) = "REF-MAT1"
        CACOLTAB1% = 1
        '
        CAITAB1% = 0
        For U& = 1 To List2.ListCount
           Z$ = List2.List(U& - 1)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = "." + Left$(Z$, 82)
        Next U&
        '
        Call PRINTDOC("FORDISTO")
        Screen.MousePointer = 1
        '
      End If
      '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    QUERSTK04.Show 1
    Command9.Enabled = True
End Sub

Private Sub configdispo_Click()
   Call Command27_Click

End Sub

Private Sub consumosproduccion_Click()
    Call MOVISTK04.Command18_Click
End Sub

Private Sub controldepositos_Click()
   TRANSTK04.Command27_Click
End Sub

Private Sub CoPreVe_Click()
Call Command7_Click
End Sub

Private Sub descartesstock_Click()
   Call MOVISTK04.Command17_Click
End Sub

Private Sub DiStImp_Click()
DISPSTK04.Option12 = True
Call QUERSTK04.Command5_Click
End Sub

Private Sub DiStPan_Click()
  DISPSTK04.Option11 = True
  Call QUERSTK04.Command5_Click
End Sub

Private Sub estadisticasdestock_Click()
  Call QUERSTK04.Command27_Click
End Sub

Private Sub ExFiImp_Click()
QUERSTK04.Option4 = True
Call QUERSTK04.Command19_Click
End Sub

Private Sub ExFiPan_Click()
QUERSTK04.Option5 = True
Call QUERSTK04.Command19_Click
End Sub

Private Sub ExistyUbi_Click()
Call QUERSTK04.Command2_Click
End Sub

Private Sub ExNegImpr_Click()
QUERSTK04.Option4 = True
Call QUERSTK04.Command30_Click
End Sub

Private Sub ExNegPan_Click()
QUERSTK04.Option5 = True
Call QUERSTK04.Command30_Click
End Sub

Private Sub FichaKardex_Click()
Call QUERSTK04.Command9_Click
End Sub

Private Sub FichKar_Click()
Call Command3_Click
End Sub

Private Sub Form_Activate()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
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
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 7650
      Set MARCOBARRA.Container = Me.Picture6
      
      
      MARCOBARRA.ZOrder 0
       
       
  End If

   Call GRATITFOR(Caption)
   WindowState = 0
   '
   OPENFORMS = DoEvents
   Call FRESHALL
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '

    If Screen.Height > 9000 Then
       Top = (Screen.Height - Height) / 2
       Left = (Screen.Width - Width) / 2
    End If
    '
    VERANTER$ = "CONSTOCK"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("PEDIDO02.EXE") > 0 Then
      Command10.Visible = True
    End If
    '
    If InStr(EPAC$, "DOSIVA") > 0 Then
      Command10.ToolTipText = "Anotador de Pedidos - Despachar Materiales Vendidos"
      HOII% = HOY(HOS$)
      If HOII% <= FECHANUM("12/05/03") Then
        Command11.Visible = True   ' recupera packing list
      End If
      Label15 = " Comprob.           Referencia                                                Entra          Sale         Fecha      Observ."
    End If
    '
    LU$ = LUDAT$
    On Error Resume Next
    Show
    On Error GoTo 0
    OPENFORMS = DoEvents
    '
    If EXISTE%(LUDAT$ + "LIPRE001.DAT") <> 1 Then
       Option2.Value = True
    End If
    '
    IDTX$ = IDENTER$
    If CONTROL$(IDTX$, "AVALSTOK") = "MAESTRO" Then
       Option2.Value = True
    End If
    '
    If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
       Option2.Value = True
    End If
    '
    Call ABREPEDCLI
    SQLX$ = " SELECT PEDENCA.NroPed FROM PEDENCA "
    SQLX$ = SQLX$ + " WHERE Status < 9 "
    Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PediModifTemp
       On Error Resume Next
       .MoveLast
       If Not (Err) Then
          On Error GoTo 0
          If .RecordCount > 0 Then
             Command10.Visible = True
          End If
       End If
    End With
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Label19_Click()

End Sub


Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub GesRese_Click()
Call QUERSTK04.Command14_Click
End Sub

Private Sub GesReser_Click()
Call QUERSTK04.Command7_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   CIII% = CODINT%(Text1)
   Picture2.Cls
   If CIII% > 0 Then
     Screen.MousePointer = 11
     FACT% = HOY(HOS$)
     FINI% = FACT% - 32 * 13 + 1
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIII%))
     STK = 0: STMAX = 0
     For KKL& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, KKL&, 4))
       XX$ = REGLEIDO$("MOVISTO", REMO&)
       FEMO% = CVI(Left$(XX$, 2))
       STK = STK + CVD(Mid$(XX$, 88, 8)) - CVD(Mid$(XX$, 96, 8))
       If FEMO% >= FINI% Then
         If STK > STMAX Then STMAX = STK
         If STMAXAN > STMAX Then STMAX = STMAXAN
       End If
       If STMAX < 0.1 Then STMAXAN = STK
     Next KKL&
     If STMAX < 0.1 Then STMAX = STMAXAN
     Picture2.ForeColor = RGB(0, 0, 0)
     Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height)-(0.97 * Picture2.Width, 0.94 * Picture2.Height)
     Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height)-(0.05 * Picture2.Width, 300)
     If STMAX > 0.1 Then
       COEFV = 0.8 * (Picture2.Height - 300) / STMAX
       COEFH = 0.88 * Picture2.Width / (FACT% - FINI%)
       XINIA = 0.05 * Picture1.Width
       YINIA = 0.94 * Picture1.Height - COEFV * STMAXAN
       Picture2.ForeColor = RGB(255, 0, 0)
       STK = 0
       For KKL& = 1 To Len(RFF$) Step 4
         REMO& = CVS(Mid$(RFF$, KKL&, 4))
         XX$ = REGLEIDO$("MOVISTO", REMO&)
         FEMO% = CVI(Left$(XX$, 2))
         STK = STK + CVD(Mid$(XX$, 88, 8)) - CVD(Mid$(XX$, 96, 8))
         STX = STK: If STX < 0 Then STX = 0
         YINI = 0.94 * Picture2.Height - COEFV * STX
         XINI = 0.05 * Picture2.Width + COEFH * (FEMO% - FINI%)
         If XINI >= 0.05 * Picture1.Width Then
           Picture2.Line (XINIA, 0.94 * Picture2.Height)-(XINI, YINIA), RGB(255, 80, 0), BF
         End If
         If XINI >= XINIA Then
           XINIA = XINI
           YINIA = YINI
         End If
       Next KKL&
       XINI = 0.05 * Picture2.Width + COEFH * (FACT% - FINI%)
       If XINI >= 0.05 * Picture1.Width Then
          Picture2.Line (XINIA, 0.94 * Picture2.Height)-(XINI, YINI), RGB(255, 80, 0), BF
       End If
       Picture2.ForeColor = RGB(0, 150, 160)
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(Label103))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(Label103))
       Picture2.ForeColor = RGB(200, 150, 0)
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(STOMI))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * XVALO(STOMI))
       Picture2.Visible = True
     End If
     Screen.MousePointer = 1
   End If
       
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Picture2.Visible = False
   Screen.MousePointer = 1
End Sub

Private Sub Image2_DblClick()
    Call ACCDIR("CONSTOCK")
End Sub

Private Sub ItStImp_Click()
QUERSTK04.Option4 = True
Call QUERSTK04.Command20_Click
End Sub

Private Sub ItStPant_Click()
QUERSTK04.Option5 = True
Call QUERSTK04.Command20_Click
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

Private Sub List1_DblClick()
   '
   Static MODOTRAZA$
   If MODOTRAZA$ = "" Then
       MODOTRAZA$ = CONTROL$("", "RETRAZA")
       If MODOTRAZA$ = "" Then MODOTRAZA$ = "FIFO"
   End If
   '
   If MODOTRAZA$ = "FIFO" Then
      If EXISTE%("TRAZAB04.EXE") > 0 Then
         Call MENSERR(24, "Aparente Error de Configuración\de Modalidad de Trazabilidad.\   \Consulte !!!")
         OPDISTK04.Show 1
         Call FINAL("")
      End If
   End If
   '
   CIXI% = CODINT%(Text1)
   '
   If CIXI% > 0 Then
     '
     If MODOTRAZA$ <> "FIFO" Then GoTo 15
     '
     STACT = STOCKACT(CIXI%)
     If STACT < 0 Then
       Call MENSERR(24, "Stock Negativo !!!\   \Imposible Consulta de Saldos por Lote.")
       GoTo 99
     End If
     '
     If Abs(STOLOTES(CIXI%) - STACT) >= 0.05 Then
        Call CREALOTES(CIXI%)    ' VERIFICA Y CREA TABLA DE LOTES DE ENTRADA
        Call VERSALIDA(CIXI%)    ' VERIFICA MOVIMIENTOS DE SALIDA
     End If
     '
15   Screen.MousePointer = 1
     SALOTES04.Label4 = Text1
     SALOTES04.Show 1
   End If
   '
99 Screen.MousePointer = 1
End Sub

Private Sub List2_Click()
   If CONTROCLIKK% > 0 Then Exit Sub
   For U& = 1 To List2.ListCount
     If U& - 1 <> List2.ListIndex Then
       List2.Selected(U& - 1) = False
     End If
   Next U&
End Sub

Private Sub List2_DblClick()
   '
   Screen.MousePointer = 11
   CONTROCLIKK% = 1
   List2.Visible = False
10 For KIL& = 2 To List2.ListCount - 1
     AAA$ = Mid$(List2.List(KIL& - 1), 65, 8)
     BBB$ = Mid$(List2.List(KIL&), 65, 8)
     If FECHANUM(AAA$) > FECHANUM(BBB$) Then
       CCC$ = List2.List(KIL& - 1)
       List2.RemoveItem (KIL& - 1)
       List2.AddItem CCC$
       GoTo 10
     End If
   Next KIL&
   List2.Visible = True
   CONTROCLIKK% = 0
   Screen.MousePointer = 1
   '
End Sub

Private Sub LIUBIC_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
End Sub

Private Sub MaDep_Click()
Call QUERSTK04.Command6_Click
End Sub

Private Sub mnuConfiguraciongeneral_Click()
   Call MOVISTK04.Command26_Click

End Sub

Private Sub MovStock_Click()
Call QUERSTK04.Command6_Click
End Sub

Private Sub Option1_Click()
   If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
      Option2.Value = True
   End If
End Sub

Private Sub OtrAccDirec_Click()
'\\\OT-05-0711-WAR-03/11/05///
    'Call MENSERR(24, "No Existen Accesos Directos para esta Aplicación")
5   RTA% = ALTERNA%("\Alta Numeros de Serie\Listado Numeros de Serie\Salida Numeros de Serie")
    If RTA% = 1 Then
      'ACA LLAMA AL OBJETO NUSERI
      VDEF$ = ""
      VDEF$ = Space(128)
10    HOII% = HOY(HO$)
      Call REPLA(VDEF$, MKI$(HOII%), 3, 2)
      OBJ$ = OBJPLA$("NUSERIE", VDEF$)
      OBJANT$ = OBJ$
      If TRIM$(OBJ$) <> "" Then
        CODI1% = CVI(Mid$(OBJ$, 1, 2))
        FA1% = CVI(Mid$(OBJ$, 3, 2))
        ORI1$ = Mid$(OBJ$, 5, 12)
        NUSERI1$ = Mid$(OBJ$, 17, 24)
        OBS1$ = Mid$(OBJ$, 41, 64)
        'VALIDACIONES
        If CODI1% > NUMAS% Or CODI1% <= 0 Then
          Call MENSERR(24, "Código de Articulo no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        If TRIM$(ORI1$) = "" Then
          Call MENSERR(24, "Origen no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        If TRIM$(NUSERI1$) = "" Then
          Call MENSERR(24, "Número de Serie no Válido")
          VDEF = OBJANT$
          GoTo 10
        End If
        'CODIGO + Nº DE SERIE NO SE PUEDEN REPETIR
        Call FILTERFILE("SERIEQUI", "!SERIEQUI", 1, MKI$(CODI1%))
        For j& = 1 To ULTREG&("!SERIEQUI")
          x$ = REGLEIDO$("!SERIEQUI", j&)
          NUMESERI$ = Mid$(x$, 17, 24)
          If UCase(NUMESERI$) = UCase(NUSERI1$) Then
            Call MENSERR(24, "Código + Número de Serie no Válido o Pre-Existente")
            VDEF = OBJANT$
            GoTo 10
          End If
        Next j&
        'GRABAR EN BASE DE DATOS
        y$ = REGBLAN$("SERIEQUI")
        Call REPLA(y$, MKI$(CODI1%), 1, 2)
        Call REPLA(y$, MKI$(FA1%), 3, 2)
        Call REPLA(y$, ORI1$, 5, 12)
        Call REPLA(y$, NUSERI1$, 17, 24)
        Call REPLA(y$, OBS1$, 123, 64)
        Call REGAPP("SERIEQUI", y$)
        '
        Call CIERRARCH("SERIEQUI")
        'PREPARO EL OBJETO PARA LA PROXIMA ALTA
        PRONUSERI$ = ""
        NUSE& = XVALO(Right$(TRIM$(NUSERI1$), 1))
        If NUSE& > 0 Then
          PRONUSE& = NUSE& + 1
          If PRONUSE& = 10 Then PRONUSE& = 1
          PRONUSE1$ = FORMATNUM$(PRONUSE&, "F1.0")
          PRONUSERI$ = Mid$(NUSERI1$, 1, Len(TRIM$(NUSERI1$)) - 1) + PRONUSE1$
        End If
        Call REPLA(OBJANT$, "", 5, 105)
        Call REPLA(OBJANT$, PRONUSERI$, 17, 24)
        VDEF$ = OBJANT$
        GoTo 10
        '
      End If
      GoTo 5
    ElseIf RTA% = 2 Then
      Call LINUSE 'ACA LISTADO
      GoTo 5
    ElseIf RTA% = 3 Then
      Call PROCENUSE 'PROCESO DE SALIDA DE NU. SERIE
      GoTo 5
    End If
'\\\OT-05-0711-WAR-FIN///
End Sub
'\\\OT-05-0711-WAR-07/11/05///
Sub PROCENUSE()
      '1) SELECCION DEL CODIGO
30    OBJ3$ = OBJPLA$("PRESCODIGO", VDEF$)
      If OBJ3$ <> "" Then
        Call FRESHECHO("!INDINUSE")
        CI3% = CVI(Mid$(OBJ3$, 1, 2))
        'VALIDACIONES
        If CI3% <= 0 Then
          Call COMUNI("Codigo no Válido o Inexistente.")
          GoTo 30
        End If
        'AVERIGUO SI EXISTEN Nº DE SERIE PARA ESE ARTICULO
        'PENDIENTES DE SALIDA
        Call FILTERFILE("SERIEQUI", "!SERIEQUI", 1, MKI$(CI3%))
        If ULTREG&("!SERIEQUI") <= 0 Then
          Call COMUNI("No existen Numeros de Serie para \ " & CODEXT$(CI3%) & " - " & DESCRIT0(CI3%) & "")
          GoTo 30
        End If
        JJ& = 0
        Call BLANARCH("!INDINUSE")
        For j& = 1 To ULTREG&("SERIEQUI")
          x$ = REGLEIDO$("SERIEQUI", j&)
          CODI1% = CVI(Mid$(x$, 1, 2))
          If CODI1% = CI3% Then
            REMI3$ = Mid$(x$, 75, 24)
            FACT3$ = Mid$(x$, 99, 24)
            If TRIM$(REMI3$) = "" And TRIM$(FACT3$) = "" Then
              FA1% = CVI(Mid$(x$, 3, 2))
              ORI1$ = Mid$(x$, 5, 12)
              NUSE1$ = Mid$(x$, 17, 24)
              REFE1$ = NUSE1$ + "   " + ORI1$ + "   " + FECHATEX$(CVINT(CVDAT(FA1%)))
              'ACA AGREGO EL NUMERO DE SERIE PENDIENTE AL INDICE
              y$ = REGBLAN$("INDINUSE")
              JJ& = JJ& + 1
              NO% = XVALO(j&) 'NUMERO DE REGISTRO
              Call REPLA(y$, MKI$(NO%), 1, 2)
              Call REPLA(y$, REFE1$, 3, 62)
              Call GRAREG("!INDINUSE", y$, JJ&)
            End If
          End If
        Next j&
        '
        If JJ& <= 0 Then
          Call COMUNI("No existen Numeros de Serie Pendientes para \ " & CODEXT$(CI3%) & " - " & DESCRIT0(CI3%) & "")
          GoTo 30
        End If
        '
        Call SETULTREG("!INDINUSE", JJ&)
        '
32      Call SELECHO("!INDINUSE")
        NREG1& = XVALO(VALACT1$("!INDINUSE"))
        If NREG1& <= 0 Then
          GoTo 30
        End If
        '
35      x$ = REGLEIDO$("SERIEQUI", NREG1&)
        CI4% = CVI(Mid$(x$, 1, 2))
        FA4% = CVI(Mid$(x$, 3, 2))
        ORI4$ = Mid$(x$, 5, 12)
        NUSE4$ = Mid$(x$, 17, 24)
        OBS4A$ = Mid$(x$, 123, 32)
        OBS4B$ = Mid$(x$, 155, 64)
        '
        Call REPLA(VDEF$, MKI$(CI4%), 1, 2)
        Call REPLA(VDEF$, MKI$(FA4%), 3, 2)
        Call REPLA(VDEF$, ORI4$, 5, 12)
        Call REPLA(VDEF$, NUSE4$, 17, 24)
        Call REPLA(VDEF$, OBS4A$, 41, 32)
        Call REPLA(VDEF$, OBS4B$, 73, 32)
        '
        HOII% = HOY(HO$)
        Call REPLA(VDEF$, MKI$(HOII%), 105, 2)
        '
37      OBJ4$ = OBJPLA$("SANUSE", VDEF$)
        OBJANT4$ = OBJ4$
        If OBJ4$ <> "" Then
          'VALIDACIONES
          REMI4$ = Mid$(OBJ4$, 109, 24)
          FACT4$ = Mid$(OBJ4$, 133, 24)
          If TRIM$(REMI4$) = "" Or TRIM$(FACT4$) = "" Then
            Call COMUNI("Número de Remito o Factura no Valido.")
            VDEF$ = OBJANT4$
            GoTo 37
          End If
          NC4% = CVI(Mid$(OBJ4$, 107, 2))
          If NC4% <= 0 Then
            Call COMUNI("Cliente / Destino no Válido.")
            VDEF$ = OBJANT4$
            GoTo 37
          End If
          'EDITA EL REGISTRO
          FS4% = CVI(Mid$(OBJ4$, 105, 2))
          '
          x$ = REGLEIDO$("SERIEQUI", NREG1&)
          Call REPLA(x$, MKI$(FS4%), 41, 2)
          Call REPLA(x$, MKI$(NC4%), 43, 2)
          Call REPLA(x$, RASOCLI$(NC4%), 45, 30)
          Call REPLA(x$, REMI4$, 75, 24)
          Call REPLA(x$, FACT4$, 99, 24)
          Call GRAREG("SERIEQUI", x$, NREG1&)
          '
        Else
          VDEF$ = ""
          GoTo 32
        End If
        '
      End If
End Sub
Sub LINUSE()
    'ACA LLAMA AL OBJETO DESHASCOD2
      VDEF$ = ""
20    HOII% = HOY(HO$)
      ULTCO% = NUMAS%
      Call REPLA(VDEF$, MKI$(1), 1, 2)
      Call REPLA(VDEF$, MKI$(ULTCO%), 3, 2)
      Call REPLA(VDEF$, Chr$(0), 5, 1)
      Call REPLA(VDEF$, MKI$(COMESACT%), 6, 2)
      Call REPLA(VDEF$, MKI$(HOII%), 8, 2)
      OBJ1$ = OBJPLA$("DESHASCOD2", VDEF$)
      If TRIM$(OBJ1$) <> "" Then
        DESA% = CVI(Mid$(OBJ1$, 1, 2))
        HASA% = CVI(Mid$(OBJ1$, 3, 2))
        DESF% = CVI(Mid$(OBJ1$, 6, 2))
        HASF% = CVI(Mid$(OBJ1$, 8, 2))
        TIP% = Asc%(Mid$(OBJ1$, 5, 1))
        '
        Call COPYSTRU("SERIEQUI", "LINUSERI")
        Z$ = REGBLAN$("LINUSERI")
        '
        JJ& = 0
        For j& = 1 To ULTREG&("SERIEQUI")
          x$ = REGLEIDO$("SERIEQUI", j&)
          CI1% = CVI(Mid$(x$, 1, 2))
          FA1% = CVI(Mid$(x$, 3, 2))
          If TIP% > 0 Then
            If TIP% <> TIMATE%(CI1%) Then
              GoTo 25
            End If
          End If
          If CI1% >= DESA% Then
            If CI1% <= HASA% Then
              If FA1% >= DESF% Then
                If FA1% <= HASF% Then
                  ORI1$ = Mid$(x$, 5, 12)
                  NUSE1$ = Mid$(x$, 17, 24)
                  FESA1% = CVI(Mid$(x$, 41, 2)) 'FECHA DE SALIDA
                  CTE1% = CVI(Mid$(x$, 43, 2)) 'CLIENTE
                  DESTI1$ = Mid$(x$, 45, 30) 'DESTINO
                  REMI1$ = Mid$(x$, 75, 24) 'REMITO
                  FACT1$ = Mid$(x$, 99, 24) 'FACTURA
                  OBS1$ = Mid$(x$, 123, 64) 'OBSERVACIONES
                  Call REPLA(y$, MKI$(CI1%), 1, 2)
                  Call REPLA(y$, MKI$(FA1%), 3, 2)
                  Call REPLA(y$, ORI1$, 5, 12)
                  Call REPLA(y$, NUSE1$, 17, 24)
                  Call REPLA(y$, MKI$(FESA1%), 41, 2)
                  Call REPLA(y$, MKI$(CTE1%), 43, 2)
                  Call REPLA(y$, DESTI1$, 45, 30)
                  Call REPLA(y$, REMI1$, 75, 24)
                  Call REPLA(y$, FACT1$, 99, 24)
                  Call REPLA(y$, OBS1$, 123, 64)
                  JJ& = JJ& + 1
                  Call GRAREG("LINUSERI", y$, JJ&)
                End If
              End If
            End If
          End If
25      Next j&
        Call SETULTREG("LINUSERI", JJ&)
        Call CIERRARCH("LINUSERI")
        If JJ& > 0 Then
          Call FILESORT("LINUSERI", "LINUSERI", 3, 3, "ASC")
          Call FILESORT("LINUSERI", "LINUSERI", 1, 1, "ASC")
          Call FINAL("*LINUSERI")
        Else
          Call COMUNI("No existen Numeros de Serie Asignados")
        End If
      End If
End Sub

Private Sub p_Click()

End Sub

'\\\OT-05-0711-WAR-FIN///
Private Sub ParFabr_Click()
Call QUERSTK04.Command3_Click
End Sub

Private Sub recepciondeproveedores_Click()
   Call MOVISTK04.Command28_Click
End Sub

Private Sub remitirmercaderias_Click()
    Call MOVISTK04.Command23_Click
End Sub

Private Sub remitoconsignacion_Click()
    Call MOVISTK04.Command22_Click
End Sub

Private Sub salir_Click()
Call Command1_Click
End Sub

Private Sub SitCob_Click()
Call Command4_Click
End Sub

Private Sub Stockminimorecomendado_Click()
   Call ESTASTK04.Command7_Click
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    Label101.Caption = ""
    Label102.Caption = ""
    Label103.Caption = ""
    Label201.Caption = ""
    Label202.Caption = ""
    Label203.Caption = ""
    Label301.Caption = ""
    Label302.Caption = ""
    Label303.Caption = ""
    Label401.Caption = ""
    Label404.Caption = ""
    Label402.Caption = ""
    Label403.Caption = ""
    List1.Clear
    List2.Clear
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
    If Option1.Value = True Then
       ASEL$ = "LIPRE001"
    End If
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
     Screen.MousePointer = 11
     Call ABREPEDCLI
     SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
     SQLX$ = SQLX$ + " FROM PEDENCA "
     SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
     SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
     SQLX$ = SQLX$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
     SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
     Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     JJ& = 0
     With PediModifTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
             NRODES& = !NroPed
           TTXX$ = MKS$(NRODES&)
             NOCO$ = !NroOcom
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
             FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
             NCLIE% = !NROCLIE
           TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
           JJ& = JJ& + 1
           Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("INDIPEP", JJ&)
     Call CIERRARCH("INDIPEP")
     '
     Screen.MousePointer = 1
     YAINDI% = 1
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 x$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(x$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(x$, 4))
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
     GoTo 2110
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
     EDITFORM.Command3.Enabled = False
'     EDITFORM.Command5.Enabled = False
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
10   List2.Clear
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
     List2.AddItem LI2TEX$
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
            CARECEP = XVALO(!CANTNOM) - XVALO(!CanToApro)
            If CARECEP > 0 Then
              NCI% = XVALO(!NumProv)
              REMIPRO$ = SAFETEXT$(!NumRem_Prov)
              NROBOL& = XVALO(!NumeBore)
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
           NROX& = XVALO(!NroPed)
           NCI% = XVALO(!NROCLIE)
           CAPEN = XVALO(!CanPed) - XVALO(!CantEnt)
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
           NUORSE& = XVALO(!N_OrServ)
           If NUORSE& < 1 Then
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = XVALO(!Cant_Ocom) * COEFCAN
             CAREC = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = XVALO(!Npro_Ocom)
               NROX& = XVALO(!Nume_Ocom)
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
             CACOM = XVALO(!Cant_Aten) * COEFCAN
             CAREC = XVALO(!Acuen_Aten) * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = XVALO(!Npro_Aten)
               NROX& = XVALO(!Nume_Aten)
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
           CACOM = XVALO(!CanProy)
           CAREC = XVALO(!CanApro)
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             NroPed& = XVALO(!NUMEORFA)
             FEX = CVINT(!FechEnSol)
             REFE$ = SAFETEXT$(!Referen)
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
           TORE# = XVALO(!CantRes)
           CACO# = XVALO(!CantCons)
           '
           If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
              TORE# = TORE# / 1000
              CACO# = CACO# / 1000
           End If
           '
           SIGNO% = Sgn(CVS(Mid$(x$, 75, 4)))
           SARE = TORE# - CACO#
           If SARE * SIGNO% < 0 Then
             SARE = 0
           End If
           CAPEN = SARE
           If Abs(CAPEN) > 0 Then
             NORFA$ = TRIM$(SAFETEXT$(!IdSubOr))
             IDLOTXX$ = TRIM$(SAFETEXT$(!IdenLote))
             OBSER$ = ""
             If IDLOTXX$ <> "" Then OBSER$ = "Lote " + IDLOTXX$
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
               CICON% = XVALO(OrdeFabr!COD_INTE)
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
             Call REPLA(LI2TEX$, OBSER$, 84, 24)
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
           TORE# = XVALO(!Cant_Entreg)
           CACO# = XVALO(!Cant_Rendid)
           CAPEN = TORE# - CACO#
           CAPENENT# = XVALO(!Cant_Asig) - XVALO(!Cant_Entreg)
           NUOC& = XVALO(Mid$(!Nro_Ocom, 4, 6))
           '
           If Abs(CAPEN) > 0 Then
              NUREX$ = "RC-" + TRIM$(!NRemi_Consig)
              NORFA$ = SAFETEXT$(!Nro_Ocom)
              NPRO% = XVALO(!Npro_Ocom)
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
                NORFA$ = SAFETEXT$(!Nro_Ocom)
                NPRO% = XVALO(!Npro_Ocom)
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
        x$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(x$, 2)) <> CI% Then
           x$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", x$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(x$, 3, 24)) <> "" Then
           y$ = Mid$(x$, 3, 24)
           LIUBIC.AddItem y$
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
     Label404.Caption = TRIM$(CSTRING$(MKS$(XVALO(Label103) - XVALO(Label202)), 4, 12, 1, 2))
     '
     DoEvents
     '
Return
'
End Sub
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

Private Sub Timer1_Timer()
    If Me.Enabled = False Then
      On Error Resume Next
      For U& = 0 To Controls.Count - 1
        If Left$(UCase$(Me.Controls(U&).Name), 5) = "COMMA" Then
           Me.Controls(U&).Refresh
        End If
      Next U&
      On Error GoTo 0
    End If
End Sub

Private Sub transferencias_Click()
   Call MOVISTK04.Command24_Click
End Sub

Private Sub transferenciastock_Click()
   TRANSTK04.Command3_Click
End Sub

Private Sub valeconsumo_Click()
    Call MOVISTK04.Command21_Click
End Sub

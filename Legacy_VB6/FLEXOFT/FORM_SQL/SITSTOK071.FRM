VERSION 5.00
Begin VB.Form SITSTOK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   ClipControls    =   0   'False
   Icon            =   "SITSTOK07.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.CommandButton Command12 
      Caption         =   "Conver-Repa"
      Height          =   195
      Left            =   8640
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "MIGRA-REPA"
      Height          =   195
      Left            =   7200
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
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
         Picture         =   "SITSTOK07.frx":014A
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
         Picture         =   "SITSTOK07.frx":07B4
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
         Picture         =   "SITSTOK07.frx":08FE
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
         Picture         =   "SITSTOK07.frx":0C08
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
         Picture         =   "SITSTOK07.frx":0F12
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
         Picture         =   "SITSTOK07.frx":105C
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
         Picture         =   "SITSTOK07.frx":11A6
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
         Picture         =   "SITSTOK07.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1560
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "SITSTOK07.frx":18F2
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
      Picture         =   "SITSTOK07.frx":3814
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Visualizar Ordenes de Venta Pendientes de Entrega"
      Top             =   210
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
         Picture         =   "SITSTOK07.frx":3B1E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Muestra / Oculta Situación de Cobertura"
         Top             =   630
         Width           =   560
      End
      Begin VB.Image Image1 
         Height          =   250
         Left            =   105
         Picture         =   "SITSTOK07.frx":3E28
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
         Left            =   3105
         TabIndex        =   48
         Top             =   675
         Width           =   1170
      End
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
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   225
      Left            =   630
      TabIndex        =   75
      Top             =   105
      Width           =   435
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   225
      Left            =   105
      TabIndex        =   74
      Top             =   105
      Width           =   435
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
         Begin VB.Menu GesRese 
            Caption         =   "Gestión de Reservas"
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
      Begin VB.Menu MaDep 
         Caption         =   "Material en Depósitos"
      End
      Begin VB.Menu GesReser 
         Caption         =   "Gestión de Reservas"
      End
   End
   Begin VB.Menu AcceDire 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MovStock 
         Caption         =   "Movimientos de Stock"
      End
      Begin VB.Menu php25 
         Caption         =   "-"
      End
      Begin VB.Menu ParFabr 
         Caption         =   "Partes de Fabricación"
      End
      Begin VB.Menu php7 
         Caption         =   "-"
      End
      Begin VB.Menu OtrAccDirec 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Confi 
      Caption         =   "Configuración"
   End
   Begin VB.Menu Ayud 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "SITSTOK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%

Private Sub Abrir_Click()
    Call Command2_Click
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
   CONPED07.Show 1
   Command10.Enabled = True
End Sub



Private Sub Command11_Click()
'Sub MIGRASTOCKS()
  'ESTE BOTON SE AGREGO PARA  MIGRAR LA TABLA MOVISTO HACIA MIGRAREPA PARA DOSIVAC
  '\\\OT-08-0819-OD-08/09/08///
  Exit Sub
  '
  Command11.Enabled = False
  Call ABRECONEXION
  'Picture1.ToolTipText = "Del - Movimientos de Stock"
  On Error Resume Next
   FLEXCONN.EXECUTE (FILTSQL$ ("DROP TABLE MOVIREPA"))
  On Error GoTo 0
  Call GENERBASE("STOCKS")
  Call APERBASDAT("STOCKS")
  '
  'Picture1.ToolTipText = "Conversión de Movimientos de Stocks"
  '
  Screen.MousePointer = 11

  FIN& = ULTREG&("MOVIREPA")
  DOCANT$ = ""
  PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
  '
  sqlx$ = "SELECT * FROM MOVIREPA "
  Dim MOVIREPATMP As ADODB.Recordset
  Set MOVIREPATMP = New ADODB.Recordset
  MOVIREPATMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MOVIREPA", U&)
    Call TRACE(20, U&, FIN&)
    'Label1 = U&
    'LECTURA
      FEC% = CVI(Mid$(X$, 1, 2)) 'FECHA DE MOVIMIENTO
      CI% = CVI(Mid$(X$, 3, 2)) 'CODIGO DE PIEZA
      LOT$ = Mid$(X$, 5, 16) 'LOTE / PARTIDA
      COMO$ = Mid$(X$, 21, 2) 'CODIGO DE MOVIMIENTO
      DOC1$ = Mid$(X$, 23, 10) 'DOCUMENTO PRIMARIO ASOCIADO
      DOC1LAR$ = DOC1$
      If COMO$ = "RP" Then
         NUMEDO$ = TRIM$(Mid$(DOC1$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC1LAR$ = Left$(DOC1$, 2) + "-X-" + PREFIREM$ + "-" + NUMEDO$
      End If
      DOC2$ = Mid$(X$, 33, 10) 'DOCUMENTO SECUNDARIO
      DOC2LAR$ = DOC2$
      If Left$(DOC2$, 2) = "FC" Then
         NUMEDO$ = TRIM$(Mid$(DOC2$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC2LAR$ = Left$(DOC2$, 2) + "-0001-" + NUMEDO$ + "-A"
      End If
      REFE$ = Mid$(X$, 43, 30) 'REFERENCIA
      VAUN# = CVD(Mid$(X$, 73, 8)) 'VALOR UNITARIO
      MONE1$ = Mid$(X$, 81, 2) 'MONEDA
      DEP% = Asc(Mid$(X$, 83, 1)) 'DEPOSITO
      NUCLI% = CVI(Mid$(X$, 84, 2)) 'NUEMRO DE CLIENTE
      NUPRO% = CVI(Mid$(X$, 86, 2)) 'NUMERO DE PROVEEDOR
      CAEN = CVD(Mid$(X$, 88, 8)) 'CANTIDAD ENTRADA
      CASA = CVD(Mid$(X$, 96, 8)) 'CANTIDAD SALIDA
      CASALD = CVD(Mid$(X$, 104, 8)) 'CANTIDAD SALDO
      FILL$ = Mid$(X$, 112, 10) 'FILLER NO IDENTIFICADO
      FEMO% = CVI(Mid$(X$, 122, 2)) 'FECHA DE MOVIMIENTO
      USU% = Asc(Mid$(X$, 126, 1)) 'USUARIO
    'GRABACIÓN
    With MOVIREPATMP
      .AddNew   ' YA CONVERTIDO
      ON ERROR RESUME NEXT
      !CodiEmpr = Codiemp%
      ON ERROR GOTO 0
        !fechMov = CVDAT(FEC%)
        !Cod_Inte = CI%
        !Cod_Exte = " " 'CODEXT$(CI%) - eliminado para ganar velocidad
        !Descrip = " " 'DESCRIT0(CI%) - eliminado para ganar velocidad
        !IdenLote = LOT$
        !CodiMovi = COMO$
        !DoPriCor = DOC1$
        !DOPRILAR = DOC1LAR$
           If DOC1$ <> DOCANT$ Then
             NORI% = 0
             DOCANT$ = DOC1$
           End If
        !DoSecCor = DOC2$
        !DOSECLAR = DOC2LAR$
        !NumRem_Cli = "" 'POR CONVERSIÓN
           NORI% = NORI% + 1
        !NuOrItem = NORI%
        !NUPEDCLI = CVS(Left$(FILL$, 4))      'POR CONVERSIÓN
          NOPXX% = CVI(Mid$(FILL$, 5, 2))
          If NOPXX% = 8224 Then NOPXX% = 0
        !NumeOrdPed = NOPXX%
        !ReferCor = REFE$
        !ValoUnit = VAUN
        !MoneVal = XVALO(MONE1$)
        !DepoMovi = DEP%
        !Nume_Clie = NUCLI%
        !Nume_Prov = NUPRO%
        !CantIngre = CAEN
        !Cantsalid = CASA
        !CantSaldo = CASALD
        !FeReMovi = CVDAT(FEMO%)
      .Update
    End With
    '
    Call TRACE(20, U&, FIN&)
  Next U&
  MOVIREPATMP.Close
  Set MOVIREPATMP = Nothing
  '
  Screen.MousePointer = 1
  Command11.Enabled = True

  '
  '\\\OT-08-0819-OD-FIN///
  '
End Sub
Private Sub Command12_Click()
    '\\\OT-08-0839-OD-12/08/08///
    '
    'ESTE BOTON SE AGREGO PAARA PASAR LOS DATOS DE REPARACIONES QUE ESTABAN EN EL MOVISTO AL MOVIREPA
    'PARA DOSIVAC
    Exit Sub
    Command12.Enabled = False
    Screen.MousePointer = 11
    FLIM$ = FETEXCOR1$(FECHANUM("22/08/08"))
    '
    Call ABRESTOCKS
    'CONSULTA AL MOVISTO CON MOVIMIENTOS SUPERIORES O IGUALES A 22/08/08
    '
    sqlx$ = "SELECT * FROM MOVISTO"
    sqlx$ = sqlx$ + " WHERE FechMov >= '" & FLIM$ & "'"
    sqlx$ = sqlx$ + " AND NuPedCli > 0"
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NumPedido& = XVALO(!NUPEDCLI)
          ResObtenido = VALOBADA("PEDENCA", "VENREP", "NroPed = " & NumPedido&)
          If ResObtenido = "R" Then 'SOLO LOS QUE SON REPARACIONES
            Dim VALORCA(128), NOMBREDECAMPO(128) 'DIMENSIONE VECTORES
            CANCAMPOS = .Fields.Count 'CUENTO LA CANTIDAD DE CAMPOS
            For I% = 1 To CANCAMPOS 'LOS RECORRO
              NOMBREDECAMPO(I%) = .Fields(I% - 1).Name 'VOY ASIGNANDO EN EL VECTOR LOS NOMBRES DE LOS CAMPOS
              VALORCA(I%) = .Fields(I% - 1).Value 'ASIGNO EL CONTENIDO DE CADA CAMPO EN EL OTRO VECTOR
            Next I%
            'ABRO EL RECORDSET DE MOVIREPA
            sqlx$ = "SELECT * FROM MOVIREPA" 'HABRO EL MOVIREPA PARA REALIZAR EL ADDNEW
            Dim RS1 As ADODB.Recordset
            Set RS1 = New ADODB.Recordset
            RS1.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            With RS1
              .AddNew   ' YA CONVERTIDO
              ON ERROR RESUME NEXT
              !CodiEmpr = Codiemp%
              ON ERROR GOTO 0
                For I% = 1 To CANCAMPOS 'DE 1 A LA CANTIDAD DE CAMPOS
                  NOCAM$ = NOMBREDECAMPO(I%) 'ASIGNO EL NOMBRE DEL CAMPO
                  For J% = 1 To .Fields.Count
                    If .Fields(J% - 1).Name = NOCAM$ Then ' SI COINCIDE EL NOMBRE CAMPO CON NOCAM$
                      .Fields(J% - 1).Value = VALORCA(I%) 'A CADA CAMPO LE ASIGNO EL VALOR
                    End If
                  Next J%
                Next I%
              .Update
              MOVITMP!DoPriCor = "MIGRADO" 'MARCA PARA DESPUES BORRAR DEL MOVISTO LOS MIGRADOS
              MOVITMP.Update
            End With
          End If
          .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    'ELIMINO LOS MOVIMIENTOS  DEL MOVISTO PASADOS AL MOVIREPA
    '
    sqlx$ = "DELETE  FROM MOVISTO "
    sqlx$ = sqlx$ + " WHERE DOPRICOR = 'MIGRADO'"
     FLEXCONN.EXECUTE (FILTSQL$ (sqlx$))
    '
    Command12.Enabled = True
    Screen.MousePointer = 1

    Call COMUNI("Proceso Terminado")
    
    '\\\OT-08-0839-OD-FIN///
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
         OPDISTK07.Show 1
       Else
         Call MENSERR(24, "La Opción Elegida Requiere Privilegios de Supervisor")
    End If
    Command27.Enabled = True
End Sub

Sub Command3_Click()
   COMMAND3.Enabled = False
   FIKARDEX07.Text1 = Text1
   FIKARDEX07.Show 1
   COMMAND3.Enabled = True
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
   MOVISTK07.Show 1
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
    QUERSTK07.Show 1
    Command9.Enabled = True
End Sub

Private Sub Confi_Click()
    Call Command27_Click
End Sub

Private Sub CoPreVe_Click()
    Call Command7_Click
End Sub

Private Sub DiStImp_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command5_Click
End Sub

Private Sub DiStPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command5_Click
End Sub

Private Sub ExFiImp_Click()
    QUERSTK07.Option4 = True
Call QUERSTK07.Command19_Click
End Sub

Private Sub ExFiPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command19_Click
End Sub

Private Sub ExistyUbi_Click()
    Call QUERSTK07.Command2_Click
End Sub

Private Sub ExNegImpr_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command30_Click
End Sub

Private Sub ExNegPan_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command30_Click
End Sub

Private Sub FichaKardex_Click()
    Call QUERSTK07.Command9_Click
End Sub

Private Sub FichKar_Click()
    Call Command3_Click
End Sub

Private Sub Form_Activate()
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
      '\\\OT-08-0643-OD-20/08/08///
      'La linea de abajo se comento segun O/T. 08-0643 para que la informacion _
       aparezcan ordenada con respecto a los encabezados (Queda para ver si _
       queda finallmente así)
      '
      'Label15 = " Comprob.           Referencia                                                Entra          Sale         Fecha      Observ."
      '\\\OT-08-0643-OD-Fin///

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
   Call FINAL("")
End Sub

Private Sub GesRese_Click()
    Call QUERSTK07.Command14_Click
End Sub

Private Sub GesReser_Click()
    Call QUERSTK07.Command7_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture2.Visible = False
   Screen.MousePointer = 1
End Sub

Private Sub Image2_DblClick()
    Call ACCDIR("CONSTOCK")
End Sub

Private Sub ItStImp_Click()
    QUERSTK07.Option4 = True
    Call QUERSTK07.Command20_Click
End Sub

Private Sub ItStPant_Click()
    QUERSTK07.Option5 = True
    Call QUERSTK07.Command20_Click
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
   CIXI% = CODINT%(Text1)
   NDEP% = XVALO(Mid$(List1.TEXT, 61, 4))
   If ESTRAZABLE(CIXI%) = 1 Then
        SALOTES07.Label9 = TRIM$(Str$(NDEP%))
        SALOTES07.Label4 = Text1
        Call ARMALILOT(CIXI%, NDEP%, 0, 1, 1)
        Call SALOTES07.MUESTRASALOT
        SALOTES07.Show 1
        Exit Sub
      Else
        Call COMUNI("Imposible Consulta por Lotes en Material no Trazable.")
   End If
   '
99 Screen.MousePointer = 1
   '
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
    Call QUERSTK07.Command6_Click
End Sub

Private Sub MovStock_Click()
    Call Command6_Click
End Sub

Private Sub Option1_Click()
   If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
      Option2.Value = True
   End If
End Sub

Private Sub OtrAccDirec_Click()
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
        For J& = 1 To ULTREG&("!SERIEQUI")
          X$ = REGLEIDO$("!SERIEQUI", J&)
          NUMESERI$ = Mid$(X$, 17, 24)
          If UCase(NUMESERI$) = UCase(NUSERI1$) Then
            Call MENSERR(24, "Código + Número de Serie no Válido o Pre-Existente")
            VDEF = OBJANT$
            GoTo 10
          End If
        Next J&
        'GRABAR EN BASE DE DATOS
        Y$ = REGBLAN$("SERIEQUI")
        Call REPLA(Y$, MKI$(CODI1%), 1, 2)
        Call REPLA(Y$, MKI$(FA1%), 3, 2)
        Call REPLA(Y$, ORI1$, 5, 12)
        Call REPLA(Y$, NUSERI1$, 17, 24)
        Call REPLA(Y$, OBS1$, 123, 64)
        Call REGAPP("SERIEQUI", Y$)
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
        For J& = 1 To ULTREG&("SERIEQUI")
          X$ = REGLEIDO$("SERIEQUI", J&)
          CODI1% = CVI(Mid$(X$, 1, 2))
          If CODI1% = CI3% Then
            REMI3$ = Mid$(X$, 75, 24)
            FACT3$ = Mid$(X$, 99, 24)
            If TRIM$(REMI3$) = "" And TRIM$(FACT3$) = "" Then
              FA1% = CVI(Mid$(X$, 3, 2))
              ORI1$ = Mid$(X$, 5, 12)
              NUSE1$ = Mid$(X$, 17, 24)
              REFE1$ = NUSE1$ + "   " + ORI1$ + "   " + FECHATEX$(CVINT(CVDAT(FA1%)))
              'ACA AGREGO EL NUMERO DE SERIE PENDIENTE AL INDICE
              Y$ = REGBLAN$("INDINUSE")
              JJ& = JJ& + 1
              NO% = XVALO(J&) 'NUMERO DE REGISTRO
              Call REPLA(Y$, MKI$(NO%), 1, 2)
              Call REPLA(Y$, REFE1$, 3, 62)
              Call GRAREG("!INDINUSE", Y$, JJ&)
            End If
          End If
        Next J&
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
35      X$ = REGLEIDO$("SERIEQUI", NREG1&)
        CI4% = CVI(Mid$(X$, 1, 2))
        FA4% = CVI(Mid$(X$, 3, 2))
        ORI4$ = Mid$(X$, 5, 12)
        NUSE4$ = Mid$(X$, 17, 24)
        OBS4A$ = Mid$(X$, 123, 32)
        OBS4B$ = Mid$(X$, 155, 64)
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
          X$ = REGLEIDO$("SERIEQUI", NREG1&)
          Call REPLA(X$, MKI$(FS4%), 41, 2)
          Call REPLA(X$, MKI$(NC4%), 43, 2)
          Call REPLA(X$, RASOCLI$(NC4%), 45, 30)
          Call REPLA(X$, REMI4$, 75, 24)
          Call REPLA(X$, FACT4$, 99, 24)
          Call GRAREG("SERIEQUI", X$, NREG1&)
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
        For J& = 1 To ULTREG&("SERIEQUI")
          X$ = REGLEIDO$("SERIEQUI", J&)
          CI1% = CVI(Mid$(X$, 1, 2))
          FA1% = CVI(Mid$(X$, 3, 2))
          If TIP% > 0 Then
            If TIP% <> TIMATE%(CI1%) Then
              GoTo 25
            End If
          End If
          If CI1% >= DESA% Then
            If CI1% <= HASA% Then
              If FA1% >= DESF% Then
                If FA1% <= HASF% Then
                  ORI1$ = Mid$(X$, 5, 12)
                  NUSE1$ = Mid$(X$, 17, 24)
                  FESA1% = CVI(Mid$(X$, 41, 2)) 'FECHA DE SALIDA
                  CTE1% = CVI(Mid$(X$, 43, 2)) 'CLIENTE
                  DESTI1$ = Mid$(X$, 45, 30) 'DESTINO
                  REMI1$ = Mid$(X$, 75, 24) 'REMITO
                  FACT1$ = Mid$(X$, 99, 24) 'FACTURA
                  OBS1$ = Mid$(X$, 123, 64) 'OBSERVACIONES
                  Call REPLA(Y$, MKI$(CI1%), 1, 2)
                  Call REPLA(Y$, MKI$(FA1%), 3, 2)
                  Call REPLA(Y$, ORI1$, 5, 12)
                  Call REPLA(Y$, NUSE1$, 17, 24)
                  Call REPLA(Y$, MKI$(FESA1%), 41, 2)
                  Call REPLA(Y$, MKI$(CTE1%), 43, 2)
                  Call REPLA(Y$, DESTI1$, 45, 30)
                  Call REPLA(Y$, REMI1$, 75, 24)
                  Call REPLA(Y$, FACT1$, 99, 24)
                  Call REPLA(Y$, OBS1$, 123, 64)
                  JJ& = JJ& + 1
                  Call GRAREG("LINUSERI", Y$, JJ&)
                End If
              End If
            End If
          End If
25      Next J&
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
'\\\OT-05-0711-WAR-FIN///
Private Sub ParFabr_Click()
    Call QUERSTK07.Command3_Click
End Sub

Private Sub Salir_Click()
    Call Command1_Click
End Sub

Private Sub SitCob_Click()
    Call Command4_Click
End Sub

Private Sub Text1_Change()
    '
    On Error GoTo 0
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
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    If CI% > 0 Then
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

Private Sub text1_lostfocus()
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
     sqlx$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
     sqlx$ = sqlx$ + " FROM PEDENCA "
     sqlx$ = sqlx$ + " INNER JOIN  PEDDETA "
     sqlx$ = sqlx$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
     sqlx$ = sqlx$ + " WHERE PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
     sqlx$ = sqlx$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     sqlx$ = sqlx$ + " AND PEDENCA.Status < 9 "
     sqlx$ = sqlx$ + " ORDER BY PEDENCA.NroPed ASC "
     'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim PediModifTemp As ADODB.Recordset
     Set PediModifTemp = New ADODB.Recordset
     PediModifTemp.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockReadOnly, adCmdText
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
             NCLIE% = !NroClie
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
             JJ& = JJ& + 1
           Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
         .MoveNext
         Loop
       End If
     End With
     PediModifTemp.Close
     Set PediModifTemp = Nothing
     
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
     EDITFORM.COMMAND3.Enabled = False
     EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$, HOYY, MODODOSI%
     '
     Screen.MousePointer = 11
     If CTX% <> 1 Then
        IDENSEDE$ = CONTROL$("", "IDENSEDE")
        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        HOYY = HOY(HOS$)
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
        CTX% = 1
        Call ABREMASTER
        Call ABRESTOCKS
        Call ABREORFAB
        Call ABRESERVA
        Call ABREPEDCLI
        Call ABRECOMPRAS
     End If
     '
     On Error GoTo 0
Label14 = XVALO(Label14) + 1
Label14.Refresh
Label17 = CI%
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
     List1.Clear
     List2.Clear
     STOPROP = 0
     If CI% > 0 Then
       For U& = 1 To NUSU%
         CAPAR = STODEPOS(CI%, COSU%(U&)) '        XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
         If Abs(CAPAR) >= 0.05 Then
           XXX$ = ADEPO$(U&) + CSTRING$(MKS$(CAPAR), 4, 8, 1, 2)
           If DISTO(U&) < 0 Then
               XXX$ = XXX$ + "(x)"
             Else
               STOPROP = STOPROP + CAPAR
           End If
           XXX$ = AJUSTI$(XXX$, 60) + FORMATNUM$(COSU%(U&), "F4.0")
           List1.AddItem XXX$
         End If
       Next U&
       GoSub PRESEDIS
     End If
     '
Exit Sub
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
     sqlx$ = " SELECT * FROM BOLRECEP "
     sqlx$ = sqlx$ + " WHERE Status = 0 "
     sqlx$ = sqlx$ + "AND Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + " ORDER BY NumeBoRe "
     '
     'Set COBRECEPTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
     Dim COBRECEPTMP As ADODB.Recordset
     Set COBRECEPTMP = New ADODB.Recordset
     COBRECEPTMP.CursorType = adOpenDynamic
     COBRECEPTMP.LockType = adLockReadOnly
     COBRECEPTMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
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
     sqlx$ = " SELECT * FROM PEDDETA "
     sqlx$ = sqlx$ + "WHERE CodiInt = " & CI% & " "
     sqlx$ = sqlx$ + "AND VenRep = 'V' "
     sqlx$ = sqlx$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     If MODODOSI% = 1 Then
          sqlx$ = sqlx$ + " AND StaDespa < 2 "
          sqlx$ = sqlx$ + "AND AcuDespa < (CanPed - 0.005) "
        Else
          sqlx$ = sqlx$ + "AND CantEnt < (CanPed - 0.005) "
     End If
     sqlx$ = sqlx$ + "ORDER BY NroPed ASC "
     '
     Dim PEDTEMP As ADODB.Recordset
     Set PEDTEMP = New ADODB.Recordset
     PEDTEMP.CursorType = adOpenDynamic
     PEDTEMP.LockType = adLockReadOnly
     PEDTEMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
     
     
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           NROX& = !NroPed
           NCI% = !NroClie
           CAPEN = !CanPed - !CantEnt
           
           SALPREP& = 0
           If MODODOSI% = 1 Then
              FEMIPED% = CVINT(!FECHEMIS)
              If FEMIPED% >= FECHANUM("07/11/05") Then
                 CAPEN = XVALO(!CanPed) - XVALO(!AcuDespa)
                 SALPREP& = XVALO(!AcuPrepa) - XVALO(!AcuDespa)
              End If
           End If
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
               If SALPREP > 0 Then
                 If OBSER$ <> "" Then OBSER$ = OBSER$ + "/ "
                 OBSER$ = OBSER$ + "Pr." + TRIM$(Str$(SALPREP&))
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
     sqlx$ = " SELECT * FROM OCOMDETA "
     sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     sqlx$ = sqlx$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     sqlx$ = sqlx$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     sqlx$ = sqlx$ + "ORDER BY Nume_Ocom ASC "
     '
     'Set OcoTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
     Dim OcoTemp As ADODB.Recordset
     Set OcoTemp = New ADODB.Recordset
     OcoTemp.CursorType = adOpenDynamic
     OcoTemp.LockType = adLockReadOnly
     OcoTemp.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
          
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !coef_unids
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
22   Call ABRECOMPRAS
     sqlx$ = " SELECT * FROM AUTOENT "
     sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + "AND Stat_Aten < 8 "   ' para que no este anulada ni cumplida
     sqlx$ = sqlx$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
     sqlx$ = sqlx$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
     sqlx$ = sqlx$ + "ORDER BY Nume_Aten ASC "
     '
     'Set AtenTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
     Dim AtenTemp As ADODB.Recordset
     Set AtenTemp = New ADODB.Recordset
     AtenTemp.CursorType = adOpenDynamic
     AtenTemp.LockType = adLockReadOnly
     AtenTemp.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
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
     sqlx$ = " SELECT * FROM ORDEFABR WITH(NOLOCK) "
     sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     sqlx$ = sqlx$ + "AND CanApro < (CanProy - 0.005) "
     sqlx$ = sqlx$ + "ORDER BY NumeOrFa ASC "
     '
     'Set OfaTemp = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
     Dim OfaTemp As ADODB.Recordset
     Set OfaTemp = New ADODB.Recordset
     OfaTemp.CursorType = adOpenDynamic
     OfaTemp.LockType = adLockReadOnly
     OfaTemp.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
     
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
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
     sqlx$ = " SELECT * FROM RESERVA "
     sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + "AND CantCons < (CantRes - 0.005) "
     sqlx$ = sqlx$ + "ORDER BY IdSubOr ASC "
     '
     'Set ReserTemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     
     Dim ReserTemp As ADODB.Recordset
     Set ReserTemp = New ADODB.Recordset
     ReserTemp.CursorType = adOpenDynamic
     ReserTemp.LockType = adLockReadOnly
     ReserTemp.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
     
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
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
             IDLOTXX$ = TRIM$(!IdenLote)
             OBSER$ = ""
             If IDLOTXX$ <> "" Then OBSER$ = "Lote " + IDLOTXX$
             FEX = CVINT(!FechRes)
             If Left$(NORFA$, 2) = "OR" Then
               REFE$ = "Reserva p/Reparación"
               GoTo 38
             End If
             '
             DESTIN$ = NORFA$
             CICON% = CODIFAB%(NORFA$)
             If CICON% > 0 Then
               DESTIN$ = TRIM$(CODEXT$(CICON%)) + " - " + TRIM$(DESCRIT0$(CICON%))
             End If
             '
             REFE$ = "Reserva p/" + DESTIN$
             '
38           LI2TEX$ = Space$(96)
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
     ReserTemp.Close
     Set ReserTemp = Nothing
     '
     GoSub PRESEDIS
     '
     STOTER = 0
     Call ABRESTOCKS
     sqlx$ = "SELECT * FROM MACONSIG "
     sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CI% & " "
     sqlx$ = sqlx$ + " ORDER BY Femi_Ocom "
     'Set LIMACONTMP = Stocks.OpenRecordset(sqlx$, 4)
     '
     Dim LIMACONTMP As ADODB.Recordset
     Set LIMACONTMP = New ADODB.Recordset
     LIMACONTMP.CursorType = adOpenDynamic
     LIMACONTMP.LockType = adLockReadOnly
     LIMACONTMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, , , adcmtext
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
     Label404.Caption = TRIM$(CSTRING$(MKS$(XVALO(Label103) - XVALO(Label202)), 4, 12, 1, 2))
     '
     DoEvents
     '
Return
'
COBRECEPTMP.Close
Set COBRECEPTMP = Nothing
PEDTEMP.Close
Set PEDTEMP = Nothing
OcoTemp.Close
Set OcoTemp = Nothing
AtenTemp.Close
Set AtenTemp = Nothing
OfaTemp.Close
Set OfaTemp = Nothing
ReserTemp.Close
Set ReserTemp = Nothing
LIMACONTMP.Close
Set LIMACONTMP = Nothing

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

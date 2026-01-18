VERSION 5.00
Begin VB.Form CODISTOK 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11850
   ClipControls    =   0   'False
   Icon            =   "CODISTOK.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5460
      Picture         =   "CODISTOK.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Visualizar Ordenes de Venta Pendientes de Entrega"
      Top             =   210
      Visible         =   0   'False
      Width           =   560
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Listado de Existencias y Ubicaciones"
      Top             =   5390
      Width           =   855
   End
   Begin VB.CommandButton Command5 
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1690
      Width           =   855
   End
   Begin VB.CommandButton COMMAND3 
      Caption         =   "Kardex"
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
      Left            =   10815
      Picture         =   "CODISTOK.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
      Top             =   950
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   435
      Left            =   4100
      TabIndex        =   52
      Top             =   210
      Width           =   1275
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   54
         Top             =   200
         Width           =   1240
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   53
         Top             =   0
         Value           =   -1  'True
         Width           =   1275
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Imprime Situación de Stocks"
      Top             =   4650
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   745
      Left            =   10815
      Picture         =   "CODISTOK.frx":13DC
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Acceso a Listas de Precios de Venta"
      Top             =   3170
      Width           =   855
   End
   Begin VB.CommandButton Command27 
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":1526
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6140
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
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
      Height          =   960
      Left            =   210
      TabIndex        =   40
      Top             =   3255
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
         Left            =   1050
         TabIndex        =   46
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St Neto:"
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
         TabIndex        =   45
         Top             =   525
         Width           =   855
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
         Left            =   3675
         TabIndex        =   44
         Top             =   420
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2310
         TabIndex        =   43
         Top             =   525
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5250
         TabIndex        =   42
         Top             =   525
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
         TabIndex        =   41
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8715
      ScaleHeight     =   75
      ScaleWidth      =   1440
      TabIndex        =   29
      Top             =   2415
      Width           =   1500
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   210
      TabIndex        =   27
      Top             =   4890
      Width           =   10410
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5460
      TabIndex        =   20
      Top             =   1470
      Width           =   2640
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
         Left            =   1260
         TabIndex        =   39
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
         Left            =   315
         TabIndex        =   38
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
         Left            =   1260
         TabIndex        =   24
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
         Left            =   315
         TabIndex        =   23
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
         Left            =   1260
         TabIndex        =   22
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
         Left            =   105
         TabIndex        =   21
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   15
      Top             =   1470
      Width           =   2430
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
         Left            =   1050
         TabIndex        =   37
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
         Left            =   105
         TabIndex        =   36
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reservas:"
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
         TabIndex        =   19
         Top             =   840
         Width           =   855
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
         Left            =   1050
         TabIndex        =   18
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
         Left            =   105
         TabIndex        =   17
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
         Left            =   1050
         TabIndex        =   16
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   210
      TabIndex        =   13
      Top             =   1470
      Width           =   2430
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
         TabIndex        =   35
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
         Left            =   105
         TabIndex        =   34
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
         Left            =   105
         TabIndex        =   33
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
         Left            =   1050
         TabIndex        =   32
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
         Left            =   105
         TabIndex        =   31
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
         Left            =   1050
         TabIndex        =   14
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   8295
      TabIndex        =   10
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
      Left            =   1155
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
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
      Height          =   600
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   735
      Width           =   5055
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
      Left            =   3570
      TabIndex        =   1
      Top             =   210
      Width           =   175
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualización de Ubicaciones de Stock"
      Top             =   2430
      Width           =   855
   End
   Begin VB.ListBox LIUBIC 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   8295
      TabIndex        =   6
      Top             =   2970
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":1AB2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Acceso a Disponibilidad General de Stocks"
      Top             =   3910
      Width           =   855
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
      Height          =   740
      Left            =   10815
      Picture         =   "CODISTOK.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
      Top             =   210
      Width           =   855
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
      TabIndex        =   57
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
      TabIndex        =   56
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10630
      X2              =   10630
      Y1              =   0
      Y2              =   7770
   End
   Begin VB.Label STOMINI 
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
      TabIndex        =   48
      Top             =   735
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   10575
      Picture         =   "CODISTOK.frx":1F06
      Stretch         =   -1  'True
      Top             =   7010
      Width           =   5250
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
      TabIndex        =   47
      Top             =   760
      Width           =   1275
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comprob.           Referencia                                                Entra          Sale         Fecha      Sucursal"
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
      TabIndex        =   28
      Top             =   4620
      Width           =   10410
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
      Left            =   8295
      TabIndex        =   26
      Top             =   2690
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
      Left            =   8295
      TabIndex        =   25
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
      Left            =   105
      TabIndex        =   12
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
      Left            =   6510
      TabIndex        =   11
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "CODISTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Text1.TEXT = ""
    IDTX$ = IDENTER$
    AVALSTOK$ = "LISTAPRE"
    If Option2.Value = True Then
       AVALSTOK$ = "MAESTRO"
    End If
    Call GRACONTROL(IDTX$, "AVALSTOK", AVALSTOK$)
    CONSTOK.Show
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
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
         OPDISTOK.Show 1
       Else
         Call MENSERR(24, "La Opción Elegida Requiere Privilegios de Supervisor")
    End If
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'Call CONECRUN("*KARDEX/" + COD$, "")
    Call GENKARDEX(COD$)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    DISPOSTO.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Command6.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
   Static CTXX%, MODOCORE%
   CODD$ = TRIM$(Text1.TEXT)
   CIXX% = CODINT%(CODD$)
   If CIXX% <= 0 Then
      Exit Sub
   End If
   '
   If MODOCORE% < 1 Then
      MODOCORE% = 1
      If TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE"))) = "PREBASE" Then
         MODOCORE% = 2
      End If
   End If
   '
   MUEPRECIO.List1.Clear
   ZZ$ = "C.Reposición" + Space$(30): COEFI = 0
   If MODOCORE% = 2 Then
      ZZ$ = "Precio Base" + Space$(30)
      COEFI = 15
   End If
   '
   COPRES = COSUNI(CIXX%) * (1 + COEFI / 100)
   Call REPLA(ZZ$, CSTRING$(MKS$(COPRES), 4, 10, 2, 2), 26, 10)
   MUEPRECIO.List1.AddItem ZZ$
   MUEPRECIO.List1.AddItem String$(80, 45)
   '
   For U& = 1 To ULTREG&("LISTAS")
      ZZ$ = Mid$(REGLEIDO$("LISTAS", U&), 2, 45)
      If TRIM$(ZZ$) <> "" Then
         CLIS$ = TRIM$(Str$(U&))
         While Len(CLIS$) < 3: CLIS$ = "0" + CLIS$: Wend
         APREC$ = "LIPRE" + CLIS$
         Call REPLA(ZZ$, CSTRING$(MKD$(PRELISTA#(APREC$, CIXX%)), 4, 10, 2, 2), 26, 10)
         AAA = PRELISTA#("LIPRE001", CIXX%)
         MUEPRECIO.List1.AddItem ZZ$
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
        DOCPARAM$(4) = TRIM$(STOMINI.Caption)
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
    Call LIUBIEXI
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If InStr(EPAC$, "DOSIVA") > 0 Then
       Command10.Visible = True
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
    Call VERNOVTRA
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
   Cancel = 1
   CONSTOK.Show
   WindowState = 1
End Sub

Private Sub Label13_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
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
    Label102.Caption = ""
    Label103.Caption = ""
    Label201.Caption = ""
    Label202.Caption = ""
    Label203.Caption = ""
    Label301.Caption = ""
    Label302.Caption = ""
    Label303.Caption = ""
    Label401.Caption = ""
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
    CONEXI.Text1.SetFocus
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
     EDITFORM.Command3.Enabled = False
     EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

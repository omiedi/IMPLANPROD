VERSION 5.00
Begin VB.Form SITUSTOK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ClipControls    =   0   'False
   Icon            =   "SITUSTOK.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.CommandButton Command11 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6195
      TabIndex        =   69
      Top             =   210
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   2680
      Left            =   8295
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   67
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
         TabIndex        =   68
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
      Left            =   8295
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   64
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
         TabIndex        =   66
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
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   210
      MultiSelect     =   1  'Simple
      TabIndex        =   22
      Top             =   4725
      Width           =   10410
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1105
      Left            =   10815
      TabIndex        =   60
      Top             =   5250
      Width           =   1000
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Realizar Movimientos de Stock"
         Top             =   265
         Width           =   750
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   10815
      TabIndex        =   58
      Top             =   6355
      Width           =   1000
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
         Height          =   480
         Left            =   105
         Picture         =   "SITUSTOK.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Imprime Situación de Stocks"
         Top             =   265
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Queries"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2560
      Left            =   10815
      TabIndex        =   54
      Top             =   2670
      Width           =   1000
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
         Left            =   105
         Picture         =   "SITUSTOK.frx":08FE
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Acceso a Consultas y Listados"
         Top             =   1735
         Width           =   750
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":0A48
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Consulta de Precios de Venta"
         Top             =   1000
         Width           =   750
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":0D52
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   265
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2570
      Left            =   10815
      TabIndex        =   50
      Top             =   105
      Width           =   1000
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":105C
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1735
         Width           =   750
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1000
         Width           =   750
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
         Height          =   740
         Left            =   105
         Picture         =   "SITUSTOK.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   265
         Width           =   750
      End
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
      Left            =   5460
      Picture         =   "SITUSTOK.frx":18F2
      Style           =   1  'Graphical
      TabIndex        =   49
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
      Left            =   4100
      TabIndex        =   44
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
         TabIndex        =   46
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
         TabIndex        =   45
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
      Left            =   210
      TabIndex        =   35
      Top             =   3150
      Width           =   7890
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
         Picture         =   "SITUSTOK.frx":1BFC
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Muestra / Oculta Situación de Cobertura"
         Top             =   630
         Width           =   560
      End
      Begin VB.Image Image1 
         Height          =   250
         Left            =   105
         Picture         =   "SITUSTOK.frx":1F06
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
         Left            =   1995
         TabIndex        =   63
         Top             =   380
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
         Left            =   2900
         TabIndex        =   62
         Top             =   670
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
         Left            =   1050
         TabIndex        =   41
         Top             =   670
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
         Left            =   105
         TabIndex        =   40
         Top             =   380
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
         Left            =   4700
         TabIndex        =   39
         Top             =   670
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
         Left            =   3885
         TabIndex        =   38
         Top             =   380
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
         Left            =   5670
         TabIndex        =   37
         Top             =   380
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
         TabIndex        =   36
         Top             =   670
         Width           =   1170
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10815
      ScaleHeight     =   75
      ScaleWidth      =   945
      TabIndex        =   24
      Top             =   7350
      Width           =   1000
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   12000
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
      Left            =   5460
      TabIndex        =   15
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         Left            =   315
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
         Left            =   1260
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
         Left            =   105
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
      Left            =   2730
      TabIndex        =   10
      Top             =   1470
      Width           =   2640
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
         Left            =   1260
         TabIndex        =   32
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
         Left            =   315
         TabIndex        =   31
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
         Left            =   105
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
         Left            =   1260
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
         Left            =   315
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
         Left            =   1260
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
      Left            =   210
      TabIndex        =   8
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
      Left            =   8295
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
      TabIndex        =   4
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
      Left            =   3570
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
      Height          =   960
      Left            =   8295
      TabIndex        =   2
      Top             =   2970
      Width           =   2325
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Histórico por Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8300
      TabIndex        =   70
      Top             =   4030
      Width           =   2325
   End
   Begin VB.Line Line2 
      X1              =   210
      X2              =   12000
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10635
      X2              =   10635
      Y1              =   0
      Y2              =   8400
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10290
      Picture         =   "SITUSTOK.frx":2210
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   7645
      Width           =   2010
   End
   Begin VB.Line Line4 
      X1              =   8295
      X2              =   10605
      Y1              =   2560
      Y2              =   2560
   End
   Begin VB.Line Line3 
      X1              =   8295
      X2              =   10605
      Y1              =   3930
      Y2              =   3930
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   23
      Top             =   4440
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
      TabIndex        =   21
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   2700
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
      Left            =   105
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
      Left            =   6510
      TabIndex        =   6
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "SITUSTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%
Dim MOCONPED%

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
   If InStr(EPAC$, "SABO") > 0 Then
       Call CONODESCLI
     ElseIf InStr(EPAC$, "DOSIVA") > 0 Then
       CONPEDID.Show 1
     Else
       OPX% = COMALTER%("Formato de Consulta de Pedidos:\\Ver y Re-Imprimir\Pendientes por Numero\Pendientes por Cliente\Pendientes por Material")
       If OPX% = 1 Then
           Call SHOWPEDIPEN
         ElseIf OPX% = 2 Then
           Call FINAL("*COPEDNO")
         ElseIf OPX% = 3 Then
            Call LIPEDIPEN
         Else
            Call LIPEDIMAT
       End If
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   FLIXX% = FECHANUM("01/05/03")
   FIN& = ULTREG&("MOVISTO")
   Screen.MousePointer = 11
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MOVISTO", U&)
     If CVI(Left$(XX$, 2)) >= FLIXX% Then
      If TRIM$(Mid$(XX$, 23, 2)) = "RT" Then
       If Val(Mid$(XX$, 26, 7)) >= 4946 Then
         FF% = CVI(Left$(XX$, 2))
         CASALD# = CVD(Mid$(XX$, 96, 8))
         CIXI% = CVI(Mid$(XX$, 3, 2))
         If CASALD# >= 0.05 Then
           For JJ& = U& + 1 To FIN&
             YY$ = REGLEIDO$("MOVISTO", JJ&)
             If Mid$(YY$, 23, 3) = "PL." Then
              If CVI(Left$(YY$, 2)) = FF% Then
               If Abs(CVD(Mid$(YY$, 88, 8)) - CASALD#) < 0.05 Then
                If CVI(Mid$(YY$, 3, 2)) = CIXI% Then
                 GoTo 19
                End If
               End If
              End If
             End If
           Next JJ&
MsgBox CODEXT$(CIXI%) & "   " & Mid$(XX$, 23, 10)
GoSub 90
         End If
       End If
      End If
     End If
19 Next U&
Screen.MousePointer = 1
               
Exit Sub
   '
90 YY$ = XX$
   CANTIDXX# = CVD(Mid$(XX$, 96, 8))
   REFEX$ = "Packing-List " + Mid$(XX$, 43, 30)
   Call REPLA(YY$, "RP", 21, 2)
   Call REPLA(YY$, "PL.", 23, 10)
   Call REPLA(YY$, REFEX$, 43, 30)
   Call REPLA(YY$, MKD$(CANTIDXX#), 88, 8)
   Call REPLA(YY$, MKD$(0), 96, 8)
   Call REGAPP("MOVISTO", YY$)
   Call CIERRARCH("MOVISTO")
   Return
   
         
End Sub


Private Sub Command12_Click()
   Command12.Enabled = False
   CIXI% = CODINT%(Text1)
   If CIXI% < 1 Then GoTo 99
   '
   Dim TLOTE$(), CALOTE#()
   '
   ' SELECCIONA LOTES DE INGRESO
   CANLOT% = 0
   RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
   For U& = 1 To Len(RFF$) Step 4
     REMO& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("MOVISTO", REMO&)
     ENTRA# = CVD(Mid$(X$, 88, 8))
     '
     If ENTRA# >= 0.05 Then
       IDENLO$ = TRIM$(Mid$(X$, 5, 16))
       If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 23, 10))
       If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 33, 10))
       IDENLO$ = AJUSTI$(IDENLO$, 16)
       '
       For KKI% = 1 To CANLOT%
         If Left$(TLOTE$(KKI%), 16) = IDENLO$ Then
           CALOTE#(KKI%) = CALOTE#(KKI%) + ENTRA#
           TTXX$ = TLOTE$(KKI%)
           Call REPLA(TTXX$, CSTRING$(MKD$(CALOTE#(KKI%)), 4, 12, 1, 2), 27, 12)
           TLOTE$(KKI%) = TTXX$
           GoTo 19
         End If
       Next KKI%
       FEX = CVI(Left$(X$, 2))
       CANLOT% = CANLOT% + 1
       ReDim Preserve TLOTE$(CANLOT%), CALOTE#(CANLOT%)
       TTXX$ = AJUSTI$(IDENLO$, 48)
       Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
       Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
       Call REPLA(TTXX$, CSTRING$(MKS$(REMO&), 3, 8, 0, 2), 41, 8)
       TLOTE$(CANLOT%) = TTXX$
       CALOTE#(KKI%) = ENTRA#
     End If
     '
19 Next U&
   '
   If CANLOT% < 1 Then
     Call COMUNI("Sin Movimientos de Ingreso")
     GoTo 99
   End If
   '
   ' ORDENA POR FECHA DE INGRESO
   For KI% = 1 To CANLOT%
     For KJ% = KI% + 1 To CANLOT%
       If FECHANUM(Mid$(TLOTE$(KI%), 19, 8)) > FECHANUM(Mid$(TLOTE$(KJ%), 19, 8)) Then
         RESER$ = TLOTE$(KI%): RESCA# = CALOTE#(KI%)
         TLOTE$(KI%) = TLOTE$(KJ%): CALOTE#(KI%) = CALOTE#(KJ%)
         TLOTE$(KJ%) = RESER$: CALOTE#(KJ%) = RESCA#
       End If
     Next KJ%
   Next KI%
   '
   Dim RESAL&(), CASAL#(), FESAL%(), IDSAL$()
   CAMOSA% = 0
   For U& = 1 To Len(RFF$) Step 4
     REMO& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("MOVISTO", REMO&)
     SALID# = CVD(Mid$(X$, 96, 8))
     If SALID# >= 0.05 Then
       IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
       For KKI% = 1 To CANLOT%
         If Left$(TLOTE$(KKI%), 16) = IDENLO$ Then GoTo 29
       Next KKI%
       IDENLO$ = Space$(16)
       Call REPLA(X$, Space$(16), 5, 16)
       Call GRAREG("MOVISTO", X$, REMO&)
       '
       CAMOSA% = CAMOSA% + 1
       ReDim Preserve RESAL&(CAMOSA%), CASAL#(CAMOSA%), FESAL%(CAMOSA%), IDSAL$(CAMOSA%)
       RESAL&(CAMOSA%) = REMO&
       CASAL#(CAMOSA%) = SALID#
       FESAL%(CAMOSA%) = CVI(Left$(X$, 2))
       IDSAL$(CAMOSA%) = IDENLO$
     End If
29 Next U&
   '
35 REBLA$ = REGBLAN$("DETALOT")
   JJ& = 0: II& = 0
   For KKI% = 1 To CANLOT%
     LOTEN$ = Left$(TLOTE$(KKI%), 16)
     CAPENLO# = CALOTE#(KKI%)
     FELOT% = FECHANUM(Mid$(TLOTE$(KKI%), 19, 8))
     RELOT& = Val(Mid$(TLOTE$(KKI%), 41, 8))
     '
     II& = II& + 1
     Call GRAREG("INDILOT", Left$(TLOTE$(KKI%), 40), II&)
     '
     XX$ = REGLEIDO$("MOVISTO", RELOT&)
     NUCO1$ = TRIM$(Mid$(XX$, 23, 10))
     NUCO2$ = TRIM$(Mid$(XX$, 33, 10))
     REFEX$ = TRIM$(Mid$(XX$, 43, 30))
     TIMOX$ = TRIM$(UCase$(Mid$(XX$, 21, 2)))
     '
     YY$ = REBLA$
     Call REPLA(YY$, LOTEN$, 1, 16)
     Call REPLA(YY$, MKI$(FELOT%), 17, 2)
     Call REPLA(YY$, NUCO1$, 19, 12)
     Call REPLA(YY$, NUCO2$, 31, 12)
     Call REPLA(YY$, REFEX$, 43, 30)
     Call REPLA(YY$, MKD$(CAPENLO#), 73, 8)
     JJ& = JJ& + 1
     Call GRAREG("DETALOT", YY$, JJ&)
     '
     If TIMOX$ = "AJ" Or TIMOX$ = "TR" Or TIMOX$ = "SJ" Then FELOT% = 0
     ' LO ANTERIOR ES PARA QUE AJUSTES Y TRANSFERENCIAS TAMBIEN NETEEN HACIA ATRAS
     '
     For KKJ% = 1 To CAMOSA%
       If CAPENLO# >= 0.05 Then
         If FESAL%(KKJ%) >= FELOT% Then
           If IDSAL$(KKJ%) = LOTEN$ Or TRIM$(IDSAL$(KKJ%)) = "" Then
             If CASAL#(KKJ%) >= 0.05 Then
               CASALI# = CASAL#(KKJ%)
               If CASALI# > CAPENLO# Then CASALI# = CAPENLO#
               '
               ZZ$ = REGLEIDO$("MOVISTO", RESAL&(KKJ%))
               FEXI% = CVI(Left$(ZZ$, 2))
               NUCO1$ = TRIM$(Mid$(ZZ$, 23, 10))
               NUCO2$ = TRIM$(Mid$(ZZ$, 33, 10))
               REFEX$ = TRIM$(Mid$(ZZ$, 43, 30))
               YY$ = REBLA$
               Call REPLA(YY$, LOTEN$, 1, 16)
               Call REPLA(YY$, MKI$(FEXI%), 17, 2)
               Call REPLA(YY$, NUCO1$, 19, 12)
               Call REPLA(YY$, NUCO2$, 31, 12)
               Call REPLA(YY$, REFEX$, 43, 30)
               Call REPLA(YY$, MKD$(CASALI#), 81, 8)
               '
               JJ& = JJ& + 1
               Call GRAREG("DETALOT", YY$, JJ&)
               CAPENLO# = CAPENLO# - CASALI#
               CASAL#(KKJ%) = CASAL#(KKJ%) - CASALI#
             End If
           End If
         End If
       End If
     Next KKJ%
   Next KKI%
   Call SETULTREG("INDILOT", II&)
   Call FRESHECHO("INDILOT")
   On Error Resume Next
   Kill LUDAT$ + "INDILOT.SEL"
   On Error GoTo 0
   Call SETULTREG("DETALOT", JJ&)
   '
   MODOX% = COMALTER%("Opciones de Acceso:\\Consulta por Lote\Listado")
   If MODOX% = 1 Then
       Call CONECRUN("*DETALOT", "Detalle Consumos por Lote" + "/WAIT")
     Else
       '\\\OT-05-0552-WAR-01/08/05///
       Call HEADERS("DETALOT", "")
       CI% = CIXI%
       CODIDESCRI$ = TRIM(CODEXT(CI%)) + " - " + DESCRIT0(CI%)
       Call HEADERS("DETALOT", "Código: " + CODIDESCRI$)
       '\\\OT-05-0552-WAR-FIN///
       Call CONECRUN("*LIDELOT", "Detalle Consumos por Lote")
   End If
   '
99 Command12.Enabled = True
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
    COMMAND3.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX/" + Text1, "Ficha Kardex de Movimientos")
      Else
        Call GENKARDEX(Text1)
    End If
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
   MOVIMSTO.Show 1
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
    QUERISTO.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_Activate()
   Call GRATITFOR(Caption)
   WindowState = 0
   '
   CDDX$ = Left$(TRIM$(Clipboard.GetText), 16)
   If CDDX$ <> "" Then
     If CODINT%(CDDX$) > 0 Then
       Text1 = CDDX$
       Clipboard.SetText ""
     End If
   End If
   '
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
    If EXISTE%("PEDIDO02.EXE") > 0 Or INTERPRE% = 1 Then
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
    Call VERNOVTRA
    '
    'MOCONPED% = 2
    'If EXISTE%("PEDIDO01.EXE") > 0 Or EXISTE%("PEDIDO02.EXE") > 0 Then
       MOCONPED% = 1
    'End If
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
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * Val(Label103))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * Val(Label103))
       Picture2.ForeColor = RGB(200, 150, 0)
       Picture2.Line (0.05 * Picture2.Width, 0.94 * Picture2.Height - COEFV * Val(STOMI))-(0.95 * Picture2.Width, 0.94 * Picture2.Height - COEFV * Val(STOMI))
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
   Dim TLOTE$(256), CALOTE#(256)
   CIXI% = CODINT%(Text1)
   If CIXI% > 0 Then
     CANLOT% = 0
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       ENTRA# = CVD(Mid$(X$, 88, 8))
       If ENTRA# >= 0.05 Then
         IDENLO$ = TRIM$(Mid$(X$, 5, 16))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 23, 10))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 33, 10))
         IDENLO$ = AJUSTI$(IDENLO$, 16)
         '
         For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Then
             CALOTE#(KKI%) = CALOTE#(KKI%) + ENTRA#
             TTXX$ = TLOTE$(KKI%)
             Call REPLA(TTXX$, CSTRING$(MKD$(CALOTE#(KKI%)), 4, 12, 1, 2), 27, 12)
             TLOTE$(KKI%) = TTXX$
             GoTo 19
           End If
         Next KKI%
         FEX = CVI(Left$(X$, 2))
         CANLOT% = CANLOT% + 1
         TTXX$ = AJUSTI$(IDENLO$, 40)
         Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
         Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
         TLOTE$(CANLOT%) = TTXX$
         CALOTE#(KKI%) = ENTRA#
         '
       End If
19   Next U&
     '
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       SALID# = CVD(Mid$(X$, 96, 8))
       If SALID# >= 0.05 Then
         IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
         IDENLOX$ = TRIM$(IDENLO$)
         '
35       For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Or IDENLOX$ = "" Then
             If SALID# <= CALOTE#(KKI%) Then
                 CALOTE#(KKI%) = CALOTE#(KKI%) - SALID#
                 SALID# = 0
               Else
                 SALID# = SALID# - CALOTE#(KKI%)
                 CALOTE#(KKI%) = 0
             End If
           End If
         Next KKI%
         If SALID# >= 0.05 Then
           IDENLOX$ = ""
           GoTo 35
         End If
         '
       End If
     Next U&
     '
     REBLA$ = REGBLAN$("STOLOTE"): JJ& = 0
     For KKI% = 1 To CANLOT%
       If CALOTE#(KKI%) > 0.05 Then
         ZZ$ = REBLA$
         Call REPLA(ZZ$, TLOTE$(KKI%), 1, 38)
         Call REPLA(ZZ$, MKD$(CALOTE#(KKI%)), 41, 8)
         JJ& = JJ& + 1
         Call GRAREG("!STOLOTE", ZZ$, JJ&)
       End If
     Next KKI%
     Call SETULTREG("!STOLOTE", JJ&)
     Call CIERRARCH("!STOLOTE")
     '
     Call FRESHECHO("!STOLOTE")
     Call SELECHO("!STOLOTE/Id.Lote                       F.Alta                Q.Orig                      Saldo")
   End If
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

Private Sub Option1_Click()
   If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
      Option2.Value = True
   End If
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
     EDITFORM.COMMAND3.Enabled = False
     EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$, SIMAT%
     '
     Screen.MousePointer = 11
     If CTX% <> 1 Then
        IDENSEDE$ = CONTROL$("", "IDENSEDE")
        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        '
        Call GENCONPEN(13)   ' PARA GENERAR ARCHIVO DE MATERIAL EN CONSIGNACIÓN
        '
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
     Call VERISAL(CI%)
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     REG& = CI% + 1
     If REG& > 0 Then
        If REG& <= NUMAS% + 1 Then
           SALCA$ = REGLEIDO$("SALSTOCK", REG&) + Mid$(REGLEIDO$("SALDEPOS", REG&), 41)
           For U& = 1 To NUSU%
               XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
               XXX$ = ADEPO$(U&) + CSTRING$(XXZ$, 4, 8, 1, 2)
               If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
               If Abs(CVS(XXZ$)) >= 0.05 Then
                  List1.AddItem XXX$
                  CAPAR = CVS(XXZ$)
                  If DISTO%(U&) >= 0 Then
                     STOPROP = STOPROP + CVS(XXZ$)
                  End If
               End If
           Next U&
           GoSub PRESEDIS
       End If
     End If
     '
10   List2.Clear
     HOII% = HOY(HO$)
     LI2TEX$ = Space$(80)
     Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 30)
     Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, 1, 2), 44, 9)
     Call REPLA(LI2TEX$, HO$, 65, 9)
     Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
     List2.AddItem LI2TEX$
     SALSTO = STOTOT
     '
12   TORECEP = 0
     RFF$ = RECFILT$("BOLREPEN", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("BOLREPEN", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("BOLREPEN", 1)
           Call GRAREG("BOLREPEN", X$, 1)
           Call CIERRARCH("BOLREPEN")
           GoTo 12
        End If
        '
        CARECEP = CVD(Mid$(X$, 89, 8))
        If CARECEP > 0 Then
           NCI% = CVI(Mid$(X$, 105, 2))
           REMIPRO$ = TRIM$(Mid$(X$, 119, 16))
           NROBOL& = CVS(Left$(X$, 4))
           FEX = CVI(Mid$(X$, 81, 2))
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2)), 1, 9)
           Call REPLA(LI2TEX$, "RT." + REMIPRO$ + " - " + RASOCLI$((-1) * NCI%), 15, 30)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CARECEP), 4, 9, 1, 2), 44, 9)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
           List2.AddItem LI2TEX$
           TOSUMIN = TOSUMIN + CARECEP
           TORECEP = TORECEP + CARECEP
        End If
     Next UK&
     '
     If TORECEP > 0.05 Then
        XXX$ = "RECEPC. " + CSTRING$(MKS$(TORECEP), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     '
     GoSub PRESEDIS
     '
     EPAC$ = EMPREAC$
     '
     If MOCONPED% <> 2 Then
       '
       UREOD& = ULTREG&("MAVENDID")
       RGXX$ = RECFILT$("MAVENDID", 20, MKI$(CI%))
       For U& = 1 To Len(RGXX$) Step 4
         REG& = CVS(Mid$(RGXX$, U&, 4))
         If REG& > 0 Then
           If REG& <= UREOD& Then
             X$ = REGLEIDO$("MAVENDID", REG&)
             If Asc(Mid$(X$, 63, 1)) <= 1 Then  ' STATUS 1
               NRODES& = CVS(Mid$(X$, 87, 4))
               NROPED& = CVS(Mid$(X$, 93, 4))
               NCI% = CVI(Mid$(X$, 1, 2))
               CAPEN = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
               FEX = CVI(Mid$(X$, 125, 2))
               If CAPEN < 0 Then CAPEN = 0
               '
               If CAPEN > 0 Then
                 TORESER = TORESER + CAPEN
                 LI2TEX$ = Space$(80)
                 Call REPLA(LI2TEX$, "OD." + TRIM$(CSTRING$(MKS$(NRODES&), 3, 6, 0, 2)), 1, 9)
                 Call REPLA(LI2TEX$, "Ped." + TRIM$(Str$(NROPED&)) + " - " + RASOCLI$(NCI%), 15, 30)
                 Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
                 Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
                 Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
                 List2.AddItem LI2TEX$
               End If
             End If
           End If
         End If
       Next U&
       '
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") > 0 Then
         UREPE& = ULTREG&("DESPAPED")
         RGXX$ = RECFILT$("DESPAPED", 12, MKI$(CI%))
         For U& = 1 To Len(RGXX$) Step 4
           REG& = CVS(Mid$(RGXX$, U&, 4))
           If REG& > 0 Then
             If REG& <= UREPE& Then
               X$ = REGLEIDO$("DESPAPED", REG&)
               '
               If Asc(Mid$(X$, 27, 1)) < 2 Or Asc(Mid$(X$, 27, 1)) = 11 Then
                 If Mid$(X$, 51, 1) <> "R" Then ' no es reparacion
                   NROPED& = CVS(Left$(X$, 4))
                   NCI% = CVI(Mid$(X$, 7, 2))
                   CAPEN = CVS(Mid$(X$, 75, 4)) - CVS(Mid$(X$, 95, 4))
                   FEX = CVI(Mid$(X$, 25, 2))
                   If CAPEN > 0 Then
                     TOPEDID = TOPEDID + CAPEN
                     LI2TEX$ = Space$(80)
                     Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
                     Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 30)
                     Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
                     Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
                     'Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
                     '
                     CATOPRE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
 
                     If CATOPRE >= 0.05 Then
                       QPRE$ = "Pr=" + TRIM$(FORMATNUM$(CATOPRE, "F10.1"))
                       Call REPLA(LI2TEX$, QPRE$, 75, 8)
                     End If
                     '
                     List2.AddItem LI2TEX$
                   End If
                 End If
               End If
             End If
           End If
         Next U&
         '
         GoSub PRESEDIS
         GoTo 15
         '
       End If
       '
       ABUSQUE$ = "PEDDETP"
14     UREPE& = ULTREG&(ABUSQUE$)
       RGXX$ = RECFILT$(ABUSQUE$, 12, MKI$(CI%))
       For U& = 1 To Len(RGXX$) Step 4
         REG& = CVS(Mid$(RGXX$, U&, 4))
         If REG& > 0 Then
           If REG& <= UREPE& Then
             X$ = REGLEIDO$(ABUSQUE$, REG&)
             '
             If ABUSQUE$ = "PEDDETP" Then
               REGPEDA& = CVS(Mid$(X$, 109, 4))
               If Left$(REGLEIDO$("PEDDETA", REGPEDA&), 108) <> Left$(X$, 108) Then
                 ABUSQUE$ = "PEDDETA"
                 GoTo 14
               End If
             End If
             '
             If Asc(Mid$(X$, 27, 1)) <= 2 Then
               If Mid$(X$, 51, 1) <> "R" Then ' no es reparacion
                 NROPED& = CVS(Left$(X$, 4))
                 NCI% = CVI(Mid$(X$, 7, 2))
                 CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                 CAPEN = CAPEN - CVS(Mid$(X$, 75, 4)) 'descuenta asignados
                 FEX = CVI(Mid$(X$, 25, 2))
                 If CAPEN < 0 Then CAPEN = 0
                 '
                 If CAPEN > 0 Then
                   TOPEDID = TOPEDID + CAPEN
                   LI2TEX$ = Space$(80)
                   Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
                   Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 30)
                   Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
                   Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
                   Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
                   List2.AddItem LI2TEX$
                 End If
               End If
             End If
           End If
         End If
       Next U&
       '
       GoSub PRESEDIS
       GoTo 15
       '
     End If
     '
     For U& = 1 To Len(PUNPEDPE$(CI%)) Step 4
        REG& = CVS(Mid$(PUNPEDPE$(CI%), U&, 4))
        RGXX$ = RECFILT$("PEDDETP", 12, MKI$(CI%))
        X$ = REGLEIDO$(APEDPE$, REG&)
        CAPEN = 0
        If MODOMAI% <> 1 Then
            If Asc(Mid$(X$, 27, 1)) <= 1 Then
              NROPED& = CVS(Left$(X$, 4))
              NCI% = CVI(Mid$(X$, 7, 2))
              CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
              FEX = CVI(Mid$(X$, 25, 2))
            End If
          Else
            If Asc(Mid$(X$, 300, 1)) <= 1 Then
               NROPED& = CVS(Left$(X$, 4))
               NCI% = CVI(Mid$(X$, 7, 2))
               CACOPED = CVS(Mid$(X$, 513, 4))
               CAITPED = CVS(Mid$(X$, 645, 4))
               CACOENT = CVS(Mid$(X$, 512 + 77, 4))
               CAPEN = (CACOPED - CACOENT) * CAITPED
               FEX = CVI(Mid$(X$, 512 + 63, 2))
            End If
        End If
        '
        If CAPEN > 0 Then
           TOPEDID = TOPEDID + CAPEN
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
           Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 30)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
           List2.AddItem LI2TEX$
        End If
        '
     Next U&
     GoSub PRESEDIS
     '
15   RFF$ = RECFILT$("OCOMPRA", 6, MKI$(CI%))
     ' esta bien - despues tambien lee de OCOMDETA - 20/11/01
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMPRA", U&)
        CIIX% = CVI(Mid$(X$, 11, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("OCOMPRA", 1)
           Call GRAREG("OCOMPRA", X$, 1)
           Call CIERRARCH("OCOMPRA")
           GoTo 15
        End If
        STAT% = Asc(Mid$(X$, 62, 1))
        If STAT% <= 0 Then
           CACOM = CVS(Mid$(X$, 13, 4))
           CAREC = CVS(Mid$(X$, 58, 4))
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
              NCI% = CVI(Mid$(X$, 7, 2))
              NROPED& = CVS(Left$(X$, 4))
              FEX = CVI(Mid$(X$, 9, 2))
              LI2TEX$ = Space$(80)
              Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 30)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
              Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
              List2.AddItem LI2TEX$
              TOSUMIN = TOSUMIN + CAPEN
           End If
        End If
     Next UK&
     '
18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        CIIX% = CVI(Mid$(X$, 9, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("OCOMDETA", 1)
           Call GRAREG("OCOMDETA", X$, 1)
           Call CIERRARCH("OCOMDETA")
           GoTo 18
        End If
        STAT% = Asc(Mid$(X$, 124, 1))
        If STAT% < 8 Then      ' no esta anulada NI dada por cumplida
          NUORSE& = CVS(Mid$(X$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(X$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              NCI% = CVI(Mid$(X$, 7, 2))
              NROPED& = CVS(Left$(X$, 4))
              FEX = CVI(Mid$(X$, 95, 2))
              LI2TEX$ = Space$(80)
              Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 30)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
              Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
              List2.AddItem LI2TEX$
              TOSUMIN = TOSUMIN + CAPEN
            End If
          End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
     If SIMAT% = 0 Then
       SIMAT% = (-1)
       If EXISTE%(LUDAT$ + "BDSIMAT.RFS") > 0 Then
         If EXISTE%(LUDAT$ + "BDSIMAT.DAT") > 0 Then
           SIMAT% = 1
         End If
       End If
     End If
     '
     If SIMAT% > 0 Then
       RFF$ = RECFILT$("BDSIMAT", 10, MKI$(CI%))
       For UK& = 1 To Len(RFF$) Step 4
          U& = CVS(Mid$(RFF$, UK&, 4))
          X$ = REGLEIDO$("BDSIMAT", U&)
          CIIX% = CVI(Mid$(X$, 65, 2))
          CAPEN = CVS(Mid$(X$, 71, 4))
          If CVS(Mid$(X$, 131, 4)) < 1 Then ' no hay orden de compra
              COBRA$ = TRIM$(Mid$(X$, 7, 8))
            NROPED& = CVS(Left$(X$, 4))
            FEX = CVI(Mid$(X$, 75, 2))
            LI2TEX$ = Space$(80)
            Call REPLA(LI2TEX$, "SIM." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
            Call REPLA(LI2TEX$, ECOARCH$("INDIOBR", COBRA$), 15, 30)
            Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
            Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
            Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
            List2.AddItem LI2TEX$
            TOSUMIN = TOSUMIN + CAPEN
          End If
       Next UK&
       '
       GoSub PRESEDIS
     End If
     '
25   RFF$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("ORDEFABR", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("ORDEFABR", 1)
           Call GRAREG("ORDEFABR", X$, 1)
           Call CIERRARCH("ORDEFABR")
           GoTo 25
        End If
        STAT% = Asc(Mid$(X$, 108, 1))
        If STAT% <= 2 Then
           CACOM = CVD(Mid$(X$, 89, 8))
           CAREC = CVD(Mid$(X$, 119, 8))
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
              NROPED& = CVS(Left$(X$, 4))
              FEX = CVI(Mid$(X$, 105, 2))
              REFE$ = Mid$(X$, 5, 64)
              LI2TEX$ = Space$(80)
              Call REPLA(LI2TEX$, "OF." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, REFE$, 15, 30)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
              Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
              List2.AddItem LI2TEX$
              TOFABRI = TOFABRI + CAPEN
           End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
35   RFF$ = RECFILT$("RESERVA", 1, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("RESERVA", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("RESERVA", 1)
           Call GRAREG("RESERVA", X$, 1)
           Call CIERRARCH("RESERVA")
           GoTo 35
        End If
        TORE# = CVD(Mid$(X$, 19, 8))
        CACO# = CVD(Mid$(X$, 27, 8))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
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
           NORFA$ = Mid$(X$, 3, 12)
           FEX = CVI(Mid$(X$, 17, 2))
           If Left$(NORFA$, 2) = "OR" Then
             REFE$ = "Reserva p/Reparación"
             GoTo 38
           End If
           '
           CICON% = CVI(Mid$(X$, 79, 2))
           If CICON% < 1 Or CICON% > NUMAS% Then
             REBUL& = REGBUS&("ORDEFABR", 3, NORFA$)
             If REBUL& > 0 Then
               XY$ = REGLEIDO$("ORDEFABR", REBUL&)
               CICON% = CVI(Mid$(XY$, 83, 2))
               Call REPLA(X$, MKI$(CICON%), 79, 2)
               Call GRAREG("RESERVA", X$, U&)
             End If
           End If
           REFE$ = "Reserva p/" + CODEXT$(CICON%)   ' NORFA$
           '
38         LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
           Call REPLA(LI2TEX$, REFE$, 15, 30)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
           List2.AddItem LI2TEX$
           TORESER = TORESER + CAPEN
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
135  RFF$ = RECFILT$("REPUREPA", 3, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("REPUREPA", U&)
        TORE# = CVS(Mid$(X$, 21, 4))
        CACO# = CVS(Mid$(X$, 25, 4))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
           TORE# = TORE# / 1000
           CACO# = CACO# / 1000
        End If
        '
        SARE = TORE# - CACO#
        If SARE < 0 Then SARE = 0
        CAPEN = SARE
        If Abs(CAPEN) > 0 Then
           NORFA$ = Mid$(X$, 33, 12)
           REFE$ = "Reserva p/Reparación"
           NOREPARA& = CVS(Left$(X$, 4))
           REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
           If REREP& > 0 Then
             YY$ = REGLEIDO$("PRESREPA", REREP&)
             If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 139 ' O/Rep Cerrada
             If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 139 ' falta O/Venta
             REFE$ = "Reserva p/Rep. " + CODEXT$(CVI(Mid$(YY$, 69, 2)))
           End If
           If TRIM$(NORFA$) = "" Then
             NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
           End If
           FEX = HOY(HOS$)
           '
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
           Call REPLA(LI2TEX$, REFE$, 15, 30)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
           List2.AddItem LI2TEX$
           TORESER = TORESER + CAPEN
        End If
139  Next UK&
     '
     GoSub PRESEDIS
     '
40   RFF$ = RECFILT$("MACONSIG", 8, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("MACONSIG", U&)
        CIIX% = CVI(Mid$(X$, 33, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("MACONSIG", 1)
           Call GRAREG("MACONSIG", X$, 1)
           Call CIERRARCH("MACONSIG")
           GoTo 40
        End If
        TORE# = CVD(Mid$(X$, 55, 8))
        CACO# = CVD(Mid$(X$, 63, 8))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
           TORE# = TORE# / 1000
           CACO# = CACO# / 1000
        End If
        '
        SARE = TORE# - CACO#
        If SARE < 0 Then
          SARE = 0
        End If
        CAPEN = SARE
        If Abs(CAPEN) > 0 Then
           NORFA$ = TRIM$(Mid$(X$, 21, 12))
           NUOCO& = Val(Mid$(NORFA$, 4, 6))
           For TYU& = 1 To Len(VECOPEN$) Step 4
             If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 42
           Next TYU&
           GoTo 44  ' FILTRA PARA O/C's PENDIENTES
           '
42         FEX = CVI(Mid$(X$, 1, 2))
           REFE$ = "Consig." + RASOCLI$((-1) * CVI(Mid$(X$, 3, 2)))
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
           Call REPLA(LI2TEX$, REFE$, 15, 30)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
           List2.AddItem LI2TEX$
           TORESER = TORESER + CAPEN
        End If
44   Next UK&
     '
     GoSub PRESEDIS
     '
45   RFF$ = RECFILT$("STOREMOT", 1, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("STOREMOT", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("STOREMOT", 1)
           Call GRAREG("STOREMOT", X$, 1)
           Call CIERRARCH("STOREMOT")
           GoTo 45
        End If
        NORFA$ = Mid$(X$, 3, 12)
        TIMO$ = UCase$(Mid$(X$, 15, 2))
        REFE$ = Mid$(X$, 17, 50)
        TORE# = CVD(Mid$(X$, 67, 8))
        CACO# = CVD(Mid$(X$, 75, 8))
        FEX = CVI(Mid$(X$, 83, 2))
        SDOR$ = Mid$(X$, 85, 8)
        CAPEN = TORE# - CACO#
        '
        If Abs(CAPEN) > 0 Then
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
           Call REPLA(LI2TEX$, REFE$, 15, 30)
           If CAPEN > 0 Then
                Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
              Else
                Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPEN)), 4, 9, 1, 2), 53, 9)
           End If
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, SDOR$, 75, 8)
           List2.AddItem LI2TEX$
           '
           If TIMO$ = "ST" Then
               STOREMO = STOREMO + CAPEN
               XXX$ = SDOR$ + CSTRING$(MKS$(CAPEN), 4, 8, 1, 2)
               List1.AddItem XXX$
           End If
           If TIMO$ = "PC" Or TIMO$ = CMPED$ Then TOPEDID = TOPEDID - CAPEN
           If TIMO$ = "OF" Then TOFABRI = TOFABRI + CAPEN
           If TIMO$ = "RF" Then TORESER = TORESER - CAPEN
           If TIMO$ = "OC" Then TOSUMIN = TOSUMIN + CAPEN
           If TIMO$ = "BR" Then TOSUMIN = TOSUMIN + CAPEN
        End If
     Next UK&
     '
46   STOTER = 0
     RFF$ = RECFILT$("LCONPEN", 8, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("LCONPEN", U&)
        NORFA$ = Mid$(X$, 21, 12)
        NPRO% = CVI(Mid$(X$, 3, 2))
        REFE$ = RASOCLI$((-1) * NPRO%)
        TORE# = CVD(Mid$(X$, 63, 8))
        CACO# = CVD(Mid$(X$, 111, 8))
        FEX = CVI(Mid$(X$, 71, 2))
        SDOR$ = ""
        CAPEN = TORE# - CACO#
        '
        If Abs(CAPEN) > 0 Then
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
           Call REPLA(LI2TEX$, REFE$, 15, 30)
           If CAPEN > 0 Then
                Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
              Else
                Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPEN)), 4, 9, 1, 2), 53, 9)
           End If
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
           Call REPLA(LI2TEX$, SDOR$, 75, 8)
           List2.AddItem LI2TEX$
           '
           STOREMO = STOREMO + CAPEN
           STOTER = STOTER + CAPEN
           '
           For KKUU& = 1 To List2.ListCount
             If TRIM$(Left$(List2.List(KKUU& - 1), 12)) = TRIM$(NORFA$) Then
               LI2TEX$ = List2.List(KKUU& - 1)
               CAPETO = CAPEN + Val(Mid$(LI2TEX$, 53, 9))
               Call REPLA(LI2TEX$, FORMATNUM$(CAPETO, "F9.1"), 53, 9)
               List2.RemoveItem KKUU& - 1
               List2.AddItem LI2TEX$, KKUU& - 1
               TORESER = TORESER + CAPEN
               GoTo 47
             End If
           Next KKUU&
        End If
47   Next UK&
     '
     If Abs(STOTER) >= 0.005 Then
        XXX$ = "TERCEROS" + CSTRING$(MKS$(STOTER), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     '
     GoSub PRESEDIS
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
     DUPLI% = 0
     For UI& = 1 To List2.ListCount - 1
       RELI$ = List2.List(UI& - 1)
       For UJ& = UI& + 1 To List2.ListCount
         RELJ$ = List2.List(UJ& - 1)
         If RELI$ = RELJ$ Then DUPLI% = DUPLI% + 1
       Next UJ&
     Next UI&
     If DUPLI% > 0 Then
       Call COMUNI("ATENCION: Movimientos Aparentemente Duplicados !!!")
     End If
     '
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
     Label404.Caption = TRIM$(CSTRING$(MKS$(Val(Label103) - Val(Label202)), 4, 12, 1, 2))
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
   STOK = Val(Label103): If STOK < 0 Then STOK = 0
   COMP = Val(Label301): If COMP < 0 Then COMP = 0
   FABR = Val(Label302): If FABR < 0 Then FABR = 0
   '
   MINI = Val(STOMI): If MINI < 0 Then MINI = 0
   RESE = Val(Label203): If RESE < 0 Then RESE = 0
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

Sub LIPEDIPEN()
    '
    Screen.MousePointer = 11
    JJ& = 0: LL& = 0
    FIN& = ULTREG&("PEDDETA")
    '
    modo% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
    If modo% < 1 Then GoTo 99
    FLI1% = -1: FLI2% = FINMESACT%
    If modo% = 2 Then
        FLI1% = FINMESACT% + 1: FLI2% = 32767
10             FNMA = FINMESACT%
        FLI1% = DIPOST(FNMA): FLI2% = 32767
        VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
        OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
        If OBX$ = "" Then GoTo 99
        '
        FLI1% = CVI(Left$(OBX$, 2))
        FLI2% = CVI(Mid$(OBX$, 3, 2))
        If FLI1% > FLI2% Then GoTo 10
        FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
        FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
      ElseIf modo% = 3 Or modo% = 4 Then
        FLI1% = -1: FLI2% = 32767
    End If
    '
    'OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
    'If OPX% < 1 Or OPX% > 2 Then GoTo 99
    OPX% = 1 ' Solo Cantidades
    '
    Call HEADERS("LIPEDPE", "")
    VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
    If modo% = 4 Then VARIANTE$ = "Pendientes sin Orden de Despacho"
    Call HEADERS("LIPEDPE", VARIANTE$)
    '
    Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
    If NCL1% < 1 Then GoTo 99
    '
    VECPEDI$ = ""
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETA", U&)
      NC% = CVI(Mid$(X$, 7, 2))
      If NC% >= NCL1% Then
       If NC% <= NCL2% Then
        If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
          CI% = CVI(Mid$(X$, 33, 2))
          If CI% > 0 And CI% <= NUMAS% Then
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            If CAPEN > 0.05 Then
              X1$ = X$
              'Call REPLA(X1$, MKS$(U&), 109, 4)
              'LL& = LL& + 1
              'Call GRAREG("PEDDETP", X1$, LL&)
              FENTRESOL% = CVI(Mid$(X$, 25, 2))
              If FENTRESOL% >= FLI1% Then
               If FENTRESOL% <= FLI2% Then
                 If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                   Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                   Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                   Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                   Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                   '
                   CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                   If CAPEN < 0 Then CAPEN = 0
                   CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                   If CANOASI < 0 Then CANOASI = 0
                   '
                   NC% = CVI(Mid$(X$, 7, 2))
                   PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                   If OPX% <> 2 Then PREUNI = 0
                   ITOTA# = CDbl(CAPEN * PREUNI)
                   INOAS# = CDbl(CANOASI * PREUNI)
                   '
                   Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                   Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                   Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                   Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                   Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                   Call REPLA(Y$, Space$(34), 127, 34)
                   '
                   JJ& = JJ& + 1
                   Call GRAREG("LIPEDPE", Y$, JJ&)
                   VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                 End If
               End If
              End If
            End If
          End If
         End If
        End If
       End If
      End If
    Next U&
    'Call SETULTREG("PEDDETP", LL&)
    'Call CIERRARCH("PEDDETP")
    GoTo 30
    '
    FIN& = ULTREG&("ORDESPA")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ORDESPA", U&)
      CADESPAX = CVS(Mid$(X$, 79, 4))
      If CADESPAX >= 0.05 Then
        ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
        For W& = 1 To Len(VECPEDI$) / 8
          If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
            Y$ = REGLEIDO$("LIPEDPE", W&)
            ODX$ = TRIM$(Mid$(Y$, 127, 34))
            If ODX$ = "" Then
                ODX$ = "OD"
              Else
                ODX$ = ODX$ + ","
            End If
            ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
            Call REPLA(Y$, ODX$, 127, 34)
            '
            If modo% = 4 Then
               CADESPAX = CVS(Mid$(X$, 79, 4))
               CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
               CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
               Call REPLA(Y$, MKS$(CAENTX), 97, 4)
               Call REPLA(Y$, MKS$(CASALX), 101, 4)
            End If
            '
            Call GRAREG("LIPEDPE", Y$, W&)
            GoTo 19
          End If
        Next W&
      End If
19         Next U&
    '
    FIN& = ULTREG&("ODESEMI")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ODESEMI", U&)
      ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
      For W& = 1 To Len(VECPEDI$) / 8
        If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
          Y$ = REGLEIDO$("LIPEDPE", W&)
          ODX$ = TRIM$(Mid$(Y$, 127, 34))
          If ODX$ = "" Then
              CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
              If Val(CASIGNA$) >= 0.05 Then
                 Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                 Call GRAREG("LIPEDPE", Y$, W&)
              End If
          End If
          GoTo 29
        End If
      Next W&
29  Next U&
    '
30  Call HEADERS("LIPEDPE", "")
    If modo% = 2 Then
       Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
    End If
    '
    If modo% = 4 Then
      KK& = JJ&: JJ& = 0
      For U& = 1 To KK&
        Y$ = REGLEIDO$("LIPEDPE", U&)
        If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
          JJ& = JJ& + 1
          Call GRAREG("LIPEDPE", Y$, JJ&)
        End If
      Next U&
    End If
    '
    Call SETULTREG("LIPEDPE", JJ&)
    Call CIERRARCH("LIPEDPE")
    '
    VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
    If modo% = 4 Then VARIANTE$ = "Pedidos Pendientes sin Orden de Despacho"
    Call HEADERS("LIPEDPE", "Corresponde a: " + VARIANTE$)
    If VENDSEL% > 0 Then
      Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
    End If
    '
    Screen.MousePointer = 1
    '
    Sali$ = "*SOPENCL"
    If FORMAPED% = 2 Then Sali$ = "*SOQENCL"
    Call FINAL(Sali$)
99  Screen.MousePointer = 1
'
End Sub

Sub LIPEDIMAT()
    '
    Screen.MousePointer = 11
    JJ& = 0: LL& = 0
    FIN& = ULTREG&("PEDDETA")
    modo% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
    If modo% < 1 Then GoTo 99
    FLI1% = -1: FLI2% = FINMESACT%
    '
    If modo% = 2 Then
10      FNMA = FINMESACT%
        FLI1% = DIPOST(FNMA): FLI2% = 32767
        VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
        OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
        If OBX$ = "" Then GoTo 99
        '
        FLI1% = CVI(Left$(OBX$, 2))
        FLI2% = CVI(Mid$(OBX$, 3, 2))
        If FLI1% > FLI2% Then GoTo 10
        FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
        FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
      ElseIf modo% = 3 Or modo% = 4 Then
        FLI1% = -1: FLI2% = 32767
    End If
    '
    'OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
    'If OPX% < 1 Or OPX% > 2 Then GoTo 99
    OPX% = 1 ' Solo Cantidades
    '
    Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
    If NCL1% < 1 Then GoTo 99
    '
    VECPEDI$ = ""
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETA", U&)
      NC% = CVI(Mid$(X$, 7, 2))
      If NC% >= NCL1% Then
       If NC% <= NCL2% Then
        If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
          CI% = CVI(Mid$(X$, 33, 2))
          If CI% > 0 And CI% <= NUMAS% Then
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            If CAPEN > 0.05 Then
              X1$ = X$
              'Call REPLA(X1$, MKS$(U&), 109, 4)
              'LL& = LL& + 1
              'Call GRAREG("PEDDETP", X1$, LL&)
              FENTRESOL% = CVI(Mid$(X$, 25, 2))
              If FENTRESOL% >= FLI1% Then
               If FENTRESOL% <= FLI2% Then
                If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                  Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                  Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                  Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                  Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                  '
                  CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                  If CAPEN < 0 Then CAPEN = 0
                  CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                  If CANOASI < 0 Then CANOASI = 0
                  '
                  NC% = CVI(Mid$(X$, 7, 2))
                  PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                  If OPX% <> 2 Then PREUNI = 0
                  ITOTA# = CDbl(CAPEN * PREUNI)
                  INOAS# = CDbl(CANOASI * PREUNI)
                  
                  Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                  Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                  Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                  Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                  Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                  Call REPLA(Y$, Space$(34), 127, 34)
                  '
                  JJ& = JJ& + 1
                  Call GRAREG("LIPEDPE", Y$, JJ&)
                  VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                End If
               End If
              End If
            End If
          End If
         End If
        End If
       End If
      End If
    Next U&
    '
    'Call SETULTREG("PEDDETP", LL&)
    'Call CIERRARCH("PEDDETP")
    '
    GoTo 30
    '
    FIN& = ULTREG&("ORDESPA")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ORDESPA", U&)
      CADESPAX = CVS(Mid$(X$, 79, 4))
      If CADESPAX >= 0.05 Then
        ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
        For W& = 1 To Len(VECPEDI$) / 8
          If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
            Y$ = REGLEIDO$("LIPEDPE", W&)
            ODX$ = TRIM$(Mid$(Y$, 127, 34))
            If ODX$ = "" Then
                ODX$ = "OD"
              Else
                ODX$ = ODX$ + ","
            End If
            ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
            Call REPLA(Y$, ODX$, 127, 34)
            '
            If modo% = 4 Then
               CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
               CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
               Call REPLA(Y$, MKS$(CAENTX), 97, 4)
               Call REPLA(Y$, MKS$(CASALX), 101, 4)
            End If
            '
            Call GRAREG("LIPEDPE", Y$, W&)
            GoTo 19
          End If
        Next W&
      End If
19         Next U&
    '
    FIN& = ULTREG&("ODESEMI")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ODESEMI", U&)
      ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
      For W& = 1 To Len(VECPEDI$) / 8
        If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
          Y$ = REGLEIDO$("LIPEDPE", W&)
          ODX$ = TRIM$(Mid$(Y$, 127, 34))
          If ODX$ = "" Then
             CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
               If Val(CASIGNA$) > 0.05 Then
                 Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                 Call GRAREG("LIPEDPE", Y$, W&)
               End If
          End If
          GoTo 29
        End If
      Next W&
29  Next U&
    '
30  Call HEADERS("LIPEDPE", "")
    If modo% = 2 Then
       Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
    End If
    '
    If modo% = 4 Then
      KK& = JJ&: JJ& = 0
      For U& = 1 To KK&
        Y$ = REGLEIDO$("LIPEDPE", U&)
        If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
          JJ& = JJ& + 1
          Call GRAREG("LIPEDPE", Y$, JJ&)
        End If
      Next U&
      Call HEADERS("LIPEDPE", "Corresponde a: Pedidos Pendientes sin Orden de Despacho")
      If VENDSEL% > 0 Then
        Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
      End If
    End If
    '
    Call SETULTREG("LIPEDPE", JJ&)
    Call CIERRARCH("LIPEDPE")
    Screen.MousePointer = 1
    '
    Sali$ = "*SOPENMA"
    If FORMAPED% = 2 Then Sali$ = "*SOQENMA"
    Call FINAL(Sali$)
99  Screen.MousePointer = 1
    '
End Sub

Sub SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
10 VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
   OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
   If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     Exit Sub
   End If
   '
   NCL1% = CVI(Left$(OBX$, 2))
   NCL2% = CVI(Mid$(OBX$, 3, 2))
   If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
   End If
   VENDSEL% = Asc(Mid$(OBX$, 5, 1))
   '
End Sub

Function CANPREPA(CIXI%, CAVERI$)
   '
   Static DEPEX%
   '
   If DEPEX% < 1 Then
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
     If DEPEX% < 1 Then DEPEX% = 1
   End If
   '
   CAPRE = 0: CAYEN = 0
   CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
   RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
   'RFF$ = RECFILT$("MOVISTO", 5, "PP") ' SELECCIONA MOVIMIENTOS DE PREPARACION DE PEDIDOS
   'RFF$ = RFF$ + RECFILT$("MOVISTO", 5, "DM")
   JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     'Call TRACE(20, U&, FIN&)
     REMO& = CVS(Mid$(RFF$, U&, 4))
     MM$ = REGLEIDO$("MOVISTO", REMO&)
     TIMO$ = Mid$(MM$, 21, 2)
     If TIMO$ <> "PP" And TIMO$ <> "DM" Then GoTo 19
     If Asc(Mid$(MM$, 83, 1)) <> DEPEX% Then GoTo 19  ' MOVIMIENTO EN DEPODITO DE EXPEDICION
     If Mid$(MM$, 112, 6) <> CAVERJ$ Then GoTo 19     ' LEE PEDIDO E ITEM
     '
     If TIMO$ = "PP" Then
          CAPRE = CAPRE + CVD(Mid$(MM$, 88, 8)) - CVD(Mid$(MM$, 96, 8))
        ElseIf TIMO$ = "DM" Then
          CAYEN = CAYEN + CVD(Mid$(MM$, 96, 8)) - CVD(Mid$(MM$, 88, 8))
     End If
19 Next U&
   '
   CANPREPA = CAPRE - CAYEN
   If CANPREPA < 0 Then CANPREPA = 0
   '
End Function


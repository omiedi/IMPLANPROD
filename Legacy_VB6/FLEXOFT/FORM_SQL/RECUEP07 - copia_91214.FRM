VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RECUEP07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Acreedores (Gestion de Compras)"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   8
      Left            =   15000
      TabIndex        =   66
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2415
      Index           =   6
      Left            =   15000
      TabIndex        =   65
      Top             =   2400
      Width           =   15
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   3720
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   9585
      TabIndex        =   42
      Top             =   7440
      Width           =   9650
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Gral."
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
         Left            =   120
         TabIndex        =   45
         Top             =   50
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   7800
         TabIndex        =   44
         Top             =   45
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   2450
      Left            =   105
      ScaleHeight     =   2385
      ScaleWidth      =   9810
      TabIndex        =   34
      Top             =   2400
      Width           =   9870
      Begin VB.Frame Frame7 
         Caption         =   "Frame11"
         Height          =   2415
         Index           =   5
         Left            =   9465
         TabIndex        =   41
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame10"
         Height          =   2415
         Index           =   4
         Left            =   7770
         TabIndex        =   40
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame9"
         Height          =   2415
         Index           =   3
         Left            =   6195
         TabIndex        =   39
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame8"
         Height          =   2415
         Index           =   2
         Left            =   4830
         TabIndex        =   38
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2415
         Index           =   1
         Left            =   2625
         TabIndex        =   37
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame6"
         Height          =   2415
         Index           =   0
         Left            =   1250
         TabIndex        =   36
         Top             =   0
         Width           =   15
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
         Left            =   0
         TabIndex        =   35
         Top             =   300
         Width           =   9765
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Fecha   Documento   Numero                  Debe        Haber        Saldo"
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
         TabIndex        =   53
         Top             =   0
         Width           =   10650
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   0
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID1 
      Height          =   765
      Left            =   2760
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      Height          =   9435
      Left            =   10080
      ScaleHeight     =   9375
      ScaleWidth      =   915
      TabIndex        =   20
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command6 
         Caption         =   "I-Cre"
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
         Left            =   140
         Picture         =   "RECUEP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imputación de Créditos y Anticipos"
         Top             =   4485
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
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
         Left            =   140
         Picture         =   "RECUEP07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Impresión de Cuenta Corriente"
         Top             =   2400
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   25
         Top             =   6930
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   26
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Note"
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
         Left            =   140
         Picture         =   "RECUEP07.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Notas de Proveedores"
         Top             =   3090
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "R-Apl"
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
         Left            =   140
         Picture         =   "RECUEP07.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   5280
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Left            =   140
         Picture         =   "RECUEP07.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   990
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
         Left            =   140
         Picture         =   "RECUEP07.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "RECUEP07.frx":105C
         Top             =   7350
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   9825
      TabIndex        =   17
      Top             =   5040
      Width           =   9885
      Begin VB.Frame Frame7 
         Caption         =   "Frame6"
         Height          =   2415
         Index           =   10
         Left            =   1425
         TabIndex        =   52
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2415
         Index           =   11
         Left            =   3705
         TabIndex        =   51
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame8"
         Height          =   2415
         Index           =   12
         Left            =   4965
         TabIndex        =   50
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame9"
         Height          =   2415
         Index           =   13
         Left            =   6435
         TabIndex        =   49
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame10"
         Height          =   2415
         Index           =   14
         Left            =   7905
         TabIndex        =   48
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame11"
         Height          =   2415
         Index           =   15
         Left            =   9450
         TabIndex        =   47
         Top             =   0
         Width           =   15
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
         Height          =   2085
         Left            =   -30
         TabIndex        =   46
         Top             =   360
         Width           =   9795
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Documento  Número              Emision     Importe     Vencim.     Saldo ($)"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   10650
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Datos de la Cuenta Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   105
      TabIndex        =   2
      Top             =   90
      Width           =   9915
      Begin VB.Frame Frame6 
         BackColor       =   &H00F0F0C0&
         Caption         =   "Moneda - Cotizacion"
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
         Left            =   5160
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   4305
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   105
            TabIndex        =   64
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   200
            Width           =   2430
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   63
            ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2400
            TabIndex        =   62
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   315
            Width           =   510
         End
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   18
         Text            =   " "
         Top             =   1275
         Width           =   1905
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7440
         TabIndex        =   14
         Text            =   " "
         Top             =   1920
         Width           =   1905
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1245
         TabIndex        =   11
         Text            =   " "
         Top             =   1920
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   1155
         Width           =   180
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1245
         TabIndex        =   8
         Text            =   " "
         Top             =   780
         Width           =   3690
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1245
         TabIndex        =   7
         Text            =   " "
         Top             =   1155
         Width           =   3690
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1245
         TabIndex        =   6
         Text            =   " "
         Top             =   1545
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2085
         TabIndex        =   5
         Text            =   " "
         Top             =   1545
         Width           =   2850
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5235
         TabIndex        =   4
         Text            =   " "
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5235
         TabIndex        =   3
         Text            =   " "
         Top             =   1275
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         TabIndex        =   28
         Top             =   1545
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   7230
         TabIndex        =   19
         Top             =   1005
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Acreedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   7200
         TabIndex        =   15
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax-Nbr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   5130
         TabIndex        =   13
         Top             =   1665
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   5130
         TabIndex        =   9
         Top             =   1005
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECUEP07.frx":2F7E
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "RECUEP07.frx":2FF8
      Top             =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1215
      Left            =   120
      TabIndex        =   55
      Top             =   7920
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   2143
      _Version        =   393216
      BackColor       =   14737632
      BackColorFixed  =   8421504
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   8000
      TabIndex        =   60
      Top             =   9165
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   7320
      TabIndex        =   59
      Top             =   9165
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Dif.Compra Vs. Facturado:"
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
      Left            =   2160
      TabIndex        =   58
      Top             =   9165
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   4920
      TabIndex        =   57
      Top             =   9165
      Width           =   1425
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9800
      Y1              =   7850
      Y2              =   7850
   End
   Begin VB.Line Line23 
      X1              =   105
      X2              =   9765
      Y1              =   4845
      Y2              =   4845
   End
   Begin VB.Line Line22 
      X1              =   9740
      X2              =   9740
      Y1              =   2415
      Y2              =   4830
   End
   Begin VB.Line Line21 
      X1              =   9740
      X2              =   9740
      Y1              =   5020
      Y2              =   7820
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
End
Attribute VB_Name = "RECUEP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text10 = ""
    Label26 = ""
    '
    List1.Clear
    List2.Clear
    Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    '
    Hide
    '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("PROVED1")
    NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    '
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    '
    Screen.MousePointer = 11
    '
    If ULTREG&("PROVED1") > 0 Then
        X$ = REGLEIDO$("PROVED1", 1)      ' fuerza la
        Call GRAREG("PROVED1", X$, 1)     ' reindexacion
        Call CIERRARCH("PROVED1")         ' del archivo
    End If
    '
    Screen.MousePointer = 11
    '
    If Int(XVALO(Text1.TEXT)) > 0 Then
        If Int(XVALO(Text1.TEXT)) <= 32767 Then
            NCMIN% = Int(XVALO(Text1.TEXT))
            NCMAX% = NCMIN%
            GoTo 10
        End If
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("PROVED1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
10  Screen.MousePointer = 1
    RESCUP07.Text10.TEXT = TRIM$(Str$(NCMIN%))
    RESCUP07.Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    RESCUP07.Text8.TEXT = TRIM$(Str$(NCMAX%))
    RESCUP07.Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    HOO% = HOY(HO$)
    RESCUP07.Text6.TEXT = "01" + Mid$(HO$, 3)
    RESCUP07.Text5.TEXT = HO$
    '
    RESCUP07.Show 1
    '
End Sub

Sub Command4_Click()
    '
    Command4.Enabled = False
    NC% = XVALO(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       GoTo 99
    End If
    '
    'If TICUEPRO%((-1) * NC%) <> 0 Then
    '   Call MENSERR(24, "Proceso Imposible - Cuenta de Gastos")
    '   GoTo 99
    'End If
    '
    If CONTROLGRABA% = 1 Then GoTo 80
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       GoTo 99
    End If
    '
80  Call RAPLIPRO(NC%)
    Text1 = ""
    Text1 = TRIM$(Str$(NC%))
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    NCLI$ = TRIM$(Text1.TEXT)
    If NCLI$ <> "" Then
        Command5.Enabled = False
        Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Proveedor")
    End If
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    CLIPROACTIVO = (-1) * XVALO(Text1)
    Call FACOCOMP07.Command3_Click     ' Imputación de Créditos
    Text1 = ""
    Text1 = TRIM$(Str$(Abs(CLIPROACTIVO)))
    CLIPROACTIVO = 0
    Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    MONRESU% = 1
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    TICAMBIS = FORMATNUM$(TICAMONE(MONRESU%), "F9.4")
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       '
       Text7.FontName = "Lucida Console"
       Text7.FontSize = 9
       Text7.FontBold = True
       Text7.Height = 300
       'Text7.Top = 970
       '
       Text6.FontName = "Lucida Console"
       Text6.FontSize = 9
       Text6.FontBold = True
       Text6.Height = 300
       'Text6.Top = 980
       '
       Frame6.Visible = True
'       Text12.FontName = "Lucida Console"
'       Text12.FontSize = 9
'       Text12.FontBold = True
'       Text12.Height = 270
'       Text12.Top = 1260
       '
'       Label2(0).Top = 720
'       Label2(2).Top = 720
       'Label(0).Top = 1040
       'Label(1).Top = 1360
       '
'       Frame6.Height = 1670
'       Frame6.Top = 600
       '
'       MONEMIS.Visible = True
       '
       Label4.Caption = Label4.Caption + "  Saldo (U$)"
       Label4.FontName = "Lucida Console"
       Label4.FontSize = 9
       Label4.FontBold = False
       Label4.Top = 90
       List1.FontName = "Lucida Console"
       List1.FontSize = 9
       List1.FontBold = False
       '
       Label9.Caption = Label9.Caption + "  ME   Cotiz."
       Label9.FontName = "Lucida Console"
       Label9.FontSize = 9
       Label9.FontBold = False
       Label9.Top = 90
       '
       List2.FontName = "Lucida Console"
       List2.FontSize = 9
       List2.FontBold = False
       '
       Frame7(0).Left = 1100
       Frame7(1).Left = 2280
       Frame7(2).Left = 4285
       Frame7(3).Left = 5470
       Frame7(4).Left = 6830
       Frame7(5).Left = 8200
       Frame7(6).Left = 8620
       '
       Frame7(10).Left = 1250
       Frame7(11).Left = 3250
       Frame7(12).Left = 4300
       Frame7(13).Left = 5590
       Frame7(14).Left = 6850
       Frame7(15).Left = 8320
       'Frame7(16).Left = 8200
       'Frame7(17).Left = 9540
       '
       Label6.Left = 3560
       Label5.Left = 6290
       Label8.Visible = True
   End If

    RECUEP07.Height = 8300
    If InStr(EPAC$, "AHP") > 0 Then
      RECUEP07.Height = 9825 'MUESTRA LA LISTA DE FACTURAS CON LOS VALORES SEGUN ORDEN DE COMPRA.
      RECUEP07.Picture3.Height = 9435
      Call CARGA_ENCABEZADO(MSF2, "Fecha/D8;Factura/A18;Imp.Compra/F12.2;Gastos Apl./F12.2;Facturado/F12.2", SELFACTUMAT)
    End If
    '
    Call APLICACIONSKINS(Me, Picture3)
    Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    Call REMPLAGRID(Me, GRID1, Picture2, Label4, List2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub GRID2_DblClick()
   NGRID = GRID2.Row
   NLIST = List1.ListCount
   If (NLIST) >= NGRID - 1 Then
     List1.ListIndex = NGRID - 1
     List1_DblClick
   End If

End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
       End If
     End If
   End If
End Sub

Private Sub Label8_Change()
    Label12 = FORMATNUM$(XVALO(Label5) + XVALO(Label8), "F12.2")
End Sub

Private Sub List1_DblClick()
   '
   RELIS$ = List1.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 11))
   If UCase$(TIDOCUM$) <> "O.PAGO" Then
      Call MUEFACO
      Exit Sub
   End If
   '
   TIDOCUM1$ = "Orden de Pago"
   NPX& = XVALO(Mid$(RELIS$, 32, 8))
   NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM1$) > 4 Then
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
3  Call COMUNI("Imposible Mostrar Documento")
   Exit Sub
   '
5  Call MUESTRADOC(DOCUNU&)
   '
End Sub

Private Sub MONEMIS_Change()
   TICAMBIS = FORMATNUM$(TICAMONE(MONRESU%), "F9.4")
End Sub

Private Sub MONEMIS_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   If Abs(XVALO(Text1)) < 32767 Then
     NCX% = XVALO(Text1)
     If NCX% > 0 Then
       If ECOARCH$("PROVED1", MKI$(NCX%)) <> "" Then
         GoTo 20
       End If
     End If
   End If
   Exit Sub
   '
20 MONEMIAR% = MONRESU%
   MONRESU% = MONRESU% + 1
   If MONRESU% > CAMON% Then MONRESU% = 1
   MONEMIS = DMONE$(MONRESU%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONRESU% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   Call CARESCUEPRO
   '
End Sub

Private Sub MSF2_DblClick()
    'AL REALIZAAR DOBLE CLICK SE POSICIONA CON LA IDEM EN LA LISTA DE ARRIBA.
    I& = MSF2.MouseRow
    j& = MSF2.MouseCol
    If I& = 0 Or I& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = I&
    MSF2.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If

   BAND% = 0 ' BANDERA PARA VERIFICAR SI ENCONTRO EL ARGUMENTO
10 ARGU$ = UCase$(TRIM$(MSF2.TextMatrix(REG&, 2)))
   If ARGU$ <> "" Then
    Screen.MousePointer = 11
    If List1.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For I& = 1 To List1.ListCount
        If TRIM$(ARGU$) = TRIM$(Mid$(List1.List(I& - 1), 22, 18)) Then
          List1.ListIndex = I& - 1
          BAND% = 1
          Exit For
        Else
        End If
     Next I&
    End If
   End If
   
   If BAND% < 1 Then Call MENSERR(24, "Argumento No Encontrado")
   Screen.MousePointer = 1

End Sub

Private Sub Text1_Change()
   If XVALO(Text1) >= 1 Then
     If REGICLI&((-1) * XVALO(Text1)) > 0 Then
       Label26 = CODICLI$((-1) * XVALO(Text1))
'       Call CARESCUEPRO
        Timer2.Enabled = True
       Exit Sub
     End If
   End If
   '
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text10 = ""
   Label26 = ""
   '
   List1.Clear
   List2.Clear
   '
   '
   Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
   Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
   '
End Sub
Private Sub Timer2_Timer()
    Static TEXTA$, FUEAMOSTRAR%
    '
    If Text1 = TEXTA$ Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '
    FUEAMOSTRAR% = 1
    Call CARESCUEPRO
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AHP") > 0 Then
      RECUEP07.Height = 9825 'MUESTRA LA LISTA DE FACTURAS CON LOS VALORES SEGUN ORDEN DE COMPRA.
      RECUEP07.Picture3.Height = 9435
      'Call CARGA_ENCABEZADO(MSF2, "Fecha/D8;Factura/A26;Importe/F20.2;Gastos Aplicado/F20.2;LIBRE/A0", SELFACTUMAT)
      Call CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(XVALO(Text1), RECUEP07.MSF2)
    End If

    FUEAMOSTRAR% = 0
    
    Timer2.Enabled = False
    '
End Sub
Sub CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(NPROV, MSF As MSFlexGrid)
    '
    If NPROV < 1 Then Exit Sub
    MSF.Rows = 1
    '
    MSF2.Rows = 1
    Label8 = FORFAC07.Label43 ' EL TOTAL DE LA FACTURA
    DIFERENCIA# = 0
    '
    SQLX1$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    SQLX1$ = SQLX1$ + " WHERE Nume_Cli = " & NPROV & " AND TIPO_FACTURA = 'M'"
    SQLX1$ = SQLX1$ + " AND VAL_COMPRA > 0 AND VAL_COMPRA <> IMPORTE_FACTURA" ' NO MUESTRO  LAS QUE ES EL MISMO VALOR P Q NO TIENE SENTIDO
    Set RS1 = READSET(SQLX1$)
    With RS1
       Do While Not .EOF
         'ESCRIBO SIEMPRE LAS 3 CAMPOS
         j& = j& + 1
         MSF.Rows = j& + 1
         
         MSF.TextMatrix(j&, 1) = CVDAT(CVINT(!FECHEMISI))
           FACTURAPROV$ = TRIM$(SAFETEXT$(!FACTURAPROV))
         MSF.TextMatrix(j&, 2) = FACTURAPROV$
           CONDI$ = "TIPO_FACTURA = 'G' AND FACTURAPROV_ASOC = '" & FACTURAPROV$ & "'"
         APLICADO# = XVALO(VALOBADA("FACTURAYGASTOS", "SUM(IMPO_GASTO)", CONDI$, "NOMESS"))
         IMPOFAC# = XVALO(!IMPORTE_FACTURA) ' PARA MOSTRAR EL IMPORTE SIEMPRE EN PESOS MAS ALLA DE COMO SE HAYA GUARDADA
             MONEX% = XVALO(!MoneVal)
             If MONEX% > 1 Then
               TICAMB# = XVALO(!TipoCam)
               COEFCON = TICAMB# / TICAMONE(1)
               IMPOFAC# = COEFCON * IMPOFAC#
               APLICADO# = COEFCON * APLICADO#
             End If
         MSF.TextMatrix(j&, 3) = FORMATNUM$(XVALO(!VAL_COMPRA), "F12.2") 'EL IMPORTE DE LA COMPRA SIEMPRE SE GUARDA EN $
         MSF.TextMatrix(j&, 4) = TRIM$(FORMATNUM$(APLICADO#, "F12.2"))
         MSF.TextMatrix(j&, 5) = TRIM$(FORMATNUM$(IMPOFAC#, "F12.2"))
         DIFERENCIA# = DIFERENCIA# + (XVALO(!VAL_COMPRA) - IMPOFAC#)

         .MoveNext
       Loop
    End With
    Label8 = FORMATNUM$(DIFERENCIA#, "F12.2")
    Call Label8_Change
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("PROVED1")
    NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    '
End Sub

Private Sub Text10_Click()
   Text1.SetFocus
End Sub

Private Sub Text2_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text3_Click()
   Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text5_Click()
   Text1.SetFocus
End Sub

Private Sub Text6_Click()
   Text1.SetFocus
End Sub

Private Sub Text7_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text8_Click()
   Text1.SetFocus
End Sub

Private Sub Text9_CLICK()
   Text1.SetFocus
End Sub

Sub MUEFACO()
   '
   TIPODOCU$ = TRIM$(Mid$(List1.TEXT, 10, 17)) 'FC-A
   NUDOCOR$ = TRIM$(Mid$(List1.TEXT, 22, 18)) 'NUMERO DE DOC COMP
   NP1& = XVALO(Text1)
   FECHDOCU$ = TRIM$(Left$(List1.TEXT, 9))
   'FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   FECDOC1 = FETEXCOR1$(FECHANUM(TRIM$(Mid$(List1.TEXT, 1, 9)))) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   'jandy sql
   Dim SHOFACOTMP As ADODB.Recordset
   Set SHOFACOTMP = New ADODB.Recordset
   SHOFACOTMP.CursorLocation = adUseClient
   SQLZ$ = " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLZ$ = SQLZ$ + " AND NuProve =  " & NP1& & " "
   SQLX$ = " SELECT * FROM CAJABANC " + SQLZ$
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
25 On Error Resume Next
   SHOFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = TIPODOCU$ + " " + !Codicom_lar + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '\\\OT-08-988-29/10/08-FL
     SHOWFACO.Label19 = ""
     '\\\OT-08-988-FIN-FL
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CUECONTAI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DescriMov
        IDEBE1# = !IDEBECOMP
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.List1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !IDEBECOMP) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !IDEBECOMP)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !ObserGen
          '\\\OT-08-988-29/10/08-FL
          'SE CAPTURA UN POSIBLE ERROR PRODUCTO QUE EL CAMPO IDENALTA CONTENGA INFORMACION INVALIDA
          On Error Resume Next
          USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
          On Error GoTo 0
          SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
          '\\\OT-08-988-FIN-FL
        End If
        '
     .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub

Sub CARESCUEPRO()
    '
    Static ENPROCESO%
    '
    Text1.Enabled = False
5   If ENPROCESO% > 0 Then Exit Sub
    ENPROCESO% = 1
    '
10  Screen.MousePointer = 11
    NC% = XVALO(Text1)
    NCX7% = XVALO(Text1)
    NICONFLIM% = NIVCONFI%
    Command4.Enabled = True
    '
    If NC% > 0 Then
        NC1% = (-1) * NC%
        Label26 = CODICLI$(NC1%)
        Text2 = RASOCLI$(NC1%)
        Text3 = DOMICLI$(NC1%)
        Text4 = CPOSCLI$(NC1%)
        Text5 = LOCACLI$(NC1%)
        Text7 = TEL1CLI$(NC1%)
        Text10 = FORMATNUM$(SALDCLI#(NC1%), "F13.2")
        TCUE% = TICUEPRO%(NC1%)
        If TCUE% = 1 Then
            Text6.TEXT = "GASTOS"
            Picture4.Height = 5340
            Frame7(0).Height = 5340
            Frame7(1).Height = 5340
            Frame7(2).Height = 5340
            Frame7(3).Height = 5340
            Frame7(4).Height = 5340
            Frame7(5).Height = 5340
            List1.Height = 4750
          Else
            Text6.TEXT = "CORRIENTE"
            Picture4.Height = 2430
            Frame7(0).Height = 2415
            Frame7(1).Height = 2415
            Frame7(2).Height = 2415
            Frame7(3).Height = 2415
            Frame7(4).Height = 2415
            Frame7(5).Height = 2415
            List1.Height = 2085
        End If
        DoEvents
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    SALDO# = 0: SARRAS# = 0
    '
    RFF$ = RECFILT$("CUECOPRO", 1, MKI$(NC%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOPRO", RELE&)
      TIPOS$ = Mid$(XX$, 5, 16)
      IMPO# = CVD(Mid$(XX$, 25, 8))
      SALDO# = SALDO# - IMPO#
      FEX = CVI(Mid$(XX$, 3, 2))
      NROL& = CVS(Mid$(XX$, 5, 4))
      NROX$ = Mid$(XX$, 33, 24)
      '
      z$ = Space$(96)
      Call REPLA(z$, FECHATEX$(FEX), 1, 8)
      Call REPLA(z$, TIPOS$, 11, 11)
      Call REPLA(z$, NROX$, 22, 18)
      IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
      SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
      If IMPO# > 0 Then
          Call REPLA(z$, IMPTEX$, 40, 11)
        ElseIf IMPO# < 0 Then
          Call REPLA(z$, IMPTEX$, 53, 11)
      End If
      Call REPLA(z$, SALTEX$, 64, 13)
      List1.AddItem z$
    Next KU&
    DoEvents
    '
    Call APERBASDAT("CABAN")
    'jandy sql
    
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    'SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, VAL(MID(CodiCom_Lar,11)) ASC;"
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, SUBSTRING(CodiCom_Lar,11,8)ASC;"
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    If NCX7% <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = Left$(SAFETEXT$(!Codicom_lar), 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             Command4.Enabled = False            ' NO PERMITE REAPLICAR
             GoTo 899
          End If
          '
          ' AGREGADO PARA AGRUPAR O/P'S MULTIPLES A EGRESOS Y GASTOS - 27/05/06
          If List1.ListCount > 0 Then
            If TIX$ = "OP" Or TIX$ = "FF" Then
              If SAFETEXT$(!Codicom_lar) = NROX$ Then
                If CVINT%(!FECHEMISI) = FEX Then
                  SALDO# = SALDO# - IMPO#
                  IMPO# = IMPO# + XVALO(Mid$(z$, 40, 11)) - XVALO(Mid$(z$, 53, 11))
                  List1.RemoveItem List1.ListCount - 1
                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
                  Call REPLA(z$, Space$(24), 40, 24)
                  If IMPO# > 0 Then
                      Call REPLA(z$, IMPTEX$, 40, 11)
                    ElseIf IMPO# < 0 Then
                      Call REPLA(z$, IMPTEX$, 53, 11)
                  End If
                  Call REPLA(z$, SALTEX$, 64, 13)
                  List1.AddItem z$
                  GoTo 899
                End If
              End If
            End If
          End If
          ' FIN AGREGADO 27/05/06
          '
890       SALDO# = SALDO# - IMPO#
          FEX = CVINT%(!FECHEMISI)
          NROX$ = SAFETEXT$(!Codicom_lar)
          '
          z$ = Space$(96)
          Call REPLA(z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(z$, TIPOS$, 11, 11)
          Call REPLA(z$, NROX$, 22, 18)
          IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
          SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
          If IMPO# > 0 Then
              Call REPLA(z$, IMPTEX$, 40, 11)
            ElseIf IMPO# < 0 Then
              Call REPLA(z$, IMPTEX$, 53, 11)
          End If
          Call REPLA(z$, SALTEX$, 64, 13)
          List1.AddItem z$
          '
899      .MoveNext
        Loop
      End If
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    List1.Refresh
    Openforms = DoEvents
    '
    Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    If TCUE% <> 0 Then GoTo 999
    '
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM SACREPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (abs(SaldAcre) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SaldDebe) >= 0.005) "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
26  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             GoTo 993
          End If
          TIX$ = UCase$(Left$(SAFETEXT$(!Codicom_lar), 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = (-1) * XVALO(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
          SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = SAFETEXT$(!Codicom_lar)
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
          '
          z$ = Space$(80)
          Call REPLA(z$, TIPOS$, 1, 11)
          Call REPLA(z$, NROX$, 12, 18)
          Call REPLA(z$, FECHATEX$(FEMI), 32, 8)
          Call REPLA(z$, IMPTEX$, 40, 12)
          Call REPLA(z$, FECHATEX$(FEVE), 56, 8)
          Call REPLA(z$, SALTEX$, 65, 12)
          '
          List2.AddItem z$
          '
993     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       If Command4.Enabled = True Then
         If CONTROLGRABA% = 0 Then
           ATENTI% = 1
          '
          ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
          ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR PAGAR APARECIA
          ' DESACTUALIZADO.
          '
          Call RAPLIPRO(NC%)
          GoTo 100
'           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar y Reconstruir\Continuar sin Reaplicar"
'           OPX1% = COMALTER%(TTXX$)
'           If OPX1% = 2 Then 'CONTINUAR Y RECONSTRUIR
'               Call RAPLIPRO(NC%)
'               GoTo 100
'           ElseIf OPX1% = 3 Then ' CONTINUAR SIN REAPLICAR
'               GoTo 995
'           Else
'               ATENTI% = 0 '  CANCELA
'               Call CIERRARCH("*.*")
'               Call FINAL("")
'           End If
         End If
       End If
    End If
    '
995 List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label5.Refresh
    Label6.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label6.Refresh
    '
    Call REMPLAGRID(Me, GRID1, Picture2, Label4, List2)
    '
999 Screen.MousePointer = 1
    ENPROCESO% = 0
    Text1.Enabled = True
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    '
End Sub
'

Private Sub Time2_Timer()

End Sub



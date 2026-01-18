VERSION 5.00
Begin VB.Form RECUEP04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Acreedores (Gestion de Compras)"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   8
      Left            =   15000
      TabIndex        =   57
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   14
      Left            =   7905
      TabIndex        =   56
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   13
      Left            =   6435
      TabIndex        =   55
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   12
      Left            =   4965
      TabIndex        =   54
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   11
      Left            =   3705
      TabIndex        =   53
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   10
      Left            =   1425
      TabIndex        =   52
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   15
      Left            =   9450
      TabIndex        =   51
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   6
      Left            =   15000
      TabIndex        =   48
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   5
      Left            =   9465
      TabIndex        =   47
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   4
      Left            =   7770
      TabIndex        =   46
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   3
      Left            =   6195
      TabIndex        =   45
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   2
      Left            =   4830
      TabIndex        =   44
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   1
      Left            =   2625
      TabIndex        =   43
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2450
      Index           =   0
      Left            =   1260
      TabIndex        =   42
      Top             =   2415
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      Height          =   7995
      Left            =   9985
      ScaleHeight     =   7935
      ScaleWidth      =   915
      TabIndex        =   32
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   975
      Begin VB.PictureBox Picture16 
         Height          =   7935
         Left            =   0
         ScaleHeight     =   7875
         ScaleWidth      =   1005
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   7905
            Left            =   0
            Picture         =   "RECUEP04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "I-Cre"
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
         Left            =   140
         Picture         =   "RECUEP04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imputación de Créditos y Anticipos"
         Top             =   4485
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   140
         Picture         =   "RECUEP04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   39
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
         TabIndex        =   37
         Top             =   6930
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Note"
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
         Left            =   140
         Picture         =   "RECUEP04.frx":2B4C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Notas de Proveedores"
         Top             =   3090
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "R-Apl"
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
         Left            =   140
         Picture         =   "RECUEP04.frx":2E56
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   5175
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP04.frx":2FA0
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   990
         Width           =   650
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
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP04.frx":32AA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "RECUEP04.frx":33F4
         Top             =   7350
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   105
      ScaleHeight     =   2370
      ScaleWidth      =   9615
      TabIndex        =   16
      Top             =   2415
      Width           =   9675
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   18
            Top             =   50
            Width           =   15000
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6500
         Left            =   0
         TabIndex        =   19
         Top             =   -1050
         Width           =   10515
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
            Height          =   2100
            IntegralHeight  =   0   'False
            Left            =   -50
            TabIndex        =   20
            Top             =   1365
            Width           =   9670
         End
         Begin VB.Line Line7 
            X1              =   1155
            X2              =   1155
            Y1              =   1470
            Y2              =   3255
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2800
      Left            =   105
      ScaleHeight     =   2745
      ScaleWidth      =   9615
      TabIndex        =   22
      Top             =   4990
      Width           =   9675
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5630
         Left            =   0
         TabIndex        =   25
         Top             =   315
         Width           =   10515
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
            Height          =   2090
            IntegralHeight  =   0   'False
            Left            =   -30
            TabIndex        =   26
            Top             =   0
            Width           =   9675
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7665
            TabIndex        =   58
            Top             =   2130
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Line Line15 
            X1              =   0
            X2              =   9660
            Y1              =   2420
            Y2              =   2420
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   9660
            Y1              =   2085
            Y2              =   2085
         End
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
            Left            =   25
            TabIndex        =   29
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4330
            TabIndex        =   28
            Top             =   2130
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7455
            TabIndex        =   27
            Top             =   2130
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   10515
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Left            =   0
            TabIndex        =   24
            Top             =   50
            Width           =   10410
         End
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   8820
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   5460
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
      Height          =   2040
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   9675
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
         Left            =   5250
         TabIndex        =   49
         Top             =   35
         Visible         =   0   'False
         Width           =   4425
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
            Left            =   2205
            TabIndex        =   60
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   315
            Width           =   750
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
            Left            =   3150
            TabIndex        =   59
            ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
            Top             =   230
            Width           =   1050
         End
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
            TabIndex        =   50
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   200
            Width           =   2430
         End
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   30
         Text            =   " "
         Top             =   800
         Width           =   1905
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   14
         Text            =   " "
         Top             =   1560
         Width           =   1905
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1365
         TabIndex        =   11
         Text            =   " "
         Top             =   1560
         Width           =   3690
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   800
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
         Left            =   1050
         TabIndex        =   1
         Top             =   800
         Width           =   180
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1365
         TabIndex        =   8
         Text            =   " "
         Top             =   420
         Width           =   3690
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1365
         TabIndex        =   7
         Text            =   " "
         Top             =   800
         Width           =   3690
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1365
         TabIndex        =   6
         Text            =   " "
         Top             =   1180
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2205
         TabIndex        =   5
         Text            =   " "
         Top             =   1180
         Width           =   2850
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5355
         TabIndex        =   4
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5355
         TabIndex        =   3
         Text            =   " "
         Top             =   800
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Height          =   300
         Left            =   210
         TabIndex        =   40
         Top             =   1180
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   7350
         TabIndex        =   31
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Acreedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   7350
         TabIndex        =   15
         Top             =   1300
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax-Nbr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   5250
         TabIndex        =   13
         Top             =   1300
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5250
         TabIndex        =   9
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.Line Line23 
      X1              =   105
      X2              =   9765
      Y1              =   4845
      Y2              =   4845
   End
   Begin VB.Line Line22 
      X1              =   9775
      X2              =   9775
      Y1              =   2415
      Y2              =   4830
   End
   Begin VB.Line Line21 
      X1              =   9775
      X2              =   9775
      Y1              =   4960
      Y2              =   7770
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7560
      Picture         =   "RECUEP04.frx":5316
      Top             =   7260
      Width           =   2010
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu imputaciondecreditos 
         Caption         =   "Imputación de Créditos y Anticipos"
      End
      Begin VB.Menu imprimircc 
         Caption         =   "Imprimir Cuenta Corriente"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Begin VB.Menu reaplicacion 
         Caption         =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu notadProveedores 
         Caption         =   "Nota de Proveedores"
      End
   End
End
Attribute VB_Name = "RECUEP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%
Dim MONRESU%

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
    '
    Hide
    '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("PROVED1")
    NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    If NC% > 0 Then Text1.TEXT = TRIM$(Str$(NC%))
    'Call CARESCUEPRO
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
    RESCUP04.Text10.TEXT = TRIM$(Str$(NCMIN%))
    RESCUP04.Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    RESCUP04.Text8.TEXT = TRIM$(Str$(NCMAX%))
    RESCUP04.Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    HOO% = HOY(HO$)
    RESCUP04.Text6.TEXT = "01" + Mid$(HO$, 3)
    RESCUP04.Text5.TEXT = HO$
    '
    RESCUP04.Show 1
    '
End Sub

Sub Command4_Click()
    '
    '
    If MONRESU% > 1 Then
      Call MENSERR(24, "Proceso Imposible - Cambie Moneda a '" + TRIM$(DEMONECO$(1)) + "'")
      Exit Sub
    End If
    '
    Command4.Enabled = False
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    '
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
    Call FACOCOMP04.Command3_Click     ' Imputación de Créditos
    Text1 = ""
    Text1 = TRIM$(Str$(Abs(CLIPROACTIVO)))
    CLIPROACTIVO = 0
    Command6.Enabled = True
End Sub

Private Sub Form_Activate()
Exit Sub
    If XVALO(Text1.TEXT) > 0 Then
        NC% = XVALO(Text1.TEXT)
      ElseIf NC% > 0 Then
        Text1.TEXT = NC%
    End If
    List1.Clear
    List2.Clear
    If NC% > 0 Then
      If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
        Call CARESCUEPRO
      End If
    End If
End Sub

Private Sub Form_Load()
    '
     If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7(0).BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 7500
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    EPAC$ = TRIM$(EMPREAC$)
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
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       '
       Text7.FontName = "Lucida Console"
       Text7.FontSize = 9
       Text7.FontBold = True
       Text7.Height = 300
       Text7.Top = 970
       '
       Text6.FontName = "Lucida Console"
       Text6.FontSize = 9
       Text6.FontBold = True
       Text6.Height = 300
       Text6.Top = 980
       '
       Frame6.Visible = True
'       Text12.FontName = "Lucida Console"
'       Text12.FontSize = 9
'       Text12.FontBold = True
'       Text12.Height = 270
'       Text12.Top = 1260
       '
       Label2(0).Top = 720
       Label2(2).Top = 720
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
'    For K% = 0 To Forms.Count - 1
'        If Forms(K%).Name = PROLLAMA$ Then
'            Forms(K%).Show
'            PROLLAMA$ = ""
'            Exit Sub
'          ElseIf Forms(K%).Name = "COGECOMP" Then
'            COGECOMP.Show
'            Exit Sub
'        End If
'    Next K%
End Sub

Private Sub imprimircc_Click()
    Call Command3_Click
End Sub

Private Sub imputaciondecreditos_Click()
    Call Command6_Click
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

Private Sub List1_DBLClick()
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
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
'5  For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   '
'   REDOCU$ = REGACT$("PDOCLST")
'   URE& = ULTREG&("PDOCSPL")
'   LI& = 0: LS& = URE&
'6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'   GoTo 8
'   '
'7  Screen.MousePointer = 1
'   Call MENSERR(24, "Documento no Registrado")
'   Exit Sub
'   '
'8  TPSP$ = ""
'   Screen.MousePointer = 11
'   While L& > 1
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then GoTo 9
'     L& = L& - 1
'   Wend
'   '
'9  CAPAGINAS% = 0
'   For KKL& = L& To URE&
'     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'     End If
'   Next KKL&
'   '
'95 SPOOLSTRING$ = TPSP$
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   ALTUPAGINA = 7170
'   TOPEPAGINA = 0
'   PAGINACTIVA% = 0
'   CAPAGINAS% = 0
'   '
'   FIN& = Len(TPSP$)
'10 PPXX% = InStr(TPSP$, Chr$(255))
'   If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'   End If
'   '
'   Screen.MousePointer = 1
'   PAGINACTIVA% = 0
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Visible = False
'     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'   Next KKI%
'   '
'   EDITFORM.Picture1(PAGINACTIVA%).Refresh
'   EDITFORM.VScroll1.Min = 0
'   EDITFORM.VScroll1.Value = 0
'   EDITFORM.VScroll1.Max = 1
'   If ALTUPAGINA > EDITFORM.Frame1.Height Then
'      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'   End If
'   EDITFORM.Command3.Enabled = False
'   EDITFORM.Command5.Enabled = False
'   EDITFORM.Show 1
'   '
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

Private Sub notadProveedores_Click()
    Call Command5_Click
End Sub

Private Sub reaplicacion_Click()
    Call Command4_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_Change()
   If XVALO(Text1) >= 1 Then
     If REGICLI&((-1) * XVALO(Text1)) > 0 Then
       Label26 = CODICLI$((-1) * XVALO(Text1))
       Call CARESCUEPRO
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
End Sub

Private Sub Text1_DblClick()
    '
    Call Command2_Click
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
   TIPODOCU$ = TRIM$(Mid$(List1.TEXT, 10, 17))
   NUDOCOR$ = TRIM$(Mid$(List1.TEXT, 22, 18)) 'NUMERO DE DOC COR
   NP1& = XVALO(Text1)
   FECHDOCU$ = TRIM$(Left$(List1.TEXT, 9))
   FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CDBL(FechEmisi) = " & CDbl(FECDOC1) & " "
   SQLX$ = SQLX$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLX$ = SQLX$ + " AND NuProve =  " & NP1& & " "
   Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
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
     SHOWFACO.Caption = TIPODOCU$ + " " + !CodiCom_Lar + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CueContaI)
        QQQ2$ = "": If IsNull(!CENCOSTO) = False Then QQQ2$ = !CENCOSTO
        REFE1$ = !Descrimov
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.List1.AddItem ADDLIS1$
        End If
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !Idebecomp) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !Idebecomp)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !OBSERGEN
        End If
        '
     .MoveNext
     Loop
   End With
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub

Sub CARESCUEPRO()
    '
10  Screen.MousePointer = 11
    NC% = XVALO(Text1)
    NCX7% = XVALO(Text1)
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
            Picture1.Height = 5340
            List1.Height = 5005
            For KKI% = 0 To 5: Frame7(KKI%).Height = 5300: Next KKI%
            For KKI% = 10 To 15: Frame7(KKI%).Visible = False: Next KKI%
            Line22.Y2 = 5330
          Else
            For KKI% = 0 To 5: Frame7(KKI%).Height = 2450: Next KKI%
            For KKI% = 10 To 15: Frame7(KKI%).Visible = True: Next KKI%
            Line22.Y2 = 4830
            Text6.TEXT = "CORRIENTE"
            Picture1.Height = 2430
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
      Z$ = Space$(96)
      Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
      Call REPLA(Z$, TIPOS$, 11, 11)
      Call REPLA(Z$, NROX$, 22, 18)
      IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
      SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
      If IMPO# > 0 Then
          Call REPLA(Z$, IMPTEX$, 40, 11)
        ElseIf IMPO# < 0 Then
          Call REPLA(Z$, IMPTEX$, 53, 11)
      End If
      Call REPLA(Z$, SALTEX$, 64, 13)
      '
      List1.AddItem Z$
    Next KU&
    DoEvents
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, VAL(MID(CodiCom_Lar,11)) ASC;"
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    If NCX7% <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = Left$(!CodiCom_Lar, 2)
          COCOLA$ = !CodiCom_Lar
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
          IMPO# = !Idebecomp - !IHABECOMP
          '
          ' AGREGADO PARA AGRUPAR O/P'S MULTIPLES A EGRESOS Y GASTOS - 27/05/06
          If List1.ListCount > 0 Then
            If TIX$ = "OP" Or TIX$ = "FF" Then
              If !CodiCom_Lar = NROX$ Then
                If CVINT%(!FechEmisi) = FEX Then
                  SALDO# = SALDO# - IMPO#
                  IMPO# = IMPO# + XVALO(Mid$(Z$, 40, 11)) - XVALO(Mid$(Z$, 53, 11))
                  List1.RemoveItem List1.ListCount - 1
                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
                  Call REPLA(Z$, Space$(24), 40, 24)
                  If IMPO# > 0 Then
                      Call REPLA(Z$, IMPTEX$, 40, 11)
                    ElseIf IMPO# < 0 Then
                      Call REPLA(Z$, IMPTEX$, 53, 11)
                  End If
                  Call REPLA(Z$, SALTEX$, 64, 13)
                  List1.AddItem Z$
                  GoTo 899
                End If
              End If
            End If
          End If
          '
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!FechEmisi)
              NROL& = !NumeComp
              Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
              Call REPLA(Z$, TIPOS$, 11, 11)
              Call REPLA(Z$, NROX$, 22, 18)
                TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
                While Len(TTX1$) < 23: TTX1$ = "." + TTX1$ + ".": Wend
              Call REPLA(Z$, " " + Left$(TTX1$, 23) + " ", 40, 25)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          '
890       SALDO# = SALDO# - IMPO#
          FEX = CVINT%(!FechEmisi)
          NROX$ = !CodiCom_Lar
          '
          Z$ = Space$(96)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(Z$, TIPOS$, 11, 11)
          Call REPLA(Z$, NROX$, 22, 18)
          IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
          SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
          If IMPO# > 0 Then
              Call REPLA(Z$, IMPTEX$, 40, 11)
            ElseIf IMPO# < 0 Then
              Call REPLA(Z$, IMPTEX$, 53, 11)
          End If
          Call REPLA(Z$, SALTEX$, 64, 13)
          '
          If MODOMONEI% = 2 Then
              MONEM% = ROUND(!Mone_Emis, 0) Mod 6   ' enjuague gramy
              TICAM = ROUND(!Tipo_Cam, 5)
              If TICAM < 0.05 Then TICAM = 1
              '
              Call REPLA(Z$, SIMBOLPES$(MONEM%), 79, 2)
              If MONEM% > 1 Then
                 Call REPLA(Z$, FORMATNUM$(TICAM, "F9.4"), 81, 9)
              End If
              Call REPLA(Z$, FORMATNUM$(MONEM%, "I2"), 103, 2)
              Call REPLA(Z$, COCOLA$, 105, 24)
              '
          End If
          '
895       List1.AddItem Z$
          '
899      .MoveNext
        Loop
      End If
990 End With
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    If TCUE% <> 0 Then GoTo 999
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (abs(SaldAcre) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SaldDebe) >= 0.005) "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SUDEUDOL# = 0
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
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
          IMPO# = (-1) * !IBruComp * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = !SaldAcre - !SaldDebe
          SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = !CodiCom_Lar
          FEMI = CVINT(!FechEmisi)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
          '
          SALTEXDOL$ = ""
          COCOLAR$ = !CodiCom_Lar
          '
          If MODOMONEI% = 2 Then
             SALDOLA# = ROUND(!SaldAcre_Dola - !SaldDebe_Dola)
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          '
20        Z$ = Space$(90)
          Call REPLA(Z$, TIPOS$, 1, 11)
          Call REPLA(Z$, NROX$, 12, 18)
          Call REPLA(Z$, FECHATEX$(FEMI), 32, 8)
          Call REPLA(Z$, IMPTEX$, 40, 12)
          Call REPLA(Z$, FECHATEX$(FEVE), 56, 8)
          Call REPLA(Z$, SALTEX$, 65, 12)
          Call REPLA(Z$, SALTEXDOL$, 77, 13)
          SUDEUDOL# = SUDEUDOL# + SALDOLA#
          '
          List2.AddItem Z$
          '
        .MoveNext
        Loop
      End If
    End With
    '
    If MONRESU% <= 1 Then
       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
         If CONTROLGRABA% = 0 Then
           ATENTI% = 1
           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
           If COMALTER%(TTXX$) = 2 Then
               Call RAPLIPRO(NC%)
               GoTo 100
             Else
               ATENTI% = 0
               Call CIERRARCH("*.*")
               Call FINAL("")
           End If
         End If
       End If
    End If
    '
    List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label5.Refresh
    Label6.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label6.Refresh
    Label8.Caption = CSTRING$(MKD$(SUDEUDOL#), 4, 14, 2, 2)
    Label8.Refresh
    '
999 Screen.MousePointer = 1
    '
End Sub
'


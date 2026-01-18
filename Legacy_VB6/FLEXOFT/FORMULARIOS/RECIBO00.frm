VERSION 5.00
Begin VB.Form RECIBO00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza en Cuenta Corriente"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command8 
         Caption         =   "Opción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   10920
         Picture         =   "RECIBO00.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
         Top             =   4535
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Next"
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
         Left            =   10920
         Picture         =   "RECIBO00.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Continúa Carga de Recibo."
         Top             =   735
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
         Left            =   10500
         TabIndex        =   50
         Top             =   735
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   435
         Left            =   10670
         TabIndex        =   51
         Top             =   670
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
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
         Height          =   850
         Left            =   10920
         Picture         =   "RECIBO00.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   5460
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2670
         Left            =   210
         ScaleHeight     =   2610
         ScaleWidth      =   10395
         TabIndex        =   32
         Top             =   1760
         Width           =   10450
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   10515
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line22 
               X1              =   4935
               X2              =   4935
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "  M-P    Medio de Pago      Numero              P/T       Cl.         Emision           Vencim.               Importe"
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
               TabIndex        =   34
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line20 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   4200
               X2              =   4200
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line7 
               X1              =   2625
               X2              =   2625
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line5 
               X1              =   7140
               X2              =   7140
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2600
            Left            =   -315
            TabIndex        =   35
            Top             =   0
            Width           =   10800
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
               Height          =   2310
               Left            =   285
               TabIndex        =   36
               Top             =   315
               Width           =   10500
            End
            Begin VB.Line Line21 
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
         Left            =   210
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   22
         Top             =   4725
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento             Nro      Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               TabIndex        =   29
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   1680
               X2              =   1680
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   23
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
               Height          =   2085
               Left            =   -30
               TabIndex        =   24
               Top             =   0
               Width           =   10450
            End
            Begin VB.ListBox List3 
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
               TabIndex        =   45
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label10 
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
               Left            =   8295
               TabIndex        =   31
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label11 
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
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Line Line12 
               X1              =   1680
               X2              =   1680
               Y1              =   1575
               Y2              =   3360
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
               Left            =   6615
               TabIndex        =   27
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label4 
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
               Left            =   3570
               TabIndex        =   26
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
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
               TabIndex        =   25
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10450
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3730
               X2              =   3730
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   2415
            End
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   980
         Left            =   10920
         Picture         =   "RECIBO00.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba - Graba e Imprime Recibo"
         Top             =   7140
         Width           =   855
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6195
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
      End
      Begin VB.CommandButton Command11 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   10920
         Picture         =   "RECIBO00.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1740
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C.Cte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   10920
         Picture         =   "RECIBO00.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2685
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   10920
         Picture         =   "RECIBO00.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Notas de Clientes"
         Top             =   3610
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "RECIBO00.frx":1956
         TabIndex        =   16
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text11 
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
         Left            =   9030
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   1650
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   3
         Text            =   " "
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10500
         TabIndex        =   4
         Top             =   315
         Width           =   180
      End
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
         Height          =   300
         Left            =   2310
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         Left            =   1470
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3690
      End
      Begin VB.TextBox Text2 
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
         Left            =   1470
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   2
         Top             =   1155
         Width           =   180
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
         TabIndex        =   1
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Line Line29 
         X1              =   10650
         X2              =   11900
         Y1              =   7030
         Y2              =   7030
      End
      Begin VB.Line Line28 
         X1              =   210
         X2              =   10660
         Y1              =   4420
         Y2              =   4420
      End
      Begin VB.Line Line27 
         X1              =   10660
         X2              =   10660
         Y1              =   1740
         Y2              =   4420
      End
      Begin VB.Line Line26 
         X1              =   10665
         X2              =   10665
         Y1              =   4730
         Y2              =   7510
      End
      Begin VB.Line Line25 
         X1              =   10650
         X2              =   11900
         Y1              =   6500
         Y2              =   6500
      End
      Begin VB.Line Line24 
         X1              =   10650
         X2              =   11900
         Y1              =   6495
         Y2              =   6495
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   10320
         Picture         =   "RECIBO00.frx":1C60
         Top             =   6510
         Width           =   2010
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000016&
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
         Left            =   8505
         TabIndex        =   49
         Top             =   735
         Visible         =   0   'False
         Width           =   3330
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talonario:"
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
         Left            =   7035
         TabIndex        =   48
         Top             =   735
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   8610
         TabIndex        =   44
         Top             =   7770
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipos:"
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
         Left            =   6300
         TabIndex        =   43
         Top             =   7875
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   5145
         TabIndex        =   42
         Top             =   7770
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicados:"
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
         Left            =   2835
         TabIndex        =   41
         Top             =   7875
         Width           =   2220
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   3780
         TabIndex        =   40
         Top             =   5775
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   1470
         TabIndex        =   39
         Top             =   5880
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   1785
         TabIndex        =   38
         Top             =   7770
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   -525
         TabIndex        =   37
         Top             =   7875
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
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
         Left            =   5355
         TabIndex        =   21
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
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
         Left            =   7560
         TabIndex        =   15
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   8505
         TabIndex        =   14
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   8
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      TabIndex        =   13
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "RECIBO00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%
Dim MODORECI%

Private Sub BOTNEXT_Click()
    '
    Call GRACONTROL("RECISTA", "COBRACTI", TRIM$(Label22.Caption))
    Call SIGUERECI1
    Command5.Enabled = True
    '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call NOTACLI(NC%, 1)
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCH_GES/AMACLI", "Archivos Maestros")
99  Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
        Else
          If TRIM$(CONTROL$("RECISTA", "AVAHOJA")) = "NEWPAGE" Then
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
          End If
          Call CIERRARCH("*.*")
          Call BAJALDISCO
          Hide
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then Call NOTACLI(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC00.Text1.TEXT = Text1.TEXT
    RECUEC00.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    If NC% <> Val(Text1.TEXT) Then
       Call MENSERR(24, "Imposible Grabar - Intente de Nuevo")
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    TOTAVAL# = CDbl(Val(Label14.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption))
    If TOTACAN# > TOTAVAL# Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder la Suma de Valores Recibidos.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    FF% = FECHANUM(Text9.TEXT)
    If FF% < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        Text9.SetFocus
        Exit Sub
    End If
    '
    Call GRARECISTA(OKEY%)
    If OKEY% = 0 Then
        Command5.Enabled = True
        Exit Sub
      ElseIf OKEY% <> 6 Then
        Call COMUNI("La Registración de este Recibo\no Pudo Completarse. (" + TRIM$(Str$(OKEY%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    Call PRIRECISTA
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    Call CIERRARCH("!CARMECO")
    Call BLANARCH("!CARMECO")
    Call CIERRARCH("!APLIRECI")
    Call BLANARCH("!APLIRECI")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    Text11.SetFocus
    '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RECICO00.Show 1
   '
   NROPRO& = 0
   Call NURECIBO(NROPRO&)
   Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
   Text11.SelStart = 3
   Text11.SelLength = Len(Text11.TEXT) - 3
   '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Form_Load()
    '
    Dim BUFLE As String * 64
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Refresh
    Frame1.Refresh
    DoEvents
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo Standard de Cobranza:\  \'" + RMCC$ + "'.")
        Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
        Call CIERRARCH("*.*")
        Call FINAL("")
      End If
    End If
    '
    CUEDEUVE$ = TRIM$(CONTROL$("AMACLI", "ECUECON"))
    If CUEDEUVE$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUEDEUVE$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    CUECHECA$ = TRIM$(CONTROL$("", "CUECHECA"))
    If CUECHECA$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUECHECA$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
20  If ULTREG&("MECOBRA") < 1 Then
      TTXX$ = "Falta Tabla de Medios de Cobranza\Puede Instalar una Tabla Preconfigurada"
      OPX% = COMALTER%(TTXX$ + "\\Instalar Ahora\Carga Manual")
      If OPX% = 1 Then
         NX1% = FILEOPEN%(CurDir$ + "\MECOBRA.DRV", 0, 64)
         For U& = 1 To LOF(NX1%) / 64
           Get #NX1%, U&, BUFLE$
           NMP% = CVI(Left$(BUFLE$, 2))
           If NMP% > 0 Then
             If NMP% < 900 Then
               XXX$ = BUFLE$
               Call REPLA(XXX$, CUECHECA$, 45, 12)
               Call REGAPP("MECOBRA", XXX$)
             End If
           End If
         Next U&
         Call CIERRARCH("MECOBRA")
         Close #NX1%
         Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf OPX% = 2 Then
         Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
      End If
      GoTo 20
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    If ULTREG&("COBRADOR") > 0 Then
       Label21.Visible = True
       Label22.Visible = True
       Command7.Visible = True
       XXX$ = REGLEIDO$("COBRADOR", 1)
       RECIBO00.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
    End If
    '
    AUTONUME% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONUME% = 1
    End If
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    If AUTONUME% <> 0 Then
      Text1.TEXT = "     "
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
    End If
    '
    Refresh
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BLOQEXE("*")
     Call LIBARCH("*.*")
     Reset
     Call BLOQEXE(App.EXEName)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Label2_DblClick()
   Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label22_DblClick()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub List1_DblClick()
    '
    Call SIGUERECI2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(Val(Label14.Caption)) - CDbl(Val(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If Val(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = Val(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    '
    ANTI# = CDbl(Val(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Anticipos:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUERECI3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Text1_Change()
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    BOTNEXT.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call NOTACLI(NC%, 1)
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$(NC%)
                Text3.TEXT = DOMICLI$(NC%)
                Text4.TEXT = CPOSCLI$(NC%)
                Text5.TEXT = LOCACLI$(NC%)
                Call NOTACLI(NC%, 1)
                On Error Resume Next
                BOTNEXT.SetFocus
                On Error GoTo 0
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_GotFocus()
    If AUTONUME% <> 0 Then
      NROPRO& = 0
      Call NURECIBO(NROPRO&)
      Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    '
    Text11.SelStart = 3
    On Error Resume Next
    Text11.SelLength = Len(Text11.TEXT) - 3
    On Error GoTo 0
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "RB." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    If NRONUE& > 999999 Then
        Call MENSERR(24, "Número de Recibo Fuera de Rango")
        Call NURECIBO(NROPRO&)
        GoTo 95
    End If
    '
    NROPRO& = NRONUE&
    Call NURECIBO(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    If FORMALTER.Visible = True Then Exit Sub
    If OPRECIB00.Visible = True Then Exit Sub
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
95  Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
   Call CONECRUN("FLSETUP/OPRECIBO", "Configuración de Funcionamiento")
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text7_GotFocus()
    Text10.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Sub NURECIBO(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENURECI")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    AUTONURE% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONURE% = 1
    End If
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = Val(Label22.Caption)
    On Error GoTo 0
    '
    CAMOLE% = 0: CAMOCOBRA% = 0
    NROMA& = 0
    For U& = ULTREG&("COBRAN") To 1 Step -1
        XX$ = REGLEIDO$("COBRAN", U&)
        CAMOLE% = CAMOLE% + 1
        If CAMOLE% > 512 Then GoTo 90
        If CAMOLE% > 128 Then
          If CAMOCOBRA% > 16 Then
            GoTo 90
          End If
        End If
        '
        FEMO% = CVI(Mid$(XX$, 7, 2))
        COBR% = CVI(Mid$(XX$, 23, 2))
        If FEMO% >= FLIM% Then
          If (COBR% = COBRACT% Or COBRACT% = 0) Then
            NROLE& = CVS(Mid$(XX$, 3, 4))
            If NROLE& = NROPRO& Then
              If CVD(Mid$(XX$, 15, 8)) >= 0.005 Or AUTONURE% = 1 Then
                NROPRO& = 0
              End If
            End If
            If NROLE& > NROMA& Then NROMA& = NROLE&
            CAMOCOBRA% = CAMOCOBRA% + 1
          End If
        End If
    Next U&
    '
90  If AUTONURE% = 1 Then
      If COBRACT% = UCOBRA% Then
        If NROMA& < UNUREC& Then
          NROMA& = UNUREC&
        End If
      End If
    End If
    '
    If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'
Sub SIGUERECI1()
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 30 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "NEUMAN") < 1 Or Val(Left$(TRIM$(Label22.Caption), 2)) <> 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
      End If
    End If
   '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        'Text7.TEXT = TEL1CLI$(NC%)
        'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If MODORECI% = 1 Then
             FF% = CVI(Mid$(X$, 19, 2)) ' orden emision
           Else
             FF% = CVI(Mid$(X$, 9, 2)) ' orden vencimiento
        End If
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Desc. Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                '
                If SALDO# >= 0.005 Then
                    SUMAFAC# = SUMAFAC# + IMPO#
                    SUMADEU# = SUMADEU# + SALDO#
                    NROL& = CVS(Mid$(X$, 5, 4))
                    FEMI = CVI(Mid$(X$, 19, 2))
                    FEVE = CVI(Mid$(X$, 9, 2))
                    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, NUSUCUVE%)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, TIPOS$, 1, 14)
                    Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 31, 14)
                    Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
                    Call REPLA(Z$, SALTEX$, 56, 14)
                    Call REPLA(Z$, DOCU$, 97, 18)
                    Call REPLA(Z$, TRIM$(Str$(REG&)), 121, 6)
                    '
                    List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    List2.ListIndex = -1
    List2.Refresh
    'Label14.Caption = CSTRING$(MKD$(SUMAFAC#), 4, 14, 2, 2)
    'Label14.Refresh
    Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    BOTNEXT.Enabled = False
    Call BLANARCH("!CARMECO")
    Screen.MousePointer = 1
    '
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso de Medios de Pago Codificados"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("CARMECOW" + ATRI$)
    Call CIERRARCH("!CARMECO")
    AAA = ULTREG&("!CARMECO")
    If VTB% < 0 Then
        BOTNEXT.Enabled = True
        Exit Sub
    End If
    '
    
    List1.Clear
    SUMAVA# = 0
    Label14 = ""
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        If IMVA# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        If UCase$(Left$(BANCO$, 2)) = "CH" Then
          FEMI = CVI(Mid$(X$, 39, 2))
          F1 = FEMI: F2 = HOI%
          If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
            Call MENSERR(24, "Fecha de Emisión no Válida")
            Call REPLA(X$, MKI$(HOI%), 39, 2)
            Call GRAREG("!CARMECO", X$, U&)
            GoTo 200
          End If
          '
          FEVE = CVI(Mid$(X$, 41, 2))
          F1 = FEVE: F2 = HOI%
          If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Then
            Call MENSERR(24, "Fecha de Vencimiento no Válida")
            Call REPLA(X$, MKI$(HOI%), 41, 2)
            Call GRAREG("!CARMECO", X$, U&)
            GoTo 200
          End If
        End If
        '
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(X$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        CLEA% = CVI(Mid$(X$, 45, 2))
        If CLEA% < 24 Then CLEA% = 24
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(X$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        Z$ = Space$(96)
        SUMAVA# = SUMAVA# + IMVA#
        IMPTEX$ = CSTRING$(MKD$(IMVA#), 3, 14, 2, 2)
        Call REPLA(Z$, CSTRING$(MKI$(MECO%), 1, 4, 0, 2), 1, 4)
        Call REPLA(Z$, BANCO$, 7, 14)
        Call REPLA(Z$, NUVA$, 23, 12)
        Call REPLA(Z$, PROTER$, 38, 2)
        Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
        Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
        Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
        Call REPLA(Z$, IMPTEX$, 70, 14)
        '
        List1.AddItem Z$
        '
290 Next U&
    '
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    '
    Label14.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    Label14.Refresh
    ANTI# = SUMAVA# - CDbl(Val(Label18.Caption))
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Anticipos:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Sub SIGUERECI3()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        TIPOS$ = Left$(Z$, 12)
        NROL& = Val(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = Val(Mid$(Z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TIPOS$, 1, 16)
        Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TIPOS$ = Left$(Y$, 16)
            NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            '
            Z$ = Space$(128)
            Call REPLA(Z$, TIPOS$, 1, 14)
            Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 31, 14)
            Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, DOCU$, 97, 18)
            Call REPLA(Z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            List2.AddItem Z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    ANTI# = CDbl(Val(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Anticipos:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub GRARECISTA(OKEY%)
    '
700 OKEY% = 0
    '
    NC% = Val(Text1.TEXT)
    NROO& = Val(Mid$(TRIM$(Text11.TEXT), 4))
    FC% = FECHANUM(TRIM$(Text9.TEXT))
    TOTAVAL# = CDbl(Val(Label14.Caption))
    ANTI# = CDbl(Val(Label20.Caption))
    URE202& = ULTREG&("SVD202")
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = Val(Label22.Caption)
    On Error GoTo 0
    '
    Call BLOQARCH("COBRAN")
    Call BLOQARCH("CAJABANC")
    Call BLOQARCH("APLICOB")
    '
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    '
    NROPRES& = Val(Mid$(Text11.TEXT, 4))
    NROPRO& = NROPRES&
    Call NURECIBO(NROPRO&)
    If NROPRO& <> NROPRES& Then
       Call CIERRARCH("COBRAN")
       Call CIERRARCH("CAJABANC")
       Call CIERRARCH("APLICOB")
       Screen.MousePointer = 1
       Call MENSERR(24, "Al Grabar este Recibo\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
       OKEY% = 0
       Exit Sub
    End If
    '
840 ' ***** aqui grabar cheques en CAJABANC.DAT
    '
    Z0$ = REGBLAN$("CAJABANC")
      NRECIB$ = TRIM$(Str$(COBRACT%))
      While Len(NRECIB$) < 2: NRECIB$ = "0" + NRECIB$: Wend
      NRECIB$ = TRIM$(Str$(NROO&)) + "-" + NRECIB$
      While Len(NRECIB$) < 9: NRECIB$ = "0" + NRECIB$: Wend
      NRECIB$ = "RB-" + NRECIB$
    Call REPLA(Z0$, NRECIB$, 1, 12)
    RASOC$ = TRIM$(RASOCLI$(NC%))
    Call REPLA(Z0$, RASOC$, 13, 48)
    '
    For U& = 1 To ULTREG&("!CARMECO")
      X$ = REGLEIDO$("!CARMECO", U&)
      TIPO% = CVI(Left$(X$, 2))
      BANCO$ = TRIM$(Mid$(X$, 3, 16))
      If BANCO$ = "" Then
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(TIPO%)))
      End If
      If Left$(UCase$(BANCO$), 3) = "CH." Then
        If Len(BANCO$) > 3 Then
          BANCO$ = Mid$(BANCO$, 4)
        End If
      End If
      NUCHE$ = TRIM$(Mid$(X$, 19, 12))
      IMCHE# = CVD(Mid$(X$, 31, 8))
      FECHE% = CVI(Mid$(X$, 39, 2))
      FEVEN% = CVI(Mid$(X$, 41, 2))
      '
      CUECARTE% = CUEMECOB%(-1)
      If IMCHE# > 0 Then
        '
        Z$ = Z0$
          FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          IDEVA$ = AJUSTI$(BANCO$, 18)
          IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
          IDEVA$ = IDEVA$ + FEVECHE$
        Call REPLA(Z$, MKI$(FC%), 61, 2)
           CUECO% = CUEMECOB%(TIPO%)
        Call REPLA(Z$, MKI$(CUECO%), 65, 2)
        Call REPLA(Z$, MKI$(NC%), 67, 2)
        Call REPLA(Z$, "CH", 69, 2)
        Call REPLA(Z$, IDEVA$, 71, 40)
          If FECHE% < 1 Then FECHE% = FC%
        Call REPLA(Z$, MKI$(FECHE%), 111, 2)
          If FEVEN% < FECHE% Then FEVEN% = FECHE%
          If FEVEN% < FC% Then FEVEN% = FC%
        Call REPLA(Z$, MKI$(FEVEN%), 63, 2)
        Call REPLA(Z$, MKI$(COBRACT%), 115, 2)
        Call REPLA(Z$, MKS$(NROO&), 117, 4)
           STATUI% = 0
           If CUECO% = CUECARTE% Then STATUI% = 1
        Call REPLA(Z$, Chr$(STATUI%), 121, 1) ' STATUS%=1 ES CHEQUE EN CARTERA
        Call REPLA(Z$, MKD$(IMCHE#), 145, 8)
        Call REPLA(Z$, MKD$(IMCHE#), 153, 8)
        Call REPLA(Z$, RASOC$, 177, 32)
        If TRIM$(NUCHE$) <> "" Then
          Call REPLA(Z$, TRIM$(BANCO$) + " " + TRIM$(NUCHE$) + " " + TRIM$(FEVECHE$), 177, 32)
        End If
        Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
        '
        Call REGAPP("CAJABANC", Z$)
        '
      End If
    Next U&
    Call CIERRARCH("CAJABANC")
    OKEY% = 6
    '
    URE202& = ULTREG&("SVD202")
    '
    If ANTI# >= 0.005 Then
        '
        Z$ = REGBLAN$("SVD202")
        Call REPLA(Z$, MKI$(NC%), 1, 2)
        Call REPLA(Z$, Chr$(4), 3, 1)
        Call REPLA(Z$, Chr$(0), 4, 1)
        Call REPLA(Z$, MKS$(NROO&), 5, 4)
        Call REPLA(Z$, MKI$(FC%), 9, 2)
        Call REPLA(Z$, MKD$(ANTI#), 11, 8)
        Call REPLA(Z$, MKI$(FC%), 19, 2)
        Call REPLA(Z$, MKD$(TOTAVAL#), 21, 8)
        Call REPLA(Z$, String$(2, 255), 31, 2)
        '
        Call BLOQARCH("SVD202")
        Call REGAPP("SVD202", Z$)
        Call CIERRARCH("SVD202")
        '
    End If
    OKEY% = 1
    '
800 ' *********************************** aqui blanquear facturas canceladas
    '
    XX% = 0
    VACON% = 1
    List2.Visible = False
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    For U& = 1 To List2.ListCount
      'List2.ListIndex = U& - 1
      X$ = List2.List(U& - 1)
      CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
      If CANCE# >= 0.005 Then
        NREG& = Val(Mid$(X$, 121, 6))
        If NREG& > 0 Then
          If NREG& <= ULTREG&("SVD202") Then
            '
            SS$ = REGLEIDO$("SVD202", NREG&)
            SALDEU# = CVD(Mid$(SS$, 11, 8)) - CANCE#
            Call REPLA(SS$, MKD$(SALDEU#), 11, 8)
            Call GRAREG("SVD202", SS$, NREG&)
            '
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("SVD202")
    OKEY% = 2
    '
750 RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
        DX$ = REGLEIDO$("DEUDOR2", RGX&)
        SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTAVAL#
        Call REPLA(DX$, MKD$(SALDO#), 109, 8)
        Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
    End If
    Call CIERRARCH("DEUDOR2")
    OKEY% = 3
    '
    URECABA& = ULTREG&("CAJABANC")
    Z$ = REGBLAN$("COBRAN")
    Call REPLA(Z$, MKI$(NC%), 1, 2)
    Call REPLA(Z$, MKS$(NROO&), 3, 4)
    Call REPLA(Z$, MKI$(FC%), 7, 2)
    Call REPLA(Z$, String$(2, 0), 9, 2)
    Call REPLA(Z$, MKS$(1 + URECABA&), 11, 4)
    Call REPLA(Z$, MKD$(TOTAVAL#), 15, 8)
    Call REPLA(Z$, MKI$(COBRACT%), 23, 2)
    '
    Call REGAPP("COBRAN", Z$)
    Call CIERRARCH("COBRAN")
    UCOBRA% = COBRACT%
    UNUREC& = NROO&
    URECOBRA& = ULTREG&("COBRAN")
    '
    Z$ = REGBLAN$("CAJABANC")
    Call REPLA(Z$, NRECIB$, 1, 12)
    Call REPLA(Z$, RASOCLI$(NC%), 13, 48)
    Call REPLA(Z$, MKI$(FC%), 61, 2)
    Call REPLA(Z$, MKI$(FC%), 63, 2)
    Call REPLA(Z$, MKI$(CUECOCLI%(NC%)), 65, 2)
    Call REPLA(Z$, MKI$(NC%), 67, 2)
    Call REPLA(Z$, "RB", 69, 2)
    Call REPLA(Z$, "Cobranza Cta.Corriente " + NRECIB$, 71, 40)
    Call REPLA(Z$, MKI$(FC%), 111, 2)
    Call REPLA(Z$, MKI$(COBRACT%), 115, 2)
    Call REPLA(Z$, MKS$(NROO&), 117, 4)
    Call REPLA(Z$, Chr$(0), 121, 1)
    Call REPLA(Z$, MKD$(TOTAVAL#), 145, 8)
    Call REPLA(Z$, MKD$(TOTAVAL#), 161, 8)
    Call REPLA(Z$, RASOC$, 177, 32)
    Call REPLA(Z$, "COBRAN", 245, 8)
    Call REPLA(Z$, MKS$(URECOBRA&), 253, 4)
    Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
' AQUI GRABAR CAJABANC PARA COBRANZA
    Call REGAPP("CAJABANC", Z$)
    OKEY% = 4
    '
    DESCAJ# = 0
    ANTI# = CDbl(Val(Label20.Caption))
    If ANTI# < 0 Then
        DESCAJ# = Abs(ANTI#)
        ANTI# = 0
    End If
    '
    If DESCAJ# >= 0.005 Then
'aun no se permiten descuentos de caja
    TIPO% = 5
     ' grabar como tipo 5 en 'FACTUR.DAT'
    End If
    '
    OKEY% = 5

830 ' **************** aqui grabar archivo APLICOB.DAT
    '
    For U& = 1 To List2.ListCount
        'List2.ListIndex = U& - 1
        X$ = List2.List(U& - 1)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = Val(Mid$(X$, 121, 6))
            If NREG& > 0 Then
                If NREG& <= URE202& Then
                    SS$ = REGLEIDO$("SVD202", NREG&)
                    NUSUCUVE% = CVI(Mid$(SS$, 29, 2))
                    DOCORI$ = FORMDOC$(CVS(Mid$(SS$, 5, 4)), Asc(Mid$(SS$, 3, 1)), Asc(Mid$(SS$, 4, 1)), NUSUCUVE%)
                    ITOTORI# = CVD(Mid$(SS$, 21, 8))
                    IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
                    '
                    REFAPLI$ = "Recibo " + TRIM$(Str$(NROO&))
                    Call APLIPAG(FC%, NC%, 4, 1, SUC%, NROO&, CANCE#, REFAPLI$, DOCORI$, ITOTORI#, IVAORI#)
                    '
                End If
            End If
        End If
    Next U&
    VACON% = 0
    Call CIERRARCH("APLICOB")
    OKEY% = 6
    '
890 Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
End Sub
'
Sub PRIRECISTA()
    '
    Dim ORCOP$(6)
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
    '    Exit Sub
    'End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!CARMECO")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("RECISTA", "AVAHOJA"))
    CANCELPRINT% = 0
    If CONTROL$("RECISTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 19
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        FEMII% = CVI(Mid$(X$, 39, 2))
        FEVEI% = CVI(Mid$(X$, 41, 2))
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        CLEA% = CVI(Mid$(X$, 45, 2))
        '
        TIPOSS$ = Space$(4)
        If MECO% > 0 Then
            For V& = 1 To ULTREG&("MECOBRA")
                MCB$ = REGLEIDO$("MECOBRA", V&)
                If CVI(Left$(MCB$, 2)) = MECO% Then
                    TIPOSS$ = Mid$(MCB$, 57, 4)
                    GoTo 10
                End If
            Next V&
        End If
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOSS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
19  Next U&
    '
    V0& = ULTREG&("!PRINRECI")
    '
    VACON% = 1
    List2.Visible = False
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        X$ = List2.TEXT
        DOCU$ = Mid$(X$, 97, 18)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    VACON% = 0
    '
    V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(Val(Label14.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption)) + CDbl(Val(Label20.Caption))
    ITEXTO$ = NUMTEX$(TOTAVAL#)
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + TRIM$(NRX3$) + " ctvs. *****************************************************************************************"
    '
    ANTI# = CDbl(Val(Label20.Caption))
    '
    If Abs(ANTI#) >= 0.005 Then
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, Space$(18), 53, 18)
        Call REPLA(Z$, MKD$(0), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
        V0& = ULTREG&("!PRINRECI")
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, "Anticipo N/Aplic.", 53, 18)
        If ANTI# < 0 Then
            Call REPLA(Z$, "Descuento de Caja", 53, 18)
        End If
        Call REPLA(Z$, MKD$(ANTI#), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORECI$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            '
            YAC = DESMM
            '
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                        'Printer.EndDoc
                        Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Recibo Nro. " + NRORECI$ + " - " + TRIM$(RASOCLI$(NC%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Recibo Numero\'" + NRORECI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Recibo") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = PORIAN
            Call SETSPOOL("ORIENTATION", Str$(PORIAN))
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Sub APLIPAG(FE%, NC%, TI%, VA%, PRF%, NROF&, IMAPLI#, REFE$, DOCORI$, ITOTORI#, IVAORI#)
    '
    ARAPLI$ = "APLICOB": If NC% < 0 Then ARAPLI$ = "APLIPAG"
    PRFS$ = TRIM$(Str$(PRF%)): While Len(PRFS$) < 4: PRFS$ = "0" + PRFS$: Wend
    NROS$ = TRIM$(Str$(NROF&)): While Len(NROS$) < 8: NROS$ = "0" + NROS$: Wend
    '
    TIX% = TI% Mod 10
    TINUVA$ = "FC"                                               ' por default
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 3 Then TINUVA$ = "ND"
        If TIX% = 4 Then
        TINUVA$ = "RB"
        If NC% < 0 Then
            TINUVA$ = "OP"
            If TI% = 14 Then TINUVA$ = "FF"
        End If
    End If
    '
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    VAD$ = Chr$(64 + VA%): If VAD$ < "A" Or VAD$ > "Z" Then VAD$ = "X"
    TINUVA$ = TINUVA$ + "-" + VAD$ + "-" + PRFS$ + "-" + NROS$
    '
    RECORD$ = MKI$(FE%)
    RECORD$ = RECORD$ + MKI$(Abs(NC%))
    '
    RECORD$ = RECORD$ + TINUVA$
    RECORD$ = RECORD$ + MKD$(IMAPLI#)
    '
    REFE$ = Left$(REFE$ + Space$(24), 24)
    RECORD$ = RECORD$ + REFE$
    '
    RECORD$ = RECORD$ + DOCORI$ + MKD$(ITOTORI#) + MKD$(IVAORI#)
    '
    RECORD$ = RECORD$ + String$(8, 0)
    '
    Call REGAPP(ARAPLI$, RECORD$)
    'Call CIERRARCH(ARAPLI$)
    '
End Sub
'



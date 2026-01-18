VERSION 5.00
Begin VB.Form OPRECIB00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Recibo Standard de Cobranza"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10770
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   4550
      Left            =   9720
      ScaleHeight     =   4485
      ScaleWidth      =   1020
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
      Begin VB.Image Image3 
         Height          =   4550
         Left            =   0
         Picture         =   "OPRECIB00.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Contabilizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   5460
      TabIndex        =   40
      Top             =   315
      Width           =   4180
      Begin VB.CommandButton Command19 
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
         Left            =   3780
         TabIndex        =   76
         Top             =   2220
         Width           =   175
      End
      Begin VB.CommandButton Command18 
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
         Left            =   3780
         TabIndex        =   70
         Top             =   1560
         Width           =   175
      End
      Begin VB.CommandButton Command15 
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
         Left            =   3780
         TabIndex        =   53
         Top             =   900
         Width           =   175
      End
      Begin VB.CommandButton Command14 
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
         Left            =   3780
         TabIndex        =   50
         Top             =   3580
         Width           =   175
      End
      Begin VB.CommandButton Command13 
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
         Left            =   3780
         TabIndex        =   46
         Top             =   2880
         Width           =   175
      End
      Begin VB.CommandButton Command12 
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
         Left            =   3780
         TabIndex        =   42
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste Redondeo:"
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
         Left            =   -120
         TabIndex        =   79
         Top             =   2280
         Width           =   2325
      End
      Begin VB.Label Label20 
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
         Left            =   2310
         TabIndex        =   78
         Top             =   2220
         Width           =   1485
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   2550
         Width           =   3690
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   270
         TabIndex        =   73
         Top             =   1900
         Width           =   3690
      End
      Begin VB.Label Label17 
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
         Left            =   2310
         TabIndex        =   72
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N/Cred.Descuentos:"
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
         Left            =   -105
         TabIndex        =   71
         Top             =   1680
         Width           =   2325
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acreedores Comerciales:"
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
         Left            =   -120
         TabIndex        =   56
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label Label14 
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
         Left            =   2310
         TabIndex        =   55
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   270
         TabIndex        =   54
         Top             =   1250
         Width           =   3690
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   270
         TabIndex        =   52
         Top             =   3920
         Width           =   3690
      End
      Begin VB.Label Label11 
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
         Left            =   2310
         TabIndex        =   51
         Top             =   3580
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques Rechazados:"
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
         Left            =   -105
         TabIndex        =   49
         Top             =   3650
         Width           =   2325
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   270
         TabIndex        =   48
         Top             =   3260
         Width           =   3690
      End
      Begin VB.Label Label8 
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
         Left            =   2310
         TabIndex        =   47
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartera de Valores:"
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
         Left            =   -105
         TabIndex        =   45
         Top             =   2950
         Width           =   2325
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   270
         TabIndex        =   44
         Top             =   600
         Width           =   3690
      End
      Begin VB.Label Label5 
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
         Left            =   2310
         TabIndex        =   43
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deudores por Ventas:"
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
         Left            =   -105
         TabIndex        =   41
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tablas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   9770
      TabIndex        =   36
      Top             =   2130
      Width           =   860
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
         Height          =   450
         Left            =   80
         Picture         =   "OPRECIB00.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   1710
         Width           =   670
      End
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
         Height          =   450
         Left            =   80
         Picture         =   "OPRECIB00.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Archivo Maestro de Cuentas Contables"
         Top             =   270
         Width           =   670
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
         Height          =   450
         Left            =   80
         Picture         =   "OPRECIB00.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Actualizar Tabla de Talonarios de Recibo y Cobradores"
         Top             =   1230
         Width           =   670
      End
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
         Height          =   450
         Left            =   80
         Picture         =   "OPRECIB00.frx":2936
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Actualizar Tabla de Valores y Medios de Cobranza"
         Top             =   750
         Width           =   670
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comprobantes Impresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5460
      TabIndex        =   30
      Top             =   4635
      Width           =   5160
      Begin VB.CommandButton Command17 
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
         Left            =   4200
         Picture         =   "OPRECIB00.frx":2A80
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1050
         Width           =   750
      End
      Begin VB.CommandButton Command16 
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
         Left            =   3780
         TabIndex        =   63
         Top             =   960
         Width           =   175
      End
      Begin VB.TextBox Text10 
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
         Left            =   2625
         TabIndex        =   64
         Text            =   " "
         Top             =   980
         Width           =   1170
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
         Height          =   285
         Left            =   210
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1330
         Width           =   3780
      End
      Begin VB.CommandButton Command9 
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
         Left            =   4200
         Picture         =   "OPRECIB00.frx":2D8A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   315
         Width           =   750
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
         Height          =   285
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   640
         Width           =   3780
      End
      Begin VB.CommandButton Command8 
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
         Left            =   3780
         TabIndex        =   32
         Top             =   300
         Width           =   175
      End
      Begin VB.TextBox Text8 
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
         Left            =   2625
         TabIndex        =   31
         Text            =   " "
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comprensación de Cuentas:"
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
         Index           =   1
         Left            =   -420
         TabIndex        =   65
         Top             =   1070
         Width           =   2955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste de Saldo:"
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
         Index           =   0
         Left            =   150
         TabIndex        =   35
         Top             =   380
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recibo de Cobranza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   5160
      Begin VB.Frame Frame14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1900
         TabIndex        =   67
         Top             =   2100
         Width           =   2130
         Begin VB.OptionButton Option10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Cancela"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   945
            TabIndex        =   75
            Top             =   490
            Width           =   1065
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   74
            Top             =   490
            Width           =   1275
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "N/Credito"
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
            Left            =   105
            TabIndex        =   69
            Top             =   210
            Width           =   1155
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Sin Impresión"
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
            Left            =   105
            TabIndex        =   68
            Top             =   840
            Width           =   1590
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reng.Det."
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
         Left            =   3675
         TabIndex        =   59
         Top             =   4095
         Width           =   1275
         Begin VB.TextBox Text6 
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
            Left            =   420
            MaxLength       =   2
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   260
            Width           =   450
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   58
         Text            =   " "
         Top             =   1545
         Width           =   3165
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   29
         Top             =   1260
         Width           =   4740
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenamiento"
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
         Left            =   210
         TabIndex        =   26
         Top             =   4875
         Width           =   1650
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Vencim."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   28
            Top             =   600
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Emision"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prefijo"
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
         Left            =   3675
         TabIndex        =   18
         Top             =   3360
         Width           =   1275
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
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
            Left            =   315
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   260
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   210
         TabIndex        =   17
         Top             =   3360
         Width           =   3330
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   480
            Left            =   315
            TabIndex        =   23
            Top             =   735
            Width           =   2850
            Begin VB.OptionButton Option6 
               BackColor       =   &H00C0E0FF&
               Caption         =   "NewPage"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   25
               Top             =   120
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00C0E0FF&
               Caption         =   "EndDoc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   24
               Top             =   120
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   435
            Left            =   315
            TabIndex        =   20
            Top             =   315
            Width           =   2850
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Formulario Continuo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Hoja Suelta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   21
               Top             =   0
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Talonario en Uso"
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
         Left            =   1995
         TabIndex        =   13
         Top             =   4875
         Width           =   2955
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
            Height          =   360
            Left            =   2580
            TabIndex        =   15
            Top             =   420
            Width           =   175
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
            Height          =   360
            Left            =   1425
            TabIndex        =   14
            Text            =   " "
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
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
            Left            =   165
            TabIndex        =   16
            Top             =   375
            Width           =   1170
         End
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
         Height          =   360
         Left            =   2415
         TabIndex        =   12
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   11
         Top             =   420
         Width           =   175
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
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   4095
         Picture         =   "OPRECIB00.frx":3094
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   210
         TabIndex        =   5
         Top             =   2100
         Width           =   1650
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pre-Impresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Auto-Numera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   6
            Top             =   630
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   4095
         TabIndex        =   3
         Top             =   2205
         Width           =   855
         Begin VB.TextBox Text5 
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
            Left            =   190
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "RB"
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Impresion:"
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
         Index           =   4
         Left            =   105
         TabIndex        =   57
         Top             =   1605
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5145
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         TabIndex        =   8
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   9765
      Picture         =   "OPRECIB00.frx":339E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   750
      Left            =   9765
      Picture         =   "OPRECIB00.frx":36A8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   400
      Width           =   855
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu cuentascontables 
         Caption         =   "Cuentas Contables"
      End
      Begin VB.Menu valoresymediosdecobranza 
         Caption         =   "Tabla de Valores y Medios de Cobranza"
      End
      Begin VB.Menu talonariosderecibosycobradores 
         Caption         =   "Tabla de Talonarios de Recibo y Cobradores"
      End
      Begin VB.Menu cotizacionytipodecambio 
         Caption         =   "Cotización Dolar y Tipo de Cambio "
      End
   End
End
Attribute VB_Name = "OPRECIB00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command12_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label5 = VDEVU$
        Label6 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command13_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CARTACOBRA
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   RETEJ% = EJER%
   EJER% = 1
   Call CARCUEINGAS
   EJER% = RETEJ%
   EJU$ = EJACT$
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command15_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label14 = VDEVU$
        Label13 = VALACT2$("ECUECON")
    End If
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Command16_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command17_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command18_Click()
    '\\\OT-07-0029-RG-26/01/07///
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label17 = VDEVU$
        Label18 = VALACT2$("ECUECON")
    End If
    '\\\OT-07-0029-RG-FIN///
End Sub

Private Sub Command19_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label20 = VDEVU$
        Label19 = VALACT2$("ECUECON")
    End If
End Sub

'\\\OT-06-0144-WAR-FIN///
Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
        FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text3.SetFocus
        Exit Sub
    End If
    Call GRACONTROL("", "FENURECI", FENUE$)
    '
    Call GRACONTROL("", "FORECIST", Text1.TEXT)
    Call GRACONTROL("", "FORAJUSA", Text8.TEXT)
    Call GRACONTROL("", "FORCOMPE", Text10.TEXT)
    Call GRACONTROL("", "CODMRECI", "RB")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NURECIST", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPARECI", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("RECISTA", "AVAHOJA", AVAHOJA$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM": MODORECI% = 2
    End If
    Call GRACONTROL("RECISTA", "MODORDE", MODORDE$)
    '
    DESCUEX$ = ""
    If Check3.Value = 1 Then DESCUEX$ = "CREDESCU"
    Call GRACONTROL("RECISTA", "DESCAJA", DESCUEX$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check1.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("CREDESCU", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("RECISTA", "SUPRINT", SUPRINT$)
    '
    PorValor$ = "SI": If Option10.Value = True Then PorValor$ = "NO"
    Call GRACONTROL("", "OPCVALOR", PorValor$)
    '
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "" Then
       If Len(PREFIREM$) <> 4 Then
          Call COMUNI("Prefijo no Válido")
          Text4.SetFocus
          Exit Sub
       End If
       For U& = 1 To 4
          UI% = Asc(Mid$(PREFIREM$, U&))
          If UI% < 48 Or UI% > 57 Then
             Call COMUNI("Prefijo no Válido")
             Text4.SetFocus
             Exit Sub
          End If
       Next U&
    End If
    Call GRACONTROL("", "PREFRECI", PREFIREM$)
    '
    Call GRACONTROL("AMACLI", "ECUECON", Label5)
    Call GRACONTROL("AMAPRO", "ECUECON", Label14)
    Call GRACONTROL("RECISTA", "CUEREDON", Label20)
     
    Call GRACONTROL("", "CUECHERE", Label11)
    Call GRACONTROL("", "CUEDESCU", Label17)
    '
    CARENDET% = Val(Text6.TEXT)
    Call GRACONTROL("RECISTA", "CARENDET", Text6)
    Call GRACONTROL(IDENTER$, "PORTREC", Combo1.TEXT)
    Call GRACONTROL("*", "PORTREC", Combo1.TEXT)
    Call CIERRARCH("*.*")
    '
    On Error Resume Next
    RECIBO00.List1.Clear
    RECIBO00.List2.Clear
    RECIBO00.Label4.Caption = ""
    RECIBO00.Label8.Caption = ""
    RECIBO00.Label10.Caption = ""
    RECIBO00.Label14.Caption = ""
    RECIBO00.Label18.Caption = ""
    RECIBO00.Label20.Caption = ""
    RECIBO00.BOTNEXT.Enabled = True
    RECIBO00.Command5.Enabled = False
    On Error GoTo 0
    '
    ' SE CAMBIA DE LUGAR PORQUE SINO EL CALL FINAL NO PERMITIRIA
    Call VERICAMCAR(Label8.Caption)
    Call FINAL("")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
       Call FINAL("")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CARTAMECO
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command9_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub cotizacionytipodecambio_Click()
   Call Command7_Click
End Sub

Private Sub cuentascontables_Click()
   Call Command11_Click
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      Me.Frame13.BackColor = &H8000000F
      Me.Frame14.BackColor = &H8000000F
      
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      Me.Option9.BackColor = &H8000000F
      Me.Option10.BackColor = &H8000000F
      
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
    End If

    Text1.TEXT = CONTROL$("", "FORECIST")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FENURECI")
    Text8.TEXT = CONTROL$("", "FORAJUSA")
    Text9.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    '\\\OT-06-0144-WAR-17/04/06///
    Text10.TEXT = CONTROL$("", "FORCOMPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    '\\\OT-06-0144-WAR-FIN///
    '
    '\\\OT-05-0513-WAR-14/07/05///
    Text6 = UCase$(CONTROL$("RECISTA", "CARENDET"))
    '\\\OT-05-0513-WAR-FIN
    Label5 = CONTROL$("AMACLI", "ECUECON")
    Label6 = ECOARCH$("ECUECON", Label5)
    Label14 = CONTROL$("AMAPRO", "ECUECON")
    Label13 = ECOARCH$("ECUECON", Label14)
    Label20 = CONTROL$("RECISTA", "CUEREDON")
    Label19 = ECOARCH$("ECUECON", Label20)
    Label8 = CONTROL$("", "CUECHECA")
    Label9 = ECOARCH$("ECUECON", Label8)
    Label11 = CONTROL$("", "CUECHERE")
    Label12 = ECOARCH$("ECUECON", Label11)
    '\\\OT-07-0029-RG-26/01/07///
    Label17 = CONTROL$("", "CUEDESCU")
    Label18 = ECOARCH$("ECUECON", Label17)
    '\\\OT-07-0029-RG-FIN///
    '
    If CONTROL$("", "NURECIST") = "AUTONUME" Then Option2.Value = True
    If CONTROL$("", "TIPARECI") = "HOJASUEL" Then Option4.Value = True
    If CONTROL$("RECISTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    '\\\OT-07-0029-RG-26/01/07///
    If CONTROL$("RECISTA", "DESCAJA") = "CREDESCU" Then Check3.Value = 1
    If CONTROL$("CREDESCU", "SUPRINT") = "SI" Then Check1.Value = 1
    '\\\OT-07-0029-RG-FIN///
    If CONTROL$("RECISTA", "SUPRINT") = "SI" Then Check2.Value = 1
    MODORECI% = 1: If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    Option9.Value = True
    If CONTROL$("", "OPCVALOR") = "NO" Then Option10.Value = True
    '
    PREFIREM$ = TRIM$(CONTROL$("", "PREFRECI"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "RB"
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = TRIM$(CONTROL$(IDENTER$, "PORTREC"))
      If Combo1.TEXT = "" Then Combo1.TEXT = CONTROL$("*", "PORTREC")
      If Combo1.TEXT = "" Then Combo1.TEXT = Printer.DeviceName
      '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label14_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label14 = VDEVU$
        Label13 = VALACT2$("ECUECON")
    End If
End Sub



Private Sub Label17_DblClick()
    '\\\OT-07-0029-RG-26/01/07///
    Call Command18_Click
    '\\\OT-07-0029-RG-FIN///
End Sub

Private Sub Label5_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label5 = VDEVU$
        Label6 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label8_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub talonariosderecibosycobradores_Click()
   Call Command10_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'\\\OT-06-0144-WAR-FIN///
Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text5_Click()
    Command2.SetFocus
End Sub
'\\\OT-06-0144-WAR-17/04/06///
Private Sub Text7_CLICK()
  Text10.SetFocus
End Sub
'\\\OT-06-0144-WAR-FIN///
Private Sub Text8_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text9_Click()
    Text8.SetFocus
End Sub

Sub CARTAMECO()     ' carga tabla medios de cobranza
    '
    If DERACCE%("ACTAMECO", "Actualizacion Tabla de Medios de Cobranza") < 2 Then
      Exit Sub
    End If
    '
    CUECHECA$ = TRIM$(CONTROL$("", "CUECHECA"))
    If CUECHECA$ = "" Or CUECOINT%(CUECHECA$) < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores.")
      Exit Sub
    End If
    '
    Call SETULTREG("!MECOBRA", 0)
    '
    LU$ = LUDAT$
    If EXISTE%(LU$ + "MECOBRA.DAT") = 1 Then
       If COPYFILE%(LU$ + "MECOBRA.DAT", "C:\FLEXOFT\MECOBRA.WKF") <> 1 Then
          Command1.Enabled = True
          Exit Sub
       End If
    End If
    '
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       X$ = REGLEIDO$("!MECOBRA", U&)
       If CVI(Left$(X$, 2)) < 900 Then
         CUECO$ = TRIM$(Mid$(X$, 45, 12))
         If CUECO$ = "" Or CUECOINT%(CUECO$) < 1 Then
           Call REPLA(X$, CUECHECA$, 45, 12)
           Call GRAREG("!MECOBRA", X$, U&)
         End If
       End If
    Next U&
    Call CIERRARCH("!MECOBRA")
    Screen.MousePointer = 1
    '
360 ATY% = VENTABU%("COMECOB/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       X$ = REGLEIDO$("!MECOBRA", U&)
       CDX% = CVI(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             J& = J& + 1
363          Call GRAREG("!MECOBRA", X$, J&)
             If TRIM$(Mid$(X$, 3, 12)) = "" Then ATX% = 1
             If TRIM$(Mid$(X$, 45, 12)) = "" Then ATX% = 1
364          If TRIM$(Mid$(X$, 57, 4)) = "" Then ATX% = 1
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MECOBRA", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        Call COMUNI("Datos Incompletos !!!")
        GoTo 360
    End If
    '
    Call TRACENTRAL("MECOBRA", "")
    Call CIERRARCH("MECOBRA")
    Call FILESORT("MECOBRA", "", 1, 1, "ASC")
    '
390 End Sub

Sub CARTACOBRA()    ' carga tabla cobradores / talonarios
    '
    Dim MCOBRA%(1000)
    '
    Call CIERRARCH("COBRADOR")
    Call CIERRARCH("!COBRADOR")
    LU$ = LUDAT$
    If EXISTE%(LU$ + "COBRADOR.DAT") = 1 Then
       If COPYFILE%(LU$ + "COBRADOR.DAT", "C:\FLEXOFT\COBRADOR.WKF") <> 1 Then
          Command6.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COBRADOR/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!COBRADOR")
       X$ = REGLEIDO$("!COBRADOR", U&)
       CDX% = Asc(Left$(X$, 2))
362    If CDX% >= 0 Then
          If CDX% <= 255 Then
             If MCOBRA%(CDX%) = 0 Then
                J& = J& + 1
363             Call GRAREG("!COBRADOR", X$, J&)
                MCOBRA%(CDX%) = 1
             End If
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!COBRADOR", J&)
    Call CIERRARCH("!COBRADOR")
    Call CIERRARCH("COBRADOR")
    '
366 Screen.MousePointer = 1
    '
    AA% = COPYFILE%("C:\FLEXOFT\COBRADOR.WKF", LU$ + "COBRADOR.DAT")
    '
390 End Sub


Sub VERICAMCAR(CTANUE$)
   '
   CTAANT$ = CONTROL$("", "CUECHECA")
   If CTAANT$ <> CTANUE$ Then
     If DERACCE%("CAMCUECA", "Modificar Cuenta Cartera de Valores") > 1 Then
       OPX% = COMALTER%("El Cambio de la Cuenta Asociada a la Cartera\Puede Alterar el Funcionamiento del Subsistema.\\¿ Está Seguro de Cambiarla ?\\No, Cancelar el Cambio\Si, Registrar Nueva Cuenta")
       If OPX% = 2 Then
         Call GRACONTROL("", "CUECHECA", CTANUE$)
         Call MENSERR(24, "Se Ha Modificado la Cuenta 'Cartera de Valores'\de la Cuenta Nro " + TRIM$(CTAANT$) + " " + TRIM$(ECOARCH$("ECUECON", CTAANT$)) + "\por la Cuenta Nro " + TRIM$(CTANUE$) + " " + TRIM$(ECOARCH$("ECUECON", CTANUE$)) + "\  \Re-Inicie Ahora la Aplicación.")
       End If
     End If
   End If
   '
End Sub

Private Sub valoresymediosdecobranza_Click()
    Call Command6_Click
End Sub

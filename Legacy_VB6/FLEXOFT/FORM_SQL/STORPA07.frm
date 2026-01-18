VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STORPA07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formularios de Comprobantes de Pago Standard"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   1515
   ClientWidth     =   11535
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
   ScaleHeight     =   6120
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1800
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   47
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   7470
      Left            =   10605
      ScaleHeight     =   7440
      ScaleWidth      =   1260
      TabIndex        =   29
      Top             =   0
      Width           =   1290
      Begin VB.CommandButton Command14 
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
         Picture         =   "STORPA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   3250
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Tbls"
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
         Picture         =   "STORPA07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   2560
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
         Picture         =   "STORPA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   250
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
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
         MaskColor       =   &H00404000&
         Picture         =   "STORPA07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   940
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Picture         =   "STORPA07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   1750
         Width           =   650
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C7D9CC&
      Caption         =   "FORMULARIOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4950
      Begin VB.CommandButton Command22 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   4200
         Width           =   645
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   79
         Text            =   " "
         Top             =   4200
         Width           =   1380
      End
      Begin VB.CommandButton Command21 
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
         Left            =   3810
         TabIndex        =   78
         Top             =   4200
         Width           =   175
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   " "
         Top             =   4560
         Width           =   3750
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   76
         Top             =   4245
         Width           =   645
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   70
         Top             =   3520
         Width           =   645
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3840
         Width           =   3750
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
         Left            =   3810
         TabIndex        =   67
         Top             =   3480
         Width           =   175
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   66
         Text            =   " "
         Top             =   3480
         Width           =   1380
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2760
         Width           =   645
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   56
         Text            =   " "
         Top             =   2760
         Width           =   1380
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
         Left            =   3810
         TabIndex        =   55
         Top             =   2760
         Width           =   175
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3090
         Width           =   3750
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "F.Fijo - Acumular el pago en un solo movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   53
         Top             =   1680
         Width           =   4650
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C7D9CC&
         Caption         =   "CÁLCULO RETENCIÓN IIBB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   240
         TabIndex        =   42
         Top             =   4920
         Width           =   2895
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Acumulación Mensual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   420
            TabIndex        =   43
            Top             =   300
            Width           =   2205
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C7D9CC&
         Caption         =   "NUMERACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3240
         TabIndex        =   38
         Top             =   4920
         Width           =   1485
         Begin VB.OptionButton Option10 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   40
            Top             =   420
            Width           =   1260
         End
         Begin VB.OptionButton Option9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   39
            Top             =   210
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2385
         Width           =   3750
      End
      Begin VB.CommandButton Command10 
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
         TabIndex        =   36
         Top             =   2055
         Width           =   175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   35
         Text            =   " "
         Top             =   2055
         Width           =   1380
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2055
         Width           =   645
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":17DA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1035
         Width           =   645
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   210
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1370
         Width           =   3750
      End
      Begin VB.CommandButton Command6 
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
         TabIndex        =   23
         Top             =   1035
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   22
         Text            =   " "
         Top             =   1035
         Width           =   1380
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         Picture         =   "STORPA07.frx":1AE4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   650
         Width           =   3750
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
         Height          =   285
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB Tucumán"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   81
         Top             =   4290
         Width           =   2085
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB Misiones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   69
         Top             =   3570
         Width           =   2085
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB CABA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   58
         Top             =   2850
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. IIBB Bs As"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   34
         Top             =   2140
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5145
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación Fondo Fijo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   26
         Top             =   1110
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   21
         Top             =   380
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CONFIGURACION DE ORDENES DE PAGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   5250
      TabIndex        =   0
      Top             =   120
      Width           =   5160
      Begin VB.CommandButton Command20 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4215
         Picture         =   "STORPA07.frx":1DEE
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4560
         Width           =   645
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
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
         Left            =   2535
         TabIndex        =   73
         Text            =   " "
         Top             =   4560
         Width           =   1380
      End
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
         Left            =   3930
         TabIndex        =   72
         Top             =   4560
         Width           =   175
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   " "
         Top             =   4920
         Width           =   3750
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Suprime Impresión al Generar el Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         TabIndex        =   64
         Top             =   5350
         Width           =   4650
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   330
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   4170
         Width           =   3750
      End
      Begin VB.CommandButton Command11 
         Caption         =   "."
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
         Left            =   3900
         TabIndex        =   61
         Top             =   3840
         Width           =   175
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
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
         Left            =   2535
         TabIndex        =   60
         Text            =   " "
         Top             =   3840
         Width           =   1380
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         Picture         =   "STORPA07.frx":20F8
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3840
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C7D9CC&
         Caption         =   "NUMERACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   210
         TabIndex        =   3
         Top             =   1440
         Width           =   4725
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Pre-Impresa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Auto-Numera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            TabIndex        =   4
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C7D9CC&
         Caption         =   "OPCIONES VARIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         TabIndex        =   27
         Top             =   240
         Width           =   3750
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "No Aplica Ordenes de Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   360
            TabIndex        =   46
            Top             =   810
            Width           =   2805
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "No Aplica Créditos Pendientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   360
            TabIndex        =   45
            Top             =   540
            Width           =   3045
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Utiliza Cartera de Valores"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   360
            TabIndex        =   28
            Top             =   270
            Width           =   2805
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C7D9CC&
         Caption         =   "ORDENAMIENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   13
         Top             =   2820
         Width           =   1650
         Begin VB.OptionButton Option8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   15
            Top             =   525
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Emisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   14
            Top             =   315
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C7D9CC&
         Caption         =   "PREFIJO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3675
         TabIndex        =   11
         Top             =   2040
         Width           =   1275
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   240
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C7D9CC&
         Caption         =   "TALONARIO EN USO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1995
         TabIndex        =   6
         Top             =   2820
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
            TabIndex        =   8
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1425
            TabIndex        =   7
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   165
            TabIndex        =   9
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C7D9CC&
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
         Height          =   1185
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   855
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   190
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "OP"
            Top             =   480
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C7D9CC&
         Caption         =   "PAPEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   210
         TabIndex        =   10
         Top             =   2040
         Width           =   3330
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "EndDoc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   52
            Top             =   220
            Width           =   1275
         End
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "NewPage"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   1800
            TabIndex        =   51
            Top             =   220
            Width           =   1275
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "H.Suelta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   1800
            TabIndex        =   50
            Top             =   470
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Form. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   49
            Top             =   470
            Value           =   -1  'True
            Width           =   1515
         End
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención de Iva:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   75
         Top             =   4650
         Width           =   2085
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención de Ganancias:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   63
         Top             =   3915
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STORPA07.frx":2402
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "STORPA07.frx":247C
      Top             =   1200
   End
End
Attribute VB_Name = "STORPA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command11_Click()
  Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text10.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text11.TEXT), "Formularios de Impresion")
   Command12.Enabled = True
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
    '
    If DERACCE%("CARMEPAG", "Actualizar Tabla de Medios de Pago") < 2 Then
       Command13.Enabled = True
       Exit Sub
    End If
    '
    LU$ = LUDAT$
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    If COPYFILE%(LU$ + "MEPAGO.DAT", "C:\FLEXOFT\MEPAGO.WKF") <> 1 Then
     Command13.Enabled = True
        Exit Sub
    End If
    '
    Dim REMEPA&(1000)
    '
360 ATY% = VENTABU%("COMEPAG/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MEPAGO")
       x$ = REGLEIDO$("!MEPAGO", U&)
       CDX% = CVI(Left$(x$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             '
             If UCase$(Mid$(x$, 57, 4)) = "CH.3" Then
                 Call MENSERR(24, "Valor '" + Mid$(x$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             If UCase$(Mid$(x$, 57, 4)) = "PROP" Then
                 Call MENSERR(24, "Valor '" + Mid$(x$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             '
             J& = J& + 1
363          Call GRAREG("!MEPAGO", x$, J&)
             If TRIM$(Mid$(x$, 3, 12)) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Falta Descripción Abreviada.")
                 ATX% = 1
             End If
             If ECOARCH$("ECUECON", TRIM$(Mid$(x$, 45, 12))) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Cuenta Contable Inexistente o no Válida.")
                 ATX% = 1
             End If
             CCII% = CUEINT%(TRIM$(Mid$(x$, 45, 12)))
             Call REPLA(x$, MKI$(CCII%), 61, 2)
             If CCII% < 1 Or (NUCUEN% > 0 And CCII% > NUCUEN%) Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Cuenta Contable Inexistente o no Válida.")
                 ATX% = 1
             End If
364          If TRIM$(Mid$(x$, 57, 4)) = "" Then
                 Call MENSERR(24, "Registro " + TRIM$(Str$(U&)) + " - Falta Código de M.Pago.")
                 ATX% = 1
             End If
             '
             If REMEPA&(CDX%) < 1 Then REMEPA&(CDX%) = J&
             '
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MEPAGO", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        For UII% = 0 To 1000: REMEPA&(UII%) = 0: Next UII%
        Call COMUNI("Datos Incompletos o no Válidos.")
        GoTo 360
    End If
    '
    JJ& = 0
    For UII% = 1 To 1000
      If REMEPA&(UII%) > 0 Then
        If REMEPA&(UII%) <= J& Then
          x$ = REGLEIDO$("!MEPAGO", REMEPA&(UII%))
          JJ& = JJ& + 1
          Call GRAREG("MEPAGO", x$, JJ&)
        End If
      End If
    Next UII%
    Call SETULTREG("MEPAGO", JJ&)
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    Call FRESHECHO("MEPAGO")
    '
390 Command13.Enabled = True
    '
End Sub

Private Sub Command14_Click()
   '\\\OT-06-0023-FG-26/01/06
   Command14.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
99 Command14.Enabled = True
   '\\\OT-06-0023-FG-FIN
End Sub

Private Sub Command15_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text13.TEXT = VDEVU$
        Text12.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command16_Click()

   Command9.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text13.TEXT), "Formularios de Impresion")
   Command9.Enabled = True

End Sub

Private Sub Command17_Click()

   Command17.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text14.TEXT), "Formularios de Impresion")
   Command17.Enabled = True
   
End Sub

Private Sub Command18_Click()

    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text14.TEXT = VDEVU$
        Text15.TEXT = VALACT2$("INDIFRM")
    End If
    
End Sub

Private Sub Command19_Click()
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) <> "" Then
        Text17.TEXT = CODFOR$
    End If
End Sub

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
    '
    Call GRACONTROL("", "FORPASTA", Text1.TEXT)
    Call GRACONTROL("", "FORLIFOF", Text6.TEXT)
    Call GRACONTROL("", "FORRETIB", Text9.TEXT)
    Call GRACONTROL("", "FOREIBCA", Text13.TEXT)
    Call GRACONTROL("", "FOREIBMI", Text14.TEXT)
    If TRIM$(Text14) <> "" Then Check7.Value = 1
    VALO$ = "NO": If Check7.Value = 1 Then VALO$ = "SI"
    Call GRACONTROL("", "RETIBMI", VALO$)
    Call GRACONTROL("", "FOREIBTU", Text19.TEXT)
    If TRIM$(Text19) <> "" Then Check8.Value = 1
    VALO$ = "NO": If Check8.Value = 1 Then VALO$ = "SI"
    Call GRACONTROL("", "RETIBTU", VALO$)
    
    '****////******
    'ACA AGREGO EL CONTROL DEL FORMULARIO DE RET DE GANANCIAS
    Call GRACONTROL("", "FORREGAN", Text11.TEXT)
    '****////*******
    Call GRACONTROL("", "FORETIVA", Text17.TEXT)
    '
    Call GRACONTROL("", "CODMORPA", "OP")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NUORPAST", MONUME$)
    '
    MONUME$ = "AUTONUME": If Option10.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NURETGAN", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPAORPA", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("ORPASTA", "AVAHOJA", AVAHOJA$)
    '
    UTICHECA$ = "NO": If Check1.Value = 1 Then UTICHECA$ = "SI"
    Call GRACONTROL("ORPASTA", "UTICHECA", UTICHECA$)
    '
    SUPRICR$ = "NO": If Check4.Value = 1 Then SUPRICR$ = "SI"
    Call GRACONTROL("ORPASTA", "SUPRICRE", SUPRICR$)
    '
    SUPRIPA$ = "NO": If Check5.Value = 1 Then SUPRIPA$ = "SI"
    Call GRACONTROL("ORPASTA", "SUPRIPAG", SUPRIPA$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("ORPASTA", "SUPRINT", SUPRINT$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM"
        MODORECI% = 2
    End If
    Call GRACONTROL("ORPASTA", "MODORDE", MODORDE$)
    '
    Call GRACONTROL("", "FENUORPA", FENUE$)
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
    Call GRACONTROL("", "PREFORPA", PREFIREM$)
    '\\\OT-05-0319-WAR-23/05/05///
    ACUMENS$ = "NO": ACUMERIB% = 0
    If Check3.Value = 1 Then ACUMENS$ = "SI": ACUMERIB% = 1
    Call GRACONTROL("ORPASTA", "ACUMENS", ACUMENS$)
    '
    ' OBTIENE LA CONFIGURACION DE LA CANTIDAD DE MOVIMIENTOS QUE GENERAR EN LA CUENTA CONTABLE
    ' DEL MEDIO DE PAGO. UNO POR EL TOTAL DE LO CANCELADO, O UNO POR CADA CBTE. CANCELADO
    CAMOPAGO$ = "VARIOS"
    If Check6.Value = 1 Then CAMOPAGO$ = "UNOSOLO"
    Call GRACONTROL("ORPASTA", "FFCAMOMP", CAMOPAGO$)
    '
    '\\\OT-05-0319-WAR-FIN///
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text17.TEXT), "Formularios de Impresion")
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()

    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text19.TEXT = VDEVU$
        Text18.TEXT = VALACT2$("INDIFRM")
    End If
    
End Sub

Private Sub Command22_Click()

   Command22.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text19.TEXT), "Formularios de Impresion")
   Command22.Enabled = True

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
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = Control$("", "FORPASTA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = Control$("", "FORLIFOF")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text9.TEXT = Control$("", "FORRETIB")
    Text13.TEXT = Control("", "FOREIBCA")
    Text14.TEXT = Control("", "FOREIBMI")
    If Control("", "RETIBMI") = "SI" Then Check7.Value = 1
    Text19.TEXT = Control("", "FOREIBTU")
    If Control("", "RETIBTU") = "SI" Then Check8.Value = 1
    Text8.TEXT = ECOARCH$("INDIFRM", Text9.TEXT)
    Text3.TEXT = Control$("", "FENUORPA")
    'agregue estas dos lineas para que cargue
    'el formulario de retencion de ganacias
    Text11.TEXT = Control$("", "FORREGAN")
    Text17.TEXT = Control("", "FORETIVA")
    Text10.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    '
    If Control$("", "NUORPAST") = "AUTONUME" Then Option2.Value = True
    If Control$("", "NURETGAN") = "MANUAL" Then Option10.Value = True
    If Control$("", "TIPAORPA") = "HOJASUEL" Then Option4.Value = True
    If Control$("ORPASTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    If Control$("ORPASTA", "UTICHECA") = "SI" Then Check1.Value = 1
    If Control$("ORPASTA", "SUPRINT") = "SI" Then Check2.Value = 1
    '
    Check4.Value = 0
    Check5.Value = 0
    If Control$("ORPASTA", "SUPRICRE") = "SI" Then Check4.Value = 1
    If Control$("ORPASTA", "SUPRIPAG") = "SI" Then Check5.Value = 1
    '
    '\\\OT-05-0319-WAR-23/05/05///
    If Control$("ORPASTA", "ACUMENS") = "SI" Then Check3.Value = 1
    '\\\OT-05-0319-WAR-FIN///
    MODORECI% = 1: If Control$("ORPASTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    PREFIREM$ = TRIM$(Control$("", "PREFORPA"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "OP"
    
    Check6.Value = 0
    If TRIM$(Control$("ORPASTA", "FFCAMOMP")) = "UNOSOLO" Then Check6.Value = 1
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Text1_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text2 = ECOARCH$("INDIFRM", Text1)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text10 = ECOARCH$("INDIFRM", Text11)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub TEXT11_DBLCLICK()
    '\\\OT-06-0570-RG-27/12/06///
    Call Command11_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text13_DblClick()

    Call Command15_Click
    
End Sub

Private Sub Text14_DblClick()

    Call Command18_Click

End Sub

Private Sub Text17_Change()
    CODFOR$ = TRIM$(Text17)
    Text16 = ECOARCH$("INDISQR", CODFOR$)
End Sub

Private Sub Text17_DblClick()

    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text17.TEXT = CODFOR$
    End If

End Sub

Private Sub Text19_DblClick()
 
    Call Command21_Click
    
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text5_Click()
    Command2.SetFocus
End Sub

Private Sub Text6_Change()
    '\\\OT-06-0570-RG-28/12/06///
    Text7 = ECOARCH$("INDIFRM", Text6)
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text9.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_Click()
    Text9.SetFocus
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("FLFORMS/" + TRIM$(Text9.TEXT), "Formularios de Impresion")
   Command9.Enabled = True
End Sub

Private Sub Text9_Change()
   '\\\OT-06-0570-RG-28/12/06///
   Text8 = ECOARCH$("INDIFRM", Text9)
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text9_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command10_Click
   '\\\OT-06-0570-RG-FIN///
End Sub

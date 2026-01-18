VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPCHEQ07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Gestión de Valores"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
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
   ScaleHeight     =   9060
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H0095C5E1&
      Caption         =   "CONCILIACION BANCARIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   50
      Top             =   8280
      Width           =   5385
      Begin VB.CheckBox chkConciliacionAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Activar Versión Automática"
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
         TabIndex        =   51
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0095C5E1&
      Caption         =   "CÓDIGO RÉGIMEN DE RETENCION DE IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   47
      Top             =   3480
      Width           =   4425
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   2880
         TabIndex        =   49
         Top             =   350
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Regimen de Ret.I.V.A."
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
         Left            =   -615
         TabIndex        =   48
         Top             =   390
         Width           =   3270
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0095C5E1&
      Caption         =   "FORMULARIOS DE IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   5370
      Begin VB.CommandButton Command9 
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
         Height          =   300
         Left            =   1480
         TabIndex        =   35
         Top             =   3480
         Width           =   175
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
         Height          =   300
         Left            =   1480
         TabIndex        =   38
         Top             =   2925
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Height          =   300
         Left            =   180
         TabIndex        =   37
         Top             =   2925
         Width           =   1335
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   180
         TabIndex        =   36
         Top             =   3480
         Width           =   1335
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
         Height          =   300
         Left            =   180
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2385
         Width           =   4215
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
         Height          =   300
         Left            =   3015
         TabIndex        =   32
         Text            =   " "
         Top             =   1995
         Width           =   1215
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
         Left            =   4215
         TabIndex        =   31
         Top             =   1995
         Width           =   175
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
         Height          =   630
         Left            =   4515
         Picture         =   "OPCHEQ07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1995
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   630
         Index           =   2
         Left            =   4515
         Picture         =   "OPCHEQ07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1220
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   630
         Index           =   0
         Left            =   4515
         Picture         =   "OPCHEQ07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   420
         Width           =   650
      End
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
         Height          =   285
         Left            =   4215
         TabIndex        =   25
         Top             =   1215
         Width           =   175
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
         Height          =   300
         Left            =   3015
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Height          =   300
         Left            =   3000
         TabIndex        =   23
         Top             =   400
         Width           =   1215
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
         Left            =   4200
         TabIndex        =   20
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega de Valores"
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
         Left            =   195
         TabIndex        =   42
         Top             =   3240
         Width           =   2430
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción de Cheques"
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
         Left            =   195
         TabIndex        =   41
         Top             =   2685
         Width           =   2430
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   40
         Top             =   2925
         Width           =   3495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   39
         Top             =   3480
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compensación de Cuentas"
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
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   2100
         Width           =   2955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Déposito Bancario"
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
         Index           =   0
         Left            =   630
         TabIndex        =   27
         Top             =   1300
         Width           =   2325
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   1590
         Width           =   4215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   165
         TabIndex        =   22
         Top             =   795
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia de Fondos"
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
         Left            =   510
         TabIndex        =   21
         Top             =   500
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   4680
      ScaleHeight     =   4275
      ScaleWidth      =   1170
      TabIndex        =   13
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   140
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   45
         Top             =   3960
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Width           =   650
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
         Height          =   560
         Index           =   1
         Left            =   140
         Picture         =   "OPCHEQ07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Acceso a Tabla de Cuentas Bancarias"
         Top             =   2565
         Width           =   650
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
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ07.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Archivo Maestro de Cuentas Contables"
         Top             =   2010
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ07.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   195
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   140
         MaskColor       =   &H00404000&
         Picture         =   "OPCHEQ07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0095C5E1&
      Caption         =   "CHEQUES RECHAZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4425
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
         Left            =   3990
         TabIndex        =   5
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0095C5E1&
         Caption         =   "PROCESAMIENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   210
         TabIndex        =   2
         Top             =   1065
         Width           =   4005
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Por Ajuste de Saldo Deudor"
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
            Height          =   450
            Left            =   210
            TabIndex        =   4
            Top             =   350
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Por Nota de Débito al Cliente"
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
            Height          =   450
            Left            =   2100
            TabIndex        =   3
            Top             =   350
            Width           =   1785
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Cheques Rechazados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   285
         Width           =   1485
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   375
         TabIndex        =   6
         Top             =   660
         Width           =   3795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Caption         =   "GESTIÓN DE VALORES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   70
      Width           =   4425
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
         Left            =   3990
         TabIndex        =   9
         Top             =   300
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartera de Valores"
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
         Left            =   105
         TabIndex        =   12
         Top             =   390
         Width           =   2325
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   375
         TabIndex        =   10
         Top             =   675
         Width           =   3795
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPCHEQ07.frx":14D0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPCHEQ07.frx":154A
      Top             =   1200
   End
End
Attribute VB_Name = "OPCHEQ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CargandoFormulario%

Private Sub chkConciliacionAuto_Click()
    
    If chkConciliacionAuto.Value = 0 Or CargandoFormulario% = 1 Then Exit Sub
    Resultado% = CLACONTRA%()
    If Resultado% = 0 Then
        chkConciliacionAuto.Value = 0
    End If
    
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call SELECHO("INDISQR")
   CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text6.SetFocus
      Else
        Text6.TEXT = CODFOR$
    End If
    Command10.Enabled = True

End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARCUEINGAS
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
'
Private Sub Command13_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub


Private Sub Command17_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

'
Private Sub Command2_Click()
    '
    Call GRACONTROL("", "CUECHECA", Label8)
    Call GRACONTROL("", "CUECHERE", Label11)
    '\\\OT-05-0440-WAR-29/06/05///
    Call GRACONTROL("", "FORTRAFO", Text1)
    '\\\OT-05-0440-WAR-FIN///
    '\\\OT-05-0475-WAR-07/07/05///
    Call GRACONTROL("", "FORDEPBA", Text2)
    '\\\OT-05-0475-WAR-FIN///
    '
    'FRM DE IMPRESION DE RECEPCION DE CHEQUES
    Call GRACONTROL("", "IMPRECHE", Text6)
    '
    'FRM DE IMPRESION DE ENTREGA DE VALORES
    Call GRACONTROL("", "IMPENTVA", Text5)
    MONUME$ = "AJUSALDO"
    If Option4.Value = True Then MONUME$ = "NOTADEBI"
    '
    Call GRACONTROL$("OPRECIBO", "COREIVAT", Text3)
    
    Call GRACONTROL$("", "FORCOMPE", TRIM$(Text10))
    '
    VALO$ = "NO"
    If chkConciliacionAuto = 1 Then VALO$ = "SI"
    Call GRACONTROL$("", "CONCIAU", VALO$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    '\\\OT-05-0440-WAR-29/06/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1 = VDEVU$
        Label3 = VALACT2$("INDIFRM")
    End If
    '\\\OT-05-0440-WAR-FIN///
End Sub
'

Private Sub Command4_Click()
    '\\\OT-05-0475-WAR-07/07/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text2 = VDEVU$
        Label2 = VALACT2$("INDIFRM")
    End If
    '\\\OT-05-0475-WAR-FIN///
End Sub

Private Sub Command6_Click(Index As Integer)
   '
   Command6(Index).Enabled = False
   Select Case Index
     Case 0
       Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
     Case 1
       '\\\OT-5-661-FG-06/10/05///
10     X% = ALTERNA%("Tabla de Cuentas Bancarias\Tabla de Cuentas de Tesorería")
       If X% = 1 Then
           Call CARCUEBANC
         '\\\OT-06-0570-RG-28/12/06///
         ElseIf X% = 2 Then
         '\\\OT-06-0570-RG-FIN///
           Call CARCUETES
       End If
       '\\\OT-5-661-FG-FIN///
     Case 2
       Call CONECRUN("FLFORMS/" + TRIM$(Text2.TEXT), "Formularios de Impresion")
   End Select
   Command6(Index).Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call HELPONLINE("CUECONTA")
   Command7.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   '
   Call SELECHO("INDISQR")
   CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text5.SetFocus
      Else
        Text5.TEXT = CODFOR$
    End If
    Command9.Enabled = True
    '
End Sub

Private Sub Form_Load()
    '
    CargandoFormulario% = 1
    Label8 = CONTROL$("", "CUECHECA")
    Label9 = ECOARCH$("ECUECON", Label8)
    Label11 = CONTROL$("", "CUECHERE")
    Label12 = ECOARCH$("ECUECON", Label11)
    '\\\OT-05-0440-WAR-29/06/05
    '///
    Text1 = CONTROL$("", "FORTRAFO")
    Label3 = ECOARCH$("INDIFRM", Text1)
    '\\\OT-05-0440-WAR-FIN///
    '\\\OT-05-0475-WAR-07/07/05///
    Text2 = CONTROL$("", "FORDEPBA")
    Label2 = ECOARCH$("INDIFRM", Text2)
    '\\\OT-05-0475-WAR-07/07/05///
    '\\\OT-06-0144-WAR-17/04/06///
    Text10.TEXT = CONTROL$("", "FORCOMPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    '
    'FRM DE IMPRESION DE RECEPCION DE CHEQUES
    Text6 = CONTROL("", "IMPRECHE")
    Label6 = ECOARCH$("INDISQR", TRIM$(Text6))
    '
    'FRM DE IMPRESION DE ENTREGA DE VALORES
    Text5 = CONTROL("", "IMPENTVA")
    '
    Label5 = ECOARCH$("INDISQR", TRIM$(Text5))
    '\\\OT-06-0144-WAR-FIN///
    Text3 = CONTROL$("OPRECIBO", "COREIVAT")
    '
    If CONTROL$("", "CONCIAU") = "SI" Then
        chkConciliacionAuto.Value = 1
        chkConciliacionAuto.Enabled = False
    End If
    '
    If CONTROL$("", "CHERECHA") = "NOTADEBI" Then Option4.Value = True
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    CargandoFormulario% = 0
    '
End Sub

Private Sub Text1_Change()
    DECUE$ = ECOARCH("INDIFRM", TRIM$(Text1))
    If DECUE$ <> "" Then
        Label3 = DECUE$
      Else
        Label3 = ""
    End If
End Sub
Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_DblClick()
    '\\\OT-06-0570-RG-28/12/06///
    Call Command4_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text5_Change()
    CODFOR$ = TRIM$(Text5)
    Label5 = ECOARCH$("INDISQR", CODFOR$)
End Sub

Private Sub Text6_Change()
    CODFOR$ = TRIM$(Text6)
    Label6 = ECOARCH$("INDISQR", CODFOR$)
 

End Sub

Private Sub Text7_CLICK()
  Text10.SetFocus
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command16_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_Change()
    '\\\OT-05-0475-WAR-07/07/05///
    DECUE$ = ECOARCH("INDIFRM", TRIM$(Text2))
    If DECUE$ <> "" Then
        Label2 = DECUE$
      Else
        Label2 = ""
    End If
    '\\\OT-05-0475-WAR-FIN///
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
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
'\\\OT-5-661-FG-06/10/05///
Sub CARCUETES()
   '
   JJ& = 0
   NUCUEN% = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("TACUECAJ")
      XX$ = REGLEIDO$("TACUECAJ", U&)
      CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
      If CXI% > 0 Then
         If CXI% <= NUCUEN% Then
            JJ& = JJ& + 1
            Call GRAREG("!TACUECAJ", XX$, JJ&)
         End If
      End If
   Next U&
   Call SETULTREG("!TACUECAJ", JJ&)
   Call CIERRARCH("!TACUECAJ")
   Call CIERRARCH("TACUECAJ")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("TACUECAJ")
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      JJ& = 0
      
      REGS& = ULTREG&("!TACUECAJ")
      ReDim VecTaCueCaj(REGS&) As String
      '
      For U& = 1 To ULTREG&("!TACUECAJ")
         XX$ = REGLEIDO$("!TACUECAJ", U&)
         CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
         Cuenta$ = TRIM$(Left$(XX$, 12))
         'VERIFICA SI YA ESTA INGRESADO
         For AX& = 1 To REGS&
           If XVALO(VecTaCueCaj(AX&)) = CXI% Then
              COMUNI ("Cuenta " & Cuenta$ & " ya Ingresada, Se Suprime la Segunda Aparición")
              GoTo 90 'Si ya esta pasa al otro registro de !Tacuecaj
           End If
         Next AX&
         VecTaCueCaj(U&) = CXI%  'En caso de NO estar registrado se registra
         '
         If CXI% > 0 Then
            If CXI% <= NUCUEN% Then
               JJ& = JJ& + 1
               Call GRAREG("TACUECAJ", XX$, JJ&)
            End If
         End If
90    Next U&
      Call SETULTREG("TACUECAJ", JJ&)
      Call CIERRARCH("TACUECAJ")
      Call FRESHECHO("TACUECAJ")
      Call CIERRARCH("!TACUECAJ")
      Screen.MousePointer = 1
   End If
   '
End Sub
Sub CONFI_CONCI_AUTO()
     '
'     If InStr(UCase$(EMPREAC$), "DOSIVA") < 1 Then
'        Call MENSERR(24, "No Disponible en la Presente Versión")
'        Exit Sub
'     End If
     '
     '
10   OPX% = COMALTER%("Configuración - Conciliación\\Tabla Gastos Extras\Tabla Posiciones de Lectura")
     '
     Select Case OPX%
       Case 1:
          Call CARGA_TABLA_GASEXTRA
          GoTo 10
       Case 2:
          Call CARGA_TABLA_CFGCONCI
          GoTo 10
       Case Else
         Exit Sub
     End Select
End Sub
Sub CARGA_TABLA_GASEXTRA()
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias")
      Call CARCUEBANC
      Exit Sub
   End If
   '
   
   Call SELECHO("CUEBANC")
   CTA_SELECCIONADA$ = VALACT1$("CUEBANC")
   If CTA_SELECCIONADA$ = "" Then
      Exit Sub
   End If
   '
   'LEO EL CENTRAL Y GRABO EN EL LOCAL SOLO
   Call SETULTREG("!GASEXTRA", 0)
   For KU& = 1 To ULTREG&("GASEXTRA")
       XX$ = REGLEIDO$("GASEXTRA", KU&)
       CUENTACON$ = Left$(XX$, 12)
       If TRIM$(CUENTACON$) = CTA_SELECCIONADA$ Then
         j& = j& + 1
         Call GRAREG("!GASEXTRA", XX$, j&)
       End If
   Next KU&
   Call SETULTREG("!GASEXTRA", j&)
   '
   '
   VTB% = VENTABU%("GASEXTRA/TITUPAN=GASTOS PARA LA CUENTA: " & CTA_SELECCIONADA$)
   If VTB% >= 0 Then
     '
     For KU& = 1 To ULTREG&("!GASEXTRA")
       XX$ = REGLEIDO$("!GASEXTRA", KU&)
       CUENTACON$ = Left$(XX$, 12)
       If TRIM$(CUENTACON$) = CTA_SELECCIONADA$ Then
         REFEGASTO$ = TRIM$(Mid$(XX$, 13, 128))
         CTA_SELECC$ = TRIM$(Mid$(XX$, 141, 12))
         If CTA_SELECC$ = "" Then
           Call COMUNI("Falta Seleccionar Cuenta Asociada Para Gasto: \" & REFEGASTO$)
           Exit Sub
         End If
       End If
     Next KU&

     'BORRO DEL CENTRAL LOS DATOS QUE MACHEAN CON LA CUENTA CONTABLE
     Call FILTERDELETE("GASEXTRA", 1, AJUSTI$(CTA_SELECCIONADA$, 12))
     '
     'GRABO EN EL LOCAL
     For KU& = 1 To ULTREG&("!GASEXTRA")
       XX$ = REGLEIDO$("!GASEXTRA", KU&)
       Call REPLA(XX$, CTA_SELECCIONADA$, 1, 12)
       Call REGAPP("GASEXTRA", XX$)
     Next KU&
     '
     'PASO AL CENTRAL
     Call CIERRARCH("!GASEXTRA")
'     Call TRACENTRAL("GASEXTRA", "")
   End If

End Sub
Sub CARGA_TABLA_CFGCONCI()
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias")
      Call CARCUEBANC
      Exit Sub
   End If
   '
   Call TRALOCAL("CFGCONCI", "")
   j& = 0
   VTB% = VENTABU%("CFGCONCI")
   If VTB% >= 0 Then
     For KU& = 1 To ULTREG&("!CFGCONCI")
       XX$ = REGLEIDO$("!CFGCONCI", KU&)
       CUENTACON$ = Left$(XX$, 12)
       If CUEINT%(CUENTACON$) > 0 Then
         j& = j& + 1
         Call GRAREG("!CFGCONCI", XX$, j&)
       End If
     Next KU&
     Call CIERRARCH("!CFGCONCI")
     Call TRACENTRAL("CFGCONCI", "")
   End If

End Sub


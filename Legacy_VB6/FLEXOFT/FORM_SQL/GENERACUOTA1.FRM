VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form GENERACUOTA1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GENERACION MASIVA DE CUOTAS"
   ClientHeight    =   5250
   ClientLeft      =   540
   ClientTop       =   2505
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   1035
      TabIndex        =   66
      Top             =   60
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   16
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Text            =   "Legajo:"
      Top             =   60
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Generación de MatrÍcula"
      Height          =   2175
      Left            =   120
      TabIndex        =   57
      Top             =   3000
      Width           =   4455
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   14
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "SELECCIONAR MES Y AÑO DE VENCIMIENTO"
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   13
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "Año"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   12
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "Mes"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Generar Matrícula"
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
         Left            =   120
         Picture         =   "GENERACUOTA1.frx":0000
         TabIndex        =   58
         Top             =   1750
         Width           =   4200
      End
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   1320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Carga de Importe de Matrícula"
      Height          =   1215
      Left            =   4680
      TabIndex        =   49
      Top             =   3960
      Width           =   6855
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Importe Matrícula"
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Mes"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text77 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5640
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Enseñanza Programática"
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Año"
         Top             =   340
         Width           =   375
      End
      Begin VB.ComboBox Combo11 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox Combo22 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   330
         Width           =   1440
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar Matrícula"
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
         Left            =   4080
         Picture         =   "GENERACUOTA1.frx":014A
         TabIndex        =   50
         Top             =   720
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Lista de Precios"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   4680
      TabIndex        =   24
      Top             =   600
      Width           =   6855
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   325
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Seleccione Mes y Año Cuota"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   325
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Mes"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox Combo10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   360
         Width           =   1320
      End
      Begin VB.ComboBox Combo20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   325
         Index           =   10
         Left            =   4910
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Año"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar Lista de Cuotas"
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
         Left            =   4080
         Picture         =   "GENERACUOTA1.frx":0294
         TabIndex        =   43
         Top             =   2760
         Width           =   2520
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Enseñanza Programática"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   45
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Total"
         ToolTipText     =   "Enseñanza Programática"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   4920
         TabIndex        =   15
         ToolTipText     =   "Enseñanza Programática"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   4080
         TabIndex        =   14
         ToolTipText     =   "Enseñanza Programática"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "Enseñanza Programática"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TEXT3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "Enseñanza Programática"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Secundaria"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "4"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   4920
         TabIndex        =   11
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   3240
         TabIndex        =   9
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TEXT3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "4to. a 7mo."
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "3"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   4920
         TabIndex        =   7
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   4080
         TabIndex        =   6
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TEXT3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "1ro. a 3ro."
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   4920
         TabIndex        =   3
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   4080
         TabIndex        =   2
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TEXT3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   2400
         TabIndex        =   0
         ToolTipText     =   "Enseñanza Programática"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "JARDIN"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Comput."
         ToolTipText     =   "Enseñanza Programática"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Ingles"
         ToolTipText     =   "Enseñanza Programática"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Emerg."
         ToolTipText     =   "Enseñanza Programática"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Ens.Prog."
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Referencia"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Lista Nº"
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Generar Cuotas Mensuales"
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   15
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "SELECCIONAR MES Y AÑO DE CUOTA"
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Cuota"
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
         Left            =   120
         Picture         =   "GENERACUOTA1.frx":03DE
         TabIndex        =   23
         Top             =   1600
         Width           =   4200
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Mes"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Año"
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   2
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Generación de Cuota / Matrícula"
      Top             =   0
      Width           =   11655
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "GENERACUOTA1.frx":0528
      Top             =   0
   End
End
Attribute VB_Name = "GENERACUOTA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CARGADATOS_LISTA()
    ''
    Call LIMPIAR_LISTA_CUOTAS
    '
    MEX% = XVALO(Mid$(Combo10.Text, 30, 4))
    ANOX% = XVALO(Combo20.Text)
    '
    If MEX% < 1 Or ANOX% < 1 Then Exit Sub
    '
    Call ABRECONEXION

    Dim AUX As ADODB.Recordset
    
    For i% = 1 To 4
       SQLX$ = "SELECT * FROM LISTAPRE "
       SQLX$ = SQLX$ + " WHERE COD_LISTA = " & i%
       SQLX$ = SQLX$ + " AND Corresponde_mes = " & MEX%
       SQLX$ = SQLX$ + " AND Corresponde_ano = " & ANOX%
       Set RS1 = BASSSEX.Execute(SQLX$)
       With RS1
         If Not .EOF Then
          Text3(i% - 1) = FORMATNUM$(XVALO(!ENSE_PROGRAMADA), "F9.2")
          Text4(i% - 1) = FORMATNUM$(XVALO(!EMERGENCIA), "F9.2")
          Text5(i% - 1) = FORMATNUM$(XVALO(!INGLES), "F9.2")
          Text6(i% - 1) = FORMATNUM$(XVALO(!COMPUTACION), "F9.2")
          Text7(i% - 1) = FORMATNUM$(XVALO(!impo_total), "F9.2")
        End If
       End With
    Next i%
    '
    RS1.Close
    Set RS1 = Nothing

End Sub

Sub LIMPIAR_LISTA_CUOTAS()
    For i% = 1 To 4
       Text3(i% - 1) = ""
       Text4(i% - 1) = ""
       Text5(i% - 1) = ""
       Text6(i% - 1) = ""
       Text7(i% - 1) = ""
    Next i%
End Sub

Private Sub Combo1_Click()
   Combo10.Text = Combo1.Text

End Sub

Private Sub Combo10_Click()
   Call CARGADATOS_LISTA
End Sub

Private Sub Combo11_Click()
    '
    Text77 = ""
    Call CARGA_MATRICULA
    '
End Sub

Sub CARGA_MATRICULA()
    '
    ANOX% = XVALO(Combo22.Text)
    '
    If ANOX% < 1 Then Exit Sub
    '
    Call ABRECONEXION
    CONDI$ = "COD_LISTA = 999 "
    CONDI$ = CONDI$ + " AND Corresponde_ano = " & ANOX%
    CONDI$ = CONDI$ + " AND REFERENCIA = 'MATRICULA'"
    Text77 = FORMATNUM$(XVALO(VALOADMIN("listapre", "IMPO_TOTAL", CONDI$, "NOMESS")), "F9.2")
End Sub

Private Sub Combo2_Click()
  Combo20.Text = Combo2.Text

End Sub

Private Sub Combo20_Click()

  Call CARGADATOS_LISTA
End Sub


Private Sub Combo22_Change()
    Text77 = ""
    Call CARGA_MATRICULA

End Sub

Private Sub Combo22_Click()
    Text77 = ""
    Call CARGA_MATRICULA

End Sub


Private Sub Command1_Click()
      Command1.Enabled = False
      '
      MEX% = XVALO(Mid$(Combo1.Text, 30, 4))
      ANOX% = XVALO(Combo2.Text)
      ANO_ACTUAL% = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7))

      'VALIDACIONES
      If MEX% < 1 Or MEX > 12 Then
         MsgBox "MES INGRESADO PARA MATRÍCULA NO VÁLIDO"
         GoTo 99
      End If
      '
      If ANOX% < 1 Then
         MsgBox "AÑO INGRESADO PARA MATRÍCULA NO VÁLIDO"
         GoTo 99
      End If

      If ANOX% < 1 Or (ANO_ACTUAL% <> ANOX%) Then
        TTXX$ = "AÑO INGRESADO PARA CUOTA DEL MES  DIFIERE DEL ACTUAL " & vbCrLf & vbCrLf
        TTXX$ = TTXX$ + "¿¿¿¿¿¿¿¿-  DESEA GRABAR DE TODOS MODOS  -????????? " & vbCrLf
        If MsgBox(TTXX$, vbYesNo, "VALIDACION DE AÑO") <> vbYes Then GoTo 99
      End If
      '
      CONDI$ = "COD_LISTA > 0 AND Corresponde_mes = " & MEX% & " AND Corresponde_ano = " & ANOX%
      EXISTELISTA& = CantRegistros&("LISTAPRE", CONDI$, "NOMESS")
     
      If EXISTELISTA& < 1 Then
         TXXT$ = "IMPOSIBLE GENERAR CUOTAS" & vbCrLf
         TXXT$ = TXXT$ + "NO EXISTE LISTAS DE PRECIOS GRABADA" & vbCrLf
         MsgBox TXXT$ + "PARA EL MES Y AÑO SELECCIONADO"
         GoTo 99
      End If
      '
      For i% = 0 To 3
         If XVALO(Text7(i%)) < 1 Then
           TX$ = "IMPOSIBLE GUARDAR CUOTAS" + vbCrLf + vbCrLf
           TX$ = TX$ + "LISTA DE PRECIOS PARA " & TEXT2(i%).Text & " SIN VALOR"
           TX$ = TX$ + " ASIGNE IMPORTE A " & TEXT2(i%).Text & " Y VUELVA A INTENTAR"
           MsgBox TX$
           GoTo 99
         End If
      Next i%
      '
      If Text7(0) < 1 Then
         MsgBox "AÑO INGRESADO PARA CUOTA NO VÁLIDO"
         GoTo 99
      End If
      '
      TX$ = "SE VA A GENERAR PROCESO DE CUOTAS " + vbCrLf
      TX$ = TX$ + "NO INTERRUMPA EL MISMO AGUARDE MENSAJE DE FINALIZACIÓN" + vbCrLf + vbCrLf
      TX$ = TX$ + "PULSE ACEPTAR PARA CONTINUAR"
      MsgBox TX$
      '
      LGJOTXT& = XVALO(Text1(5))
      Screen.MousePointer = 11
      '
      'BARRA DE PROGRESO
      CONDI$ = " BECADO < 1 AND STATUS >= 0"
      If LGJOTXT& > 0 Then CONDI$ = CONDI$ + " AND LEGAJO = " & LGJOTXT&
      
      REG& = CantRegistros&("ALUMNOS", CONDI$, "NOMESS")
      
      If REG& < 1 Then MsgBox "SIN ALUMNOS CARGADOS": GoTo 99
      ProgressBar1.Min = 0
      ProgressBar1.Value = 0
      ProgressBar1.Max = REG&
   
      'RECORRE LOS ALUMNOS Y LE GRABA LA CUOTA PARA EL MES CORRESPONDIENTE
      SQLX$ = "SELECT * FROM ALUMNOS "
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      Set RS1 = BASSSEX.Execute(SQLX$)
      With RS1
        Do While Not .EOF
           LGJO& = XVALO(!LEGAJO)
           'BARRA DE PROGRESO
           PROGRE% = PROGRE% + 1
           ProgressBar1.Visible = True
           ProgressBar1.Value = PROGRE%
           Command1.Caption = SAFETEXT$(CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %")
           DoEvents
           If LGJO& > 0 Then
             ENSE_PROGRAMADA# = PRECUOTA#(LGJO&, MEX%, ANOX%, "ENSE_PROGRAMADA")
             EMERGENCIA# = PRECUOTA#(LGJO&, MEX%, ANOX%, "EMERGENCIA")
             INGLES# = PRECUOTA#(LGJO&, MEX%, ANOX%, "INGLES")
             COMPUTACION# = PRECUOTA#(LGJO&, MEX%, ANOX%, "COMPUTACION")
             IMPO_TOTA# = ENSE_PROGRAMADA# + EMERGENCIA# + INGLES# + COMPUTACION#
             '
             TIPOGENERACION$ = "Generación de Cuotas Masivas"
             If LGJOTXT& > 0 Then TIPOGENERACION$ = "Generación de Cuota Individual"
             Call GRABA_CUOTA_CAJABANC(LGJO&, MEX%, ANOX%, ENSE_PROGRAMADA#, EMERGENCIA#, INGLES#, COMPUTACION#, 0, "CUOTA", "CUOTA", TIPOGENERACION$)
            End If
50          .MoveNext
        Loop
       End With
       RS1.Close
       Set RS1 = Nothing
       '
       MsgBox "HECHO"
99     Command1.Enabled = True
       Command1.Caption = "Generar Cuotas"
      
       Screen.MousePointer = 1
       ProgressBar1.Value = 0
       '
       'SI VIENE DE EDIRECIBO
       If LGJOTXT& > 0 Then
          Text1(5) = ""
          Unload Me
       End If
       
End Sub


Private Sub Command2_Click()

    Command2.Enabled = False
    VUELTA% = 0
    MEX% = XVALO(Mid$(Combo10.Text, 30, 4))
    ANOX% = XVALO(Combo20.Text)
    ANO_ACTUAL% = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7))
    
    'VALIDACIONES
    If MEX% < 1 Or MEX > 12 Then
       MsgBox "MES INGRESADO PARA CUOTA NO VÁLIDO"
       GoTo 99
    End If
    '
    If ANOX% < 1 Then
       MsgBox "AÑO INGRESADO PARA CUOTA NO VÁLIDO"
       GoTo 99
    End If
    '
    If ANOX% < 1 Or (ANO_ACTUAL% <> ANOX%) Then
       TTXX$ = "AÑO INGRESADO PARA CUOTA DIFIERE DEL ACTUAL " & vbCrLf & vbCrLf
       TTXX$ = TTXX$ + "¿¿¿¿¿¿¿¿-  DESEA GRABAR DE TODOS MODOS  -????????? " & vbCrLf
       If MsgBox(TTXX$, vbYesNo, "VALIDACION DE AÑO") <> vbYes Then GoTo 99
    End If
    '
    'VALIDO QUE TODAS LAS LISTAS TENGAN IMPORTE
    For i% = 1 To 4
       IMPOR_TOTAL# = XVALO(Text7(i% - 1))
       If IMPOR_TOTAL# < 0.05 Then
          TX$ = "IMPOSIBLE GRABAR" + vbCrLf
          TX$ = TX$ + "FALTA COMPLETAR IMPORTE PARA LISTA " & i%
          MsgBox TX$
          GoTo 99
       End If
    Next i%
    '
    
    For i% = 1 To 4
      SQLX$ = "SELECT * FROM LISTAPRE"
      SQLX$ = SQLX$ + " WHERE COD_LISTA = " & i%
      SQLX$ = SQLX$ + " AND Corresponde_mes = " & MEX%
      SQLX$ = SQLX$ + " AND Corresponde_ano = " & ANOX%
      On Error GoTo ERROR_GUARDAR
      Set AUX = New ADODB.Recordset
      AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
      If VUELTA% < 1 Then BASSSEX.BeginTrans
      With AUX
         If (.BOF And .EOF) Then
          .AddNew
          !COD_LISTA = i%
          !Corresponde_mes = MEX%
          !Corresponde_ano = ANOX%
          !REFERENCIA = TEXT2(i% - 1)
        End If
          !ENSE_PROGRAMADA = XVALO(Text3(i% - 1))
          !EMERGENCIA = XVALO(Text4(i% - 1))
          !INGLES = XVALO(Text5(i% - 1))
          !COMPUTACION = XVALO(Text6(i% - 1))
             IMPOR_TOTAL# = XVALO(Text7(i% - 1))
          !impo_total = IMPOR_TOTAL#
          .Update
           VUELTA% = VUELTA% + 1
      End With
    Next i%
    
ERROR_GUARDAR:
    If Err <> 0 Then
      BASSSEX.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      BASSSEX.CommitTrans
      AUX.Close
      Set AUX = Nothing
    End If
    MsgBox "HECHO"
    
99  Command2.Enabled = True
End Sub

Private Sub Command25_Click()
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    
    ANOX% = XVALO(Combo22.Text)
    IMPOR_TOTAL# = XVALO(Text77)
    ANO_ACTUAL% = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7))
    'VALIDACIONES
'    If MEX% < 1 Or MEX > 12 Then
'       MsgBox "MES INGRESADO PARA MATRÍCULA NO VÁLIDO"
'       GoTo 99
'    End If
    '
    If ANOX% < 1 Or (ANO_ACTUAL% <> ANOX%) Then
       TTXX$ = "AÑO INGRESADO PARA MATRÍCULA DIFIERE DEL ACTUAL " & vbCrLf & vbCrLf
       TTXX$ = TTXX$ + "¿¿¿¿¿¿¿¿-  DESEA GRABAR DE TODOS MODOS  -????????? " & vbCrLf
       If MsgBox(TTXX$, vbYesNo, "VALIDACION DE AÑO") <> vbYes Then GoTo 99
    End If
    
    If IMPOR_TOTAL# < 0.05 Then
       MsgBox "FALTA INGRESAR IMPORTE DE MATRÍCULA"
       GoTo 99
    End If
    '
    'VALIDO SI YA EXISTE EN LA LISTA
    CONDI$ = "COD_LISTA = 999"
    CONDI$ = CONDI$ + " AND Corresponde_mes = 999"
    CONDI$ = CONDI$ + " AND Corresponde_ano = " & ANOX%
    '
    If CantRegistros&("LISTAPRE", CONDI$, "NOMESS") > 0 Then
       TTXX$ = "MATRÍCULA EXISTENTE PARA ESTE AÑO " & vbCrLf & vbCrLf
       TTXX$ = TTXX$ + "SI DECIDE SOBRE-ESCRIBIR LA INFORMACIÓN" + vbCrLf
       TTXX$ = TTXX$ + "DEBERA ACTUALIZAR LA GENERACION DE MATRICULAS" + vbCrLf + vbCrLf + vbCrLf
       TTXX$ = TTXX$ + "¿¿¿¿¿¿¿¿-  DESEA SOBRE-ESCRIBIR LA INFORMACIÓN  -????????? "
       If MsgBox(TTXX$, vbYesNo, "VALIDACION DE AÑO") <> vbYes Then GoTo 99
    End If
    
    SQLX$ = "SELECT * FROM LISTAPRE"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    On Error GoTo ERROR_GUARDAR
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    BASSSEX.BeginTrans
    With AUX
        If (.BOF And .EOF) Then
          .AddNew
          !COD_LISTA = 999
          !Corresponde_mes = 999
          !Corresponde_ano = ANOX%
          !REFERENCIA = "MATRICULA"
        End If
        !impo_total = XVALO(Text77)
        !ENSE_PROGRAMADA = 0
        !EMERGENCIA = 0
        !INGLES = 0
        !COMPUTACION = 0
        !impo_total = IMPOR_TOTAL#
        .Update
    End With
    
ERROR_GUARDAR:
    If Err <> 0 Then
      BASSSEX.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      BASSSEX.CommitTrans
      AUX.Close
      Set AUX = Nothing
    End If
    MsgBox "HECHO"
    
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
      Command4.Enabled = False
      '
      MEX% = XVALO(Mid$(Combo3.Text, 30, 4))
      ANOX% = XVALO(Combo4.Text)
      ANO_ACTUAL% = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7))
    
      'VALIDACIONES
      If MEX% < 1 Or MEX > 12 Then
         MsgBox "MES INGRESADO PARA MATRÍCULA NO VÁLIDO"
         GoTo 99
      End If
      '
      If ANOX% < 1 Then
         MsgBox "AÑO INGRESADO PARA MATRÍCULA NO VÁLIDO"
         GoTo 99
      End If
      '
      
      If ANO_ACTUAL% <> ANOX% Then
        TTXX$ = "AÑO INGRESADO PARA MATRÍCULA DIFIERE DEL ACTUAL " & vbCrLf & vbCrLf
        TTXX$ = TTXX$ + "¿¿¿¿¿¿¿¿-  DESEA GRABAR DE TODOS MODOS  -????????? " & vbCrLf
        If MsgBox(TTXX$, vbYesNo, "VALIDACION DE AÑO") <> vbYes Then GoTo 99
      End If
      '
      CONDI$ = "COD_LISTA = 999 AND Corresponde_ano = " & ANOX% & " AND IMPO_TOTAL > 0 "
      EXISTELISTA& = CantRegistros&("LISTAPRE", CONDI$, "NOMESS")
      
      If EXISTELISTA& < 1 Then
         MsgBox "IMPOSIBLE GENERAR CUOTAS" & vbCrLf & "NO EXISTEN LISTAS DE MATRICULAS " & vbCrLf & "GENERADAS PARA EL AÑO SELECCIONADO"
         GoTo 99
      End If
      '
      LGJOTXT& = XVALO(Text1(5))
      '
      Screen.MousePointer = 11
      'BARRA DE PROGRESO
      CONDI$ = " BECADO < 1 AND STATUS >= 0"
      REG& = CantRegistros&("ALUMNOS", CONDI$, "NOMESS")
      If REG& < 1 Then MsgBox "SIN ALUMNOS CARGADOS": GoTo 99

      ProgressBar3.Min = 0
      ProgressBar3.Value = 0
      ProgressBar3.Max = REG&
      
      SQLX$ = "SELECT * FROM ALUMNOS WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE " 'BECADO < 1" 'COMENTADO P.Q. A LOS BECADOS IGUAL CORRESPONDE MATRICULA
      SQLX$ = SQLX$ + " STATUS >= 0"
      If LGJOTXT& > 0 Then SQLX$ = SQLX$ + " AND LEGAJO = " & LGJOTXT&

      Set RS1 = BASSSEX.Execute(SQLX$)
      With RS1
        Do While Not .EOF
           LGJO& = XVALO(!LEGAJO)
           'BARRA DE PROGRESO
           FILA& = FILA& + 1
           ProgressBar3.Visible = True
           ProgressBar3.Value = FILA&
           Command4.Caption = SAFETEXT$(CLng((ProgressBar3.Value * 100) / ProgressBar3.Max) & " %")
           DoEvents

           If LGJO& > 0 Then
             MATRICULA1# = PRECUOTA#(LGJO&, 999, ANOX%, "MATRICULA")
             If MATRICULA1# < 1 Then GoTo 98
             TIPOGENERACION$ = "Generación de Matrícula Masivas"
             If LGJOTXT& > 0 Then TIPOGENERACION$ = "Generación de Matrícula Individual"

             Call GRABA_CUOTA_CAJABANC(LGJO&, MEX%, ANOX%, 0, 0, 0, 0, MATRICULA1#, "MATRICULA", "MATRICULA", TIPOGENERACION$)
           End If
           
           .MoveNext
        Loop
       End With
       
98     RS1.Close
       Set RS1 = Nothing
       '
       MsgBox "HECHO"
99     Command4.Enabled = True
       Command4.Caption = "Generar Matrícula"
       ProgressBar3.Value = 0
       Screen.MousePointer = 1
       '
       'SI VIENE DE EDIRECIBO
       If LGJOTXT& > 0 Then
          Text1(5) = ""
          Unload Me
       End If

End Sub

Private Sub Command98_Click()
   Command98.Enabled = False
   Unload Me
   Command98.Enabled = True
End Sub


Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
'NOTA
'LA CUOTAS SE GUARDAN EN LA TABLA CAJABANC_COL CON TIPOMOVIM = CUOTA
'CUANDO SE GRABAN DE MANERA MASIVA VERIFICA SI TIENE OTRO VALOR LO ACTUALIZA
    Call CARGAVECTOR_MES

    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call CARGAMESES(GENERACUOTA.Combo1)
    Call CARGAAÑOS(GENERACUOTA.Combo2)
    Call CARGAMESES(GENERACUOTA.Combo10)
    Call CARGAMESES(GENERACUOTA.Combo3)
    Call CARGAAÑOS(GENERACUOTA.Combo20)
    Call CARGAAÑOS(GENERACUOTA.Combo4)
    Call CARGAMESES(GENERACUOTA.Combo11)
    Call CARGAAÑOS(GENERACUOTA.Combo22)

'    Call CARGARLISTAPRECIOS
    
End Sub
Sub CARGAMESES(COMBOMES As ComboBox)

    For i% = 1 To 12
        Call REPLA(TX$, MES(i%), 1, 12)
        Call REPLA(TX$, SAFETEXT$(i%), 30, 4)
        COMBOMES.AddItem TX$
    Next i%
End Sub
Sub CARGAAÑOS(COMBOAÑOS As ComboBox)
        
    ANOX% = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7))
    For j% = 0 To 2
        For i% = ANOX% - 1 To ANOX% + 10
           COMBOAÑOS.AddItem SAFETEXT$(i%)
        Next i%
    Next j%
End Sub
    

'''Sub CARGAMESES(COMBOFECHA As ComboBox)
'''    Dim MES$(12), DIAS%(12)
'''    'TOMO EL AÑO ACTUAL MENOS 1 AÑO
'''    COMBOFECHA.Clear
'''    ANO$ = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7)) - 1 ' PARA QUE EMPIECE UN AÑO ATRAS
'''    For i% = 1 To 3
'''        MES$(1) = "ENERO      " & ANO$
'''        MES$(2) = "FEBRERO    " & ANO$
'''        MES$(3) = "MARZO      " & ANO$
'''        MES$(4) = "ABRIL      " & ANO$
'''        MES$(5) = "MAYO       " & ANO$
'''        MES$(6) = "JUNIO      " & ANO$
'''        MES$(7) = "JULIO      " & ANO$
'''        MES$(8) = "AGOSTO     " & ANO$
'''        MES$(9) = "SETIEMBRE  " & ANO$
'''        MES$(10) = "OCTUBRE    " & ANO$
'''        MES$(11) = "NOVIEMBRE  " & ANO$
'''        MES$(12) = "DICIEMBRE  " & ANO$
'''        For I2% = 1 To 12
'''           COMBOFECHA.AddItem MES$(I2%)
'''        Next I2%
'''        ANO$ = SAFETEXT$(XVALO(ANO$) + 1)
'''    Next i%
'''    'LE PONGO AL TEXT TAPA EL MES ACUTAL
'''
'''
'''''    VUELTA% = VUELTA% + 1
'''''    If VUELTA% < 2 Then
'''''       ANO$ = SAFETEXT$(XVALO(ANO$) + 1)
'''''    End If
'''
'''End Sub


Private Sub TEXT3_Change(Index As Integer)
   SUMA = XVALO(Text3(Index)) + XVALO(Text4(Index)) + XVALO(Text5(Index)) + XVALO(Text6(Index))
   Text7(Index) = FORMATNUM$(SUMA, "F9.2")
End Sub

Private Sub TEXT3_LostFocus(Index As Integer)
   Text3(Index) = FORMATNUM$(Text3(Index), "F9.2")
End Sub

Private Sub Text4_Change(Index As Integer)
   SUMA = XVALO(Text3(Index)) + XVALO(Text4(Index)) + XVALO(Text5(Index)) + XVALO(Text6(Index))
   Text7(Index) = FORMATNUM$(SUMA, "F9.2")

End Sub

Private Sub Text4_LostFocus(Index As Integer)
   Text4(Index) = FORMATNUM$(Text4(Index), "F9.2")

End Sub

Private Sub Text5_Change(Index As Integer)
   SUMA = XVALO(Text3(Index)) + XVALO(Text4(Index)) + XVALO(Text5(Index)) + XVALO(Text6(Index))
   Text7(Index) = FORMATNUM$(SUMA, "F9.2")

End Sub

Private Sub Text5_LostFocus(Index As Integer)
   Text5(Index) = FORMATNUM$(Text5(Index), "F9.2")

End Sub

Private Sub Text6_Change(Index As Integer)
   SUMA = XVALO(Text3(Index)) + XVALO(Text4(Index)) + XVALO(Text5(Index)) + XVALO(Text6(Index))
   Text7(Index) = FORMATNUM$(SUMA, "F9.2")

End Sub

Private Sub Text6_LostFocus(Index As Integer)
   Text6(Index) = FORMATNUM$(Text6(Index), "F9.2")

End Sub

Private Sub Text7_LostFocus(Index As Integer)
   Text7(Index) = FORMATNUM$(Text7(Index), "F9.2")

End Sub

Private Sub Text77_LostFocus()
    Text77 = FORMATNUM$(Text77, "F9.2")
End Sub



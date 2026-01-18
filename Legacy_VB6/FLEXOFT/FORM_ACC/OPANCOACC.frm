VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OpAnCoAcc 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Costos - Tablas y Configuracion"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "PRECIO DE LISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4680
      TabIndex        =   42
      Top             =   3600
      Width           =   2715
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Presenta Primera aparición encontrada o la subsiguiente"
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
         Height          =   750
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1800
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONVERSIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   7560
      TabIndex        =   38
      Top             =   3600
      Width           =   1755
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Habilita Fórmula Por Millar"
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
         Height          =   585
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   9450
      ScaleHeight     =   5340
      ScaleWidth      =   1350
      TabIndex        =   33
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   45
         Top             =   4320
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
      Begin VB.PictureBox Picture16 
         Height          =   4900
         Left            =   840
         ScaleHeight     =   4845
         ScaleWidth      =   1035
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   1100
         Begin VB.Image Image3 
            Height          =   4905
            Left            =   600
            Picture         =   "OPANCOACC.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         TabIndex        =   37
         ToolTipText     =   "Coeficientes de Gastos Generales"
         Top             =   2330
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "OPANCOACC.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1640
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "OPANCOACC.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Picture         =   "OPANCOACC.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Guarda - Cierra - Abandona la Aplicación"
         Top             =   240
         Width           =   650
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "TABLAS DE VALIDACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "INDICES DE AJUSTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   210
         TabIndex        =   29
         Top             =   3380
         Width           =   4005
         Begin VB.CommandButton Command4 
            Height          =   330
            Left            =   3360
            Picture         =   "OPANCOACC.frx":2AF6
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Cotización Moneda Extranjera"
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
            Index           =   12
            Left            =   315
            TabIndex        =   32
            Top             =   315
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Indices de Ajuste por Inflación"
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
            Index           =   11
            Left            =   315
            TabIndex        =   31
            Top             =   565
            Width           =   3375
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "CRITERIOS Y TABLAS GENERALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   4005
         Begin VB.CommandButton Command5 
            Height          =   330
            Left            =   3360
            Picture         =   "OPANCOACC.frx":2C40
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Unidades de Medida"
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
            Index           =   9
            Left            =   315
            TabIndex        =   8
            Top             =   565
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Tipos de Material"
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
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   315
            Width           =   3480
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONFIGURACIÓN INDUSTRIAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   1
         Top             =   1500
         Width           =   4005
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Proveedores Internos"
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
            Index           =   10
            Left            =   315
            TabIndex        =   28
            Top             =   1315
            Width           =   3500
         End
         Begin VB.CommandButton Command6 
            Height          =   330
            Left            =   3360
            Picture         =   "OPANCOACC.frx":2D8A
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Categorias de Operarios Productivos"
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
            Index           =   8
            Left            =   315
            TabIndex        =   7
            Top             =   1065
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Máquinas y Equipos de Producción"
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
            Index           =   7
            Left            =   315
            TabIndex        =   6
            Top             =   816
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Secciones y Centros de Costos"
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
            Index           =   6
            Left            =   315
            TabIndex        =   5
            Top             =   565
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Depósitos y Almacenes"
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
            Index           =   5
            Left            =   315
            TabIndex        =   4
            Top             =   310
            Width           =   3500
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "ACTUALIZACIÓN MONETARIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   4680
      TabIndex        =   16
      Top             =   120
      Width           =   4635
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "AJUSTE POR INFLACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   210
         TabIndex        =   24
         Top             =   250
         Width           =   2910
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Ajuste Automático por Inflación (Costos en Moneda Argentina)"
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
            Height          =   750
            Left            =   360
            TabIndex        =   25
            Top             =   315
            Width           =   2430
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   210
         TabIndex        =   22
         Top             =   1890
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         Caption         =   "DECIMALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   3210
         TabIndex        =   17
         Top             =   250
         Width           =   1215
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "0"
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
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   21
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "2"
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
            Height          =   225
            Index           =   2
            Left            =   210
            TabIndex        =   20
            Top             =   525
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "4"
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
            Height          =   225
            Index           =   4
            Left            =   210
            TabIndex        =   19
            Top             =   735
            Value           =   -1  'True
            Width           =   435
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "5"
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
            Height          =   225
            Index           =   5
            Left            =   210
            TabIndex        =   18
            Top             =   945
            Width           =   435
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda / Coeficiente         T-C"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   1600
         Width           =   4215
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   4440
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fuera de uso"
      Height          =   1800
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
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
         Index           =   3
         Left            =   1260
         TabIndex        =   15
         Top             =   525
         Width           =   435
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
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
         Left            =   1260
         TabIndex        =   14
         Top             =   525
         Width           =   435
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grupos de Venta según Tasa I.V.A."
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
         Height          =   285
         Index           =   4
         Left            =   105
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterios Técnicos de Clasificación"
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
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   945
         Width           =   3795
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grupos Contables de Venta"
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
         Height          =   285
         Index           =   3
         Left            =   105
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tablas de Clasificación Técnica"
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
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   315
         Width           =   3690
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPANCOACC.frx":2ED4
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "OPANCOACC.frx":3108
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu coeficientesgg 
         Caption         =   "Coeficientes de Gastos Generales"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OpAnCoAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Check1_Click()
   Option1(11).Enabled = False
   If Check1.Value = 1 Then
     Option1(11).Enabled = True
   End If
End Sub

Private Sub coeficientesgg_Click()
   Call Command7_Click
End Sub

Private Sub Command1_Click()
   '
   Call Command2_Click
   Call CIERRARCH("*.*")
   Hide
   Call FRESHALL
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
   AJUIN$ = "NO"
   If Check1.Value = 1 Then AJUIN$ = "SI"
   Call GRACONTROL$("ANACOSTO", "AJUINFLA", AJUIN$)
   '
   PORMILES$ = "NO": HABIPORMI% = 0
   If Check1.Value = 1 Then PORMILES$ = "SI": HABIPORMI% = 1
   Call GRACONTROL("AMASTO", "PORMILES", PORMILES$)
   '
   MONUME$ = "NO"
   If Check3.Value = 1 Then MONUME$ = "SI"
   Call GRACONTROL$("ANACOSTO", "PRELISTA", MONUME$)
   '
   '
   Command2.Enabled = True
End Sub

Private Sub CARINDAJU()
   Option1(11).Enabled = False
   '
   ' arma tabla indices de ajuste inflacion
   '
   Screen.MousePointer = 11
   U& = 0
   Dim MES$(12)
   '
   MES$(1) = "ENERO"
   MES$(2) = "FEBRERO"
   MES$(3) = "MARZO"
   MES$(4) = "ABRIL"
   MES$(5) = "MAYO"
   MES$(6) = "JUNIO"
   MES$(7) = "JULIO"
   MES$(8) = "AGOSTO"
   MES$(9) = "SETIEMBRE"
   MES$(10) = "OCTUBRE"
   MES$(11) = "NOVIEMBRE"
   MES$(12) = "DICIEMBRE"
   '
   HOO% = HOY(HOS$)
   MMFF% = XVALO(Mid$(HOS$, 4, 2)) + 2
   AAFF% = XVALO(Mid$(HOS$, 7, 2))
   If MMFF% > 12 Then
      MMFF% = MMFF% - 12
      AAFF% = AAFF% + 1
   End If
   '
   MMXX% = 12: AAXX% = 1
   U& = 0: UINAJU# = 1
   REBLA$ = REGBLAN$("INDIAJU")
   '
10 R2$ = TRIM$(Str$(AAXX%)): While Len(R2$) < 2: R2$ = "0" + R2$: Wend
   R1$ = MES$(MMXX%)
   '
   REMESA$ = R1$ + " " + R2$
   REMESB$ = MESANO$(REMESA$, DES%, HAS%)
   U& = U& + 1
   If U& <= 1 Then DES% = (-32767)
   If U& <= ULTREG&("INDIAJU") Then
       TTXX$ = REGLEIDO$("INDIAJU", U&)
     Else
       TTXX$ = REBLA$
       Call REPLA(TTXX$, REMESB$, 1, 16)
       Call REPLA(TTXX$, MKI$(DES%), 17, 2)
       Call REPLA(TTXX$, MKI$(HAS%), 19, 2)
       Call REPLA(TTXX$, MKD$(UINAJU#), 21, 8)
   End If
   Call GRAREG("INDIAJU", TTXX$, U&)
   MMXX% = MMXX% + 1
   If MMXX% > 12 Then
     MMXX% = MMXX% - 12
     AAXX% = AAXX% + 1
   End If
   If 13 * AAXX% + MMXX% <= 13 * AAFF% + MMFF% Then GoTo 10
   Call CIERRARCH("INDIAJU")
   Screen.MousePointer = 1
   '
   Call TRALOCAL("INDIAJU", "")
   VTB% = VENTABU%("INDIAJU")
   If VTB% >= 0 Then
     Call TRACENTRAL("INDIAJU", "")
   End If
   Option1(11).Enabled = True
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Call Command5_Click
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    For Index = 0 To 12
      If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
      End If
   Next Index
99 Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call Command5_Click
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
  '
  Command7.Enabled = False
  '1.-Selección por Grupo Contable
  Call SELECHO("GRUCOVEN")
  VALONUE$ = VALACT1$("GRUCOVEN")
  If VALONUE$ <> "" Then
  '2.- Lista de Porcentajes de Gastos Generales
    GC1% = XVALO(VALONUE$)
15  Call BLANARCH("!GAGENCOS")
    Call SETULTREG("!GAGENCOS", 0)
    Call FILTERFILE("GAGENCOS", "!GAGENCOS", 1, MKI$(GC1%))
    FIN& = ULTREG&("!GAGENCOS")
    If FIN& < 7 Then
      Call SETULTREG("!GAGENCOS", 0)
      XBLA1$ = REGBLAN$("GAGENCOS")
      X1$ = XBLA1$
      Call REPLA(X1$, MKI$(GC1%), 1, 2)
      Call REPLA(X1$, MKI$(1), 3, 2)
      Call REPLA(X1$, "G.G. Materiales", 5, 32)
      Call REPLA(X1$, MKS$(0), 37, 4)
      Call REPLA(X1$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X1$)
      '
      X2$ = XBLA1$
      Call REPLA(X2$, MKI$(GC1%), 1, 2)
      Call REPLA(X2$, MKI$(2), 3, 2)
      Call REPLA(X2$, "G.G. Elaboración", 5, 32)
      Call REPLA(X2$, MKS$(0), 37, 4)
      Call REPLA(X2$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X2$)
      '
      X3$ = XBLA1$
      Call REPLA(X3$, MKI$(GC1%), 1, 2)
      Call REPLA(X3$, MKI$(3), 3, 2)
      Call REPLA(X3$, "G.G. Administrac. y Ventas", 5, 32)
      Call REPLA(X3$, MKS$(0), 37, 4)
      Call REPLA(X3$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X3$)
      '
      X4$ = XBLA1$
      Call REPLA(X4$, MKI$(GC1%), 1, 2)
      Call REPLA(X4$, MKI$(4), 3, 2)
      Call REPLA(X4$, "Costos sobre Venta", 5, 32)
      Call REPLA(X4$, MKS$(0), 37, 4)
      Call REPLA(X4$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X4$)
      '
      X5$ = XBLA1$
      Call REPLA(X5$, MKI$(GC1%), 1, 2)
      Call REPLA(X5$, MKI$(5), 3, 2)
      Call REPLA(X5$, "Impuestos", 5, 32)
      Call REPLA(X5$, MKS$(0), 37, 4)
      Call REPLA(X5$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X5$)
      '
      X6$ = XBLA1$
      Call REPLA(X6$, MKI$(GC1%), 1, 2)
      Call REPLA(X6$, MKI$(6), 3, 2)
      Call REPLA(X6$, "Utilidad de Cálculo", 5, 32)
      Call REPLA(X6$, MKS$(0), 37, 4)
      Call REPLA(X6$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X6$)
      '
      X7$ = XBLA1$
      Call REPLA(X7$, MKI$(GC1%), 1, 2)
      Call REPLA(X7$, MKI$(7), 3, 2)
      Call REPLA(X7$, "Descuento Standard", 5, 32)
      Call REPLA(X7$, MKS$(0), 37, 4)
      Call REPLA(X7$, MKI$(0), 41, 4)
      Call REGAPP("!GAGENCOS", X7$)
      '
      Call CIERRARCH("!GAGENCOS")
      '
    End If
    '
25  ECOGC1$ = ECOARCH("GRUCOVEN", MKI$(GC1%))
    VTB1% = VENTABU%("GAGENCOS/NC/TITUPAN=Porcentajes de Gastos Generales - " & ECOGC1$)
    If VTB1% >= 0 Then
      'Validaciones
      For JX& = 1 To ULTREG&("!GAGENCOS")
        XX$ = REGLEIDO$("!GAGENCOS", JX&)
        PO1 = CVS(Mid$(XX$, 37, 4))
        MO1% = CVI(Mid$(XX$, 41, 2))
        '
        If PO1 < 0 Or PO1 > 99.9 Then
          Call MENSERR(24, "Porcentaje no Válido")
          GoTo 25
        End If
        If MO1% < 1 Or MO1% > 2 Then
          Call MENSERR(24, "Modo de Cálculo no Válido")
          GoTo 25
        End If
      Next JX&
      '3.-Guardar en Tabla GAGENCOS.DAT
      Call FILTERUPDATE("GAGENCOS", "!GAGENCOS", 1, MKI$(GC1%))
      '4.- Cierra Todos Los Archivos
      Call CIERRARCH("*.*")
    End If
  End If
  Command7.Enabled = True
End Sub

Private Sub Form_Activate()
   List1.Clear
   For U& = 1 To ULTREG&("TAMONED")
     XX$ = REGLEIDO$("TAMONED", U&)
     TTXX$ = Mid$(XX$, 2, 27) + Space$(21)
     TICAX = CVS(Mid$(XX$, 29, 4))
     Call REPLA(TTXX$, CSTRING$(MKS$(TICAX), 3, 12, 4, 2), 21, 12)
     List1.AddItem TTXX$
   Next U&
   Call CIERRARCH("TAMONED")
End Sub

'
Private Sub Form_Load()
  '
  '
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &H8000000F
      Picture16.Visible = True
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F

      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      
      For i% = 0 To Option1.UBound
         Option1(i%).BackColor = &H8000000F
      Next i%
      For i% = 0 To Option2.UBound
         Option2(i%).BackColor = &H8000000F
      Next i%
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Label1.BackColor = &H8000000F
      MARCOBARRA.Top = 4500
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
  End If

  'Call APLICACIONSKINS(Me, Picture1)
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "TACRITEC"
  TABLA$(2) = "TACLATEC"
  TABLA$(3) = "GRUCOVEN"       ' grupos contables por articulo
  TABLA$(4) = "TASIVA"
  TABLA$(5) = "TADEPOSI"
  TABLA$(6) = "SECPROD"
  TABLA$(7) = "PADMAQUI"
  TABLA$(8) = "CATOPER"
  TABLA$(9) = "UMEDIDA"
  TABLA$(10) = "ZONAREP"
  '
  DCC$ = TRIM$(CONTROL$("ANACOSTO", "DECICAL"))
  If DCC$ = "" Then
      DECICALCOS% = 4
      Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
    Else
      DECICALCOS% = XVALO(DCC$)
  End If
  If DECICALCOS% < 0 Or DECICALCOS% > 5 Then DECICALCOS% = 4
  Option2(DECICALCOS%).Value = True
  '
  Option1(11).Enabled = False
  If CONTROL$("ANACOSTO", "AJUINFLA") = "SI" Then
    Check1.Value = 1
    Option1(11).Enabled = True
  End If
  '
  If CONTROL$("AMASTO", "PORMILES") = "SI" Then
    Check1.Value = 1
  End If
  '
  Check3.Value = 0
  If CONTROL$("ANACOSTO", "PRELISTA") = "SI" Then
     Check3.Value = 1
  End If
  
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub List1_DblClick()
   Call CARTAMONED
End Sub

'
Private Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 12
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
End Sub

Sub CARTABLA(Index)
   '
   If Index = 0 Then
       Call CARTATIMA
    ElseIf Index = 1 Then
       Call CARCRITEC
    ElseIf Index = 2 Then
       Call CARTACLAS
    ElseIf Index = 3 Then
       Call CARGRUCOVE
    ElseIf Index = 4 Then
       Call CARTASIVA
    ElseIf Index = 5 Then
       Call CARTADEPOSI
    ElseIf Index = 6 Then
       Call CARSECPROD
    ElseIf Index = 7 Then
       Call CARMAQUI
    ElseIf Index = 8 Then
       Call CARCATOPER
    ElseIf Index = 9 Then
       Call CARUMEDIDA
    ElseIf Index = 10 Then
       Call CARTAPROINTER
    ElseIf Index = 11 Then
       Call CARINDAJU
    ElseIf Index = 12 Then
       Call CARTAMONED
   End If
   '
End Sub
'
Private Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
End Sub

Private Sub Option2_Click(Index As Integer)
   DECICALCOS% = Index
End Sub

Sub CARTATIMA()
    '
    Call BLANARCH("!TABTIMA")
    For i& = 1 To ULTREG&("TABTIMA")
       x$ = REGLEIDO$("TABTIMA", i&)
       If Asc(Left$(x$, 1)) > 0 Then
           Call REGAPP("!TABTIMA", x$)
       End If
    Next i&
    Call CIERRARCH("!TABTIMA")
    '
10  VTB% = VENTABU%("TABTIMA")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!TABTIMA")
       x$ = REGLEIDO$("!TABTIMA", i&)
       If Asc(Left$(x$, 1)) > 0 Then
          DPP$ = Mid$(x$, 17, 1)
          If ECOARCH$("ECODEP", DPP$) = "" Then
             Call MENSERR(24, "Depósito no Definido")
             GoTo 10
          End If
          If Asc(DPP$) = 99 Then
             Call MENSERR(24, "Depósito 99 no Válido")
             GoTo 10
          End If
          J& = J& + 1
          Call GRAREG("TABTIMA", x$, J&)
          Call GRAREG("TATIMAT", x$, J&)
       End If
    Next i&
    Call SETULTREG("TABTIMA", J&)
    Call CIERRARCH("TABTIMA")
    Call SETULTREG("TATIMAT", J&)
    Call CIERRARCH("TATIMAT")
    Call CIERRARCH("!TABTIMA")
    '
End Sub
'

Sub CARGRUCOVE()
    '
    Call BLANARCH("!GRUCOVEN")
    For i& = 1 To ULTREG&("GRUCOVEN")
        x$ = REGLEIDO$("GRUCOVEN", i&)
        If CVI(Left$(x$, 2)) > 0 Then
            Call REGAPP("!GRUCOVEN", x$)
        End If
    Next i&
    Call CIERRARCH("GRUCOVEN")
    '
    VTB% = VENTABU%("GRUCOVEN/NC/TITUPAN=Grupos Contables")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!GRUCOVEN")
        x$ = REGLEIDO$("!GRUCOVEN", i&)
        If CVI(Left$(x$, 2)) > 0 Then
            J& = J& + 1
            Call GRAREG("GRUCOVEN", x$, J&)
        End If
    Next i&
    Call SETULTREG("GRUCOVEN", J&)
    Call CIERRARCH("GRUCOVEN")
    '
End Sub

Sub CARTACLAS()
   '
10 x$ = REGSEL$("TACRICLA")
   If x$ = "" Then Exit Sub
   '
   CLASI% = Asc(Left$(x$, 1))
   CLASX$ = TRIM$(Mid$(x$, 2, 31))
   If CLASI% < 1 Or CLASI% > 8 Or CLASX$ = "" Then Exit Sub
   '
   CLASX$ = REPLACAR$(CLASX$, "/", "-")
   ALISX$ = TRIM$(Str$(CLASI%))
   While Len(ALISX$) < 3: ALISX$ = "0" + ALISX$: Wend
   ACLASI$ = "CLASI" + ALISX$
   Call COPYSTRU("CLASISTO", ACLASI$)
   '
   Call TRALOCAL(ACLASI$, "CLASISTO")
   VTB% = VENTABU%("CLASISTO/NC/TITUPAN=Clasificación Técnica: " + CLASX$)
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!CLASISTO")
         x$ = REGLEIDO$("!CLASISTO", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG(ACLASI$, x$, U&)
      Next U&
      Call SETULTREG(ACLASI$, U&)
      Call CIERRARCH(ACLASI$)
   End If
   Call CIERRARCH("*.*")
   GoTo 10
   '
End Sub

Sub CARCRITEC()
   Call TRALOCAL("TACRICLA", "")
   VTB% = VENTABU%("TACRICLA/NC/TITUPAN=Criterios Clasificación Técnica")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACRICLA")
         x$ = REGLEIDO$("!TACRICLA", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TACRICLA", x$, U&)
      Next U&
   End If
   If ULTREG&("TACRICLA") > 8 Then
      Call SETULTREG("TACRICLA", 8)
   End If
   Call CIERRARCH("TACRICLA")
End Sub

Sub CARTASIVA()
   Call TRALOCAL("GRUPIVA", "")
   VTB% = VENTABU%("GRUPIVA/NC/TITUPAN=Tasa IVA por Categoria")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUPIVA")
         x$ = REGLEIDO$("!GRUPIVA", U&)
         Call REPLA(x$, Chr$(U& - 1), 1, 1)
         Call GRAREG("GRUPIVA", x$, U&)
      Next U&
      FIN& = ULTREG&("!GRUPIVA")
      Call SETULTREG("GRUPIVA", FIN&)
   End If
   If ULTREG&("GRUPIVA") > 3 Then
      Call SETULTREG("GRUPIVA", 3)
   End If
   Call CIERRARCH("GRUPIVA")
End Sub

Sub CARTADEPOSI()
    '
    Call BLANARCH("!TADEPOSI")
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        x$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(x$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 99 Then
                 J& = J& + 1
                 y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(y$, Left$(x$, 63), 2, 63)
                 Call REGAPP("!TADEPOSI", y$)
              End If
           End If
        End If
    Next U&
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        x$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(x$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               y$ = REGBLAN$("SUCURSAL")
               Call REPLA(y$, Mid$(x$, 2, 63), 1, 63)
               Call REPLA(y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", y$, J&)
               y$ = REGBLAN$("ECODEP")
               Call REPLA(y$, Left$(x$, 21), 1, 21)
               Call GRAREG("ECODEP", y$, J&)
               Call GRAREG("TADEPOSI", x$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    y$ = REGBLAN$("ECODEP")
    Call REPLA(y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        x$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(x$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", x$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        x$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(x$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", x$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub

Sub CARCATOPER()
   '
   Call TRALOCAL("CATOPER", "CATOPER")
   VTB% = VENTABU%("CATOPER")
   If VTB% >= 0 Then
      PUCONTROL$ = String$(255, 0)
      JJ& = 0
      For U& = 1 To ULTREG&("!CATOPER")
         x$ = REGLEIDO$("!CATOPER", U&)
         COCAT% = Asc(Left$(x$, 1))
         If COCAT% > 0 Then
           If COCAT% < 256 Then
             If Mid$(PUCONTROL$, COCAT%, 1) = Chr$(0) Then
               Mid$(PUCONTROL$, COCAT%, 1) = Chr$(1)
               JJ& = JJ& + 1
               Call GRAREG("CATOPER", x$, JJ&)
             End If
           End If
         End If
      Next U&
      Call SETULTREG("CATOPER", JJ&)
      Call CIERRARCH("CATOPER")
   End If
   '
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         x$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(x$, 4)) <> "" Then
           If Len(TRIM$(Left$(x$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           y$ = Left$(x$, 2) + Mid$(x$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", x$, JJ&)
           Call GRAREG("UMEDIDA", y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub

Sub CARTAMONED()
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
   '
   List1.Clear
   For U& = 1 To ULTREG&("TAMONED")
     XX$ = REGLEIDO$("TAMONED", U&)
     TTXX$ = Mid$(XX$, 2, 27) + Space$(21)
     TICAX = CVS(Mid$(XX$, 29, 4))
     Call REPLA(TTXX$, CSTRING$(MKS$(TICAX), 3, 12, 4, 2), 21, 12)
     List1.AddItem TTXX$
   Next U&
   Call CIERRARCH("TAMONED")
End Sub

Sub CARTAPROINTER()
   Call TRALOCAL("TAPROINT", "")
   VTB% = VENTABU%("TAPROINT/NC/TITUPAN=Tabla de Proveedores Internos")
   If VTB% >= 0 Then
      J& = 0
      For U& = 1 To ULTREG&("!TAPROINT")
         x$ = REGLEIDO$("!TAPROINT", U&)
         If CVI(Left$(x$, 2)) > 0 Then
           J& = J& + 1
           Call GRAREG("TAPROINT", x$, J&)
         End If
      Next U&
      Call SETULTREG("TAPROINT", J&)
      Call CIERRARCH("TAPROINT")
   End If
End Sub


Sub CARMAQUI()
    '
    J& = 0
    For U& = 1 To ULTREG&("PADMAQ")
        x$ = REGLEIDO$("PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(x$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("!PADMAQ", x$, J&)
        End If
    Next U&
    Call SETULTREG("!PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
    VTB% = VENTABU%("CARMAQUI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!PADMAQ")
        x$ = REGLEIDO$("!PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(x$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("PADMAQ", x$, J&)
        End If
    Next U&
    '
    Call SETULTREG("PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
End Sub






Private Sub salir_Click()
    Call Command1_Click
End Sub

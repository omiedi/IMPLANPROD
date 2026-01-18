VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPPRES07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento - Presupuestos"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11175
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
   ScaleHeight     =   5145
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Rutas de Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5130
      TabIndex        =   63
      Top             =   4080
      Width           =   5055
      Begin VB.TextBox TXTRUTA2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Text            =   "OPPRES07.frx":0000
         Top             =   180
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectividad de Cotizaciones vs. Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   65
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Rutas de Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   60
      Top             =   4080
      Width           =   5055
      Begin VB.TextBox TXTRUTA1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Text            =   "OPPRES07.frx":0002
         Top             =   180
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Items sugeridos para ingresar a stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   61
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "INICIALIZACION DE FACTURA PROFORMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5130
      TabIndex        =   56
      Top             =   3200
      Width           =   5055
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
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
         Left            =   4470
         TabIndex        =   58
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A partir del N°."
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
         Left            =   1320
         TabIndex        =   57
         Top             =   315
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PUERTO DE IMPRESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5130
      TabIndex        =   42
      Top             =   0
      Width           =   5055
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   210
         Width           =   2685
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pto.Impresión:"
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
         Left            =   315
         TabIndex        =   44
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   480
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles Operativos - Campos Mandatorios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   4950
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plazo de Garantía"
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
         Index           =   8
         Left            =   2900
         TabIndex        =   26
         Top             =   1365
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Validez de Oferta"
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
         Index           =   7
         Left            =   2900
         TabIndex        =   25
         Top             =   1050
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lugar de Entrega"
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
         Index           =   6
         Left            =   2900
         TabIndex        =   24
         Top             =   735
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plazo de Entrega"
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
         Index           =   5
         Left            =   2900
         TabIndex        =   23
         Top             =   420
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vendedor"
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
         Index           =   4
         Left            =   350
         TabIndex        =   22
         Top             =   1680
         Width           =   2250
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Condición de Pago"
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
         Index           =   3
         Left            =   350
         TabIndex        =   21
         Top             =   1365
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lista de Precios"
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
         Index           =   2
         Left            =   350
         TabIndex        =   20
         Top             =   1050
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Referencia"
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
         Index           =   1
         Left            =   350
         TabIndex        =   19
         Top             =   735
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Atención Sr./Sra./Srta"
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
         Index           =   0
         Left            =   350
         TabIndex        =   18
         Top             =   420
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FORMULARIO PRESUPUESTO - COTIZACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4950
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
         Height          =   330
         Left            =   2415
         TabIndex        =   53
         Text            =   " "
         Top             =   1870
         Width           =   1380
      End
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
         Height          =   360
         Left            =   3780
         TabIndex        =   52
         Top             =   1870
         Width           =   175
      End
      Begin VB.TextBox Text21 
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
         Height          =   330
         Left            =   210
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2250
         Width           =   3750
      End
      Begin VB.CommandButton Command60 
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
         Height          =   680
         Left            =   4095
         Picture         =   "OPPRES07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1870
         Width           =   645
      End
      Begin VB.CommandButton Command61 
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
         Height          =   680
         Left            =   4080
         Picture         =   "OPPRES07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2640
         Width           =   645
      End
      Begin VB.TextBox Text22 
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
         Height          =   330
         Left            =   210
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3010
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
         Height          =   360
         Left            =   3780
         TabIndex        =   47
         Top             =   2640
         Width           =   175
      End
      Begin VB.TextBox Text12 
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
         Height          =   330
         Left            =   2415
         TabIndex        =   46
         Text            =   " "
         Top             =   2640
         Width           =   1380
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
         Height          =   330
         Index           =   2
         Left            =   2310
         TabIndex        =   38
         Text            =   " "
         Top             =   3585
         Width           =   2430
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Incluir Especificaciones"
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
         Height          =   465
         Left            =   210
         TabIndex        =   36
         Top             =   3465
         Width           =   1920
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
         Height          =   330
         Index           =   1
         Left            =   2415
         TabIndex        =   30
         Text            =   " "
         Top             =   1050
         Width           =   1380
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
         Height          =   360
         Index           =   1
         Left            =   3780
         TabIndex        =   29
         Top             =   1050
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1425
         Width           =   3750
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
         Height          =   680
         Index           =   1
         Left            =   4095
         Picture         =   "OPPRES07.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1050
         Width           =   645
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
         Height          =   680
         Index           =   0
         Left            =   4095
         Picture         =   "OPPRES07.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   265
         Width           =   645
      End
      Begin VB.TextBox Text2 
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
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   660
         Width           =   3750
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
         Height          =   360
         Index           =   0
         Left            =   3780
         TabIndex        =   8
         Top             =   265
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
         Height          =   330
         Index           =   0
         Left            =   2415
         TabIndex        =   7
         Text            =   " "
         Top             =   265
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proforma Item Único"
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
         Index           =   4
         Left            =   0
         TabIndex        =   55
         Top             =   2000
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proforma Lista Repuestos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   54
         Top             =   2780
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía"
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
         Index           =   2
         Left            =   2160
         TabIndex        =   39
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Lista de  Repuestos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   31
         Top             =   1180
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unico"
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
         Left            =   0
         TabIndex        =   11
         Top             =   400
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ANEXOS - DOCUMENTOS GRAFICOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   5130
      TabIndex        =   0
      Top             =   675
      Width           =   5055
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Control de Impresion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   13
         Top             =   1140
         Width           =   4530
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Fotografía - 2"
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
            Index           =   3
            Left            =   2520
            TabIndex        =   17
            Top             =   630
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Fotografía - 1"
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
            Index           =   2
            Left            =   2520
            TabIndex        =   16
            Top             =   315
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Croquis"
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
            Index           =   1
            Left            =   525
            TabIndex        =   15
            Top             =   630
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Plano de Pieza"
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
            Index           =   0
            Left            =   525
            TabIndex        =   14
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.TextBox Text5 
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
         Height          =   330
         Left            =   2415
         TabIndex        =   5
         Text            =   " "
         Top             =   220
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
         TabIndex        =   4
         Top             =   220
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
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   3750
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
         Height          =   720
         Left            =   4095
         Picture         =   "OPPRES07.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   195
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
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
         TabIndex        =   1
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   10275
      ScaleHeight     =   7470
      ScaleWidth      =   1140
      TabIndex        =   32
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command44 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "OPPRES07.frx":0F36
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Tabla de Textos de Observaciones"
         Top             =   2400
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "OPPRES07.frx":1240
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Lugares de Entrega"
         Top             =   1680
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
         Height          =   640
         Left            =   105
         Picture         =   "OPPRES07.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   315
         Width           =   650
      End
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
         Height          =   660
         Left            =   105
         Picture         =   "OPPRES07.frx":1694
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   955
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
         Height          =   640
         Left            =   120
         Picture         =   "OPPRES07.frx":199E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3240
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPPRES07.frx":1CA8
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPPRES07.frx":1EDC
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "OPPRES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text12.SetFocus
      Else
        Text12.TEXT = CODFOR$
    End If

End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    If DERACCE%("INIPROFO", "Inicialización de Numeración de Factura Proforma") >= 2 Then
        Call SELECHO("INDIPRES")
        NPX& = XVALO(VALACT1$("INDIPRES"))
        If NPX& < 1 Then GoTo 99
        Text7 = NPX&
    End If
99  Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORPRESU", Text1(0).TEXT)
    Call GRACONTROL("", "FORPRERE", Text1(1).TEXT)
    Call GRACONTROL("", "FORGRAFI", Text5.TEXT)
    Call GRACONTROL("", "GARANDEF", Text1(2).TEXT)
    '
    For KKI% = 0 To 3
      MONUME$ = "NO": If Check1(KKI%).Value > 0 Then MONUME$ = "SI"
      Call GRACONTROL("PRESUPUE", "PRIGRA-" + TRIM$(Str$(KKI%)), MONUME$)
    Next KKI%
    '
    For KKI% = 0 To 8
      MONUME$ = "NO": If Check2(KKI%).Value > 0 Then MONUME$ = "SI"
      Call GRACONTROL("PRESUPUE", "CTRLX-" + TRIM$(Str$(KKI%)), MONUME$)
    Next KKI%
    '
    Call GRACONTROL("PRESIUNI", "PROFLREP", TRIM$(Text11))
    Call GRACONTROL("PRESUPUE", "PROFLREP", TRIM$(Text12))
    Call GRACONTROL("PRESUPUE", "INIPROFO", TRIM$(Text7))
    VALORX$ = "NO"
    If Check10.Value > 0 Then VALORX$ = "SI"
    Call GRACONTROL("PRESUVEN", "INCLUSPE", VALORX$)
    '
    Call GRACONTROL(IDENTER$, "PORTPRES", Combo2.TEXT)
    Call GRACONTROL("*", "PORTPRES", Combo2.TEXT)
    
    Call GRACONTROL("PRESUAHP", "RUTAEXC1", TXTRUTA1)
    Call GRACONTROL("PRESUAHP", "RUTAEXC2", TXTRUTA2)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        Text2(Index).TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command44_Click()
   Command44.Enabled = False
   Call CREACAMPO("", "TAB_TEXTO", "CODINT_OBS", 3, 0, 1)
   Call CREACAMPO("", "TAB_TEXTO", "TITU_OBS", 10, 64)
   Call CREACAMPO("", "TAB_TEXTO", "TEXTO_OBS", 12, 0)
   Call TRALOCAL("TEXTOBSE", "")
2  OPX% = COMALTER%("Carga de Textos Para Observaciones\\Ingresar Nuevo Texto\Editar Texto Existente\Eliminar")
   
   If OPX% = 1 Then ' NUEVO TEXTO DE OBSERVACION
5     TEXTO_TITULO$ = TRIM$(InputBox("Título", "Ingreso Título de Texto de Observaciones", TITUTEXT$))
      If TEXTO_TITULO$ = "" Then GoTo 99
      If CantRegistros&("TAB_TEXTO", "TITU_OBS  ='" & TEXTO_TITULO$ & "'", "NOMESS") > 0 Then
         Call COMUNI("Título Existente")
         GoTo 5
      End If
      '
      Call SETULTREG("!TEXTOBSE", 0)
8     VTB% = VENTABU%("TEXTOBSE/TITUPAN=" + TEXTO_TITULO$)
      If VTB% >= 0 Then
        EPAC$ = TRIM$(EMPREAC$)
        If InStr(UCase$(EPAC$), "AHP") > 0 Then
            CODFOR$ = TRIM$(Control$("", "FORPRERE"))
            If CODFOR$ = "PRRS-AHP" Then
                FORMATOX$ = TRIM$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO"))
                POS1% = InStr(FORMATOX$, "/")
                SEPASO% = 0
                If POS1% > 0 Then
                   CANTIRECORD% = XVALO(Mid$(FORMATOX$, POS1% + 1))
                End If
                TEXTX$ = ""
                For U& = 1 To ULTREG&("!TEXTOBSE")
                  XX$ = REGLEIDO$("!TEXTOBSE", U&)
                  XX1$ = Mid$(XX$, 3, 62)
                  TEXTX$ = TEXTX$ + XX1$ + vbCrLf
                  If CANTIRECORD% < U& Then SEPASO% = 1
                Next U&
                If SEPASO% > 0 Then
                   OPX% = COMALTER%("El Texto Ingresado Supera el Máximo de Renglones\Configurado Para las Observaciones (" & CANTIRECORD% & ")\\Modificar\Continuar")
                   If OPX% = 1 Then GoTo 8
                End If
            End If
        End If
        '
10      RNDX$ = RANDSTRING$(10)
        CONDI$ = "TITU_OBS = '" & RNDX$ & "'"
        If CantRegistros("TAB_TEXTO", "TITU_OBS = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
        Call CreaRegistro("TAB_TEXTO", "TITU_OBS", RNDX$, "NOMESS")
          PROXIMOOBS% = 1 + XVALO(VALOBADA("TAB_TEXTO", "MAX(CODINT_OBS)", "CODINT_OBS > 0", "NOMESS"))
        Call ACTUREGISTRO("TAB_TEXTO", "CODINT_OBS", CONDI$, PROXIMOOBS%, REGAF&, "NOMESS")
        Call ACTUREGISTRO("TAB_TEXTO", "TEXTO_OBS", CONDI$, TEXTX$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("TAB_TEXTO", "TITU_OBS", CONDI$, TEXTO_TITULO$, REGAF&, "NOMESS")
        GoTo 2
      End If
      '
    ElseIf OPX% = 2 Then
      'EDICION DE OBSERVACION EXISTENTE
      CONDI$ = "CODINT_OBS > 0 ORDER BY CODINT_OBS"
      Call CARGALISEL("LISTTITU", "TAB_TEXTO", "CODINT_OBS/F5;TITU_OBS/A48", CONDI$, "NOMESS/DISTINCT")
      Call SELECHO("LISTTITU/TITUTLOS DE TEXTO DE OBSERVACIONES")
      If VALACT1$("LISTTITU") = "" Then GoTo 2 ' COMALTER
      '
      CODITITU% = XVALO(VALACT1$("LISTTITU"))
      CONDI$ = "CODINT_OBS = " & CODITITU%
      textox$ = VALOBADA("TAB_TEXTO", "TEXTO_OBS", CONDI$, "NOMESS")
      'SEPARO EL TEXTO SEGUN LOS GOLPES DE ENTER
      X = TEXTOLOTES$(textox$, Chr(13))
      TXX1$ = Space(62)
      Call SETULTREG("!TEXTOBSE", 0)
      For J% = 1 To UBound(CADENATEX$)
         TROZO$ = CADENATEX$(J%)
         TXX1$ = TROZO$
         Call REPLA(TXX1$, MKI$(J%), 1, 2)
         Call REPLA(TXX1$, TROZO$, 3, 62)
         Call REGAPP("!TEXTOBSE", TXX1$)
      Next J%
      VTB% = VENTABU%("TEXTOBSE/TITUPAN=" + VALACT2$("LISTTITU"))
      TXXT$ = ""
      If VTB% >= 0 Then
        For i& = 1 To ULTREG&("!TEXTOBSE")
          XX$ = REGLEIDO$("!TEXTOBSE", i&)
          TXX1$ = TRIM$(Mid$(XX$, 3, 62))
          TXXT$ = TXXT$ & TXX1$ + vbCrLf
        Next i&
        Call ACTUREGISTRO("TAB_TEXTO", "TEXTO_OBS", CONDI$, TXXT$, REGAF&, "NOMESS")
        GoTo 2
      End If
   ElseIf OPX% = 3 Then
      'ELIMINAR COMPLETAMENTE UN TEXTO DE OBSERVACION EXISTENTE.
      CONDI$ = "CODINT_OBS > 0 ORDER BY CODINT_OBS"
      Call CARGALISEL("LISTTITU", "TAB_TEXTO", "CODINT_OBS/F5;TITU_OBS/A48", CONDI$, "NOMESS/DISTINCT")
      Call SELECHO("LISTTITU/TITULO DE TEXTO DE OBSERVACIONES")
      If VALACT1$("LISTTITU") = "" Then GoTo 2 ' COMALTER
      '
      CODITITU% = XVALO(VALACT1$("LISTTITU"))
      Call BORRAREGISTROS("TAB_TEXTO", "CODINT_OBS = " & CODITITU%, REGAF&, "NOMESS")
      Call COMUNI("Texto de Observación Seleccionado 'Eliminado'")
      GoTo 2
   End If
   '
   Call TRACENTRAL("TEXTOBSE", "")
   '
99 Command44.Enabled = True
End Sub





Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text5.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click(Index As Integer)
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1(Index).TEXT), "Formularios de Impresion")
End Sub

Private Sub Command60_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text11.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command61_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text12.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command8_Click()
    Call TRALOCAL("PRELUENT", "")
    
    VTB% = VENTABU%("PRELUENT")
    If VTB% < 0 Then Exit Sub
    
    Call TRACENTRAL("PRELUENT", "")
End Sub

Private Sub Command9_Click()
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text11.SetFocus
      Else
        Text11.TEXT = CODFOR$
    End If
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
    Text1(0).TEXT = Control$("", "FORPRESU")
    Text2(0).TEXT = ECOARCH$("INDIFRM", Text1(0).TEXT)
    Text1(1).TEXT = Control$("", "FORPRERE")
    Text2(1).TEXT = ECOARCH$("INDIFRM", Text1(1).TEXT)
    Text1(2).TEXT = Control$("", "GARANDEF")
    Text5.TEXT = Control$("", "FORGRAFI")
    Text6.TEXT = ECOARCH$("INDIFRM", Text5.TEXT)
    '
    For KKI% = 0 To 3
      Check1(KKI%).Value = 0
      If Control$("PRESUPUE", "PRIGRA-" + TRIM$(Str$(KKI%))) = "SI" Then
        Check1(KKI%).Value = 1
      End If
    Next KKI%
    '
    For KKI% = 0 To 8
      Check2(KKI%).Value = 0
      If Control$("PRESUPUE", "CTRLX-" + TRIM$(Str$(KKI%))) = "SI" Then
        Check2(KKI%).Value = 1
      End If
    Next KKI%
    '
    Check10.Value = 0
    If Control("PRESUVEN", "INCLUSPE") = "SI" Then
      Check10.Value = 1
    End If
    '
    Text11 = Control("PRESIUNI", "PROFLREP")
    Text12 = Control("PRESUPUE", "PROFLREP")
    Text7 = Control("PRESUPUE", "INIPROFO")
    '
    TXTRUTA1 = Control("PRESUAHP", "RUTAEXC1")
    TXTRUTA2 = Control("PRESUAHP", "RUTAEXC2")

    Dim PRX As Printer
    PPRES$ = Control$(IDENTER$, "PORTPRES")
      If PPRES$ = "" Then PPRES$ = Control$("*", "PORTPRES")
    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PPRES$ Then
        Combo2.TEXT = PPRES$
      End If
    Next
    '
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        Text2(Index).TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text11_Change()
    CODFOR$ = TRIM$(Text11)
    Text21 = ECOARCH$("INDISQR", CODFOR$)
End Sub

Private Sub Text12_Change()
    CODFOR$ = TRIM$(Text12)
    Text22 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub Text2_CLICK(Index As Integer)
    Text1(Index).SetFocus
End Sub

Private Sub TXTRUTA2_DblClick()
 TXTRUTA2 = SELECCION_Archivo$("*", CMD1)
End Sub

Private Sub Text6_Click()
    Text5.SetFocus
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Text7_DblClick()
   If DERACCE%("INIPROFO", "Inicialización de Numeración de Factura Proforma") >= 2 Then
      Text7 = ""
   End If
End Sub


Private Sub TXTRUTA1_DblClick()
 TXTRUTA1 = SELECCION_Archivo$("*", CMD1)
End Sub


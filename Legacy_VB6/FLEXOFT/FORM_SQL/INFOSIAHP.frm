VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFOSIAHP 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE SITUACIÓN DE STOCKS DE ITEM ACTIVO"
   ClientHeight    =   8700
   ClientLeft      =   -3150
   ClientTop       =   540
   ClientWidth     =   15210
   ControlBox      =   0   'False
   DrawWidth       =   10
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15210
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton Command3 
         Caption         =   "Ing.Sin Stock"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text40 
         Appearance      =   0  'Flat
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
         Left            =   6515
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Vía de Entr."
         Top             =   960
         Width           =   2625
      End
      Begin VB.TextBox TXTVIAENTREGA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   2385
      End
      Begin VB.TextBox TXTCOSTO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   4920
         TabIndex        =   2
         ToolTipText     =   "Costo de Maestro de Articulos"
         Top             =   1200
         Width           =   1590
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Proveedor:"
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox TXTNUMEPROVE 
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "F2 - Para Desplegar Lista de Proveedores"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Lista de Proveedores Relacionados con el Item Activo"
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Código Cliente:"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtcodigo 
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   510
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   240
         Width           =   3600
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   510
         Width           =   3600
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Cant. Solicitada:"
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TXTSOLICITADA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   380
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   510
         Width           =   1560
      End
      Begin VB.TextBox TXTSERVTECNICO 
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
         Height          =   1305
         Left            =   10560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         ToolTipText     =   "Especificación Del Maestro de Artículos"
         Top             =   360
         Width           =   4380
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Texto Técnico:"
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TXTPESO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   380
         Left            =   7215
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   510
         Width           =   1440
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   7215
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Peso u. Kg.:"
         Top             =   240
         Width           =   1440
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir sin Modificar"
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
         Left            =   9240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1215
      End
      Begin VB.TextBox TXTCODICLIENTE 
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   6150
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Combo6"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5160
         TabIndex        =   51
         Top             =   1630
         Width           =   800
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tot.Disponible.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3720
         TabIndex        =   50
         Top             =   1635
         Width           =   1420
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido/s:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5955
         TabIndex        =   49
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   48
         Top             =   1635
         Width           =   8100
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2800
         TabIndex        =   47
         Top             =   1630
         Width           =   900
      End
      Begin VB.Label lblAlertaReservaSinLote 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad Reservada Sin Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1635
         Width           =   2700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4440
      Left            =   0
      TabIndex        =   12
      Top             =   4560
      Width           =   15135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   " F2 Despliega Lista de Proveedores      -     F5 Aprueba Ingreso     - "
         Top             =   4080
         Width           =   14865
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cotizador"
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
         Left            =   8820
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ingreso con Stock"
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox TXTSELECCIONADA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Cant.Selecc.:"
         Top             =   240
         Width           =   1440
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "En transito (Otros):"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TXTTRANSITO_OTROS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox Text36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Últimas 3 Cotizaciones"
         Top             =   600
         Width           =   9750
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Últimas 3 Ventas"
         Top             =   2400
         Width           =   9750
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Estadisticas de Ventas"
         ToolTipText     =   "CANTIDAD DE UNIDADES VENDIDAS (DESCUENTA NOTA DE CRÉDITO) "
         Top             =   2400
         Width           =   5025
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Estadísticas de Cotizaciones"
         Top             =   600
         Width           =   5025
      End
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   1395
         Left            =   9960
         TabIndex        =   21
         Top             =   840
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   2461
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
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1395
         Left            =   9960
         TabIndex        =   22
         Top             =   2640
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   2461
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   8421504
         AllowUserResizing=   3
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1395
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   2461
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   16501405
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   1395
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   2461
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   16501405
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selección "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   15135
      Begin VB.TextBox TXTMOVIL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2355
         Left            =   120
         TabIndex        =   10
         Top             =   200
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   4154
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   16501405
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   360
      Top             =   630
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "INFOSIAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODO1$
Public FILAX&
Public PRESENTA_COTIZADOR%
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Sub ACTUALIZAPESO_ENMAESTRO(CI1%, PESO#)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    CONDI$ = "CODINT =  " & CI1%
    Call ACTUREGISTRO("MASTER", "PESUNI", CONDI$, FORMATNUM$(PESO#, "F14.4"), REGAF&, "NOMESS")
    If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
    End If

End Sub


Sub EstadoVisibleReservasSinLote(visibleOfalse As Boolean)
   lblAlertaReservaSinLote.Visible = visibleOfalse
   Label2.Visible = visibleOfalse
   Label3.Visible = visibleOfalse
   Label4.Visible = visibleOfalse
   Label5.Visible = visibleOfalse
   Label6.Visible = visibleOfalse
   If visibleOfalse = False Then
     Label2 = "": Label3 = "":   Label6 = ""
   End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
   Call TXTVIAENTREGA_KeyPress(KeyAscii)
End Sub

'1.CODIGO
'2.CODIGO CLIENTE
'3.DESCRIPCION
'4.CANTIDAD
'5.NUMERO PROVEEDOR ' OCULTO
'6.R.S.PROVEEDOR
'7.COSTO
'8.PR.VENTA
'9.IMPORTE
'10.VIA ENTREGA
'11.LOTES OCULTO
'12.CANTIDAD DE LOTES OCULTO
'13.NUMERO DE PROVEEDOR OCULTO
'14.KILOS OCULTO
'15.DEMORA OCULTO
Private Sub Command1_Click()
    Command1.Enabled = False
    Timer1.Enabled = False
    '
    If TRIM$(txtcodigo.TEXT) = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 90
    End If
    '
    CI1% = CODINT%(txtcodigo.TEXT)
    CANTSELEC = XVALO(TXTSELECCIONADA)
    NECESIDAD = XVALO(TXTSOLICITADA)
    If CANTSELEC > NECESIDAD Then
       Call COMUNI("Cantidad Seleccionada Supera a la Cantidad a Cotizar")
       GoTo 99
    End If
    '
    If NECESIDAD > CANTSELEC Then
       Call COMUNI("Cantidad Seleccionada No Alcanza Para la Cantidad a Cotizar\Modifique Cantidad")
       GoTo 99
    End If
    '
    If XVALO(Label2) > 0 Then
        Dispo = 0: CANTI = 0
        For i& = 1 To MSF1.Rows - 1
            Dispo = Dispo + XVALO(MSF1.TextMatrix(i&, 12))
            CANTI = CANTI + XVALO(MSF1.TextMatrix(i&, 11))
        Next i&
        If CANTI > (Dispo - XVALO(Label2)) Then
          Call COMUNI("Cantidades Ingresadas, menos Cantidades Reservadas Sin Lote, Superan la Disponibilidad")
          GoTo 99
        End If
    End If
    
    'VALIDO QUE NO SE HAYA SELECCIONADO MAS CANTIDAD DE LOS LOTES DISPONIBLES
    If FILAX& = 0 Then
      Call CONTROL_CANT_LOTE_MSF1(OK%)
      If OK% < 1 Then GoTo 99
    End If
    '
    If XVALO(TXTSELECCIONADA) < 1 Then
      Call COMUNI("Falta Seleccionar Cantidad de Lote")
      GoTo 99
    End If
    '
    EDICION% = 0
    '
    Call ACTUALIZAPESO_ENMAESTRO(CI1%, XVALO(TXTPESO))

    For i& = 1 To MSF1.Rows - 1
     CANTI = XVALO(MSF1.TextMatrix(i&, 11))
     If CANTI > 0 Then
        FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
        FRMPRESUAHP.Text15 = FORMATNUM$(XVALO(MSF1.TextMatrix(i&, 8)), "F12.2")
        '
        FRMPRESUAHP.Text38 = "Stock Propio" ' MSF1.TextMatrix(I&, 2)
'          LOTE$ = MSF1.TextMatrix(I&, 13)
'        FRMPRESUAHP.TXTPREUNIT = FRMPRESUAHP.PRECIOVENTA_X_FORMULA(XVALO(FRMPRESUAHP.Text15), ViaEntrega(LOTE$), CI1%, 1)
        FRMPRESUAHP.TXTPREUNIT = MSF1.TextMatrix(i&, 10)

        FRMPRESUAHP.TXTCANTIDAD = TRIM$(FORMATNUM$(CANTI, "F9.1"))
        FRMPRESUAHP.TXTPESOP = TRIM$(FORMATNUM$(XVALO(TXTPESO), "F9.2"))
        Call FRMPRESUAHP.cargar_datos ' cargo los datos para crear el nuevo registro en la pantalla de presupuesto
          NPROV = XVALO(MSF1.TextMatrix(i&, 14))
          'EL NUMERO DEL PROVEEDOR SE ESTA GUARDANDO EN DOS CELDAS DIFERENTES !!!HAY QUE CONTROLARLO!!!
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 5) = NPROV
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 13) = NPROV
        
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 6) = TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 11) = MSF1.TextMatrix(i&, 13)
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 14) = XVALO(TXTPESO) * CANTI
        
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 15) = "Inmediata" 'DEBERIA SER SIEMPRE INMEDIATA POR QUE AQUI ES CON STOCKS.
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 16) = MSF1.TextMatrix(i&, 7) 'valor fob
     End If
    Next i&
    '
    Call FRMPRESUAHP.ACTUPESO
    Call FRMPRESUAHP.COLOREARSTOCKS(FRMPRESUAHP.MSF1.Rows - 1)

    MODO1$ = "ACTUALIZAR"
90  Hide
    
99  Command1.Enabled = True
End Sub



Sub ACTUCANTI()
   ACUMULACANTI = 0
   'If TXTMOVIL.Visible = True Then Exit Sub
   For i& = 1 To MSF1.Rows - 1
        CANTI = XVALO(MSF1.TextMatrix(i&, 11))
        ACUMULACANTI = ACUMULACANTI + CANTI
   Next i&
   TXTSELECCIONADA = FORMATNUM$(ACUMULACANTI, "F10.1")
End Sub
Sub CONTROL_CANT_LOTE_MSF1(OK%)
   OK% = 1
   CI1% = CODINT(txtcodigo)
   For i& = 1 To MSF1.Rows - 1
        LOTE$ = MSF1.TextMatrix(i&, 13)
        ST = XVALO(MSF1.TextMatrix(i&, 12))
        CANTI = XVALO(MSF1.TextMatrix(i&, 11))
        '
        If CANTI > ST Then
             Call COMUNI("Cantidad Seleccionada para Lote: " & LOTE$ & " Supera el Stock del Mismo.")
             OK% = 0
             Exit For
        End If
        '
        'CONSIDERAR TODO ESTE IF DE AQUI ABAJO POR QUE NO TIENE DEMASIADO SENTIDO YA QUE LA VARIABLE ST
        'TRAE EL DISPONIBLE YA DESCONTANDO SI ANTERIORMENTE SE HABIA CARGADO MISMO CODIGO Y LOTE??
        'LO DEJE POR QUE  TIENE EL MENSAJE CON EL DETALLE EXPLICANDO LA SITUACION.
        
        'VALIDA SI HAY CANTIDAD YA SELECCIONADA EN EL PRESUPUESTO + LA SELECCIONADA SI ALCANZA EL STOCKS
        CANTIGRI = CANTILOTE_EN_GRILLA(CI1%, LOTE$) ' CANTIDAD YA INGRESADA EN GRILLA DE PRESUPUESTO DEL MISMO LOTE
        If CANTIGRI > 0 Then
           If (CANTI + CANTIGRI) > (ST + CANTIGRI) Then
              TX$ = "La Cantidad (" & CANTI & ") Seleccionada para Lote: " & LOTE$
              TX$ = TX$ + "\Sumada a La Cantidad (" & CANTIGRI & ") Ya ingresada en Detalle de Presupuesto Activo"
              TX$ = TX$ + "\Supera La Cantidad (" & ST & ") del Stock Disponible."
              Call COMUNI(TX$)
              OK% = 0
              Exit For
           End If
        End If
   Next i&
End Sub

Function CANTILOTE_EN_GRILLA(CI1%, LOTEPARAM$)
    ACUCANTI = 0
    With FRMPRESUAHP
        LOTEPARAM$ = TRIM$(LOTEPARAM$)
        For i& = 1 To .MSF1.Rows - 1
           STRINGLOTES$ = .MSF1.TextMatrix(i&, 11)
           If InStr(STRINGLOTES$, LOTEPARAM$) > 0 Then
                  CANTI = XVALO(.MSF1.TextMatrix(i&, 4))
                  ACUCANTI = ACUCANTI + CANTI
           End If
        Next i&
    End With
    '
    CANTILOTE_EN_GRILLA = ACUCANTI
    '
End Function

Private Sub Command1_GotFocus()
 Command1.BackColor = vbYellow
End Sub

Private Sub Command1_LostFocus()
  Command1.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
   FRMPRESUAHP.EDICION = 0
   MODO1$ = "CANCEL"
   On Error Resume Next
   Timer1.Enabled = False
   Hide
   On Error GoTo 0
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
     CANTI = XVALO(TXTSOLICITADA)
     '

     If CANTI > 0 Then
        FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
        FRMPRESUAHP.Text15 = FORMATNUM$(XVALO(TXTCOSTO), "F12.2")
            CI1% = CODINT%(txtcodigo)
         Call ACTUALIZAPESO_ENMAESTRO(CI1%, XVALO(TXTPESO))
           
        FRMPRESUAHP.TXTPREUNIT = PRECIOVENTA_X_FORMULA(XVALO(FRMPRESUAHP.Text15), XVALO(Left$(TXTVIAENTREGA, 2)), CI1%, 1, XVALO(Me.TXTNUMEPROVE), 1)
        '
        '
        FRMPRESUAHP.Text38 = TXTVIAENTREGA
        'FRMPRESUAHP.Text38 = FRMPRESUAHP.DEMORA_X_VIA(XVALO(Left$(TXTVIAENTREGA, 2)), XVALO(TXTNUMEPROVE)) & "-Días"
        FRMPRESUAHP.TXTCANTIDAD = TRIM$(FORMATNUM$(XVALO(TXTSOLICITADA), "F9.1"))
        FRMPRESUAHP.TXTPESOP = TRIM$(FORMATNUM$(XVALO(TXTPESO), "F9.2"))
        NPROV = XVALO(TXTNUMEPROVE)
        FRMPRESUAHP.TXTNUMPROV = NPROV 'TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor

        Call FRMPRESUAHP.cargar_datos_sin_stock ' cargo los datos para crear el nuevo registro en la pantalla de presupuesto
          'EL NUMERO DEL PROVEEDOR SE ESTA GUARDANDO EN DOS CELDAS DIFERENTES !!!HAY QUE CONTROLARLO!!!
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 5) = NPROV
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 13) = NPROV
        
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 6) = TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 11) = ""
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 14) = CANTI * XVALO(TXTPESO)
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 15) = FRMPRESUAHP.DEMORA_X_VIA(XVALO(Left$(TXTVIAENTREGA, 2)), XVALO(TXTNUMEPROVE))
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 16) = TXTCOSTO 'valor fob INGRESADO MANUALMENTE
        Hide
        Call FRMPRESUAHP.COLOREARSTOCKS(FRMPRESUAHP.MSF1.Rows - 1)
    
     End If
    Command3.Enabled = True
End Sub

Private Sub Command3_GotFocus()
    Command3.BackColor = vbYellow
End Sub


Private Sub Command3_LostFocus()
   Command3.BackColor = &H8000000F
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Timer1.Enabled = False
   PRESENTA_COTIZADOR% = 1
   
   Call Command2_Click
   Command4.Enabled = True

End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
           Call SELECHO("PROVED1")
           '
           NPRO% = XVALO(VALACT1("PROVED1"))
           If NPRO% < 1 Then GoTo 99
           TXTNUMEPROVE = NPRO%
           On Error Resume Next
           TXTNUMEPROVE.SetFocus
           On Error GoTo 0
99   Command5.Enabled = True
End Sub

Private Sub Form_Activate()
  Call CARGARCOMBOVIAS_TODOS
  If MSF1.Rows > 1 Then
    MSF1.COL = 11
    MSF1.Row = 1
    'MSF1.SetFocus
    With TXTMOVIL
      .ZOrder 0
      .Width = MSF1.CellWidth
      .Left = MSF1.CellLeft + MSF1.Left
      .Top = MSF1.CellTop + MSF1.Top
      .TEXT = MSF1.TEXT
      .Visible = True
      .ToolTipText = MSF1.Row
      .SetFocus
    End With
  End If
  
  If TXTNUMEPROVE.Visible = True Then
     Call PINTATEXT(TXTPESO)
     Command3.BackColor = &H8000000F
  End If
  TXTNUMEPROVE.TEXT = ""
  TXTVIAENTREGA.TEXT = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
       If Command1.Visible = True Then
          Command1_Click
       Else
          Command3_Click
       End If
    End If
    If KeyCode = vbKeyF2 Then
      If Command1.Visible = True Then
          Command5_Click
      End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Command2_Click
    End If
End Sub


Sub MSF1_Click()
          COLUM = MSF1.MouseCol
          If COLUM <> 11 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF1.COL = 11
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            .TEXT = MSF1.TEXT
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
          Call ACTUCANTI
End Sub

Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
    IDLOT$ = TRIM$(MSF1.TextMatrix(REG&, 13))
    If IDLOT$ <> "" Then
      On Error Resume Next
      CONSUMLOT07.Label4 = TRIM(txtcodigo)
      If Err Then Exit Sub
      On Error GoTo 0
      CONSUMLOT07.Label1 = IDLOT$
      CONSUMLOT07.Show 1
    End If

End Sub

Private Sub MSF1_Scroll()
   If TXTMOVIL.Visible = True Then
     'TXTMOVIL.Left = MSF1.CellLeft
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF1.TEXT = TXTMOVIL
     TXTMOVIL.Visible = False
   End If
   TXTMOVIL.Left = MSF1.CellLeft + 125
End Sub

Private Sub Timer1_Timer()
    CI1% = CODINT%(txtcodigo)
    If CI1% > 0 Then
        Call ULTIMOS3PRESUPUESTOS(INFOSIAHP.MSF3, CI1%, 0)
        Call ULTIMOS3VENTAS(INFOSIAHP.msf2, CI1%, 0)
        Call ESTADIS_COTIZA(INFOSIAHP.MSF4, CI1%)
        Call ESTADIS_VENTAS(INFOSIAHP.MSF5, CI1%)
        TXTSERVTECNICO = VectArticulos.ESPECIFICACION(CI1%) 'ESPECIFIT$(CI1%)
     End If
     Timer1.Enabled = False
End Sub

Private Sub txtcodigo_Change()
   Call TXTDESCRIPCION_Change
End Sub

 Sub TXTCOSTO_DblClick()

'    If TXTCOSTO.Locked = False Then
'       TXTCOSTO.BackColor = &HE0E0E0
'       TXTCOSTO.Locked = True
'    Else
'       TXTCOSTO.Locked = False
'       TXTCOSTO.BackColor = &H80000005
'       Call PINTATEXT(TXTCOSTO)
'    End If
    

End Sub

Private Sub TXTCOSTO_GotFocus()
    Call PINTATEXT(TXTCOSTO)
    TXTCOSTO.BackColor = vbYellow
End Sub


Private Sub TXTCOSTO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     On Error Resume Next
        SendKeys "{tab}"
     On Error GoTo 0
        TXTVIAENTREGA.BackColor = vbYellow
    End If

End Sub

Private Sub TXTCOSTO_LostFocus()
    TXTCOSTO = TRIM$(FORMATNUM$(XVALO(TXTCOSTO), "f12.2"))
    TXTCOSTO.BackColor = vbWhite
'    TXTCOSTO.BackColor = &HE0E0E0
'    TXTCOSTO.Locked = True
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub

Sub TXTDESCRIPCION_Change()
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF3, "Presup./f7.0;Cliente/A22;Proveedor/A12;Cto.Fob U$S/F10.2;M./A3;P.Unit./F10.2;Fecha/D8", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.msf2, "Factura./A20;Cliente/A24;M./A3;P.Unit./F10.2;Fecha/D8", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF4, "Referencia/A28;Cant.Cotiz./F10", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF5, "Referencia/A28;Cant.Ventas/F10", INFOSIAHP)
    '
    NC = XVALO(FRMPRESUAHP.Text1)
    INFOSIAHP.msf2.Rows = 1: INFOSIAHP.MSF3.Rows = 1: INFOSIAHP.MSF4.Rows = 1: INFOSIAHP.MSF5.Rows = 1
    If TRIM$(TXTDESCRIPCION.TEXT) <> "" Then
       CI1% = CODINT(txtcodigo.TEXT)
       If CI1% > 0 Then
'          Timer1.Enabled = True
'          Call ULTIMOS3PRESUPUESTOS(INFOSIAHP.MSF3, CI1%, 0)
'          Call ULTIMOS3VENTAS(INFOSIAHP.MSF2, CI1%, 0)
'          Call ESTADIS_COTIZA(INFOSIAHP.MSF4, CI1%)
'          Call ESTADIS_VENTAS(INFOSIAHP.MSF5, CI1%)
'          TXTSERVTECNICO = ESPECIFIT$(CI1%)
       End If
    End If

End Sub
Sub ESTADIS_VENTAS(MSF As msFlexGrid, CI1%)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT90%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_90 = SALE - INGRE: If ULT_90 < 1 Then ULT_90 = 0
    '
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT180%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_180 = SALE - INGRE: If ULT_180 < 1 Then ULT_180 = 0
    '
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT365%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_365 = SALE - INGRE: If ULT_365 < 1 Then ULT_365 = 0
    '
    MSF.Rows = 4: MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    '
    
End Sub

Sub ESTADIS_COTIZA(MSF As msFlexGrid, CI1%)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT90%) & "' AND CodiInt = " & CI1%
    ULT_90 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT180%) & "' AND CodiInt = " & CI1%
    ULT_180 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT365%) & "' AND CodiInt = " & CI1%
    ULT_365 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    MSF.Rows = 4
    MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    '
End Sub
Sub ULTIMOS3PRESUPUESTOS(MSF As msFlexGrid, CI1%, NC)
     MSF.Rows = 1: MSF.Cols = 8
     SQLX$ = "SELECT TOP 3 NroPres, RasoClie, PreUnit,NroClie,MonedEmi,COTIZAMONEDA, FechEmis, Num_Prov, COSTO_UNIT,LOTES_Reservados, COSTOFOB"
     SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CODIINT =  " & CI1%
     If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND NroClie =  " & NC
     SQLX$ = SQLX$ + " ORDER BY NroPres desc"
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        J& = J& + 1
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = XVALO(!NroPres)
        NC = XVALO(!NroClie)
        MSF.TextMatrix(J&, 2) = VectCliente.nombre(NC) ' rasocli$(XVALO(!NroClie))
        MON% = XVALO(!MonedEmi)
        PRECIOUNITARIO# = XVALO(!PreUnit) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        'VEO PRIMERO SI EL COSTOFOB TIENE INFORMACION POR QUE SE AGREGO POSTERIORMENTE.
        If XVALO(!COSTOFOB) > 0 Then 'SI ENTRA AQUI SE VA IF MON%=1
           COSTOUNITARIO# = XVALO(!COSTOFOB)
        Else
            'TOMO EL VALOR DEL FOB SEGUN EL LOTE VERIFICO QUE NO SE HAYA GUARDADO MAS DE UN LOTE EN EL PRESUPUESTO PARA UN MISMO CODIGO.
            LOTE$ = SAFETEXT$(!LOTES_Reservados)
            COSFOB# = 0
            If LOTE$ <> "" Then
              POS1% = InStr(LOTE$, ",")
              If POS1% > 0 Then LOTE$ = Left$(LOTE$, POS1% - 1)
              COSFOB# = COSTOFOB#(LOTE$)
              COSTOUNITARIO# = COSFOB#
              
            End If
            'SI NO OBTUVO EL COSTO FOB LE PONE EL COSTO NACIONALIZADO
            If COSFOB# <= 0 Then
              COSTOUNITARIO# = XVALO(!COSTO_UNIT)
            End If
        End If
        '
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZAMONEDA) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(MON%) ' CALCULA
               COSTOUNITARIO# = XVALO(!COSTO_UNIT) / TICAMONE(MON%) ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               MON% = 2
           End If
        End If
        MSF.TextMatrix(J&, 3) = TEL2CLI$(-1 * XVALO(!Num_Prov))
        MSF.TextMatrix(J&, 4) = FORMATNUM$(COSTOUNITARIO#, "F14.2")
        MSF.TextMatrix(J&, 5) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
          FE% = CVINT(!FECHEMIS)
        MSF.TextMatrix(J&, 7) = FECHATEX$(FE%)
        .MoveNext
       Loop
     End With

End Sub
Sub ULTIMOS3VENTAS(MSF As msFlexGrid, CI1%, NC)

     MSF.Rows = 1 ': MSF.Cols = 5 CAMBIE TOP 10 PARA QUE SI HAY UNA MISMA FACTURA QUE CONTINE EL MISMO ITEM
                                  'MAS DE UNA VEZ PASE A OTRO, IGUAL CUANDO LLEGA A 3 CORTA LA CARGA
     SQLX$ = "SELECT TOP 10 DOSECLAR, VALOUNIT,Nume_Clie,MONEVAL, FECHMOV, TIPOCAM "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE COD_INTE =  " & CI1%
     If NC > 0 Then SQLX$ = SQLX$ + " AND  Nume_Clie =  " & NC
     SQLX$ = SQLX$ + " AND CODIMOVI = 'RT'"
     SQLX$ = SQLX$ + " ORDER BY FECHMOV DESC,DOSECLAR DESC"
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        FCT$ = SAFETEXT$(!DOSECLAR)
        If FACTURA_ANT$ = FCT$ And TRIM$(UCase$(FCT$)) <> "SIN CARGO" Then GoTo 50   'MOVENEXT
        
        J& = J& + 1
        If J& > 3 Then GoTo 60
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = SAFETEXT$(!DOSECLAR)
        MSF.TextMatrix(J&, 2) = rasocli$(XVALO(!Nume_Clie))
        MSF.TextMatrix(J&, 3) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
        '
'        MON% = XVALO(!MONEVAL)
'        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
'        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
'           TIPOCAMBIODEPRESUP# = XVALO(!TIPOCAM) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
'           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
'               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMBIODEPRESUP# ' CALCULA
'           Else
'               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
'               MON% = 2
'           End If
'        End If
        MON% = XVALO(!MoneVal)
        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% <= 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!TIPOCAM) 'TOMA EL CAMBIO EN QUE FUE GUARDADO EN MOVISTO
           If TIPOCAMBIODEPRESUP# > 0.0005 And Left$(MSF.TextMatrix(J&, 1), 2) = "RT" Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMBIODEPRESUP# ' CALCULA
               MON% = 2
           Else
               'SI NO TIENE EL TIPO DE CAMBIO EN EL MOVISTO LO BUSCA EN LA FACTURA
               TIPOCAMFACTURA# = XVALO(VALOBADA("CAJABANC", "COTIZA_DOLAR", "CODICOM_LAR = '" & MSF.TextMatrix(J&, 1) & "'", "NOMESS"))
               If TIPOCAMFACTURA# > 0.0005 Then
                 PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMFACTURA#
               Else
                 PRECIOUNITARIO# = XVALO(!VALOUNIT) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               End If
               MON% = 2
           End If
        End If
        
        
        MSF.TextMatrix(J&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        MSF.TextMatrix(J&, 5) = TRIM$(FECHATEX$(CVINT(!FechMov)))
        '
        FACTURA_ANT$ = SAFETEXT$(!DOSECLAR)
50      .MoveNext
       Loop
     End With
60   RS.Close
     Set RS = Nothing
End Sub


Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     REG& = TXTMOVIL.ToolTipText
     If XVALO(TXTMOVIL) > XVALO(MSF1.TextMatrix(MSF1.Row, 12)) Then
       Call COMUNI("Cantidad Ingresada Supera la Disponible, Verifique!!!!")
       TXTMOVIL = ""
     Else
      MSF1.TextMatrix(REG&, 11) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
      TXTMOVIL.Visible = False
      On Error Resume Next
      Command1.SetFocus
      On Error GoTo 0
      Call ACTUCANTI
     End If
   End If
End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
       REG& = XVALO(TXTMOVIL.ToolTipText)
       MSF1.TextMatrix(REG&, 11) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
       TXTMOVIL.Visible = False
       Call ACTUCANTI
     End If
End Sub

Private Sub TXTNUMEPROVE_Change()
    NPROV& = XVALO(TXTNUMEPROVE)
    If NPROV& > 0 Then
      Text1 = VALOBADA("Proved12", "RASO_CLI", "Nume_Cli= " & NPROV&, "NOMESS")
    Else
      Text1 = ""
    End If
    Call CARGARCOMBOVIAS_TODOS(NPROV&)
End Sub
Private Sub Combo6_Click()
    TXTVIAENTREGA.TEXT = Combo6.TEXT
End Sub

Private Sub TXTNUMEPROVE_DblClick()
   Call Command5_Click
End Sub

Private Sub TXTNUMEPROVE_GotFocus()
    TXTNUMEPROVE.BackColor = vbYellow
    
End Sub

Private Sub TXTNUMEPROVE_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       TXTCOSTO.SetFocus
   End If
End Sub

Private Sub TXTNUMEPROVE_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 113 'pulso  f2
        Call Command5_Click
    End Select

End Sub

Private Sub TXTNUMEPROVE_LostFocus()
   TXTNUMEPROVE.BackColor = &HE0E0E0
End Sub

Private Sub TXTPESO_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       'SendKeys "{tab}"
       Call PINTATEXT(TXTNUMEPROVE)
 End If
End Sub

Private Sub TXTPESO_LostFocus()
   TXTPESO = FORMATNUM$(XVALO(TXTPESO), "F12.2")
End Sub

'Private Sub TXTMOVIL_LostFocus()
'   Call TXTMOVIL_KeyPress(13)
'End Sub
Private Sub TXTSOLICITADA_Change()

End Sub

Private Sub TXTSOLICITADA_DblClick()
    If TXTSOLICITADA.Locked = False Then
       TXTSOLICITADA.BackColor = &HE0E0E0
       TXTSOLICITADA.Locked = True
    Else
       TXTSOLICITADA.Locked = False
       TXTSOLICITADA.BackColor = &H80000005
       Call PINTATEXT(TXTSOLICITADA)
    End If

End Sub

Sub CARGARCOMBOVIAS_TODOS(Optional NPROV&)
     Combo6.Clear
     
     '
     
     If NPROV& > 0 Then
            'RECORRO LA VIA DE ENTREGA Y CARGO EN UN VECTOR LOS CODIGOS DE VIAS DE ENTREGA FILTRADO POR PROVEEDOR
            RFF$ = RECFILT$("PORCCOEF", 6, MKI$(NPROV&)) '4
            FIN& = Len(RFF$)
            Dim Control()
            Erase Control
            ESCRIBE% = 0
            For UK& = 1 To FIN& Step 4
                U& = CVS(Mid$(RFF$, UK&, 4))
                X$ = REGLEIDO$("PORCCOEF", U&)
                VIAX1 = CVI(Mid$(X$, 1, 2))
                VIAX$ = SAFETEXT$(VIAX1)
                REFVIA$ = ECOARCH$("VIASENTR", MKI$(VIAX1))
                Call REPLA(TX$, VIAX$, 1, 2)
                Call REPLA(TX$, "-", 3, 1)
                Call REPLA(TX$, REFVIA$, 4, 20)
                ReDim Preserve Control(VIAX1)
                If Control(VIAX1) < 1 Then
                  Combo6.AddItem TX$
                  Control(VIAX1) = 1
                  ESCRIBE% = 1
                End If
            Next UK&
            If ESCRIBE% > 0 Then Exit Sub
     End If
           FIN& = ULTREG&("KILOVIAS")

            For i& = 1 To FIN&
               X$ = REGLEIDO$("KILOVIAS", i&)
               VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
               REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
               Call REPLA(TX$, VIAX$, 1, 2)
               Call REPLA(TX$, "-", 3, 1)
               Call REPLA(TX$, REFVIA$, 4, 20)
               Combo6.AddItem TX$
            Next i&
     '

End Sub


Function VIAENTREGA(LOTE$)
   CONDI$ = "IdenLote = '" & LOTE$ & "'"
   NOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
   If NOC > 0 Then
      CONDI$ = "Nume_Ocom = " & NOC&
      VIAENTREGA = XVALO(VALOBADA("OCOMENCA", "Via_Entrega", CONDI$, "NOMESS"))
   End If
End Function

Sub Combo6_GotFocus()
End Sub

Private Sub TXTVIAENTREGA_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       Command3.BackColor = vbYellow
       Command3.SetFocus
 End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    MSF1.COL = 11
    Select Case KeyCode
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF1.TEXT = TXTMOVIL.TEXT

            'MSF1.SetFocus
            If MSF1.Row > 1 Then MSF1.Row = MSF1.Row - 1
          MSF1.COL = 11
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            
             .TEXT = TRIM$(MSF1.TEXT)
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
            Call VERIFICA_CANTI(MSF1.Row - 1)
            Call ACTUCANTI
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            'MSF1.SetFocus
            MSF1.TEXT = TXTMOVIL.TEXT

            If MSF1.Row < MSF1.Rows Then MSF1.Row = MSF1.Row + 1
          MSF1.COL = 11

          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
             .TEXT = TRIM$(MSF1.TEXT)
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
          Call VERIFICA_CANTI(MSF1.Row - 1)
          Call ACTUCANTI
          On Error GoTo 0
    End Select

    '
End Sub
Sub VERIFICA_CANTI(REG1&)
        If XVALO(MSF1.TextMatrix(REG1&, 11)) > XVALO(MSF1.TextMatrix(REG1&, 12)) Then
          Call COMUNI("Cantidad Ingresada en Supera la Disponible, Verifique!!!!")
          MSF1.TextMatrix(REG1&, 11) = ""
        End If
End Sub

Private Sub TXTVIAENTREGA_LostFocus()
    TXTVIAENTREGA.BackColor = vbWhite
End Sub

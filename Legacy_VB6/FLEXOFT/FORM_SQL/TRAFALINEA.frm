VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TRAFALINEA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transferencias de Ingreso a Linea"
   ClientHeight    =   9930
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   15735
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9930
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15735
      Begin VB.TextBox Text2 
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
         Height          =   2610
         Index           =   3
         Left            =   3240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "TRAFALINEA.frx":0000
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14220
         TabIndex        =   43
         Top             =   9600
         Width           =   1455
      End
      Begin VB.TextBox TXTOPERACIONSELECCIONADA 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox Text14 
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
         Index           =   11
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Transferido a Linea"
         Top             =   7155
         Width           =   5460
      End
      Begin VB.TextBox Text5 
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
         Height          =   350
         Left            =   3600
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Posicionar Cursor Para Leer Código QR de Etiqueta de Pallet para Transferir a Linea"
         Top             =   6720
         Width           =   10335
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Ver Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command43 
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
         Left            =   2280
         Picture         =   "TRAFALINEA.frx":005A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Revertir Selección"
         Top             =   4200
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14220
         TabIndex        =   13
         Top             =   6630
         Width           =   1455
      End
      Begin VB.TextBox Text14 
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
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Detalle"
         Top             =   4200
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox Text14 
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
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Materiales en Rack de Semiterminados"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox TXTOT 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   360
         Width           =   1725
      End
      Begin VB.CommandButton CMDOT 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1995
         TabIndex        =   8
         ToolTipText     =   "Ordenes de Fabricación Abiertas"
         Top             =   360
         Width           =   210
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   15435
         Begin VB.CommandButton Command14 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   7.5
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2160
            TabIndex        =   45
            ToolTipText     =   "Ordenes de Fabricación Abiertas y Cerradas"
            Top             =   240
            Width           =   210
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   150
            Width           =   1800
         End
         Begin VB.TextBox Text11 
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
            Height          =   350
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   420
            Width           =   1800
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   4000
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   150
            Width           =   4000
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Cant.Proyec."
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Left            =   4620
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "Cant.Aprob."
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   4620
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   150
            Width           =   1140
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox TXTRAZOCLI 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   12360
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
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
            Left            =   12360
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Cliente"
            Top             =   150
            Width           =   3000
         End
         Begin VB.TextBox Text14 
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
            Index           =   10
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Cuenta"
            Top             =   150
            Width           =   860
         End
         Begin VB.TextBox TXTNC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   350
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   860
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   10800
         TabIndex        =   4
         Top             =   960
         Width           =   4755
         Begin VB.TextBox Text14 
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
            Index           =   6
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "F.Técnica"
            Top             =   140
            Width           =   1140
         End
         Begin VB.TextBox TXTFTECNICA 
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
            Height          =   350
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   380
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Index           =   5
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "Cotización"
            Top             =   140
            Width           =   1140
         End
         Begin VB.TextBox TXTNROCOTIZA 
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
            Height          =   350
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   380
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Cód.Interno"
            Top             =   140
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "Pedido"
            Top             =   140
            Width           =   1140
         End
         Begin VB.TextBox TXTNROPED 
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
            Height          =   350
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   380
            Width           =   1140
         End
         Begin VB.TextBox TXTCODINT 
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
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   360
            Width           =   1155
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2145
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Clic Para ver detalle según Fila Seleccionada"
         Top             =   1920
         Visible         =   0   'False
         Width           =   15435
         _ExtentX        =   27226
         _ExtentY        =   3784
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   12632256
         GridColor       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
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
         Height          =   1545
         Left            =   7320
         TabIndex        =   7
         Top             =   3960
         Visible         =   0   'False
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   2725
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   12632256
         GridColor       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   4665
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Doble Clic Para Transferira a Linea"
         Top             =   1920
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   8229
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   12632256
         GridColor       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   2025
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Doble Clic Para Revertir Ingreso a Linea"
         Top             =   7500
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   3572
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   12632256
         GridColor       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lectura de QR   >>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   6720
         Width           =   3855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "TRAFALINEA.frx":01A4
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "TRAFALINEA.frx":03D8
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "TRAFALINEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub CargarMsf3()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NORFA$ = TRIM$(TXTOT)
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   'armo lista de racks de talleres
   Talleres$ = ""
   For i% = 30 To 50
     If i% = 30 Then 'la primera vuelta
       Talleres$ = i%
     Else
       Talleres$ = Talleres$ & ";" & i%
     End If
   Next i%
   'fin
   
   CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
   If (XVALO(Mid$(NORFA$, 4, 6)) >= 2341 Or XVALO(Mid$(NORFA$, 4, 6)) = 2325) And InStr(CODFAMI$, "BB") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then  'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
        Call ctp.NuevaCargaOperacionesInformadasSegunUbicaciones(NORFA$, CI1%, MSF3, 0, "", 99, "7;8;9;10;11;" & Talleres$, "")
        Text14(11) = "Saldos de Cortes Transferidos a Linea (Agrupados)"
   Else
        Call ctp.CargaOperacionesInformadasSegunUbicaciones(NORFA$, CI1%, MSF3, 0, "", 99, "7;8;9;10;11;" & Talleres$, "")
   End If
   
   Set ctp = Nothing
End Sub

Private Sub Check1_Click()
   Check1.ForeColor = &HFFFFFF
   If Check1.Value > 0 Then
      Check1.ForeColor = vbYellow
   End If
End Sub

Private Sub Check2_Click()
   Check2.ForeColor = &HFFFFFF
   If Check2.Value > 0 Then
      Check2.ForeColor = vbYellow
   End If
End Sub



Private Sub CMDOT_Click()
   CMDOT.Enabled = False
   NORFA$ = ORFASEL_ENVAPLAST$
   If NORFA$ <> "" Then
     CI1% = CODIFAB%(NORFA$)
     TXTCODINT = CI1%
     Text10 = NORFA$
     Text11 = CODEXT$(CI1%)
     Text4 = DESCRIT0$(CI1%)
     TXTOT = NORFA$
   End If
   CMDOT.Enabled = True
End Sub

Sub printformu(MSF1 As msFlexGrid, i&, COL&)

   If CODINT(Text11) < 1 Then
      Call COMUNI("Código No Válido o Inexistente")
      GoTo 99
   End If
   
   CONDI$ = "CODIMASTER = '" & Text11 & "'"
   NROFTEC& = Me.TXTFTECNICA
   NROCOT& = XVALO(Me.TXTNROCOTIZA)
   NORFA$ = Me.Text10
   If TRIM$(MSF1.TextMatrix(i&, 1)) = "IMPRESION" Then
      FORIPRE$ = TRIM$(MSF1.TextMatrix(i&, 4))
      FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&))
      Call STDFORM("OCCU-CTP", FILTX$)
   Else
      CI1% = XVALO(MSF1.TextMatrix(i&, 5)) 'CODIGO VISIBLE EN LA GRILLA
      CI2% = XVALO(MSF1.TextMatrix(i&, 6)) 'CODIGO ORIGINAL DE LA COTIZACION
      If CI2% > 0 Then
         CODIGOORIGINAL$ = CODEXT$(CI2%)
         CODIGORANGOX1$ = CODEXT$(CI1%)
         CODIGOCORTE$ = CODIGOORIGINAL$
      Else
         CODIGOORIGINAL$ = CODEXT$(CI1%)
         CODIGORANGOX1$ = "----"
         CODIGOCORTE$ = CODIGOORIGINAL$
      End If
      
      If TRIM$(CODIGOORIGINAL$) = "" Then CODIGOORIGINAL$ = "-----" 'PARA QUE NO APAREZCA LEYENDA COMPAR-1
      
      FORIPRE$ = TRIM$(MSF1.TextMatrix(i&, 4))
      FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&)) & ";" & TRIM$(CODIGORANGOX1$) & ";" & TRIM$(CODIGOORIGINAL$)
      Call STDFORM(FORIPRE$, FILTX$)
      End If
   
   
'   'ETIQUETA
'   CI1% = CODINT(Text11)
'   If Check1.Value > 0 Then
'      Call PRINTETIQUETACTP(CI1%, NORFA$, CODIGOEDITABLE$)
'   End If
'
'   'DOBLADO
'   If Check2.Value > 0 Then
'      Call PRINTDOBLADO(CI1%, NORFA$)
'   End If
99
End Sub
Sub PRINTDOBLADO(CI1%, NORFA$)
    If CI1% < 1 Then
       Call COMUNI("No Hay un Código Activo")
       Exit Sub
    End If
    '
    NORFA$ = TRIM$(NORFA$)
    If NORFA$ = "" Then
       Call COMUNI("Falta Orden de Fabriación")
       Exit Sub
    End If
    '
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")

    '
    
    RUTAEXCEL$ = Control("COTICONT", "EXCEDOBL")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
'datos a imprimir
'cliente
'NORFA$
'CANTIDAD

    Dim ctp As New Conteplast
    NP& = XVALO(TXTNROPED)
    CANTI = XVALO(Text3)
    NC = cliente_SEGUN_PEDIDO(NP&)
    CI1% = XVALO(Me.TXTCODINT)
    RASO$ = rasocli$(NC)
    CODI$ = TRIM$(Text11)
    DESKRI$ = TRIM$(Text4)
    '
    With VENTABNEW.MSF
       .Rows = 4: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = NORFA$
       .TextMatrix(3, 1) = TRIM$(FORMATNUM$(CANTI, "F10.0"))
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "DOBLADO", VENTABNEW.MSF, PDES$, "1", 1, 1, "", 1, 0)                    'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Set ctp = Nothing
    Unload VENTABNEW
End Sub

Sub LIMPIAR()

    Text11 = ""
    Text4 = ""
    Text10 = ""
    MSF.Rows = 1
    '
    Text6 = ""
    MSF1.Rows = 1

End Sub

Private Sub Command1_Click()
    command1.Enabled = False
    '
    NORFA$ = TXTOT
    Campos$ = "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0"
    Call CARGA_TABLA_IMPRE(MSF3, Campos$, "INFOUBIC", "Operaciones x O.F. Sector Previo a Linea", "TABLA_IMPRE", , 1)
    '*****************************************************'
    FILTX$ = TRIM$(FECHATEX$(HOY(HO$))) & ";" & TRIM$(NORFA$)
    Call STDFORM("UBIK-CTP", FILTX$)
    command1.Enabled = True
End Sub

Private Sub Command14_Click()

   Command14.Enabled = False
   NORFA$ = ORFASEL_ENVAPLAST$(1)
   If NORFA$ <> "" Then
     CI1% = CODIFAB%(NORFA$)
     TXTCODINT = CI1%
     Text10 = NORFA$
     Text11 = CODEXT$(CI1%)
     Text4 = DESCRIT0$(CI1%)
     TXTOT = NORFA$
   End If
   Command14.Enabled = True

End Sub

Private Sub Command2_Click()
    If TRIM$(TXTOT) = "" Then Exit Sub
    TXTOPERACIONSELECCIONADA = ""
    Dim ctp As New Conteplast
    Call ctp.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF3, 0, "", 99)
    Set ctp = Nothing

End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    NORFA$ = TXTOT
    Campos$ = "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0"
    Call CARGA_TABLA_IMPRE(MSF4, Campos$, "INFOUBIC", "Operaciones x O.F. Sector Previo a Linea", "TABLA_IMPRE", , 1)
    '*****************************************************'
    FILTX$ = TRIM$(FECHATEX$(HOY(HO$))) & ";" & TRIM$(NORFA$)
    Call STDFORM("UBIL-CTP", FILTX$)
    Command3.Enabled = True
End Sub

Private Sub Form_Load()
 CAMPOS1$ = "Operación/A30" '1
 CAMPOS1$ = CAMPOS1$ + ";Código./A0" '2
 CAMPOS1$ = CAMPOS1$ + ";Descripción/A0" '3
 CAMPOS1$ = CAMPOS1$ + ";Formulario/A0" '4
 CAMPOS1$ = CAMPOS1$ + ";CI/A0" '5
 CAMPOS1$ = CAMPOS1$ + ";CIORIG/A0" '6
 CAMPOS1$ = CAMPOS1$ + ";Necesidad/F10.1" '7
 CAMPOS1$ = CAMPOS1$ + ";Cant.Inf/F6.1" '8
 CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/F8" '9
 CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/A16" '10
 CAMPOS1$ = CAMPOS1$ + ";Cant.S.P./F9.1" '11
 CAMPOS1$ = CAMPOS1$ + ";Saldo/f9.1" '12
 Call CARGA_ENCABEZADO(MSF1, CAMPOS1$, Me)
 
 Call CARGA_ENCABEZADO(MSF2, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;CI/f6", Me)

 Call CARGA_ENCABEZADO(MSF3, "Código/A14;Descripción/A32;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;Inf/F3;IDUbic/a0;N°Operario/a0;Nombre/A0;Peso/f8.2;IdReg/a0", Me)
 Call CARGA_ENCABEZADO(MSF4, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0", Me)
End Sub

'
Private Sub MSF1_Click()
'    REG& = MSF1.MouseRow
'    J& = MSF1.MouseCol
'    If REG& = 0 Or REG& > MSF1.Rows Then Exit Sub
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbYellow, REG&)
    REG& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
    DESCOPERACION$ = MSF1.TextMatrix(REG&, 1)
    Dim ctp As New Conteplast
    TXTOPERACIONSELECCIONADA = DESCOPERACION$
    Call ctp.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF3, 0, TXTOPERACIONSELECCIONADA, 99)
    Call ctp.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF4, 100, "", 199)

    Set ctp = Nothing

End Sub

Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If REG& = 1 Or REG& > MSF1.Rows Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbYellow, REG&)
    Call printformu(MSF1, REG&, J&)
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbWhite, REG&)


End Sub


Private Sub MSF2_Click()
    REG& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If REG& = 0 Or REG& > MSF2.Rows Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF2, vbYellow, REG&)
End Sub


  Sub CARGARMSF1_respaldo()
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   MSF1.Rows = 1
   SELPRINOT.Text11 = CODEXT(CI1%)
   SELPRINOT.Text4 = DESCRIT0$(CI1%)
   TXTFTECNICA = ctp.NroFTecnicaXCodigo(CI1%)
   NROCOTIZAX& = XVALO(TXTNROCOTIZA)
   CANTIPROY = XVALO(Text3)
   NORFA$ = TRIM$(Text10)
   REG& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1% & ""
   SQLX$ = SQLX$ + " AND LEFT(COMPONENTE ,4) <> 'HILO' ORDER BY NUORIT"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
       REG& = REG& + 1
         CIELEM% = XVALO(!codielem)
         MSF1.Rows = REG& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         MSF1.TextMatrix(REG&, 1) = NOMBRESECTOR$
         MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!CodXElem)
         MSF1.TextMatrix(REG&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         MSF1.TextMatrix(REG&, 4) = ctp.FRMSECTOR(NOMBRESECTOR$)
         MSF1.TextMatrix(REG&, 5) = CIELEM%
         CIRANGO = ""
         If ESCODIGORANGO(CIELEM%) > 0 Then
            'en esta fila guarda el codigo interno cargado en la cotizacion
            MSF1.TextMatrix(REG&, 6) = ctp.CodigoOriginalCoti1SegunComponente(NROCOTIZAX&, NOMBRESECTOR$)
         End If
         CANTICORTE = XVALO(!moduxbb): If CANTICORTE < 1 Then CANTICORTE = 1
         NECESIDAD = CANTICORTE * CANTIPROY
         MSF1.TextMatrix(REG&, 7) = NECESIDAD
        'MODOX% = 1 CON IDREGPADRE > 0
        'MODOX% = 2 CON OPERFUCIONADA > 0
        'MODOX% = 3 CON IDREGPADRE > 0 Y OPERFUCIONADA <> ''
        'MODOX% = 4 CON IDREGPADRE < 1"
         INFORMADO = ctp.CANTIDADINFORMADA(NORFA$, NOMBRESECTOR$, 0, 0)
         MSF1.TextMatrix(REG&, 8) = INFORMADO
         '
         INFORMADO = ctp.CANTIDADINFORMADA(NORFA$, "", 0, 1)
         MSF1.TextMatrix(REG&, 9) = INFORMADO
         
         CantSubproc = CantAcumuladaSubProceso&
         SALDO = XVALO(NECESIDAD) - INFORMADO: If SALDO < 0 Then SALDO = 0
         MSF1.TextMatrix(REG&, 12) = SALDO
       .MoveNext
     Loop
   End With
   Set ctp = Nothing
   
CAMPOS1$ = "Operación/A30" '1
CAMPOS1$ = CAMPOS1$ + ";Código./A0" '2
CAMPOS1$ = CAMPOS1$ + ";Descripción/A0" '3
CAMPOS1$ = CAMPOS1$ + ";Formulario/A0" '4
CAMPOS1$ = CAMPOS1$ + ";CI/A0" '5
CAMPOS1$ = CAMPOS1$ + ";CIORIG/A0" '6
CAMPOS1$ = CAMPOS1$ + ";Necesidad/F10.1" '7
CAMPOS1$ = CAMPOS1$ + ";Cant.Inf/F6.1" '8
CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/F8" '9
CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/A16" '10
CAMPOS1$ = CAMPOS1$ + ";Cant.S.P./F9.1" '11
CAMPOS1$ = CAMPOS1$ + ";Saldo/f9.1" '12

End Sub

Sub CARGARMSF1()
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   MSF1.Rows = 1
   SELPRINOT.Text11 = CODEXT(CI1%)
   SELPRINOT.Text4 = DESCRIT0$(CI1%)
   TXTFTECNICA = ctp.NroFTecnicaXCodigo(CI1%)
   NROCOTIZAX& = XVALO(TXTNROCOTIZA)
   CANTIPROY = XVALO(Text3)
   NORFA$ = TRIM$(Text10)
   REG& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1% & ""
   SQLX$ = SQLX$ + " AND LEFT(COMPONENTE ,4) <> 'HILO' ORDER BY NUORIT"
   
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
       REG& = REG& + 1
         CIELEM% = XVALO(!codielem)
         MSF1.Rows = REG& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         MSF1.TextMatrix(REG&, 1) = NOMBRESECTOR$
         MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!CodXElem)
         MSF1.TextMatrix(REG&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         MSF1.TextMatrix(REG&, 4) = ctp.FRMSECTOR(NOMBRESECTOR$)
         MSF1.TextMatrix(REG&, 5) = CIELEM%
         CIRANGO = ""
         If ESCODIGORANGO(CIELEM%) > 0 Then
            'en esta fila guarda el codigo interno cargado en la cotizacion
            MSF1.TextMatrix(REG&, 6) = ctp.CodigoOriginalCoti1SegunComponente(NROCOTIZAX&, NOMBRESECTOR$)
         End If
         CANTICORTE = XVALO(!moduxbb): If CANTICORTE < 1 Then CANTICORTE = 1
         NECESIDAD = CANTICORTE * CANTIPROY
         MSF1.TextMatrix(REG&, 7) = NECESIDAD
        'MODOX% = 1 CON IDREGPADRE > 0
        'MODOX% = 2 CON OPERFUCIONADA > 0
        'MODOX% = 3 CON IDREGPADRE > 0 Y OPERFUCIONADA <> ''
        'MODOX% = 4 CON IDREGPADRE < 1"
         INFORMADO = ctp.CANTIDADINFORMADA(NORFA$, NOMBRESECTOR$, 0, 0)
         MSF1.TextMatrix(REG&, 8) = INFORMADO
         '
         INFORMADO = ctp.CANTIDADINFORMADA(NORFA$, "", 0, 1)
         MSF1.TextMatrix(REG&, 9) = INFORMADO
         
         CantSubproc = CantAcumuladaSubProceso&
         SALDO = XVALO(NECESIDAD) - INFORMADO: If SALDO < 0 Then SALDO = 0
         MSF1.TextMatrix(REG&, 12) = SALDO
       .MoveNext
     Loop
   End With
   Set ctp = Nothing
   
CAMPOS1$ = "Operación/A30" '1
CAMPOS1$ = CAMPOS1$ + ";Código./A0" '2
CAMPOS1$ = CAMPOS1$ + ";Descripción/A0" '3
CAMPOS1$ = CAMPOS1$ + ";Formulario/A0" '4
CAMPOS1$ = CAMPOS1$ + ";CI/A0" '5
CAMPOS1$ = CAMPOS1$ + ";CIORIG/A0" '6
CAMPOS1$ = CAMPOS1$ + ";Necesidad/F10.1" '7
CAMPOS1$ = CAMPOS1$ + ";Cant.Inf/F6.1" '8
CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/F8" '9
CAMPOS1$ = CAMPOS1$ + ";SUB-PROC/A16" '10
CAMPOS1$ = CAMPOS1$ + ";Cant.S.P./F9.1" '11
CAMPOS1$ = CAMPOS1$ + ";Saldo/f9.1" '12

End Sub


Private Sub MSF3_DblClick()
    REG& = MSF3.MouseRow
    J& = MSF3.MouseCol
    If REG& < 1 Or REG& >= MSF3.Rows Then Exit Sub
    Dim MSF As msFlexGrid
    Set MSF = MSF3
    '
    'BLOQUE:LE ASIGNO EL CODIGO TAPACOMPLETA PARA QUE EN EL DOBLE CLIC PUEDA BAJAR
    'DADO QUE HAY UNA VALIDACION QUE SIN CODIGO NO BAJA
    'LUEGO AL BAJAR VA A MOVER TAMBIEN LOS 2 ITEMS POR EL CUAL SE GENERO LA FUCION DE LA TAPA
    CI1% = XVALO(MSF.TextMatrix(REG&, 5))
    NOP% = XVALO(MSF.TextMatrix(REG&, 6))
    If CI1% < 1 And NOP% = 1201 Then     'tapa completa
       MSF.TextMatrix(REG&, 9) = CODEXT$(1081)   ' 1081 CODIGO INTERNO DE TAPACOMPLETA
       MSF.TextMatrix(REG&, 5) = "1081"
       ID& = XVALO(MSF.TextMatrix(REG&, 16))
       CONDI$ = "INFOFAB_ID = " & ID&
       Call ACTUREGISTRO("INFOFAB", "CODXELEM", CONDI$, CODEXT$(1081), REGAF&, "NOMESS")
       Call ACTUREGISTRO("INFOFAB", "CODIELEM", CONDI$, 1081, REGAF&, "NOMESS")
    End If
    'FIN BLOQUE
    
    Codigo$ = TRIM$(MSF.TextMatrix(REG&, 9))
    LOTE$ = TRIM$(MSF.TextMatrix(REG&, 3))
    CANTI = XVALO(MSF.TextMatrix(REG&, 4))
    NOP% = XVALO(MSF.TextMatrix(REG&, 6))
    NC = XVALO(TXTNC)
    NROPED& = XVALO(TXTNROPED)
    NROOF$ = TRIM$(TXTOT)
    ORDENINFO = XVALO(MSF.TextMatrix(REG&, 11))
    UBICACION$ = TRIM$(MSF.TextMatrix(REG&, 10))
    IDUBICACION = XVALO(MSF.TextMatrix(REG&, 12))

    TEXTOQR$ = Codigo$ & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|" & NC & "|" & NROPED&
    TEXTOQR$ = TEXTOQR$ + "|" & NROOF$ & "|" & ORDENINFO & "|" & UBICACION$ & "|" & IDUBICACION
    Text5 = TEXTOQR$
    Call Text5_KeyPress(13)
    
    
   
    Text5 = ""
 '1"Código/A16;
 '2Descripción/A28;
 '3Lote/A8;
 '4Cant.Cortes/F10.1;
 '5CI/a0;
 '6N°Op./F5;
 '7Descrip/A12;
 '8Hora/A0;
 '9Material/A16;
 '10Ubicación/A12;
 '11NroInf/a0;
 '12IDUbic/a0"
End Sub


Private Sub MSF3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ESTA TRANSACCION PERMITE DIVIDIR UN REGISTRO Y MOVERLO HACIA LA LINEA
    'EJ, SI HAY UN INFORME DE CORTE DE 100 UNIDADES Y SOLICITA PARCIALIZAR 60 BAJA LOS 60 Y MANTIENE ARRIBA
    ' EL DE 40
    '  EL REGISTRO DUPICADO SE LE ACTUALIZA EL N DE OPERACION Y CANTIDAD
    '   AL REGISGTRO ORIGINAL SOLO LA CANTIDAD.
    '!!!!!    A LA LINEA DEBERIA IR LO QUE SE VA A CONSUMIR EN LOS CIERRES   !!!!!
    ' *****   Y MANTENER PREVIO A LA LINEA LO QUE SE HAN LLEVADO ASI CUANDO VUELVA LA MERCADERIA SE GENEREN LOS CIERRES Y SE CONSUMA
    If Button <> 2 Then Exit Sub
    REG& = MSF3.MouseRow
    COL% = MSF3.MouseCol
    If REG& = 0 Or REG& >= MSF3.Rows Then Exit Sub
    Call SELECCIONARFILA(MSF3, REG&)
    IDINFOFAB& = XVALO(MSF3.TextMatrix(REG&, 16))
    
    OPX% = COMALTER%("Envio Parcial a Linea\\Canelar\Ingresar Cantidad Parcial Para Enviar a Linea")
    If OPX% < 2 Then Exit Sub
    
    CI1% = XVALO(MSF3.TextMatrix(REG&, 5))
    NOP% = XVALO(MSF3.TextMatrix(REG&, 6))
    If CI1% < 1 And NOP% = 1201 Then     'tapa completa
      Call COMUNI("Atenciòn!! Imposible Continuar Tapa Completa Bloqueada Para Parcializaciòn")
      Exit Sub
    End If
     

    CantidadOriginal$ = TRIM$(MSF3.TextMatrix(REG&, 4))
    CantidadInformadaActual$ = TRIM$(MSF3.TextMatrix(REG&, 4))
    CANTIDADINFORMADA$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad Total Informada", CantidadInformadaActual$))
    If XVALO(CANTIDADINFORMADA$) < 1 Then Exit Sub
    
    If XVALO(CANTIDADINFORMADA$) >= XVALO(CantidadOriginal$) Then
       Call COMUNI("Imposible Continuar La Cantidad Ingresada Debe ser Inferior a la Cantidad Original")
       Exit Sub
    End If
    
    Codigo$ = TRIM$(MSF3.TextMatrix(REG&, 9))
    LOTE$ = TRIM$(MSF3.TextMatrix(REG&, 3))
    CANTI = XVALO(MSF3.TextMatrix(REG&, 4))
    NOP% = XVALO(MSF3.TextMatrix(REG&, 6))
    NC = XVALO(TXTNC)
    NROPED& = XVALO(TXTNROPED)
    NROOF$ = TRIM$(TXTOT)
    ORDENINFO = XVALO(MSF3.TextMatrix(REG&, 11))
    UBICACION$ = TRIM$(MSF3.TextMatrix(REG&, 10))
    IDUBICACION = XVALO(MSF3.TextMatrix(REG&, 12))
    ID& = XVALO(MSF3.TextMatrix(REG&, 16))
    CantiParcial& = XVALO(CANTIDADINFORMADA$)
    CantiOrigDescontandoParcial& = XVALO(CantidadOriginal$) - CantiParcial
    
    'duplico el registro
    NEWID& = 1 'hay que pasarle un valor a nuevoid para que en duplicaregistro devuelva el id del nuevo registro
    Call DuplicarRegistro("INFOFAB", "INFOFAB_ID = " & ID&, NEWID&)
    '
    'actualizo cantidades
    Call ACTUREGISTRO("INFOFAB", "CANTREAL", "INFOFAB_ID = " & ID&, CantiOrigDescontandoParcial&, REGAF&, "NOMESS")
    Call ACTUREGISTRO("INFOFAB", "CANTREAL", "INFOFAB_ID = " & NEWID&, CantiParcial&, REGAF&, "NOMESS")
        'actualizo el n orden en el nuevo registro
        CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOP%
        NUEVOMONUMORDENOP% = 1 + XVALO(VALOBADA("INFOFAB", "MAX(nuorinf)", CONDI$, "NOMESS"))
    Call ACTUREGISTRO("INFOFAB", "nuorinf", "INFOFAB_ID = " & NEWID&, NUEVOMONUMORDENOP%, REGAF&, "NOMESS")
    '
    'AHORA LO QUE CARGA SON LOS DATOS DEL NUEVO REGISTROS (CantiParcial&,NUEVOMONUMORDENOP%, EL RESTO QUEDA IGUAL)
    TEXTOQR$ = Codigo$ & "|" & LOTE$ & "|" & CantiParcial& & "|" & TRIM$(str$(NOP%)) & "|" & NC & "|" & NROPED&
    TEXTOQR$ = TEXTOQR$ + "|" & NROOF$ & "|" & TRIM$(str$(NUEVOMONUMORDENOP%)) & "|" & UBICACION$ & "|" & IDUBICACION
    Text5 = TEXTOQR$
    Call Text5_KeyPress(13)
    Text5 = ""
End Sub


Private Sub MSF4_DblClick()
    REG& = MSF4.MouseRow
    If REG& = 0 Or REG& >= MSF4.Rows Then Exit Sub
    IDINFOFAB& = XVALO(MSF4.TextMatrix(REG&, 16))
    CONDI$ = "INFOFAB_ID =" & IDINFOFAB&
    If XVALO(VALOBADA("INFOFAB", "CONSUMO", CONDI$, "NOMESS")) > 0 Then 'TIENE INFORMADO CONSUMO
       Call COMUNI("Imposible Revertir\ El Registro Seleccionado Yá Tiene Informado Consumo Por Orden de Fabricación")
       Exit Sub
    End If
    VENTANA.Inicializar = 0
    Campos$ = "Código/A16N;" '1
    Campos$ = Campos$ + "Descripción/A28;" '2
    Campos$ = Campos$ + "Lote/A8;" '3
    Campos$ = Campos$ + "Cant.Cortes/F10.1;" '4
    Campos$ = Campos$ + "CI/a0;" '5
    Campos$ = Campos$ + "N°Op./F5;" '6
    Campos$ = Campos$ + "Descrip/A12;" '7
    Campos$ = Campos$ + "Hora/A0;" '8
    Campos$ = Campos$ + "Material/A16;" '9
    Campos$ = Campos$ + "Ubicación/A12;" '10
    Campos$ = Campos$ + "NroInf/F5;" '11
    Campos$ = Campos$ + "IDUbic/a0" '12
    VENTANA.Campox = Campos$
    
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Revertir Transferencia a Linea"
    VENTANA.CampEditables = "10"  'no muestro lista de seleccion para no cambiar la estrucutra que se usa
    'en diferentes lugares con una estructura definida
    
'    VENTANA.ColEcoCampo(10) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'SECTPLANTA' AND CONVERT(INT,CAMPO1) < 100"
'    Condi$ = "LISTA = 'SECTPLANTA' AND CONVERT(INT,CAMPO1) = " & ID&
'    UBICACIONSEGUNID$ = VALOBADA("DATASQL", "CAMPO2", Condi$, "NOMESS")

    VENTANA.Inicializar = 1
    SECTORANTERIOR$ = TRIM$(UCase$(MSF4.TextMatrix(REG&, 10)))
   
25  Call EDITARFILA(MSF4, VENTABNEW.MSF, REG&)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       With VENTABNEW
         Dim ctp As New Conteplast
         SECTOR$ = UCase$(TRIM$(.MSF.TextMatrix(1, 10)))
         IDSECTOR& = ctp.IDSEGUNUBICACION&(SECTOR$)
         Set ctp = Nothing
         'VALIDAR SI ES LOS SECTORES POSIBLES SOLO PUEDE VOLVER HACIA ATRAS
         If TRIM$(SECTORANTERIOR$) = SECTOR$ Then Exit Sub 'SI ES EL MISMO SALE
         
         If XVALO(SECTOR$) < 1 Then
            If IDSECTOR& < 1 Or IDSECTOR& >= 100 Then
               Call COMUNI("Atención Sector Ingresado No Válido, Seleccione un Sector Válido")
               'SI EL INGESO ES INCORRECTO MUESTRA LISTA
               X = LISTASELECGENERAL("SECTPLANTA", "", "", 0, "CONVERT(int, campo1) < 100")
               IDSECTOR = XVALO(RESP_ZELE_VECT(1))
               SECTOR$ = TRIM$(RESP_ZELE_VECT(2))
            End If
         ElseIf XVALO(SECTOR$) > 0 Then
            If InStr(SECTOR$, ";") > 0 Then
              POS1% = InStr(SECTOR$, ";")
              RESTOSECTOR$ = Mid(SECTOR$, POS1% + 1)
              If InStr(RESTOSECTOR$, ";") < 1 Then
                 GoTo 70
              End If
            Else
              GoTo 70
            End If
         End If
         'SI NO ELIGE SALE
         If SECTOR$ = "" Then Exit Sub
         NORFA$ = TRIM$(TXTOT)
         NOP% = XVALO(.MSF.TextMatrix(1, 6))
         LOTE$ = .MSF.TextMatrix(1, 3)
         CI1% = CODINT(.MSF.TextMatrix(1, 9))
         NUORD% = XVALO(.MSF.TextMatrix(1, 11))
         CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
         CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
         CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
         CONDI$ = CONDI$ + " AND  CODIELEM = " & CI1%
         CONDI$ = CONDI$ + " AND  NUORINF = " & NUORD%
         CONDI$ = "INFOFAB_ID = " & IDINFOFAB&
         Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, SECTOR$, REGAF&, "NOMESS")
         Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, IDSECTOR, REGAF&, "NOMESS")
         'VUELVO PARA ATRAS LA FECHA DE TRANSFERENICA A LINEA
         Call ACTUREGISTRO("INFOFAB", "FeTranLinea", CONDI$, CVDAT(0), REGAF&, "NOMESS")
         'Call CTP.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF3, 0, TXTOPERACIONSELECCIONADA, 99)
         Call CargarMsf3
         Call ctp.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF4, 100, "", 199)
         Set ctp = Nothing
       End With
    End If
    Unload VENTABNEW
    Exit Sub
70  Call COMUNI("Atención Formato Posición de Rack Ingresada, No Válida, Seleccione una Posición con Formato Correcto (Ej.: 1;3;4)")

    Unload VENTABNEW

End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
 '1"Código/A16;
 '2Descripción/A28;
 '3Lote/A8;
 '4Cant.Cortes/F10.1;
 '5CI/a0;
 '6N°Op./F5;
 '7Descrip/A12;
 '8Hora/A0;
 '9Material/A16;
 '10Ubicación/A12;
 '11NroInf/a0;
 '12IDUbic/a0"
'     CODEXT$(CIELEM%) & "|" & LOTE$ & "|" & CANTI & "|" & NOP% & "|"
'     & NC & "|" & NROPED& & "|" & NROOF$ & "|" & ORDENINFO & "|"
'     & UBICACION$ & "|" & IDUBICACION
        FE% = HOY(HO$)
        FEHOY$ = FETEXCOR1$(FE%)
        TEXTOQR$ = Text5
        LECTURA$ = TEXTOQR$
        On Error GoTo 0
        strArray = Split(LECTURA$, "|")
        Codigo$ = strArray(0)
        CIELEM% = CODINT(Codigo$)
        If CIELEM% < 1 Then On Error Resume Next: Exit Sub
        '
        Dim ctp As New Conteplast
        NORFA$ = TRIM$(TXTOT)
        LOTE$ = TRIM$(strArray(1))
        NOP% = XVALO(strArray(3))
        NUORD% = XVALO(strArray(7))
        On Error Resume Next
        
        UBICACION$ = ctp.UBICACIONSEGUNID(100)
        'SI ES UBICACION EN RACK ES VALOR NUMERICO REEMPLAZO "-" POR ";" QUE LO GUARDO CON GUION P Q DA ERROR AL GENERAR QR
        If XVALO(UBICACION$) > 0 Then Call REPLACAR(UBICACION$, "-", ";")
        
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
        CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
        CONDI$ = CONDI$ + " AND  CODIELEM = " & CIELEM%
        CONDI$ = CONDI$ + " AND  NUORINF = " & NUORD%
        IDSELEC& = XVALO(VALOBADA("INFOFAB", "INFOFAB_ID", CONDI$, "NOMESS"))
        
        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, UBICACION$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, 100, REGAF&, "NOMESS")
        Call ACTUREGISTRO("INFOFAB", "FeTranLinea", CONDI$, Now, REGAF&, "NOMESS")
        
        'TAMBIEN BAJA LOS ITEMS FUCIONADOS SI EL ITEM SELECCIONADO FUERA EL CASO
        Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", "OperFucionada = " & IDSELEC&, 100, REGAF&, "NOMESS")
        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", "OperFucionada = " & IDSELEC&, UBICACION$, REGAF&, "NOMESS")
        
        'Call CTP.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF3, 0, TXTOPERACIONSELECCIONADA, 99)
        Call CargarMsf3
        Call ctp.CargaOperacionesInformadas(NORFA$, CIELEM%, MSF4, 100, "", 199)
        Call seleccionFilaLinea(MSF4, NOP%, NUORD%)
        Set ctp = Nothing
        TEXT = ""
End Sub

Sub seleccionFilaLinea(MSF4, NOP%, NUORD%)
   Dim MSF As msFlexGrid
   Set MSF = MSF4
   For REG& = 1 To MSF.Rows - 1
      If XVALO(MSF.TextMatrix(REG&, 6)) = NOP% Then
         If XVALO(MSF.TextMatrix(REG&, 11)) = NUORD% Then
            Call SELECCIONARFILA(MSF, REG&)
            Exit For
         End If
      End If
   Next REG&

End Sub
Private Sub TXTOT_Change()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NORFA$ = TXTOT
   Text3 = CantiProyectadaOF(NORFA$)
   Text1 = CantiAprobadaOF(NORFA$)
   TXTCODINT = CODIFAB%(NORFA$)
   TXTNROPED = NroPedidoSegunOf&(NORFA$)
   TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPED))
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   TXTFTECNICA = ctp.NroFTecnicaXCodigo(CI1%)
   Text11 = CODEXT$(CI1%)
   Text4 = DESCRIT0$(CI1%)
   
   NP& = XVALO(TXTNROPED)
   NC = cliente_SEGUN_PEDIDO(NP&): TXTNC = NC
   RASO$ = rasocli$(NC): Me.TXTRAZOCLI = RASO$

   Call CARGARMSF1
   Call CargarMsf3
   CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
   If (XVALO(Mid$(NORFA$, 4, 6)) >= 2341 Or XVALO(Mid$(NORFA$, 4, 6)) = 2325) And InStr(CODFAMI$, "BB") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then  'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
        'la estructura es la misma solo cambio el titulo de NroInf x CantReg
        Call CARGA_ENCABEZADO(MSF4, "Código/A13;Descripción/A32;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F4;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;CantReg/F6;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0", Me)
        Call ctp.NuevaCargaSaldoOperacionesInformadasALinea(NORFA$, CI1%, MSF4, 100, "", 100, "")
   Else
        Call CARGA_ENCABEZADO(MSF4, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0", Me)
        Call ctp.CargaOperacionesInformadas(NORFA$, CI1%, MSF4, 100, "", 199)
   End If
    
    CONDI$ = "NUMRECETACTP  = " & XVALO(TXTFTECNICA)
    IMPRE$ = TRIM$(VALOBADA("RECETACTP", "IMPRESION", CONDI$, "NOMESS"))
    If IMPRE$ = "SI" Then
        REG& = 0
        REG& = REG& + 1
        MSF2.Rows = REG& + 1
        
        MSF2.TextMatrix(REG&, 1) = "IMPRESIÓN"
        MSF2.TextMatrix(REG&, 2) = ""
        MSF2.TextMatrix(REG&, 3) = ""
        MSF2.TextMatrix(REG&, 4) = ctp.FRMSECTOR("IMPRESION")
    End If
    Set ctp = Nothing
End Sub



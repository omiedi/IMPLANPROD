VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TRANFTALLER 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transferencia de Materiales a Taller "
   ClientHeight    =   9300
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   15735
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9300
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   9975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15735
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
         Index           =   13
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Detalle a Entregar a Taller"
         Top             =   6840
         Width           =   3540
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
         Index           =   12
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Sub-Totales Agrupados x Tipo de Corte"
         Top             =   4755
         Width           =   3540
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   825
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   8655
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "Razón Social"
            Top             =   120
            Width           =   6405
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
            Index           =   7
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "Cuenta Proveedor"
            Top             =   120
            Width           =   1725
         End
         Begin VB.TextBox Text7 
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   360
            Width           =   6405
         End
         Begin VB.CommandButton Command2 
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
            Left            =   1875
            TabIndex        =   40
            ToolTipText     =   "Ordenes de Fabricación Abiertas"
            Top             =   360
            Width           =   210
         End
         Begin VB.TextBox Text6 
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
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Grabar e Imprimir"
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
         Left            =   12660
         TabIndex        =   36
         Top             =   9000
         Width           =   3015
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Posicionar Cursor Para Leer Código QR de Etiqueta de Pallet para Transferir a Linea"
         Top             =   6840
         Visible         =   0   'False
         Width           =   10335
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
         Left            =   14160
         TabIndex        =   10
         Top             =   1800
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
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Materiales en Semiterminados"
         Top             =   1800
         Width           =   3540
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
         TabIndex        =   7
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
         TabIndex        =   6
         ToolTipText     =   "Ordenes de Fabricación Abiertas"
         Top             =   360
         Width           =   210
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         TabIndex        =   5
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
            TabIndex        =   37
            ToolTipText     =   "Ordenes de Fabricación Abiertas y Cerradas"
            Top             =   240
            Visible         =   0   'False
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   860
         End
      End
      Begin VB.Frame Frame5 
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   380
            Width           =   1155
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2715
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Doble Clic Para Transferira a Taller"
         Top             =   2040
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   4789
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
         Height          =   1860
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Doble Clic Para Revertir Ingreso a Linea"
         Top             =   7120
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   3281
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
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1785
         Left            =   120
         TabIndex        =   44
         Top             =   5040
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   3149
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
         TabIndex        =   12
         Top             =   6720
         Visible         =   0   'False
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
      OleObjectBlob   =   "TRANFTALLER.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "TRANFTALLER.frx":0234
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
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAnulTransfTaller 
         Caption         =   "Anular Transferencia a Taller"
      End
   End
   Begin VB.Menu mnuCtasList 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
   End
End
Attribute VB_Name = "TRANFTALLER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub actualizaGrillaAgrupada(MSFALEER As msFlexGrid, MSFAESCRIBIR As msFlexGrid)
' Call CARGA_ENCABEZADO(MSF3, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/a0;Nombre/A0;Peso/f8.2;IdReg/a0;1/A12;2/a12", Me)
' Call CARGA_ENCABEZADO(MSF4, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0;1/A0;2/a0", Me)
' Call CARGA_ENCABEZADO(MSF5, "Código/A16;Descripción/A36;Cant.Cortes/F10.1;N°Op./F5;Descrip/A12;Peso/f8.2,Metros/f8.2", Me)
 Dim OPERACION%()
 ReDim Preserve OPERACION%(1)
 MSFAESCRIBIR.Rows = 1
 i& = 0
    For REG& = 1 To MSFALEER.Rows - 1
        NUMEOPER% = XVALO(MSFALEER.TextMatrix(REG&, 6))
        If UBound(OPERACION%) < NUMEOPER% Then ReDim Preserve OPERACION%(NUMEOPER%)
        If OPERACION%(NUMEOPER%) < 1 Then
           i& = i& + 1
           MSFAESCRIBIR.Rows = i& + 1
           CI1% = CODINT(MSFALEER.TextMatrix(REG&, 9))
           MSFAESCRIBIR.TextMatrix(i&, 1) = MSFALEER.TextMatrix(REG&, 9)
           MSFAESCRIBIR.TextMatrix(i&, 2) = DESCRIT0$(CI1%)
           MSFAESCRIBIR.TextMatrix(i&, 3) = SUMA_CANT_SEGUNDATO(MSFALEER, str$(NUMEOPER%), 6, 4)   'SUMA LOS CORTES SEGUN LA OPERACION
           MSFAESCRIBIR.TextMatrix(i&, 4) = MSFALEER.TextMatrix(REG&, 5)
           MSFAESCRIBIR.TextMatrix(i&, 5) = MSFALEER.TextMatrix(REG&, 6)
           MSFAESCRIBIR.TextMatrix(i&, 6) = MSFALEER.TextMatrix(REG&, 7)
           MSFAESCRIBIR.TextMatrix(i&, 7) = SUMA_CANT_SEGUNDATO(MSFALEER, str$(NUMEOPER%), 6, 15) 'SUMA LOS KG.
            ID_INFOFAB = MSFALEER.TextMatrix(REG&, 16)
            corte = XVALO(VALOBADA("INFOFAB", "CORTE", "INFOFAB_ID= " & ID_INFOFAB, "NOMESS")): If corte <= 0 Then corte = 1
           MSFAESCRIBIR.TextMatrix(i&, 8) = corte * XVALO(MSFAESCRIBIR.TextMatrix(i&, 3)) 'metros = corte x cantidad de cortes

           OPERACION%(NUMEOPER%) = 1
         End If
    Next REG&
End Sub

Sub CargarMsf3()
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
   Call ctp.CargaOperacionesInformadasSegunUbicaciones(NORFA$, CI1%, MSF3, 0, "", 99, "7;8;9;10;11;" & Talleres$, "")
   
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



Sub GrabarEntregaTaller(NORFA$, MSF As msFlexGrid, NPROTALLER%, NROPED&, RETDOCUORIG$)
    '1.-CARGAR EN INFOFAB COMO IDUBICACION = 400 Y TEXTO TALLER(NROTALLER)
    '2.-AGREGAR A MOVISTO LOS REGISTROS DE SALIDA EN LA TABLA MOVISTO
    '3.-AL RECIBIR SE DARA INGRESO DEL PRODUCTO TERMINADO (1 REGISTRO) (AL ENTREGAR NO HABRA REGISTRO DEL  PRODUCTO TERMINADO)
    '4.-AL RECIBIR SE GENERARAN MOVIMIENTOS DE LOS CORTES QUE INGRESARON. SEGUN LA FORMULA.
    '5.-CODIMOVI = ST (SALIDA TALLER)
    'DOSECLAR LA OF, DOPRILAR EL NRO DE TRANSACCION REMITO.
    'PARA AL IMPRESION UTILIZAR EL MISMO FRM DE IMPRIMIR CON EL ENCABEZADO QUE INDIQUE PROVEEDOR, OF. FECHA,FIRMA
    'MOVTALLER
    
    '
    RETDOCUORIG$ = "" 'SI RETORNA CON VALOR ES QUE SE GRABO BIEN
    PRONUM_MOVTALLER& = 1 + XVALO(VALOBADA("MOVTALLER", "Max (CAST(SUBSTRING(DOCUORIG,4,8) AS INTEGER))", "DOCUORIG  like 'RT-%'", "NOMESS"))  'SOLO PARA SALIDAS
    DOCUORIG$ = "RT-" & FORMATOCON0$(PRONUM_MOVTALLER&, 8)
    'DOPRI$ = "ST-" & str$(PRONUM_MOVTALLER) va como norfa$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
    ORDEN% = 0
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    For REG& = 1 To MSF.Rows - 1
       TXTTALLER$ = "TALLER(" & TRIM$(str$(NPROTALLER%)) & ")"
       IDREG& = XVALO(MSF.TextMatrix(REG&, 16))
       SQLX$ = "UPDATE INFOFAB SET IDCOORDUBIC = 400"
       SQLX$ = SQLX$ + " ,COORDUBIC = '" & TXTTALLER$ & "'"
       SQLX$ = SQLX$ + " WHERE INFOFAB_ID= " & IDREG&
       FLEXCONN.Execute (SQLX$)
       
       '
       Codigo$ = MSF.TextMatrix(REG&, 9)
       LOTE$ = MSF.TextMatrix(REG&, 3)
       CANTI = XVALO(MSF.TextMatrix(REG&, 4))
       CI1% = XVALO(MSF.TextMatrix(REG&, 5))
       NUMEOPER = XVALO(MSF.TextMatrix(REG&, 6)) 'SE GUARDA EN ID_OPERACION
       DESCRIOPERACION$ = MSF.TextMatrix(REG&, 7) 'SE GUARDA EN REFERCOR
       NROINFORME% = XVALO(MSF.TextMatrix(REG&, 11)) 'COMO NUMERO DE ORDEN DEL PEDIDO
       IDCOORDUBIC_Orig = XVALO(MSF.TextMatrix(REG&, 12))
       Noperario = XVALO(MSF.TextMatrix(REG&, 13))
       NameOperario = MSF.TextMatrix(REG&, 14)
       Peso# = XVALO(MSF.TextMatrix(REG&, 15))
       ID_INFOFAB& = XVALO(MSF.TextMatrix(REG&, 16)) 'guardadado en tabla como IDCALIDAD
       corte = XVALO(VALOBADA("INFOFAB", "CORTE", "INFOFAB_ID= " & ID_INFOFAB, "NOMESS")): If corte <= 0 Then corte = 1
       
       ORDEN% = ORDEN% + 1
       Call MOVISTOK_MOVTALLER(HOY(HO$), CI1%, LOTE$, "ST", NORFA$, NORFA$, DOSEC$, DOSEC$, DESCRIOPERACION$, ORDEN, 0, 0, (-1) * NPROTALLER%, NPROTALLER%, (-1) * CANTI, NROPED&, NROINFORME%, "Entrega Taller a Proveedor: " & NPROTALLER%, DOCUORIG$, 0, 0, 0, "", ID_INFOFAB&, 0, CANTI * corte, Peso#, 0, XVALO(TXTNROCOTIZA), NUMEOPER)
    Next REG&
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      RETDOCUORIG$ = DOCUORIG$
      Call COMUNI("Transacción Realizada Exitosamente")
    End If


End Sub

Sub MOVISTOK_MOVTALLER(FECHA%, CI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NCLIE, DEPO%, CANTIX, Optional NUPECLIEN&, Optional NORITPE%, Optional DELARGA$, Optional DOCOMPRA$, Optional IDMOVIM$, Optional PRECONIVA#, Optional TIPOTRAN%, Optional ORDCOM_RTOABIER$, Optional IdCalidad&, Optional IDOPERACION%, Optional METROS#, Optional PESOKG#, Optional DOPRINUM1&, Optional NROCOTIZACION&, Optional NUMEOPER)
   '
   '
   SMP = Screen.MousePointer
   DOPRILA1$ = TRIM$(DOPRILA$)
   Dim Movisto As ADODB.Recordset
   Set Movisto = New ADODB.Recordset
   TABLAMR$ = "MOVTALLER"
   strSQL = "Select *  from " & TABLAMR$
   strSQL = strSQL + " WHERE DoPriLar = '" & DOPRILA1$ & "'"
25 On Error Resume Next
   'Movisto.Open strsql, FLEXCONN, , , adCmdText
   Movisto.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With Movisto
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CODIEMPR = CodiEmp%
     On Error GoTo 0
     !NUMPVTA = NUMPVTA1%
     !DOPRINUM = DOPRINUM1&
     !FechMov = CVDAT(FECHA%)
     !cod_inte = CI%
     !Cod_Exte = Left$(CODEXT$(CI%), 24)
        DESCIRIX1$ = Left(DESCRIT0$(CI%), 64)
     !Descrip = DESCIRIX1$
     !idenlote = Left$(LOTE$, 16)
     !CODIMOVI = Left$(CMOV$, 2)
     !dopricor = Left$(DOPRICO$, 12)
       DOPRILAX$ = TRIM$(DOPRILA$): If DOPRILAX$ = "" Then DOPRILAX$ = DOPRICO$
     !DOPRILAR = Left(DOPRILAX$, 24)
     !DoSecCor = Left$(DOSECO$, 12)
       DOSECLAX$ = DOSELA$: If DOSECLAX$ = "" Then DOSECLAX$ = DOSECO$
     !DOSECLAR = Left(DOSECLAX$, 24)
     !NuOrItem = NUORIT%
     !DOCUORIG = Left$(DOCOMPRA$, 24)
     !IdenMovi = Left$(IDMOVIM$, 32)
     !NUPEDCLI = 0: !NUMEORDPED = 0
     !DESCRI_LAR = DELARGA$ ' DESCRIPCION DE TEXTO LIBRE
        If TRIM$(DELARGA$) = "" Then !DESCRI_LAR = DESCIRIX1$
       If Left$(DOSECO$, 3) = "PC." Then
           !NUPEDCLI = XVALO(Mid$(DOSECO$, 4))
         Else
           On Error Resume Next
           !NUPEDCLI = NUPECLIEN
           !NUMEORDPED = NORITPE%
           On Error GoTo 0
       End If
     '\\\OT-08-0751-OD-12/09/08///
     
     '\\\OT-08-0751-OD-FIN///
        If Left$(REFE$, 7) = "RECHAZO" Or Left$(REFE$, 8) = "FALTANTE" Then
          locon$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(Abs(CANTIX), "F8." & CNTDC$ & "")
          !LotConsum = locon$
        End If
     !VALOUNIT = VUNI#
     !ValTotIt = ROUND(VUNI# * (ROUND(Abs(CANTIX), CADEC%)), 2)
     !MoneVal = MONEII%
       On Error Resume Next      ' POR SI NO ESTA DEFINIDO EL CAMPO
     !TIPOCAM = TICAMONE(MONEII%)
       On Error GoTo 0
     !DEPOMOVI = DEPOXI%
     If NCLIE >= 0 Then
         !NUME_CLIE = NCLIE
         !Nume_Prov = 0
       Else
         !NUME_CLIE = 0
         !Nume_Prov = Abs(NCLIE)
     End If
     If CANTIX >= 0 Then
         !CANTINGRE = ROUND#(CANTIX, CADEC%) ' CADEC% = CANTIDAD DE DECIMALES CONFIGURADOS
         !CANTSALID = 0
       Else
         !CANTINGRE = 0
         !CANTSALID = ROUND#(Abs(CANTIX), CADEC%)
     End If
     !CantSaldo = 0
     !FeReMovi = Now
     !NumUsuar = USERN%
     !VALOUNITIVA = PRECONIVA#
     !ValTotItIva = ROUND(PRECONIVA# * (ROUND(Abs(CANTIX), CADEC%)), 2)
     !VERIFICADOPOR = 0
     !FHor_Verif = FETEXCOR1$(CVINT(NUNCA$))
     On Error Resume Next
     !TIPOTRAN = TIPOTRAN% 'TIPO DE TRANSFERENCIA AGREGADO POR ALEJPA
     On Error GoTo 0
     !OCOMPRA_RTOABIERTO = Left$(ORDCOM_RTOABIER$, 64)
     
     !IdCalidad = IdCalidad&
     !ID_OPERACION = IDOPERACION%
     
     If PESOKG# > 0 Or METROS# > 0 Then
        !PESOKG = PESOKG#
        !METROS = METROS#
     End If
     !NROCOTIZACION = XVALO(NROCOTIZACION&)
     !ID_OPERACION = XVALO(NUMEOPER)
     !ReferCor = Left(REFE$, 64)
     .Update
   End With
   '
   Screen.MousePointer = SMP
   '
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
    Command1.Enabled = False
    '
    NORFA$ = TXTOT
    Campos$ = "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0"
    Call CARGA_TABLA_IMPRE(MSF3, Campos$, "INFOUBIC", "Operaciones x O.F. Sector Previo a Linea", "TABLA_IMPRE", , 1)
    '*****************************************************'
    FILTX$ = TRIM$(FECHATEX$(HOY(HO$))) & ";" & TRIM$(NORFA$)
    Call STDFORM("UBIK-CTP", FILTX$)
    Command1.Enabled = True
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
    NCXI% = ListaSeleccionProveedores()
    If NCXI% > 0 Then
        NC% = NCXI%
        Text6.TEXT = TRIM$(str$(NC%))
        Text7.TEXT = rasocli$((-1) * NC%)
    End If
    
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    If XVALO(Text6) < 1 Then
       Call COMUNI("Falta Cta de Proveedor (Taller)")
       GoTo 99
    End If
    
    Call GrabarEntregaTaller(TRIM$(TXTOT), Me.MSF4, XVALO(Text6), XVALO(TXTNROPED), RETDOCUORIG$)
    If RETDOCUORIG$ = "" Then GoTo 99
    NORFA$ = TXTOT
    Campos$ = "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0;1/A0;2/a0"

    Call CARGA_TABLA_IMPRE(MSF4, Campos$, "RtoEntTa", "Remito de Entrega a Taller", "TABLA_IMPRE", , 1)
    '*****************************************************'
    FILTX$ = TRIM$(FECHATEX$(HOY(HO$))) & ";" & TRIM$(NORFA$) & ";" & TRIM$(RETDOCUORIG$) & ";" & TRIM$(Text11) & ";" & TRIM$(Text4)
    
    Call STDFORM("RTAL-CTP", FILTX$)
    MSF4.Rows = 1
    MSF5.Rows = 1
99  Command3.Enabled = True
End Sub



Private Sub Form_Load()
 Call OFABRI07.CrearTablaMovTaller
   
 Call CARGA_ENCABEZADO(MSF3, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/a0;Nombre/A0;Peso/f8.2;IdReg/a0;1/A0;2/a0", Me)
 Call CARGA_ENCABEZADO(MSF4, "Código/A16;Descripción/A36;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/F5;IDUbic/a0;N°Operario/A0;Nombre/A0;Peso/f8.2;IdReg/a0;1/A0;2/a0", Me)
 Call CARGA_ENCABEZADO(MSF5, "Código/A16;Descripción/A36;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Peso/f8.2;Metros/f8.2", Me)

End Sub

'
Private Sub MSF1_Click()
'    REG& = MSF1.MouseRow
'    J& = MSF1.MouseCol
'    If REG& = 0 Or REG& > MSF1.Rows Then Exit Sub
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbYellow, REG&)
''    REG& = MSF1.MouseRow
''    j& = MSF1.MouseCol
''    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
''    DESCOPERACION$ = MSF1.TextMatrix(REG&, 1)
''    Dim CTP As New Conteplast
''    TXTOPERACIONSELECCIONADA = DESCOPERACION$
''    Call CTP.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF3, 0, TXTOPERACIONSELECCIONADA, 99)
''    Call CTP.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF4, 100, "", 199)
''
''    Set CTP = Nothing

End Sub

Private Sub MSF1_DblClick()
'    REG& = MSF1.MouseRow
'    j& = MSF1.MouseCol
'    If REG& = 1 Or REG& > MSF1.Rows Then Exit Sub
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbYellow, REG&)
'    Call printformu(MSF1, REG&, j&)
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbWhite, REG&)


End Sub


Private Sub MSF2_Click()
'    REG& = MSF2.MouseRow
'    j& = MSF2.MouseCol
'    If REG& = 0 Or REG& > MSF2.Rows Then Exit Sub
'    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF2, vbYellow, REG&)
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
   Set RS = READSET(SQLX$)
   With RS
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
   
   Set RS = READSET(SQLX$)
   With RS
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


Private Sub mnuAnulTransfTaller_Click()
'   1.-cargar zelecho
    '2.-VA A VER QUE AGREGAR UN CAMPO PARA INDICAR SI LA OF REAL SE CERRO Y EN BASE A ESTO GENERAR
    'ANULACION DE LOS ENVIOS A TALLER POR EL NRO DE REMITO
    'MODIFICAR A NRO 100 LOS LOTES INFORMADOS QUE SE FUERON EN EL INFOFAB CON EL NRO DEL CAMPO IDCALIDAD
    Dim ctp As New Conteplast
    UBICACION$ = ctp.UBICACIONSEGUNID(3)

10  NROREMITO$ = TRIM$(ctp.ListaSelecRemitosTallerConteplast$)

    If NROREMITO$ = "" Then Exit Sub
    NORFA$ = TRIM(RESP_ZELE_VECT(1))
    OPX% = COMALTER%("Confirma Eliminar de Manera Permanente Los Materiales  Enviados con Remito: " & NROREMITO$ & "\\Cancelar\Eliminar")
    If OPX% < 2 Then Exit Sub
    
    If CantRegistros("MOVTALLER", "CODIMOVI = 'DT' AND DOPRILAR = '" & NORFA$ & "'", "NOMESS") > 0 Then
       Call COMUNI("Imposible Anular!!\El Remito Seleccionado\Està Vinculado a una O.F. que yà Posee Ingresos del Proveedor")
       Exit Sub
    End If
    
    LISTAIDINFOFAAB$ = ctp.ListaIdInfofabSegunRemitoMovTaller$(NROREMITO$) 'TRAE LA LISTA DE LOS ID DE INFOFAB
    CONDI$ = "INFOFAB_ID IN( " & LISTAIDINFOFAAB$ & ")"
    Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, 3, REGAF&, "NOMESS") 'LO VUELVE A 100
    Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, UBICACION$, REGAF&, "NOMESS") 'LE ASIGNA SU TEXTO CORRESPONDIENTE A EL IDUBIC 100
    
    CONDI$ = "DOCUORIG = '" & NROREMITO$ & "'"
    Call BORRAREGISTROS("MOVTALLER", CONDI$, REGAF&, "NOMESS")
    
    Call COMUNI("Movimientos de Remito: " & NROREMITO$ & " Eliminados")
    Call TXTOT_Change
    GoTo 10
    
End Sub

Private Sub MSF3_DblClick()
    REG& = MSF3.MouseRow
    J& = MSF3.MouseCol
    If REG& < 1 Or REG& >= MSF3.Rows Then Exit Sub
    Dim MSF As msFlexGrid
    Set MSF = MSF3
    '
    
    Call AGREGAR_FILA_OTRO_MSF(MSF3, MSF4, REG&)
    Call actualizaGrillaAgrupada(MSF4, MSF5)  'REG POSITIVO  INCREMENTA
    Call cmdEliminaItem(MSF3, REG&)

    Exit Sub

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
    Call AGREGAR_FILA_OTRO_MSF(MSF4, MSF3, REG&)
    Call cmdEliminaItem(MSF4, REG&)
    Call actualizaGrillaAgrupada(MSF4, MSF5)  'NEGATIVO DESCUENTA

    Exit Sub
    
'codigo anterior
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


Private Sub MSFlexGrid1_Click()

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
   MSF3.Rows = 1
   MSF4.Rows = 1
   MSF5.Rows = 1
   
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
   
   CIFAB% = CODINT%(Text11)
   Call ctp.CargaOperacionesInformadasMasSubProcesos(NORFA$, CIFAB%, MSF3, 3, "", 3)  'SOLO SEMITERMINADOS
   
   Set ctp = Nothing

'   Call CARGARMSF1
'   Call CargarMsf3
'   Call CTP.CargaOperacionesInformadas(TXTOT, XVALO(TXTCODINT), MSF4, 100, "", 199)
    
'    CONDI$ = "NUMRECETACTP  = " & XVALO(TXTFTECNICA)
'    IMPRE$ = TRIM$(VALOBADA("RECETACTP", "IMPRESION", CONDI$, "NOMESS"))
'    If IMPRE$ = "SI" Then
'        REG& = 0
'        REG& = REG& + 1
'        MSF2.Rows = REG& + 1
'
'        MSF2.TextMatrix(REG&, 1) = "IMPRESIÓN"
'        MSF2.TextMatrix(REG&, 2) = ""
'        MSF2.TextMatrix(REG&, 3) = ""
'        MSF2.TextMatrix(REG&, 4) = CTP.FRMSECTOR("IMPRESION")
'    End If
End Sub



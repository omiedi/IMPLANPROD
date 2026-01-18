VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SELPRINOT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Impresión de Formularios de Corte"
   ClientHeight    =   9105
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   12015
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9105
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Grabar Datos Adicionales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   8880
         TabIndex        =   46
         ToolTipText     =   "Graba Producción Estandar,Anexos y Observaciones Adicionales"
         Top             =   6440
         Width           =   3015
      End
      Begin VB.TextBox Text7 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Producción estandar"
         Top             =   5760
         Width           =   2235
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   42
         Top             =   6030
         Width           =   2235
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
         Left            =   2340
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   6030
         Width           =   9615
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
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Anexos"
         Top             =   5760
         Width           =   9615
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Imprimir Liners"
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
         Height          =   315
         Left            =   3600
         TabIndex        =   34
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Archivos Adjuntos Asociados a la Ficha Técnica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   8280
         Width           =   11835
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir Ficha Técnica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   384
            Left            =   7560
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir Ordenes de Corte Seleccionadas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   384
            Left            =   9480
            TabIndex        =   38
            Top             =   240
            Width           =   2250
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   300
            Width           =   6045
         End
         Begin VB.CommandButton Command45 
            Caption         =   "Ver"
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
            Left            =   6240
            TabIndex        =   32
            Top             =   300
            Width           =   850
         End
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
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Impresión"
         Top             =   4470
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Imprimir Etiqueta"
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
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Imprimir Doblado"
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
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   11835
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir Orden de Fabricación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   384
            Left            =   3360
            TabIndex        =   37
            Top             =   240
            Width           =   3015
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
            Left            =   11265
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "N°.O."
            Top             =   180
            Width           =   500
         End
         Begin VB.TextBox TXTNORPED 
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
            Left            =   11265
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   420
            Width           =   500
         End
         Begin VB.CommandButton Command15 
            Appearance      =   0  'Flat
            Caption         =   "Imprimir Orden de Corte General"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   384
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   3015
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   420
            Width           =   1155
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
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   23
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
            Index           =   3
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "Pedido"
            Top             =   180
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Cód.Interno"
            Top             =   180
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   20
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
            Index           =   5
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Cotización"
            Top             =   180
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
            Left            =   10125
            Locked          =   -1  'True
            TabIndex        =   18
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
            Index           =   6
            Left            =   10125
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "F.Técnica"
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   11835
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
            Left            =   10620
            Locked          =   -1  'True
            TabIndex        =   15
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
            Index           =   2
            Left            =   10620
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Cant.Aprob."
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
            Index           =   1
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Cant.Proyec."
            Top             =   150
            Width           =   1140
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
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   150
            Width           =   6015
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
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   6015
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
            Left            =   8340
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   420
            Width           =   2235
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   150
            Width           =   2235
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
            Left            =   8340
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   150
            Width           =   1140
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2745
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   4842
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   945
         Left            =   120
         TabIndex        =   26
         Top             =   4800
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1667
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   6840
         Width           =   11820
      End
      Begin VB.TextBox TXTLABEL1 
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
         Height          =   330
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Observaciones Adicionales"
         Top             =   6480
         Width           =   3050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   4440
         Width           =   10905
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
      OleObjectBlob   =   "SELPRINOT.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELPRINOT.frx":0234
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
      Begin VB.Menu mnuRegenerarOperfab 
         Caption         =   "RegenerarOperfab"
      End
   End
End
Attribute VB_Name = "SELPRINOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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

Private Sub Check3_Click()
   Check3.ForeColor = vbYellow
   
   If Check3.Value = 1 Then
        Check3.ForeColor = vbYellow
        NROFT& = XVALO(TXTFTECNICA)
        LINER_SI_NO$ = UCase$(TRIM$(VALOBADA("RECETACTP", "LINERSINO", "NUMRECETACTP = " & NROFT&, "NOMESS")))
        If LINER_SI_NO$ <> "SI" Then
          Call COMUNI("La Orden de Fabricación Seleccionada \Referencia una Ficha Técnica Sin Liner\Se Podrá Imprmir igual Pero con Datos Incompletos")
        End If
   End If
End Sub

Private Sub Command1_Click()
   NRO = XVALO(TXTFTECNICA)
   If NRO < 1 Then Exit Sub
   COTICONT.Text5 = NRO
   COTICONT.mnuimprimirfichaactiva_Click
   Unload COTICONT
End Sub

 Sub Command15_Click()
      NROCOT& = XVALO(TXTNROCOTIZA)
      NORFA$ = Text10
      Campos$ = "Tipo de Corte/a24;Descrip.M.P./A36;Cantidad/F9;Med.Corte(Cm)/F13.1;Cons.Tot.(Mts)/F14.1"
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 18000)
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , 18000)
      FRMZELECHO.Width = 19000
      FRMZELECHO.Caption = ""
      
      SQLX$ = "SELECT COTI1.COMPONENTE, COTI1.CODINTELEM, COTI1.CORTE, ORDEFABR.CanProy,COTI1.MODUXBB"
      SQLX$ = SQLX$ + " FROM COTI1 WITH(NOLOCK)"
      SQLX$ = SQLX$ + " INNER JOIN ORDEFABR ON COTI1.CODINTCONJ = ORDEFABR.Cod_Inte "
      SQLX$ = SQLX$ + " WHERE COTI1.NROCOTIZACION =  " & NROCOT&
      SQLX$ = SQLX$ + " AND COTI1.CODINTELEM > '0' AND ORDEFABR.IDSUBOR = '" & NORFA$ & "'"
      '****************************************************************
      Screen.MousePointer = 11
      Set rs = READSET(SQLX$)
      With rs
        i& = 0
        Do While Not .EOF
          i& = i& + 1
          FRMZELECHO.MSF2.Rows = i& + 1
          FRMZELECHO.MSF2.TextMatrix(i&, 1) = SAFETEXT$(!COMPONENTE)
          CIMP% = XVALO(!CODINTELEM) 'CODIGO DE MATERIA PRIMA GUARDADO
          CIRANGO = CODIGORANGO&(CIMP%)
          If CIRANGO > 0 Then 'SI TIENE CODIGO RANGO
             CIMP% = CIRANGO   'ASIGNO
          End If
          CANTI = ROUND(XVALO(!CANPROY) * XVALO(!moduxbb))
          FRMZELECHO.MSF2.TextMatrix(i&, 2) = DESCRIT0$(CIMP%)   ' ES EL CODIGO COMPONENTE O EL RANGO SEGUN VALIDACION
          FRMZELECHO.MSF2.TextMatrix(i&, 3) = CANTI
          FRMZELECHO.MSF2.TextMatrix(i&, 4) = ROUND(XVALO(!corte))
            'CONSUMO = XVALO(!Corte) * XVALO(!CANPROY)
            CONSTOTAL = (CANTI * XVALO(!corte)) / 100 ' POR QUE LO LLEVO A METROS
          'FRMZELECHO.MSF2.TextMatrix(i&, 5) = CONSUMO
          FRMZELECHO.MSF2.TextMatrix(i&, 5) = CONSTOTAL
         .MoveNext
        Loop
       End With
       Set obj = Nothing
       rs.Close
       Set rs = Nothing
'       FRMZELECHO.Width = 16500
'       FRMZELECHO.msf2.Font.Size = 9
'       FRMZELECHO.MSF1.Font.Size = 9
'       FRMZELECHO.mnuImprimir.Visible = True
'       Screen.MousePointer = 1
'
'60     FRMZELECHO.Show 1
'       ARGUBUS$ = TRIM$(FRMZELECHO.TXTBUSCAR)
'       If FRMZELECHO.IMPRIMIR > 0 Then
        Dim MSF As msFlexGrid
        Set MSF = FRMZELECHO.MSF2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOGRAL", "ORDEN DE CORTE GENERAL: " & NROCOT&, "TABLA_IMPRE", RET_NAMECONS$, 1)
        '*****************************************************'
       NROFTEC& = XVALO(TXTFTECNICA)
       FORIPRE$ = Control("", "FORMORFA")
       FILTX$ = TRIM$(Text10) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&))
       Call STDFORM(FORIPRE$, FILTX$)
End Sub

Private Sub Command2_Click()
    CICONJ% = XVALO(TXTCODINT)
    If CICONJ% < 1 Then GoTo 99
   
    CI1% = CODINT%(Text11)
    If CI1% < 1 Then
     Call COMUNI("Falta Seleccionar Código de Tela")
     GoTo 99
    End If
    '
    If XVALO(TXTNROCOTIZA.TEXT) < 1 And XVALO(TXTFTECNICA.TEXT) < 1 Then
     'IMPRESION PARA BOLSAS ENVAPLAST QUE NO TIENEN FICHA TENCNCIA NI COTIZACION
     If TRIM$(UCase$(CODFAMIL$(CICONJ%))) = "BO" Then
        FILTX$ = TRIM$(Text10) & ";" & TRIM$(str$(CICONJ%))
        Call STDFORM("ODFA-EYM", FILTX$, "PRINT")
        Exit Sub
     End If
   End If


    
    If XVALO(TXTFTECNICA) <= 0 Then
     Call COMUNI("Falta N°: de Ficha Técnica)")
     GoTo 99
    End If

    '
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")

    '
    RUTAEXCEL$ = Control("PLANVENT", "IMPOFBIG")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    Dim CTP As New Conteplast
    NPED& = XVALO(TXTNROPED)
    ORDENPEDIDO% = XVALO(TXTNORPED)
    NC = cliente_SEGUN_PEDIDO(NPED&)
    RASO$ = rasocli$(NC)
    NROFT& = XVALO(TXTFTECNICA)
    FechaEntrega% = CVINT(VALOBADA("PEDDETA", "FESOLENT", "NROPED=" & NPED& & " AND NUORITEM =" & ORDENPEDIDO%, "NOMESS"))
    FechaEmisi% = CVINT(VALOBADA("pedenca", "Fechemis", "NROPED=" & NPED&, "NOMESS"))
    CODIESPECIFICLIENTE$ = CTP.CodEspTecCliente(CI1%)
    CODIGOBB$ = CTP.CODIGOBIGBAG(CI1%)


    ORCOMPRA$ = VALOBADA("PEDENCA", "NROOCOM", "NROPED= " & NPED&, "NOMESS")
    NVEND% = XVALO(VALOBADA("PEDENCA", "Vended", "NROPED= " & NPED&, "NOMESS"))
    NOMBVENDEDOR$ = ECOARCH$("VENDEDOR", Chr$(NVEND%))
    TIPOPALLEX$ = CTP.TIPOPALLET(CI1%)
    EMBUFARDOX$ = CTP.EMBUFARDO(CI1%)
    UNXPALLEX = CTP.UNXPALLET(CI1%)
    OBSGENERALX$ = CTP.OBSGENERAL(CI1%)
    JUSTAX$ = CTP.JUSTA(CI1%)
    
    With VENTABNEW.MSF
    .Rows = 22: .Cols = 2
    .TextMatrix(1, 1) = Text10 'Orden de fabricación 1
    .TextMatrix(2, 1) = Text11 'Código bolsa 1
    .TextMatrix(3, 1) = NROFT& 'F.Técnica Nº1
    .TextMatrix(4, 1) = RASO$ 'Razón Social 1
    .TextMatrix(5, 1) = FECHAFORMATOLARGO$(FechaEntrega%)  'Fecha de entrega1
'    .TextMatrix(5, 1) = FechaEntrega%
    .TextMatrix(6, 1) = XVALO(Text3)
    .TextMatrix(7, 1) = NPED& ''Nota de Pedido 1
    .TextMatrix(8, 1) = CODIESPECIFICLIENTE$ 'orden de compra
    .TextMatrix(9, 1) = ORCOMPRA$ 'Nota de Pedido 1
    .TextMatrix(10, 1) = NOMBVENDEDOR$  'Vendedor 1
    .TextMatrix(11, 1) = CODIGOBB$ 'CODIGO DE CONFECCION 1
    
    .TextMatrix(12, 1) = XVALO(Text6) 'PRODUCCION ESTANDAR
    .TextMatrix(13, 1) = Text5 'ANEXO
    
    .TextMatrix(14, 1) = TIPOPALLEX$ ''Tipo de pallet 14
    .TextMatrix(15, 1) = EMBUFARDOX$ 'U / Fardo      15
    .TextMatrix(16, 1) = UNXPALLEX 'U / Pallet     16
    .TextMatrix(17, 1) = OBSGENERALX$ 'Observaciones Adicionales  18
    
    .TextMatrix(18, 1) = Text28 'SE AGREGA EN PANTALLA OBSERVACIONES ADICIONALES
    
    .TextMatrix(19, 1) = JUSTAX$ 'Cantidad justa               19
    .TextMatrix(20, 1) = FECHATEX$(FechaEmisi%)  'FECHA NP1
'    .TextMatrix(20, 1) = FechaEmisi%  'FECHA NP1

    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "Hoja1", VENTABNEW.MSF, PDES$, "1", 1, 8, "", 1, 0)                       'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    Set CTP = Nothing
99  Command3.Enabled = True
End Sub




Private Sub Command3_Click()

   Command3.Enabled = True
   If CODINT(Text11) < 1 Then
      Call COMUNI("Código No Válido o Inexistente")
      GoTo 99
   End If
   'codigo editable, si lo deja vacio pone el codigo real
   If Check1.Value > 0 Then
      CODIGOEDITABLE$ = TRIM$(InputBox$("Informar", "Ingrese Código", Text11))
      If CODIGOEDITABLE$ = "" Then
         CODIGOEDITABLE$ = Text11
      End If
   End If
   
   CONDI$ = "CODIMASTER = '" & Text11 & "'"
   NROFTEC& = Me.TXTFTECNICA
   NROCOT& = XVALO(Me.TXTNROCOTIZA)
   NORFA$ = Me.Text10
   For i& = 1 To MSF1.Rows - 1
      If TRIM$(MSF1.TextMatrix(i&, 0)) = "*" Then
         If TRIM$(MSF1.TextMatrix(i&, 1)) = "IMPRESION" Then
           FORIPRE$ = TRIM$(MSF1.TextMatrix(i&, 4))
           FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&))
           Call STDFORM("OCCU-CTP", FILTX$)
         Else
           CI1% = XVALO(MSF1.TextMatrix(i&, 5)) 'CODIGO VISIBLE EN LA GRILLA
           CI2% = XVALO(MSF1.TextMatrix(i&, 6)) 'CODIGO ORIGINAL DE LA COTIZACION
           CIRANGO% = CODIGORANGO&(CI1%)        'SI EL CODIGO VISIBLE ES CODIGO M.PRIIMA ASOCIADO A UN CODIGO RANGO
           If CIRANGO% > 0 Then
             CODIGORANGOX1$ = CODEXT$(CIRANGO%)
           Else
             If ESCODIGORANGO&(CI1%) > 0 Then  'SI CODIGO RANGO MUESTRA ESTE
                CODIGORANGOX1$ = CODEXT$(CI1%)
             Else
                CODIGORANGOX1$ = "----"        'SI NO NO IMPRIME NADA POR QUE SOLO DEBE MOSTRAR CODIGOS RANGO NO ORIGINALES
             End If
           End If
           
           If CI2% > 0 Then
              CODIGOORIGINAL$ = CODEXT$(CI2%)
           Else
              CODIGOORIGINAL$ = CODEXT$(CI1%)
           End If
           
           If TRIM$(CODIGOORIGINAL$) = "" Then CODIGOORIGINAL$ = "-----" 'PARA QUE NO APAREZCA LEYENDA COMPAR-1
           
           FORIPRE$ = TRIM$(MSF1.TextMatrix(i&, 4))
           FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&)) & ";" & TRIM$(CODIGORANGOX1$) & ";" & TRIM$(CODIGOORIGINAL$)
           Call STDFORM(FORIPRE$, FILTX$)
           End If
      End If
   Next i&
   If MSF2.Rows > 1 Then
      If TRIM$(MSF2.TextMatrix(1, 0)) = "*" Then
           FORIPRE$ = TRIM$(MSF2.TextMatrix(1, 4))
           FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(str$(NROFTEC&)) & ";" & TRIM$(str$(NROCOT&))
           Call STDFORM(FORIPRE$, FILTX$)
      End If
   End If
   
   'ETIQUETA
   CI1% = CODINT(Text11)
   If Check1.Value > 0 Then
      Call PRINTETIQUETACTP_DESDEOF(CI1%, NORFA$, CODIGOEDITABLE$)
   End If
   
   'DOBLADO
   If Check2.Value > 0 Then
      Call PRINTDOBLADO(CI1%, NORFA$)
   End If
   '
'   LINERS
   If Check3.Value > 0 Then
     Call PRINTLINERS(CI1%, NORFA$)
   End If

99  Command3.Enabled = True
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

    Dim CTP As New Conteplast
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
    Unload VENTABNEW
End Sub

Sub PRINTLINERS(CI1%, NORFA$)
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
    
    RUTAEXCEL$ = Control("COTICONT", "EXCLINER")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    Dim CTP As New Conteplast
    NPED& = XVALO(TXTNROPED)
    ORDENPEDIDO% = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
    NC = cliente_SEGUN_PEDIDO(NPED&)
    RASO$ = rasocli$(NC)
    NROFT& = XVALO(TXTFTECNICA)
    FechaEntrega% = CVINT(VALOBADA("PEDDETA", "FESOLENT", "NROPED=" & NumeroPedido& & " AND NUORITEM =" & ORDENPEDIDO%))
    Cantidad = XVALO(VALOBADA("PEDDETA", "CANPED", "NROPED = " & NPED&, "NOMESS"))
    CODCONFECCION$ = CTP.CODIGOBIGBAG(CI1%)
    DescripBIGBAG$ = CTP.DESCRIBIGBAG(CODCONFECCION$)
    Fuelle_LINER$ = VALOBADA("RECETACTP", "FUELLELINERCU", "NUMRECETACTP = " & NROFT&, "NOMESS")
    ANCHO_LINER$ = VALOBADA("RECETACTP", "LINERANCHO", "NUMRECETACTP = " & NROFT&, "NOMESS")
    ALTO_LINER$ = VALOBADA("RECETACTP", "LINERALTO", "NUMRECETACTP = " & NROFT&, "NOMESS")
    TipoLINER$ = VALOBADA("RECETACTP", "LINERTIPO", "NUMRECETACTP = " & NROFT&, "NOMESS")
    Micronaje_LINER$ = VALOBADA("RECETACTP", "LINERMICRON", "NUMRECETACTP = " & NROFT&, "NOMESS")
    LINER_SI_NO$ = VALOBADA("RECETACTP", "LINERSINO", "NUMRECETACTP = " & NROFT&, "NOMESS")

    '
    With VENTABNEW.MSF
       .Rows = 14: .Cols = 2
       .TextMatrix(1, 1) = NORFA$
       .TextMatrix(2, 1) = RASO$
       .TextMatrix(3, 1) = CODEXT$(CI1%)
       
       .TextMatrix(4, 1) = FECHATEX$(FechaEntrega%)
       .TextMatrix(5, 1) = Cantidad
       .TextMatrix(6, 1) = CODCONFECCION$
      
       .TextMatrix(7, 1) = DescripBIGBAG$
       .TextMatrix(8, 1) = Fuelle_LINER$
       .TextMatrix(9, 1) = ANCHO_LINER$
      
       .TextMatrix(10, 1) = ALTO_LINER$
       .TextMatrix(11, 1) = TipoLINER$
       .TextMatrix(12, 1) = Micronaje_LINER$
      
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "ORDEN_PROVISION_LINER", VENTABNEW.MSF, PDES$, "1", 1, 19, "", 1, 0)                       'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    Set CTP = Nothing
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

Private Sub Command4_Click()
   Command4.Enabled = False
   NORFA$ = TRIM$(Text10)
   If NORFA$ = "" Then GoTo 99
'   On Error GoTo ERROR_GUARDAR
'   FLEXCONN.BeginTrans
'LE QUITE POR QUE ESTABA DEVOLVIENDO UN ERROR QUE SUPERO EL MAXIMO DE TRANSACCIONES
'EN OBSERVACION EL BEGIN

   SQLX$ = "UPDATE ORDEFABR SET ProdEstandar = " & XVALO(Text6)
   SQLX$ = SQLX$ + ", ANEXOS = '" & Text5 & "'"
   SQLX$ = SQLX$ + ", OBSERVA = '" & Left$(Text28, 1024) & "'"
   SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
   FLEXCONN.Execute (SQLX$)

   
'ERROR_GUARDAR:
'    If Err <> 0 Then
'       MsgBox Err.Description
'       On Error Resume Next
'       FLEXCONN.RollbackTrans
'       On Error GoTo 0
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
'       Screen.MousePointer = 1
'       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
'      Call COMUNI("Grabación Realizado Exitosamente")
'     End If
   
99 Command4.Enabled = True
End Sub

Private Sub Command45_Click()
   Command45.Enabled = False

   If Combo2.ListIndex >= 0 Then
     Linea$ = Combo2.List(Combo2.ListIndex)
     Archi$ = TRIM$(Mid$(Linea$, 25))
     If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
        GoTo 99
     End If
   Else
      Call COMUNI("Atención No Hay Ningún Archivo Seleccionado")
      GoTo 99
   End If
    
99 Command45.Enabled = True
End Sub


Private Sub mnuRegenerarOperfab_Click()
   NORFA$ = TRIM$(Text10)
   If NORFA$ = "" Then
      Call COMUNI("Debe Haber Una Orden de Fabricación Activa")
      Exit Sub
   End If
   If CLACONTRA < 1 Then Exit Sub
   If CantRegistros&("INFOFAB", "IDSUBOR = '" & NORFA$ & "'", "NOMESS") Then
      OPX% = COMALTER%("Atención !! Ya hay Informes en Operaciones Informadas\\Cancelar\Regenerar Operaciones")
      If OPX% < 2 Then Exit Sub
   End If
   Call REGENOPER(NORFA$)

End Sub

Private Sub MSF1_Click()
    Reg& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If Reg& = 0 Or Reg& > MSF1.Rows Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, vbYellow, Reg&)

End Sub

Private Sub MSF2_Click()
    Reg& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If Reg& = 0 Or Reg& > MSF2.Rows Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF2, vbYellow, Reg&)
End Sub


Private Sub Text10_Change()
   NORFA$ = TRIM$(Text10)
   Text3 = CantiProyectadaOF(NORFA$)
   Text1 = CantiAprobadaOF(NORFA$)
   TXTNROPED = NroPedidoSegunOf&(NORFA$)
   ORDENPEDIDO% = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
   Me.TXTNORPED = ORDENPEDIDO%
   TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPED), ORDENPEDIDO%)
   
   CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
   ANEXOS = VALOBADA("ORDEFABR", "ANEXOS", CONDI$, "NOMESS")
   PRODSTAN = XVALO(VALOBADA("ORDEFABR", "ProdEstandar", CONDI$, "NOMESS"))
   OBSERVACIONESADI$ = VALOBADA("ORDEFABR", "Observa", CONDI$, "NOMESS")
   
   Text5 = ANEXOS
   Text28 = OBSERVACIONESADI$
   Text6 = FORMATNUM$(PRODSTAN, "F12.0")
  Call CARGARMSF1
End Sub

Private Sub Text6_LostFocus()
   Text6 = FORMATNUM$(XVALO(Text6), "f12")
End Sub

Private Sub TXTCODINT_Change()
   CI1% = XVALO(TXTCODINT)
   
End Sub
  Sub CARGARMSF1()
   CI1% = XVALO(TXTCODINT)
   Dim CTP As New Conteplast
   SELPRINOT.Text11 = CODEXT(CI1%)
   SELPRINOT.Text4 = DESCRIT0$(CI1%)
   NROFT& = CTP.NroFTecnicaXCodigo(CI1%)
   TXTFTECNICA = NROFT&
   NROCOTIZAX& = XVALO(TXTNROCOTIZA)
   Call CTP.CARGARARCHIVOSENCOMBO(XVALO(TXTFTECNICA), Combo2)

   Call CARGA_ENCABEZADO(MSF1, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;CI/f6;CIORIG/A8", Me)
   Reg& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1% & " ORDER BY NUORIT"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
       Reg& = Reg& + 1
         CIELEM% = XVALO(!codielem)
         MSF1.Rows = Reg& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         MSF1.TextMatrix(Reg&, 1) = NOMBRESECTOR$
         MSF1.TextMatrix(Reg&, 2) = SAFETEXT$(!CodXElem)
         MSF1.TextMatrix(Reg&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         MSF1.TextMatrix(Reg&, 4) = CTP.FRMSECTOR(NOMBRESECTOR$)
         MSF1.TextMatrix(Reg&, 5) = CIELEM%
         CIRANGO = ""
         If ESCODIGORANGO(CIELEM%) > 0 Then
            'en esta fila guarda el codigo interno cargado en la cotizacion
            MSF1.TextMatrix(Reg&, 6) = CTP.CodigoOriginalCoti1SegunComponente(NROCOTIZAX&, NOMBRESECTOR$)
         End If
       .MoveNext
     Loop
   End With
   Set CTP = Nothing
End Sub

Private Sub TXTFTECNICA_Change()
    Call CARGA_ENCABEZADO(MSF2, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;CI/f6", Me)
    Dim CTP As New Conteplast
    Call CARGARMSF1
    
    CONDI$ = "NUMRECETACTP  = " & XVALO(TXTFTECNICA)
    IMPRE$ = TRIM$(VALOBADA("RECETACTP", "IMPRESION", CONDI$, "NOMESS"))
    If IMPRE$ = "SI" Then
        Reg& = 0
        Reg& = Reg& + 1
        MSF2.Rows = Reg& + 1
        
        MSF2.TextMatrix(Reg&, 1) = "IMPRESIÓN"
        MSF2.TextMatrix(Reg&, 2) = ""
        MSF2.TextMatrix(Reg&, 3) = ""
        MSF2.TextMatrix(Reg&, 4) = CTP.FRMSECTOR("IMPRESION")
    End If
End Sub
Sub PRINTETIQUETACTP_DESDEOF(CI1%, NORFA$, CODIGOEDITABLE$)
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
    RUTAEXCEL$ = Control("COTICONT", "EXCEETIQ")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    IMAG$ = "QR_ETIQPTERM.JPG"

    On Error Resume Next
    Kill "C:\FLEXOFT\" & IMAG$
    On Error GoTo 0
'1 razon social
'2 orden de compra
'3 codigo p.t.
'4 medidas
'5 descripcion de p.t.
'6 orden de fabricacion
    Dim CTP As New Conteplast
    NC = cliente_SEGUN_PEDIDO(XVALO(Me.TXTNROPED))
    NP& = XVALO(TXTNROPED)
    CI1% = XVALO(Me.TXTCODINT)
    RASO$ = rasocli$(NC)
    OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    CODI$ = CODIGOEDITABLE$
    DESKRI$ = TRIM$(Text4)
    NF& = XVALO(TXTFTECNICA)
    ANCHOX = CTP.ANCHOCTP(CI1%)
    LARGOX = CTP.LARGOCTP(CI1%)
    ALTOX = CTP.ALTOCTP(CI1%)
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    FEFA% = FECHGENORFA%(NORFA$)
    '
    With VENTABNEW.MSF
       .Rows = 10: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = OC$
       .TextMatrix(3, 1) = CODI$
       .TextMatrix(4, 1) = CTP.CodEspTecCliente$(CI1%)
       .TextMatrix(5, 1) = DESKRI$
       .TextMatrix(6, 1) = NORFA$
       .TextMatrix(7, 1) = FECHATEX$(FEFA%)
       .TextMatrix(8, 1) = ""
       .TextMatrix(9, 1) = ""
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "ETIQOF", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0)                  'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
End Sub


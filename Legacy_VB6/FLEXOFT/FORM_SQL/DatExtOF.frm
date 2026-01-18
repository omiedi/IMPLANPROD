VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DatExtOF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Impresión de Formularios de Orden de Fabricación de Bolsas"
   ClientHeight    =   6075
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   10155
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Command4 
         Caption         =   "."
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
         Left            =   3945
         TabIndex        =   44
         Top             =   2280
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Grabar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   39
         Top             =   2280
         Width           =   1330
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Observaciones Adicionales"
         Top             =   2760
         Width           =   3050
      End
      Begin VB.TextBox Text6 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Borrar"
         Top             =   2280
         Width           =   1740
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Cantidad de Hilos:"
         Top             =   2280
         Width           =   1905
      End
      Begin VB.CommandButton Command44 
         Caption         =   "."
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
         Left            =   3945
         TabIndex        =   34
         Top             =   1800
         Width           =   175
      End
      Begin VB.TextBox TXTCODIGO 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Borrar"
         Top             =   1800
         Width           =   1740
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Index           =   29
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Código de Tela:"
         Top             =   1800
         Width           =   1905
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   3120
         Width           =   9780
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H009FBE85&
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
         TabIndex        =   20
         Top             =   840
         Width           =   9915
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   150
            Width           =   1140
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
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   150
            Width           =   2235
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
            TabIndex        =   28
            Top             =   420
            Width           =   2235
         End
         Begin VB.TextBox Text10 
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   4100
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
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   150
            Width           =   4100
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
            Left            =   7565
            Locked          =   -1  'True
            TabIndex        =   24
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
            Left            =   7565
            Locked          =   -1  'True
            TabIndex        =   23
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
            Left            =   8705
            Locked          =   -1  'True
            TabIndex        =   22
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
            Left            =   8705
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H009FBE85&
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   9915
         Begin VB.TextBox Text8 
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
            Left            =   6650
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   420
            Width           =   3200
         End
         Begin VB.TextBox Text7 
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
            Height          =   350
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   420
            Width           =   1500
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
            Left            =   6650
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "Material"
            Top             =   180
            Width           =   3200
         End
         Begin VB.TextBox Text14 
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
            Height          =   405
            Index           =   7
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "Largo  Corte cm "
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Index           =   8
            Left            =   4665
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "N°.O."
            Top             =   180
            Width           =   500
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Index           =   6
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "F.Técnica"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTFTECNICA 
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
            Height          =   350
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Index           =   5
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Cotización"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTNROCOTIZA 
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
            Height          =   350
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Cód.Interno"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Index           =   3
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Pedido"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTNROPED 
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
            Height          =   350
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox TXTCODINT 
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
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   420
            Width           =   1155
         End
         Begin VB.TextBox TXTNORPED 
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
            Height          =   350
            Left            =   4665
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   420
            Width           =   500
         End
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir Orden Fabricación"
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
         Left            =   6360
         TabIndex        =   8
         Top             =   4680
         Width           =   3690
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H009FBE85&
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
         TabIndex        =   5
         Top             =   5160
         Width           =   9915
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
            TabIndex        =   7
            Top             =   300
            Width           =   850
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
            TabIndex        =   6
            Top             =   300
            Width           =   6045
         End
      End
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
         Left            =   240
         TabIndex        =   4
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atención !!! El codigo de Tela debe ser un Código Rango"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   4200
         TabIndex        =   45
         Top             =   2280
         Width           =   4095
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      OleObjectBlob   =   "DatExtOF.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DatExtOF.frx":0234
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
      Caption         =   "Transacciones"
      Begin VB.Menu mnuINGTELBOL 
         Caption         =   "Ingreso de Telas a Dep. Corte P/Bolsas"
      End
   End
   Begin VB.Menu mnuCtalis 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
   End
   Begin VB.Menu mnuacces 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "DatExtOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ActualizaListaSeleccion As Boolean




Private Sub Command1_Click()
   NRO = XVALO(TXTFTECNICA)
   If NRO < 1 Then Exit Sub
   COTICONT.Text5 = NRO
   COTICONT.mnuimprimirfichaactiva_Click
   Unload COTICONT
End Sub





Sub LIMPIAR()

    Text11 = ""
    Text4 = ""
    Text10 = ""
    MSF.Rows = 1
    '
    Text6 = ""

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   CICONJ% = XVALO(TXTCODINT)
   If CICONJ% < 1 Then GoTo 99
   
   NORFA$ = TRIM$(Text10)
   If NORFA$ = "" Then
     Call COMUNI("Falta Seleccionar Orden de Fabricación")
     GoTo 99
   End If
   
   CF$ = TRIM$(UCase$(CODFAMIL$(CICONJ%))) 'CODIGO DE FAMILIA
   'SI ES CORTES
   If CF$ = "CT" Or CF$ = "SA" Then
      'Call GrabarDatosCorte(NORFA$, CICONJ$) ' POR AHORA NO GUARDO NI TROQUEL NI OBSERVACION
      Call ANOTAORF_NEW(NORFA$, "TROQUEL: " & Text5) 'LO DEJO EN EL ANOTADOR
      Call ANOTAORF_NEW(NORFA$, "Observaciones: " & Text28)
      
      GoTo 99 'ACA TERMINA
   End If

   
   CI1% = CODINT%(Me.TXTCODIGO)
   If CI1% < 1 Then
     Call COMUNI("Falta Seleccionar Código de Tela")
     GoTo 99
   End If
   '
   LARGOCORTE = XVALO(Text7)
   If XVALO(Text7) <= 0 Then
     Call COMUNI("Falta Largo de Corte (Verifique la Ficha Técnica)")
     GoTo 99
   End If
   '
   If XVALO(TXTFTECNICA) <= 0 Then
     Call COMUNI("Falta N°: de Ficha Técnica)")
     GoTo 99
   End If
   '
   ValorAdicional = 4
   Material$ = TRIM$(UCase$(Text8))
   If Material$ = "BOLSA POLIPAPEL" Then
      ValorAdicional = 2
   End If
   LARGOCORTE = LARGOCORTE + ValorAdicional
   '
   
   'BORRO LA FORMULA DE ESTE CODIGO
   CONDI$ = "CODICONJ = " & CICONJ%
   CONDI$ = CONDI$ + " AND codiempr = " & CODIEMPR%
   SQLX$ = "DELETE FORMULAS WHERE " & CONDI$
   FLEXCONN.Execute (SQLX$)
  
   '
   SQLX$ = "SELECT * FROM Formulas WHERE CodIConj < 1"
   Dim rs1 As ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With rs1

      .AddNew
      !CODICONJ = CICONJ%
      !CODIEMPR = CodiEmp%
      !CODXCONJ = Left(CODEXT$(CICONJ%), 16)
      !codielem = CI1%
      !CodXElem = Left(CODEXT$(CI1%), 16)
       NUORIT% = 1
      !NUORIT = NUORIT%
      Uso1# = XVALO(LARGOCORTE)
      UNXX$ = Unimed$(CI1%)
      If UNXX$ = "M." Then
        Uso1# = Uso1# / 100  'LO PASO A MTS
      End If
'      USOX$ = FORMATNUM$(XVALO(Uso1#), "F12.4")

      !UsoBruto = Uso1#
      !Unis_Form = Uso1#
      !UMED_FORM = UNXX$
      !LAR_CORTE = 0
      !Pladispo = 0
      !RefeComp = ""
      !MARBORRA = 0
      !COMPONENTE = "FABRICACIÓN DE BOLSA"
      !moduxbb = 1
      !corte = Uso1#
      .Update
   End With
   rs1.Close
   Set rs1 = Nothing

   SQLX$ = "UPDATE ORDEFABR SET CantHilos = " & XVALO(Text5)
   SQLX$ = SQLX$ + ", OBSERVA = '" & Left$(Text28, 1024) & "'"
   SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
   FLEXCONN.Execute (SQLX$)
   
   Call REFRESCA_RESERVA(CICONJ%, XVALO(Text3), NORFA$) 'GENERA LA RESERVA P.Q. AL GENERAR LA OF NO ESTABA EL DATO DEL MATERIAL
   

     
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()

    CICONJ% = XVALO(TXTCODINT)
    If CICONJ% < 1 Then GoTo 99
   
    CF$ = TRIM$(UCase$(CODFAMIL$(CICONJ%))) 'CODIGO DE FAMILIA
    'SI ES CORTES
    If CF$ = "CT" Or CF$ = "SA" Then
      Call PRINTOFCORTES(TRIM$(Text10), CICONJ%)
      Exit Sub
    End If
    
    If XVALO(TXTNROCOTIZA.TEXT) < 1 And XVALO(TXTFTECNICA.TEXT) < 1 Then
     'IMPRESION PARA BOLSAS ENVAPLAST QUE NO TIENEN FICHA TENCNCIA NI COTIZACION
     If TRIM$(UCase$(CODFAMIL$(CICONJ%))) = "BO" Then
        FILTX$ = TRIM$(Text10) & ";" & TRIM$(str$(CICONJ%))
        Call STDFORM("ODFA-EYM", FILTX$, "PRINT")
        Exit Sub
     End If
   End If
   
    CI1% = CODINT%(Text11)
    If CI1% < 1 Then
     Call COMUNI("Falta Seleccionar Código de Tela")
     GoTo 99
    End If
    '
    '

   
    If XVALO(TXTFTECNICA) <= 0 Then
     Call COMUNI("Falta N°: de Ficha Técnica)")
     GoTo 99
    End If
    '
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")

    '
    
    RUTAEXCEL$ = Control("PLANVENT", "IMPOFBOL")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    Dim ctp As New Conteplast
    NPED& = XVALO(TXTNROPED)
    ORDENPEDIDO% = XVALO(TXTNORPED)
    NC = cliente_SEGUN_PEDIDO(NPED&)
    RASO$ = rasocli$(NC)
    NROFT& = XVALO(TXTFTECNICA)
    FechaEntrega% = CVINT(VALOBADA("PEDDETA", "FESOLENT", "NROPED=" & NPED& & " AND NUORITEM =" & ORDENPEDIDO%, "NOMESS"))
    FechaEmisi% = CVINT(VALOBADA("pedenca", "Fechemis", "NROPED=" & NPED&, "NOMESS"))
    '

'    Text10
'    Text11
'    NROFT&
'    RASO$
'    FechaEntrega%
'    Text8
    Call CREACAMPO("", "P_VTA_BOLSAS", "NROPVTA", 4, 0, 1)
    CONDI$ = "NROPVTA = " & NROFT&
    PRODUCTO = VALOBADA("P_VTA_BOLSAS", "PRODUCTO", CONDI$, "NOMESS")
    ANCTOTO = VALOBADA("P_VTA_BOLSAS", "ANCHOTOTCM", CONDI$, "NOMESS")
    ORCOMPRA$ = VALOBADA("PEDENCA", "NROOCOM", "NROPED= " & NPED&, "NOMESS")
       NVEND% = XVALO(VALOBADA("PEDENCA", "Vended", "NROPED= " & NPED&, "NOMESS"))
    NOMBVENDEDOR$ = ECOARCH$("VENDEDOR", Chr$(NVEND%))
    IMPRE$ = VALOBADA("P_VTA_BOLSAS", "IMPRESION", CONDI$, "NOMESS")
    CANTCOLORES% = XVALO(VALOBADA("P_VTA_BOLSAS", "CANTCOLORES", CONDI$, "NOMESS"))
    FUELLE$ = VALOBADA("P_VTA_BOLSAS", "FUELLE", CONDI$, "NOMESS")
    SisCarga$ = VALOBADA("P_VTA_BOLSAS", "SisCarga", CONDI$, "NOMESS")
    SISDESCA$ = VALOBADA("P_VTA_BOLSAS", "SISDESCA", CONDI$, "NOMESS")
    MICROPERFO$ = VALOBADA("P_VTA_BOLSAS", "MICROPERFO", CONDI$, "NOMESS")
    'CANTHILOS% = XVALO(VALOBADA("P_VTA_BOLSAS", "CantHilos", CONDI$, "NOMESS")) 'LO TOMO DE PANTALLA
    CIELEM% = VALOBADA("FORMULAS", "CODIELEM", "CODICONJ = " & CI1%, "NOMESS")
    CODIGOTELA$ = VALOBADA("MASTER", "DESCRIPCION", "CODINT = " & CIELEM%, "NOMESS")
    UNIDPAQU% = Int(XVALO(VALOBADA("P_VTA_BOLSAS", "UNIDPAQU", CONDI$, "NOMESS")))
    Observaciones$ = VALOBADA("P_VTA_BOLSAS", "Observaciones", CONDI$, "NOMESS")
    CantProyectada = FORMATNUM$(XVALO(Text3), "f10.0")
    CAPKGS = TRIM$(FORMATNUM$(XVALO(VALOBADA("P_VTA_BOLSAS", "CAPKGS", CONDI$, "NOMESS")), "f12.2"))
    AltoUtil = XVALO(ctp.AltoUtilFTecnicaBolsa(NROFT&))
    
    CANTICARAS% = XVALO(VALOBADA("P_VTA_BOLSAS", "CANTICARAS", CONDI$, "NOMESS"))
    MEDIFUELLE = XVALO(VALOBADA("P_VTA_BOLSAS", "MEDIFUELLE", CONDI$, "NOMESS"))  'Medida fuelle (de FT)  x  2
    LINER$ = VALOBADA("P_VTA_BOLSAS", "LINER", CONDI$, "NOMESS")
    DescarCOLOR$ = VALOBADA("P_VTA_BOLSAS", "DescarCOLOR", CONDI$, "NOMESS")
    AmchoCadeneta$ = "6.5 mm."
    TIPOHILO$ = VALOBADA("P_VTA_BOLSAS", "TIPOHILO", CONDI$, "NOMESS")
    LCORTE = Text7 ' YA ESTA CON LA FORMULA DEL EXCEL INCLUIDA
    UNIPALLET = XVALO(VALOBADA("P_VTA_BOLSAS", "UNIPALLET", CONDI$, "NOMESS"))
    
    With VENTABNEW.MSF
    .Rows = 36: .Cols = 2
    .TextMatrix(1, 1) = Text10 'Orden de fabricación 1
    .TextMatrix(2, 1) = Text11 'Código bolsa 1
    .TextMatrix(3, 1) = NROFT& 'F.Técnica Nº1
    .TextMatrix(4, 1) = RASO$ 'Razón Social 1
    .TextMatrix(5, 1) = FECHAFORMATOLARGO$(FechaEntrega%)  'Fecha de entrega1
    .TextMatrix(6, 1) = Text8 'Material (bolsa)1
    .TextMatrix(7, 1) = PRODUCTO 'Producto 1
    .TextMatrix(8, 1) = ANCTOTO 'ANCHO Total(cm) 1
    .TextMatrix(9, 1) = NPED& 'Nota de Pedido 1
    .TextMatrix(10, 1) = ORCOMPRA$ 'Orden de Compra Nº 1
    .TextMatrix(11, 1) = NOMBVENDEDOR$ 'Vendedor 1
    .TextMatrix(12, 1) = IMPRE$ 'Impresión1
    .TextMatrix(13, 1) = CANTCOLORES% 'CANT.Colores 1
    .TextMatrix(14, 1) = FUELLE$ 'Fuelle1
    .TextMatrix(15, 1) = SisCarga$ 'Sistema de Carga1
    .TextMatrix(16, 1) = SISDESCA$ 'Sistema de Descarga1
    .TextMatrix(17, 1) = MICROPERFO$ 'Microperforado1
    .TextMatrix(18, 1) = XVALO(Text5) 'Cantidad de Hilos1
    .TextMatrix(19, 1) = CODIGOTELA$ 'Codigo de Tela 1
    .TextMatrix(20, 1) = UNIDPAQU% 'Un por paquete'1
    .TextMatrix(21, 1) = Observaciones$ 'Observaciones Generales1
    .TextMatrix(22, 1) = CantProyectada 'Cantidad1
    .TextMatrix(23, 1) = CAPKGS 'Capacidad1
    .TextMatrix(24, 1) = AltoUtil 'ALTO útil(cm) 1
    .TextMatrix(25, 1) = FECHATEX$(FechaEmisi%)  'FECHA NP1
    .TextMatrix(26, 1) = CANTICARAS% 'CANT.Caras1
    .TextMatrix(27, 1) = MEDIFUELLE * 2 'Medida fuelle (de FT)  x  2
    .TextMatrix(28, 1) = LINER$ 'Liner1
    .TextMatrix(29, 1) = DescarCOLOR$ 'COLOR1
    .TextMatrix(30, 1) = AmchoCadeneta$ 'Ancho de Cadeneta1
    .TextMatrix(31, 1) = TIPOHILO$ 'Tipo de Hilo1
    .TextMatrix(32, 1) = LCORTE ' LCORTE1 REALIZAR FORMULA
    .TextMatrix(33, 1) = UNIPALLET 'Paquetes por pallet1
    .TextMatrix(34, 1) = Text28 'Observaciones Adicionales1

    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "PARA BOLSAS", VENTABNEW.MSF, PDES$, "1", 1, 10, "", 1, 0)                       'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    Set ctp = Nothing
99  Command3.Enabled = True
End Sub
Sub PRINTOFCORTES(NORFA$, CI1%)
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")
    '
    RUTAEXCEL$ = Control("ORDEFABR", "IMPOFCOR")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    Dim ctp As New Conteplast
    NPED& = XVALO(TXTNROPED)
    ORDENPEDIDO% = XVALO(TXTNORPED)
    NC = cliente_SEGUN_PEDIDO(NPED&)
    RASO$ = rasocli$(NC)
    If NPED& > 0 Then
       FechaEntrega% = CVINT(VALOBADA("PEDDETA", "FESOLENT", "NROPED=" & NPED& & " AND NUORITEM =" & ORDENPEDIDO%, "NOMESS"))
    Else
       'SI NO TIENE NRO DE PEDIDO BUSCA LA FECHA DE ENTREGA DE ORDEFABR CUANDO SE GENERA LA MISMA
       FechaEntrega% = CVINT(VALOBADA("ORDEFABR", "FechEnSol", "IDSUBOR='" & NORFA$ & "'", "NOMESS"))
    End If
    FechaEmisi% = CVINT(VALOBADA("pedenca", "Fechemis", "NROPED=" & NPED&, "NOMESS"))
    Codigo$ = Text11
    CIELEM% = XVALO(VALOBADA("FORMULAS", "CODIELEM", "CODICONJ = " & CI1%, "NOMESS"))
    LARGOCORTE$ = FORMATNUM$(XVALO(VALOBADA("FORMULAS", "usobruto", "CODICONJ = " & CI1%, "NOMESS")), "F10.2") 'USO
    CODIGORANGOX$ = VALOBADA("FORMULAS", "codxelem", "CODICONJ = " & CI1%, "NOMESS")
    DESCRIRANGO$ = DESCRIT0$(CIELEM%)
    ORCOMPRA$ = VALOBADA("PEDENCA", "NROOCOM", "NROPED= " & NPED&, "NOMESS")
       NVEND% = XVALO(VALOBADA("PEDENCA", "Vended", "NROPED= " & NPED&, "NOMESS"))
    NOMBVENDEDOR$ = ECOARCH$("VENDEDOR", Chr$(NVEND%))
    Observaciones$ = VALOBADA("PEDENCA", "Observa", "NROPED=" & NPED&, "NOMESS")
    With VENTABNEW.MSF
    .Rows = 18: .Cols = 2
    .TextMatrix(1, 1) = Text10 'Orden de fabricación 1
    .TextMatrix(2, 1) = NPED& 'Código bolsa 1
    .TextMatrix(3, 1) = FECHATEX$(FechaEmisi%)
    .TextMatrix(4, 1) = RASO$ 'Razón Social 1
    .TextMatrix(5, 1) = FECHAFORMATOLARGO$(FechaEntrega%)  'Fecha de entrega1
    .TextMatrix(6, 1) = Codigo$
    .TextMatrix(7, 1) = DESCRIRANGO$
    .TextMatrix(8, 1) = Text5 'TROQUEL
    .TextMatrix(9, 1) = ORCOMPRA$
    .TextMatrix(10, 1) = NOMBVENDEDOR$ '
    .TextMatrix(11, 1) = Text6 'descripcion TELA
    .TextMatrix(12, 1) = LARGOCORTE$
    .TextMatrix(13, 1) = Observaciones$ 'DEL PEDIDO
    .TextMatrix(14, 1) = Text28 'DE PANTALLA
    .TextMatrix(15, 1) = CODIGORANGOX$
    .TextMatrix(16, 1) = XVALO(Text3)
   

    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "Hoja1", VENTABNEW.MSF, PDES$, "1", 1, 9, "", 1, 0)                         'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
    Set ctp = Nothing
99  Command3.Enabled = True
End Sub


Private Sub Command4_Click()
   X = LISTASELECGENERAL("DIAMECOR", , , 1)
   Me.Text5 = TRIM$(RESP_ZELE_VECT(2))
End Sub

Private Sub Command44_Click()
     COD$ = ListaSeleccionArticulos("", ActualizaListaSeleccion)
     ActualizaListaSeleccion = False
     TXTCODIGO = COD$
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

Private Sub Form_Load()
   Call CREARTABLASOLICOTI
   Call CREATABLA_P_VTA_BOLSAS
End Sub

Private Sub mnuINGTELBOL_Click()

    Call COMUNI("Realizar el Ingreso de Telas\Desde la Pantalla de Ingeso de Bobinas a Corte\Del Modo Habitual")
'    INGTELBOL.TXTOF = Text10
'    INGTELBOL.Show 1
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
   CANTHILOS = XVALO(VALOBADA("ORDEFABR", "CantHilos", CONDI$, "NOMESS"))
   OBSERVACIONESADI$ = VALOBADA("ORDEFABR", "Observa", CONDI$, "NOMESS")
   Text5 = FORMATNUM$(CANTHILOS, "F10.O")
   Text28 = OBSERVACIONESADI$

End Sub

Private Sub Text5_LostFocus()
  Text5 = FORMATNUM$(XVALO(Text5), "F10.0")
End Sub

Private Sub txtCodigo_Change()
   CI1% = CODINT%(TXTCODIGO)
   If CI1% > 0 Then
      Text6 = DESCRIT0$(CI1%)
   Else
      Text6 = ""
   End If
End Sub

Private Sub TXTCODINT_Change()
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   NROFT& = ctp.NroFTecnicaXCodigoBolsa(CI1%)
   TXTFTECNICA = NROFT&
   Text11 = CODEXT(CI1%)
   Text4 = DESCRIT0$(CI1%)
   CONDI$ = "CODICONJ = " & CI1%
   TXTCODIGO = VALOBADA("FORMULAS", "CODXELEM", CONDI$, "NOMESS")
   
   CF$ = TRIM$(UCase$(CODFAMIL$(CI1%))) 'CODIGO DE FAMILIA
   '
   
   'SI ES CORTES NO CARGA MAS
   If CF$ = "CT" Or CF$ = "SA" Then
      Exit Sub
   End If
   
   AltoUtil = XVALO(ctp.AltoUtilFTecnicaBolsa(NROFT&))
   MatBolsa$ = TRIM$(UCase$(ctp.MatBolsaFTecnicaBolsa(NROFT&)))
   Text8 = MatBolsa$
   'YA PRESENTA EL LARGO DE CORTE CALCULADO
   ValorAdicional = 4
   Material$ = TRIM$(UCase$(Text8))
   If Material$ = "BOLSA POLIPAPEL" Then
      ValorAdicional = 2
   End If
   
   LARGOCORTE = AltoUtil + ValorAdicional
   Text7 = FORMATNUM$(LARGOCORTE, "F12.2")
   
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
    Dim ctp As New Conteplast
    NC = cliente_SEGUN_PEDIDO(XVALO(Me.TXTNROPED))
    NP& = XVALO(TXTNROPED)
    CI1% = XVALO(Me.TXTCODINT)
    RASO$ = rasocli$(NC)
    OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    CODI$ = CODIGOEDITABLE$
    DESKRI$ = TRIM$(Text4)
    NF& = XVALO(TXTFTECNICA)
    ANCHOX = ctp.ANCHOCTP(CI1%)
    LARGOX = ctp.LARGOCTP(CI1%)
    ALTOX = ctp.ALTOCTP(CI1%)
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    '
    With VENTABNEW.MSF
       .Rows = 10: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = OC$
       .TextMatrix(3, 1) = CODI$
       .TextMatrix(4, 1) = ctp.CodEspTecCliente$(CI1%)
       .TextMatrix(5, 1) = DESKRI$
       .TextMatrix(6, 1) = NORFA$
       .TextMatrix(7, 1) = ""
       .TextMatrix(8, 1) = ""
       .TextMatrix(9, 1) = ""
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "ETIQOF", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0)                  'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
End Sub


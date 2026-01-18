VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FRMPRESUAHP 
   BackColor       =   &H00FFFFFF&
   Caption         =   "COTIZACIÓN AHP"
   ClientHeight    =   9165
   ClientLeft      =   0
   ClientTop       =   -555
   ClientWidth     =   14520
   DrawMode        =   11  'Not Xor Pen
   ForeColor       =   &H00000000&
   Icon            =   "FRMPRESUAHP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14520
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCONOTA 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   0
      ScaleHeight     =   3510
      ScaleWidth      =   14505
      TabIndex        =   106
      Top             =   2880
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton Command15 
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
         Height          =   570
         Left            =   13800
         Picture         =   "FRMPRESUAHP.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Cancela Pedido de Cliente"
         Top             =   1080
         Width           =   570
      End
      Begin VB.CommandButton Command14 
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
         Height          =   570
         Left            =   13785
         Picture         =   "FRMPRESUAHP.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Continúa Generación del Pedido de Cliente"
         Top             =   420
         Width           =   570
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F9FD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   109
         Top             =   370
         Width           =   13725
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   -30
         ScaleHeight     =   405
         ScaleWidth      =   14535
         TabIndex        =   107
         Top             =   -30
         Width           =   14565
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ANOTACIONES DE CLIENTES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   0
            TabIndex        =   108
            Top             =   105
            Width           =   14460
         End
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8520
      TabIndex        =   104
      Top             =   7440
      Width           =   2955
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Generar Solicitud de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11700
      TabIndex        =   101
      Top             =   2600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   4440
      TabIndex        =   100
      Top             =   1852
      Width           =   4500
      Begin VB.Image Image1 
         Height          =   800
         Left            =   0
         Picture         =   "FRMPRESUAHP.frx":0620
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4500
      End
   End
   Begin VB.TextBox TXTMOVIL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   315
      Left            =   8040
      TabIndex        =   97
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      _Version        =   393216
      BackColorFixed  =   16501405
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   6720
      Top             =   4320
   End
   Begin VB.TextBox LENTREGA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   0
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   8280
      Width           =   7920
   End
   Begin VB.TextBox LENTREGA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1335
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   7920
      Width           =   6585
   End
   Begin VB.TextBox Text19 
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
      Height          =   315
      HelpContextID   =   15
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Text            =   "Entrega"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text41 
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
      Height          =   315
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Text            =   "Vendedor"
      Top             =   7575
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1935
      TabIndex        =   91
      Top             =   7575
      Width           =   255
   End
   Begin VB.TextBox TXTNUMVENDEDOR 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1335
      TabIndex        =   90
      Top             =   7575
      Width           =   585
   End
   Begin VB.TextBox TXTNOMVENDEDOR 
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
      Height          =   315
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   7575
      Width           =   5730
   End
   Begin VB.TextBox TXT36 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   87
      Text            =   "Total Kg.:"
      Top             =   8250
      Width           =   1275
   End
   Begin VB.TextBox TXTTOTALPESO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   8250
      Width           =   1395
   End
   Begin VB.TextBox TXTCODCLIENTE 
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
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   84
      ToolTipText     =   "Código de Cliente - Doble Click Para Eliminar Relación "
      Top             =   2580
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      ToolTipText     =   "Seleccion de Artículo (F2)"
      Top             =   1850
      Width           =   570
   End
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sel.Manual"
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
      Height          =   255
      Index           =   2
      Left            =   13080
      TabIndex        =   83
      ToolTipText     =   " Modo de Carga: Selección Manual de Lotes"
      Top             =   2600
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Verif. Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   11520
      TabIndex        =   82
      ToolTipText     =   "Modo de Carga: Valida Stock de Maera General"
      Top             =   2600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Carga Rapida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   81
      ToolTipText     =   "Modo de Carga: Directamente sin Validar Stock "
      Top             =   2600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12840
      TabIndex        =   78
      Top             =   1850
      Width           =   1620
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Graba e Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   77
      Top             =   7800
      Width           =   2955
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1695
      Left            =   15
      TabIndex        =   46
      Top             =   0
      Width           =   14520
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   13320
         ScaleHeight     =   165
         ScaleWidth      =   1065
         TabIndex        =   112
         Top             =   1440
         Width           =   1100
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   113
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7320
         TabIndex        =   102
         Top             =   1090
         Width           =   255
      End
      Begin VB.TextBox TXTREFERENCIA 
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
         Height          =   285
         Left            =   8895
         TabIndex        =   2
         Top             =   1080
         Width           =   3200
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Height          =   1120
         Left            =   120
         TabIndex        =   56
         Top             =   210
         Width           =   1400
         Begin VB.CommandButton Command19 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1100
            TabIndex        =   57
            ToolTipText     =   "Seleccion de Cliente"
            Top             =   480
            Width           =   250
         End
         Begin VB.TextBox Text1 
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
            Height          =   300
            Left            =   120
            TabIndex        =   0
            Top             =   460
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   240
            Left            =   100
            OleObjectBlob   =   "FRMPRESUAHP.frx":1B0662
            TabIndex        =   58
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label26 
            BackColor       =   &H00EEEEEE&
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
            Height          =   300
            Left            =   100
            TabIndex        =   59
            Top             =   765
            Width           =   1215
         End
      End
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   220
         Width           =   4950
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   505
         Width           =   4950
      End
      Begin VB.TextBox Text4 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   795
         Width           =   4950
      End
      Begin VB.TextBox Text18 
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
         Left            =   8895
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   790
         Width           =   5520
      End
      Begin VB.TextBox Text5 
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
         Left            =   8895
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   505
         Width           =   5520
      End
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
         Height          =   285
         Left            =   8895
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   220
         Width           =   2895
      End
      Begin VB.TextBox Text17 
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
         Left            =   13290
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1100
         Width           =   1125
      End
      Begin VB.TextBox Text20 
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
         Left            =   12300
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox Text35 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   1080
         Width           =   4670
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   47
         ToolTipText     =   "Seleccion de Cliente"
         Top             =   1400
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   240
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B06CE
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B073A
         TabIndex        =   61
         Top             =   525
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B07AC
         TabIndex        =   62
         Top             =   815
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   240
         Left            =   7815
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B081E
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   7815
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B088E
         TabIndex        =   64
         Top             =   525
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   7815
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B08FC
         TabIndex        =   65
         Top             =   810
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   240
         Left            =   11865
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B0964
         TabIndex        =   66
         Top             =   240
         Width           =   1065
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   285
         Left            =   12120
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B09D0
         TabIndex        =   67
         Top             =   1100
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   13020
         TabIndex        =   68
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   134086657
         CurrentDate     =   40810
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   285
         Left            =   11900
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B0A48
         TabIndex        =   69
         Top             =   1120
         Visible         =   0   'False
         Width           =   360
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B0AAE
         TabIndex        =   70
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   7815
         OleObjectBlob   =   "FRMPRESUAHP.frx":1B0B1E
         TabIndex        =   88
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   33
      Top             =   1560
      Width           =   14520
      Begin VB.TextBox TXTCANTIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6300
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "   "
         ToolTipText     =   "Enter para Continuar"
         Top             =   540
         Width           =   930
      End
      Begin VB.CommandButton Command12 
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
         Height          =   540
         Left            =   7320
         TabIndex        =   103
         ToolTipText     =   "Agregar Observación en Detalle"
         Top             =   320
         Width           =   375
      End
      Begin VB.TextBox TXTNUMPROV 
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
         Height          =   285
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   600
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TXTPESOP 
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
         Height          =   285
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text38 
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
         Height          =   380
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   540
         Visible         =   0   'False
         Width           =   1300
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Vía de Entr."
         Top             =   270
         Visible         =   0   'False
         Width           =   1545
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   540
         Width           =   3855
      End
      Begin VB.TextBox Text9 
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
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "P.Venta"
         Top             =   270
         Visible         =   0   'False
         Width           =   1350
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
         Height          =   285
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   270
         Width           =   930
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
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   3855
      End
      Begin VB.TextBox Text12 
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
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtcodigo 
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
         Height          =   380
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Código Original"
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox TXTPREUNIT 
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
         Height          =   380
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox Text6 
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
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Importe"
         Top             =   270
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox TXTIMPORTE 
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
         Height          =   380
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
         Top             =   540
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox Text15 
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
         Height          =   380
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Costo de Maestro de Articulos"
         Top             =   540
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox Text16 
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   270
         Visible         =   0   'False
         Width           =   1470
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
         Left            =   7950
         Sorted          =   -1  'True
         TabIndex        =   96
         Text            =   "Combo6"
         Top             =   540
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1935
      TabIndex        =   32
      Top             =   6540
      Width           =   255
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1330
      TabIndex        =   31
      Top             =   6555
      Width           =   585
   End
   Begin VB.TextBox Text22 
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
      Height          =   315
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6540
      Width           =   5730
   End
   Begin VB.TextBox Text39 
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
      Height          =   315
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "L. Precios"
      Top             =   6540
      Width           =   1215
   End
   Begin VB.TextBox Text10 
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
      Left            =   11775
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Total:"
      Top             =   7935
      Width           =   975
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
      Left            =   12855
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7935
      Width           =   1635
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   360
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "MODIFICACIÓN DE PRESUPUESTO"
      ToolTipText     =   "Doble Click Para Salir de Modo Edición"
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   5295
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox TICAMBIS1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   10080
      MaxLength       =   14
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Text24 
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
      Height          =   380
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "[F2] Presenta Lista de Códigos -  [F4] Edición de Precio de Venta -  [F5] Graba e Imprime Presupuesto Activo."
      Top             =   8720
      Width           =   14475
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
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
      Left            =   12855
      TabIndex        =   21
      Text            =   " "
      Top             =   6600
      Width           =   1620
   End
   Begin VB.TextBox PORDESCU 
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
      Left            =   11415
      MaxLength       =   14
      TabIndex        =   20
      Top             =   7080
      Width           =   1065
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
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
      Left            =   12855
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   " "
      Top             =   6960
      Width           =   1620
   End
   Begin VB.TextBox Text25 
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
      Left            =   11655
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox TXTDESCRIPAGO 
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
      Height          =   315
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6885
      Width           =   5730
   End
   Begin VB.TextBox TXTCONDIPAGO 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1335
      TabIndex        =   16
      Top             =   6885
      Width           =   585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1935
      TabIndex        =   15
      Top             =   6885
      Width           =   255
   End
   Begin VB.TextBox Text27 
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
      Height          =   315
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Cond. Pago"
      Top             =   6885
      Width           =   1215
   End
   Begin VB.TextBox Text28 
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
      Height          =   315
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7230
      Width           =   5730
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1335
      TabIndex        =   12
      Top             =   7230
      Width           =   585
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1935
      TabIndex        =   11
      Top             =   7230
      Width           =   255
   End
   Begin VB.TextBox Text30 
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
      Height          =   315
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Val.Cotización"
      Top             =   7230
      Width           =   1215
   End
   Begin VB.TextBox Text31 
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
      Left            =   12855
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1635
   End
   Begin VB.TextBox Text32 
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
      Left            =   11775
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "I.V.A.:"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox Text33 
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
      Left            =   12855
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1635
   End
   Begin VB.TextBox Text34 
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
      Left            =   11775
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8565
      TabIndex        =   5
      ToolTipText     =   "Seleccion de Cliente"
      Top             =   1810
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   15
      OleObjectBlob   =   "FRMPRESUAHP.frx":1B0B92
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   6975
      OleObjectBlob   =   "FRMPRESUAHP.frx":1B0DC6
      TabIndex        =   71
      Top             =   2160
      Visible         =   0   'False
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   270
      Left            =   11295
      OleObjectBlob   =   "FRMPRESUAHP.frx":1B0E24
      TabIndex        =   72
      Top             =   6885
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3555
      Left            =   0
      TabIndex        =   79
      Top             =   2880
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   6271
      _Version        =   393216
      BackColorFixed  =   16501405
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label SIGRABO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10320
      TabIndex        =   105
      ToolTipText     =   "Doble-Click para 'Dólares'"
      Top             =   6960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label TICAMBIS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8520
      TabIndex        =   80
      ToolTipText     =   "Doble Click para Cambiar Cotizacion"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1935
      TabIndex        =   76
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8520
      TabIndex        =   75
      ToolTipText     =   "Doble-Click para 'Dólares'"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "T-C:"
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
      Left            =   8040
      TabIndex        =   74
      Top             =   7095
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%:"
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
      Left            =   12570
      TabIndex        =   73
      Top             =   7035
      Width           =   225
   End
   Begin VB.Menu mnu 
      Caption         =   "Menú"
      Begin VB.Menu MNUEDITPRESUP 
         Caption         =   "Editar Presupuesto"
      End
      Begin VB.Menu mnumodificarCliente 
         Caption         =   "Modificar Cliente a Pedido Activo"
      End
   End
   Begin VB.Menu pxcopia 
      Caption         =   "Pres.X.Copia"
      Begin VB.Menu mnupcopia 
         Caption         =   "Pres.x.Copia"
      End
   End
   Begin VB.Menu mnuvis 
      Caption         =   "Ver"
      Begin VB.Menu mnuvisualiza 
         Caption         =   "Visualizar Presupuesto"
      End
   End
   Begin VB.Menu mnuconfigrep 
      Caption         =   "Tablas de configuración"
      Begin VB.Menu MnuTabEntr 
         Caption         =   "Tabla de Vias de Entrega"
      End
      Begin VB.Menu MnuArt_Prov 
         Caption         =   "Asociar Art. a Proveedor"
      End
      Begin VB.Menu MnuProv_Art 
         Caption         =   "Asociar Prov. a Artículo"
      End
      Begin VB.Menu mnuDemora 
         Caption         =   "Tabla de Demoras"
      End
      Begin VB.Menu MnuLimvias 
         Caption         =   "Límiites Segun Vias"
      End
      Begin VB.Menu mnutcoef 
         Caption         =   "Tipo de Coeficientes"
      End
      Begin VB.Menu mnuasigcoef 
         Caption         =   "Asignación de Coeficientes"
      End
      Begin VB.Menu mnuvalofer 
         Caption         =   "Validez de Oferta"
      End
   End
   Begin VB.Menu mnuA 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuP 
         Caption         =   "Pedidos de Clientes"
      End
      Begin VB.Menu mnutamonedas 
         Caption         =   "Tabla de Monedas"
      End
      Begin VB.Menu muntexobs 
         Caption         =   "Texto de Observaciones"
      End
      Begin VB.Menu mnutexmail 
         Caption         =   "Texto de Cuerpo de Mail"
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu eliminaitemssinmovimientos 
         Caption         =   "Eliminar Items Sin Movimientos"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu desdecodigooriginal 
         Caption         =   "Eliminar Relación Código Cliente"
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu AltaContacto 
         Caption         =   "Alta de Contacto"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTablaEquiClientes 
         Caption         =   "Tabla de Equipos de Clientes"
      End
      Begin VB.Menu mnutablaequiposxcliente 
         Caption         =   "Tabla de Equipos x Cliente"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnurepocli 
         Caption         =   "Reportes de Equipos de Clientes"
      End
   End
   Begin VB.Menu MNUPR 
      Caption         =   "Procesos"
      Begin VB.Menu mnutraz 
         Caption         =   "Asignar Marca de Trazabilidad 1 a Todos los Artículos"
      End
      Begin VB.Menu mnumigramaster 
         Caption         =   "Migrar a Tabla Master2 (Solo Con Supervisión Sysflom)"
      End
   End
   Begin VB.Menu mnuAgrLoteAjus 
      Caption         =   "Agregar Lote a Ajustes"
   End
   Begin VB.Menu mnus 
      Caption         =   "Setup"
      Begin VB.Menu mnusetup 
         Caption         =   "Setup"
      End
   End
   Begin VB.Menu mnur 
      Caption         =   "Reportes"
      Begin VB.Menu mnuconfig 
         Caption         =   "Configuración de Reportes"
      End
      Begin VB.Menu mnuestacotiVSventas 
         Caption         =   "Estadísticas de Cotizaciones Vs. Ventas"
      End
      Begin VB.Menu mnuItemSugeridosIncorporarStock 
         Caption         =   "Items  Sugeridos Para incorporar Stock"
      End
      Begin VB.Menu mnuEfectividadCotizcionesVSVentas 
         Caption         =   "Efectividad de Cotizaciones vs. Ventas"
      End
      Begin VB.Menu mnucrt 
         Caption         =   "Selección de Reportes"
      End
   End
   Begin VB.Menu mnuoc 
      Caption         =   "Orden de Compra"
      Begin VB.Menu mnuctoc 
         Caption         =   "Control de Generación de O.Compra"
      End
      Begin VB.Menu munanusol 
         Caption         =   "Anular Solicitudes de Orden de Compra"
      End
      Begin VB.Menu mnuvidetasol 
         Caption         =   "Visualizar Detalle de Solicitud de Compra"
      End
   End
End
Attribute VB_Name = "FRMPRESUAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HAY QUE DESPLEGAR DEL BOTON DE SELECCION DE CODIGO QUE MUESTRE LOS ITEMS DE LA LISTA DE PRECIOS
'ACTIVA , DEBERIA MOSTRAR LOS ITEMS CON PRECIO.
'VER TAMBIEN EL CASO EN QUE SE UTILICE LISTA 0 PARA QUE APAREZCAN TODOS LOS ITEMS EN EL BUSCADOR SIN PRECIO
'CARGA DE DETALLE POR TEXTO LIBRE
Public EDICION%
Dim CTRLX As Boolean
Dim CE As Boolean


Sub ACTIVA_OBJ_SIN_STOCK()
          INFOSIAHP.Frame1.Top = 1920
          INFOSIAHP.Height = 6500
          INFOSIAHP.Text6(3).Visible = True
          INFOSIAHP.TXTNUMEPROVE.Visible = True
          INFOSIAHP.Command5.Visible = True
          INFOSIAHP.Text1.Visible = True
          INFOSIAHP.Command3.Visible = True
          INFOSIAHP.Text16.Visible = True
          INFOSIAHP.Text40.Visible = True
          INFOSIAHP.TXTCOSTO.Visible = True
          INFOSIAHP.TXTVIAENTREGA.Visible = True
          INFOSIAHP.Combo6.Visible = True
          INFOSIAHP.Command1.Visible = False
          INFOSIAHP.Command4.Visible = False
End Sub

Sub ACTIVA_OBJ_CON_STOCK()
          INFOSIAHP.Frame1.Top = 4560
          INFOSIAHP.Height = 9135
          INFOSIAHP.Text6(3).Visible = False
          INFOSIAHP.TXTNUMEPROVE.Visible = False
          INFOSIAHP.Command5.Visible = False
          INFOSIAHP.Text1.Visible = False
          INFOSIAHP.Command3.Visible = False
          INFOSIAHP.Text16.Visible = False
          INFOSIAHP.Text40.Visible = False
          INFOSIAHP.TXTCOSTO.Visible = False
          INFOSIAHP.TXTVIAENTREGA.Visible = False
          INFOSIAHP.Combo6.Visible = False
          INFOSIAHP.Command1.Visible = True
          INFOSIAHP.Command4.Visible = True
End Sub

Sub ACTUPESO()
   TXTTOTALPESO = 0
   For i% = 1 To MSF1.Rows - 1
     PESOUNIT = XVALO(MSF1.TextMatrix(i%, 14))
     CANTI = XVALO(MSF1.TextMatrix(i%, 4))
     TXTTOTALPESO = TXTTOTALPESO + (PESOUNIT)
     TXTTOTALPESO.Refresh
  Next i%
  TXTTOTALPESO = FORMATNUM$(XVALO(TXTTOTALPESO), "F10.4")
  
End Sub

Sub BorrarLotesenMovimientosNoCorresponde()
'PRIMERO MODIFICA EN CANTSALID O CANTINGRE QUE ESTEN CON UN VALOR < 0
'EN REALIDAD HAY SOLO 9 REGISTROS QUE INVOLUCRAN A 4 CODIGOS SE PASA EL VALOR NEGATIVO A VALOR POSITIVO
'SE COMENTA POR QUE EN AHP NO SE DETECTARON CASOS
'SQLX$ = "UPDATE MOVISTO SET CANTINGRE = ABS(CANTINGRE) WHERE CANTINGRE < 0":
'FLEXCONN.Execute (SQLX$)
'SQLX$ = "UPDATE MOVISTO SET CANTSALID = ABS(CANTSALID) WHERE CANTSALID < 0 AND (CODIMOVI = 'RT' OR CODIMOVI = 'CI')"
'FLEXCONN.Execute (SQLX$)

'ESTE PROCESO LE BORRA EL LOTE A MOVIMIENTOS QUE SE GENERARON CUANDO SE HACIA EL REMITO DESDE ESE MODULO
'Y EL ITEM ERA UN KIT, Y LE APLICABA EL LOTE DEL KIT AL COMPONENTE, EN UNA RUTINA QUE GENERA UNA SALIDA Y UNA ENTRADA
'PARA EL KIT Y LUEGO GENERA LAS SALIDAS DE LOS COMPONENTES DEL KIT

    SQLX$ = " SELECT m.COD_INTE, m.IDENLOTE"
    SQLX$ = SQLX$ + " FROM MOVISTO m WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN ("
    SQLX$ = SQLX$ + "     SELECT IDENLOTE"
    SQLX$ = SQLX$ + "     From Movisto"
    SQLX$ = SQLX$ + "     GROUP BY IDENLOTE"
    SQLX$ = SQLX$ + "     HAVING COUNT(DISTINCT COD_INTE) > 1 and idenlote <> ''"
    SQLX$ = SQLX$ + " ) subquery ON m.IDENLOTE = subquery.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE  not EXISTS ("
    SQLX$ = SQLX$ + "     SELECT 1"
    SQLX$ = SQLX$ + "     FROM BOLRECEP r"
    SQLX$ = SQLX$ + "     Where R.Cod_Inte = m.Cod_Inte And R.IDENLOTE = m.IDENLOTE"
    SQLX$ = SQLX$ + " )"
    SQLX$ = SQLX$ + " order by m.idenlote;"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        LOTE$ = TRIM$(SAFETEXT$(!idenlote))
        CONDI$ = " IDENLOTE = '" & LOTE$ & "' And COD_INTE = " & CI1%
        Call ACTUREGISTRO("MOVISTO", "IDENLOTE", CONDI$, "", RegAf&, "NOMESS")
        .MoveNext
      Loop
    End With
    Set rs = Nothing
'Explicación de la consulta:
'
'
'La subconsulta subquery obtiene los lotes que están repetidos para más de un código en la tabla "movimientos".
'Utilizamos GROUP BY lote y HAVING COUNT(DISTINCT codigo) > 1 para asegurarnos de que solo se seleccionen los lotes
'que tienen más de un código asociado.
'
'Luego, la consulta principal selecciona los códigos y lotes de la tabla "movimientos" que coinciden con los lotes
'obtenidos en la subconsulta. Utilizamos un INNER JOIN con la subconsulta para obtener solo los registros con lotes
'repetidos.
'
'Finalmente, utilizamos un WHERE EXISTS para asegurarnos de que la combinación de código y lote también exista
'en la tabla "recepcion". Esto garantiza que solo se seleccionen los registros cuya combinación sea válida en ambas tablas.

'****  LE CAMBIE EL EXISTS POR NOT EXISTS PARA QUE DEVUELVA LOS QUE NO SE CORRESPONDEN CON EL INGRESO DEL BOLRECEP

End Sub

Function COLORSTATUS$(CANTI, ST, PI, PC)
    'VERDE HAY STOCK ' AHORA SE CAMBIA A AMARILLO PARA QUE SELECCIONE QUE LOTES ELIGE.
    ' AMARILLO: LA CANTIDAD PROYECTADA  +  EL STOCK ES - LO PEDIDO >= A LA DEMANDA
    'ROJO
    'LA CANTIDAD PROYECTADA  +  EL STOCK ES - LO PEDIDO <= A LA DEMANDA
    '
    CANTI1 = XVALO(CANTI)
    If CANTI > 0 Then
      ST1 = XVALO(ST) 'STOCK
      PI1 = XVALO(PI) 'POR INGRESAR
      PC1 = XVALO(PC) 'COMPRAS PENDIENTES
      '
      If CANTI1 <= ST1 Then
         COLORSTATUS$ = "A" 'ANTES ERA VERDE SE CAMBIA A AMARILLO PARA QUE SELECCIONE EL LOTE CORRESPONDIENTE.
      ElseIf (ST1 + PI1) - PC1 >= CANTI1 Then
         COLORSTATUS$ = "A"
      Else
         COLORSTATUS$ = "R"
      End If
    End If
End Function





Sub CREAR_STORED_PROCEDURE_LOCAL()
    If EXISTE_SP%("VECES_Y_CANT_FACTURADA") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE  VECES_Y_CANT_FACTURADA "
    SQLX$ = SQLX$ + "(@cod_inte INT,  @FECHDESDE smalldatetime, @FECHHASTA smalldatetime, @CODIEMPR AS SMALLINT) "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "IF @cod_inte > 0 "
    SQLX$ = SQLX$ + "BEGIN "
    SQLX$ = SQLX$ + "select count(*) AS VECES, cod_inte, cod_exte,descrip,  sum(cantsalid) CANTTOTALFACTURADA ,sum(CANTINGRE) AS DEVOLU "
    SQLX$ = SQLX$ + "FROM MOVISTO where CodiMovi= 'RT' AND "
    SQLX$ = SQLX$ + "FechMov>= @FECHDESDE "
    SQLX$ = SQLX$ + "and FechMov<= @FECHHASTA "
    SQLX$ = SQLX$ + "AND COD_INTE = @cod_inte "
    SQLX$ = SQLX$ + "AND NUMEORDPED >= 0 "
    SQLX$ = SQLX$ + "AND CODIEMPR = @CODIEMPR "
    SQLX$ = SQLX$ + "AND CANTSALID > 0 "
    SQLX$ = SQLX$ + "GROUP BY cod_inte, cod_exte,descrip,CodiMovi "
    SQLX$ = SQLX$ + "END "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If

End Sub

Function DISPONIB_LIBRE(CI1%)
   '
   'STOCK ACTUAL MENOS RESERVADO DE PEDIDOS
   If CI1% > 0 Then
    '
    ST_RESERVADO = SALDO_LOTE_PEDIDO(CI1%, "", 0)
    STAC = STOCKACT(CI1%)
    DISPONIB_LIBRE = STAC - ST_RESERVADO
    If DISPONIB_LIBRE < 0 Then DISPONIB_LIBRE = 0
   End If
   '
End Function

Sub GENERA_TABLA_RELACION_ART_PROV()
   Call CREACAMPO("", "RELACODPROV", "Num_Prov", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Cod_inte", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Precio", 7, 0)
   Call CREACAMPO("", "RELACODPROV", "Moneda", 3, 2)
   Call CREACAMPO("", "RELACODPROV", "Fecha", 8, 0)
   Call CREACAMPO("", "RELACODPROV", "MarBorra", 3, 0)
   Call CREACAMPO("", "RELACODPROV", "Via", 3, 0)
End Sub



Sub importarExcelSugeridosParaIngresarStock(Emision$, Familias$, Gcontable$, CotizadosMas$, UltimosMeses$, Rango$)
    'RUTAEXCEL$ = "X:\AHP\Flexoft\Excel\ItemSugeridosParaIngresarAlStock.XLSX"
    RUTAEXCEL$ = Control("PRESUAHP", "RUTAEXC1")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Reporte Items sugeridos para ingresar al stock")
       GoTo 99
    End If
    
    IMPRESORA$ = Printer.DeviceName
    '
    Screen.MousePointer = 11
    'imprime los datos fijos
    Unload VENTABNEW
    With VENTABNEW.MSF
       .Rows = 7: .Cols = 2
       .TextMatrix(1, 1) = "Emisión: " & Emision$
       If Familias$ <> "" Then
          .TextMatrix(2, 1) = "Familia: " & Familias$
       Else
          .TextMatrix(2, 1) = "Familia: Todas"
       End If
       If Gcontable$ <> "" Then
          .TextMatrix(3, 1) = "G.Contable: " & Gcontable$
       Else
          .TextMatrix(3, 1) = "G.Contable: Todos"
       End If

         .TextMatrix(4, 1) = "Cotizado más de: " & CotizadosMas$ & " veces"
         .TextMatrix(5, 1) = "En los Últimos : " & UltimosMeses$ & " meses"
         .TextMatrix(6, 1) = "Rango: " & Rango$
         
    End With
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If GRILLAGRAL.TXTBUSCAR <> "" Then Set MSF = GRILLAGRAL.MSF_2
    
    If TRIM$(Control("BLOQEXCE", "BL0INGST")) <> "SI" Then 'BLOQUEO LISTA DE DESPACHOS PARA NO ABRIR 2 AL MISMO TIEMPO
      Call GRACONTROL("BLOQEXCE", "BL0INGST", "SI")
      'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
      'VENTABNEW.Show 1
      If IMPRESORA$ = "" Then IMPRESORA$ = Printer.DeviceName
      Call Param_Print_Excel(RUTAEXCEL$, "HOJA1", VENTABNEW.MSF, IMPRESORA$, "1", 1, 19, "", 1, 1)                            'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
      Unload VENTABNEW
'     A = FilaMaximaConContenidoDentroDeUnRangodeColumnas&("C:\Users\Omar\Desktop\Libro2.xlsx", "Hoja1", "C:J", 4)
     Call BorrarContenidoExcelXColumnas(RUTAEXCEL$, "Hoja1", "A5:L", 5)
     Call ExportarAExcel_MSF_IA(MSF, RUTAEXCEL$, "Hoja1", IMPRESORA$, 4, 0, "A1:L", 4, "$1:$4")
     Call GRACONTROL("BLOQEXCE", "BL0INGST", "NO")
    Else
      Call COMUNI("Archivo Bloqueado, Aguarde y Vuelva a Intentar")
      GoTo 99
    End If
99  Screen.MousePointer = 1
End Sub
Sub importarExcelPorcentajeEfectividadCotizacionesVsVentas(Emision$, Familias$, Gcontable$, PORCENTAJE$, UltimosMeses$, Rango$)
    'RUTAEXCEL$ = "X:\AHP\Flexoft\Excel\ItemSugeridosParaIngresarAlStock.XLSX"
    RUTAEXCEL$ = Control("PRESUAHP", "RUTAEXC2")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Reporte Items sugeridos para ingresar al stock")
       GoTo 99
    End If
    
    IMPRESORA$ = Printer.DeviceName
    '
    Screen.MousePointer = 11
    'imprime los datos fijos
    Unload VENTABNEW
    With VENTABNEW.MSF
       .Rows = 7: .Cols = 2
       .TextMatrix(1, 1) = "Emisión: " & Emision$
       If Familias$ <> "" Then
          .TextMatrix(2, 1) = "Familia: " & Familias$
       Else
          .TextMatrix(2, 1) = "Familia: Todas"
       End If
       If Gcontable$ <> "" Then
          .TextMatrix(3, 1) = "G.Contable: " & Gcontable$
       Else
          .TextMatrix(3, 1) = "G.Contable: Todos"
       End If

         .TextMatrix(4, 1) = "Efectividad Mayor al: " & PORCENTAJE$ & " %"
         .TextMatrix(5, 1) = "En los Últimos : " & UltimosMeses$ & " meses"
         .TextMatrix(6, 1) = "Rango: " & Rango$
         
    End With
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If GRILLAGRAL.TXTBUSCAR <> "" Then Set MSF = GRILLAGRAL.MSF_2
    
    If TRIM$(Control("BLOQEXCE", "BLOEFCOT")) <> "SI" Then 'BLOQUEO LISTA DE DESPACHOS PARA NO ABRIR 2 AL MISMO TIEMPO
      Call GRACONTROL("BLOQEXCE", "BLOEFCOT", "SI")
      'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
      'VENTABNEW.Show 1
      If IMPRESORA$ = "" Then IMPRESORA$ = Printer.DeviceName
      Call Param_Print_Excel(RUTAEXCEL$, "HOJA1", VENTABNEW.MSF, IMPRESORA$, "1", 1, 19, "", 1, 1)                            'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
      Unload VENTABNEW
'     A = FilaMaximaConContenidoDentroDeUnRangodeColumnas&("C:\Users\Omar\Desktop\Libro2.xlsx", "Hoja1", "C:J", 4)
     Call BorrarContenidoExcelXColumnas(RUTAEXCEL$, "Hoja1", "A5:L", 5)
     Call ExportarAExcel_MSF_IA(MSF, RUTAEXCEL$, "Hoja1", IMPRESORA$, 4, 0, "A1:L", 4, "$1:$4")
     Call GRACONTROL("BLOQEXCE", "BLOEFCOT", "NO")
    Else
      Call COMUNI("Archivo Bloqueado, Aguarde y Vuelva a Intentar")
      GoTo 99
    End If
99  Screen.MousePointer = 1
End Sub


Sub MOVER_ARRIBA(REG&)
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

    txtcodigo = MSF1.TextMatrix(REG&, 1)
    TXTCODCLIENTE = MSF1.TextMatrix(REG&, 2)
    TXTDESCRIPCION = MSF1.TextMatrix(REG&, 2)
    TXTCANTIDAD = MSF1.TextMatrix(REG&, 4)
    Text15 = MSF1.TextMatrix(REG&, 7)
    TXTPREUNIT = MSF1.TextMatrix(REG&, 8)
    TXTIMPORTE = MSF1.TextMatrix(REG&, 9)
    Text38 = MSF1.TextMatrix(REG&, 10)
End Sub

Sub PRESENTA_INFO_ARTICULO(CIXI%, Optional Fila&, Optional SINSTOCK%)
   '
   Screen.MousePointer = 11
   TXX1TRAZA$ = TXX1TRAZA$ + "PRESENTA_INFO_ARTICULO Inicio: " & Now & vbCrLf
   Call ProcesoAjusteDeLotes(CIXI%)
   '
   INFOSIAHP.TXTTRANSITO_OTROS = 0
   INFOSIAHP.txtcodigo = VectArticulos.Codigo(CIXI%)
   DoEvents
   INFOSIAHP.TXTDESCRIPCION = VectArticulos.Descripcion(CIXI%)
   INFOSIAHP.TXTCODICLIENTE = TXTCODCLIENTE
   INFOSIAHP.TXTSOLICITADA = XVALO(TXTCANTIDAD)
   INFOSIAHP.TXTPESO = FORMATNUM$(VectArticulos.PesoUn(CIXI%), "F10.2")
   If XVALO(INFOSIAHP.TXTPESO) < 0.0005 Then INFOSIAHP.TXTPESO.BackColor = &H80000005: INFOSIAHP.TXTPESO.Locked = False ' LO PONE EN BLANCO
'   ESTO SE MUESTRA EN EL CASO QUE SEA SIN STOCK, Y SI ES ASI NO DEBERIA MOSTAR COSTO SI NO QUE SE DEBERIA CARGAR.
   INFOSIAHP.TXTCOSTO = FORMATNUM$(XVALO(Text15), "F10.2")
     If XVALO(INFOSIAHP.TXTCOSTO) = 0 Then INFOSIAHP.TXTCOSTO.BackColor = &H80000005: INFOSIAHP.TXTCOSTO.Locked = False ' LO PONE EN BLANCO

   F$ = "F11." & CNTDC$(CI1%)  'FORMATO
   ENCA$ = "Proveedor/A12;Vía Ent./A10;Stock/F6;Reserv./F8;Tránsito/F8;F.Ing./D8;Cto.Fob U$S/F8.2;Costo Nac./F9.2;%/F6.2;P.Vta.U$S/F9.2;Cant./F6;Dispon./F7;Lote/A12;NUMEROPROVE/A0"
   Call CARGA_ENCABEZADO(INFOSIAHP.MSF1, ENCA$, INFOSIAHP)
   '

   PESOX = VectArticulos.PesoUn(CIXI%)

   CI1% = CIXI%
   INFOSIAHP.TXTTRANSITO_OTROS = 0

   If SINSTOCK% > 0 Then
      COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F8.0")     ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
      CANRECEP1 = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(0, CI1%), "F8.0")
      ST_TRANSITO = XVALO(COMPRAX) + XVALO(CANRECEP1)
      INFOSIAHP.TXTTRANSITO_OTROS = FORMATNUM$(ST_TRANSITO, "F10.1")
      Call INFOSIAHP.EstadoVisibleReservasSinLote(False)
      GoTo 90
   End If
   acudisponible = 0
    Call APERBASDAT("STOCKS")

    'AQUI TOMO LAS RESERVAS TOTALES DEL ITEM.
    SQLX$ = " SELECT * FROM SALOTEDEP WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI1%
    SQLX$ = SQLX$ + " AND SALDO > 0"
    SQLX$ = SQLX$ + " ORDER BY FECHMOV ASC "
    '
    Set LOTINGTEMP = READSET(SQLX$)
   
    JJ& = 0
    'AGREGADO VALIDACION DISPONIBLES SIN LOTE
    ReservasDePedidoSinLote = SALDO_SIN_LOTE_PEDIDO(CI1%, RetornoStringPedidos$)
     
    With LOTINGTEMP
25     If Not (.BOF And .EOF) Then
         Do While Not .EOF
           TTXX$ = REBLA$
           'FEX% = CVINT(!FechAlta)
           FEX% = CVINT(!FECHMOV)
           LOTE$ = SAFETEXT$(!idenlote)
           SLOTE = XVALO(!SALDO) 'SaldoLote(CIXI%, LOTE$, DEPO%)
           '
           
           ST_RESERVADO = SALDO_LOTE_PEDIDO(CIXI%, LOTE$, 0)
           
'           CANRES = XVALO(!CanReser) + ST_RESERVADO
           CANRES = ST_RESERVADO
           DISPONIBLE = SLOTE
           
           'SI TIENE STOCK LO MUESTRA EN CASO CONTRARIO PASA AL SIGUIENTE.
           If DISPONIBLE <= 0.05 Then GoTo 29

           
           NPROV = PROVELOTE%(LOTE$)
           VIA_LOTE% = VIALOTE%(LOTE$)
           J& = J& + 1
           INFOSIAHP.MSF1.Rows = J& + 1
           INFOSIAHP.MSF1.TextMatrix(J&, 1) = TEL2CLI$(-1 * NPROV)
           INFOSIAHP.MSF1.TextMatrix(J&, 2) = ECOARCH$("VIASENTR", MKI$(XVALO(INFOSIAHP.VIAENTREGA(LOTE$))))  'DEVUELVE EL LOTE LA VIA DE ENT (SE CARGA EN LA O.C.)
           INFOSIAHP.MSF1.row = J&
           INFOSIAHP.MSF1.COL = 3
           INFOSIAHP.MSF1.CellFontBold = True
           INFOSIAHP.MSF1.TextMatrix(J&, 3) = DISPONIBLE
           INFOSIAHP.MSF1.TextMatrix(J&, 4) = CANRES
           COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(NPROV, CI1%), "F8.0")     ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
           CANRECEP1 = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(NPROV, CI1%), "F8.0")
           
           ST_TRANSITO = XVALO(COMPRAX) + XVALO(CANRECEP1)
           INFOSIAHP.MSF1.TextMatrix(J&, 5) = ST_TRANSITO
           INFOSIAHP.MSF1.TextMatrix(J&, 6) = FECHATEX$(FEX%)
           COSFOB# = ROUND(COSTOFOB#(LOTE$), 2)  ' TRAE EL COSTO DE LA ORDEN DE COMPRA
             
             MONECOMP% = MONELOTE%(LOTE$)
             COEFCON = 1
             If MONECOMP% <> MONEMI% Then
                'ESTO ES PARA QEU CUANDO SE COMPRA EN $ TENGA EL TIPO DE CAM DE LA OC Y LO PASE A U$S
                
                TipoCambioCom = TipoCambioCompra(LOTE$)
                
                If TipoCambioCom > 0 Then
                   COEFCON = TICAMONE(MONECOMP%) / TipoCambioCom
                Else
                   'COEFCON = TICAMONE(MONECOMP%) / TICAMONE(MONEMI%)
                   COEFCON = TICAMONE(MONECOMP%) / 9.42 'ultimo valor antes de antes de guardar el valor de la cotizadcion en la o.c
                End If
             End If
             
             COSFOB# = COEFCON * COSFOB#
             'SIMBOMON$ = SIMBOLPES$(MONELOTE%(LOTE$)) 'MONELOTE% DEVUELVE LA MONEDA CON LA CUAL SE HIZO LA ORDEN DE COMPRA DE ESTE LOTE
           'INFOSIAHP.MSF1.TextMatrix(j&, 7) = SIMBOMON$ & " " & TRIM$(FORMATNUM$(COSFOB#, "F10.2"))
           INFOSIAHP.MSF1.TextMatrix(J&, 7) = TRIM$(FORMATNUM$(COSFOB#, "F10.2"))
             NUMLOTE& = XVALO(Mid$(LOTE$, 4))
          'CONDI$ = " VAL_COMPRA > 13490 AND VAL_COMPRA < 13510 "
          
           PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NUMLOTE&, , MONECOMP%)
           GASNAC# = XVALO(COSFOB# * PORCEN_GASNAC#)
           MODONACIONALIZACION% = 0

           If GASNAC# <= 0 Then
              GASNAC# = COSFOB#
              MODONACIONALIZACION% = 1
           End If
           INFOSIAHP.MSF1.TextMatrix(J&, 8) = TRIM$(FORMATNUM$(GASNAC#, "F10.2"))
             
             PREVENTAX# = PRECIOVENTA_X_FORMULA(XVALO(GASNAC#), VIA_LOTE%, CI1%, MODONACIONALIZACION%, NPROV, 1)
           INFOSIAHP.MSF1.TextMatrix(J&, 9) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(PREVENTAX#, (GASNAC#)), "F9.2")
           If XVALO(INFOSIAHP.MSF1.TextMatrix(J&, 9)) <= 0 Then Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, vbRed, J&, 9, 1)
           INFOSIAHP.MSF1.TextMatrix(J&, 10) = FORMATNUM$(PREVENTAX#, "F12.2")
           INFOSIAHP.MSF1.TextMatrix(J&, 11) = ""
           Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &HFFFF&, J&, 11)
                INFOSIAHP.MSF1.COL = 12
                INFOSIAHP.MSF1.CellFontBold = True
                FILA1& = 0: If EDICION% > 0 Then FILA1& = J& ' SI ESTA EN EDICION NO TIENE EN CUENTA LA FILA A EDITAR
                Disponibilidad = XVALO(INFOSIAHP.MSF1.TextMatrix(J&, 3)) - XVALO(INFOSIAHP.MSF1.TextMatrix(J&, 4)) - CARGADO_ENGRILLA(LOTE$, FILA1&)
                If Disponibilidad < 0 Then Disponibilidad = 0
                JJJ% = J&
                
                If Disponibilidad = 0 Then Call borrafila(JJJ%, INFOSIAHP, INFOSIAHP.MSF1): J& = J& - 1: GoTo 29
           INFOSIAHP.MSF1.TextMatrix(J&, 12) = Disponibilidad
           INFOSIAHP.MSF1.TextMatrix(J&, 13) = LOTE$
           INFOSIAHP.MSF1.TextMatrix(J&, 14) = NPROV
           acudisponible = acudisponible + Disponibilidad
           '
29       .MoveNext
         Loop
       End If
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
     'AGREGADO VALIDACION DISPONIBLES SIN LOTE (CODIGO EN PEDIDOS SIN LOTE ASIGNADO)
    INFOSIAHP.Label2 = "": INFOSIAHP.Label3 = "": INFOSIAHP.Label6 = ""
    If ReservasDePedidoSinLote > 0 Then
        Call INFOSIAHP.EstadoVisibleReservasSinLote(True)
        acudisponible = acudisponible - ReservasDePedidoSinLote 'DESCUENTO TAMBIEN LOS RESERVADOS SIN LOTE
        INFOSIAHP.Label2 = TRIM$(FORMATNUM$(ReservasDePedidoSinLote, "F10.1"))
        INFOSIAHP.Label3 = RetornoStringPedidos$
        INFOSIAHP.Label6 = acudisponible
    End If
    '
    If acudisponible < 0 Then acudisponible = 0
    
    'SI NO TIENE DISPONIBLE PRESENTA VENTANA SIN STOCK
    If acudisponible <= 0 Then
          Call ACTIVA_OBJ_SIN_STOCK
          Call INFOSIAHP.EstadoVisibleReservasSinLote(False)
          Call INFOSIAHP.TXTDESCRIPCION_Change
          
'          Call PRESENTA_INFO_ARTICULO(CIXI%, 0, 1)
    End If
    'AHORA VEO SI HAY COMPRAS O RECEPCIONES NO APROBADAS DE OTROS PROVEEDORES
    'UTILIZANDO LA MISMA FUNCION SIN EL PARAMENTRO DEL PROVEEDOR
    'SI LA CANTIAD SUPERA LO QUE ESTA EN TRANSITO, YA CARGADO MUESTRA UNA FILA MAS INDICANDO COMO TRANSITO DE OTROS EL SALDO EN TRANSITO
    'ORDENO LA GRILLA POR PROVEEDOR Y VIA
    MSF1.Redraw = False
    ACUMULO_TRANSITO = 0
    MSF1.COL = 1: MSF1.Sort = 7
    MSF1.COL = 2: MSF1.Sort = 7
    '
    'RECORRO UNA VEZ ORDENDADO Y ACUMULO LO YA PRESENTADO COMO EN TRANSTO
    For i& = 1 To INFOSIAHP.MSF1.Rows - 1
       PROVEEX$ = TRIM$(UCase$(INFOSIAHP.MSF1.TextMatrix(i&, 1)))
       If PROVEEX$ = PROVEEX_ANT$ Then INFOSIAHP.MSF1.TextMatrix(i&, 4) = "": INFOSIAHP.MSF1.TextMatrix(i&, 5) = ""
       PROVEEX_ANT$ = TRIM$(UCase$(INFOSIAHP.MSF1.TextMatrix(i&, 1)))
       ACUMULO_TRANSITO = ACUMULO_TRANSITO + XVALO(INFOSIAHP.MSF1.TextMatrix(i&, 5))
       STOCKSINRESERVAR = XVALO(INFOSIAHP.MSF1.TextMatrix(i&, 3))
       DISPONIB = XVALO(INFOSIAHP.MSF1.TextMatrix(i&, 12))
       INFOSIAHP.MSF1.TextMatrix(i&, 4) = TRIM$(STOCKSINRESERVAR - DISPONIB) ' PRESENTO LA CANTIDAD RESERVADA POR DIFERENCIA.
    Next i&
    '
    'COMPARO SI EL TOTAL EN TRANSITO (SIN FILTRAR POR PROVEEDOR) ES MAYOR QUE LO PRESENTADO EN LA GRILLA MUESTRO LA DIFERENCIA, COMO TRANSITO EN OTROS.
    COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F8.0")     ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
    CANRECEP1 = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(0, CI1%), "F8.0")
    
    TOTAL_TRANSITO = XVALO(COMPRAX) + XVALO(CANRECEP1)
    If TOTAL_TRANSITO > ACUMULO_TRANSITO Then
       INFOSIAHP.TXTTRANSITO_OTROS = TOTAL_TRANSITO - ACUMULO_TRANSITO
    End If
    '
    Screen.MousePointer = 1
    '
90 INFOSIAHP.txtcodigo = txtcodigo
   INFOSIAHP.TXTDESCRIPCION = TXTDESCRIPCION
   Screen.MousePointer = 1
   INFOSIAHP.FILAX& = Fila&
   INFOSIAHP.Text3 = MONEMIS.Caption
   '
   'VUELVO AL ORDENAMIENTO ORIGINAL
   MSF1.COL = 0
   MSF1.Sort = 3 'ordena de manera numerica
   MSF1.Redraw = True
   '
   TXX1TRAZA$ = TXX1TRAZA$ + "Fin de PRESENTA_INO_ARTICULO Antes del show INFOSIAHP: " & Now & vbCrLf

   INFOSIAHP.Timer1.Enabled = True
   INFOSIAHP.Show 1
   '
End Sub
''Function PRECIOVENTA_X_FORMULA(COSTO1#, VIA%, CI%, CALCULAR_VALOR_SOBRE_FOB%, Optional NPROV)
''   RESPALDO 28/08/15
''   'EN COSTO1# YA VIENE CON EL COSTO NACIONALIZADO, COMO BANDERA TIENE EL PARAMETRO 'CALCULAR_VALOR_SOBRE_FOB%' SI ESTA ES > 0 NO LE CALCULA EL COSTOSOBREFOB DE
''   'LA FORMULA STANDARD.
'''   NPROV = XVALO(NPROV)
''   PRECIOVENTA_X_FORMULA = 0
''   Select Case VIA%
''    Case 1
''        If NPROV > 0 Then
''            RFF$ = RECFILT$("!PORCCOEF", 6, MKI$(NPROV))
''            FIN& = Len(RFF$)
''            For UK& = 1 To FIN& Step 4
''               U& = CVS(Mid$(RFF$, UK&, 4))
''               X$ = REGLEIDO$("!PORCCOEF", U&)
''               VIA1% = CVI(Mid$(X$, 1, 2))
''               If VIA1% = VIA% Then
''                  PORC = CVD(Mid$(X$, 5, 8))
''                  PORC = 1 + (PORC / 100)
''                  Exit For
''               End If
''            Next UK&
''        Else
''
''          X$ = FILTERGETRECORD("PORCCOEF", 1, MKI$(VIA%))
''          PORC = CVD(Mid$(X$, 5, 8))
''          PORC = 1 + (PORC / 100)
''        End If
''        '
''        PRECIOVENTA_X_FORMULA = XVALO(COSTO1#) * PORC  '"SEGUN FORMULA"
''        Exit Function
''    Case Else
''        RFF$ = RECFILT$("PORCCOEF", 1, MKI$(VIA%))
''        For KKI% = 1 To Len(RFF$) Step 4
''         RELI& = CVS(Mid$(RFF$, KKI%, 4))
''         If RELI& > 0 Then
''           ZZ$ = REGLEIDO$("PORCCOEF", RELI&)
''           If CVI(Mid$(ZZ$, 13, 2)) = NPROV Or NPROV = 0 Then 'VALIDACION DE PROVEEDOR
''                If CVI(Mid$(ZZ$, 3, 2)) = 2 Then
''                    PORCSOBREFOB = CVD(Mid$(ZZ$, 5, 8))
''                    If Abs(PORCSOBREFOB) < 0.005 Then
''                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Sobre Valor Fob")
''                       Exit Function
''                    End If
''                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 3 Then
''                    COSTOKG = CVD(Mid$(ZZ$, 5, 8))
''                    If Abs(COSTOKG) < 0.005 Then
''                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Costo x Kg.")
''                       Exit Function
''                    End If
''                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 4 Then
''                    INVERSO = CVD(Mid$(ZZ$, 5, 8))
''                    If Abs(INVERSO) < 0.005 Then
''                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Inverso")
''                       Exit Function
''                    End If
''                End If
''           End If
''         End If
''        Next KKI%
''   End Select
''   'PVTA EN USD = (FOB * 1,3150 + KG de peso * 11,0625) / 0,70
''
''   PORCSOBREFOB = 1 + (PORCSOBREFOB / 100)
''   'SI TIENE GASTOS DE NACIONALIZACION CARGADOS NO DEBE CALCUALAR EL PORCENTAJE SOBRE FOB
''   If CALCULAR_VALOR_SOBRE_FOB% > 0 Then VALOR1 = XVALO(COSTO1#) * PORCSOBREFOB ' COSTO  * PORCENTAJE SOBRE FOB
''   CONDI$ = "CODINT = " & CI%
''   VALOR2 = COSTOKG * XVALO(VALOBADA("MASTER", "PESUNI", CONDI$, "NOMESS"))
''   If INVERSO <> 0 Then
''      PRECIOVENTA_X_FORMULA = (VALOR1 + VALOR2) / INVERSO
''   End If
''
''End Function





Sub PRINTPRES()

    Call PRIPRESU(1)

End Sub
Sub PRIPRESU(VIPREVIA%)
    '
'    FORIPRE$ = "PRRS-AHP"
    NOIMPRIMECODIGO% = 0
    If COMALTER%("Desea Imprimir Códigos de Los Items de Este Presupuesto\\Sí, Imprimir Códigos\No, No imprimir Códigos") = 2 Then
       NOIMPRIMECODIGO% = 1
    End If
    FORIPRE$ = TRIM$(Control$("", "FORPRERE"))
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       TX$ = "FALTA FORMULARIO DE IMPRESION DE PRESUPUESTOS"
       TX$ = TX$ + "EL PRESUPUESTO SE HA GUARDADO, NO SALDRÁ IMPRESO"
       MsgBox TX$
       Exit Sub
    End If
    '
    Call LIMPIATETAB
    '

    NUMEDOC$ = TRIM$(Str$(XVALO(Text17)))
    NC% = XVALO(Text1.TEXT)
    TIPODOC$ = "Presupuesto"
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIPODOC$ = "COTIZACION"

    DESTIDOC% = NC%
    FECHDOC$ = FECHATEX$(CVINT(DTPicker1.Value))
'    CANCELPRINT% = 1
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHATEX$(CVINT(DTPicker1.Value))   'FECHA
    CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = TRIM$(Text2)
    CODPARAM$(3) = "DOMI-CLI": DOCPARAM$(3) = TRIM$(Text3)
    CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = TRIM$(Text4)
    CODPARAM$(5) = "TELE-CLI": DOCPARAM$(5) = TRIM$(Text8)
    CODPARAM$(6) = "REF-MAT3":  DOCPARAM$(6) = TRIM$(TXTREFERENCIA) 'REFERENCIA
    CODPARAM$(7) = "EQUIPO1":  DOCPARAM$(7) = TRIM$(Text5) 'CELLUALR
    CODPARAM$(8) = "IMP-IVA":  DOCPARAM$(8) = TRIM$(Text31)
    CODPARAM$(9) = "IMP-TOTA":  DOCPARAM$(9) = TRIM$(Text14.TEXT) 'NONTO TOTAL
    CODPARAM$(10) = "NUME-COM":  DOCPARAM$(10) = Text17   'NUMERO DE DOCUMENTO
    CODIPROVINCIA% = XVALO(VALOBADA("DEUDOR12", "Prov_Cli", CONDI$, "NOMESS"))
'    CODPARAM$(11) = "PCIA-CLI":  DOCPARAM$(11) = ECOPROVINCIA$(CODIPROVINCIA%) 'PROVINCIA
    CODPARAM$(12) = "OBSERVA": DOCPARAM$(12) = ENTRAOBS.Text28
    CODPARAM$(13) = "PES-COMP": DOCPARAM$(13) = TXTTOTALPESO
    CODPARAM$(14) = "IMP-NETO": DOCPARAM$(14) = TRIM$(Total(o))
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = PORDESCU
    CODPARAM$(16) = "IMP-DESC": DOCPARAM$(16) = TRIM$(Total(1))
    CODPARAM$(17) = "IMPNEDES": DOCPARAM$(17) = TRIM$(Text33)
    CODPARAM$(18) = "REF-MAT2": DOCPARAM$(18) = TRIM$(Text35) 'CONTACTO
    CODPARAM$(19) = "REF-MAT4": DOCPARAM$(19) = TRIM$(Text18) 'MAIL
    CODPARAM$(20) = "IMP-TOTA": DOCPARAM$(20) = TRIM$(Text14) ' TOTAL
    '
    CODPARAM$(21) = "COND-PAG": DOCPARAM$(21) = TRIM$(TXTDESCRIPAGO)
    CODPARAM$(22) = "DESTINO": DOCPARAM$(22) = TRIM$(LENTREGA(0) + TRIM$(LENTREGA(1)))
    CODPARAM$(23) = "REF-PLCO": DOCPARAM$(23) = TRIM$(Text28)
    CODPARAM$(24) = "VENDEDOR": DOCPARAM$(24) = TRIM$(TXTNOMVENDEDOR) 'VENDEDOR
        NTRAI% = NUTRANCLI%(NC%)
    CODPARAM$(25) = "TRANSPOR": DOCPARAM$(25) = DENOTRANS$(NTRAI%)
    CODPARAM$(26) = "DOMI-TRA": DOCPARAM$(26) = DOMITRANS$(NTRAI%)
    CODPARAM$(27) = "LOCA-TRA": DOCPARAM$(27) = TRIM$(LOCATRANS$(NTRAI%))
    CODPARAM$(28) = "TELE-EMP": DOCPARAM$(28) = TRIM$(TELETRANS$(NTRAI%))

    CAPARDOC% = 28
    '
    CACOLTAB1% = 1
    CODTABU1$(1) = "CANTIDAD"
    CODTABU1$(2) = "COD-MAT"
    CODTABU1$(3) = "DES-MAT"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "PRE-LIST"
    CODTABU1$(6) = "NET-ITEM"
    CODTABU1$(7) = "REF-MAT5"
    CODTABU1$(8) = "REF-MAT6"
    '
    '''''''''''''''''''''''''''
    CODTABU1$(9) = "LAR-CUTE"  'VIA DE ENTREGA
    CODTABU1$(10) = "LAR-MATE" 'PESO
    CODTABU1$(11) = "ORD-ITEM" 'ORDEN
    CODTABU1$(12) = "COD-MPRI" 'CODIGO DE CLIENTE
    CODTABU1$(13) = "REF-MAT1" 'OBSERVACION EN DETALLE
    CACOLTAB1% = 13
    '
    CAITAB1% = 0
    TIOPCI% = 0 ' BANDERA PARA IMPRIMIR TITULO DE OPCIONALES
    For U& = 1 To MSF1.Rows - 1
        CODE$ = TRIM$(MSF1.TextMatrix(U&, 1))
        CODICLIENTE1$ = TRIM$(MSF1.TextMatrix(U&, 2))
        CIIX% = CODINT%(CODE$)
        DESCRIX$ = MSF1.TextMatrix(U&, 3)
        CAN = XVALO(MSF1.TextMatrix(U&, 4))
        NPROV& = XVALO(MSF1.TextMatrix(U&, 5))
        NOM_FANTASIA$ = MSF1.TextMatrix(U&, 6)
        COSTO# = XVALO(MSF1.TextMatrix(U&, 7))
        PREUNI# = XVALO(MSF1.TextMatrix(U&, 8))
        IMPORTX# = XVALO(MSF1.TextMatrix(U&, 9))
        Via_Entrega$ = Mid$(MSF1.TextMatrix(U&, 10), 4)
        VIA% = XVALO(Left$(MSF1.TextMatrix(U&, 10), 2))
        STRINGLOTES = MSF1.TextMatrix(U&, 11)
        CANTILOTES = XVALO(MSF1.TextMatrix(U&, 12))
        DEMORAX$ = MSF1.TextMatrix(U&, 15)
        
        If DEMORAX$ <> "Inmediata" Then  'INMEDIATA ES QUE POSEE STOCK
          If XVALO(DEMORAX$) < 0 Then ' EN EL CASO QUE DEVUELVA -1 ES QUE FIGURA EN LA TABLA CON DEMORA 0
             DEMORAX$ = "Inmediata"
          ElseIf XVALO(DEMORAX$) >= 0 Then 'SI DEVUELVE 0 ES QUE NO FIGURA EN LA TABLA ,
           DEMORAX$ = DEMORAX$ + " Días"    'SI DEVUELVE UN VALOR POSITIVO FIGURA EN LA TABLA CON VALOR > 0
          End If
        End If
                
        '
        DESCRIOBSERVA$ = DESCRIX$

        If CODE$ = "------------" Or TRIM$(CODE$) = "" Then
           If DESCRIOBSERVA$ <> "" Then
             CAITAB1% = CAITAB1% + 1

             Call PRINTEXT(DESCRIOBSERVA$, 13, "REF-MAT1", "PRRS-AHP")
             TETABU1$(1, CAITAB1%) = TRIM$(FORMATNUM$(CAN, "F8.1"))
             TETABU1$(2, CAITAB1%) = "------------"
             TETABU1$(3, CAITAB1%) = ""

             TETABU1$(4, CAITAB1%) = ""
             TETABU1$(5, CAITAB1%) = ""
             TETABU1$(6, CAITAB1%) = ""
             TETABU1$(7, CAITAB1%) = ""
             TETABU1$(8, CAITAB1%) = ""
             TETABU1$(9, CAITAB1%) = ""
             TETABU1$(10, CAITAB1%) = ""
             TETABU1$(11, CAITAB1%) = U& ' ORDEN
             TETABU1$(12, CAITAB1%) = "------------"
             GoTo 30
           End If
        End If
      '
        'PARA QUE CORTE EN EL PRIMER ESPACIO LA DESCRIPCION, SOLICITADO POR ESTEBAN AHP
        CAITAB1% = CAITAB1% + 1

        DESCRIX$ = TRIM$(DESCRIT0$(CIIX%))
        POS1% = InStr(DESCRIX$, " ")
        If POS1% > 0 Then
          DESCRIX$ = Left$(DESCRIX$, POS1% - 1)
        End If
     '
        TETABU1$(3, CAITAB1%) = DESCRIX$
        TETABU1$(1, CAITAB1%) = TRIM$(FORMATNUM$(CAN, "F8.1"))
        If NOIMPRIMECODIGO% = 0 Then
          TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
          TETABU1$(12, CAITAB1%) = CODICLIENTE1$ ' CODIGO CLIENTE
        Else
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(12, CAITAB1%) = ""
        End If
        TETABU1$(4, CAITAB1%) = TRIM$(Unimed$(CIIX%))
        TETABU1$(5, CAITAB1%) = FORMATNUM$(PREUNI#, "F12.2")
        TETABU1$(6, CAITAB1%) = FORMATNUM$(IMPORTX#, "F12.2")
        TETABU1$(7, CAITAB1%) = PRESENTA$
        TETABU1$(8, CAITAB1%) = ""
        TETABU1$(9, CAITAB1%) = CONDICOM$(-1 * NPROV&)
        TETABU1$(10, CAITAB1%) = DEMORAX$
        TETABU1$(11, CAITAB1%) = U& ' ORDEN
        TETABU1$(13, CAITAB1%) = ""
      
 '
30      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo

        '
    Next U&    '
    ARCHIVOADJUNTOMAIL$ = "DOC_ADJ"
    Call PRINTDOC("@" + FORIPRE$)

    SENDBYMAIL$ = ""
    
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
'    DOCUNU& = DOCID&("PRESUPUESTO", NUMEDOC$, "")
    NPR& = XVALO(Text17)
    Call MUESTRADOC(DOCUNU&, "", TIPODOC$, Str$(NPR&), Str$(NPR&))

    CANCELPRINT% = 0

99  Exit Sub
    '
End Sub




Sub ProcesoAjusteDeLotes(CIXI%)
  'NO ESTA CONSIDERANDO EL CASO EN QUE POR  EJ. HAY VARIOS LOTES CON 1 UNIDAD DE SALDO POSITIVO
  'Y DEBE CUBRIR REGISTROS QUE LOS MOVIMIENTOS SON MAYORES POR EJ. LA DIFERENCIA NEGATIVA ES DE 3 Y NO HAY NINGUN LOTE
  'CON SALDO >= 3 TIENE 3 LOTES DIFERENTES DE 1 C.U.
10 CI1% = CIXI%
  CONDI$ = "COD_INTE = " & CI1%
  'VERIFICA EL SALDO COMO EL KARDEX NO CONSIDERA LOTES

  StockFichaKardex = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE -CANTSALID )", CONDI$, "NOMESS"))
  'VERIFICA EL SALDO SEGUN LOTES
  resultadoDeLotes = 0
  SQLX$ = " select sum(t.total) as resultado from ("
  SQLX$ = SQLX$ + " select sum(cantingre - cantsalid) as total,idenlote from movisto WITH(NOLOCK) where cod_inte = " & CI1%
  SQLX$ = SQLX$ + " group by idenlote having  sum(cantingre - cantsalid) > 0) t"
  Set rs = READSET(SQLX$)
  With rs
    If rs.EOF = False Then
      resultadoDeLotes = XVALO(!Resultado)
    End If
  End With
  Set rs = Nothing
  '
  'VALIDO SI HAY DIFERENCIAS
  If resultadoDeLotes = StockFichaKardex Then Exit Sub 'SI ES IGUAL NO  HACE NADA, SE PODRIA QUITAR ESTO
  
  'LOTES CON SALDO NEGATIVOS INCLUYE LOS QUE NO POSEEN LOTE
  SQLX$ = " select cod_inte,cod_exte,idenlote,sum(cantingre-cantsalid) AS SALDO from movisto WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
  SQLX$ = SQLX$ + " group by cod_inte,cod_exte,idenlote having sum(cantingre-cantsalid) < 0 "
  SQLX$ = SQLX$ + " Order by sum(cantingre-cantsalid)"
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
        'VER EJEMPLO DE CONSULTA A PIE DE RUTINA BUSCA DEL LOTE QUE LA CONSULTA DE ARRIBA
        'MUESTRA COMO QUE ESTAN EN NEGATIVO, LO MUESTRA COMO SI FUERA UN KARDEX Y DONDE GENERA EL NEGATIVO
        'BUSCA DE COMPENSAR CON UN LOTE POSITIVO
        CI1% = XVALO(!cod_inte)
        SALDONEGATIVO = XVALO(!SALDO)
        LOTENEGATIVO$ = TRIM$(SAFETEXT$(!idenlote))
    'SE COMENTA LA QUERY (QUE MUESTRA EN OTRA COLUMNA EL SALDO) Y SE USA OTRA POR QUE OVER NO FUNCIONA EN < 2012
'        SQLX$ = "  SELECT MOVISTO_id,IDENLOTE,CANTINGRE,CANTSALID,SUM(CANTINGRE - CANTSALID)"
'        SQLX$ = SQLX$ + " OVER (ORDER BY CANTSALID,MOVISTO_ID ) AS Diferencia"
'        SQLX$ = SQLX$ + " From  Movisto WITH(NOLOCK)"
'        SQLX$ = SQLX$ + " WHERE COD_INTE=" & CI1%
'        SQLX$ = SQLX$ + " AND IDENLOTE ='" & LOTENEGATIVO$ & "'"
'        SQLX$ = SQLX$ + " Order By CANTSALID,MOVISTO_ID"
        ResultadoDiferencia = 0
        SQLX$ = "  SELECT MOVISTO_id,IDENLOTE,CANTINGRE,CANTSALID"
        'SQLX$ = SQLX$ + " OVER (ORDER BY CANTSALID,MOVISTO_ID ) AS Diferencia"
        SQLX$ = SQLX$ + " From  Movisto WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE COD_INTE=" & CI1%
        SQLX$ = SQLX$ + " AND IDENLOTE ='" & LOTENEGATIVO$ & "'"
        SQLX$ = SQLX$ + " Order By CANTSALID,MOVISTO_ID"
      
        Set RS1 = READSET(SQLX$)
        With RS1
           Do While Not .EOF
           'VER SI ResultadoDiferencia NO HABRIA QUE INICIALIZARLA
           'ver controlar la vuelta de ResultadoDiferencia
           '-2
            ResultadoDiferencia = ResultadoDiferencia + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
'            ResultadoDiferencia = XVALO(!CANTINGRE) - XVALO(!CantSalid)
            idMovLoteNegativo = XVALO(!MOVISTO_ID)
            If ResultadoDiferencia < 0 And idMovLoteNegativo > 0 Then
               'ACA TRAIGO EL LOTE Y LA CANTIDAD DE UN LOTE CON SALDO POSITIVO
               idMovLoteNegativo = XVALO(!MOVISTO_ID)
               SQLX$ = "select cod_inte,idenlote,sum(cantingre-cantsalid) as SaldoPos from movisto WITH(NOLOCK) where idenlote <> ''"
               SQLX$ = SQLX$ + " and cod_INTE =  " & CI1%
               SQLX$ = SQLX$ + " group by cod_inte,idenlote having sum(cantingre-cantsalid) > 0"
               'SQLX$ = SQLX$ + " Order by sum(cantingre-cantsalid) DESC"
               SQLX$ = SQLX$ + " Order by IDENLOTE"
               Set RSloteCSaldo = READSET(SQLX$)
               If RSloteCSaldo.EOF Then GoTo 99
                 
               saldoPositivo = XVALO(RSloteCSaldo!SaldoPos)
               LOTEPOSITIVO$ = TRIM$(SAFETEXT$(RSloteCSaldo!idenlote))
               '
               If saldoPositivo >= Abs(ResultadoDiferencia) Then
                 cantSalidaNegativo = XVALO(RS1!CANTSALID)
                 If Abs(ResultadoDiferencia) = cantSalidaNegativo Then
                      'SI EL NEGATIVO ES IGUAL LA CANTIDAD SALIDA SOLO REEMPLAZA EL LOTE EN EL MISMO REGISTRO
                      CONDI$ = "MOVISTO_ID = " & idMovLoteNegativo
                      Call ACTUREGISTRO("MOVISTO", "IDENLOTE", CONDI$, LOTEPOSITIVO$, RegAf&, "NOMESS")
                      Call ACTUREGISTRO("MOVISTO", "CodCarpImp", CONDI$, LOTENEGATIVO$ & LOTEPOSITIVO$, RegAf&, "NOMESS")
                      rs.Close
                      RS1.Close
                      RSloteCSaldo.Close
                      GoTo 10
                      'APLICA SEGUN ID EL LOTE POSITIVO AL MISMO REGISTRO
                 Else 'SI LA CANTIDAD SALIDA ES <> A LA DIFERENCIA (EJ. SALEN 10 Y HABIAN 3 DIFERENCIA 7
                      'LE MODIFICA LA CANTIDAD SALIDA A 3 Y CREA UN NUEVO REGISTRO IGUAL PERO CON UN LOTE CON SALDO Y
                      'CANTIDAD SALIDA = 7
                      'DuplicarRegistro
                      'ResultadoDiferencia siempre va a ser negativo se suma pero en realidad lo resta
                      cantAAplicaraCantidadSalida = cantSalidaNegativo + ResultadoDiferencia
                      CONDI$ = "MOVISTO_ID = " & idMovLoteNegativo
                      Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, cantAAplicaraCantidadSalida, RegAf&, "NOMESS")
                      Call ACTUREGISTRO("MOVISTO", "CodCarpImp", CONDI$, LOTENEGATIVO$, RegAf&, "NOMESS")

                      '
                      NEWID& = 1 'PARA QUE DEVUELVA ID
                      Call DuplicarRegistro("MOVISTO", CONDI$, NEWID&)
                      CONDI$ = "MOVISTO_ID = " & NEWID&
                      Call ACTUREGISTRO("MOVISTO", "IDENLOTE", CONDI$, LOTEPOSITIVO$, RegAf&, "NOMESS")
                      Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, Abs(ResultadoDiferencia), RegAf&, "NOMESS")
                      Call ACTUREGISTRO("MOVISTO", "CodCarpImp", CONDI$, LOTENEGATIVO$, RegAf&, "NOMESS")
                      rs.Close
                      RS1.Close
                      RSloteCSaldo.Close

                      GoTo 10
                      'CAMBIA LA CANTIDAD SALIDA AL REGISTRO INGRESANDO X CANTSALIDA + (-DIFERENCIA)
                      'DUPLICAR REGISTRO CON EL LOTE POSITIVO, Y CANTSALID = ABS(DIFERENCIA)
                      
                 End If
               Else ' CASO QUE EL SALDO POSITIVO DEL LOTE NO ALCANCE PARA CUBRIR LA CANTIDAD SALIDA
                    PuedeAplicar = saldoPositivo
                    cantSalidaNegativo = XVALO(RS1!CANTSALID)
                    cantAAplicaraCantidadSalida = cantSalidaNegativo - PuedeAplicar
                     CONDI$ = "MOVISTO_ID = " & idMovLoteNegativo
                    Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, cantAAplicaraCantidadSalida, RegAf&, "NOMESS")
                    Call ACTUREGISTRO("MOVISTO", "CodCarpImp", CONDI$, LOTENEGATIVO$, RegAf&, "NOMESS")
                    '
                    NEWID& = 1 'PARA QUE DEVUELVA ID
                    Call DuplicarRegistro("MOVISTO", CONDI$, NEWID&)
                    CONDI$ = "MOVISTO_ID = " & NEWID&
                    Call ACTUREGISTRO("MOVISTO", "IDENLOTE", CONDI$, LOTEPOSITIVO$, RegAf&, "NOMESS")
                    Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, PuedeAplicar, RegAf&, "NOMESS")
                    Call ACTUREGISTRO("MOVISTO", "CodCarpImp", CONDI$, LOTENEGATIVO$, RegAf&, "NOMESS")
                    rs.Close
                    RS1.Close
                    RSloteCSaldo.Close

                    GoTo 10
                 
               End If
            End If
            .MoveNext
            Loop
            End With
            RS1.Close
            Set RS1 = Nothing
           .MoveNext
           Loop
        End With

99   rs.Close
     Set rs = Nothing
' VER SI EN EL CASO DEL CODIGO VOE14663702 NO ENTRA A GRABAR EN NINGUN CASO POR QUE ESTAN TODOS LOS LOTES EN 0 O NEGATIVOS.
' VER PARA ESTOS CASOS HACER UN PROCESO QUE APLIQUE UN LOTE LA-
' VER POR QUE GRABA EN NEGATIVO LA CANTIDAD SALIDA.
 
'  MOVISTO_id  IDENLOTE    CANTINGRE   CANTSALID   Diferencia
'        36994   BR-2813-9       200           0   200
'        11353   BR-2813-9           0         4   196
'        68982   BR-2813-9           0         4   192
'        69222   BR-2813-9           0         4   188
'        42730   BR-2813-9           0         5   183
'        34980   BR-2813-9           0         6   177
'        79049   BR-2813-9   0                 7   170
'        15481   BR-2813-9   0                 10  160
'        58060   BR-2813-9   0                 11  149
'        34565   BR-2813-9   0                 24  125
'        35544   BR-2813-9   0                 30   95
'        41808   BR-2813-9   0                 110 -15


'SELECT MOVISTO_id,IDENLOTE,CANTINGRE,CANTSALID,SUM(CANTINGRE - CANTSALID)
'--OVER (ORDER BY CANTSALID,MOVISTO_ID ) AS Diferencia
'From  Movisto WITH(NOLOCK)
'Where COD_INTE = 1
' AND IDENLOTE ='BR-1880-13'
' GROUP BY MOVISTO_id,IDENLOTE,CANTINGRE,CANTSALID
'Order By CANTSALID,MOVISTO_ID
'ESTO REEMPLAZARIA A LA CONSULTA CON OVER
'recorrer e ir llevando el saldo como si fuera una cuenta corriente,(como hace ahora la diferencia
'sea negativa y ver de hacer como esta ahora).
'______________________________________________________
'ESTA CONSULTA TRAE MISMO LOTE CON DIFERENTES CODIGOS
'SELECT IDENLOTE
'From Movisto
'GROUP BY IDENLOTE
'HAVING COUNT(DISTINCT COD_INTE) > 1;

'ESTO ES PARA ESTUDIAR CASO POR CASO, AL DIA 25/07/23 ERAN 17 LOTES.
'SELECT * FROM MOVISTO WHERE IdenLote = 'BR-2245-46'
'VER SI NO TIENE QUE VOLVER A LA CONSULTA VERIFICAR LOS .MOVENEXT SI NO DEBERIA VOLVER A EJECUTAR LA PRIMERA CONSULTA.
'_____________________________________________________
End Sub


Function VALIDAR_OCOMP_GENERADA%(NUMERO_PRESUPUESTO&)
    CONDI$ = "NroPres = " & NUMERO_PRESUPUESTO&
    If CantRegistros&("PRESUCOMPRA", CONDI$, "NOMESS") Then
            VALIDAR_OCOMP_GENERADA% = 1
            OPX% = COMALTER%("Este Presupuesto Ya Genero Solicitudes de Compras\Si Necesita Modificar:\Anule la-s Solicitud-es y Vuelva a Generar\\Ver Solicitudes Generadas\Cancelar")
            If OPX% < 1 Or OPX% > 1 Then Exit Function
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
            If ANTOT < 5500 Then ANTOT = 5500
            FRMZELECHO.Width = ANTOT + 1100
            FRMZELECHO.Caption = "Lista de Ordenes de Compra - Presupuesto N°: " & NUMERO_PRESUPUESTO&
            '
            FRMZELECHO.MSF1.Rows = 1
            SQLX$ = "SELECT DISTINCT(NroPres),Num_Prov, NumSolCompra FROM PRESUCOMPRA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE NroPres = " & NUMERO_PRESUPUESTO&
            SQLX$ = SQLX$ + " ORDER BY Num_Prov"
            Set rs = READSET(SQLX$)
            With rs
              Do While Not .EOF
                i& = i& + 1
                FRMZELECHO.msf2.Rows = i& + 1
                NP% = XVALO(!Num_Prov)
                FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
                FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
                   NSOLIC& = XVALO(!NumSolCompra)
                   OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
                FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
                If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
                FRMZELECHO.msf2.TextMatrix(i&, 4) = NSOLIC&
                .MoveNext
              Loop
            End With
            rs.Close
            Set rs = Nothing
            FRMZELECHO.Show 1
            
    End If
    '
End Function

Function VALIDAREPECODI%(CI1%)
   VALIDAREPECODI% = 0
   If EDICION% > 0 Then Exit Function
   If CI1% > 0 Then
      For i% = 1 To MSF1.Rows - 1
         CI2% = CODINT%(MSF1.TextMatrix(i%, 1))
         If CI1% = CI2% Then
            VALIDAREPECODI% = i%
         End If
      Next i%
   End If
   
End Function


Sub VOLVER_A_DOLAR()
    'DOLAR POR DEFECTO
    MONEMI% = 2  ' DOLAR
    MONEMIS = DEMONECO$(2)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
    End If

End Sub

Private Sub AltaContacto_Click()
    COCLI$ = TRIM$(Label26.Caption)
    If COCLI$ <> "" Then
      FRMMAILS.Label26 = COCLI$
      FRMMAILS.Show 1
    Else
      Call COMUNI("Código de Cliente No Válido o Inexistente")
      GoTo 99
    End If
99

End Sub

Private Sub CARGA_Click(Index As Integer)
  CARGA(Index).ForeColor = &HFF&
  CARGA(Index).FontUnderline = True
  For i% = 0 To 2
    If i% <> Index Then
      CARGA(i%).ForeColor = &H80000008
      CARGA(i%).FontUnderline = False
    End If
  Next i%
    
End Sub

Private Sub Combo6_Click()
    Text38.TEXT = Combo6.TEXT
End Sub

Sub Command1_Click()
   Command1.Enabled = False
'   Call SELECHO("MASTER")
   Call ZELECHOAHP(1)
   COD$ = RESP_ZELE_VECT(2)
'   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     txtcodigo = ""
     txtcodigo = COD$
     TXTCODCLIENTE = ""
     Call txtcodigo_KeyPress(13)
   End If
   Command1.Enabled = True
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   NC = XVALO(Text1)
   If NC < 1 Then GoTo 99
   
   CODIXCLI$ = CODICLI$(NC)
   CONDI$ = "CODICLI= '" & CODIXCLI$ & "'  ORDER BY NOME_CTAC "
   Call CARGALISEL("CONTACTO", "CONTACSCOM", "NOME_CTAC/A18;CELU_CTAC/A18;MAIL_CTAC/A48", CONDI$, "NOMESS")

   Call SELECHO("CONTACTO")
   NAMECONTACTO$ = TRIM$(VALACT1$("CONTACTO"))
   If NAMECONTACTO$ = "" Then GoTo 99
   '
   CELUCONTACTO$ = TRIM$(Left(VALACT2$("CONTACTO"), 18))
   MAILCONTACTO$ = TRIM$(Mid$(VALACT2$("CONTACTO"), 19))
   '
   Text35 = NAMECONTACTO$
   Text5 = CELUCONTACTO$
   Text18 = MAILCONTACTO$
   '
99 Command10.Enabled = True
   '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
'   Call CARGA_INDICE_LPRECIO
'   If RESPUESTA_BUSQUEDA$ <> "" Then Text23 = RESPUESTA_BUSQUEDA$
'   Command11.Enabled = True
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   'OJO TAMBIEN SE USA PARA MOSTRAR TEXTO EN MAIL
   OBSDETALLE.Show 1
   If OBSDETALLE.Label1 = "OK" Then
    REG& = MSF1.Rows
    MSF1.Rows = REG& + 1
    MSF1.TextMatrix(REG&, 1) = "------------"
    MSF1.TextMatrix(REG&, 2) = "------------"
    MSF1.TextMatrix(REG&, 3) = OBSDETALLE.Text28
    MSF1.TextMatrix(REG&, 4) = 1
    MSF1.TextMatrix(REG&, 5) = "------------"
    MSF1.TextMatrix(REG&, 6) = "------------"
   End If
   Call Command6_Click
   'Call numerargrilla(FRMPRESUAHP.MSF1, 0)
   Command12.Enabled = True
   '
End Sub

'
Private Sub Command13_Click()
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   Command3.Enabled = False
   Text1.Locked = True
   Command19.Enabled = False
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then
      MsgBox "FALTA NÚMERO DE CLIENTE"
      GoTo 99
   End If
   '
   'VALIDACION DE DETALLE
   If MSF1.Rows <= 1 Then
     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
     MsgBox TTXX$
     GoTo 99
   End If
   
   Call Grabar_Presupuesto(OK%)
   If OK% > 0 Then
     SIGRABO.Caption = "SI"
     OPX% = COMALTER%("Presupuesto Grabado\\Continuar con Este Presupuesto\Nuevo Presupuesto")
     If OPX% < 2 Then Exit Sub
   Else
     Call COMUNI("El presupuesto No se Guardó Correctamente\Vuelva a Intentarlo")
   End If
   '
   'ESTO PIDIO ESTEBAN PARA QEU CUANDO GRABA LIMPIE TODA LA INFORMACION
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
   Call LIMPIARPRESU
   LIMPIAR_TEXT_DETALLE
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   Text21.Width = 10
   If Command9.Visible = True Then Command9.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   Call VOLVER_A_DOLAR
99 Command3.Enabled = True
End Sub


Private Sub Command14_Click()
   On Error Resume Next
   MARCONOTA.Visible = False
   txtcodigo.Locked = False
   TXTCANTIDAD.Locked = False
   On Error GoTo 0
End Sub

Private Sub Command15_Click()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   txtcodigo.Locked = False
   TXTCANTIDAD.Locked = False

   On Error GoTo 0
End Sub


''Sub Command13_Click()
''   Command3.Enabled = False
''   Call RUTINACOMMAND13(0)
''99 Command3.Enabled = True
''End Sub
''Sub RUTINACOMMAND13(UPDATE_AUTOMAT%)
''   Text1.Locked = True
''   Command19.Enabled = False
''   CUENTA& = XVALO(Text1)
''   If CUENTA& < 1 Then
''      MsgBox "FALTA NÚMERO DE CLIENTE"
''      Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
''      Exit Sub
''   End If
''   '
''   'VALIDACION DE DETALLE
''   If MSF1.Rows <= 1 Then
''     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
''     MsgBox TTXX$
''     Exit Sub
''   End If
''
''   Call Grabar_Presupuesto(OK%, UPDATE_AUTOMAT%)
''   If OK% > 0 Then
''     If UPDATE_AUTOMAT% < 1 Then ' SI  UPDATE_AUTO% > 0 ES POR QUE SE LO LLAMA PARA QUE VAYA ACTUALIZANDO AUTOMATICAMENTE
''        Call COMUNI("Presupuesto Grabado")
''     Else
''        Exit Sub
''     End If
''   Else
''     Call COMUNI("El presupuesto No se Guardó Correctamente\Vuelva a Intentarlo")
''   End If
''   '
''   'ESTO PIDIO ESTEBAN PARA QEU CUANDO GRABA LIMPIE TODA LA INFORMACION
''   Text1 = ""
''   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
''   Call LIMPIARPRESU
''   LIMPIAR_TEXT_DETALLE
''   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
''   If Command9.Visible = True Then Command9.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
''   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
''   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
''   Call VOLVER_A_DOLAR
''End Sub
Private Sub Command19_Click()
    If XVALO(Text20) > 0 Then Exit Sub 'si es modificacion
    Command19.Enabled = False
    ''''''''''''''''''
    VENZELECHO.NombreZelechoArmado = "DEUDOR12"
    VENZELECHO.Inicializar = 1
    FRMZELECHO2.Show 1
    NCLIE = XVALO(RESP_ZELE_VECT(1))
    ''''''''''''''''''''''
    If NCLIE > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE))
      'Call CARDACLI
      On Error Resume Next
      txtcodigo.SetFocus
      On Error GoTo 0
    End If

    Screen.MousePointer = 1
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call SELECHO("CONPAG")
   VDEVU$ = VALACT1$("CONPAG")
   If TRIM$(VDEVU$) <> "" Then
      TXTCONDIPAGO = VDEVU$
   End If
   
   
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("VALIOFER")
   VDEVU$ = VALACT1$("VALIOFER")
   If TRIM$(VDEVU$) <> "" Then
      Text29 = VDEVU$
   End If
   Command3.Enabled = True

End Sub

Private Sub Command4_Click()
   If EmpresaNoRealizaCambios% > 0 Then Exit Sub
   EPAC$ = TRIM$(EMPREAC$)
   Command4.Enabled = False
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Call APERBASDAT("CLIEN")

   Dim Tranclie As ADODB.Recordset
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(Control$("", "PRONUCLI"))
   'If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1
   '
   'Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & PRONUCLI&")
   'sql
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & NROPRO&))
   If Not (Tranclie.EOF And Tranclie.BOF) Then 'si la encuentra (si el record no esta vacío)
      GoTo 3
   End If
'   '
   NCLIE = NROPRO&
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(Int(NCLIE)), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 29999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE%))
     Call REPLA(VDEF$, COCLINU$, 1, 12)
     'Call MENSERR(24, "Codigo de Cuenta no Válido")
     'GoTo 10
   End If
   '
  
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Codi_Cli FROM Deudor12 where Codi_cli='" & COCLINU$ & "'"))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 where Nume_cli=" & NCLINUE))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If PIVCLI% = 0 And Len(CUTT$) > 12 Then
        Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
        GoTo 10
   End If
   '
   If CUTT$ <> "" Then
     'sql
     Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Cuit_Cli = '" & CUTT$ & "'"))
     If Not (Tranclie.EOF And Tranclie.BOF) Then
         If Tranclie!nume_cli > 0 Then
            NCLIX% = Tranclie!nume_cli
            If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(rasocli$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
                GoTo 10
            End If
         End If
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenKeyset
   Rdeudor12.LockType = adLockPessimistic
   Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLINUE%), FLEXCONN, , , adCmdText
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !nume_cli = TRIM$(Str$(NCLINUE%))
      !raso_cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !Stat_Cli = 1
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
   End With
   '
   Rdeudor12.Close
   Set Rdeudor12 = Nothing
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
     CONDI$ = "nume_cli=" & NCLINUE%
     Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Deudor12", "MINEPARTS", "Deudor12", CONDI$)
   End If

   UPDATCLI% = 1
   
   Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
   '
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

''Private Sub Command5_Click()
''   Command5.Enabled = False
''   Text1.Locked = True
''   Command19.Enabled = False
''   CUENTA& = XVALO(Text1)
''   If CUENTA& < 1 Then
''      MsgBox "FALTA NÚMERO DE CLIENTE"
''      GoTo 99
''   End If
''   '
''   'VALIDACION DE DETALLE
''   If MSF1.Rows <= 1 Then
''     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
''     MsgBox TTXX$
''     GoTo 99
''   End If
''   '
''   If XVALO(Total(0).TEXT) < 0.05 Then
''      Call COMUNI("Falta(n) Precio(s) de Lista")
''      GoTo 99
''   End If
''   '
''   If XVALO(Text14) < 0.05 Then
''      Call COMUNI("Importe no Válido")
''      GoTo 99
''   End If
''
''   'VALIDO QUE TODOS TENGAN PRECIO DE VENTA.
''   For I% = 1 To MSF1.Rows - 1
''         PRECIOVENTA# = XVALO(MSF1.TextMatrix(I%, 8))
''         If PRECIOVENTA# < 0.005 Then
''            Call COMUNI("Falta Precio de Venta en Fila: " & I%)
''            OPX% = COMALTER("Falta Precio de Venta en Fila: " & I% & "\\Continuar\Corregir")
''            If OPX% = 2 Then GoTo 99
''         End If
''   Next I%
''
''   SENDBYMAIL$ = ""
''   NC = XVALO(Text1)
''   'TOMA EL MAIL DE PANTALLA POR QUE SE PUEDE CAMBIAR
''   SENDBYMAIL$ = TRIM$(Text18) 'EMAILCLI$(NC)
''   'VALIDACION DE MAIL
''35 OPX% = COMALTER%("Envio por Mail\\Imprimir\ Enviar Por Mail\Cancelar Grabación")
''        If OPX% = 2 Then
''            SENDBYMAIL$ = TRIM$(InputBox("Envío de Mail", "Ingreso Mail Destinatario", SENDBYMAIL$))
''             If InStr(SENDBYMAIL$, "@") < 1 Then  'TIENE DIRECCION VALIDA
''                Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
''                GoTo 35
''             End If
''        ElseIf OPX% = 3 Then ' CANCELAR
''             GoTo 99
''        Else
''             SENDBYMAIL$ = ""
''        End If
''
''   Call Grabar_Presupuesto(OK%)
''   If OK% < 1 Then GoTo 99 'SI HUBO ALGUN PROBLEMA EN LA GRABACION VUELVE CON OK = 0
''   '
''   Call PRIPRESU(0)
''   '
''   Text1 = ""
''   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
''   'Text20 = "-00"
''   Call LIMPIARPRESU
''   LIMPIAR_TEXT_DETALLE
''   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
''   If Command9.Visible = True Then Command9.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
''
''   Call VOLVER_A_DOLAR
''
''99 Command5.Enabled = True
''   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
''   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
''
''   '
''End Sub
Private Sub Command5_Click()
   Command5.Enabled = False
   Text1.Locked = True
   Command19.Enabled = False
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then
      MsgBox "FALTA NÚMERO DE CLIENTE"
      GoTo 99
   End If
   '
   'VALIDACION DE DETALLE
   If MSF1.Rows <= 1 Then
     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
     MsgBox TTXX$
     GoTo 99
   End If
   '
   If XVALO(Total(0).TEXT) < 0.05 Then
      Call COMUNI("Falta(n) Precio(s) de Lista")
      GoTo 99
   End If
   '
   If XVALO(Text14) < 0.05 Then
      Call COMUNI("Importe no Válido")
      GoTo 99
   End If

   'VALIDO QUE TODOS TENGAN PRECIO DE VENTA.
   For i% = 1 To MSF1.Rows - 1
         PRECIOVENTA# = XVALO(MSF1.TextMatrix(i%, 8))
         If PRECIOVENTA# < 0.005 Then
            Call COMUNI("Falta Precio de Venta en Fila: " & i%)
            OPX% = COMALTER("Falta Precio de Venta en Fila: " & i% & "\\Continuar\Corregir")
            If OPX% = 2 Then GoTo 99
         End If
   Next i%
   
   SENDBYMAIL$ = ""
   NC = XVALO(Text1)
   'TOMA EL MAIL DE PANTALLA POR QUE SE PUEDE CAMBIAR
   SENDBYMAIL$ = TRIM$(Text18) 'EMAILCLI$(NC)
   'VALIDACION DE MAIL
35 OPX% = COMALTER%("Envio por Mail\\Imprimir\ Enviar Por Mail\Cancelar Grabación")
        If OPX% = 2 Then
            SENDBYMAIL$ = TRIM$(InputBox("Envío de Mail", "Ingreso Mail Destinatario", SENDBYMAIL$))
            Text18 = SENDBYMAIL$
             If InStr(SENDBYMAIL$, "@") < 1 Then  'TIENE DIRECCION VALIDA
                Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
                GoTo 35
             End If
        ElseIf OPX% = 3 Then ' CANCELAR
             GoTo 99
        Else
             SENDBYMAIL$ = ""
        End If
   
   Call Grabar_Presupuesto(OK%)
   If OK% < 1 Then GoTo 99 'SI HUBO ALGUN PROBLEMA EN LA GRABACION VUELVE CON OK = 0
   '
   Call PRIPRESU(0)
   '
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
   'Text20 = "-00"
   Call LIMPIARPRESU
   LIMPIAR_TEXT_DETALLE
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   Text21.Width = 7000 ' LO VUELVO AL TAMAÑO NORMAL
   If Command9.Visible = True Then Command9.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO

   Call VOLVER_A_DOLAR

99 Command5.Enabled = True
   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   
   '
End Sub

Sub Grabar_Presupuesto(OKX%, Optional UPDATE_AUTOMAT%)
    '
   OKX% = 0
   Call BLOQARCH("PRESUPU")
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   '
   HOII% = HOY(HOS$)
   FEMI% = CVINT(DTPicker1.Value)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If

   MODIFICA% = 0
   If Text21.Visible = True Then MODIFICA% = 1
   CONDI$ = "NroPres= " & XVALO(Text17.TEXT)
   If MODIFICA% < 1 Then
      EXISTEPRESU% = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
      If EXISTEPRESU% > 0 Then
       Call COMUNI("Se va a Cambiar el Nro de Presupuesto\ Por que el Mismo Esta Ocupado")
       Text17.TEXT = TRIM$(Str$(NUPRESUP&))
'       Call CIERRARCH("PRESUPU")
'       Exit Sub
    End If
   End If

'   If MODIFICA% < 1 Then
'    NPP& = NUPRESUP&
'    If NPP& > XVALO(Text17.TEXT) Then
'       Call CIERRARCH("PRESUPU")
'       MERRO$ = "Número de Presupuesto Ocupado, Vuelva a Grabar"
'       Text17.TEXT = TRIM$(Str$(NPP&))
'       GoTo 98
'    End If
'   End If
   '
   Text21.Visible = True: Text21.Width = 10
   Call ABREPRESVEN
   '\\\OT-06-0168-WAR-FIN///
   '
   '
   '
   NUOIT% = 0
   
   Call ABREPRESVEN
'   CONDI$ = "NroPres= " & XVALO(Text17.TEXT)
'   STRSQL = "Select * from Presuven"
'   STRSQL = STRSQL + " where " & CONDI$
'   Dim PresuVen As ADODB.Recordset
'   Set PresuVen = New ADODB.Recordset
'   PresuVen.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'   On Error Resume Next
'   With PresuVen
'   If MODIFICA% < 1 Then
'    If Not (.EOF And .BOF) Then
'       Call CIERRARCH("PRESUPU")
'       Call COMUNI("Número Ocupado")
'       Text11.TEXT = TRIM$(Str$(NPP&))
'       Exit Sub
'    End If
'   End If
   '
   Call ACTUREGISTRO("PRESUVEN", "MARBORRA", CONDI$, 1, RegAf&, "NOMESS")
   
   CONDI$ = "NroPres= " & XVALO(Text17.TEXT)
   strSQL = "Select * from Presuven"
   strSQL = strSQL + " where " & CONDI$
   Dim PresuVen As ADODB.Recordset
   Set PresuVen = New ADODB.Recordset
   PresuVen.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With PresuVen
    DESC$ = PORDESCU
    COEFIDESCU = COEFDESCU(DESC$)
    For U& = 1 To MSF1.Rows - 1
         CIIX% = CODINT%(MSF1.TextMatrix(U&, 1))
         CAN = XVALO(MSF1.TextMatrix(U&, 4))
         INETO = XVALO(MSF1.TextMatrix(U&, 8)) * (100 - XVALO(COEFIDESCU)) / (100)

         If CAN > 0.05 Then
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CODIEMPR = CodiEmp%
           On Error GoTo 0
             !NroPres = XVALO(Text17.TEXT)
             '\\\OT-06-0168-WAR-09/05/06///
             !NumeCRubro = 0
             !Cod_Rubro = ""
             !Cod_Anio = ""
             '\\\OT-06-0168-WAR-FIN///
             !FECHEMIS = CVDAT(FEMI%)
             !NroClie = NCLIE%
             !RasoClie = Left$(Text2, 64)
             !Referen = Left$(TXTREFERENCIA, 16)
             !AteSen = Left$(Text35, 48) 'el contactdo de pantalla
             !LisPre = XVALO(Text23)
             !ConPag = XVALO(Text26)
             !VENDED = 0
             !MonedEmi = MONEMI%
               MON% = MONEMI% '
             'aaaa = TICAMONE(MON%)
               If MONEMI% = 1 Then MON% = 2 'SI ES PESOS GUARDA LA COTIZACION EN DOLARES, SI NO LA QUE CORRESPONDE
               'bbb = FORMATNUM$(XVALO(TICAMONE(MON%)), "F12.4")
             !COTIZAMONEDA = TICAMONE(MON%)
             !NuOrdIt = NUOIT%
               NUOIT% = NUOIT% + 1
             !CodiInt = CIIX%
             !CodiExt = Left$(CODEXT$(CIIX%), 24)
             !Descrip = Left$(DESCRIT0$(CIIX%), 64)
             !Especif = MSF1.TextMatrix(U&, 3) 'AHP GRABAR
             !PreList = INETO
             !TDESCUEN = Left$(PORDESCU, 32)
                DESC$ = PORDESCU
                TTT1 = COEFDESCU(DESC$)
             !PorcDesc = TTT1
             !ImpoDesc = INETO * TTT1 / 100
             !PreUnit = INETO * (1 - TTT1 / 100)
             !Cantit = CAN
             !ImToNe = CAN * INETO * (1 - TTT1 / 100)
             !FechVent = CVDAT(0)
             !NroPedi = 0
             !NroFac_Cor = ""
             !NroFac_Lar = ""
             !Codi_Cli = CODICL$
             !Vers_Presu = 0
             !Num_Prov = XVALO(MSF1.TextMatrix(U&, 5))
             !COSTO_UNIT = Val(MSF1.TextMatrix(U&, 7))
             !Via_Entrega = XVALO(Left$(MSF1.TextMatrix(U&, 10), 2))
             !LOTES_Reservados = MSF1.TextMatrix(U&, 11)
             !Canti_Lotes = XVALO(MSF1.TextMatrix(U&, 12))
             !CodigoDelCliente = MSF1.TextMatrix(U&, 2)
             !Mail_Cli = Text18
             !MARBORRA = 0
             !COSTOFOB = XVALO(MSF1.TextMatrix(U&, 16))
             .Update
         End If
        Next U&
   End With
   PresuVen.Close
   Set PresuVen = Nothing
   'ESTO LO COMENTE POR QUE SE GRABA EN INFOSIAHP EN RUTINA ACTUALIZAPESO_ENMAESTRO EN INGRESO DE STOCK Y SIN STOCK
   
'   'RECORRE LA GRILLA PARA ACTUALIZAR EL PESO EN EL CASO QUE TENGA VALOR > 0
'   If UPDATE_AUTOMAT% < 1 Then
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    For U& = 1 To MSF1.Rows - 1
'       CI1% = CODINT%(MSF1.TextMatrix(U&, 1))
'       PESOX = XVALO(MSF1.TextMatrix(U&, 14))
'       If PESOX > 0.005 Then
'         Call ACTUREGISTRO("MASTER", "PESUNI", "CODINT= " & CI1%, FORMATNUM$(PESOX, "F16.5"), REGAF&, "NOMESS")
'         If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
'          Condix$ = "CODINT = " & CI1%
'          Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", Condix$)
'         End If
'
'       End If
'    Next U&
'   End If
   'FIN ACTUALIZACION DE PESO UNITARIO DE MASTER
   '
   CONDI$ = CONDI$ & " AND MARBORRA = 1"
   Call BORRAREGISTROS("PRESUVEN", CONDI$, RegAf&, "NOMESS")
   MODIFICA% = 0
   Call OPLIPRE07.ACTINDIPRES(XVALO(Text17))
   Call CIERRARCH("PRESUPU")
   '
   '
   OKX% = 1 + UPDATE_AUTOMAT%
   Exit Sub
   '
98 Call CIERRARCH("PRESUPU")
   Call CIERRARCH("INDIPRES")
   Call MENSERR(24, MERRO$)
   '
End Sub


Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         VCC% = 1
       End If
     End If
   End If
   If VCC% = 0 Then
      CODIXX$ = TRIM$(Label26)
      CODICL$ = SAFETEXT$(VALOBADA("BSEGCOM", "Codi_Cli", "Codi_Cli ='" & CODIXX$ & "'"))
      If CODICL$ <> "" Then VCC% = 1
   End If
   '
99 VALICLIE% = VCC%
   '
End Function


Private Sub Command6_Click()

  Command6.Enabled = False
  Screen.MousePointer = 11
  Call COLOREARSTOCKS(i&)
'  For I& = 1 To MSF1.Rows - 1
'       Call COLOREARSTOCKS(I&)
'  Next I&
'  Call numerargrilla(FRMPRESUAHP.MSF1, 0)
  Call POSICIONAR_ULTIMAFILA(FRMPRESUAHP.MSF1)
  Command6.Enabled = True
  Screen.MousePointer = 1
  
End Sub

Sub COLOREARSTOCKS(REGI&)
        
    '   TOMAR EL DISPONIBLE DEL LOTE E IR DESCONTANDO. (VER LA VALIDACION YA REALIZADA)
    '   LOS CASOS QUE TIENEN PARTIDA SELECCIONADA SI NO HAY DE LOS QUE ESTAN GRABADOS O NO ALCANZAN VAN EN ROJO.
    '   LOS CASOS QUE NO TIENEN PARTIDA SELECCIONADA SE VALIDA STOCK EN GENERAL SI NO POSEE SE PONEN EN ROJO, SI POSEE STOCK DISPONIBLE TOTAL O PARCIAL VA EN AMARILLO , PARA QUE SE ELIJA
    '   LOS CASOS QUE NO TIENEN PARTIDA SELECCIONADA SE VALIDA
    
    '****************************************************************************
     'PASO A UN GRID OCULTO PARTE DE LA INFORMACION PARA LUEGO ORDENARLA POR LOTE
     msf2.Cols = 6
     Call numerargrilla(FRMPRESUAHP.MSF1, 0)
     For i& = 1 To MSF1.Rows - 1
       msf2.Rows = i& + 1
       Call COLOREAR_GRILLA_SOLAFILA(FRMPRESUAHP.MSF1, &HFFFF&, REGI&, 0, 1, 1)
       msf2.TextMatrix(i&, 0) = XVALO(MSF1.TextMatrix(i&, 0)) ' ORDEN
       msf2.TextMatrix(i&, 1) = CODINT%(MSF1.TextMatrix(i&, 1)) 'CODIGO INTERNO
       msf2.TextMatrix(i&, 2) = XVALO(MSF1.TextMatrix(i&, 4))  'CANTIDAD
       msf2.TextMatrix(i&, 3) = XVALO(MSF1.TextMatrix(i&, 5)) 'NPROV
       msf2.TextMatrix(i&, 4) = MSF1.TextMatrix(i&, 11)       'LOTE
     Next i&
     '*************************************************************************
     'ORDENAR POR LOTE
     msf2.COL = 4
     msf2.Sort = 7
     '*************************************************************************
     
     'RECORRO GRILLA ORDENADA POR LOTE
     For i& = 1 To msf2.Rows - 1
       NUORD = XVALO(msf2.TextMatrix(i&, 0))
       LOTE$ = TRIM$(msf2.TextMatrix(i&, 4))
       CI1% = XVALO(msf2.TextMatrix(i&, 1))
       CANTI = XVALO(msf2.TextMatrix(i&, 2))
       If TRIM$(LOTE$) <> "" Then 'SI TIENE LOTES
          If LOTE_ANT$ = LOTE$ Then
             ACUCANTI = ACUCANTI + CANTI
          Else
             ACUCANTI = CANTI
             ST_RESERVADO = SALDO_LOTE_PEDIDO(CI1%, LOTE$, 0)
             SLOTE = SaldoLote(CI1%, LOTE$, 0)
             '
             CANT_DISPONIBLE = SLOTE - ST_RESERVADO
             If CANT_DISPONIBLE < 0 Then CANT_DISPONIBLE = 0
          End If
          '
          If CANT_DISPONIBLE < ACUCANTI Then
             msf2.TextMatrix(i&, 5) = "R"
          Else
             msf2.TextMatrix(i&, 5) = "V"
          End If
          LOTE_ANT$ = LOTE$
       Else
          ST = FORMATNUM$(STOCKACT(CI1%), "F8.0")  ' STOCK
          CANPED = CANT_PEDIDA_PENDIENTE(CI1%)     ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
          COMPRAX = COMPRA_PEND_DE_ENTR(0, CI1%)   ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
          CANRECEP1 = CANT_RECEP_PENDIENTE_APROB(0, CI1%)
          CantidadCargadaEnLosRegistrosAnteriores = ITEM_CARGADO_ENGRILLA(CI1%, FRMPRESUAHP.MSF1, i&, i&)
          CANTI = CANTI + CantidadCargadaEnLosRegistrosAnteriores 'ACUMULA LA CANTIDAD DEL REGISTRO + LA DE LOS REGISTROS ANTERIORES
          COLOXX$ = COLORSTATUS$(CANTI, ST, COMPRAX + CANRECEP1, CANPED)
          msf2.TextMatrix(i&, 5) = COLOXX$
       End If
       COLO$ = msf2.TextMatrix(i&, 5)
       MSF1.row = NUORD
       MSF1.COL = 0

       If COLO$ = "V" Then
           MSF1.CellBackColor = vbGreen
'           Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbGreen, NUORD, 0)
           FRMPRESUAHP.MSF1.TextMatrix(NUORD, 15) = "Inmediata" 'DEBERIA SER SIEMPRE INMEDIATA POR QUE AQUI ES CON STOCKS.
       End If
       If COLO$ = "A" Then MSF1.CellBackColor = vbYellow ' Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbYellow, NUORD, 0)
       If COLO$ = "R" Then MSF1.CellBackColor = vbRed '  Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbRed, NUORD, 0)
       Call COLOREAR_GRILLA_SOLAFILA(FRMPRESUAHP.MSF1, &HFEE7D6, NUORD, 0, 1, 1)
     Next i&
     '
'    With MSF1
'        .Row = 1
'        .COL = 0
'        .CellBackColor = vbGreen
''        .ColSel = .Cols - 1
'    End With
     MSF1.Redraw = False
     MSF1.COL = 0
     MSF1.Sort = 3 ' ordena de manera numerica
     
     On Error Resume Next
     TPXI& = MSF1.Rows - 11: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     MSF1.TopRow = TPXI&
     On Error GoTo 0
      
99   Screen.MousePointer = 1
     MSF1.Redraw = True
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "ACTUALIZA EL SEMAFORO SEGUN CASOS" & vbCrLf
       TXTO$ = TXTO$ + "COLOR VERDE:  " + vbCrLf
       TXTO$ = TXTO$ + "  a)-CARGADO CON STOCK O SE MANTIENE SITUACION DE LOTES ELEGIDOS."
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf

       TXTO$ = TXTO$ + "COLOR AMARILLO:  " + vbCrLf
       TXTO$ = TXTO$ + "  a)-CASO QUE SE CARGO EL PRESUPUESTO ORIGINAL SIN STOCK, Y AL MOMENTO"
       TXTO$ = TXTO$ + "  -DE REFRESCAR HAY STOCK DISPONIBLE PARA SELECCIONAR."
       TXTO$ = TXTO$ + "  b)-CASO QUE SE CARGO EL PRESUPUESTO ORIGINAL SIN STOCK, HAY CANTIDAD EN"
       TXTO$ = TXTO$ + "  -TRANSITO TOTAL O PARCIAL."
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       
       TXTO$ = TXTO$ + "COLOR ROJO:  " + vbCrLf
       TXTO$ = TXTO$ + "  a)-CASO QUE SE CARGO EL PRESUPUESTO ORIGINAL SIN STOCK Y ASI SE MANTIENE."
       TXTO$ = TXTO$ + "  b)-CASO QUE LAS PARTIDAS SELECCIONADAS, NO ALCANCEN PARA CUMPLIR CON LA CANTIDAD SOLICITADA."
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub Command7_Click()
   Call txtcodigo_Change
'  PROVEEDOR , VIA, DEMORA

End Sub

Private Sub Command8_Click()

   Command8.Enabled = False
   Call SELECHO("VENDEDOR")
   VDEVU$ = VALACT1$("VENDEDOR")
   If TRIM$(VDEVU$) <> "" Then
      TXTNUMVENDEDOR = VDEVU$
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call Command6_Click
   NPRES& = XVALO(Text17)
   If NPRES& < 1 Then GoTo 99
   '
   'VALIDO QUE TODOS TENGAN PROVEEDOR
   FALTA_PROVEEDOR% = 0
   For i& = 1 To MSF1.Rows - 1
      NPROV& = XVALO(MSF1.TextMatrix(i&, 5))
      If NPROV& < 1 Then
         FALTA_PROVEEDOR% = 1
         FILAS_SIN_PROV$ = FILAS_SIN_PROV$ & i& & "; "
      End If
   Next i&
   '
   Call GENERA_TABLA_COMPRAS_DE_PRESUPUESTO
   If FALTA_PROVEEDOR% > 0 Then
       POS1% = InStrRev(FILAS_SIN_PROV$, ";")
       If POS1% > 0 Then
          FILAS_SIN_PROV$ = Left$(FILAS_SIN_PROV$, Len(FILAS_SIN_PROV$) - POS1%)
       End If
      Call COMUNI("Falta Proveedor Para la/s Fila/s " & FILAS_SIN_PROV$)
      GoTo 99
   End If
   '
   'ORDENO POR LA FILA DE PROVEEDOR
    MSF1.COL = 5 'COLUMNA PROVEEDOR
    MSF1.Sort = 7
    msf2.COL = 3
    msf2.Sort = 7
    '
    OC_GENERADA% = VALIDAR_OCOMP_GENERADA%(NPRES&)
    If OC_GENERADA% > 0 Then
       GoTo 99
    End If
    '
    NORD% = 0: NPROV_ANT& = 0: VUELTA% = 0
    SQLX$ = "SELECT * FROM PRESUCOMPRA "
    SQLX$ = SQLX$ + " WHERE NroPres = " & NPRES&
    Set RS2 = New ADODB.Recordset
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS2
        For i& = 1 To MSF1.Rows - 1
             MSF1.COL = 0
             NPROV& = XVALO(MSF1.TextMatrix(i&, 5))
             If msf2.TextMatrix(i&, 5) = "R" Then
                If NPROV& <> NPROV_ANT& Then
                   NORD% = 0
                   NUMSOLCOMP& = ProxNumSolCompra&
                End If
                CI1% = CODINT%(MSF1.TextMatrix(i&, 1))
                CANTI = XVALO(MSF1.TextMatrix(i&, 4))
                COSTO1# = XVALO(MSF1.TextMatrix(i&, 7))
                NCLI = XVALO(Text1)
                NORD% = NORD% + 1
                .AddNew
                !NroPres = NPRES&
                !nume_cli = NCLI
                !Cantidad = CANTI
                  FEC1 = CVDAT(HOY(HO$))
                !FECHEMISI = FEC1
                !Num_Prov = NPROV&
                !CODINT = CI1%
                !MARBORRA = 0
                !Status = 0
                !ORDEN = NORD%
                !COSTO = COSTO1#
                !NumSolCompra = NUMSOLCOMP&
                !mone_cos = MONEMI%
                .Update
    '10          RNDX1& = XVALO(RANDSTRING$(-8))
    '            If CantRegistros("PRESUCOMPRA", "NroPres = " & RNDX1&, "NOMESS") > 0 Then GoTo 10
    '            CONDI$ = "NroPres = " & RNDX1&
    '            FEC1 = FETEXCOR1$(HOY(HO$))
    '            Call CreaRegistro("PRESUCOMPRA", "NroPres", RNDX1&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "Nume_Cli", CONDI$, NCLI, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "Cantidad", CONDI$, CANTI, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "FECHEMISI", CONDI$, FEC1, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "Num_Prov", CONDI$, NPROV&, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "Codint", CONDI$, CI1%, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "ORDEN", CONDI$, NORD%, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "COSTO", CONDI$, FORMATNUM$(COSTO1#, "F16.5"), REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "NumSolCompra", CONDI$, NUMSOLCOMP&, REGAF&, "NOMESS")
    '            Call ACTUREGISTRO("PRESUCOMPRA", "Mone_Cos", CONDI$, MONEMI%, REGAF&, "NOMESS")
'
'                '
'                Call ACTUREGISTRO("PRESUCOMPRA", "NroPres", CONDI$, NPRES&, REGAF&, "NOMESS")
                VUELTA% = 1
             End If
             '
          NPROV_ANT& = NPROV&
       Next i&
   End With
   '
   If VUELTA% < 1 Then
      Call COMUNI("No se Detectaron Items Sin Stock")
      GoTo 99
   End If
   
   'CARGAR LA GRILLA MOSTRANDO LOS ITEMS CARGADO
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Solicitudes de Compra Generadas a Partir de Presupuesto N°: " & NPRES&
    '
    i& = 0
    FRMZELECHO.msf2.Rows = 1
    SQLX$ = "SELECT DISTINCT(NroPres),Num_Prov FROM PRESUCOMPRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPres = " & NPRES&
    SQLX$ = SQLX$ + " ORDER BY Num_Prov"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        NP% = XVALO(!Num_Prov)
        FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
        FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    FRMZELECHO.Show 1
    
    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
    '
    
    If NPROV& > 0 Then
       Call COMUNI("Para Generar Ordenes de Compra Deberá Realizarlas desde el Menú > Control de Generación de Orden de Compra")
'       APLI$ = "CPRE" & NPROV& & ";" & NPRES&
'       Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra x Presupuesto")
    End If
    '
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

99 Command9.Enabled = True
    MSF1.COL = 0 'COLUMNA PROVEEDOR
    MSF1.Sort = 3

End Sub

Private Sub desdecodigooriginal_Click()
   
10  CONDI$ = "CODINT > 0"
    If CantRegistros&("CODICLIENTE", CONDI$, "NOMESS") < 1 Then Call COMUNI("Sin Códigos Relacionados"): Exit Sub
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Código Orig./A20;Descripción/A32;Código Cliente/A20;Orden/F0;CODINT/F0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Código Orig./A20;Descripción/A32;Código Cliente/A20;Orden/F0;CODINT/F0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Códigos Relacionados"
    'ESTILO DE COLUMNA MANTIENE SIN SEPARACION SI REPITE EL NUMERO DE CLIENTE.
    FRMZELECHO.MSF1.MergeCells = flexMergeRestrictColumns
    FRMZELECHO.msf2.MergeCells = flexMergeRestrictColumns
    FRMZELECHO.MSF1.MergeCol(1) = True
    FRMZELECHO.msf2.MergeCol(1) = True
    '
    FRMZELECHO.MSF1.Rows = 1
    FRMZELECHO.MSF1.Redraw = False
    i& = 0
'    SQLX$ = "SELECT CODINT,Codigo_Cliente FROM CODICLIENTE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE " & CONDI$
'    SQLX$ = SQLX$ + " ORDER BY CODINT"
    SQLX$ = "SELECT C.CODINT,C.Codigo_Cliente,M.DESCRIPCION,M.CODIGO "
    SQLX$ = SQLX$ + " FROM CODICLIENTE C WITH(NOLOCK) INNER JOIN MASTER M"
    SQLX$ = SQLX$ + " ON C.CODINT = M.CODINT"
    SQLX$ = SQLX$ + " WHERE C.CODINT > 0 "
    SQLX$ = SQLX$ + " ORDER BY CODINT"

    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF1.Rows = i& + 1
        CI1% = XVALO(!CODINT)
        '
        FRMZELECHO.msf2.Rows = i& + 1
        '
        FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!Codigo)
        FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!Codigo_Cliente)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = i&
        FRMZELECHO.msf2.TextMatrix(i&, 5) = CI1%

        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    FRMZELECHO.MSF1.Redraw = False
    FRMZELECHO.Show 1
    '
    CI1% = XVALO(RESP_ZELE_VECT(5))
    CODCLIEN$ = TRIM$(RESP_ZELE_VECT(3))
    ORDEX% = XVALO(RESP_ZELE_VECT(4))
    '
    If CI1% > 0 Then
        OPX% = COMALTER%("Realmente Desea Elimnar el Código Relacionado: " & CODCLIEN$ & "\\Cancelar\Eliminar")
        If OPX% = 2 Then
            CONDI$ = "CODIGO_CLIENTE = '" & CODCLIEN$ & "'  AND CODINT = " & CI1%
            Call BORRAREGISTROS("CODICLIENTE", CONDI$, RegAf&, "NOMESS")
            Call COMUNI("Código Eliminado Correctamente")
        End If
        Unload FRMZELECHO
        GoTo 10
    End If
    '
End Sub

Private Sub eliminaitemssinmovimientos_Click()
   Call SELECHO("MASTER")
   If TRIM$(VALACT1$("MASTER")) = "" Then Exit Sub
   
   COD$ = TRIM$(VALACT1$("MASTER"))
   CI1% = CODINT%(COD$)
   If CI1% > 0 Then
      CONDI$ = "COD_INTE = " & CI1%
      If CantRegistros&("MOVISTO", CONDI$, "NOMESS") > 0 Then
         Call COMUNI("Imposible Eliminar Este Código Posee Movmientos de Stocks Registrados")
         Exit Sub
      End If
      CONDI$ = "CodiInt = " & CI1%
      If CantRegistros&("PEDDETA", CONDI$, "NOMESS") > 0 Then
         Call COMUNI("Imposible Eliminar, Este Código Posee Pedidos Registrados")
         Exit Sub
      End If
      '
      CONDI$ = "CodiInt = " & CI1%
      If CantRegistros&("PRESUVEN", CONDI$, "NOMESS") > 0 Then
         Call COMUNI("Imposible Eliminar, Este Código Posee Presupuestos Registrados")
         Exit Sub
      End If
      '
      CONDI$ = "CODINT = " & CI1%
      If CantRegistros&("CODICLIENTE", CONDI$, "NOMESS") > 0 Then
         Call COMUNI("Imposible Eliminar, Este Código Posee Código/s de Cliente Relacionado")
         Exit Sub
      End If
      
      CONDI$ = "Codint = " & CI1%
      Call BORRAREGISTROS("MASTER", CONDI$, RegAf&, "NOMESS")
      Call COMUNI("Item Eliminado")
   End If

End Sub

Private Sub Form_Activate()
' If INFOSIAHP.PRESENTA_COTIZADOR% > 0 Then
'    TXTCANTIDAD.SetFocus
' End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 113 'pulso  f2
        Call Command1_Click
        Case 116 'pulso  f5
        Command5_Click
    End Select
    '
End Sub

Private Sub mnuconfig_Click()
   Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnucrt_Click()
   Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuctoc_Click()
    CONDI$ = "NroPres > 0   ORDER BY NroPres DESC"
10  Call CARGALISEL("INDICEP1", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEP1/INDICE DE PRESUPUESTOS")
    NumePresu& = XVALO(VALACT1("INDICEP1"))
    If NumePresu& < 1 Then Exit Sub
    '
    Call GENERA_TABLA_COMPRAS_DE_PRESUPUESTO
    '
    'ARMO LA PRESENTACION DE LA GRILLA EN EL FRMZELECHO
50  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Generación de Ordenes de Compra - Presupuesto N°: " & NPRES&
    '
    'PRESENTO TODAS LAS SOLICITUD DE Compras RELACIONADAS CON EL PRESUPUESTO
  i& = 0
    SQLX$ = "SELECT DISTINCT(Num_Prov),NumSolCompra FROM PRESUCOMPRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPres = " & NumePresu&
    SQLX$ = SQLX$ + " ORDER BY Num_Prov,NumSolCompra"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        NP% = XVALO(!Num_Prov)
        FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
        FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
           NSOLIC& = XVALO(!NumSolCompra)
           OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
        If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = NSOLIC&
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'SI NO TIENE SOLICITUD AVISA Y VUELVE A LISTA DE PRESUPUESTOS
    If i& < 1 Then
       Call COMUNI("No se Detectaron Solicitud de Compras Para El Presupuesto  Seleccionado")
       GoTo 10
    End If
    FRMZELECHO.Show 1
    '
    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
    OCOMPRA& = XVALO(RESP_ZELE_VECT(3))
    '
    'SI TIENE ORDEN DE COMPRA AVISA Y VUELVE A LA LISTA DE SELECCION
    If OCOMPRA& > 0 Then
       Call COMUNI("Solicitud Seleccionada Generó La Orden de Compra: " & OCOMPRA& & "\Modifique Orden de Compra o Anular Para Generar una Nueva.")
       GoTo 50
    End If
    '
    If NPROV& > 0 Then
       APLI$ = "CPRE" & NPROV& & "_" & NumePresu&
       Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra x Presupuesto")
    Else
       GoTo 10
    End If
    '
End Sub

Private Sub mnuestacotiVSventas_ORIGINAL_Click()
    Screen.MousePointer = 11
20  Campos$ = "CI/F0;Codigo/A20;Descripción/A32;Cotizaciones/F12;T.Un./F4;Facturas/f12;T.Unidades/f12;Stock/F8;A Recibir/F9"
    Call CARGA_ENCABEZADO(REPESTAHP.GRID, Campos$, REPESTAHP, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(REPESTAHP.MSF_2, Campos$, REPESTAHP, 1, 0, ANTOT, , 1)
    
    Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    
    
    OBX$ = OBJPLA$("GCON-FAMI", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    '
    FAM$ = TRIM$(Mid$(OBX$, 1, 2))
    GCONT% = CVI(Mid$(OBX$, 3, 2))
    '
    REPESTAHP.GRID.Redraw = False
    
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)
    CantiDias = DIENTRE%(DESDEX%, HASTAX%)
    COEF = CantiDias / 30
    If COEF < 1 Then COEF = 1
    'CARGA PRIMERO LA INFO DEL PRESUPUESTO
'select count(*) AS VECES ,CodiInt, CodiExt, Descrip, sum(Cantit) as canticotizadas,M.FAMILIA, M.GRUCOVEN from presuven p with(NOLOCK)
'inner join master M ON M.CODINT = P.CODIINT
' where FechEmis >= '01/NOV/2017' and FechEmis  <= '12/DIC/2018' and codiint > 0 AND M.FAMILIA = 6 AND GRUCOVEN = 0
'group by codiint, CodiExt, Descrip,M.FAMILIA, M.GRUCOVEN order by Descrip
    Call VectArticulos.cargarvector
30  SQLX$ = "select count(*) AS VECES ,CodiInt, CodiExt, sum(Cantit) as canticotizadas ,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + "  from presuven p with(NOLOCK) inner join master M ON M.CODINT = P.CODIINT "
    SQLX$ = SQLX$ + " where FechEmis >= '" & DESD & "'"
    SQLX$ = SQLX$ + " and FechEmis  <= '" & HAST & "'"
    SQLX$ = SQLX$ + " and codiint > 0"
    If FAM$ <> "" Then SQLX$ = SQLX$ + " AND M.FAMILIA =  '" & FAM$ & "'"
    If GCONT% > 0 Then SQLX$ = SQLX$ + " AND GRUCOVEN = " & GCONT%
    SQLX$ = SQLX$ + " group by codiint, CodiExt,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + " order by CodiExt"
    Set RS1 = READSET(SQLX$)
    J& = 0
    With RS1
      Do While Not .EOF
        CI = XVALO(!CodiInt)
        Codigo$ = SAFETEXT$(!CodiExt)
        DESCRI$ = VectArticulos.Descripcion(CI)
        Veces = XVALO(!Veces)
        CANTICOTI = XVALO(!canticotizadas)
        '
        J& = J& + 1
        REPESTAHP.GRID.Rows = J& + 1
        REPESTAHP.GRID.TextMatrix(J&, 1) = CI
        REPESTAHP.GRID.TextMatrix(J&, 2) = Codigo$
        REPESTAHP.GRID.TextMatrix(J&, 3) = DESCRI$
        REPESTAHP.GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(Veces, "F10"))
        REPESTAHP.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(CANTICOTI, "F10"))
        .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    If J& > 0 Then
       'AHORA LEE LA GRILLA QUE YA CARGO, Y LE AGREGA LOS DATOS DEL MOVISTO NO LO HAGO ARRIBA POR UNA CUESTION DE VELOCIDAD.
       For i& = 1 To REPESTAHP.GRID.Rows - 1
           CI = REPESTAHP.GRID.TextMatrix(i&, 1)
           Call VECESCANTFACTURADA(CI, DESD, HAST, RET_VECES, RET_CANTIFACTU)
           REPESTAHP.GRID.TextMatrix(i&, 6) = TRIM$(FORMATNUM$(RET_VECES, "F10"))
           REPESTAHP.GRID.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(RET_CANTIFACTU, "F10"))
       Next i&
       '
       For i& = 1 To REPESTAHP.GRID.Rows - 1
           CI = REPESTAHP.GRID.TextMatrix(i&, 1)
           REPESTAHP.GRID.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(STOCKACT(CI), "F10"))
       Next i&
       '
       For i& = 1 To REPESTAHP.GRID.Rows - 1
           CI = REPESTAHP.GRID.TextMatrix(i&, 1)
           CANTIPENREC = CANRECEP(CI): If CANTIPENREC < 0 Then CANTIPENREC = 0
           CANTIPENCOMP = STOCOPEN(CI): If CANTIPENCOMP < 0 Then CANTIPENREC = 0
           ARECIBIR = CANTIPENCOMP + CANTIPENREC
           REPESTAHP.GRID.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(ARECIBIR, "F10")) ' PENDIENTE A INGRESAR POR COMPRAS O RECPCIONES PENDIENTES DE APROBASR
       Next i&
       '

       Call COLOREAR_GRILLASQL(REPESTAHP.GRID, &HE0E0E0)
       REPESTAHP.GRID.Redraw = True
       REPESTAHP.GRID.Visible = True
       REPESTAHP.GRID.ZOrder 0
       REPESTAHP.Caption = "Estadística de Cotizaciones y Ventas - Período Seleccionado: " & PERIO$
       
       Screen.MousePointer = 1
       REPESTAHP.Show 1
       GoTo 20
    Else
       Call COMUNI("No se Han Encontrado Artículos Según Filtros")
       Screen.MousePointer = 1
       GoTo 20
    End If
99
End Sub

Private Sub mnuEfectividadCotizcionesVSVentas_Click()
    Screen.MousePointer = 11
    NMESES& = 0
    
20  Campos$ = "CI/F0;Codigo/A20;Descripción/A32;Cotizaciones/F12;T.U.Cot./F9;Facturas/f12;T.U.Fact./f9;Stock/F8;A Recibir/F9;Costo/F10.2;%/F8.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    '
   Static ABUSQUE$
25  ABUSQUE$ = XVALO(InputBox$("Cantidad de Meses", "Ingrese Cantidad de Meses a Considerar", ABUSQUE$))
    NMESES& = XVALO(ABUSQUE$)
    If NMESES& < 1 Then
       Call COMUNI("Dato Ingresado Nó Válido")
       GoTo 99
    End If
    '
    Static ABUSQUE2$
28  ABUSQUE2$ = XVALO(InputBox$("Porcentaje de Efectividad", "Ingrese Porcentaje", ABUSQUE2$))
    PORCENTAJE# = XVALO(ABUSQUE2$)
    If PORCENTAJE# < 1 Then
       Call COMUNI("Dato Ingresado Nó Válido")
       GoTo 28
    End If
    
    ' Obtener la fecha actual
    fechaActual = Date
    
    ' Calcular la fecha "desde" restando la cantidad de meses
    fechaDesde = DateAdd("m", -1 * NMESES&, fechaActual)
    fechaHasta = DateAdd("d", -1, DateAdd("m", NMESES&, fechaDesde))
    DESDEX% = CVINT(fechaDesde)
    HASTAX% = CVINT(fechaHasta)
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)


'    Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
'    If PERIO$ = "" Then GoTo 99
    
    PERIO$ = " En Los Últimos " & NMESES& & " Meses (" & fechaDesde & " al " & fechaHasta & ")"
    TITU$ = " Efectividad de Ventas Mayor al " & PORCENTAJE# & " %"
    TITU$ = TITU$ & PERIO$
    
    OBX$ = OBJPLA$("GCON-FAMI", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    '
    FAM$ = TRIM$(Mid$(OBX$, 1, 2))
    GCONT% = CVI(Mid$(OBX$, 3, 2))
    '
    GRILLAGRAL.GRID.Redraw = False
    
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)
    CantiDias = DIENTRE%(DESDEX%, HASTAX%)
    COEF = CantiDias / 30
    If COEF < 1 Then COEF = 1
    'CARGA PRIMERO LA INFO DEL PRESUPUESTO
'select count(*) AS VECES ,CodiInt, CodiExt, Descrip, sum(Cantit) as canticotizadas,M.FAMILIA, M.GRUCOVEN from presuven p with(NOLOCK)
'inner join master M ON M.CODINT = P.CODIINT
' where FechEmis >= '01/NOV/2017' and FechEmis  <= '12/DIC/2018' and codiint > 0 AND M.FAMILIA = 6 AND GRUCOVEN = 0
'group by codiint, CodiExt, Descrip,M.FAMILIA, M.GRUCOVEN order by Descrip
    Call VectArticulos.cargarvector
30  SQLX$ = "select count(*) AS VECES ,CodiInt, CodiExt, sum(Cantit) as canticotizadas ,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + "  from presuven p with(NOLOCK) inner join master M ON M.CODINT = P.CODIINT "
    SQLX$ = SQLX$ + " where FechEmis >= '" & DESD & "'"
    SQLX$ = SQLX$ + " and FechEmis  <= '" & HAST & "'"
    SQLX$ = SQLX$ + " and codiint > 0"
    If FAM$ <> "" Then SQLX$ = SQLX$ + " AND M.FAMILIA =  '" & FAM$ & "'"
    If GCONT% > 0 Then SQLX$ = SQLX$ + " AND GRUCOVEN = " & GCONT%
    SQLX$ = SQLX$ + " group by codiint, CodiExt,M.FAMILIA, M.GRUCOVEN "
'    SQLX$ = SQLX$ + " HAVING count(*) > " & nVeces&
    SQLX$ = SQLX$ + " order by CodiExt"
    Set RS1 = READSET(SQLX$)
    J& = 0
    With RS1
      Do While Not .EOF
        CI = XVALO(!CodiInt)
        Codigo$ = SAFETEXT$(!CodiExt)
        DESCRI$ = VectArticulos.Descripcion(CI)
        Veces = XVALO(!Veces)
        CANTICOTI = XVALO(!canticotizadas)
        '
        Call VECESCANTFACTURADA(CI, DESD, HAST, RET_VECES, RET_CANTIFACTU)
        If RET_CANTIFACTU > 0 Then 'si no se facturo nada next
          PorcentajeCalculado# = (RET_CANTIFACTU / CANTICOTI) * 100
          If PORCENTAJE# > PorcentajeCalculado# Then GoTo 35  ' o si el porcentaje ingresado es mayor al calculado tambien next , si no sigue e imprime en la grilla
        Else
           GoTo 35
        End If
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESCRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(Veces, "F10"))
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(CANTICOTI, "F10"))
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = TRIM$(FORMATNUM$(PorcentajeCalculado#, "F10.2")) ' Porcentaje
35      .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    If J& > 0 Then
       'AHORA LEE LA GRILLA QUE YA CARGO, Y LE AGREGA LOS DATOS DEL MOVISTO NO LO HAGO ARRIBA POR UNA CUESTION DE VELOCIDAD.
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           Call VECESCANTFACTURADA(CI, DESD, HAST, RET_VECES, RET_CANTIFACTU)
           GRILLAGRAL.GRID.TextMatrix(i&, 6) = TRIM$(FORMATNUM$(RET_VECES, "F10"))
           GRILLAGRAL.GRID.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(RET_CANTIFACTU, "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           GRILLAGRAL.GRID.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(STOCKACT(CI), "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           CANTIPENREC = CANRECEP(CI): If CANTIPENREC < 0 Then CANTIPENREC = 0
           CANTIPENCOMP = STOCOPEN(CI): If CANTIPENCOMP < 0 Then CANTIPENREC = 0
           ARECIBIR = CANTIPENCOMP + CANTIPENREC
           UltCompra# = CostoUltimaCompra#(CI%)
           GRILLAGRAL.GRID.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(ARECIBIR, "F10")) ' PENDIENTE A INGRESAR POR COMPRAS O RECPCIONES PENDIENTES DE APROBASR
           GRILLAGRAL.GRID.TextMatrix(i&, 10) = TRIM$(FORMATNUM$(UltCompra, "F10.2")) ' Ult.Compra
       Next i&
       '

       Call COLOREAR_GRILLASQL(GRILLAGRAL.GRID, &HE0E0E0)
       GRILLAGRAL.GRID.Redraw = True
       GRILLAGRAL.GRID.Visible = True
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.Caption = TITU$
       
       GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
       GRILLAGRAL.mnuMenu1.Visible = True
       GRILLAGRAL.TXTBUSCAR = ""
       GRILLAGRAL.SinEfectoClickEnGrilla% = 1
       
       Screen.MousePointer = 1
       GRILLAGRAL.Show 1
       
       Dim MSF As msFlexGrid
       If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call importarExcelPorcentajeEfectividadCotizacionesVsVentas(FECHATEX$(HOY(HOS$)), TRIM$(ECOARCH$("TAFAMIL", FAM$)), TRIM$(ECOARCH$("GRUCOVEN", MKI$(GCONT%))), TRIM$(FORMATNUM$(PORCENTAJE#, "F8.2")), SAFETEXT$(NMESES&), "" & fechaDesde & " al " & fechaHasta & "")
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
       End If
       GoTo 99
    Else
       Call COMUNI("No se Han Encontrado Artículos Según Filtros")
       Screen.MousePointer = 1
       GoTo 99
    End If
    'FALTA CREAR EL EXCEL -AGREGANDO COLUMNA DE PORCENTAJE
99
End Sub

Private Sub mnuestacotiVSventas_Click()
    Screen.MousePointer = 11
20  Campos$ = "CI/F0;Codigo/A20;Descripción/A32;Cotizaciones/F12;T.U.Cot./F9;Facturas/f12;T.U.Fact./f9;Stock/F8;A Recibir/F9"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    
    Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
    If PERIO$ = "" Then GoTo 99
     
    
    
    OBX$ = OBJPLA$("GCON-FAMI", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    '
    FAM$ = TRIM$(Mid$(OBX$, 1, 2))
    GCONT% = CVI(Mid$(OBX$, 3, 2))
    '
    GRILLAGRAL.GRID.Redraw = False
    
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)
    CantiDias = DIENTRE%(DESDEX%, HASTAX%)
    COEF = CantiDias / 30
    If COEF < 1 Then COEF = 1
    'CARGA PRIMERO LA INFO DEL PRESUPUESTO
'select count(*) AS VECES ,CodiInt, CodiExt, Descrip, sum(Cantit) as canticotizadas,M.FAMILIA, M.GRUCOVEN from presuven p with(NOLOCK)
'inner join master M ON M.CODINT = P.CODIINT
' where FechEmis >= '01/NOV/2017' and FechEmis  <= '12/DIC/2018' and codiint > 0 AND M.FAMILIA = 6 AND GRUCOVEN = 0
'group by codiint, CodiExt, Descrip,M.FAMILIA, M.GRUCOVEN order by Descrip
    Call VectArticulos.cargarvector
30  SQLX$ = "select count(*) AS VECES ,CodiInt, CodiExt, sum(Cantit) as canticotizadas ,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + "  from presuven p with(NOLOCK) inner join master M ON M.CODINT = P.CODIINT "
    SQLX$ = SQLX$ + " where FechEmis >= '" & DESD & "'"
    SQLX$ = SQLX$ + " and FechEmis  <= '" & HAST & "'"
    SQLX$ = SQLX$ + " and codiint > 0"
    If FAM$ <> "" Then SQLX$ = SQLX$ + " AND M.FAMILIA =  '" & FAM$ & "'"
    If GCONT% > 0 Then SQLX$ = SQLX$ + " AND GRUCOVEN = " & GCONT%
    SQLX$ = SQLX$ + " group by codiint, CodiExt,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + " order by CodiExt"
    Set RS1 = READSET(SQLX$)
    J& = 0
    With RS1
      Do While Not .EOF
        CI = XVALO(!CodiInt)
        Codigo$ = SAFETEXT$(!CodiExt)
        DESCRI$ = VectArticulos.Descripcion(CI)
        Veces = XVALO(!Veces)
        CANTICOTI = XVALO(!canticotizadas)
        '
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESCRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(Veces, "F10"))
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(CANTICOTI, "F10"))
        .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    If J& > 0 Then
       'AHORA LEE LA GRILLA QUE YA CARGO, Y LE AGREGA LOS DATOS DEL MOVISTO NO LO HAGO ARRIBA POR UNA CUESTION DE VELOCIDAD.
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           Call VECESCANTFACTURADA(CI, DESD, HAST, RET_VECES, RET_CANTIFACTU)
           GRILLAGRAL.GRID.TextMatrix(i&, 6) = TRIM$(FORMATNUM$(RET_VECES, "F10"))
           GRILLAGRAL.GRID.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(RET_CANTIFACTU, "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           GRILLAGRAL.GRID.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(STOCKACT(CI), "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           CANTIPENREC = CANRECEP(CI): If CANTIPENREC < 0 Then CANTIPENREC = 0
           CANTIPENCOMP = STOCOPEN(CI): If CANTIPENCOMP < 0 Then CANTIPENREC = 0
           ARECIBIR = CANTIPENCOMP + CANTIPENREC
           GRILLAGRAL.GRID.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(ARECIBIR, "F10")) ' PENDIENTE A INGRESAR POR COMPRAS O RECPCIONES PENDIENTES DE APROBASR
       Next i&
       '

       Call COLOREAR_GRILLASQL(GRILLAGRAL.GRID, &HE0E0E0)
       GRILLAGRAL.GRID.Redraw = True
       GRILLAGRAL.GRID.Visible = True
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.Caption = "Estadística de Cotizaciones y Ventas - Período Seleccionado: " & PERIO$
       
       Screen.MousePointer = 1
       GRILLAGRAL.Show 1
       GoTo 20
    Else
       Call COMUNI("No se Han Encontrado Artículos Según Filtros")
       Screen.MousePointer = 1
       GoTo 20
    End If
99  Screen.MousePointer = 1
End Sub


Sub VECESCANTFACTURADA(CI, DESDEX, HASTAX, RET_VECES, RET_CANTIFACTU)

            RET_VECES = 0: RET_CANTIFACTU = 0 ' VALORES DE RETORNO
            SQLX$ = " EXECUTE VECES_Y_CANT_FACTURADA " & CI
            SQLX$ = SQLX$ + ", '" & DESDEX & "'"
            SQLX$ = SQLX$ + ",  '" & HASTAX & "'"
            SQLX$ = SQLX$ + ", " & CodiEmp%
            Set rs = FLEXCONN.Execute(SQLX$)
            If Not rs.EOF Then
                RET_VECES = XVALO(rs.Fields(0))
                RET_CANTIFACTU = XVALO(rs.Fields(4)) - XVALO(rs.Fields(5)) '4 ES CANTSALID Y 5 ES CANTINGRE
            End If
            rs.Close
            Set rs = Nothing
End Sub

Private Sub mnuItemSugeridosIncorporarStock_Click()
    Screen.MousePointer = 11
    NMESES& = 0
    
20  Campos$ = "CI/F0;Codigo/A20;Descripción/A32;Cotizaciones/F12;T.U.Cot./F9;Facturas/f12;T.U.Fact./f9;Stock/F8;A Recibir/F9;Costo/F10.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    '
    Static ABUSQUE$
25  ABUSQUE$ = XVALO(InputBox$("Cantidad de Meses", "Ingrese Cantidad de Meses a Considerar", ABUSQUE$))
    NMESES& = XVALO(ABUSQUE$)
    If NMESES& < 1 Then
       Call COMUNI("Dato Ingresado Nó Válido")
       GoTo 99
    End If
    '
    Static ABUSQUE2$
28  ABUSQUE2$ = XVALO(InputBox$("Cantidad de Veces Cotizado", "Ingrese Cantidad de Veces a Considerar", ABUSQUE2$))
    NVECES& = XVALO(ABUSQUE2$)
    If XVALO(NVECES&) < 1 Then
       Call COMUNI("Dato Ingresado Nó Válido")
       GoTo 99
    End If
    


    ' Obtener la fecha actual
    fechaActual = Date
    
    ' Calcular la fecha "desde" restando la cantidad de meses
    fechaDesde = DateAdd("m", -1 * NMESES&, fechaActual)
    fechaHasta = DateAdd("d", -1, DateAdd("m", NMESES&, fechaDesde))
    DESDEX% = CVINT(fechaDesde)
    HASTAX% = CVINT(fechaHasta)
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)


'    Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
'    If PERIO$ = "" Then GoTo 99
    
    
    TITU$ = "Items  Sugeridos Para incorporar al Stock Cotizados Más de: " & NVECES& & " Veces en los  Últimos " & NMESES& & " Meses"
    PERIO$ = " (" & fechaDesde & " al " & fechaHasta & ")"
    TITU$ = TITU$ + PERIO$
    
    OBX$ = OBJPLA$("GCON-FAMI", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    '
    FAM$ = TRIM$(Mid$(OBX$, 1, 2))
    GCONT% = CVI(Mid$(OBX$, 3, 2))
    '
    GRILLAGRAL.GRID.Redraw = False
    
    DESD = FETEXCOR1$(DESDEX%)
    HAST = FETEXCOR1$(HASTAX%)
    CantiDias = DIENTRE%(DESDEX%, HASTAX%)
    COEF = CantiDias / 30
    If COEF < 1 Then COEF = 1
    'CARGA PRIMERO LA INFO DEL PRESUPUESTO
'select count(*) AS VECES ,CodiInt, CodiExt, Descrip, sum(Cantit) as canticotizadas,M.FAMILIA, M.GRUCOVEN from presuven p with(NOLOCK)
'inner join master M ON M.CODINT = P.CODIINT
' where FechEmis >= '01/NOV/2017' and FechEmis  <= '12/DIC/2018' and codiint > 0 AND M.FAMILIA = 6 AND GRUCOVEN = 0
'group by codiint, CodiExt, Descrip,M.FAMILIA, M.GRUCOVEN order by Descrip
    Call VectArticulos.cargarvector
30  SQLX$ = "select count(*) AS VECES ,CodiInt, CodiExt, sum(Cantit) as canticotizadas ,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + "  from presuven p with(NOLOCK) inner join master M ON M.CODINT = P.CODIINT "
    SQLX$ = SQLX$ + " where FechEmis >= '" & DESD & "'"
    SQLX$ = SQLX$ + " and FechEmis  <= '" & HAST & "'"
    SQLX$ = SQLX$ + " and codiint > 0"
    If FAM$ <> "" Then SQLX$ = SQLX$ + " AND M.FAMILIA =  '" & FAM$ & "'"
    If GCONT% > 0 Then SQLX$ = SQLX$ + " AND GRUCOVEN = " & GCONT%
    SQLX$ = SQLX$ + " group by codiint, CodiExt,M.FAMILIA, M.GRUCOVEN "
    SQLX$ = SQLX$ + " HAVING count(*) > " & NVECES&
    SQLX$ = SQLX$ + " order by CodiExt"
    Set RS1 = READSET(SQLX$)
    J& = 0
    With RS1
      Do While Not .EOF
        CI = XVALO(!CodiInt)
        Codigo$ = SAFETEXT$(!CodiExt)
        DESCRI$ = VectArticulos.Descripcion(CI)
        Veces = XVALO(!Veces)
        CANTICOTI = XVALO(!canticotizadas)
        '
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESCRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(Veces, "F10"))
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(CANTICOTI, "F10"))
        .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    If J& > 0 Then
       'AHORA LEE LA GRILLA QUE YA CARGO, Y LE AGREGA LOS DATOS DEL MOVISTO NO LO HAGO ARRIBA POR UNA CUESTION DE VELOCIDAD.
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           Call VECESCANTFACTURADA(CI, DESD, HAST, RET_VECES, RET_CANTIFACTU)
           GRILLAGRAL.GRID.TextMatrix(i&, 6) = TRIM$(FORMATNUM$(RET_VECES, "F10"))
           GRILLAGRAL.GRID.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(RET_CANTIFACTU, "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           GRILLAGRAL.GRID.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(STOCKACT(CI), "F10"))
       Next i&
       '
       For i& = 1 To GRILLAGRAL.GRID.Rows - 1
           CI = GRILLAGRAL.GRID.TextMatrix(i&, 1)
           CANTIPENREC = CANRECEP(CI): If CANTIPENREC < 0 Then CANTIPENREC = 0
           CANTIPENCOMP = STOCOPEN(CI): If CANTIPENCOMP < 0 Then CANTIPENREC = 0
           ARECIBIR = CANTIPENCOMP + CANTIPENREC
           UltCompra# = CostoUltimaCompra#(CI%)
           GRILLAGRAL.GRID.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(ARECIBIR, "F10")) ' PENDIENTE A INGRESAR POR COMPRAS O RECPCIONES PENDIENTES DE APROBASR
           GRILLAGRAL.GRID.TextMatrix(i&, 10) = TRIM$(FORMATNUM$(UltCompra, "F10.2")) ' Ult.Compra
       Next i&
       '

       Call COLOREAR_GRILLASQL(GRILLAGRAL.GRID, &HE0E0E0)
       GRILLAGRAL.GRID.Redraw = True
       GRILLAGRAL.GRID.Visible = True
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.Caption = TITU$
       
       GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
       GRILLAGRAL.mnuMenu1.Visible = True
       GRILLAGRAL.TXTBUSCAR = ""
       GRILLAGRAL.SinEfectoClickEnGrilla% = 1


       Screen.MousePointer = 1
       GRILLAGRAL.Show 1
        Dim MSF As msFlexGrid
       If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
               Call importarExcelSugeridosParaIngresarStock(FECHATEX$(HOY(HOS$)), TRIM$(ECOARCH$("TAFAMIL", FAM$)), TRIM$(ECOARCH$("GRUCOVEN", MKI$(GCONT%))), SAFETEXT$(NVECES&), SAFETEXT$(NMESES&), "" & fechaDesde & " al " & fechaHasta & "")

            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
       End If
       GoTo 99
       GoTo 99
    Else
       Call COMUNI("No se Han Encontrado Artículos Según Filtros")
       Screen.MousePointer = 1
       GoTo 99
    End If
99  Screen.MousePointer = 1
End Sub

Private Sub mnumigramaster_Click()
   mnumigramaster.Enabled = False
   If CLACONTRA% < 1 Then GoTo 99
   MIGRADOX = 0
   Call VectArticulos.cargarvector
   
   'RECORRO LA TABLA MASTER JUNTANDO CON PEDDETA, PRESUVEN,OCOMDETA,BOLRECEP Y MOVISTO
   SQLX$ = " SELECT M.CODINT AS MASTERCI,P.CODIINT AS PRESUVENCI,D.CODIINT AS PEDDETACI,MO.COD_INTE AS MOVISTOCI, "
   SQLX$ = SQLX$ + " B.COD_INTE AS BOLRECEPCI, O.COD_INTE AS OCOMDETACI ,"
   SQLX$ = SQLX$ + " RE.COD_INTE AS RESERVACI,ORDF.COD_INTE AS ORDEFABRCI "
   SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) "
   SQLX$ = SQLX$ + " LEFT JOIN  PRESUVEN P ON M.CODINT = P.CODIINT"
   SQLX$ = SQLX$ + " LEFT JOIN PEDDETA D ON M.CODINT = D.CODIINT"
   SQLX$ = SQLX$ + " LEFT JOIN MOVISTO MO ON M.CODINT = MO.COD_INTE "
   SQLX$ = SQLX$ + " LEFT JOIN OCOMDETA O ON M.CODINT = O.COD_INTE"
   SQLX$ = SQLX$ + " LEFT JOIN BOLRECEP B ON M.CODINT = B.COD_INTE "
   
   SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR ORDF ON M.CODINT = ORDF.COD_INTE "
   SQLX$ = SQLX$ + " LEFT JOIN RESERVA RE ON M.CODINT = RE.COD_INTE "
   
   
   SQLX$ = SQLX$ + " ORDER BY M.CODINT"
   Set rs = READSET(SQLX$)
   With rs
    Screen.MousePointer = 11
    Do While Not .EOF
      CIMAST = XVALO(!MASTERCI)
      If CIMAST <> CIMAST_ant Then
        CIPRESU = XVALO(!PRESUVENCI) 'SI EXISTE EN PRESUVEN DEVUELVE EL CODIGO INTERNO
        CIPEDD = XVALO(!PEDDETACI)   'SI EXISTE EN PEDDETA DEVUELVE EL CODIGO INTERNO
        CIMOVISTO = XVALO(!MOVISTOCI) 'SI EXISTE EN MOVISTO DEVUELVE EL CODIGO INTERNO
        CIBOLRECEP = XVALO(!BOLRECEPCI) 'SI EXISTE EN PRESUVEN DEVUELVE EL CODIGO INTERNO
        CIOCOMDETA = XVALO(!OCOMDETACI) 'SI EXISTE EN PRESUVEN DEVUELVE EL CODIGO INTERNO
        CIORDEFABR = XVALO(!ORDEFABRCI) 'SI EXISTE EN ORDEFABR DEVUELVE EL CODIGO INTERNO
        CIRESEREVA = XVALO(!RESERVACI) 'SI EXISTE EN RESERVA DEVUELVE EL CODIGO INTERNO
        '
        resultadox = CIMAST - (CIPRESU + CIPEDD + CIMOVISTO + CIBOLRECEP + CIOCOMDETA + CIORDEFABR + CIRESEREVA)
        If resultadox > 0 Then
           If EXISMASTER2%(VectArticulos.Codigo(CIMAST)) < 1 Then
             Call Agregar_Registro_Master_A_Master2(CIMAST, 1) ' (PARAMETRO=1) VALIDA SI TIENE FORMULA O ES PARTE DE
             MIGRADOX = MIGRADOX + 1
           End If
        End If
      End If
      CIMAST_ant = CIMAST
        
50    .MoveNext

    Loop
   End With
   rs.Close
   Set rs = Nothing
   Call COMUNI("Proceso Terminado (Cantidad Migrada: " & MIGRADOX & ")")
99 Screen.MousePointer = 1
   Set VectArticulos = Nothing
   mnumigramaster.Enabled = True
End Sub
Private Sub mnumodificarCliente_Click()
   NPR& = XVALO(Text17)
   If Text21.Visible = True Then ' SI ESTA MODIFICANDO
     If Text21.Width >= 7000 Then  'SI TIENE EL TAMAÑO DE MODIFICACION (POR QUE EN PEQUEÑO LO ABRE COMO BANDERA CUANDO GRABA AUTOMATICAMENTE.
        If NPR& > 0 Then 'SI TIENE NUMERO DE PRESUPUESTO
           If XVALO(Text1) > 0 Then
             CONDI$ = "NroPres = " & NPR&
             NROPE& = XVALO(VALOBADA("PRESUVEN", "NroPedi", CONDI$, "NOMESS"))
             If NROPE& > 0 Then
                Call COMUNI("Imposible Modificar Cliente - Presupuesto Seleccionado Posee Pedido Generado (" & NROPE& & ")")
                Exit Sub
           End If
           
           OPX% = COMALTER%("Confirmación de Cambio de Cliente\Confirma Modificar Cliente: " & XVALO(Text1) & "\\Cancelar\Continuar")
           If OPX% < 2 Then Exit Sub
               Call SELECHO("DEUDOR1")
               NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
               If NCLIE > 0 Then
                  RACLI$ = rasocli$(NCLIE)
                  CONDI$ = "NroPres = " & NPR&
                  Call ACTUREGISTRO("PRESUVEN", "NroClie", CONDI$, NCLIE, RegAf&, "NOMESS")
                  Call ACTUREGISTRO("PRESUVEN", "RASOCLIE", CONDI$, RACLI$, RegAf&, "NOMESS")
                  Text1 = ""
                  Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
                  Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
               End If
           End If
        End If
     End If
    Else
        Call COMUNI("Imposible Modificar Cliente - Debe Primero Abrir en Modo de Edición el Presupuesto")
        Exit Sub
     
    End If

End Sub

Private Sub mnuP_Click()
    Call CONECRUN("PEDIDO07", "PEDIDO DE CLIENTES")
End Sub

Private Sub mnurepocli_Click()
   REPOEQUI.Show 1
End Sub

Private Sub mnusetup_Click()
   OPPRES07.Show 1
End Sub

Private Sub mnuTablaEquiClientes_Click()
       ABMEQUIPO.Show 1
'        CAMPOS$ = "ID/f0;R.Social/A20;Equipo/A20;Codigo/Descripcion/D8;Costo/F10.2;Importe/F12.2;Formula/F10.2;%/F7.2;Venta/F10.2;NUMORDEN/F0;O.C./F7;Entregar/D8"
'        Call CARGA_ENCABEZADO(VENTAB2.MSF, CAMPOS$, VENTAB2)
'        Call CARGA_ENCABEZADO(VENTAB2.MSFFORMULA, CAMPOS$, VENTAB2)
''''        FORMATOS$ = GENERA_FORMATOS$(CAMPOS$) ' DEVUELVE ESTO = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
''''        FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(CAMPOS$)
''''        'CAMPOS$ = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
''''        Call CARGA_ENCABEZADO(VENTAB2.MSFTOTALES, FORMATOSCONBARRA$, VENTAB2)
''''
''''        VENTAB2.MSFFORMULA.Rows = 2  'DEBEN IR EN FILA 1
''''        VENTAB2.MSFFORMULA.TextMatrix(1, 7) = "4;*;6"
''''        VENTAB2.MSFFORMULA.TextMatrix(1, 9) = "=PORCENT_ENTRE_2_VALORES;10;6" 'DEBE LLEVAR SIMBOLO =
''''        Call AGREGAFILA(VENTAB2.MSFFORMULA, FORMATOS$, 2) 'DEBEN IR EN LA FILA 2
''''        '
''''        'EDITABLES
''''        VENTAB2.EDITABLES.Caption = "4;10"
''''        VENTAB2.MSFTOTALES.Rows = 2
''''        VENTAB2.MSFTOTALES.TextMatrix(1, 7) = "TOTAL"
''''
''''
''''        '
''''        j& = 0
''''        FIN& = ULTREG&("!BOLREDET")
''''        For YX& = 1 To FIN&
''''           YY$ = REGLEIDO$("!BOLREDET", YX&)
''''           CI1% = CVI(Mid$(YY$, 83, 2))
''''           CAN# = CVD(Mid$(YY$, 89, 8))
''''           If CI1% > 0 Then
''''               If CAN# > 0.005 Then
''''                 '
''''                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
''''                 If UNID$ = "" Then UNID$ = UNIMED$(CI1%)
''''                 FEX = CVI(Mid$(YY$, 155, 2))
''''                 FESOLEN$ = FECHATEX$(FEX)
''''                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
''''                 '
''''                 If NULOTE$ = "" Then
''''                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
''''                 End If
''''                 NUOC& = CVS(Mid$(YY$, 151, 4))
''''                 If NUOC& <> NUOC_ANT& Then 'SI CAMBIA LA ORDEN DE COMPRA VUELVE A BUSCAR EL COEFICIENTE
''''                   COEFI = COEFI_COMPRA_SIEMPRE_DOLAR(NUOC&)
''''                 End If
''''                 If COEFI < 1 Then COEFI = 1
''''
''''                 NORD% = Asc(Mid$(YY$, 98, 1))
''''                 ESPRIN$ = Mid$(YY$, 5, 64)
''''                 NUOC& = CVS(Mid$(YY$, 151, 4))
''''                 FENTRESO% = CVI(Mid$(YY$, 155, 2))
''''
''''                 With VENTAB2.MSF
''''                    j& = j& + 1
''''                    VENTAB2.MSF.Rows = j& + 1
''''                    .TextMatrix(j&, 1) = CODEXT$(CI1%)
''''                    .TextMatrix(j&, 2) = DESCRIT0$(CI1%)
''''                    .TextMatrix(j&, 3) = UNID$
''''                    .TextMatrix(j&, 4) = FORMATNUM$(CAN#, "F10.1")
''''                    .TextMatrix(j&, 5) = FESOLEN$
''''                        COSTOFOB1# = COSTO_SEGUN_COMPRA(CI1%, NUOC&, NORD%)
''''                        COSTOFOB1# = COSTOFOB1# / COEFI
''''                    .TextMatrix(j&, 6) = TRIM$(FORMATNUM$(COSTOFOB1#, "F12.2"))
''''                    .TextMatrix(j&, 7) = TRIM$(FORMATNUM$(COSTOFOB1# * CAN#, "F12.2"))
''''                        VIAENTR% = VIAENTRE_COMPRA(NUOC&)
''''                        PVTA# = PRECIOVENTA_X_FORMULA(COSTOFOB1#, VIAENTR%, CI1%, 0, NPROV, (-1))
''''                        PVTAEDITADO# = PRECIODEVENTA_AHP(CI1%, NPROV, VIAENTR%)
''''                        'If PVTAEDITADO# < 0.005 Then PVTAEDITADO# = PVTA# 'PRESENTO VACIO SI NO TIENE IMPORTE.
''''                    .TextMatrix(j&, 8) = TRIM$(FORMATNUM$(PVTA#, "F12.2"))
''''                    .TextMatrix(j&, 9) = TRIM$(FORMATNUM$(PORCENT_ENTRE_2_VALORES(PVTAEDITADO#, COSTOFOB1#), "F9.2"))
''''                    .TextMatrix(j&, 10) = TRIM$(FORMATNUM$(PVTAEDITADO#, "F9.2"))
''''                    .TextMatrix(j&, 11) = SAFETEXT$(NORD%)
''''                    .TextMatrix(j&, 12) = SAFETEXT$(NUOC&)
''''                    .TextMatrix(j&, 13) = FECHATEX$(FENTRESO%)
''''                    NUOC_ANT& = NUOC&
''''
''''                 End With
''''               End If
''''           End If
''''         Next YX&
''''         If VENTAB2.MSF.Rows > 1 Then
''''            VENTAB2.Show 1
''''            If VENTAB2.APROBADO% > 0 Then
''''             JJ& = 0: CONDI$ = ""
''''             Dim viaentregax()
''''             ReDim Preserve viaentregax(1)
''''             With VENTAB2
''''               For I& = 1 To .MSF.Rows - 1
''''                 Y$ = REGBLAN$("!BOLREDET")
''''                   CI1% = XVALO(CODINT(.MSF.TextMatrix(I&, 1)))
''''                   prexiventa# = XVALO(.MSF.TextMatrix(I&, 10))
''''                   If prexiventa# > 0 Then ' cargo en un  VECTOR PARA LEER SOLO UNA VEZ  Y HACER MAS RAPIDA LA ACTUALIZACION
''''                     If NUOC& > UBound(viaentregax) Then ReDim Preserve viaentregax(NUOC&)
''''                     If XVALO(viaentregax(NUOC&)) < 1 Then viaentregax(NUOC&) = VIAENTRE_COMPRA(NUOC&)
''''                     Call RUTINA_ACTU_COSTOPROVE(NPROV, CodiEmp%, CI1%, XVALO(viaentregax(NUOC&)), prexiventa#)
''''                   End If
''''                 Call REPLA(Y$, MKI$(CI1%), 83, 2)
''''                 Call REPLA(Y$, .MSF.TextMatrix(I&, 3), 85, 4)
''''                   CAPEN = XVALO(.MSF.TextMatrix(I&, 4))
''''                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 89, 8)
'''''                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
''''                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
''''                 '
''''                 Call REPLA(Y$, Chr$(XVALO(.MSF.TextMatrix(I&, 11))), 98, 1)
''''                 Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
''''                 '
''''                 NUOC& = XVALO(.MSF.TextMatrix(I&, 12))
''''                 Call REPLA(Y$, MKS$(NUOC&), 151, 4)
''''                   FENTRESO% = FECHANUM(.MSF.TextMatrix(I&, 11))
''''                 Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
''''                 Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
''''                 '
''''                 JJ& = JJ& + 1
''''                 Call GRAREG("!BOLREDET", Y$, JJ&)
''''               Next I&
''''             End With
''''            End If
''''         End If
''''         Call SETULTREG("!BOLREDET", JJ&)
         '
'         'RECORRO SOLO LOS CODIGOS QUE SE LES AGREGO PRECIO DE VENTA EN AL GRILLA Y LE GRABO EL PRECIO DEL VECTOR QUE CARGUE CUANDO LEI LA GRILLA.
'         CONDI$ = Left$(CONDI$, Len(CONDI$) - 2) ' LE QUITO LA COMA FINAL
'         CONDI$ = "(" & CONDI$ & ")"
'         SQLX$ = "SELECT PRECIOVENTA,CODINT FROM COSTOXPROVE "
'         SQLX$ = SQLX$ + " WHERE " & CONDI$
'         SQLX$ = SQLX$ + " and Nume_Cli = " & NPROV
'         'ver si esta bien el campo de proveedor y falta la via de entregea
'         SQLX$ = SQLX$ + " ORDER BY CODINT"
'         Dim RS As ADODB.Recordset
'         Set RS = New ADODB.Recordset
'         RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'         With RS
'            Do While Not .EOF
'                CI1% = XVALO(!CODINT)
'                !PRECIOVENTA = XVALO(prexiventa#(CI1%))
'                .Update
'              .MoveNext
'            Loop
'         End With
'         Erase viaentregax
         
'     Set RS = READSET(SQLX$)
'       Y$ = REGBLAN$("!BOLREDET")
'             Call REPLA(Y$, MKI$(CI%), 83, 2)
'             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
'                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
'             '
'             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
'             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
'             '
'             Call REPLA(Y$, LOTECO$, 107, 12)
'             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
'             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
'             Call REPLA(Y$, MKS$(!N_ORSERV), 189, 4)
    

End Sub



Private Sub mnutablaequiposxcliente_Click()
    NC = XVALO(Text1)
    If NC > 0 Then
      ABMEQUICLI.TXTCTACLIEN = NC
    End If
    '
    ABMEQUICLI.Show 1
    
End Sub

Private Sub mnutamonedas_Click()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")

End Sub

Private Sub mnutexmail_Click()
    TEXTOMAIL.Show 1
End Sub

Private Sub mnutraz_Click()

    If CLACONTRA% > 0 Then
      Call PoneEnCeroNull("MASTER", "TRAZABLE", RegAf&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "TRAZABLE", " CODINT > 0 AND TRAZABLE < 1", 1, RegAf&, "NOMESS")
      Call COMUNI("Se han Modificado: " & RegAf& & " Registros")
    End If
    '
End Sub

Private Sub mnuvalofer_Click()
      Call TRALOCAL("VALIOFER", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("VALIOFER")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!VALIOFER")
         X$ = REGLEIDO$("!VALIOFER", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("VALIOFER", "")
      Call FRESHECHO("VALIOFER")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")

End Sub

Private Sub mnuvidetasol_Click()
    CONDI$ = "NroPres > 0   ORDER BY NroPres DESC"
10  Call CARGALISEL("INDICEP1", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEP1/INDICE DE PRESUPUESTOS")
    NumePresu& = XVALO(VALACT1("INDICEP1"))
    If NumePresu& < 1 Then Exit Sub
    '
    Call GENERA_TABLA_COMPRAS_DE_PRESUPUESTO
    '
    'ARMO LA PRESENTACION DE LA GRILLA EN EL FRMZELECHO
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Generación de Ordenes de Compra - Presupuesto N°: " & NPRES&
    '
    'PRESENTO TODAS LAS SOLICITUD DE Compras RELACIONADAS CON EL PRESUPUESTO
    i& = 0
    SQLX$ = "SELECT DISTINCT(Num_Prov),NumSolCompra FROM PRESUCOMPRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPres = " & NumePresu&
    SQLX$ = SQLX$ + " ORDER BY Num_Prov,NumSolCompra"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        NP% = XVALO(!Num_Prov)
        FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
        FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
           NSOLIC& = XVALO(!NumSolCompra)
           OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
        If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = NSOLIC&
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'SI NO TIENE SOLICITUD AVISA Y VUELVE A LISTA DE PRESUPUESTOS
    If i& < 1 Then
       Call COMUNI("No se Detectaron Solicitud de Compras Para El Presupuesto  Seleccionado")
       GoTo 10
    End If
50  FRMZELECHO.Show 1
    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
    OCOMPRA& = XVALO(RESP_ZELE_VECT(3))
    NSOLIC& = XVALO(RESP_ZELE_VECT(4))
    '
    If NSOLIC& > 0 Then
       Call FINAL("? PRESCOM2 (" & TRIM$(SAFETEXT$(NSOLIC&)) & ")")
    End If
    'Call FINAL("? MOVCONSU (" & NORFA$ & ")")

End Sub

Private Sub MSF1_Click()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows - 1 Then Exit Sub

          COLUM = MSF1.MouseCol
          If COLUM <> 8 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF1.COL = 8
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            .Visible = True
            .ToolTipText = MSF1.row
            .TEXT = MSF1.TEXT
            .SetFocus
          End With
          Call ACTU_TOTAL
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          '
    REG& = MSF1.Rows
    If REG& < 2 Then Exit Sub
    If Button = 2 Then
        REG& = MSF1.MouseRow
        If REG& < 1 Then Exit Sub
        ARGUBUS$ = MSF1.TextMatrix(REG&, 1)
        Call MUESTRADATCOD(ARGUBUS$)
     End If
     '
End Sub
Private Sub MSF1_MouseUp(Button As Integer, _
         Shift As Integer, X As Single, Y As Single)
          Static CurrentWidth As Single
          If MSF1.Rows <= 1 Then Exit Sub
          If MSF1.CellWidth <> CurrentWidth Then
              On Error Resume Next
              TXTMOVIL.Width = MSF1.CellWidth
              CurrentWidth = MSF1.CellWidth
              On Error GoTo 0
          End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSF1_Scroll()
   If TXTMOVIL.Visible = True Then
     'TXTMOVIL.Left = MSF1.CellLeft
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF1.TEXT = TXTMOVIL
     TXTMOVIL.Visible = False
   End If
End Sub

Private Sub munanusol_Click()


10  CONDI$ = "NroPres > 0  ORDER BY NroPres DESC"
    Call CARGALISEL("INDICEP1", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEP1/INDICE DE PRESUPUESTOS")
    NumePresu& = XVALO(VALACT1("INDICEP1"))
    If NumePresu& < 1 Then Exit Sub
    '
    Call GENERA_TABLA_COMPRAS_DE_PRESUPUESTO
    '
    'ARMO LA PRESENTACION DE LA GRILLA EN EL FRMZELECHO
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cta.Prov/F9;R.Social/A40;O.Compra/F12;N.Solic./F8.0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Anulación de Solicitud de Compras - Presupuesto N°: " & NPRES&
    '
    'PRESENTO TODAS LAS SOLICITUD DE Compras RELACIONADAS CON EL PRESUPUESTO
30  i& = 0
    FRMZELECHO.msf2.Rows = 1
    SQLX$ = "SELECT DISTINCT(Num_Prov),NumSolCompra FROM PRESUCOMPRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPres = " & NumePresu&
    SQLX$ = SQLX$ + " ORDER BY Num_Prov"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        NP% = XVALO(!Num_Prov)
        FRMZELECHO.msf2.TextMatrix(i&, 1) = NP%
        FRMZELECHO.msf2.TextMatrix(i&, 2) = rasocli$(-1 * NP%)
           NSOLIC& = XVALO(!NumSolCompra)
           OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
        If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = NSOLIC&
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'SI NO TIENE SOLICITUD AVISA Y VUELVE A LISTA DE PRESUPUESTOS
    If i& < 1 Then
       Call COMUNI("Sin Solicitudes de Compra Para El Presupuesto  Seleccionado")
       GoTo 10
    End If
    
50  FRMZELECHO.Show 1
    '
    
    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
    OCOMPRA& = XVALO(RESP_ZELE_VECT(3))
    NSOLIC& = XVALO(RESP_ZELE_VECT(4))
    '
    CONDI$ = " NUME_OCOM = " & OCOMPRA& & " AND Stat_Ocom < 9"
    OCOMPRAOCOMENCA& = XVALO(VALOBADA("OCOMENCA", "NUME_OCOM", CONDI$, "NOMESS"))
    'SI TIENE ORDEN DE COMPRA AVISA Y VUELVE A LA LISTA DE SELECCION
    If OCOMPRAOCOMENCA& > 0 Then
       Call COMUNI("Solicitud Seleccionada Generó La Orden de Compra: " & OCOMPRA& & "\Modifique Orden de Compra o Anular Para Generar una Nueva.")
       GoTo 50
    End If
    '
    If NSOLIC& > 0 Then
        OPX% = COMALTER%("Confirma Realmente Anular Solicitud de Compra N°: " & NSOLIC& & "\\ No,Cancelar\Si,Anular")
        If OPX% = 2 Then
           CONDI$ = "NumSolCompra = " & NSOLIC&
           Call BORRAREGISTROS("PRESUCOMPRA", CONDI$, RegAf&, "NOMESS")
           Call COMUNI("Anulación Procesada")
           GoTo 30
        Else
           GoTo 50
        End If
    End If
    '
End Sub

Private Sub muntexobs_Click()
   '
   CargaObservacionesPredeterminadas
End Sub

Private Sub Text1_GotFocus()
   Text1.BackColor = vbYellow
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(txtcodigo)
    End If

End Sub

Private Sub Text1_LostFocus()
  Text1.BackColor = vbWhite
End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(XVALO(Text15), "F10.2")
End Sub

Private Sub Text2_Change()
   Call VectArticulos.cargarvector
   Call VectCliente.cargarvector
End Sub

Private Sub Text29_Change()
   Text28 = ""
   If XVALO(Text29) > 0 Then
      Text28 = ECOARCH$("VALIOFER", MKI(XVALO(Text29)))
   End If

End Sub

Private Sub Text35_DblClick()
    If Text35.Locked = False Then
       Text35.BackColor = &HE0E0E0
       Text35.Locked = True
    Else
       Text35.Locked = False
       Text35.BackColor = &H80000005
       Call PINTATEXT(Text35)
    End If

End Sub

Private Sub Text35_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTREFERENCIA)
    End If
End Sub


Private Sub Text35_LostFocus()
       Text35.BackColor = &HE0E0E0
       Text35.Locked = True

End Sub

Private Sub Timer1_Timer()
 Call PINTATEXT(FRMPRESUAHP.Text1)
 Timer1.Enabled = False
End Sub

Private Sub TXTCANTIDAD_GotFocus()
   TXTCANTIDAD.BackColor = vbYellow
End Sub

Private Sub TXTCANTIDAD_LostFocus()
   TXTCANTIDAD.BackColor = vbWhite
End Sub

Private Sub TXTCODCLIENTE_DblClick()
   If TRIM$(TXTCODCLIENTE.TEXT) <> "" And TRIM$(txtcodigo) <> "" Then
     OPX% = COMALTER%("Realmente Desea Eliminar Este Código de Cliente\\No, Cancelar\Sí Eliminar")
     If OPX% = 2 Then
        CONDI$ = "CODIGO_CLIENTE = '" & TRIM$(TXTCODCLIENTE.TEXT) & "' AND CODINT = " & CODINT%(TRIM$(txtcodigo))
        Call BORRAREGISTROS("CODICLIENTE", CONDI$, RegAf&, "NOMESS")
        If RegAf& > 0 Then
           Call COMUNI("Código Eliminado Correctamente")
           Call LIMPIAR_TEXT_DETALLE
        Else
           Call COMUNI("Código Relacionado Inexistente o no Válido")
        End If
     End If
   End If
End Sub


Function DEMORA_X_VIA(VIA%, Optional NPROV&)
       DEMORAX% = 0
       If NPROV& > 0 Then
            RFF$ = RECFILT$("DEMORAS", 1, MKI$(NPROV&))
            FIN& = Len(RFF$)

            For UK& = 1 To FIN& Step 4
               Call TRACE(20, UK&, FIN&)
               U& = CVS(Mid$(RFF$, UK&, 4))
               X$ = REGLEIDO$("DEMORAS", U&)
               VIALEIDA% = CVI(Mid$(X$, 3, 2))
               If VIALEIDA% = VIA% Then
                  DEMORAX% = CVI(Mid$(X$, 5, 2))
                  If DEMORAX% = 0 Then
                     DEMORAX% = -1
                  End If
                  Exit For
               End If
            Next UK&
       Else
         DEMORAX% = 0
       End If
       DEMORA_X_VIA = DEMORAX%
End Function


Private Sub txtcodigo_DblClick()
    Call Command1_Click
End Sub

Private Sub txtcodigo_LostFocus()
    txtcodigo.BackColor = vbWhite
End Sub

Private Sub TXTCONDIPAGO_Change()
   TXTDESCRIPAGO = ""
   If XVALO(TXTCONDIPAGO) > 0 Then
      TXTDESCRIPAGO = ECOARCH$("CONPAG", Chr(XVALO(TXTCONDIPAGO)))
   End If
End Sub

Private Sub TXTDESCRIPCION_Change()
    On Error Resume Next
    If TRIM$(TXTDESCRIPCION) <> "" Then TXTCANTIDAD.SetFocus
    
    On Error GoTo 0
End Sub

Private Sub TXTMOVIL_Change()
   REG& = XVALO(TXTMOVIL.ToolTipText)
   Resultado# = XVALO(MSF1.TextMatrix(REG&, 4)) * XVALO(TXTMOVIL.TEXT)
   MSF1.TextMatrix(REG&, 9) = FORMATNUM$(Resultado#, "F14.2")
   
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
    'RUTINA PARA QUE AL DAR ENTRER O ESCAPE SAQUE EL FOCO Y PONGA COMO NO VISIBLE EL TEXT DE EDICION
    If KeyAscii = 13 Or KeyAscii = 27 Then
      REG& = XVALO(TXTMOVIL.ToolTipText)
      MSF1.TEXT = TXTMOVIL
      TXTMOVIL.Visible = False
      'Call ACTUCANTI
    End If
       
End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
       REG& = XVALO(TXTMOVIL.ToolTipText)
       MSF1.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(TXTMOVIL), "F12.2")
       TXTMOVIL.Visible = False
     End If
     Call ACTU_TOTAL
     If XVALO(MSF1.TextMatrix(REG&, 8)) < XVALO(MSF1.TextMatrix(REG&, 7)) Then
        Call COMUNI("Atención !!! El Precio de Venta es Menor al Costo")
     End If
End Sub
Private Sub Form_Load()

   'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
   'Call iniciar_variables
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")

    Call CENTRAFORM(Me)
'    Show
    Refresh
    '
    '
    EPAC$ = TRIM$(EMPREAC$)

    Screen.MousePointer = 11
    Call ABRECONEXION
    If InStr(1, EPAC$, "AHP") > 0 Then
       'todos los item de ahp deben ser trazables
       CONDI$ = "(TRAZABLE IS NULL  OR TRAZABLE < 1)"
       Call ACTUREGISTRO("MASTER", "TRAZABLE", CONDI$, 1, RegAf&, "NOMESS")
       Call ACTUREGISTRO("MASTER2", "TRAZABLE", CONDI$, 1, RegAf&, "NOMESS")
       Call BorrarLotesenMovimientosNoCorresponde
       Call RecuperarLoteNotaDeCreditoParciales
       Call RecuperarLote
    End If
    
    Screen.MousePointer = 1
    
15
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    If InStr(1, EPAC$, "AHP") = 0 Then Image1.Picture = LoadPicture("")
    '
'    HOII% = HOY(HOS$)
'    Text9.TEXT = HOS$
    Text17.TEXT = TRIM$(Str$(NUPRESUP&))
    'Call NUCORUB
    '
    ARCHIPRE$ = "MASTER"
    '
    'DOLAR POR DEFECTO
    MONEMI% = 2  ' DOLAR
    MONEMIS = DEMONECO$(2)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    'CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    '
    For U& = 1 To 2
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVAX = CVS(Mid$(X$, 13, 4))
       If U& = 1 Then
           PORIVA0 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
         Else
           PORIVA1 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
       End If
    Next U&
    Call CIERRARCH("GRUPIVA")
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
'
    DTPicker1.Value = Date
    
    Text17 = NUPRESUP&
    'Text20 = "-00"
    '
    Call CREARESECOTI ' CREA CAMPOS
    Call GENERA_TABLA_RELACION_ART_PROV

    Call CARGA_ENCABEZADO(FRMPRESUAHP.MSF1, "Código/A22;Cod.Cte/A22;Descripción/A22;Q/F4.A;Prov/AO;Proveedor/A12;Costo (U$S)/F10.2;Pr.Venta/F8.2;Importe/F10.2;Via Apr./A10;LOTES/A0;CANTI/F0;NPROV/F0;KILOS/f12.2;DEMORA/F7;ValorFob/F0", FRMPRESUAHP, , , , 5)
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
    Call ABRECONEXION
    Call CARGARCOMBOVIAS_TODOS
    Call CREARCAMPOS_FACTURAYGASTOS
    Call CREACAMPO("", "PRESUVEN", "CodigoDelCliente", 10, 32)
    Call CREATABLA_EQUIPOAHP
    Call CREAR_STORED_PROCEDURE
    Call CREAR_STORED_PROCEDURE_LOCAL
    Call ABRECLIEN
    Call VectCliente.cargarvector

    '
'    AAAAAA = PRECIOVENTA_X_FORMULA(1000, 2, 3457)
    Timer1.Enabled = True
End Sub

Sub CARGARCOMBOVIAS_TODOS()
     Combo6.Clear
     FIN& = ULTREG&("KILOVIAS")
     '
     For i& = 1 To FIN&
        X$ = REGLEIDO$("KILOVIAS", i&)
        VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
        REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
        Call REPLA(TX$, VIAX$, 1, 2)
        Call REPLA(TX$, "-", 3, 1)
        Call REPLA(TX$, REFVIA$, 4, 8)
        Combo6.AddItem TX$
     Next i&
     '

End Sub
    Sub CREARESECOTI()
        Call CREACAMPO("VENTAS", "RESE_COTI", "NroPres", 6, 4)
        Call CREACAMPO("VENTAS", "RESE_COTI", "Status", 3, 0)
        Call CREACAMPO("VENTAS", "RESE_COTI", "CodiInt", 4, 0)
        Call CREACAMPO("VENTAS", "RESE_COTI", "IDENLOTE", 10, 24)
        Call CREACAMPO("VENTAS", "RESE_COTI", "CANTIDAD", 6, 0)
        Call CREACAMPO("VENTAS", "RESE_COTI", "ESTADO", 3, 0)
        Call CREACAMPO("VENTAS", "RESE_COTI", "USUARIO", 3, 0)
    End Sub
Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   Dim MIRS As ADODB.Recordset
   strSQL = ("SELECT max(NroPres) as maximo FROM PresuVen")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not IsNull(MIRS!Maximo) Then
      NPP& = 1 + MIRS!Maximo
   End If
   NUPRESUP& = NPP&
   MIRS.Close
   Set MIRS = Nothing
   '
End Function


Private Sub mnuedit_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")
End Sub

Sub CARGA_PRESUPUESTOS()
    Screen.MousePointer = 11
    RESPUESTA_BUSQUEDA$ = ""

    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NroPres > 0 "
    CNRG& = CantRegistros&("PRESUVENTA", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Presupuestos Cargados"
      Exit Sub
    End If

    '
    LISTASELEC.List1.Clear
    LISTASELEC.List2.Visible = False
    LISTASELEC.List1.Visible = True
    '
    'ASIGNO EL ANCHO QUE VA A TENER CADA CAMPO A ESTAS 2 VARIABLES PUBLICAS
    CODIGO_LONG% = 14
    DESCRIP_LONG% = 48
    
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVENTA "
    SQLX$ = SQLX$ + " ORDER BY NroPres "

    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With TABTMP
        Do While (.EOF = False)
          NPRX$ = SAFETEXT$(!NroPres_Version)
          NCLIE& = XVALO(!NroClie)
          RAZONCLI$ = SAFETEXT$(!rasocli)
          FEX = CVINT(!FECHEMIS)
          '
          YY2$ = " - " + AJUSTI$(RAZONCLI$, 20)
          YY2$ = YY2$ + " - " + AJUSTD$(FECHATEX$(FEX), 8)
          COCU$ = TRIM$(NPRX$)
          
          Call REPLA(TTXX$, COCU$, 1, CODIGO_LONG%)
          Call REPLA(TTXX$, YY2$, CODIGO_LONG% + 2, DESCRIP_LONG%)
          JJ& = JJ& + 1
          LISTASELEC.List1.AddItem TTXX$
          '
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          NPXA& = NPX&
          '
        .MoveNext
        Loop
      
    End With
    '
    LISTASELEC.Caption = "INDICE DE PRESUPUESTOS"
    LISTASELEC.List1.ListIndex = LISTASELEC.List1.ListCount - 1
    Screen.MousePointer = 1
60  LISTASELEC.Show 1

End Sub


Private Sub mnuabmcte_Click()
    ABMCLIENTES.Show 1
End Sub

Private Sub MnuArt_Prov_Click()
      '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("MASTER") ' DESPLIEGA LA LISTADO DE ARTICULOS
         '
         COD$ = TRIM$(VALACT1$("MASTER"))
         If COD$ = "" Then Exit Sub
         '
         CI% = CODINT%(COD$)
         ' BUSCO EN LA TABLA PROVEDCOD TODOS LOS REGISTROS QUE COINCIDAN CON EL CI%
         SQLX$ = "SELECT * FROM RELACODPROV WITH (NOLOCK)"
         SQLX$ = SQLX$ + "WHERE COD_INTE = " & CI% & " AND MARBORRA = 0 "
         Dim rs As ADODB.Recordset
         Set rs = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         Call SETULTREG("!PROVIADE", 0)
         JJ& = 0
         With rs
            Do While Not .EOF ' RECORRO LA TABLA Y GRABO EN EL ARCHIVO PARA MOSTRAR LA VENTANA
                X$ = REGBLAN$("PROVIADE")
                Call REPLA(X$, MKI(XVALO(!Num_Prov)), 1, 2)
                Call REPLA(X$, MKD(XVALO(!Precio)), 3, 8)
                Call REPLA(X$, Chr(XVALO(!Moneda)), 11, 1)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 12, 2)
                JJ& = JJ& + 1
                Call GRAREG("!PROVIADE", X$, JJ&)
            .MoveNext
            Loop
         End With
         rs.Close
         Set rs = Nothing
         
10       XYX% = VENTABU%("PROVIADE") ' MUESTRO LA VENTANA
         If XYX% < 0 Then GoTo 5     ' SI CIERRA VUELVE AL SELECHO
         '
         Screen.MousePointer = 11
         For U& = 1 To ULTREG&("!PROVIADE")
            X$ = REGLEIDO$("!PROVIADE", U&)
            COCO% = Asc(Left$(X$, 1))
            If COCO% < 1 Then
               Call MENSERR(24, "Código debe ser Numérico.")
               GoTo 10
            End If
            '
            DECO$ = TRIM$(Mid$(X$, 2, 32))
            If DECO$ = "" Then
               Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
               GoTo 10
            End If
         '
         Next U&
      
         CONDI$ = "COD_INTE=" & CI%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&) ' HACE UN MARBORRA

         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI%
         Set RS2 = New ADODB.Recordset
         RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          With RS2
            For i& = 1 To ULTREG&("!PROVIADE") ' RECORRO EL ARCHIVO Y GRABO
              XX$ = REGLEIDO$("!PROVIADE", i&) ' LA INFORMACION DE LA VENTANA
              NPRO% = CVI(Mid$(XX$, 1, 2))
              PRECIOX# = CVD(Mid$(XX$, 3, 8))
              MON% = Asc(Mid$(XX$, 11, 1))
              FE% = CVI(Mid$(XX$, 12, 2))
              
              .AddNew   ' YA CONVERTIDO
              
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !Num_Prov = NPRO%
              !cod_inte = CI%
              !VIA = 0
              !Precio = PRECIOX#

              If ROUND(PRECIOX#, 2) <> ROUND(XVALO(!Precio), 2) Then
                FE% = HOY(HO$)
              End If
              !Moneda = MON%

              '
              If FE% < 1 Then FE% = HOY(HO$)
              !FECHA = CVDAT(FE%)
              !MARBORRA = 0
              .Update
            Next i&
          End With
         RS2.Close
         Set RS2 = Nothing
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&) ' BORRO LOS MARCADOS CON 1

         Call TRACENTRAL("PROVIADE", "")
         Call FRESHECHO("PROVIADE")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")
End Sub

Private Sub mnuasigcoef_Click()
      Call TRALOCAL("PORCCOEF", "")
      '
      NPROV& = 0
      Call SELECHO("PROVED1")
      
      
      NPROV& = XVALO(VALACT1("PROVED1"))
'      SI TIENE PROVEEDOR SELECCIONADO FILTRA
      
      If NPROV& > 0 Then
            'RECORRO LA VIA DE ENTREGA Y CARGO EN UN VECTOR LOS CODIGOS DE VIAS DE ENTREGA FILTRADO POR PROVEEDOR
            RFF$ = RECFILT$("!PORCCOEF", 6, MKI$(NPROV&)) '4
            FIN& = Len(RFF$)
            For UK& = 1 To FIN& Step 4
               Call TRACE(20, UK&, FIN&)
               U& = CVS(Mid$(RFF$, UK&, 4))
               X$ = REGLEIDO$("!PORCCOEF", U&)
               J& = J& + 1
               Call GRAREG("!PORCCOEF", X$, J&)
            Next UK&
            '
            Call SETULTREG("!PORCCOEF", J&)
      End If
      
      
10        Screen.MousePointer = 1
          VTB% = VENTABU%("PORCCOEF")
          If VTB% < 0 Then Exit Sub
      '
          Screen.MousePointer = 11
          For U& = 1 To ULTREG&("!PORCCOEF")
             X$ = REGLEIDO$("!PORCCOEF", U&)
             COCO% = Asc(Left$(X$, 1))
             If COCO% < 1 Then
                Call MENSERR(24, "Falta Código de Via de Aprovisionamiento - Fila: " & U&)
                GoTo 10
             End If
             '
             COCO2% = Asc(Mid$(X$, 3, 2))
             If COCO2% < 1 Then
                Call MENSERR(24, "Falta Código de Tipo de Coeficiente - Fila: " & U&)
                GoTo 10
             End If
             '
             PORC = Asc(Mid$(X$, 3, 2))
             If PORC < 0.005 Then
                Call MENSERR(24, "Falta Coeficiente - Fila: " & U&)
                GoTo 10
             End If
    
         Next U&
         
      '
            'SI SELECCIONO PROVEEDOR BORRO LOS REGISTROS PARA LUEGO AGREGARLOS AL FINAL
       If NPROV& > 0 Then
         Call FILTERDELETE("PORCCOEF", 6, MKI$(NPROV&))
         For U& = 1 To ULTREG&("!PORCCOEF")
            X$ = REGLEIDO$("!PORCCOEF", U&)
            Call REGAPP("PORCCOEF", X$)
         Next U&
       Else
         Call TRACENTRAL("PORCCOEF", "")
       End If

      Call FRESHECHO("PORCCOEF")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")


End Sub

'''Private Sub mnuDemora_Click()
'''      '
'''      NPROV& = 0
'''      Call SELECHO("PROVED1")
'''      '
'''      NPROV& = XVALO(VALACT1("PROVED1"))
'''      'SI TIENE PROVEEDOR SELECCIONADO FILTRA
'''      Call TRALOCAL("DEMORAS", "")
'''      If NPROV& > 0 Then
'''            RFF$ = RECFILT$("!DEMORAS", 1, MKI$(NPROV&))
'''            FIN& = Len(RFF$)
'''            For UK& = 1 To FIN& Step 4
'''               Call TRACE(20, UK&, FIN&)
'''               U& = CVS(Mid$(RFF$, UK&, 4))
'''               X$ = REGLEIDO$("!DEMORAS", U&)
'''               j& = j& + 1
'''               Call GRAREG("!DEMORAS", X$, j&)
'''            Next UK&
'''            Call SETULTREG("!DEMORAS", j&)
'''      Else  'SI NO SE SELECCIONA PRESENTA TODOS
'''           'ORDENA POR PROVEEDOR
'''           Call TRACENTRAL("DEMORAS", "")
'''           Call FILESORT("DEMORAS", "", 1, 1)
'''           '
'''           Call TRALOCAL("DEMORAS", "")
'''           '
'''10         Screen.MousePointer = 1
'''      End If
'''      '
'''      VTB% = VENTABU%("DEMORAS/TITUPAN=Demoras Según Proveedor y Via de Entrega")
'''      If VTB% < 0 Then Exit Sub
'''      '
'''      Screen.MousePointer = 11
'''      For U& = 1 To ULTREG&("!DEMORAS")
'''         X$ = REGLEIDO$("!DEMORAS", U&)
'''         COCO% = Asc(Mid$(X$, 3, 2))
'''         If COCO% < 1 Then
'''            Call MENSERR(24, "Código debe Vía de Entrega Debe ser Numérico No Válido en Fila: " & U&)
'''            GoTo 10
'''         End If
'''         '
'''      Next U&
'''      '
'''      For U& = 1 To ULTREG&("!DEMORAS")
'''         X$ = REGLEIDO$("!DEMORAS", U&)
'''         NP% = Asc(Mid$(X$, 1, 2))
'''         If NP% < 1 Then
'''            Call MENSERR(24, "Código de Proveedor No Válido en Fila: " & U&)
'''            GoTo 10
'''         End If
'''         '
'''      Next U&
'''
'''      'VALIDO QUE NO SE REPITA LA COMBINACION DE CODIGO PROVEEDOR Y VIA DE ENTREGA
'''      For U& = 1 To ULTREG&("!DEMORAS")
'''         X$ = REGLEIDO$("!DEMORAS", U&)
'''         COCO% = CVI(Left$(X$, 2))
'''         CODEMORA% = CVI(Mid$(X$, 3, 2))
'''         For j& = 1 To ULTREG&("!DEMORAS")
'''            If j& <> U& Then
'''              X$ = REGLEIDO$("!DEMORAS", j&)
'''              COCO2% = Asc(Left$(X$, 2))
'''              CODEMORA2% = Asc(Mid$(X$, 3, 2))
'''              If COCO% = COCO2% Then
'''                If CODEMORA% = CODEMORA2% Then
'''                   Call MENSERR(24, "Código Proveedor (" & COCO2% & ") y Código Demora  (" & CODEMORA% & ") Repetido.")
'''                   GoTo 10
'''                End If
'''              End If
'''            End If
'''          Next j&
'''      Next U&
'''      '
'''      'SI SELECCIONO PROVEEDOR BORRO LOS REGISTROS PARA LUEGO AGREGARLOS AL FINAL
'''      If NPROV& > 0 Then
'''         Call FILTERDELETE("DEMORAS", 1, MKI$(NPROV&))
'''         For U& = 1 To ULTREG&("!DEMORAS")
'''            X$ = REGLEIDO$("!DEMORAS", U&)
'''            Call REGAPP("DEMORAS", X$)
'''         Next U&
'''      Else
'''         Call TRACENTRAL("DEMORAS", "")
'''      End If
'''
'''      Call FRESHECHO("DEMORAS")
'''      '
'''      Screen.MousePointer = 1
'''      Call COMUNI("Grabación Finalizada")
'''      '
'''End Sub
'''
'''

Private Sub mnuDemora_Click()
      '
      
      NPROV& = 0
      Call SELECHO("PROVED1")
      '
      NPROV& = XVALO(VALACT1("PROVED1"))
      'SI TIENE PROVEEDOR SELECCIONADO FILTRA
      Call TRALOCAL("DEMORAS", "")
      If NPROV& > 0 Then
            RFF$ = RECFILT$("!DEMORAS", 1, MKI$(NPROV&))
            FIN& = Len(RFF$)
            For UK& = 1 To FIN& Step 4
               Call TRACE(20, UK&, FIN&)
               U& = CVS(Mid$(RFF$, UK&, 4))
               X$ = REGLEIDO$("!DEMORAS", U&)
               J& = J& + 1
               Call GRAREG("!DEMORAS", X$, J&)
            Next UK&
            Call SETULTREG("!DEMORAS", J&)
      Else  'SI NO SE SELECCIONA PRESENTA TODOS
           'ORDENA POR PROVEEDOR
           Call TRACENTRAL("DEMORAS", "")
           Call FILESORT("DEMORAS", "", 1, 1)
           '
           Call TRALOCAL("DEMORAS", "")
           '
10         Screen.MousePointer = 1
      End If
      '
      VTB% = VENTABU%("DEMORAS/TITUPAN=Demoras Según Proveedor y Via de Entrega")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = Asc(Mid$(X$, 3, 2))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe Vía de Entrega Debe ser Numérico No Válido en Fila: " & U&)
            GoTo 10
         End If
         '
      Next U&
      '
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         NP% = Asc(Mid$(X$, 1, 2))
         If NP% < 1 Then
            Call MENSERR(24, "Código de Proveedor No Válido en Fila: " & U&)
            GoTo 10
         End If
         '
      Next U&

      'VALIDO QUE NO SE REPITA LA COMBINACION DE CODIGO PROVEEDOR Y VIA DE ENTREGA
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = CVI(Left$(X$, 2))
         CODEMORA% = CVI(Mid$(X$, 3, 2))
         For J& = 1 To ULTREG&("!DEMORAS")
            If J& <> U& Then
              X$ = REGLEIDO$("!DEMORAS", J&)
              COCO2% = Asc(Left$(X$, 2))
              CODEMORA2% = Asc(Mid$(X$, 3, 2))
              If COCO% = COCO2% Then
                If CODEMORA% = CODEMORA2% Then
                   Call MENSERR(24, "Código Proveedor (" & COCO2% & ") y Código Demora  (" & CODEMORA% & ") Repetido.")
                   GoTo 10
                End If
              End If
            End If
          Next J&
      Next U&
      '
      'SI SELECCIONO PROVEEDOR BORRO LOS REGISTROS PARA LUEGO AGREGARLOS AL FINAL
      If NPROV& > 0 Then
         Call FILTERDELETE("DEMORAS", 1, MKI$(NPROV&))
         For U& = 1 To ULTREG&("!DEMORAS")
            X$ = REGLEIDO$("!DEMORAS", U&)
            Call REGAPP("DEMORAS", X$)
         Next U&
      Else
         Call TRACENTRAL("DEMORAS", "")
      End If

      Call FRESHECHO("DEMORAS")
      '
      Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
      '
End Sub


 



Private Sub MNUEDITPRESUP_Click()
    '
    CONDI$ = "NroPres > 0  ORDER BY NroPres DESC"
10  Call CARGALISEL("INDICEPR", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEPR/INDICE DE PRESUPUESTOS")
    NPR& = XVALO(VALACT1("INDICEPR"))
    If NPR& < 1 Then Exit Sub
    
    '
    Text17 = TRIM(Str$(NPR&))
    Text1 = XVALO(Left$(VALACT2("INDICEPR"), 6))
    Text20 = "" 'para que siempre dispare el evento change
    Text1.Locked = True
    Command19.Enabled = False
    '
    Call PREPORCOP(NPR&)
    CONDI$ = "NroPres = " & NPR&
    NROPE& = XVALO(VALOBADA("PRESUVEN", "NroPedi", CONDI$, "NOMESS"))
    If NROPE& < 1 Then
       Text21.TEXT = "MODIFICACIÓN DE PRESUPUESTO "
    Else
       Text21.TEXT = "MODIFICACIÓN DE PRESUPUESTO - PEDIDO: " & NROPE&
    End If
    
    Text21.Visible = True
    Text21.Width = 7000
    Command9.Visible = True

    '
End Sub
Function SIG_VERSION%(PRESUPUESTOX$)
       'PR-00000063-00
       NUMEPRE& = XVALO(Mid$(PRESUPUESTOX$, 4, 8))
       CONDI$ = "NroPres = " & NUMEPRE&
       SIG_VERSION% = 1 + VALOBADA("PRESUVENTA", "MAX(RIGHT(NroPres_Version,2))", CONDI$, "NOMESS")
End Function

Private Sub MnuLimvias_Click()
      Call TRALOCAL("KILOVIAS", "")
      '
10    Screen.MousePointer = 1

      VTB% = VENTABU%("KILOVIAS")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!KILOVIAS")
         X$ = REGLEIDO$("!KILOVIAS", U&)
         COCO% = Asc(Left$(X$, 2))
         If COCO% < 1 Then
            Call MENSERR(24, "Falta Código de Vía de Entrega")
            GoTo 10
         End If
         '
      Next U&
      '
      'VALIDO QUE NO SE REPITA LA COMBINACION DE CODIGO PROVEEDOR Y VIA DE ENTREGA
      For U& = 1 To ULTREG&("!KILOVIAS")
         X$ = REGLEIDO$("!KILOVIAS", U&)
         COCO% = Asc(Left$(X$, 2))
         For J& = 1 To ULTREG&("!KILOVIAS")
            If J& <> U& Then
              X$ = REGLEIDO$("!KILOVIAS", J&)
              COCO2% = Asc(Left$(X$, 2))
              If COCO% = COCO2% Then
                 Call MENSERR(24, "Código Via Repetido (" & COCO2% & ") ")
                 GoTo 10
              End If
            End If
          Next J&
      Next U&
      '
      Call TRACENTRAL("KILOVIAS", "")
      Call FRESHECHO("KILOVIAS")
      '
      Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
      '
End Sub

Private Sub mnupcopia_Click()

'   If XVALO(Text20) > 0 Then Exit Sub ' modo modificacion
    CONDI$ = "NroPres > 0 ORDER BY NroPres DESC"
10  Call CARGALISEL("INDICEPR", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEPR/INDICE DE PRESUPUESTOS")
    NPR& = XVALO(VALACT1("INDICEPR"))
    If NPR& < 1 Then Exit Sub
    
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% < 1 Then
      GoTo 10
    End If
    '
    Text1 = NCLIE%
    Call PREPORCOP(XVALO(NPR&))
   '
End Sub
Sub PREPORCOP(NPRESU&)   '
   NUCLI% = NCLIE% 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
   CODICLI1$ = Label26
   Call ABREPRESVEN
   '
   Call VectProveedor.cargarvector
   Call VectArticulos.cargarvector
   
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres= " & NPRESU&
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt "
   '
   Dim PRESUTEMP As ADODB.Recordset
   Set PRESUTEMP = READSET(SQLX$)
   '
   MSF1.Rows = 1
   Screen.MousePointer = 11
   COTIMONEY = 1
   With PRESUTEMP
      Do While Not .EOF
       i& = i& + 1
        If i& = 1 Then ' EN EL CASO QUE SE HAYA GRABADO EN PESOS LO PRESENTA PRIMERO EN DOLARES CON EL TIPO DE CAMBIO QUE EN EL CUAL SE GRABO
          MONEMI% = XVALO(!MonedEmi) ' LUEGO AL FINAL LO LLEVA A PESOS NUEVAMENTE. ESTO LO SOLICITARON PARA QUE SIEMPRE SE MANTENGA EL VALOR EN DOLARES
          If MONEMI% = 1 Then        ' EJEMPLO SI SE GRABA EN PESOS A UN COSTO DE 10 U$S TIPO DE CAMBIO $9
               If MONEMI% = 1 Then   'CUANDO SE MUESTRA EN $ VA A MOSTRAR 90 PERO SI SE PASA A DOLARES Y EL MISMO CAMBIO DE $9 A $12
                 MON% = MONEMI%      'SE VA A MOSTRAR QUE EL COSTO EN DOLARES ES (90 / 12) = 7.50
                 COTIMONEY = XVALO(!COTIZAMONEDA)
                 MONEMI% = 2
               End If
          End If
          TXTREFERENCIA = SAFETEXT$(!Referen)
          Text18 = SAFETEXT$(!Mail_Cli)

        End If
        CI1% = XVALO(!CodiInt)
        If CI1% > 0 Then
           'MODO NORMAL
           DESCRIX$ = SAFETEXT$(!Descrip)
           CODE$ = SAFETEXT$(!CodiExt)
           CODIGOCLIENTE$ = SAFETEXT$(!CodigoDelCliente)
           CANTI = XVALO(!Cantit)
           INETO = XVALO(!PreUnit) / COTIMONEY
           PREUNI# = INETO * (1 + TTT1 / 100)
           IMPORTX# = CANTI * PREUNI#
           CONTACTO$ = TRIM$(SAFETEXT$(!AteSen))

           MSF1.Rows = MSF1.Rows + 1
           MSF1.TextMatrix(i&, 1) = CODE$
           MSF1.TextMatrix(i&, 2) = CODIGOCLIENTE$
           MSF1.TextMatrix(i&, 3) = DESCRIX$
        
           MSF1.TextMatrix(i&, 4) = TRIM$(FORMATNUM$(CANTI, "F8.1"))
           MSF1.TextMatrix(i&, 5) = TRIM$(XVALO(!Num_Prov))
           'MSF1.TextMatrix(i&, 6) = TRIM$(rasocli(-1 * XVALO(!Num_Prov)))
           MSF1.TextMatrix(i&, 6) = VectProveedor.nombre(XVALO(!Num_Prov))
'           MSF1.TextMatrix(I&, 7) = TRIM$(FORMATNUM$(XVALO(!COSTO_UNIT) / COTIMONEY, "F12.2"))
           MSF1.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(XVALO(!COSTO_UNIT), "F12.2"))  'PARA QUE SIEMPRE APAREZCA EN LA MONEDA DOLAR QUE SERIA EN LA CUAL SE GUARDO
           MSF1.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(PREUNI#, "F12.2"))
           MSF1.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(IMPORTX#, "F12.2"))
           MSF1.TextMatrix(i&, 10) = SAFETEXT$(!Via_Entrega)
           MSF1.TextMatrix(i&, 11) = SAFETEXT$(!LOTES_Reservados)
           MSF1.TextMatrix(i&, 12) = FORMATNUM$(XVALO(!Canti_Lotes), "F7.1")
           PESOUNITAR = VectArticulos.PesoUn(CI1%)
           MSF1.TextMatrix(i&, 14) = PESOUNITAR * CANTI
           
           MSF1.TextMatrix(i&, 15) = DEMORA_X_VIA(XVALO(!Via_Entrega), XVALO(!Num_Prov))
           MSF1.TextMatrix(i&, 16) = FORMATNUM$(XVALO(!COSTOFOB), "F10.2")
           If i& = 1 Then PORDESCU = SAFETEXT$(!PorcDesc)

        Else
           ' OBSERVACION EN DETALLE
           DESCRIX$ = SAFETEXT$(!Especif)
           CODE$ = "------------"
           CODIGOCLIENTE$ = "------------"
           MSF1.Rows = MSF1.Rows + 1
           MSF1.TextMatrix(i&, 1) = CODE$
           MSF1.TextMatrix(i&, 2) = CODIGOCLIENTE$
           MSF1.TextMatrix(i&, 3) = DESCRIX$
        
           MSF1.TextMatrix(i&, 4) = TRIM$(FORMATNUM$(CANTI, "F8.1"))
           MSF1.TextMatrix(i&, 5) = ""
           MSF1.TextMatrix(i&, 6) = ""
           MSF1.TextMatrix(i&, 7) = ""
           MSF1.TextMatrix(i&, 8) = ""
           MSF1.TextMatrix(i&, 9) = ""
           MSF1.TextMatrix(i&, 10) = ""
           MSF1.TextMatrix(i&, 11) = ""
           MSF1.TextMatrix(i&, 12) = ""
           MSF1.TextMatrix(i&, 14) = ""
           MSF1.TextMatrix(i&, 15) = ""
           MSF1.TextMatrix(i&, 16) = ""

           'PORDESCU = ""

        End If
        
        

      .MoveNext
      Loop
   End With
   PRESUTEMP.Close
   Set PRESUTEMP = Nothing
   '
   'SI EL CONTACTO ES DIFERNTE AL CONTACTO PRINCIPAL DEL MAESTRO SE FIJA SI COINCIDE CON LOS DE LA LISTA DEL MAESTRO.
   Text35 = CONTACTO$
   If UCase$(CONTACTO$) <> TRIM$(UCase$(CONTACTCLI$(XVALO(Text1)))) Then
      CONDI$ = "CODICLI =  '" & CODICLI1$ & "'"
      Text5 = VALOBADA("CONTACSCOM", "CELU_CTAC", CONDI$, "NOMESS")
      'Text18 = VALOBADA("CONTACSCOM", "MAIL_CTAC", CONDI$, "NOMESS")
   End If
   '
    MONEMIS = DEMONECO$(MONEMI%)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    'CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
   '
85 If MON% = 1 Then
     Call MONEMIS_DblClick 'BUSCA DE LLEGAR A $ , DE ESTE MODO ACTUALIZA EL VALOR EN $ SI EL DOLAR CAMBIO LA COTIZACION , PERO EL COSTO EN DOLARES SIEMPRE SE MANTIENE.
     If MONEMI% <> MON% Then GoTo 85
   End If
   
   Screen.MousePointer = 1
   Call ACTU_TOTAL
   Call ACTUPESO
   ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
   Set articulos = Nothing
   Set Proveedor = Nothing
End Sub

Private Sub MnuProv_Art_Click()
         '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         '
         SQLX$ = "SELECT * FROM RELACODPROV  WITH(NOLOCK)"
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO% & " AND MARBORRA = 0 "
         Set rs = READSET(SQLX$)
         '
         Call SETULTREG("!ARTPROVD", 0)
         With rs
            On Error Resume Next
            .MoveFirst
            If Err Then
                On Error GoTo 0
                GoTo 10
            End If
            JJ& = 0
             Do While (.EOF = False)
                X$ = REGBLAN$("ARTPROVD")
                Q = !VIA
                Call REPLA(X$, MKI(XVALO(!cod_inte)), 1, 2)
                Call REPLA(X$, MKD(XVALO(!Precio)), 3, 8)
                Call REPLA(X$, Chr$(XVALO(!Moneda)), 11, 1)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 12, 2)
                JJ& = JJ& + 1
                Call GRAREG("!ARTPROVD", X$, JJ&)
            .MoveNext
            Loop
         End With
         rs.Close
         Set rs = Nothing
         '
10       XYX% = VENTABU%("ARTPROVD")
         If XYX% < 0 Then GoTo 5
         '
         CONDI$ = "NUM_PROV =" & NPRO%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&)
    
         
         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO%
         Set RS1 = New ADODB.Recordset
         RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With RS1
             For i& = 1 To ULTREG&("!ARTPROVD")
               XX$ = REGLEIDO$("!ARTPROVD", i&)
               
               CI% = CVI(Mid$(XX$, 1, 2))
               PRECIOX# = CVD(Mid$(XX$, 3, 8))
               MON% = Asc(Mid$(XX$, 11, 1))
               FE% = CVI(Mid$(XX$, 12, 2))
               CI% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CODIEMPR = CodiEmp%
               On Error GoTo 0
               !Num_Prov = NPRO%
               !cod_inte = CI%
               !Precio = PRECIOX#
               !Moneda = MON%
               If FE% < 1 Then FE% = FE% = HOY(HO$)
               If ROUND(PRECIOX#, 2) <> ROUND(XVALO(!Precio)) Then FE% = HOY(HO$)
               !FECHA = CVDAT(FE%)
               !VIA = 0
               !MARBORRA = 0
               .Update
             Next i&
           End With
         RS1.Close
         Set RS1 = Nothing
         CONDI$ = CONDI + "AND  MARBORRA = 1  "
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&)
         
         Call TRACENTRAL("ARTPROVD", "")
         Call FRESHECHO("ARTPROVD")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")

End Sub

Private Sub MnuTabEntr_Click()
      Call TRALOCAL("VIASENTR", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("VIASENTR")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!VIASENTR")
         X$ = REGLEIDO$("!VIASENTR", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("VIASENTR", "")
      Call FRESHECHO("VIASENTR")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
      'POSICIONES
      'DE PORCENTAJE 35,8
     '
End Sub

Private Sub mnutcoef_Click()
      Call TRALOCAL("TIPOCOEF", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("TIPOCOEF")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!TIPOCOEF")
         X$ = REGLEIDO$("!TIPOCOEF", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("TIPOCOEF", "")
      Call FRESHECHO("TIPOCOEF")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")

End Sub

Private Sub mnuvisualiza_Click()

    CONDI$ = "NroPres > 0 ORDER BY NroPres DESC"
10  Call CARGALISEL("INDICEPR", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEPR/INDICE DE PRESUPUESTOS")
    NPR& = XVALO(VALACT1("INDICEPR"))
    If NPR& < 1 Then Exit Sub
    
    TIDOCUM$ = "Presupuesto "
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "COTIZACION"

    Screen.MousePointer = 11
    NUDOCU$ = SAFETEXT$(NPR&)
    DOCUNU& = DOCID&(TIDOCUM$, Str$(NPR&), NUDOCU$)
      If DOCUNU& < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("No se Encontró el Documento Impreso")
      End If
      Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPR&), Str$(NPR&))
      '
      GoTo 10
      '
99
End Sub


Private Sub MONEMIS_DblClick()
   '
   TICAMBIS.Visible = False
   'VALIDO QUE NO HAYA NINGUN DATO EN LOS CASILLEROS DE ARRIBA ANTES DE MODIFICAR LA MONEDA X Q ESTA YA ESTA CONVERTIDA A LA MONEDA ACTIVA.
   If TRIM$(txtcodigo) <> "" Then
      Call COMUNI("Imposible Cambiar de Moneda Cuando Hay Información en Casilleros de Ingreso\Borre la Información y Vuelva a Intentar")
      Exit Sub
   End If
   
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   '
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     'TICAMBIS.Enabled = False
   End If
   '
   Total(0) = FORMATNUM$(XVALO(Total(0)) * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%), "F12.2")
   '
      COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
      For U& = 1 To MSF1.Rows - 1

        CAN = XVALO(MSF1.TextMatrix(U&, 3))
        'PORDESCUIN = CVS(Mid$(X$, 69, 4))
'        COSTOX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 7))
        COSTOX# = 1 * XVALO(MSF1.TextMatrix(U&, 7)) 'QUIERE QUE SE MANTENGA SIEMPRE EL COSTO EN U$s
        PRECVENTA# = COEFCON * XVALO(MSF1.TextMatrix(U&, 8))
        IMPORTEX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 9))
        MSF1.TextMatrix(U&, 7) = FORMATNUM$(COSTOX#, "F12.2")
        MSF1.TextMatrix(U&, 8) = FORMATNUM$(PRECVENTA#, "F12.2")
        MSF1.TextMatrix(U&, 9) = FORMATNUM$(IMPORTEX#, "F12.2")
        
      Next U&
      '
      Call ACTU_TOTAL
      INFOSIAHP.Text3 = MONEMIS.Caption
      INFOSIAHP.Text3.ForeColor = MONEMIS.ForeColor
'      Call CARLISPRES
'      Call CIERRARCH("!CARDETPR")
'      Call CALTOPRES
'   End If
   '
End Sub




Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.row = i&
    MSF1.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '
    '
    OPX% = COMALTER%("Artículo Seleccionado: " & MSF1.TextMatrix(REG&, 3) & " \\Cancelar\Quitar\Mover Registros")
    If OPX% < 2 Then
       Exit Sub
    ElseIf OPX% = 2 Then
       REGI% = REG&
       Call borrafila(REGI%, FRMPRESUAHP, MSF1)
       Call ACTUPESO
    ElseIf OPX% = 3 Then
       MSF1.COL = 0
'       If MSF1.CellBackColor = vbRed Or MSF1.CellBackColor = vbYellow Then
'          Call COMUNI("Item Seleccionado No Posee Stocks, Para Editar Elimine y Vuelva a Cargar")
'          Exit Sub
'       End If
       'CAMPOS$ = "Código/A22;Cod.Cte/A22;Descripción/A22;Q/F4.A;Prov/AO;Proveedor/A12;Costo (U$S)/F10.2;Pr.Venta/F8.2;Importe/F10.2;Via Apr./A10;LOTES/A0;CANTI/F0;NPROV/F0;KILOS/A0;DEMORA/F7;"
       Campos$ = "Código/A22;Cod.Cte/A22;Descripción/A22;Q/F4.A;Prov/AO;Proveedor/A12;Costo (U$S)/F10.2;Pr.Venta/F8.2;Importe/F10.2;Via Apr./A10;LOTES/A0;CANTI/F0;NPROV/F0;KILOS/f12.2;DEMORA/F7;ValorFob/F0"
       Call INSERTARFILA(MSF1, i&, FRMPRESUAHP, Campos$)
    End If
    '
    Call ACTU_TOTAL

    Call Command6_Click ' BOTON ACTUALIZAR
'    Call Grabar_Presupuesto(OK%, 1) 'GRABA EL PRESUPUESTO
    '
    Screen.MousePointer = 1
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    '
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.row - VISIROWS Then
          MSF1.TopRow = MSF1.row - VISIROWS
       End If
    End If


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
        'VALIDA SI HAY CANTIDAD YA SELECCIONADA EN EL PRESUPUESTO + LA SELECCIONADA SI ALCANZA EL STOCKS
        CANTIGRI = CANTILOTE_EN_GRILLA(CI1%, LOTE$) ' CANTIDAD YA INGRESADA EN GRILLA DE PRESUPUESTO DEL MISMO LOTE
        If CANTIGRI > 0 Then
           If (CANTI + CANTIGRI) > ST Then
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
                  CANTI = XVALO(.MSF1.TextMatrix(i&, 11))
                  ACUCANTI = ACUCANTI + CANTI
           End If
        Next i&
    End With
    '
    CANTILOTE_EN_GRILLA = ACUCANTI
    '
End Function


Function CARGADO_ENGRILLA(LOTEPARAM$, Optional Fila&)
    'RECORRO LA GRILLA DEL PRESUPUESTO Y ACUMULO LA CANTIDAD YA CARGADA
    ACUCANTI = 0
    With FRMPRESUAHP
        LOTEPARAM$ = TRIM$(LOTEPARAM$)
        For i& = 1 To .MSF1.Rows - 1
           If Fila& = i& Then GoTo 50 'SI VIENE CON VALOR DE FILA NO SUMA LA FILA PARAMETRO
           STRINGLOTES$ = MSF1.TextMatrix(i&, 11)
           If STRINGLOTES$ = LOTEPARAM$ Then
              CANTI = XVALO(MSF1.TextMatrix(i&, 4))
              ACUCANTI = ACUCANTI + CANTI
           End If
50      Next i&
    End With
    '
    CARGADO_ENGRILLA = ACUCANTI
    '
End Function

Function ITEM_CARGADO_ENGRILLA(CI1%, MSF As msFlexGrid, Fila&, Optional ULTIMA&)
    FIL& = Fila&
    ULTIM& = ULTIMA&
    'RECORRO LA GRILLA DEL PRESUPUESTO Y ACUMULO LA CANTIDAD YA CARGADA
    'ULTIMA DEBE SER SIEMPRE + 1 POR QUE SI SE PASA EL REGISTRO 1 NO VA A ENTRAR TOTAL LA VA A REBOTAR EL SEGUNDO IF.
    ACUCANTI = 0
    If ULTIM& > 0 And ULTIM& = FIL& Then ULTIM& = ULTIM& + 1
    With MSF1
        For i& = 1 To .Rows - 1
           If ULTIM& > i& Or i& = 1 Then GoTo 90
           If FIL& = i& Then GoTo 50 'SI VIENE CON VALOR DE FILA NO SUMA LA FILA PARAMETRO
           CANTI = XVALO(.TextMatrix(i&, 4))
           ACUCANTI = ACUCANTI + CANTI
50      Next i&
    End With
    '
90  ITEM_CARGADO_ENGRILLA = ACUCANTI
    '
End Function


Private Sub MSFlexGrid2_Click()

End Sub

'Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'     If Button = 2 Then
'        REG& = MSF1.MouseRow
'        If REG& < 1 Then Exit Sub
'        CI1% = CODINT%(MSF1.TextMatrix(REG&, 1))
'        Call PRESENTA_INFO_ARTICULO(CI1%)
'        Call cargar_datos
'     End If
'
'End Sub

Private Sub PORDESCU_Change()
   Call ACTU_TOTAL
End Sub

Private Sub Text1_Change()
    
   Call LIMPIAR_TEXT_DETALLE
   Call LIMPIARPRESU
   MSF1.Rows = 1
   
   If TRIM$(Text1.TEXT) <> "" Then
      Call CARGADATOS_CLIENTE
   End If
   NC% = XVALO(Text1)
   'VALIDO SI LA LISTA ACTIVA ES VALIDA
   If XVALO(Text1) > 0 Then
    NLIST% = XVALO(Text23)
   End If
   Label26 = CODICLI$(NC%)
   '
   'FECHA
   DTPicker1.Value = Date
   For KKI% = 0 To 1
     LENTREGA(KKI%) = LUENTRECLI$(NC%, KKI%)
   Next KKI%
   '
  FRMPRESUAHP.TXTCONDIPAGO = TRIM$(Str$(CPAGCLI%(NC%)))
  FRMPRESUAHP.TXTNUMVENDEDOR = TRIM$(Str$(VENDCLI%(NC%)))
  On Error Resume Next
  Text1.SetFocus
  On Error GoTo 0
'   If Command1.Enabled = True Then
'       Call FRMPRESUAHP.Command1_Click
'   End If
End Sub

Sub LIMPIARPRESU() 'SE LE PUEDE PASAR UN TEXTBOX QUE ESTE NO SE LIMPIE
    
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text8 = ""
    Text18 = ""
    ENTRAOBS.Text28 = ""
    Text14 = ""
    Text19 = ""
    Text23 = ""
    Text20 = ""
    Text35 = ""
    TXTREFERENCIA = ""
    '
    Total(0) = ""
    Total(1) = ""
    Text31 = ""
    Text33 = ""
    PORDESCU = ""
    TXTTOTALPESO = ""
    MONEMI% = 2
    MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    MONEMIS.ForeColor = RGB(0, 0, 255)
    Text21.Visible = False: Text21.Width = 7000
    
'    TICAMBIS.Visible = False

End Sub
Sub CARGADATOS_CLIENTE()
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then: Call LIMPIARPRESU: Exit Sub
   '
   CONDI$ = "Nume_Cli = " & CUENTA&
   '
   'VALIDO SI EXISTE
   CNRG& = CantRegistros&("DEUDOR12", CONDI$)
   If CNRG& < 1 Then Call LIMPIARPRESU: Exit Sub
   '
   MARCONOTA.Visible = False
   'CARGO DATOS DE CLIENTE
   Text2 = VectCliente.nombre(CUENTA&)
   Text3 = VectCliente.Domicilio(CUENTA&)
   Text4 = VectCliente.localidad(CUENTA&)
   Text8 = VectCliente.TeleCli(CUENTA&)
   Text5 = VectCliente.CeluCli(CUENTA&)
   Text18 = VectCliente.mailcliente(CUENTA&)
   Text19 = VectCliente.NotaCli(CUENTA&)
   Text23 = VectCliente.LpreCli(CUENTA&)
   Text35 = VectCliente.CtacCli(CUENTA&)
   
   'Text2 = VALOBADA("DEUDOR12", "Raso_Cli", CONDI$, "NOMESS")
   'Text3 = VALOBADA("DEUDOR12", "Domi_Cli", CONDI$, "NOMESS")
'   Text4 = VALOBADA("DEUDOR12", "LOCA_CLI", CONDI$, "NOMESS")
'   Text8 = VALOBADA("DEUDOR12", "Tele_Cli", CONDI$, "NOMESS")
'   Text5 = VALOBADA("DEUDOR12", "CELU_CLI", CONDI$, "NOMESS")
'   Text18 = VALOBADA("DEUDOR12", "Mail_Cli", CONDI$, "NOMESS")
'   Text19 = VALOBADA("DEUDOR12", "Nota_Cli", CONDI$, "NOMESS")
'   Text23 = CEROIGUALVACIO(XVALO(VALOBADA("DEUDOR12", "Lpre_Cli", CONDI$, "NOMESS")))
'   Text35 = CONTACTCLI$(CUENTA&)
   '
    
    
    If REGICLI(CUENTA&) > 0 Then
       textodelaNota$ = ANOTACLI(CUENTA&)
       If TRIM$(textodelaNota$) <> "" Then
            TEXT13 = textodelaNota$
            MARCONOTA.Visible = True
            txtcodigo.Locked = True
            TXTCANTIDAD.Locked = True
       End If
    End If

End Sub

Private Sub Text20_Change()
  If XVALO(Text20) > 0 Then
     Text21.Visible = True
     Command9.Visible = True
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\GOLDEN.skn")
  Else
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
     Text21.Visible = False
     Command9.Visible = False
  End If
End Sub


Private Sub Text21_DblClick()
   'ME FIJO SI ESTA VISIBLE PERO EN ANCHO REDUCIDO SE VA PARA QUE NO LIMPIE  LA INFO
   If Text21.Width < 5000 Then Exit Sub
   
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
   'Text20 = "-00"

   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   If Command9.Visible = True Then Command9.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   MSF1.Rows = 1
   On Error Resume Next
   Unload INFOSIAHP
   On Error GoTo 0
   Call ACTU_TOTAL
   PORDESCU = ""
   
   
End Sub


Private Sub Text23_Change()
'   CODI% = XVALO(Text23)
'   Text22 = ECOLPRECIOS$(CODI%)

End Sub

Private Sub TICAMBIS_DblClick()
   If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
      ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If
   End If
   '

End Sub


Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub TXTCANTIDAD_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")
   '
End Sub

Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        '
   TXX1TRAZA$ = "TXTCANTIDAD_KeyPress Inicio: (txtcodigo) " & Now & vbCrLf
        'VALIDO EL CODIGO
        CI1% = CODINT%(txtcodigo)
   TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress POST CODINT : " & Now & vbCrLf
        
        If CI1% < 1 Then
           CONDI$ = "Codigo_Cliente = '" & TRIM$(txtcodigo) & "'"
           TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress PRE EXISTECODNUE% : " & Now & vbCrLf
           EXISTECODNUE% = CantRegistros&("CODICLIENTE", CONDI$, "NOMESS")
           TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress POST EXISTECODNUE% : " & Now & vbCrLf

           If EXISTECODNUE% < 1 Then
             Call COMUNI("CÓDIGO INEXISTENTE O NO VÁLIDO")
             Exit Sub
           End If
        End If
        'VALIDA REPETICION DE CODIGO
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress PRE VALIDAREPECODI% : " & Now & vbCrLf

        Linea% = VALIDAREPECODI%(CODINT%(txtcodigo))
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress POST VALIDAREPECODI% : " & Now & vbCrLf
        
        If Linea% > 0 Then
           Call COMUNI("Código Repetido en Fila: " & Linea%)
'           txtcodigo.SetFocus
'           Exit Sub
        End If

        'VALIDO QUE TENGA INGRESADA CANTIDAD
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           Call COMUNI("FALTA CANTIDAD")
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        '
        'VERIFICA SI HAY STOCKS
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress PRE DISPONIB_LIBRE : " & Now & vbCrLf
        
        If DISPONIB_LIBRE(CI1%) < 1 Or INFOSIAHP.PRESENTA_COTIZADOR% > 0 Then
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress POST DISPONIB_LIBRE : " & Now & vbCrLf
        
          INFOSIAHP.PRESENTA_COTIZADOR% = 0 '24/05/21 ESTO ESTABA COMENTADO CAMBIO REALIZADO DEBIDO A QUE EN ALGUNOS CASOS PRESENTA EL COTIZADOR SIN STOCK CUANDO REALMENTE HAY
          Call ACTIVA_OBJ_SIN_STOCK         'ESTOY SUPONIENDO SIN PODERLO REPRODUCIR QUE ESTA VARIABLE GLOBAL A NIVEL DE FRM QUEDE CON VALOR = 1
'          Call CARGARCOMBOVIAS_TODOS
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress PRE TXTDESCRIPCION_Change : " & Now & vbCrLf

          Call INFOSIAHP.TXTDESCRIPCION_Change
TXX1TRAZA$ = TXX1TRAZA$ + "TXTCANTIDAD_KeyPress POST TXTDESCRIPCION_Change : " & Now & vbCrLf
          
          Call PRESENTA_INFO_ARTICULO(CI1%, 0, 1)
          'Call PRESENTA_INFO_ARTICULO(CI1%, INFOSIAHP.FILAX, 1)
          If INFOSIAHP.PRESENTA_COTIZADOR% > 0 Then
             INFOSIAHP.PRESENTA_COTIZADOR% = 0
             TXTCANTIDAD.SetFocus
          End If
          
          If INFOSIAHP.MODO1$ = "CANCEL" Then
            INFOSIAHP.MODO1$ = ""
            Exit Sub ' CASO QUE CIERRE EL FRM INFOSIAHP
          End If
          
'          Call cargar_datos_sin_stock
          txtcodigo = ""
          txtcodigo.SetFocus
          Call LIMPIAR_TEXT_DETALLE
          If XVALO(MSF1.Rows - 1) Mod 10 = 0 Then
            Command13.Caption = "******************"
            Call Grabar_Presupuesto(OK%)
            Command13.Caption = "Grabar"
            If OK% > 0 Then SIGRABO.Caption = "SI"
          End If
          Call POSICIONAR_ULTIMAFILA(MSF1)
          Exit Sub
        End If
        '
        If CARGA(2) = False Then
           TXTPREUNIT.Locked = False
           TXTPREUNIT.BackColor = &H80000005
           TXTPREUNIT.SetFocus
           Exit Sub
        Else
          'OPCION DE CARGA CON SELECCION MANUAL DE LOTES
          Call ACTIVA_OBJ_CON_STOCK
          
          INFOSIAHP.Frame1.Top = 4560
          INFOSIAHP.Height = 9135
          INFOSIAHP.Text6(3).Visible = False
          INFOSIAHP.TXTNUMEPROVE.Visible = False
          INFOSIAHP.Command5.Visible = False
          INFOSIAHP.Text1.Visible = False
          INFOSIAHP.Command3.Visible = False
          Call INFOSIAHP.TXTDESCRIPCION_Change
          Call PRESENTA_INFO_ARTICULO(CI1%)
          If INFOSIAHP.MODO1$ = "CANCEL" Then
            INFOSIAHP.MODO1$ = ""
            Exit Sub ' CASO QUE CIERRE EL FRM INFOSIAHP
          End If
          If MSF1.Rows - 1 Mod 10 = 0 Then
            Command13.Caption = "******************"
            Call Grabar_Presupuesto(OK%)
            Command13.Caption = "Grabar"
            If OK% > 0 Then SIGRABO.Caption = "SI"
          End If
          Call POSICIONAR_ULTIMAFILA(MSF1)
          On Error Resume Next
          On Error GoTo 0
        End If
        '
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        
        If CARGA(0).Value = True Then Call cargar_datos   ' CARGA RAPIDA
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
        '
        'CIERRO EL FORMULARIO DE DETALLE DE STOCKS
        Unload INFOSIAHP
        '

    End If


    

End Sub
Sub POSICIONAR_ULTIMAFILA(MSF As msFlexGrid)
     With MSF
        If .Rows > 1 Then
          .row = MSF.Rows - 1
          .TopRow = MSF.row
          .RowSel = MSF.row
          .COL = 0
          .ColSel = .Cols - 1
        End If
     End With
'codigo anterior
'On Error Resume Next
'    VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
'    If Err > 0 Then
'      VISIROWS = 1
'    End If
'    On Error GoTo 0
'    '
'    If VISIROWS > 0 Then
'       If MSF.TopRow < MSF.Row - VISIROWS Then
'          MSF.TopRow = (MSF.Row - VISIROWS)
'       End If
'    End If
End Sub
Private Sub cargar_datos_ANTERIOR()
    '
    
'    MSF1.Rows = MSF1.Rows + 1
'    MSF1.Row = MSF1.Rows - 1
'    CI1% = CODINT%(txtcodigo.TEXT)
'    CANTI = XVALO(TXTCANTIDAD)
'    If CARGA(1) = True Then 'VALIDA STOCKS
'        ST = FORMATNUM$(STOCKACT(CI1%), "F8.0")
'        CANPED = FORMATNUM$(CANT_PEDIDA_PENDIENTE(CI1%), "F8.0")      ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
'        COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F8.0")  ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
'        CANRECEP = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(0, CI1%), "F8.0")
'        COLO$ = COLORSTATUS$(CANTI, ST, COMPRAX + CANRECEP, CANPED)
'        '
'        MSF1.TextMatrix(MSF1.Rows - 1, 0) = COLO$
'        If COLO$ = "V" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbGreen, MSF1.Rows - 1, 0)
'        If COLO$ = "A" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbYellow, MSF1.Rows - 1, 0)
'        If COLO$ = "R" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbRed, FRMPRESUAHP.MSF1.Rows - 1, 0)
'    End If
'    '
'    MSF1.TextMatrix(MSF1.Rows - 1, 1) = txtcodigo.TEXT
'    MSF1.TextMatrix(MSF1.Rows - 1, 2) = INFOSIAHP.TXTCODICLIENTE.TEXT
'    MSF1.TextMatrix(MSF1.Rows - 1, 3) = DESCRIT0$(CI1%)
'    MSF1.TextMatrix(MSF1.Rows - 1, 4) = FORMATNUM$(TXTCANTIDAD.TEXT, "F10.1")
'    MSF1.TextMatrix(MSF1.Rows - 1, 5) = INFOSIAHP.TXTCOSTO
'    MSF1.TextMatrix(MSF1.Rows - 1, 5) = INFOSIAHP.TXTNUMEPROVE
'    MSF1.TextMatrix(MSF1.Rows - 1, 6) = INFOSIAHP.TXTPROVEEDOR
'    If CARGA(2) = True Then 'CASO QUE ES CARGA DE SELECCION MANUAL DE LOTES TRAE LA INFO DE INFOSIAHP
'        MSF1.TextMatrix(MSF1.Rows - 1, 7) = FORMATNUM$(INFOSIAHP.TXTCOSTO.TEXT, "f12.2")
'        MSF1.TextMatrix(MSF1.Rows - 1, 8) = FORMATNUM$(INFOSIAHP.TXTPRECIOVENTA, "f12.2")
'        MSF1.TextMatrix(MSF1.Rows - 1, 9) = FORMATNUM$(XVALO(INFOSIAHP.TXTPRECIOVENTA) * XVALO(TXTCANTIDAD), "f14.2")
'        MSF1.TextMatrix(MSF1.Rows - 1, 11) = CARGALOTES$
'        MSF1.TextMatrix(REG&, 12) = CANTI_ACUMU_LOTES ' FUNCION
'        CANTILOTES = XVALO(MSF1.TextMatrix(REG&, 12))
'        CANTISOLICI = XVALO(MSF1.TextMatrix(REG&, 4))
'        If CANTILOTES <= CANTISOLICI Then COLO$ = "A"
'        If CANTILOTES = 0 Then COLO$ = "R"
'        If CANTILOTES >= CANTISOLICI Then COLO$ = "V"
'        MSF1.TextMatrix(REG&, 0) = COLO$
'
'    Else
'        MSF1.TextMatrix(MSF1.Rows - 1, 7) = FORMATNUM$(Text15.TEXT, "f12.2")
'        MSF1.TextMatrix(MSF1.Rows - 1, 8) = FORMATNUM$(TXTPREUNIT, "f12.2")
'        MSF1.TextMatrix(MSF1.Rows - 1, 9) = FORMATNUM$(XVALO(TXTPREUNIT) * XVALO(TXTCANTIDAD), "f14.2")
'    End If
'    MSF1.TextMatrix(MSF1.Rows - 1, 10) = Text38
'
''    MSF1.TextMatrix(MSF1.Rows - 1, 6) = "0" ' INDICADOR DE CODIGO OPCIONAL
''    If Check1.Value = 1 Then
''        MSF1.TextMatrix(MSF1.Rows - 1, 6) = "1"
''        MSF1.TextMatrix(MSF1.Rows - 1, 5) = ""
''    End If
''    MSF1.ColWidth(6) = 0 ' OCULTA LA COLUMNA
''    Check1.Value = 0 ' SETEA NO OPCIONAL POR DEFAULT
'    '
'    ' ORDENA GRILLA POR COLUMNA 6 Y EN FORMA ASCENDENTE (VALOR 7 DE SORT DATOS DE TIPO STRING)
''    MSF1.Col = 6
''    MSF1.Sort = 7
'    '
'    'SUMO EL IMPORTE
'    Call ACTU_TOTAL
'    '
'    On Error Resume Next
'    VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
'    If Err > 0 Then
'      VISIROWS = 1
'    End If
'    On Error GoTo 0
'    If VISIROWS > 0 Then
'       If MSF1.TopRow < MSF1.Row - VISIROWS Then
'          MSF1.TopRow = MSF1.Row - VISIROWS
'       End If
'    End If
'    '
'    'desplazo verticalmente a la ultima columna
'    'MSF1.LeftCol = MSF1.Cols - 1
End Sub
 Sub cargar_datos_sin_stock()
    '
    CI1% = CODINT%(txtcodigo.TEXT)
    If INFOSIAHP.MODO1$ = "CANCEL" Then
      INFOSIAHP.MODO1$ = ""
      Exit Sub
    End If
    REG& = MSF1.Rows
    '
    If INFOSIAHP.FILAX& > 0 Then
       MSF1.row = INFOSIAHP.FILAX&
       REG& = INFOSIAHP.FILAX&
       TXTCANTIDAD = INFOSIAHP.TXTSOLICITADA
       MSF1.TextMatrix(REG&, 0) = FILAX&
       INFOSIAHP.FILAX& = 0
    Else
       MSF1.Rows = MSF1.Rows + 1
       'TXTCANTIDAD = INFOSIAHP.TXTSOLICITADA
       MSF1.TextMatrix(REG&, 0) = REG&
    End If

    CANTI = XVALO(TXTCANTIDAD)
    '
    MSF1.TextMatrix(REG&, 1) = txtcodigo.TEXT
    MSF1.TextMatrix(REG&, 2) = INFOSIAHP.TXTCODICLIENTE.TEXT
    MSF1.TextMatrix(REG&, 3) = DESCRIT0$(CI1%)
    MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(TXTCANTIDAD), "F10.1")
      NPROV = XVALO(INFOSIAHP.TXTNUMEPROVE)
    MSF1.TextMatrix(REG&, 5) = NPROV
    MSF1.TextMatrix(REG&, 6) = TEL2CLI$(-1 * NPROV)
    MSF1.TextMatrix(REG&, 7) = FORMATNUM$(Text15.TEXT, "f12.2")
    MSF1.TextMatrix(REG&, 8) = FORMATNUM$(TXTPREUNIT, "f12.2")
    MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(TXTPREUNIT) * XVALO(TXTCANTIDAD), "F14.2")
'    MSF1.TextMatrix(REG&, 0) = ""
    MSF1.TextMatrix(REG&, 10) = Text38
    MSF1.TextMatrix(REG&, 14) = XVALO(TXTCANTIDAD) * XVALO(INFOSIAHP.TXTPESO)
    Call ACTUPESO
    '
    'SUMO EL IMPORTE
    Call ACTU_TOTAL
    '
    '
    On Error Resume Next
    TPXI& = MSF1.Rows - 11: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0    '
    
    'desplazo verticalmente a la ultima columna
    'MSF1.LeftCol = MSF1.Cols - 1
End Sub
 Sub cargar_datos()
    '
    CI1% = CODINT%(txtcodigo.TEXT)
    If INFOSIAHP.MODO1$ = "CANCEL" Then
      INFOSIAHP.MODO1$ = ""
      Exit Sub
    End If
    REG& = MSF1.Rows
    '
    If INFOSIAHP.FILAX& > 0 Then
       MSF1.row = INFOSIAHP.FILAX&
       REG& = INFOSIAHP.FILAX&
       TXTCANTIDAD = INFOSIAHP.TXTSOLICITADA
       MSF1.TextMatrix(REG&, 0) = FILAX&
       INFOSIAHP.FILAX& = 0
    Else
       MSF1.Rows = MSF1.Rows + 1
       'TXTCANTIDAD = INFOSIAHP.TXTSOLICITADA
       MSF1.TextMatrix(REG&, 0) = REG&
    End If

    CANTI = XVALO(TXTCANTIDAD)
    '
    MSF1.TextMatrix(REG&, 1) = txtcodigo.TEXT
    MSF1.TextMatrix(REG&, 2) = INFOSIAHP.TXTCODICLIENTE.TEXT
    MSF1.TextMatrix(REG&, 3) = DESCRIT0$(CI1%)
    MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(TXTCANTIDAD), "F10.1")
'      NPROV = PROVELOTE%(LOTE$)
'    MSF1.TextMatrix(REG&, 5) = NPROV
'    MSF1.TextMatrix(REG&, 6) = TEL2CLI$(-1 * NPROV)
    MSF1.TextMatrix(REG&, 7) = FORMATNUM$(Text15.TEXT, "f12.2")
    MSF1.TextMatrix(REG&, 8) = FORMATNUM$(TXTPREUNIT, "f12.2")
    MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(TXTPREUNIT) * XVALO(TXTCANTIDAD), "F14.2")
'    MSF1.TextMatrix(REG&, 0) = ""
    MSF1.TextMatrix(REG&, 10) = Text38
    MSF1.TextMatrix(REG&, 14) = XVALO(TXTCANTIDAD) * XVALO(INFOSIAHP.TXTPESO)
    Call ACTUPESO
    '
    'SUMO EL IMPORTE
    Call ACTU_TOTAL
    '
    '
    On Error Resume Next
    TPXI& = MSF1.Rows - 11: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0
    
    'desplazo verticalmente a la ultima columna
'    MSF1.LeftCol = MSF1.Cols - (MSF1.Cols - 1)
End Sub

Private Sub cargar_datos_respaldo()
    '
''    MSF1.Rows = MSF1.Rows + 1
''    MSF1.Row = MSF1.Rows - 1
'    CI1% = CODINT%(txtcodigo.TEXT)
'    If INFOSIAHP.MODO1$ = "CANCEL" Then
'      INFOSIAHP.MODO1$ = ""
'      Exit Sub
'    End If
'    REG& = MSF1.Rows
'
'
'    If INFOSIAHP.FILAX& > 0 Then
'       MSF1.Row = INFOSIAHP.FILAX&
'       REG& = INFOSIAHP.FILAX&
'       TXTCANTIDAD = INFOSIAHP.Text55
'       MSF1.TextMatrix(REG&, 0) = FILAX&
'       INFOSIAHP.FILAX& = 0
'    Else
'       MSF1.Rows = MSF1.Rows + 1
'       MSF1.TextMatrix(REG&, 0) = REG&
'    End If
'
'    CANTI = XVALO(TXTCANTIDAD)
'    If CARGA(1) = True Then 'VALIDA STOCKS
''        ST = FORMATNUM$(STOCKACT(CI1%), "F8.0")
''        CANPED = FORMATNUM$(CANT_PEDIDA_PENDIENTE(CI1%), "F8.0")      ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
''        COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F8.0")  ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
''        CANRECEP = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(0, CI1%), "F8.0")
''        COLO$ = COLORSTATUS$(CANTI, ST, COMPRAX + CANRECEP, CANPED)
''        '
''        MSF1.TextMatrix(REG&, 0) = COLO$
''        If COLO$ = "V" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbGreen, REG&, 0)
''        If COLO$ = "A" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbYellow, REG&, 0)
''        If COLO$ = "R" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbRed, REG&, 0)
'    End If
'    '
'
'    MSF1.TextMatrix(REG&, 1) = txtcodigo.TEXT
'    MSF1.TextMatrix(REG&, 2) = INFOSIAHP.TXTCODICLIENTE.TEXT
'    MSF1.TextMatrix(REG&, 3) = DESCRIT0$(CI1%)
'    MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(TXTCANTIDAD), "F10.1")
'    'MSF1.TextMatrix(REG&, 5) = INFOSIAHP.TXTCOSTO
'    MSF1.TextMatrix(REG&, 5) = INFOSIAHP.TXTNUMEPROVE
'    MSF1.TextMatrix(REG&, 6) = INFOSIAHP.TXTPROVEEDOR
'    If CARGA(2).Value = True Then 'CASO QUE ES CARGA DE SELECCION MANUAL DE LOTES TRAE LA INFO DE INFOSIAHP
'        MSF1.TextMatrix(REG&, 7) = FORMATNUM$(INFOSIAHP.TXTCOSTO.TEXT, "f12.2")
'        MSF1.TextMatrix(REG&, 8) = FORMATNUM$(INFOSIAHP.TXTPRECIOVENTA, "f12.2")
'        MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(INFOSIAHP.TXTPRECIOVENTA) * XVALO(TXTCANTIDAD), "f14.2")
'        MSF1.TextMatrix(REG&, 11) = CARGALOTES$  ' FUNCION
'        MSF1.TextMatrix(REG&, 12) = CANTI_ACUMU_LOTES  ' FUNCION
'        'Call CANTI_SEGUN_LOTES_SELECCIONADOS(CARGALOTES$) 'ACA HABRIA QUE DECODIFICAR LA CANTIDAD DE LOTES SELECCIONADOS
'        ' **** Y QUE HACER EL COLORSTATUS PARA VER QUE COLOR DEVUELVE, EN ESTE CASO NO SE DEBERIA CONTEMPLAR LA CANTIDAD EN TRANSITO.
'        'VALIDA EL COLOR
''        Call COLOREARSTOCKS(REG&)
''        COLO$ = COLORSTATUS$(TXTCANTIDAD, XVALO(MSF1.TextMatrix(REG&, 12)), 0, 0)
''        CANTILOTES = XVALO(MSF1.TextMatrix(REG&, 12))
''        CANTISOLICI = XVALO(MSF1.TextMatrix(REG&, 4))
''        If CANTILOTES <= CANTISOLICI Then COLO$ = "A"
''        If CANTILOTES = 0 Then COLO$ = "R"
''        If CANTILOTES >= CANTISOLICI Then COLO$ = "V"
''        MSF1.TextMatrix(REG&, 0) = COLO$
'        '
'    Else
'        MSF1.TextMatrix(REG&, 7) = FORMATNUM$(Text15.TEXT, "f12.2")
'        MSF1.TextMatrix(REG&, 8) = FORMATNUM$(TXTPREUNIT, "f12.2")
'        MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(TXTPREUNIT) * XVALO(TXTCANTIDAD), "f14.2")
'        MSF1.TextMatrix(REG&, 0) = ""
'
'    End If
'    MSF1.TextMatrix(REG&, 10) = Text38
'    MSF1.TextMatrix(REG&, 13) = TXTPESOP
'    MSF1.TextMatrix(REG&, 14) = INFOSIAHP.TXTDIASDEMORA
'    Call ACTUPESO
'    '
'    'SUMO EL IMPORTE
'    Call ACTU_TOTAL
'    '
'    If CARGA(0).Value <> True Then
'       Call COLOREARSTOCKS(REG&)
'    End If
'    '
'    On Error Resume Next
'    VISIROWS = (MSF1.Height / MSF1.RowHeight(j&)) - 4
'    If Err > 0 Then
'      VISIROWS = 1
'    End If
'    On Error GoTo 0
'    If VISIROWS > 0 Then
'       If MSF1.TopRow < MSF1.Row - VISIROWS Then
'          MSF1.TopRow = MSF1.Row - VISIROWS
'       End If
'    End If
'    '
'    'desplazo verticalmente a la ultima columna
'    'MSF1.LeftCol = MSF1.Cols - 1
End Sub


Function CANTI_ACUMU_LOTES()
   'SUMA LA CANTIDAD SELECCIONADA ACUMULADA DE LOTES
   '
   CANTI_ACUMU_LOTES = 0
   With INFOSIAHP
      For i& = 1 To .MSF5.Rows - 1
        LOTE$ = TRIM$(.MSF5.TextMatrix(i&, 1))
        CANTI = XVALO(.MSF5.TextMatrix(i&, 5))
        If LOTE$ <> "" Then
          If CANTI > 0 Then
            CANTI_ACUMU_LOTES = CANTI_ACUMU_LOTES + CANTI
          End If
        End If
      Next
   End With
   '
End Function

Function CARGALOTES$()
   'CARGA LOS LOTES SELECCIONADOS PARA EL ITEM - SOLO MODO DE CARGA MANUAL
   With INFOSIAHP
      For i& = 1 To .MSF5.Rows - 1
        LOTE$ = TRIM$(.MSF5.TextMatrix(i&, 1))
        CANTIX$ = SAFETEXT$(XVALO(.MSF5.TextMatrix(i&, 5)))
'        If LOTE$ <> "" Then
          If XVALO(CANTIX$) > 0 Then
            CARGALOTES$ = CARGALOTES$ & LOTE$ & ";" & CANTIX$ & ";"
          End If
'        End If
      Next
   End With
   '

   If InStrRev(CARGALOTES$, ";") = Len(CARGALOTES$) And CARGALOTES$ <> "" Then
      CARGALOTES$ = Left$(CARGALOTES$, Len(CARGALOTES$) - 1)
   End If
   '
   Debug.Print CARGALOTES$
   '
End Function


Sub ACTU_TOTAL()


   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   For U& = 1 To MSF1.Rows - 1
     CIXXI% = CODINT%(MSF1.TextMatrix(U&, 1))
     CAN = XVALO(MSF1.TextMatrix(U&, 4))
'     PORDESCUIN = CVS(Mid$(X$, 69, 4))
'     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = XVALO(MSF1.TextMatrix(U&, 8))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
'       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
'       Call GRAREG("!CARDETPR", X$, U&)
     End If
   Next U&
   '
   Total(0).TEXT = FORMATNUM$(SUTOTPED#, "F12.2")
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   TTTX# = SUNEFA2# * (100 - PODES) / 100
3   TTT1 = XVALO(Total(0).TEXT) * PODES / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   '
    ACUMU_IMPORTE# = 0
'    For I& = 1 To MSF1.Rows - 1
'        IMPORTE# = XVALO(MSF1.TextMatrix(I&, 5))
'        ACUMU_IMPORTE# = ACUMU_IMPORTE# + IMPORTE#
'    Next I&
    TTTX# = SUIVA# * (100 - PODES) / 100
    Text33 = FORMATNUM$(XVALO(Total(0)) - XVALO(Total(1)), "F12.2")
    Text31 = FORMATNUM$(TTTX#, "f12.2")
    Text14 = FORMATNUM$(XVALO(Text33) + XVALO(Text31), "F12.2")

End Sub

Private Sub txtcodigo_Change()

    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    Text15 = "" ' COSTO
    'TXTCODCLIENTE = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
    TXTPREUNIT.Locked = True 'DESHABILITO
    'MSF3.Rows = 1
    TXTPREUNIT.ToolTipText = ""

    If TRIM$(txtcodigo) = "" Then Exit Sub
    'VALIDO SI LA LISTA ACTIVA ES VALIDA
    NLIST% = XVALO(Text23)
    If NLIST% < 1 Or TRIM$(Text22) = "" Then
      NLIST% = 1
    End If
    '
    'BUSCO LOS DATOS
    CI1% = CODINT%(txtcodigo)
    If CI1% < 1 Then
       Exit Sub
    Else
'       TXTDESCRIPCION = DESCRIT0$(CI1%)
'       Text15 = FORMATNUM$(actualiza_costo(CI1%), "F10.2")
    End If
    '
'    INFOSIAHP.FILAX = 0
'    On Error Resume Next
'    Call PINTATEXT(TXTCANTIDAD)
'    TXTCANTIDAD.SetFocus
'    On Error GoTo 0
    
'    CONDI$ = " CODINT = " & CI1%
  
    
    
End Sub



Private Sub txtcodigo_GotFocus()
   txtcodigo.BackColor = vbYellow
End Sub


Private Sub TXTNUMVENDEDOR_Change()
    TXTNOMVENDEDOR = ""
    If XVALO(TXTNUMVENDEDOR) > 0 Then
       TXTNOMVENDEDOR = ECOARCH$("VENDEDOR", Chr(XVALO(TXTNUMVENDEDOR)))
    End If
    '
End Sub

'Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'     CI1% = CODINT(txtcodigo)
'     If CI1% < 1 Then
'        Call COMUNI("Código Inexistente o No válido")
'        Exit Sub
'     End If
'
'     If KeyAscii = 13 Then
'        TXTDESCRIPCION = DESCRIT0(CI1%)
'        TXTCANTIDAD = ""
'        'TXTPREUNIT = FORMATNUM$(PRELISTA(TRIM$(Str$(NLIST%)), CI1%), "F12.2")
'        TXTIMPORTE = ""
'        '
'        Call PRESENTA_INFO_ARTICULO(CI1%)
'        '
'        ' SI SE SELECCIONO DE LA LISTA PASA EL FOCO A LA CANTIDAD
'
'        If Command1.Enabled = False Then
'           TXTCANTIDAD.SetFocus
'        End If
'     End If
'   End If
'End Sub

Private Sub TXTPREUNIT_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")

End Sub
Sub LIMPIAR_TEXT_DETALLE()
    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    txtcodigo = ""
    TXTCODCLIENTE = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.Locked = False
    TXTPREUNIT.ToolTipText = ""
    'Text35 = ""
    Text15 = ""
    Text38 = ""
    SIGRABO.Caption = ""
End Sub

Sub TXTPREUNIT_DblClick()

    If TXTPREUNIT.Locked = False Then
       TXTPREUNIT.BackColor = &HE0E0E0
       TXTPREUNIT.Locked = True
    Else
       TXTPREUNIT.Locked = False
       TXTPREUNIT.BackColor = &H80000005
       Call PINTATEXT(TXTPREUNIT)
    End If
    
End Sub

Sub TEXT15_DblClick()

    If Text15.Locked = False Then
       Text15.BackColor = &HE0E0E0
       Text15.Locked = True
    Else
       Text15.Locked = False
       Text15.BackColor = &H80000005
       Call PINTATEXT(Text15)
    End If
    
End Sub

Private Sub TXTPREUNIT_KeyPress(KeyAscii As Integer)


    If TXTPREUNIT.Locked = True Then Exit Sub
    '
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        
        Linea% = VALIDAREPECODI%(CODINT%(txtcodigo))
        If Linea% > 0 Then
           Call COMUNI("Código Repetido en Fila: " & Linea%)
'           txtcodigo.SetFocus
'           Exit Sub
        End If

'        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
'           KeyAscii = 0
'           MsgBox "FALTA PRECIO UNITARIO"
'           TXTPREUNIT.SetFocus
'           Exit Sub
'        End If

'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        If CARGA(2).Value <> True Then
          cargar_datos
        Else
          TXTCANTIDAD_KeyPress (13) 'CASO QUE ESTE PARADO EN PREUNIT Y SELECCIONADO CARGA MANUAL
          Exit Sub
        End If
        
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
'        If Command1.Enabled = True Then
'           Call FRMPRESUAHP.Command1_Click
'        End If

     End If
    
End Sub
   
Private Sub TXTPREUNIT_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 115 'pulso  f4
        Call TXTPREUNIT_DblClick

    End Select

End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     NC = XVALO(Text1)
     If NC <= 0 Then Call COMUNI("Falta Ingresar N° de Cliente"): Call PINTATEXT(Text1): Exit Sub
     CI1% = CODINT(txtcodigo)
'     If CI1% < 1 Then
'       'SI CODINT ES < 1 VERIFICA SI EXISTE EN TABLA MASTER2 - SI EXISTE LO DA DE ALTA EN MASTER
'       If EXISMASTER2%(txtcodigo) Then Call AltaM2AMaster(txtcodigo)
'       CI1% = CODINT(txtcodigo) 'VUELVE A LEER EL CODIGO INTERNO CON CODINT
'     End If
     If CI1% < 1 Then
        CICLI& = CODICLIENTE&(txtcodigo)
        If CICLI& > 0 Then
          CI1% = CICLI&
          TXTCODCLIENTE = UCase$(txtcodigo) ' CONMUTO LA POSICION PARA QUE SIGA HABITUALMENTE
          txtcodigo.TEXT = CODEXT$(CI1%) 'CON EL CODIGO ORIGINAL
        Else
          OPX% = COMALTER("Código Inexistente o No Válido\\Dar de Alta\Cancelar")
          If OPX% = 1 Then
            If EmpresaNoRealizaCambios% > 0 Then
               Exit Sub
            End If
            'ALTA DE CODIGO Y/O RELACION
            Call ALTACODIRELA
          End If
          Exit Sub
        End If
     End If
     
     TXTDESCRIPCION = DESCRIT0(CI1%)
     TXTCANTIDAD = ""
     'TXTPREUNIT = FORMATNUM$(PRELISTA(TRIM$(Str$(NLIST%)), CI1%), "F12.2")
     TXTIMPORTE = ""
     '
     'VALIDA REPETICION DE CODIGO
     Linea% = VALIDAREPECODI%(CI1%)
     If Linea% > 0 Then
        Call COMUNI("Código Repetido en Fila: " & Linea%)
     End If
     '
     ' SI SE SELECCIONO DE LA LISTA PASA EL FOCO A LA CANTIDAD
'     If Command1.Enabled = False Then
       TXTCANTIDAD.SetFocus
'     End If
   End If
End Sub
Function actualiza_costo#(CI1%)
     'PASA ACTUALIZA EL COSTO A LA MONEDA ACTIVA.
     MONEI% = MONECOSTO%(CI1%)
     COSTOX# = COSUNI(CI1%)
     COEFCON = 1
     If MONEI% <> MONEMI% Then
        COEFCON = TICAMONE(MONEI%) / TICAMONE(MONEMI%)
     End If
     actualiza_costo# = COEFCON * XVALO(COSTOX#)
End Function
'
Sub ALTACODIRELA()
    VDEF$ = String$(101, 0)
    CODI$ = TRIM$(txtcodigo)
    Call REPLA(VDEF$, CODI$, 77, 24)
    CI_NEW = CODINT_NEW

    Call REPLA(VDEF$, MKI$(CI_NEW), 1, 2)
10  OBX$ = OBJPLA$("RELACODI", VDEF$)
    If OBX$ = "" Then Exit Sub
    
    Call ALTARELACODIGO(OBX$, OK%)
    If OK% = 1 Then
      Call COMUNI("Alta Realizada Correctamente")
    ElseIf OK% = 2 Then
      Call COMUNI("Código Relacionado Correctamente")
    Else
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    Screen.MousePointer = 1
    UPDATMASTER% = 1
    Call GENECOMAS
   
    
End Sub

Private Sub TXTREFERENCIA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)

       'Verifica las teclas Alt - Ctrl y la A
        If KeyCode = 17 Then CTRLX = True
        If KeyCode = 67 Then CE = True
        'Si las 2 variables están en True
        If CTRLX And CE Then
            REG& = MSF1.MouseRow
            If REG& < 1 Then Exit Sub
            strContenidoCelda$ = MSF1.TextMatrix(MSF1.MouseRow, MSF1.MouseCol)
            Clipboard.Clear
            Clipboard.SetText strContenidoCelda
            MSF1_KeyUp 0, 0
        End If
  
End Sub
Private Sub MSF1_KeyUp(KeyCode As Integer, Shift As Integer)
    'Cuando se produce KeyUp ponemos las variables en false
    CTRLX = False
    CE = False

End Sub


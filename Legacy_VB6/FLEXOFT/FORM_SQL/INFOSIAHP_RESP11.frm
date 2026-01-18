VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFOSIAHP_RESP11 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INFORME DE SITUACIÓN DE STOCKS DE ITEM ACTIVO"
   ClientHeight    =   10140
   ClientLeft      =   -3135
   ClientTop       =   555
   ClientWidth     =   15150
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   15150
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      Caption         =   "Selección  Por Proveedor y Lotes con Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   43
      Top             =   1920
      Width           =   15135
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
         Left            =   2280
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ingreso con Lotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13050
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2000
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Stocks y Costos x Lotes"
         Top             =   2880
         Width           =   7200
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Stocks y Costos x Lotes"
         Top             =   2880
         Width           =   7560
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Cant. Solicitada:"
         Top             =   5100
         Width           =   1560
      End
      Begin VB.TextBox Text5 
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
         Index           =   2
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1200
      End
      Begin VB.TextBox Text55 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5100
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Cant.Selecc.:"
         Top             =   5100
         Width           =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2595
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   4577
         _Version        =   393216
         BackColorFixed  =   16501405
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
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   2235
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   3942
         _Version        =   393216
         BackColorFixed  =   8421504
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
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1875
         Left            =   7440
         TabIndex        =   52
         Top             =   3120
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   3307
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15135
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
         Height          =   300
         Left            =   7680
         TabIndex        =   42
         ToolTipText     =   "Lista de Proveedores Relacionados con el Item Activo"
         Top             =   360
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ingreso sin Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13000
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2000
      End
      Begin VB.TextBox TXTNACIONALIZACION 
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
         Left            =   9960
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Gastos de Nacionalización"
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11750
         TabIndex        =   8
         ToolTipText     =   "CAMBIAR MONEDA"
         Top             =   990
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
         Height          =   380
         Index           =   1
         Left            =   11990
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Cód.Cliente"
         Top             =   990
         Width           =   1030
      End
      Begin VB.TextBox Text11 
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
         Index           =   0
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "N° Prov."
         ToolTipText     =   "Doble click - Presena Lista de Selección de todos los Proveedores."
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox TXTFECHACOSTEO 
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
         Left            =   7930
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Fecha de Costeo"
         Top             =   990
         Width           =   1575
      End
      Begin VB.TextBox TXTPESO 
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
         Left            =   13905
         TabIndex        =   32
         ToolTipText     =   "Peso de la Pieza"
         Top             =   630
         Width           =   1110
      End
      Begin VB.TextBox Text2 
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
         Left            =   13905
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Peso"
         Top             =   360
         Width           =   1110
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
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Proveedor"
         Top             =   360
         Width           =   2530
      End
      Begin VB.TextBox TXTPROVEEDOR 
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
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   630
         Width           =   2530
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
         TabIndex        =   28
         Top             =   630
         Width           =   3345
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
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   360
         Width           =   3345
      End
      Begin VB.TextBox Text4 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TXTCOSTO 
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
         Left            =   9960
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Costo "
         Top             =   630
         Width           =   1335
      End
      Begin VB.TextBox TXCOMBO6 
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
         Height          =   380
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Via de entrega segun peso"
         Top             =   630
         Width           =   1820
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
         Height          =   285
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Vìa de Entrega"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text9 
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
         Left            =   13170
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Demora"
         Top             =   360
         Width           =   785
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   360
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
         TabIndex        =   20
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox TXTDIASDEMORA 
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
         Left            =   13170
         TabIndex        =   19
         ToolTipText     =   "Demora en Dias"
         Top             =   630
         Width           =   785
      End
      Begin VB.TextBox TXTPRECIOVENTA 
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
         Left            =   11730
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "   "
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox Text8 
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
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "P.Venta"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TXTMONEDA 
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
         Height          =   380
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Tipo de Moneda"
         Top             =   630
         Width           =   460
      End
      Begin VB.TextBox Text10 
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
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Mon."
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox Text11 
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "ST"
         ToolTipText     =   "Stock Físico"
         Top             =   990
         Width           =   375
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
         Height          =   380
         Index           =   0
         Left            =   495
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   990
         Width           =   1375
      End
      Begin VB.TextBox Text11 
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
         Index           =   2
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "PI"
         ToolTipText     =   "Proyección de Ingreso"
         Top             =   990
         Width           =   385
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
         Height          =   380
         Index           =   1
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   1375
      End
      Begin VB.TextBox Text11 
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
         Index           =   3
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "PC"
         Top             =   990
         Width           =   375
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
         Height          =   380
         Index           =   2
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Pedidos de Clientes"
         Top             =   990
         Width           =   1365
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
         Left            =   13020
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   2000
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
         Left            =   7680
         Sorted          =   -1  'True
         TabIndex        =   37
         Text            =   "Combo6"
         Top             =   615
         Width           =   2295
      End
      Begin VB.TextBox TXTSIMBOLO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   380
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Descripción de Moneda"
         Top             =   990
         Width           =   460
      End
      Begin VB.TextBox TXTNUMEPROVE 
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
         Height          =   380
         Left            =   6465
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Estadísticos"
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
      TabIndex        =   2
      Top             =   7440
      Width           =   15135
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Últimas 3 Ventas"
         Top             =   360
         Width           =   7560
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
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Últimas 3 Cotizaciones"
         Top             =   360
         Width           =   7200
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1875
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   3307
         _Version        =   393216
         BackColorFixed  =   16501405
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
         Height          =   1875
         Left            =   7440
         TabIndex        =   39
         Top             =   600
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   3307
         _Version        =   393216
         BackColorFixed  =   16501405
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "INFOSIAHP_RESP11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON
Public MODO1$
Public FILAX&

Sub ACTUALIZA_INFO_SUPERIOR(Optional KILOX_NO%, Optional COSTO_NO%)
  
  NPROV& = XVALO(TXTNUMEPROVE)
  CI1% = CODINT%(txtcodigo)
  KILOX = FORMATNUM$(PESUNI(CI1%), "F12.2")
  VIAX$ = VIA_SEGUN_KG(NPROV&, KILOX)
  TXCOMBO6.TEXT = VIAX$

  VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))

  If COSTO_NO% < 1 Then TXTCOSTO.TEXT = FORMATNUM$(COSTO_X_PROV_VIA(NPROV&, 0, CI1%), "F12.2")
  TXTFECHACOSTEO = FECHATEX$(FECHACOSTO%(NPROV&, VIA%, CI1%))
  TXTMONEDA = MONEDAEMISION%(NPROV&, VIA%, CI1%)
  TXTSIMBOLO = SIMBOLPES$(XVALO(TXTMONEDA))
  If TXTPOSIVA = "LOCAL" Then
     TXTDIASDEMORA = "0"
     TXCOMBO6 = "LOCAL"
  Else
     TXTDIASDEMORA = DEMORA_X_VIA(VIA%, NPROV&)
  End If
  '
  'PARA QUE SI SE EDITA EL PESO NO LIMPIE EL CASILLERO
  If KILOX_NO% < 1 Then KILOX = FORMATNUM$(PESUNI(CI1%), "F10.4"): TXTPESO.TEXT = KILOX
  If COSTO_NO% < 1 And TXTPROVEEDOR = "Todos" Then TXTCOSTO.TEXT = COSUNI(CI1%)
'  Call CARGARCOMBOVIAS(NPROV&, KILOX)

End Sub

Sub ACTUCANTI()
   ACUMULACANTI = 0
   'If TXTMOVIL.Visible = True Then Exit Sub
   For i& = 1 To MSF5.Rows - 1
        CANTI = XVALO(MSF5.TextMatrix(i&, 5))
        ACUMULACANTI = ACUMULACANTI + CANTI
   Next i&
   Text5(2) = FORMATNUM$(ACUMULACANTI, "F10.1")
End Sub

Function COSTO_X_PROV_VIA(NPROV&, VIAX%, CI1%)
   If NPROV& < 1 Then Exit Function
   If CI1% < 1 Then Exit Function
   CONDI$ = " NUM_PROV = " & NPROV&
   If VIAX% > 0 Then CONDI$ = CONDI$ + " AND VIA = " & VIAX%
   CONDI$ = CONDI$ + " AND MARBORRA < 1 "
   CONDI$ = CONDI$ + "  AND COD_INTE = " & CI1%
   COSTO_X_PROV_VIA = XVALO(VALOBADA("RELACODPROV", "PRECIO", CONDI$, "NOMESS"))
End Function
Function FECHACOSTO%(NPROV&, VIAX%, CI1%)
   If NPROV& < 1 Then Exit Function
   If CI1% < 1 Then Exit Function
   CONDI$ = " NUM_PROV = " & NPROV&
   If VIAX% > 0 Then CONDI$ = CONDI$ + " AND VIA = " & VIAX%
   CONDI$ = CONDI$ + " AND MARBORRA < 1 "
   CONDI$ = CONDI$ + "  AND COD_INTE = " & CI1%
   FECHACOSTO% = CVINT(VALOBADA("RELACODPROV", "FECHA", CONDI$, "NOMESS"))

End Function

Sub GRABA_RESE_COTI()
       
       SQLX$ = "Select * from  RESE_COTI WHERE " + CONDI$
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
25     On Error Resume Next
       RS.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       '
       U& = 0
       User% = USNBR% Mod 100
       '
       CI1% = CODINT%(txtcodigo)
       With RS
          .AddNew
          !NroPres = FRMPRESUAHP.Text17
          !CodiInt = CI1%
          !IDENLOTE = LOTE$
          !Cantidad = CANTI
          !ESTADO = -1
          !USUARIO = User%
          .Update
       End With
       '
       RS.Close
       Set RS = Nothing
       '
End Sub

Function MONEDAEMISION%(NPROV&, VIA%, CI1%)
   If NPROV& < 1 Then
    MONEDAEMISION% = 1
    Exit Function
   End If
   
   If CI1% < 1 Then Exit Function
   CONDI$ = " NUM_PROV = " & NPROV&
   If VIAX% > 0 Then CONDI$ = CONDI$ + " AND VIA = " & VIA%
   CONDI$ = CONDI$ + " AND MARBORRA < 1 "
   CONDI$ = CONDI$ + "  AND COD_INTE = " & CI1%
   MONEDAEMISION% = XVALO(VALOBADA("RELACODPROV", "MONEDA", CONDI$, "NOMESS"))
   If MONEDAEMISION% < 1 Then MONEDAEMISION% = 1
   '
End Function

Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub
Private Sub cmdelinminar_Click()

End Sub

Private Sub Combo6_Click()
     
    TXCOMBO6.TEXT = Combo6.TEXT
End Sub

Private Sub Command1_Click()
   'ACTUALIZAR
   Command1.Enabled = False
    If TRIM$(txtcodigo.TEXT) = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 90
    End If
    '
    CANTSELEC = XVALO(Text5(2))
    NECESIDAD = XVALO(Text55)
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
    'VALIDO QUE NO SE HAYA SELECCIONADO MAS CANTIDAD DE LOS LOTES DISPONIBLES
    If FILAX& = 0 Then
      Call CONTROL_CANT_LOTE_MSF5(OK%)
      If OK% < 1 Then GoTo 99
    End If

    '
    If XVALO(Text5(2)) < 1 Then
      Call COMUNI("Falta Seleccionar Cantidad de Lote")
      GoTo 99
    End If
    '
    
    'SI LEEGA ACA ES POR QUE HAY STOCK CON LOS LOTES.
    TXTDIASDEMORA = "Inmediata"
    '
    KILOX = XVALO(TXTPESO)
    NPROV& = XVALO(TXTNUMEPROVE)
    If InStr(TXCOMBO6, "LOCAL") < 1 Then
       VIAX$ = VIA_SEGUN_KG(NPROV&, KILOX)
       If TRIM$(TXCOMBO6.TEXT) <> TRIM$(VIAX$) Then
         OPX% = COMALTER("Atención !!! \ Vía de Aprovisionamiento No Coincide con el Peso o Posición de Proveedor\\Continuar\Corregir")
         If OPX% = 2 Then GoTo 99
       End If
    End If
    '
'    If TXTPOSIVA = "LOCAL" And TRIM$(Mid$(TXCOMBO6, 4)) <> "LOCAL" Then
'         OPX% = COMALTER("Atención !!! \ Tipo de Proveedor No Conincide Con el Modo de Aprovisionamiento\\Continuar\Corregir")
'         If OPX% = 2 Then GoTo 99
'    End If

    If XVALO(TXTCOSTO) < 0.005 Then
         OPX% = COMALTER("Atención !!! \ Falta Costo Para el Item Activo\\Continuar\Corregir")
         Call PINTATEXT(TXTCOSTO)
         TXTCOSTO.SetFocus
         If OPX% = 2 Then GoTo 99
    End If
    '
    COEFCON = 1
    If MONEMI% <> XVALO(TXTMONEDA) Then
        COEFCON = TICAMONE(XVALO(TXTMONEDA)) / TICAMONE(MONEMI%)
    End If
    '
    FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
    FRMPRESUAHP.Text15 = FORMATNUM$(COEFCON * XVALO(TXTCOSTO), "F12.2")
    '
    FRMPRESUAHP.Text38 = TXCOMBO6
    FRMPRESUAHP.TXTPREUNIT = FORMATNUM$(COEFCON * XVALO(TXTPRECIOVENTA), "F12.2")
    FRMPRESUAHP.TXTCANTIDAD = TRIM$(FORMATNUM$(XVALO(Text55), "F9.1"))
    '
    FRMPRESUAHP.TXTPESOP = FORMATNUM$(XVALO(TXTPESO), "F12.4")
    MODO1$ = "ACTUALIZAR"
90  Hide
99  Command1.Enabled = True
End Sub
Private Sub Command2_Click()
  Command2.Enabled = False
  FIN& = ULTREG&("TAMONED")
  POSPANTALLA% = XVALO(TXTMONEDA)
  If XVALO(TXTMONEDA) < FIN& Then
     REGIS% = POSPANTALLA% + 1
  Else
     REGIS% = 1
  End If
  TXTMONEDA = REGIS%
  
  Command2.Enabled = True
End Sub

Private Sub COMMAND3_Click()
   Command3.Enabled = False
    If TRIM$(txtcodigo.TEXT) = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 90
    End If
    '
    CANTSELEC = XVALO(Text5(2))
    NECESIDAD = XVALO(Text55)
    If CANTSELEC > 0 Then
       OPX% = COMALTER("Atención !!! \ Se Han Seleccionado Lotes y Cantidad\Desea de Todos Modos Continuar con Carga sin Stock\\Continuar\Corregir")
       If OPX% = 2 Then GoTo 99
    End If
    '
    KILOX = XVALO(TXTPESO)
    NPROV& = XVALO(TXTNUMEPROVE)
    If InStr(TXCOMBO6, "LOCAL") < 1 Then
       VIAX$ = VIA_SEGUN_KG(NPROV&, KILOX)
       If TRIM$(TXCOMBO6.TEXT) <> TRIM$(VIAX$) Then
         OPX% = COMALTER("Atención !!! \ Vía de Aprovisionamiento No Coincide con el Peso o Posición de Proveedor\\Continuar\Corregir")
         If OPX% = 2 Then GoTo 99
       End If
    End If
    '
    If TXTPOSIVA = "LOCAL" And TRIM$(Mid$(TXCOMBO6, 4)) <> "LOCAL" Then
         OPX% = COMALTER("Atención !!! \ Tipo de Proveedor No Conincide Con el Modo de Aprovisionamiento\\Continuar\Corregir")
         If OPX% = 2 Then GoTo 99
    End If

    If XVALO(TXTCOSTO) < 0.005 Then
         OPX% = COMALTER("Atención !!! \ Falta Costo Para el Item Activo\\Continuar\Corregir")
         Call PINTATEXT(TXTCOSTO)
         TXTCOSTO.SetFocus
         If OPX% = 2 Then GoTo 99
    End If
    '
    COEFCON = 1
    If MONEMI% <> XVALO(TXTMONEDA) Then
        COEFCON = TICAMONE(XVALO(TXTMONEDA)) / TICAMONE(MONEMI%)
    End If
    '
    FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
    FRMPRESUAHP.Text15 = FORMATNUM$(COEFCON * XVALO(TXTCOSTO), "F12.2")
    
    FRMPRESUAHP.Text38 = TXCOMBO6
'    FRMPRESUAHP.TXTPREUNIT = FORMATNUM$(COEFCON * XVALO(INFOSIAHP.TXTPRECIOVENTA), "F12.2")
    
    FRMPRESUAHP.TXTPESOP = FORMATNUM$(XVALO(TXTPESO), "F12.4")
    FRMPRESUAHP.TXTCANTIDAD = XVALO(Text55)
    '
    'Call GRABA_RESE_COTI ' SE CAMBIO POR QUE GUARDA EN LA GRILLA
    
    '
    MODO1$ = "ACTUALIZAR"
90  Hide
99  Command3.Enabled = True

End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call SELECHO("PROVERELA/Proveedores Relacionados")
    NP% = XVALO(VALACT1$("PROVERELA"))
    If NP% > 0 Then
      TXTNUMEPROVE = NP%
    End If
    Call T_Espera(2) 'PARA EVITAR QUE EL DOBLE CLICK EN LA LISTA DE PROVEEDORES IMPACTE EN LA GRILLA
    On Error Resume Next
    txtcodigo.SetFocus
    On Error GoTo 0
    Command5.Enabled = True

End Sub


Private Sub Form_Activate()
    On Error Resume Next
    MSF1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     If Cancel = 0 Then
        MODO1$ = "CANCEL"
        INFOSIAHP.FILAX = 0
     End If
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "CASO DE TODOS X FIFO " & vbCrLf
       TXTO$ = TXTO$ + "TRAE DE TODOS LOS PROVEEDORES ORDENADOS POR FIFO" + vbCrLf
       TXTO$ = TXTO$ + "(PUEDE INCLUIR LOTES DE PROVEEDORES QUE NO SEAN LOS RELACIONADOS)" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub

Private Sub MSF5_Click()
          If MSF5.COL <> 5 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF5.COL = 5
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF5.CellWidth
            .Left = MSF5.CellLeft + MSF5.Left
            .Top = MSF5.CellTop + MSF5.Top
            .TEXT = MSF5.TEXT
            .Visible = True
            .ToolTipText = MSF5.Row
            .SetFocus
          End With
          Call ACTUCANTI
         
         'SI HAY UN SOLO REGISTRO PONE EL COSTO EN EL CASILLERO DE ARRIBA Y EL  GASTO DE NACIONALIZACION


End Sub

Private Sub MSF5_DblClick()
    i& = MSF5.MouseRow
    J& = MSF5.MouseCol
    If i& = 0 Or i& > MSF5.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF5.Row = i&
    MSF5.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '

    OPX% = COMALTER%("Realmente Desea Quitar la Fila Seleccionada\\Cancelar\Quitar")
    If OPX% = 2 Then
        REGI% = REG&
        'VUELVO A PINTAR DE BLANCO LA FILA CON EL MISMO LOTE QUE SE HA QUITADO
        LOTE5$ = TRIM$(MSF5.TextMatrix(REG&, 1))
        For i& = 1 To MSF4.Rows - 1
            LOTE4$ = TRIM$(MSF4.TextMatrix(i&, 1))
            If LOTE4$ = LOTE5$ Then
               Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFOSIAHP.MSF4, &HFFFFFF, i&)
               Exit For
            End If
        Next i&
        '
        
        REGI% = MSF5.Row
        If REGI% > 0 Then
           Call borrafila(REGI%, INFOSIAHP, INFOSIAHP.MSF5)
        End If
        
        MSF4.COL = 1
        MSF4.Sort = 1

        Call ACTUCANTI
    End If
         TXTNACIONALIZACION = ""
         TXTCOSTO = ""
         If MSF5.Rows <= 2 Then
          COSTOX = XVALO(MSF5.TextMatrix(1, 3)) 'REGISTRO 1 POR QUE SOLO VA A LEER CUANDO HAYA UN SOLO REGISTRO
          GSTNAC = XVALO(MSF5.TextMatrix(1, 4))

          TXTNACIONALIZACION = GSTNAC
          TXTCOSTO = COSTOX
         End If

End Sub

Private Sub MSF5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        MSF5_Click
   End If
End Sub

Private Sub MSF5_SelChange()
   Call ACTUCANTI
'   If MSF5.Col = 5 Then
'     MSF5.Col = 5
'     MSF5_Click
'
'   End If
End Sub

Private Sub Text11_DblClick(Index As Integer)
         Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         TXTNUMEPROVE = NPRO%
         Call T_Espera(2) 'PARA EVITAR QUE EL DOBLE CLICK EN LA LISTA DE PROVEEDORES IMPACTE EN LA GRILLA
         On Error Resume Next
         txtcodigo.SetFocus
         On Error GoTo 0

End Sub


Private Sub Text55_DblClick()
    If Text55.Locked = False Then
       Text55.BackColor = &HE0E0E0
       Text55.Locked = True
    Else
       Text55.Locked = False
       Text55.BackColor = &H80000005
       Call PINTATEXT(Text55)
    End If

End Sub


Private Sub txtcodigo_Change()
    CI1% = CODINT%(txtcodigo)
    CONDI$ = "COD_INTE = " & CI1%
    Call CARGALISEL("PROVERELA", "RELACODPROV", "Num_Prov/F9+PROVED12/A48", CONDI$, "NOMESS")
    '
    TXTPESO = FORMATNUM$(PESUNI(CI1%), "F10.4")
    '
End Sub

Private Sub TXTDIASDEMORA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "LA INFORMACION DEL TIEMPO DE DEMORA TRAE LA INFO DE LA TABLA DEMORAS" & vbCrLf
       TXTO$ = TXTO$ + "CONSIDERANDO EL PROVEEDOR Y LA VIA DE ENTREGA" + vbCrLf
       TXTO$ = TXTO$ + "SI SE TOMA CANTIDADES CON LOTES IMPRIME AQUI LA LEYENDA 'INMEDIATA' " + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "SI NO ESTA EN LA TABLA NO ESCRIBE NADA TOMA EL VALOR QUE SE ELIJA MANUALMENTE." + vbCrLf
       TXTO$ = TXTO$ + "CASO CONTRARIO TRAE EL COSTO COMBINACION ARTICULO PROVEEDOR" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub TXTMOVIL_Click()
    
'     MSF5.TextMatrix(MSF5.Row, MSF5.Col) = TXTMOVIL.TEXT
'     TXTMOVIL.Visible = False
End Sub

Sub ULTIMOS3PRESUPUESTOS(MSF As MSFlexGrid, CI1&, NC)
     MSF.Rows = 1
     SQLX$ = "SELECT TOP 3 NroPres, RasoClie, PreUnit,NroClie,MonedEmi,COTIZAMONEDA, FechEmis"
     SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CODIINT =  " & CI1&
     If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND NroClie =  " & NC
     SQLX$ = SQLX$ + " ORDER BY FechEmis  "
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        J& = J& + 1
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = XVALO(!NroPres)
        MSF.TextMatrix(J&, 2) = rasocli$(XVALO(!NroClie))
        MON% = XVALO(!MonedEmi)
        PRECIOUNITARIO# = XVALO(!PreUnit) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZAMONEDA) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TIPOCAMBIODEPRESUP# ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               MON% = 2
           End If
        End If
        MSF.TextMatrix(J&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        FE% = CVINT(!FechEmis)
        MSF.TextMatrix(J&, 5) = FECHATEX$(FE%)
        .MoveNext
       Loop
     End With

End Sub


Sub ULTIMOS3VENTAS(MSF As MSFlexGrid, CI1&, NC)

     MSF.Rows = 1
     SQLX$ = "SELECT TOP 3 DOSECLAR, VALOUNIT,Nume_Clie"
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE COD_INTE =  " & CI1&
     SQLX$ = SQLX$ + " AND  Nume_Clie =  " & NC
     SQLX$ = SQLX$ + " AND CODIMOVI = 'RT'"
     SQLX$ = SQLX$ + " ORDER BY DOSECLAR DESC "
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        J& = J& + 1
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = SAFETEXT$(!DOSECLAR)
        MSF.TextMatrix(J&, 2) = rasocli$(XVALO(!Nume_Clie))
        MSF.TextMatrix(J&, 3) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
        '
        MON% = XVALO(!Mone_Emis)
        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZA_DOLAR) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TIPOCAMBIODEPRESUP# ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               MON% = 2
           End If
        End If
        MSF.TextMatrix(J&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        '
        .MoveNext
       Loop
     End With
     
End Sub
Private Sub Form_Load()
    '
    '
    'Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
    Call APLICACIONSKINS(Me)
    ENCA$ = "Lote/A16;Stock/F9.1;Fob/F10.2;Gast.Nac./F10.2;Fecha/D8"
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF4, ENCA$, INFOSIAHP)
    ENCA$ = "Lote/A16;Stock/F9.1;Fob/F10.2;Gast.Nac./F10.2;Cant.Sol/F7.0;LOTES$/A0"
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF5, ENCA$, INFOSIAHP)
    

 End Sub
   
Private Sub MSF1_DblClick()
   
   If (MSF1.MouseRow = 0) Then Exit Sub
   REG& = MSF1.Row
   If MSF1.TextMatrix(REG&, 1) = "TOTALES:" Then Exit Sub
   '
   CI1% = CODINT%(txtcodigo)
   NPROV& = XVALO(MSF1.TextMatrix(REG&, 7))
   TXTNUMEPROVE = SAFETEXT$(NPROV&)
   TXTPROVEEDOR = MSF1.TextMatrix(REG&, 1)
   'TXTCOSTO = MSF1.TextMatrix(REG&, 4)
   KILOX = FORMATNUM$(PESUNI(CI1%), "F12.2")
   '
   AUTOFIFO% = 0
   If NPROV& = 0 And TRIM$(MSF1.TextMatrix(REG&, 1)) = "Todos x AutoFifo" Then AUTOFIFO% = 1
      
   MSF4.Rows = 1
   '
     i& = 0
   'CARGA  DE LOTE INICIAL
   If NPROV& = 0 And AUTOFIFO% < 1 Then
   SQLX$ = "SELECT SUM(CANTINGRE)AS INGRE, SUM(CANTSALID) AS SALID,IDENLOTE,FECHMOV FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE LEFT(IDENLOTE,2) = 'LI' "
   SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
   SQLX$ = SQLX$ + " GROUP BY COD_INTE,IDENLOTE, FECHMOV "
   Set RS1 = READSET(SQLX$)
   With RS1
      Do While Not .EOF
        LOTE$ = SAFETEXT$(!IDENLOTE)
          ST = XVALO(!INGRE) - XVALO(!SALID)
          RESERVA_PEDIDO = SALDO_LOTE_PEDIDO(CI1%, LOTE$)
        ST = ST - RESERVA_PEDIDO
        If ST < 0.005 Then GoTo 50
        'If ST < 0.005 Then ST = 0
        
        FE% = CVINT(!FECHMOV)
        '
       i& = i& + 1
       MSF4.Rows = i& + 1
        With INFOSIAHP.MSF4
         .TextMatrix(i&, 1) = LOTE$
         .TextMatrix(i&, 2) = FORMATNUM$(ST, "F9.1")
         .TextMatrix(i&, 3) = 0
         .TextMatrix(i&, 4) = 0
         .TextMatrix(i&, 5) = FECHATEX$(FE%)
        End With

50    .MoveNext
      Loop
   End With
   RS1.Close
'   INFOSIAHP.Show 1
   End If
   SQLX$ = " SELECT  SUM(M.CANTSALID) AS SALID , SUM(M.CANTINGRE) AS INGRE, M.IDENLOTE"
   SQLX$ = SQLX$ + " ,B.FechRecep as Fecha FROM BOLRECEP B WITH(NOLOCK)  INNER JOIN MOVISTO  M WITH(NOLOCK) ON B.IDENLOTE = M.IDENLOTE "
   SQLX$ = SQLX$ + "  Where M.COD_INTE = " & CI1%
   If NPROV& > 0 And AUTOFIFO% < 1 Then SQLX$ = SQLX$ + "  AND B.NUMPROV = " & NPROV&
   SQLX$ = SQLX$ + "  GROUP BY M.COD_INTE, B.NUMPROV,M.IDENLOTE, B.FechRecep "
   'SE ORDENA POR EL LOTE PERO DIRECTAMENTE DESDE LA GRILLA
   
   Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        ST = XVALO(!INGRE) - XVALO(!SALID)
        LOTE$ = SAFETEXT$(!IDENLOTE)
        RESERVA_PEDIDO = SALDO_LOTE_PEDIDO(CI1%, LOTE$)
        ST = ST - RESERVA_PEDIDO
        If ST < 0.005 Then GoTo 52

        FE% = CVINT(!FECHA)
        '
        i& = i& + 1
        With MSF4
        .Rows = i& + 1
        .TextMatrix(i&, 1) = LOTE$
        .TextMatrix(i&, 2) = FORMATNUM$(ST, "F9.1")
        .TextMatrix(i&, 3) = 0
        .TextMatrix(i&, 4) = 0
        .TextMatrix(i&, 5) = FECHATEX$(FE%)
        
        End With
52    .MoveNext
      Loop
    End With
    '
    RS.Close
    Set RS = Nothing
    MSF4.COL = 1
    MSF4.Sort = 1
    
    VIAX$ = VIA_SEGUN_KG(NPROV&, KILOX)
    TXCOMBO6.TEXT = VIAX$
    VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))
   'TOMO EL COSTO SEGUN EL PROVEEDOR
   If TXTPROVEEDOR = "Todos" Or XVALO(TXTNUMEPROVE) < 1 Then
      COSTOX# = COSUNI(CI1%)
      MONEX% = MONECOSTO%(CI1%)
   Else
     'SI ES UN PROVEEDOR EXISTE EN LA TABLA DE RELACION CODIGO PROVEEDOR
     COSTOX# = COSTO_X_PROV_DOLAR#(NPROV&, CI1%)
     TXTMONEDA = SAFETEXT$(2)
    End If
''''
   TXTCOSTO.TEXT = FORMATNUM$(COSTOX#, "F12.2")
   TXTFECHACOSTEO = FECHATEX$(FECHACOSTO%(NPROV&, 0, CI1%))
   'TXTPRECIOVENTA = "SEGUN FORMULA" LO CARGA EN EL EVENTO CHANGE DEL COSTO
   TXTDIASDEMORA = DEMORA_X_VIA(VIA%, NPROV&)
   TXTPESO.TEXT = KILOX
'   Call CARGARCOMBOVIAS(NPROV&, KILOX)
   Command1.Enabled = True
   Text33(0) = MSF1.TextMatrix(REG&, 2) 'STOCK
   PI = XVALO(MSF1.TextMatrix(REG&, 3)) + XVALO(MSF1.TextMatrix(REG&, 4))
   Text33(1) = SAFETEXT$(PI) 'PROYECCION DE INGRESO
   Text33(2) = MSF1.TextMatrix(1, 6) 'PEDIDO ' REGISTRO 1 POR QUE ESTA SOLO UNA VEZ EN EL PRIMER REGISTRO
''''
'''''   Call borrafila(RECALPEND09.MSF1.Row, RECALPEND09, RECALPEND09.MSF1)
'''''   MSF2.Sort = 1
End Sub

Function COSTO_X_PROV_DOLAR#(NPROV&, CI1%)

   COSTOX# = COSTO_X_PROV_VIA(NPROV&, 0, CI1%)
   MONEX% = MONEDAEMISION%(NPROV&, 0, CI1%)
   If COSTOX# > 0.005 Then
    If XVALO(MONEX%) <> 1 Then
       TIPOCA = TICAMONE(MONEX%)
       COSTOPESOS# = COSTOX# * TIPOCA
    Else
       COSTOPESOS# = COSTOX#
    End If
   End If
'''''
   COSTO_X_PROV_DOLAR# = COSTOPESOS# / TICAMONE(2)
End Function
Private Sub MSF1_SelChange()
'   Call MSF1_DblClick
End Sub

Private Sub MSF4_Click()
   If (MSF4.MouseRow = 0) Then Exit Sub
   If MSF4.Rows = 1 Then Exit Sub
   REG% = MSF4.Row
   '
   LOTE4$ = TRIM$(MSF4.TextMatrix(REG%, 1))
   
   'VALIDO SI EXISTE YA SELECCIONADA
   For i& = 1 To MSF5.Rows - 1
        LOTE5$ = TRIM$(MSF5.TextMatrix(i&, 1))
        If LOTE4$ = LOTE5$ Then
           Call COMUNI("Lote Yá Seleccionado")
           Exit Sub
        End If
   Next i&
   '
   'SI NO ESTA SELECCIONADO AGREGO A NUEVA LISTA
   MSF5.Rows = MSF5.Rows + 1
   MSF5.TextMatrix(MSF5.Rows - 1, 1) = MSF4.TextMatrix(REG%, 1)
   MSF5.TextMatrix(MSF5.Rows - 1, 2) = MSF4.TextMatrix(REG%, 2)
   MSF5.TextMatrix(MSF5.Rows - 1, 3) = MSF4.TextMatrix(REG%, 3)
   MSF5.TextMatrix(MSF5.Rows - 1, 4) = MSF4.TextMatrix(REG%, 4)
   MSF5.COL = 5
   '
   'ACTUALIZO EL GASTO DE NACIONALIZACION
   COSTOX = XVALO(MSF4.TextMatrix(REG%, 3))
   GSTNAC = XVALO(MSF4.TextMatrix(REG%, 4))
   '
   'SI HAY UN SOLO REGISTRO PONE EL COSTO EN EL CASILLERO DE ARRIBA Y EL  GASTO DE NACIONALIZACION
   TXTNACIONALIZACION = ""
   TXTCOSTO = ""
   If MSF5.Rows <= 2 Then
     TXTNACIONALIZACION = GSTNAC
     TXTCOSTO = COSTOX
   Else
     TXTCOSTO = "???"
   End If
   'PINTO LISTA SELECCIONADA
   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(INFOSIAHP.MSF4, &HC0C0FF, REG%)
   MSF5.Row = MSF5.Rows - 1: MSF5.COL = 5
   MSF5.SetFocus
   Call MSF5_Click
   
   
End Sub
Private Sub Timer1_Timer()
'    Call PINTATEXT(TXTPRECIOVENTA)
'    Timer1.Enabled = False
End Sub

Private Sub TXTLISTA_GotFocus()
   txtlista.SelStart = 0
   txtlista.SelLength = Len(txtlista.TEXT)

End Sub

'Private Sub txtlista_LostFocus()
'
'   txtlista.SelStart = 0
'   txtlista.SelLength = Len(txtlista.TEXT)
'   NN% = XVALO(ACLISPRE07.Text3)
'   If NN% = 0 Then NN% = 4
'   txtlista = FORMATNUM$(txtlista, "F14." & CStr(NN%))
'
'End Sub
Sub ACTUALIZA_FLEX(MSF As MSFlexGrid, RENGLON%, FILA%, VALOR1$)
    '
    If MSF.Row < 1 Then Exit Sub
    '
    MSF.TextMatrix(RENGLON%, FILA%) = TRIM$(VALOR1$)
    
End Sub

Function VIA_SEGUN_KG$(NPROV&, KILOX)
'     'VALIDAR SI ES LOCAL QUE TOME LA PRIMERA
     If TIPO_PROV$(NPROV) = "LOCAL" Then
         X$ = REGLEIDO$("KILOVIAS", 1)
         VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
         GoTo 50
     End If
     '
     'RECORRO LA TABLA DE KILOVIAS Y BUSCO EN QUE KILOS CUADRA
     FIN& = ULTREG&("KILOVIAS")
     For i& = 1 To FIN&
       X$ = REGLEIDO$("KILOVIAS", i&)
       MINIMOX = CVD(Mid$(X$, 3, 8))
       MAXIMOX = CVD(Mid$(X$, 11, 8))
       If XVALO(KILOX) < 0.005 Then
         VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
         Exit For
       End If
       'SI NO TIENE VALOR EN NINGUNO NO LO CONSIDERA
       If MAXIMOX < 0.005 And MINIMOX < 0.005 Then GoTo 30
       If XVALO(KILOX) >= MINIMOX And XVALO(KILOX) <= MAXIMOX Then
          VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
          Exit For
       End If
30   Next i&
     '
     'EN CASO QUE PESA + QUE EL MAXIMO, TOMA LA ULTIMA QUE LEYÓ
     If VIAX$ = "" Then
       VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
     End If
     '
     'ARMA EL STRING CON EL NUMERO DE CODIGO DE LA VIA - DESCRIPCION.
50   Call REPLA(TX$, VIAX$, 1, 2)
     REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
     Call REPLA(TX$, "-", 3, 1)
     Call REPLA(TX$, REFVIA$, 4, 12)
     '
     VIA_SEGUN_KG$ = TX$
     
End Function

''Sub CARGARCOMBOVIAS(NPROV&, KILOX)
'''     'VALIDAR SI ES LOCAL QUE TOME LA PRIMERA
''     'RECORRO LA TABLA DE KILOVIAS Y BUSCO EN QUE KILOS CUADRA
''     Combo6.Clear
''     FIN& = ULTREG&("KILOVIAS")
''     For I& = 1 To FIN&
''       X$ = REGLEIDO$("KILOVIAS", I&)
''       MINIMOX = CVD(Mid$(X$, 3, 8))
''       MAXIMOX = CVD(Mid$(X$, 11, 8))
''       'SI NO TIENE VALOR EN NINGUNO NO LO CONSIDERA
'''       If MAXIMOX < 0.005 And MINIMOX < 0.005 Then GoTo 30
''       If (XVALO(KILOX) >= MINIMOX And XVALO(KILOX) <= MAXIMOX) Or (XVALO(KILOX) = 0) Then
''         For j& = I& To FIN&
''          X$ = REGLEIDO$("KILOVIAS", j&)
''          VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
''          REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
''          Call REPLA(TX$, VIAX$, 1, 2)
''          Call REPLA(TX$, "-", 3, 1)
''          Call REPLA(TX$, REFVIA$, 4, 8)
''          Combo6.AddItem TX$
''         Next j&
''         Exit For
''       End If
''30   Next I&
''     '
''     If j& = 0 Then
''        X$ = REGLEIDO$("KILOVIAS", FIN&)
''        VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
''        REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
''        Call REPLA(TX$, VIAX$, 1, 2)
''        Call REPLA(TX$, "-", 3, 1)
''        Call REPLA(TX$, REFVIA$, 4, 8)
''        Combo6.AddItem TX$
''     End If
''
''
''End Sub
Sub CARGARCOMBOVIAS()
'     'VALIDAR SI ES LOCAL QUE TOME LA PRIMERA
     'RECORRO LA TABLA DE KILOVIAS Y BUSCO EN QUE KILOS CUADRA
     Combo6.Clear
     FIN& = ULTREG&("KILOVIAS")
'     For I& = 1 To FIN&
'       X$ = REGLEIDO$("KILOVIAS", I&)
'       MINIMOX = CVD(Mid$(X$, 3, 8))
'       MAXIMOX = CVD(Mid$(X$, 11, 8))
'       'SI NO TIENE VALOR EN NINGUNO NO LO CONSIDERA
''       If MAXIMOX < 0.005 And MINIMOX < 0.005 Then GoTo 30
'       If (XVALO(KILOX) >= MINIMOX And XVALO(KILOX) <= MAXIMOX) Or (XVALO(KILOX) = 0) Then
'         For j& = I& To FIN&
'          X$ = REGLEIDO$("KILOVIAS", j&)
'          VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
'          REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
'          Call REPLA(TX$, VIAX$, 1, 2)
'          Call REPLA(TX$, "-", 3, 1)
'          Call REPLA(TX$, REFVIA$, 4, 8)
'          Combo6.AddItem TX$
'         Next j&
'         Exit For
'       End If
'30   Next I&
     '
     If J& = 0 Then
        X$ = REGLEIDO$("KILOVIAS", FIN&)
        VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
        REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
        Call REPLA(TX$, VIAX$, 1, 2)
        Call REPLA(TX$, "-", 3, 1)
        Call REPLA(TX$, REFVIA$, 4, 8)
        Combo6.AddItem TX$
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
                  Exit For
               End If
            Next UK&
       Else
'         X$ = FILTERGETRECORD("DEMORAS", 1, MKI$(VIA%))
         DEMORAX% = 0
       End If
       DEMORA_X_VIA = DEMORAX%
End Function

Private Sub TXCOMBO6_Change()
  
   KILOX = XVALO(TXTPESO)
   NPROV& = XVALO(TXTNUMEPROVE.TEXT)
   VIAX$ = TXCOMBO6.TEXT
   VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))
   CI1% = CODINT%(txtcodigo)
'   TXTCOSTO.TEXT = FORMATNUM$(COSTO_X_PROV_VIA(NPROV&, 0, CI1%), "F12.2")
   If TXTPROVEEDOR = "Todos" Then TXTCOSTO.TEXT = COSUNI(CI1%)
   TXTDIASDEMORA = DEMORA_X_VIA(VIA%, NPROV&)
   'Call ACTUALIZA_INFO_SUPERIOR(1, 1)
   Call TXTCOSTO_Change
   '
End Sub

Private Sub TXCOMBO6_DblClick()
   TXCOMBO6.TEXT = ""
End Sub

Private Sub TXCOMBO6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
        TXTO$ = " EVENTO CHANGE ACTUALIZA LA FECHA DE COSTEO SEGUN EL Nª PROV,LA VIA " + vbCrLf
        TXTO$ = TXTO$ + " SI LA VIA DE ENTREGA ESTA VACIA TRAE LA DEL LOCAL" + vbCrLf
        TXTO$ = TXTO$ + " PRECIO DE VENTA = COSTO * PORC  (SEGUN TABLA Y VIA DE ENTREGA)" + vbCrLf
        TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
        MsgBox TXTO$
     End If
   End If

End Sub

Private Sub TXTCOSTO_Change()
   '
   TXTPRECIOVENTA.TEXT = ""
   If TRIM$(TXTCOSTO.TEXT) = "" Then Exit Sub
   VIAX$ = TXCOMBO6.TEXT
   VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))
   If VIA% < 1 Then VIA% = 1
   '
   TXTFECHACOSTEO = FECHATEX$(FECHACOSTO%(XVALO(TXTNUMEPROVE), 0, CODINT%(txtcodigo)))
   'SI ESTA VACIO CONSIDERA COMO LOCAL Y CALCULA SEGUN PORCENTAJE LOCAL
   
'   TXTPRECIOVENTA = PRECIOVENTA_X_FORMULA(XVALO(TXTCOSTO), VIA%, 1, 1)
   '
End Sub
Function PRECIOVENTA_X_FORMULA(COSTO1#, VIA%)
   PRECIOVENTA_X_FORMULA = 0
   Select Case VIA%
    Case 1
        X$ = FILTERGETRECORD("PORCCOEF", 1, MKI$(VIA%))
        PORC = CVD(Mid$(X$, 5, 8))
        PORC = 1 + (PORC / 100)
        '
        PRECIOVENTA_X_FORMULA = XVALO(TXTCOSTO) * PORC  '"SEGUN FORMULA"
        Exit Function
    Case Else
        RFF$ = RECFILT$("PORCCOEF", 1, MKI$(VIA%))
        For KKI% = 1 To Len(RFF$) Step 4
         RELI& = CVS(Mid$(RFF$, KKI%, 4))
         If RELI& > 0 Then
           ZZ$ = REGLEIDO$("PORCCOEF", RELI&)
           If CVI(Mid$(ZZ$, 3, 2)) = 2 Then
               PORCSOBREFOB = CVD(Mid$(ZZ$, 5, 8))
               If Abs(PORCSOBREFOB) < 0.005 Then
                  Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Sobre Valor Fob")
                  Exit Function
               End If
           ElseIf CVI(Mid$(ZZ$, 3, 2)) = 3 Then
               COSTOKG = CVD(Mid$(ZZ$, 5, 8))
               If Abs(COSTOKG) < 0.005 Then
                  Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Costo x Kg.")
                  Exit Function
               End If
           ElseIf CVI(Mid$(ZZ$, 3, 2)) = 4 Then
               INVERSO = CVD(Mid$(ZZ$, 5, 8))
               If Abs(INVERSO) < 0.005 Then
                  Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Inverso")
                  Exit Function
               End If
               
           End If
         End If
        Next KKI%
   End Select
   'PVTA EN USD = (FOB * 1,3150 + KG de peso * 11,0625) / 0,70
   PORCSOBREFOB = 1 + (PORCSOBREFOB / 100)
   VALOR1 = XVALO(TXTCOSTO.TEXT) * PORCSOBREFOB
   VALOR2 = COSTOKG * XVALO(TXTPESO)
   If INVERSO <> 0 Then
      PRECIOVENTA_X_FORMULA = FORMATNUM$((VALOR1 + VALOR2) / INVERSO, "F12.2")
   End If
   
End Function

Private Sub TXTCOSTO_LostFocus()
    CI1% = CODINT%(txtcodigo)
    NPROV& = XVALO(TXTNUMEPROVE)
    If XVALO(TXTNUMEPROVE) < 1 Then
       If ROUND(XVALO(TXTCOSTO), 4) <> ROUND(COSUNI(CI1%), 4) Then
         CONDI$ = "CODINT = " & CI1%
         Call ACTUREGISTRO("MASTER", "COSUNI", CONDI$, FORMATNUM$(XVALO(TXTCOSTO), "F16.5"), REGAF&, "NOMESS")
         Call ACTUREGISTRO("MASTER", "Monecos", CONDI$, XVALO(TXTMONEDA), REGAF&, "NOMESS")
         Call ACTUREGISTRO("MASTER", "FECOSTO", CONDI$, FETEXCOR1$(HOY(HO$)), REGAF&, "NOMESS")
       End If
    Else
       COSTO_ACTUAL = COSTO_SEGUN_PROV(NPROV&, CI1%)
       If ROUND(XVALO(TXTCOSTO), 4) <> ROUND(COSTO_ACTUAL, 4) Then
         CONDI$ = "NUM_PROV = " & NPROV& & " AND COD_INTE = " & CI1%
         Call ACTUREGISTRO("RELACODPROV", "PRECIO", CONDI$, FORMATNUM$(XVALO(TXTCOSTO), "F16.5"), REGAF&, "NOMESS")
         Call ACTUREGISTRO("RELACODPROV", "Moneda", CONDI$, XVALO(TXTMONEDA), REGAF&, "NOMESS")
         Call ACTUREGISTRO("RELACODPROV", "FECHA", CONDI$, FETEXCOR1$(HOY(HO$)), REGAF&, "NOMESS")
         If PROVHABI%(CI1%) = NPROV& Then
           CONDI$ = "CODINT = " & CI1%
           Call ACTUREGISTRO("MASTER", "COSUNI", CONDI$, FORMATNUM$(XVALO(TXTCOSTO), "F16.5"), REGAF&, "NOMESS")
           Call ACTUREGISTRO("MASTER", "Monecos", CONDI$, XVALO(TXTMONEDA), REGAF&, "NOMESS")
           Call ACTUREGISTRO("MASTER", "FECOSTO", CONDI$, FETEXCOR1$(HOY(HO$)), REGAF&, "NOMESS")
         End If
       End If
          
    End If
     
End Sub

Private Sub TXTCOSTO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "EL EVENTO CHANGE ACTUALIZA EL PRECIO DE VENTA" & vbCrLf
       TXTO$ = TXTO$ + "RESPETA EL RESTO DE LA INFORMACION TAL CUAL ESTA" + vbCrLf
       TXTO$ = TXTO$ + "SI SE CAMBIA OTRA INFORMACION PUEDE CAMBIAR INCLUSO EL COSTO " + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "SI EL PROVEEDOR ES TODOS TRAE EL COSTO DEL MAESTRO" + vbCrLf
       TXTO$ = TXTO$ + "CASO CONTRARIO TRAE EL COSTO COMBINACION ARTICULO PROVEEDOR" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If
End Sub


Private Sub TXTDESCRIPCION_Change()

    Call CARGA_ENCABEZADO(INFOSIAHP.MSF3, "Presup./f8.0;Cliente/A24;Proveedor/A24;Cto.Fob U$S/A24;M./A3;P.Unit./F10.2;Fecha/D8", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.msf2, "Factura./A20;Cliente/A24;P.Unit./F10.2", INFOSIAHP)
    NC = XVALO(INFOSIAHP.Text1)
    INFOSIAHP.MSF3.Rows = 1
    If TRIM$(TXTDESCRIPCION.TEXT) <> "" Then
       CI1& = CODINT(txtcodigo.TEXT)
       If CI1& > 0 Then
          Call ULTIMOS3PRESUPUESTOS(INFOSIAHP.MSF3, CI1&, NC)
          Call ULTIMOS3VENTAS(INFOSIAHP.msf2, CI1&, NC)
       End If
    End If
End Sub


Private Sub TXTMONEDA_Change()
   TXTSIMBOLO = SIMBOLPES$(XVALO(TXTMONEDA))
End Sub

Private Sub TXTMONEDA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "EL CAMBIO DE MONEDA NO ACTULIZA AQUI, POR QUE CUANDO SE ACEPTA" & vbCrLf
       TXTO$ = TXTO$ + "COMPARA CON LA MONEDA ACTIVA Y EN  CASO QUE DIFIERA REALIZA LA CONVERSION" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)

End Sub
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
    'RUTINA PARA QUE AL DAR ENTRER O ESCAPE SAQUE EL FOCO Y PONGA COMO NO VISIBLE EL TEXT DE EDICION
    If KeyAscii = 13 Or KeyAscii = 27 Then
       REG& = XVALO(TXTMOVIL.ToolTipText)
       If MSF5.Rows - 1 > REG& Then
          MSF5.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL.TEXT), "F10.1")
          MSF5.Row = REG& + 1
          MSF5_Click
          
          On Error Resume Next
          TXTMOVIL.SetFocus
          Call PINTATEXT(TXTMOVIL)
          On Error GoTo 0
          'Exit Sub
          'Call MSF5_MouseDown(1, 0, 1, 1)
       Else
          MSF5.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL.TEXT), "F10.1")
          MSF5.COL = 1
          
          TXTMOVIL.Visible = False
          
       End If
       '
       MSF5.SetFocus
       Call ACTUCANTI
       If TXTMOVIL.Visible = True Then TXTMOVIL.SetFocus
    End If
       
End Sub

Private Sub TXTMOVIL_LostFocus()
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF5.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
End Sub

Private Sub TXTNUMEPROVE_Change()
   TXTPOSIVA = TIPO_PROV$(XVALO(TXTNUMEPROVE))
   TXTPROVEEDOR = TEL2CLI$(XVALO(TXTNUMEPROVE) * -1)
   '
   CI1% = CODINT(txtcodigo)
   NPROV& = XVALO(TXTNUMEPROVE)
   If TXTPROVEEDOR = "Todos" Or XVALO(TXTNUMEPROVE) < 1 Then
      COSTOX# = COSUNI(CI1%)
      MONEX% = MONECOSTO%(CI1%)
   Else
     'SI ES UN PROVEEDOR EXISTE EN LA TABLA DE RELACION CODIGO PROVEEDOR
     COSTOX# = COSTO_X_PROV_DOLAR#(NPROV&, CI1%)
     TXTMONEDA = SAFETEXT$(2)
    End If
''''
   TXTCOSTO.TEXT = FORMATNUM$(COSTOX#, "F12.2")

End Sub
Function TIPO_PROV$(NPROVE)
   NPRO = NPROVE
   If NPRO > 0 Then NPRO = NPRO * -1
   NUME% = XVALO(Str$(PIVACLI%(NPRO)))
   If NUME% <> 5 Then
      TIPO_PROV$ = "LOCAL"
   Else
      TIPO_PROV$ = "EXTERIOR"
   End If
   
End Function

Private Sub TXTNUMEPROVE_DblClick()
   TXTNUMEPROVE = ""
End Sub

Private Sub TXTPESO_Change()
'COMENTADO POR QUE CUANDO CAMBIA EL PESO ACTUALIZABA EL PLRECIO DE VENTA Y SI LO HABIA CAMBIADO MANUALMENTE, NO CORRESPONDIA
'   VIAX$ = TXCOMBO6.TEXT
'   VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))
'   TXTPRECIOVENTA = PRECIOVENTA_X_FORMULA(XVALO(TXTCOSTO), VIA%)
End Sub

Private Sub TXTPESO_DblClick()
    If TXTPESO.Locked = False Then
       TXTPESO.BackColor = &HE0E0E0
       TXTPESO.Locked = True
    Else
       TXTPESO.Locked = False
       TXTPESO.BackColor = &H80000005
       Call PINTATEXT(TXTPESO)
    End If

End Sub


Private Sub TXTPESO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TXTPESO_LostFocus
End Sub

Private Sub TXTPESO_LostFocus()
  If TXTPESO.Locked = True Then Exit Sub
  '
  TXTPESO = FORMATNUM$(XVALO(TXTPESO), "F10.4")
  NPROV& = XVALO(TXTNUMEPROVE)
  CI1% = CODINT%(txtcodigo)
  KILOX = XVALO(TXTPESO)
  VIAX$ = VIA_SEGUN_KG(NPROV&, KILOX)
  'TXCOMBO6.TEXT = VIAX$

  VIA% = XVALO(TRIM$(Left$(VIAX$, 2)))

  TXTFECHACOSTEO = FECHATEX$(FECHACOSTO%(NPROV&, VIA%, CI1%))
  '
'  Call CARGARCOMBOVIAS(NPROV&, KILOX)


'  'SI ESTA EDITANDO ACTUALIZA
'  Call ACTUALIZA_INFO_SUPERIOR(1)
'  '
'  ESTABIEN% = 0
'  For I% = 1 To Combo6.ListCount - 1
'     If Mid$(Combo6.TEXT, 4, 8) = TRIM$(TXCOMBO6.TEXT) Then
'       ESTABIEN% = 1
'       Exit For
'     End If
'  Next I%
  
'  If ESTABIEN% < 1 And UCase$(InStr(TXCOMBO6.TEXT, "LOCAL")) < 1 Then
'     Call COMUNI("Via de Provisión No Corresponde")
'  End If
  '
'  NPROV& = XVALO(TXTNUMEPROVE)
'  KILOX = XVALO(TXTPESO)
'  Call CARGARCOMBOVIAS(NPROV&, KILOX)
End Sub

Private Sub TXTPRECIOVENTA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "EL P.VENTA SE ACTUALIZA SEGUN LA VIA DE ENTREGA CON EL PORCENTAJE DE LA TABLA 'VIASENTR'" & vbCrLf
       TXTO$ = TXTO$ + "SI LA VIA DE ENTREGA ESTA VACIA TRAE EL PORCENTAJE DE LA VIA LOCAL (1)" + vbCrLf
       TXTO$ = TXTO$ + "EJEMPLO:  VENTA = COSTO * PORCENTAJE (SEGUN VIA DE ENTREGA)" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.Row > 1 Then MSF1.Row = MSF1.Row - 1
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.Row < MSF1.Rows Then MSF1.Row = MSF1.Row + 1
            On Error GoTo 0
    End Select
    '
End Sub
Private Sub MSF5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          '
          If MSF5.MouseCol <> 5 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF5.COL = 5
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF5.CellWidth
            .Left = MSF5.CellLeft + MSF5.Left
            .Top = MSF5.CellTop + MSF5.Top
            .TEXT = MSF5.TEXT
            .Visible = True
            .ToolTipText = MSF5.Row
            .BackColor = &HC0C0FF
            .SetFocus
            
          End With
          '
          
End Sub
Private Sub MSF5_MouseUp(Button As Integer, _
         Shift As Integer, X As Single, Y As Single)
          Static CurrentWidth As Single
          If MSF5.Rows <= 1 Then Exit Sub
          If MSF5.CellWidth <> CurrentWidth Then
              On Error Resume Next
              TXTMOVIL.Width = MSF5.CellWidth
              CurrentWidth = MSF5.CellWidth
              On Error GoTo 0
          End If
End Sub

Private Sub TXTPROVEEDOR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "NOMBRE FANTASIA DEL PROVEEDOR" & vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub
Sub CONTROL_CANT_LOTE_MSF5(OK%)
   OK% = 1
   CI1% = CODINT(txtcodigo)
   For i& = 1 To MSF5.Rows - 1
        LOTE$ = MSF5.TextMatrix(i&, 1)
        ST = XVALO(MSF5.TextMatrix(i&, 2))
        CANTI = XVALO(MSF5.TextMatrix(i&, 5))
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
             A = TEXTOLOTES$(STRINGLOTES$, ";")
             For J% = 1 To UBound(CADENATEX$) Step 2
                LOTE$ = CADENATEX$(J%)
                If LOTE$ = LOTEPARAM$ Then
                  CANTI = XVALO(CADENATEX$(J% + 1))
                  ACUCANTI = ACUCANTI + CANTI
                End If
             Next J%
           End If
        Next i&
    End With
    '
    CANTILOTE_EN_GRILLA = ACUCANTI
    '
End Function


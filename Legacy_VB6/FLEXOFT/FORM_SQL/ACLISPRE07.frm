VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ACLISPRE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precios de Venta"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11670
   Icon            =   "ACLISPRE07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11670
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Indice de Listas de Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4260
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Width           =   3795
      End
   End
   Begin VB.TextBox Text5 
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
      Left            =   7030
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3270
      Width           =   625
   End
   Begin VB.TextBox Text4 
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
      Left            =   7030
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Text            =   "M"
      Top             =   3000
      Width           =   625
   End
   Begin VB.TextBox TXTNROORIGINAL 
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
      Left            =   7650
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3270
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
      Left            =   7650
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Nro. Original"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Grabar Modificaciones Realizadas"
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
      Left            =   7320
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
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
      Left            =   2080
      TabIndex        =   58
      ToolTipText     =   "Lista de Selección de Artículos Completa"
      Top             =   3285
      Width           =   220
   End
   Begin VB.TextBox txtlista 
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
      Left            =   9090
      TabIndex        =   55
      ToolTipText     =   "Enter para Ingresar Nuevo Item a la Lista"
      Top             =   3270
      Width           =   1575
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
      Left            =   150
      TabIndex        =   51
      Top             =   3270
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
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "Código"
      Top             =   3000
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Descripción"
      Top             =   3000
      Width           =   4665
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
      Left            =   9090
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "Precio"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox TXTDESCRIPCION 
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
      Left            =   2400
      TabIndex        =   52
      Top             =   3270
      Width           =   4665
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4455
      Left            =   150
      TabIndex        =   45
      Top             =   3720
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   7858
      _Version        =   393216
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton CMDCONVERSIONASQL 
      Caption         =   "SQL"
      Height          =   300
      Left            =   10920
      TabIndex        =   33
      Top             =   7320
      Width           =   645
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004040&
      Height          =   8290
      Left            =   10815
      ScaleHeight     =   8235
      ScaleWidth      =   1845
      TabIndex        =   30
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Terminar - Salir de Listas de Precios de Venta"
         Top             =   120
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   810
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Emite Lista de Precios Impresa"
         Top             =   6000
         Width           =   600
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   2490
         Width           =   600
      End
      Begin VB.CommandButton Command9 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Opciones Complementarias"
         Top             =   3870
         Width           =   600
      End
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Guarda Configuración de la Lista"
         Top             =   1800
         Width           =   600
      End
      Begin VB.CommandButton Command7 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Nueva Lista de Precios"
         Top             =   3180
         Width           =   600
      End
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ACLISPRE07.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Busqueda de Listas de Precios por Denominación"
         Top             =   5040
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   525
         TabIndex        =   31
         Top             =   7050
         Width           =   585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ACLISPRE07.frx":2836
         Top             =   7680
         Width           =   1515
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Utilidades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   8920
      TabIndex        =   19
      Top             =   120
      Width           =   1740
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   26
         Top             =   1890
         Width           =   1450
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "ACLISPRE07.frx":4758
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Control y Depuración de Lista Activa"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Control"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   28
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   23
         Top             =   1155
         Width           =   1450
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "ACLISPRE07.frx":4A62
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Ordena por Código de Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort por Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   25
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   20
         Top             =   420
         Width           =   1450
         Begin VB.CommandButton Command11 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Actualización por Porcentaje"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste +/- (%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   22
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Lista de Precio Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   4690
      TabIndex        =   2
      Top             =   120
      Width           =   3945
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Incl. IVA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3030
         TabIndex        =   34
         Top             =   1470
         Width           =   765
      End
      Begin VB.TextBox Text3 
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
         Left            =   3360
         TabIndex        =   18
         Top             =   2280
         Width           =   435
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
         Height          =   285
         Left            =   945
         TabIndex        =   14
         Top             =   2220
         Width           =   175
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
         Height          =   285
         Left            =   945
         TabIndex        =   10
         Top             =   1590
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Height          =   285
         Left            =   3570
         TabIndex        =   8
         Top             =   990
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   630
         TabIndex        =   5
         Top             =   1000
         Width           =   1590
      End
      Begin VB.TextBox Text1 
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
         Left            =   200
         TabIndex        =   3
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal:"
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
         Left            =   2730
         TabIndex        =   17
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1260
         TabIndex        =   16
         Top             =   2235
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
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
         Left            =   0
         TabIndex        =   15
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   630
         TabIndex        =   13
         Top             =   2235
         Width           =   330
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1260
         TabIndex        =   12
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formato:"
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
         Left            =   0
         TabIndex        =   11
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   1605
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   2730
         TabIndex        =   7
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vigencia:"
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
         Left            =   2100
         TabIndex        =   6
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abrev:"
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
         Left            =   -105
         TabIndex        =   4
         Top             =   750
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ACLISPRE07.frx":4D6C
      TabIndex        =   37
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "ACLISPRE07.frx":4DE6
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   4215
      Left            =   240
      TabIndex        =   57
      Top             =   3720
      Visible         =   0   'False
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   12648447
      FocusRect       =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   315
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   11355
      X2              =   11355
      Y1              =   -105
      Y2              =   8295
   End
   Begin VB.Menu mnuTransaccion 
      Caption         =   "Transacciones"
      Begin VB.Menu MnuGrabar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnuNuevoItem 
         Caption         =   "Nuevo Item de Stock"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "Nueva Lista de Precios"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActualizacionporcent 
         Caption         =   "Actualización por Porcentaje"
      End
      Begin VB.Menu mnusort 
         Caption         =   "Ordena por Código de Artículo"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReunir 
         Caption         =   "Re-Unir Listas de Precios"
      End
      Begin VB.Menu mnuimporta 
         Caption         =   "Importador de listas de precios"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActulistaxg 
         Caption         =   "Actualización de Listas de Precios por Grupo de Actualización"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnubuscar 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu mnucl 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuListadePrecios 
         Caption         =   "Lista de Precios Impresa"
      End
      Begin VB.Menu mnuVerAntlp 
         Caption         =   "Ver Versiones Anteriores de Listas de Precios"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTablaAux 
         Caption         =   "Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "ACLISPRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LISTACTIVA&
Dim Ban_Cancel As Integer
Dim MODIFICA%

Sub RESPALDOTABLA(LISTACTIVA&, RET_OK%)
      On Error GoTo ERROR_GUARDAR
      OK% = 0: J& = 0
      H$ = Time
      Screen.MousePointer = 11
      MAX_ID = 1 + XVALO(VALOBADA("PREVENTA_RESP", "max(PREVENTA_RESP_ID)", "COD_LISTA > 0 ", "NOMESS"))
      SQLX$ = "SELECT * FROM PREVENTA WITH(NOLOCK) WHERE Cod_Lista = " & LISTACTIVA&
      Set RS = READSET(SQLX$)
      'GUARDO EL INDICE DE LISTA
      FECHALTAX = FETEXCOR1$(HOY(HO$))
      FE% = HOY(HO$)
      '
      strSQL = "SELECT * FROM PREVENTA_RESP "
      strSQL = strSQL & " WHERE Cod_Lista < 1"
            
      Dim RSRESP As ADODB.Recordset
      Set RSRESP = New ADODB.Recordset
27     On Error Resume Next
      RSRESP.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
      End If
      On Error GoTo 0
    
      With RS
        Do While Not .EOF
            RSRESP.AddNew
            J& = J& + 1
            RSRESP!Cod_Lista = XVALO(!Cod_Lista)
            RSRESP!Cod_Pieza = XVALO(!Cod_Pieza)
            RSRESP!Num_Orden = XVALO(!Num_Orden)
            RSRESP!Desc_Pieza = SAFETEXT$(!Desc_Pieza)
            RSRESP!Denom_Pieza = SAFETEXT$(!Denom_Pieza)
            RSRESP!Carac_Control = SAFETEXT$(!Carac_Control)
            RSRESP!Num_Original = SAFETEXT$(!Num_Original)
            RSRESP!Precio_Unitario = XVALO(!Precio_Unitario)
            RSRESP!Cant_Limite = XVALO(!Cant_Limite)
            RSRESP!Porc_Limite = XVALO(!Cod_Lista)
            RSRESP!MARBORRA = 0
            RSRESP!CODIEMPR = CodiEmp%
            RSRESP!FECHMOD = CVDAT(FE%)
            RSRESP!IDMODIFIC = MAX_ID
            RSRESP!HORA = H$
            RSRESP.Update
           .MoveNext
       Loop
    End With
    RS.Close
    Set RS = Nothing
    RSRESP.Close
    Set RSRESP = Nothing
    GoTo 99
    Exit Sub
   '            sqlx2$ = "INSERT INTO PREVENTA_RESP"
'            sqlx2$ = sqlx2$ + " (Cod_Lista,Cod_Pieza,Num_Orden,Desc_Pieza,Denom_Pieza,Carac_Control,"
'            sqlx2$ = sqlx2$ + " Num_Original,Precio_Unitario,Cant_Limite,Porc_Limite,MarBorra,CODIEMPR,FECHMOD,IDMODIFIC,HORA) "
'
'            sqlx2$ = sqlx2$ + " VALUES(" & XVALO(!COD_LISTA)
'            sqlx2$ = sqlx2$ + " ," & XVALO(!Cod_Pieza)
'            sqlx2$ = sqlx2$ + " ," & XVALO(!Num_Orden)
'            sqlx2$ = sqlx2$ + " ,'" & SAFETEXT$(!Desc_Pieza) & "'"
'            sqlx2$ = sqlx2$ + " ,'" & SAFETEXT$(!Denom_Pieza) & "'"
'            sqlx2$ = sqlx2$ + " ,'" & SAFETEXT$(!Carac_Control) & "'"
'            sqlx2$ = sqlx2$ + " ,'" & SAFETEXT$(!Num_Original) & "'"
'            sqlx2$ = sqlx2$ + " ," & XVALO(!Precio_Unitario)
'            sqlx2$ = sqlx2$ + " ," & XVALO(!Cant_Limite)
'            sqlx2$ = sqlx2$ + " ,'" & XVALO(!Porc_Limite) & "'"
'            sqlx2$ = sqlx2$ + " ," & XVALO(!MarBorra)
'            sqlx2$ = sqlx2$ + " ," & CodiEmp%
'            sqlx2$ = sqlx2$ + " ,'" & FECHALTAX & "'"
'            sqlx2$ = sqlx2$ + " ," & MAX_ID
'            sqlx2$ = sqlx2$ + " ,'" & H$ & "'"
'            sqlx2$ = sqlx2$ + " )"
'            FLEXCONN.Execute (sqlx2$)

ERROR_GUARDAR:
    If Err <> 0 Then
       MsgBox Err.Description
       RET_OK% = -1
    End If
99  Screen.MousePointer = 1

End Sub

Sub VERIFICA_MODIFICACION()
   
   If MODIFICA% > 0 Then
      If COMALTER%("¿Desea Guardar Los Cambios Realizados?\\No, Cancelar\Si, Actualizar Cambios Realizados") <> 2 Then Exit Sub
      Call Command3_Click ' GRABA
      MODIFICA% = 0
   End If
   
End Sub

Private Sub CMDCONVERSIONASQL_Click()
   CMDCONVERSIONASQL.Enabled = False
   If DERACCE%("CONASQL", "Convertir archivos a SQL") < 2 Then
      Exit Sub
   End If
   ' No permite modificar datos futuros.
    If CantRegistros("PREVENTA") > 0 Then
        Call COMUNI("Imposible Migrar Lista de Precios.\La Tabla Destino de Precios\Ya Contiene Informacion.")
        GoTo 99
    End If
    '
    If CLACONTRA% < 1 Then GoTo 99
    
  '//OT-08-0033//
  'BORRO LOS DATOS EN LISTAPRE
  SQLX$ = "DELETE  FROM LISTAPRE "
  FLEXCONN.Execute (FILTSQL$(SQLX$))
  
  'BORRO LOS DATOS EN PREVENTA
'  SQLX$ = "DELETE  FROM PREVENTA "
'  FLEXCONN.Execute (SQLX$)
  '
  SQLX$ = "UPDATE PREVENTA SET MARBORRA = 1"  'MARCA MARBORRA EN 1 PARA BORRAR LUEGO DE REALIZAR EL UPDATE
   FLEXCONN.Execute (FILTSQL$(SQLX$))
  '
  'CONSULTA PARA LISTAPRE
  SQLX$ = "SELECT * FROM LISTAPRE"
  Dim INDICELISTA As ADODB.Recordset
  Set INDICELISTA = New ADODB.Recordset
  INDICELISTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  
  'CONSULTA PARA PREVENTA
  SQLX$ = "SELECT * FROM PREVENTA"
  Dim PRECIOLISTA As ADODB.Recordset
  Set PRECIOLISTA = New ADODB.Recordset
  PRECIOLISTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  
  'GUARDO EL INDICE DE LISTA
  With INDICELISTA
    FIN& = ULTREG("LISTAS")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      X$ = REGLEIDO$("LISTAS", i&)
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !Cod_Lista = Asc(Mid$(X$, 1, 1))
      !DENOM_LISTA = Mid$(X$, 2, 46)
      ' SE CORRIGE PORQUE SINO SE PRODUCE UN ERROR DE FECHAS
      FF% = CVI(Mid$(X$, 63, 2))
      !VIGENCIA = CVDAT(FF%)
      !Formato = Asc(Mid$(X$, 48, 1))
      !DECIMALES = Asc(Mid$(X$, 49, 1))
      !Moneda = Asc(Mid$(X$, 50, 1))
      !DESC_ABREV = Mid$(X$, 51, 12)
      .Update
      With PRECIOLISTA 'GUARDO LOS LOS ARTICULOS SEGUN LISTA
        '
        'GENERO EL NUMERO DE LISTA DE 3 POSICIONES
        'Y CONCATENO CON "LIPRE"
        ARCHIPRE$ = TRIM$(Str$(i&))
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        'ARCHIVO SEGUN NUMERO DE LISTA
        Call COPYSTRU("LIPREVEN", ARCHIPRE$)
        
        FIN1& = ULTREG&(ARCHIPRE$)
        For U& = 1 To FIN1&
          NXX$ = REGLEIDO$(ARCHIPRE$, U&)
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CODIEMPR = CodiEmp%
          On Error GoTo 0
          !Cod_Lista = i&
          CII% = CVI(Mid$(NXX$, 1, 2))
          !Cod_Pieza = CVI(Mid$(NXX$, 1, 2))
          !Num_Orden = U&
          !Desc_Pieza = Mid$(NXX$, 3, 42)
          !Denom_Pieza = DESCRIT0$(CII%)
          !Carac_Control = Mid$(NXX$, 45, 1)
          !Num_Original = Mid$(NXX$, 46, 16)
          !Precio_Unitario = CVD(Mid$(NXX$, 62, 8))
          !Cant_Limite = CVI(Mid$(NXX$, 70, 2))
          !Porc_Limite = CVI(Mid$(NXX$, 72, 1))
          .Update
        Next U&
      End With
    Next i&
  End With
  '
  INDICELISTA.Close
  Set INDICELISTA = Nothing
  PRECIOLISTA.Close
  Set PRECIOLISTA = Nothing
  '
  '//OT-08-0033//
  '\\\OT-08-1046-OD-17/11/08
  'BORRRAR LOS REGISTROS MARCADOS CUYO CAMPO MARBORRA ESTA EN 1
  
  SQLX$ = "DELETE FROM PREVENTA WHERE MARBORRA = 1"
   FLEXCONN.Execute (FILTSQL$(SQLX$))
  '\\\OT-08-1046-OD-FIN
   Call COMUNI("Proceso Finalizado")
   
99 CMDCONVERSIONASQL.Enabled = True
End Sub

Private Sub Command1_Click()
   If Ban_Cancel < 1 Then
    If MODIFICA% > 0 Then
       If COMALTER%("¿Desea Guardar Los Cambios Realizados?\\No, Cancelar\Si, Actualizar Cambios Realizados") <> 2 Then Exit Sub
       Call Command3_Click ' GRABA
       MODIFICA% = 0
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
   End If
End Sub

Private Sub Command10_Click()
    
'   Call SELECHO("VALIMONE")
'   If VALACT1$("VALIMONE") <> "" Then
'     Label7 = VALACT1$("VALIMONE")
'   End If
    Call SELECHO("TAMONED")
    If VALACT1$("TAMONED") <> "" Then
      Label7 = VALACT1$("TAMONED")
    End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") = 2 Then
      Call APREPOR
      Screen.MousePointer = 1
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SORLIPRE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call COMUNI("Opción No Disponible Momentaneamente")
   Command13.Enabled = True
   Exit Sub
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then GoTo 99
   End If
   '
   If DERACCE%("DEPUPREC", "Depurar Listas de Precios de Venta") < 2 Then
     GoTo 99
   End If
   '
   'ABRE EL FORMULARIO DE SELECCION DE OPERACION
   CLEANPRE.Show 1
   DoEvents
   '
   If CLEANPRE.Check1(2).Value = Checked Then
     Call VEPRECOS
     GoTo 99
   End If
   '
   LICOBAJ% = 0: If CLEANPRE.Check1(0).Value = Checked Then LICOBAJ% = 1
   LISINST% = 0: If CLEANPRE.Check1(1).Value = Checked Then LISINST% = 1
   CLEANPRE.Check1(0).Value = Unchecked
   CLEANPRE.Check1(1).Value = Unchecked
   CLEANPRE.Check1(2).Value = Unchecked
   DoEvents
   '
   ' SI SELECCIANA CHECK1(ELIMINAR ITEMS DE BAJA)
   If LICOBAJ% > 0 Or LISINST% > 0 Then
     If LISTACTIVA& < 1 Then GoTo 99
     ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
     While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ARCHIPRE$
     '
     Screen.MousePointer = 11
     JJ& = 0
     'FIN& = ULTREG&(ARCHIPRE$)
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + " ORDER BY Num_Orden ;"
     'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
     Dim ARTSELEC As ADODB.Recordset
     Set ARTSELEC = New ADODB.Recordset
     ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With ARTSELEC
         On Error Resume Next
        .MoveFirst
        FIN& = .RecordCount
        If Err > 0 Then
          On Error GoTo 0
          Call MENSERR(24, "Lista Vacía")
          Exit Sub
        End If
        On Error GoTo 0
        '
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          'X$ = REGLEIDO$(ARCHIPRE$, U&)
          'CIXI% = CVI(Left$(X$, 2))
          X$ = Space$(64)
          CIXI% = XVALO(!Cod_Pieza)
          cx = CODEXT$(CIXI%)
          Call REPLA(X$, MKI(CIXI%), 1, 2)
          Call REPLA(X$, DESCRIT$(CIXI%), 3, 42)
          Call REPLA(X$, SAFETEXT$(!Carac_Control), 45, 1)
          Call REPLA(X$, SAFETEXT$(!Num_Original), 46, 16)
             PRUNI# = XVALO(!Precio_Unitario)
          Call REPLA(X$, MKD$(PRUNI#), 62, 8)
          .MoveNext
            '****
          If LICOBAJ% = 1 Then
            'If Mid$(CODEXT$(CIXI%), 16, 1) = Chr$(96) Then
            ESTADO& = STATITEM%(CIXI%)
            If ESTADO& = -1 Then
              GoTo 19
            End If
          End If
          '
          If LISINST% = 1 Then
            If STOCKACT(CIXI%) < 0.05 Then
              GoTo 19
            End If
          End If
          '
          JJ& = JJ& + 1
          Call GRAREG("!LIPREVEN", X$, JJ&)
          '
19      Next U&
     End With
     Set ARTSELEC = Nothing
     '
     Call SETULTREG("!LIPREVEN", JJ&)
     Call CIERRARCH("!LIPREVEN")
     Screen.MousePointer = 1
     '
     Call EDITLISTA(ARCHIPRE$)
     
   End If
   '
   
99 Screen.MousePointer = 1
   Command13.Enabled = True
   '
End Sub

Sub VEPRECOS()
   '
   If LISTACTIVA& < 1 Then Exit Sub
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   'MUESTRA UN FORMULARIO SOLICITANDO EL RANGO DE PRECIOS
10 VDEF$ = MKS$(1.5) + MKS$(4)
   OBX$ = OBJPLA$("COEFICOSPRE", VDEF$)
   If OBX$ <> "" Then
     Screen.MousePointer = 11
     COEFIMI = CVS(Left$(OBX$, 4))
     COEFIMA = CVS(Mid$(OBX$, 5, 4))
     If COEFIMI > COEFIMA Then GoTo 10
     '
     'ARMA EL TITULO
     X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
     TITU$ = TRIM$(Mid$(X$, 2, 46))
     Call HEADERS("LIPREVEN", "")
     Call HEADERS("LIPREVEN", "Lista: " + TITU$ + " - Precios a Revisar por Costos")
     '
     'LEE LOS REGISTROS DE LA LISTA SELECCIONADA
     JJ& = 0
     
     DECIMALEX% = XVALO(Text3) '
     If DECIMALEX% < 2 Or DECIMALEX% > 4 Then DECIMALEX% = 2
     '
     'FIN& = ULTREG&(ARCHIPRE$)
     SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + " ORDER BY Num_Orden ;"
     'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
     Set ARTSELEC = READSET(SQLX$)
'     Dim ARTSELEC As ADODB.Recordset
'     Set ARTSELEC = New ADODB.Recordset
'     ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With ARTSELEC
      On Error Resume Next
      .MoveFirst
      FIN& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Lista Vacía")
        Exit Sub
      End If
      On Error GoTo 0
'      For U& = 1 To FIN&
       Do While Not .EOF
        Call TRACE(20, U&, FIN&)
        'XX$ = REGLEIDO$(ARCHIPRE$, U&)
        'CIXI% = CVI(Left$(XX$, 2))
        CIXI% = XVALO(!Cod_Pieza)
        'PRE# = CVD(Mid$(XX$, 62, 8))
        PRE# = XVALO(!Precio_Unitario)
        COCAL = COSCALCU(CIXI%, MODOCO%)
        If ROUND(PRE#, DECIMALEX%) < ROUND(COEFIMI * COCAL, DECIMALEX%) Or ROUND(PRE#, DECIMALEX%) > ROUND(COEFIMA * COCAL, DECIMALEX%) Then
          XX$ = Space$(64)
          Call REPLA(XX$, MKI$(CIXI%), 1, 2)
          Call REPLA(XX$, DESCRIT$(CIXI%), 3, 42)
          Call REPLA(XX$, !Denom_Pieza, 45, 0)
          Call REPLA(XX$, !Carac_Control, 45, 1)
          Call REPLA(XX$, !Num_Original, 46, 16)
          Call REPLA(XX$, MKD$(PRE#), 62, 8)
          Call REPLA(XX$, MKS$(XVALO(!Cant_Limite)), 70, 2)
          Call REPLA(XX$, MKD$(XVALO(!Porc_Limite)), 72, 1)
          JJ& = JJ& + 1
          Call GRAREG("LIPREVEN", XX$, JJ&)
        End If
       .MoveNext
       Loop
     End With
     Call SETULTREG("LIPREVEN", JJ&)
     '
     If JJ& < 1 Then
          Call MENSERR(24, "Sin Valores Fuera de Rango.")
        Else
          Screen.MousePointer = 1
          Call FINAL("*LILIPRE") '
     End If
   End If
   '
End Sub

Private Sub Command14_Click()
       Command14.Enabled = False
       Ban_Cancel = 1
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") > 0 Then
          If VERILLAVE% <> 1 Then Exit Sub
       End If
       '
       On Error Resume Next
       'LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
       On Error GoTo 0

'       ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
'       While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
'       ARCHIPRE$ = "LIPRE" + ARCHIPRE$
'       '
'       ATRI$ = "/NC"
'       ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
'       ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
'       ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
'       '
'       n% = XVALO(TRIM$(Str$(LISTACTIVA&)))
'       '\\OT-07-0424
'       Call LeeLaLista(n%)
'       '\\OT-07-0424
'       XII% = VENTABU%("LIPREVEN" + ATRI$)
'       If XII% >= 0 Then
'         Call BLOQARCH("LISTAS")
'         Ban_Cancel = 1 ' para que no se cierre el formulario si esta grabando
'
'         CADECI% = Val(Text3)
'         If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
'         List2.Clear
'         '
'         FIN& = ULTREG&("!LIPREVEN")
'         Screen.MousePointer = 11
'         '
'         HOII% = HOY(HOS$)
'         If Val(Label4) = 3 Then
'           For I& = 1 To FIN&
'             NXX$ = REGLEIDO$("!LIPREVEN", I&)
'             CI% = CVI(Left$(NXX$, 2))
'             If CI% > 0 Then
'               If CI% <= NUMAS% Then
'                 PRELI# = CVD(Mid$(NXX$, 62, 8))
'               End If
'             End If
'             Call REPLA(NXX$, HOS$, 46, 16)
'             Call GRAREG("!LIPREVEN", NXX$, I&)
'89         Next I&
'         End If
'         '
'         PUNCODIAC$ = String$(NUMAS%, 0): JJ& = 0
'         '\\OT-07-0424
'    '     SQLX$ = "DELETE  FROM Preventa "
'    '     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & N% & " ;"
'    '     FLEXCONN.Execute (SQLX$)
         '
         'RECORRO Y GUARDO EL PRECIO DE LA LISTA EN UN VECTOR PARA COMPARAR SI VARIO
         Dim PRLSTA(32768)
         For L& = 1 To 32768
           PRLSTA(L&) = 0
         Next L&
         SQLX$ = "SELECT Precio_Unitario,Cod_Pieza  FROM Preventa WITH(NOLOCK)"
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA&
         Set RS = READSET(SQLX$)
         With RS
           Do While Not .EOF
             CI1% = XVALO(!Cod_Pieza)
             PRLSTA(CI1%) = XVALO(!Precio_Unitario)
             .MoveNext
           Loop
         End With
         RS.Close
         Set RS = Nothing
         '
         Screen.MousePointer = 11
         CONDI$ = "Cod_Lista = " & LISTACTIVA&
         Call ACTUREGISTRO("PREVENTA", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
         '
         SQLX$ = "SELECT * FROM Preventa "
         SQLX$ = SQLX$ + "WHERE Cod_Lista = -1"
         '
         PUNCODIAC$ = String$(NUMAS%, 0): JJ& = 0
         Dim PREVVVENTA As ADODB.Recordset
         Set PREVVVENTA = New ADODB.Recordset
         PREVVVENTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         '
         'VER TEMA DE LA FECHA Y LOS ULTIMOS DOS CAMPOS QUE SE ESTAN GUARDANDO
         FIN& = MSF1.Rows - 1
         For i& = 1 To FIN&
            Call TRACE(20, i&, FIN&)
             
            Ci% = CODINT%(MSF1.TextMatrix(i&, 1))
            If Ci% > 0 Then
              If Ci% <= NUMAS% Then
                If Asc(Mid$(PUNCODIAC$, Ci%, 1)) > 0 Then
                  Call MENSERR(24, "Código '" + TRIM$(CODEXT$(Ci%)) + "' Repetido en Lista.\Suprimida Segunda Aparición.")
                  GoTo 94
                End If
                '
                DE$ = MSF1.TextMatrix(i&, 2)
                CRTERCONTROL$ = Left$(MSF1.TextMatrix(i&, 3), 1)
                NroOriginal = MSF1.TextMatrix(i&, 4)
                PRELI# = XVALO(MSF1.TextMatrix(i&, 5))
                FE% = XVALO(MSF1.TextMatrix(i&, 6))

'                ' QUEDA PENDIENTE VER PARA QUE ES ESTO
'                PRELI# = CVD(Mid$(NXX$, 62, 8))
'                PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
                '
                With PREVVVENTA
                   
                  .AddNew   ' YA CONVERTIDO
                  On Error Resume Next
                  !CODIEMPR = CodiEmp%
                  On Error GoTo 0
                  !Cod_Lista = LISTACTIVA&
                  !Cod_Pieza = Ci%
                  !Num_Orden = i&
                  !Desc_Pieza = DE$
                  !Denom_Pieza = DE$
                  !Carac_Control = CRTERCONTROL$
                  !Num_Original = NroOriginal
                  'SE AGREGA ESTO PARA QUE GRABE LA FECHA DE MODIFICACION COMO SE GRABA EN FRM ANACOS
                  'POR QUE SI NO CUANDO MODIFCABA EVENTUALMENTE DE LA LISTA LE BORRABA LA FECHA
                     CNTLIMTE = FE% ' LO LEE DE LA GRILLA (CUANDO ARMA LA GRILLA LO TRAE DE LA BASE)
                  If InStr(EMPREAC$, "J Y M") > 0 Then
                     If ROUND(PRELI#, 4) <> ROUND(PRLSTA(Ci%), 4) Then CNTLIMTE = HOY(HO$)
                  End If
                  'FIN CAMBIO JYM
                  
                  !Precio_Unitario = PRELI#
                  !Cant_Limite = CNTLIMTE
                  !Porc_Limite = 0
                  .Update
                End With
                '
                Mid$(PUNCODIAC$, Ci%, 1) = Chr$(1)
              End If
            End If
94       Next i&
         PREVVVENTA.Close
         Set PREVVVENTA = Nothing
         '
         SQLX$ = "DELETE FROM PREVENTA "     'BORRRAR LOS REGISTROS MARCADOS CUYO CAMPO MARBORRA ESTA EN 1
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " AND MARBORRA = 1"
         FLEXCONN.Execute (FILTSQL$(SQLX$))
         '
         '
         MODIFICA% = 0
         NLIST% = LISTACTIVA&
         GENERAR_DAT (NLIST%)
         'Call CARLISTA
         Ban_Cancel = 0 ' desmarca la bandera de cierre por que ya grabo
         'Call LIBARCH("LISTAS")
         'Call Command3_Click ' GRABA EN LISTAPRE
         Screen.MousePointer = 1
       
99  Command14.Enabled = True
End Sub

'
Private Sub Command2_Click()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label3 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command20_Click()

    Command20.Enabled = False
    '
    Call SELECHO("LISTAS")
    LISNOM$ = VALACT2$("LISTAS")
    If LISNOM$ = "" Then GoTo 99
    FIN& = List1.ListCount
    For U& = 1 To FIN&
    ABUS$ = UCase$(VALACT2$("LISTAS"))
    If InStr(UCase$(List1.List(U& - 1)), ABUS$) > 0 Then
        List1.ListIndex = U& - 1
        GoTo 99
    End If
    Next U&
    '
99  Command20.Enabled = True

End Sub

Private Sub Command21_Click()
   Command1.Enabled = False
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   '
   Command1.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   'LISTACTIVA& = List1.ListIndex + 1
   On Error Resume Next
   'LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
   On Error GoTo 0
   If LISTACTIVA < 1 Then
     Call MENSERR(24, "Imposible Procesar Transacción.\Ninguna Lista de Precios\Marcada Como Lista Activa.")
     GoTo 99
   End If
'   ***********
'   //OT -07 - 0434//
   If MODIFICA% > 0 Then Call RESPALDOTABLA(LISTACTIVA&, OK%)
   If OK% < 0 Then Call COMUNI("Imposible Continuar !!!\No se Pudo Crear Tabla de Respaldo\Vuelva a Intentarlo"): GoTo 99
   
   SQLX$ = "SELECT * FROM LISTAPRE "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
   'Set LISPRE = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim LisPre As ADODB.Recordset
   Set LisPre = New ADODB.Recordset
   LisPre.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   On Error Resume Next
   With LisPre
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      .MoveFirst
      '.Edit
      !DENOM_LISTA = Text1.TEXT
      !VIGENCIA = CVDAT(FECHANUM(Label3.Caption))
      !Formato = (Label4.Caption)
      If Val(Text3) < 0 Or Val(Text3) > 5 Then Text3 = "2"
      !DECIMALES = XVALO(Text3.TEXT)
      !Moneda = (Label7.Caption)
      !DESC_ABREV = Text2.TEXT
        !IncluIva = Check1.Value ' si check1 esta tildado va a ser 1 (iva incluido)caso contrario va a ser 0
      .Update
   End With
   Call CARILISTAS
   If MODIFICA% = 1 Then Call Command14_Click
   Call COMUNI("Proceso de Grabación Finalizado")
   
   '//OT -07 - 0434//
   'XX$ = REGLEIDO$("LISTAS", LISTACTIVA&)
   'Mid$(XX$, 2, 46) = AJUSTI$(Text1, 46)
   'Mid$(XX$, 51, 12) = AJUSTI$(Text2, 12)
   'FEXI% = FECHANUM(Label3)
   'Mid$(XX$, 63, 2) = MKI$(FEXI%)
   'Mid$(XX$, 48, 1) = Chr$(Val(Label4))
   'Mid$(XX$, 50, 1) = Chr$(Val(Label7))
   'Mid$(XX$, 49, 1) = Chr$(Val(Text3))
   'Call GRAREG("LISTAS", XX$, LISTACTIVA&)
   'Call CIERRARCH("LISTAS")
   'Call CARILISTAS
99 Command3.Enabled = True
'   List1.ListIndex = LISTACTIVA& - 1
End Sub

Private Sub Command4_Click()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     TXTCODIGO = COD$
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
        '
        Command7.Enabled = False

3000    ' ************************* nueva lista de precios
        '
        '//OT-07-0434//
        CantListas& = CantRegistros("LISTAPRE")
        '
'''        SQLX$ = "SELECT COUNT (*)AS CantReg FROM LISTAPRE"
'''        Dim CANTIREGIS As ADODB.Recordset
'''        Set CANTIREGIS = New ADODB.Recordset
'''        CANTIREGIS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'''        On Error Resume Next
'''        With CANTIREGIS
'''          .MoveFirst
'''          If Err < 1 Then
'''            CantListas& = !CantReg
'''            On Error GoTo 0
'''          End If
'''          On Error GoTo 0
'''        End With
'''        Set CANTIREGIS = Nothing
        
        SQLX$ = "SELECT * FROM LISTAPRE"
        Dim LISTTTAPRE As ADODB.Recordset
        Set LISTTTAPRE = New ADODB.Recordset
        LISTTTAPRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        
        With LISTTTAPRE
            
            If CantListas& >= 255 Then
                Call COMUNI("Hay 255 Listas de Precios\ya Definidas\-\Capacidad del Sistema Saturada.")
                Exit Sub
            End If
            '
            NULIS% = 1 + CantListas&
            X$ = Chr$(NULIS%) + Mid$(REGBLAN$("LISTAS"), 2)
3010        X$ = OBJPLA$("LISPRE-NUEVA", X$)
            If X$ = "" Then
                GoTo 99
            End If
            '
            If TRIM$(Mid$(X$, 2, 46)) <> "" Then
              If Asc(Mid$(X$, 48, 1)) > 0 Then
                If Asc(Mid$(X$, 48, 1)) <= 2 Then
                  If Asc(Mid$(X$, 49, 1)) <= 4 Then
                    If Asc(Mid$(X$, 50, 1)) > 0 Then
                      If Asc(Mid$(X$, 50, 1)) <= 2 Then
                        If TRIM$(Mid$(X$, 51, 12)) <> "" Then
                          If CVI(Mid$(X$, 63, 2)) > 0 Then
                            .AddNew   ' YA CONVERTIDO
                            On Error Resume Next
                            !CODIEMPR = CodiEmp%
                            On Error GoTo 0
                            .Fields("Cod_Lista") = Asc(Mid$(X$, 1, 1))
                            .Fields("Denom_Lista") = (Mid$(X$, 2, 46))
                            .Fields("Vigencia") = FECHATEX(CVI(Mid$(X$, 63, 2)))
                            .Fields("Formato") = Asc(Mid$(X$, 48, 1))
                            .Fields("Decimales") = Asc(Mid$(X$, 49, 1))
                            .Fields("Moneda") = Asc(Mid$(X$, 50, 1))
                            .Fields("Desc_Abrev") = (Mid$(X$, 51, 12))
                            .Update
                            Call CARILISTAS
                            List1.ListIndex = NULIS% - 1

                            '//OT-07-0434//
                            GoTo 99
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
            '
            Call COMUNI("Datos Incompletos\o No Válidos")
            GoTo 3010
          On Error GoTo 0
       End With
       Set LISTTTAPRE = Nothing
99     Command7.Enabled = True
       
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call LISTAPRE
  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPLIPREC07.Show 1
   Command9.Enabled = True
End Sub


Private Sub Form_Load()
   '
   MODIFICA% = 0
   
   VERANTER$ = "ADMIPRE"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'       Caption = Caption + " - " + EPAC$
'   End If
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
    EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
    Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
   End If
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture4)
   '
   Call ABRECONEXION
   Call CREATABLA_ENCAPRF
   
   Call GENERAR_ESTRUCTURAS_SQL
   Call ABREVENTAS
   'VALIDACION POR SI SE LLAMA DESDE OTRO PROYECTO EN MODO NO VISIBLE
   If ACLISPRE07.Visible = True Then
     Show
   End If
   Refresh
   '
   Screen.MousePointer = 11
   If ULTREG&("VALIFORM") < 3 Then
      Call BLANARCH("VALIFORM")
      Call REGAPP("VALIFORM", Chr$(1) + "Nro. Original")
      Call REGAPP("VALIFORM", Chr$(2) + "Descuento por Cantidad")
      Call REGAPP("VALIFORM", Chr$(3) + "Fecha Ultima Actualizacion")
      Call CIERRARCH("VALIFORM")
   End If
   '
   If ULTREG&("VALIMONE") < 2 Then
      Call BLANARCH("VALIMONE")
      Call REGAPP("VALIMONE", Chr$(1) + "Pesos")
      Call REGAPP("VALIMONE", Chr$(2) + "Dolares U$S")
      Call CIERRARCH("VALIMONE")
   End If
   '
   Call carga_enca
   Call CARILISTAS
   '

   ILIOO& = 0
   If XVALO(APLINVO$) > 0 Then ILI00& = XVALO(APLINVO$) - 1
   '
   LISTACTIVA& = -1
   If InStr(1, EMPREAC$, "FERVI") < 1 Then
        On Error Resume Next
        List1.ListIndex = ILI00&
        LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
        On Error GoTo 0
   End If
   Screen.MousePointer = 1
   '
End Sub
Sub carga_enca(Optional FORMATO1$)
   '
   TEPRUEBA.Font = MSF1.Font
   '
   MSF1.Cols = 7
   msf2.Cols = 7
   '
   If TRIM$(FORMATO1$) = "" Then
     FORMATO1$ = "N.Original"
   End If
   
   NUCOLU% = 1: TEPRUEBA = "12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = ALINEA1%: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 2: TEPRUEBA = "12345678901234567890": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = ALINEA1%: MSF1.TextMatrix(0, NUCOLU% - 1) = "Codigo"
   NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLLOPQRSTWXYZABCDEFGCH1234": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 4: TEPRUEBA = "123456789012345": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Control"
   NUCOLU% = 5: TEPRUEBA = "1234567890123": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = FORMATO1$
   NUCOLU% = 6: TEPRUEBA = "1234567890123456789": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "P.Unitario"
   NUCOLU% = 7: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   '
   NUCOLU% = 1: TEPRUEBA = "12": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = ALINEA1%: msf2.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 2: TEPRUEBA = "12345678901234567890": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = ALINEA1%: msf2.TextMatrix(0, NUCOLU% - 1) = "Codigo"
   NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLLOPQRSTWXYZABCDEFGCH1234": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 4: TEPRUEBA = "123456789012345": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "Control"
   NUCOLU% = 5: TEPRUEBA = "1234567890123": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = FORMATO1$
   NUCOLU% = 6: TEPRUEBA = "1234567890123456789": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "P.Unitario"
   NUCOLU% = 7: TEPRUEBA = "": msf2.ColWidth(NUCOLU% - 1) = 0: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = ""
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MODIFICA% > 0 Then
      MODIFICA% = 0
      If COMALTER%("¿Desea Guardar Los Cambios Realizados?\\No, Cancelar\Si, Actualizar Cambios Realizados") <> 2 Then Exit Sub
      Call Command3_Click ' GRABA
   End If
   
   Cancel = Ban_Cancel ' no se cierra si se le pasa a Ban_Cancel = 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = Ban_Cancel
   If Ban_Cancel < 1 Then ' no se cierra si se le pasa a Ban_Cancel = 1
    Call CIERRARCH("*.*")
    Call FINAL("")
   End If
End Sub

Private Sub Label15_DblClick()
   Call List2_DblClick
End Sub

Private Sub Label4_Change()
   Label6 = ECOARCH$("VALIFORM", Chr$(Val(Label4)))
   If Val(Label4) = 3 Then
       Call carga_enca("Fecha Act.")
       Text8 = "Fecha Act."
     ElseIf Val(Label4) = 2 Then
       Call carga_enca("Dto.X.Cant.")
       Text8 = "Dto.X.Cant."
     Else
       Call carga_enca("Nro.Original")
       Text8 = "Nro.Original"
   End If
   '
End Sub

Private Sub Label4_DblClick()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Label7_Change()
   Label9 = ECOARCH$("TAMONED", Chr$(Val(Label7)))
End Sub

Private Sub Label7_DblClick()
   Call SELECHO("TAMONED")
   If VALACT1$("TAMONED") <> "" Then
     Label7 = VALACT1$("TAMONED")
   End If
End Sub

Private Sub List1_Click()

   If List1.ListIndex <> LISTACTIVA& - 1 Then
        Call VERIFICA_MODIFICACION
        Command3.Enabled = False 'deshabilito el boton de grabar mientras esta cargando la lista.
        Call CARLISTA
        Command3.Enabled = True
   End If
   
End Sub
   '\\OT-07-0424
Sub CARLISTA()
   '
   Static LOCKUP%
   '
   Screen.MousePointer = 11
   'CONTROLA QUE NO SE CHOQUEN DOS PROCESO DE GRABACION Y SE MEZCLE LA INFORMACION
   If LOCKUP% = 1 Then
        Call COMUNI("Imposible Continuar.\Debe Esperar Que Finalice El Proceso de Carga Anterior")
        List1.ListIndex = LISTACTIVA& - 1
        Exit Sub 'SALE EN FORMA ABRUPTA PARA QUE NO CAMBIE EL VALOR DE LOCKUP%
   End If
   If Ban_Cancel = 1 Then
        Call COMUNI("Imposible Continuar.\Debe Esperar Que Finalice La Grabación de la Lista Anterior")
        List1.ListIndex = LISTACTIVA& - 1
        Exit Sub 'SALE EN FORMA ABRUPTA PARA QUE NO CAMBIE EL VALOR DE Ban_Cancel
   End If
   LOCKUP% = 1
   MSF1.Rows = 1
   msf2.Rows = 1

   'LISTACTIVA& = List1.ListIndex + 1
   On Error Resume Next
   LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
   On Error GoTo 0
   APERBASDAT ("PEDCLI")
   Call PoneEnCeroNull("PREVENTA", "MARBORRA", REGAF&, "NOMESS")
   
   SQLX$ = "SELECT * FROM LISTAPRE WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Cod_Lista = " & LISTACTIVA&
   
   'Set LISTASEL = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   'Dim LISTASEL As ADODB.Recordset
   'Set LISTASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set LISTASEL = READSET(SQLX$)
   InIva% = 0
   With LISTASEL
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 90
      End If
      On Error GoTo 0
      Text1 = TRIM$(SAFETEXT$(!DENOM_LISTA))
      Text2 = TRIM$(SAFETEXT$(!DESC_ABREV))
      Label3 = (Format$(!VIGENCIA, "dd/mm/yy"))
      Label4 = XVALO(!Formato)
      Label7 = XVALO(!Moneda)
      Text3 = XVALO(!DECIMALES)
      InIva% = XVALO(!IncluIva) ' LEE CAMPO PARA SABER SI LISTA INCLUYE IVA
   End With
   Set LISTASEL = Nothing
   Check1.Value = 0
   If InIva% = 1 Then Check1.Value = 1
   '
'Call CARGALISTA(List2, "PREVENTA", "Cod_Pieza/C16;Desc_Pieza/A31; Carac_Control/A4; Num_Original/A13; Precio_Unitario/ F15.4", "Cod_Lista = " & LISTACTIVA& & " ORDER BY Num_Orden ")
'GoTo 90
   '
   
   NN% = XVALO(Text3)
   If NN% = 0 Then NN% = 4
            
   SQLX$ = "SELECT MASTER.CODIGO, MASTER.DESCRIPCION,PREVENTA.Carac_Control, PREVENTA.Num_Original,PREVENTA.Precio_Unitario,Cant_Limite  FROM PREVENTA INNER JOIN MASTER WITH(NOLOCK) ON PREVENTA.Cod_Pieza = MASTER.CODINT  "
   SQLX$ = SQLX$ + " WHERE PREVENTA.Cod_Lista = " & LISTACTIVA&
   SQLX$ = SQLX$ + " AND PREVENTA.MARBORRA = 0"
   SQLX$ = SQLX$ + "ORDER BY Num_Orden ;"
   'Set PRODUC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   'Dim PRODUC As ADODB.Recordset
   'Set PRODUC = New ADODB.Recordset
   'PRODUC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set PRODUC = READSET(SQLX$)
   MSF1.Redraw = False
   On Error Resume Next
   With PRODUC
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        MSF1.Rows = 1
        MSF1.Redraw = True
        Screen.MousePointer = 1
        GoTo 90
      End If
      On Error GoTo 0
      FIN& = CantRegistros("PREVENTA", "Cod_Lista = " & LISTACTIVA& & " AND MARBORRA = 0")
      .MoveFirst
      If FIN& > 0 Then
        i& = 0
        Do While (.EOF <> True)
        i& = i& + 1
        Call TRACE(24, i&, FIN&)
        PRE# = XVALO(!Precio_Unitario)
        MSF1.Rows = i&
        MSF1.Rows = MSF1.Rows + 1
        MSF1.TextMatrix(i&, 1) = SAFETEXT$(!Codigo)
        MSF1.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
        MSF1.TextMatrix(i&, 3) = SAFETEXT$(!Carac_Control)
        MSF1.TextMatrix(i&, 4) = SAFETEXT$(!Num_Original)
        MSF1.TextMatrix(i&, 5) = FORMATNUM$(PRE#, "F14." & CStr(NN%))
        MSF1.TextMatrix(i&, 6) = XVALO(!Cant_Limite)
'        Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
'        Call REPLA(TTXX$, DESCRIT$(CII%), 18, 31)
'        Call REPLA(TTXX$, SAFETEXT$(!Carac_Control), 51, 2)
'        Call REPLA(TTXX$, SAFETEXT$(!Num_Original), 56, 13)
'        Call REPLA(TTXX$, FORMATNUM$(PRE#, "F12,4"), 72, 12)
'        List2.AddItem TTXX$
        .MoveNext
        Loop

      End If
   End With
   MSF1.Redraw = True
   Set PRODUC = Nothing
90 Screen.MousePointer = 1
   LOCKUP% = 0
   '\\OT-07-0424
   '
End Sub

Private Sub List2_DblClick()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
   End If
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
   '
   n% = XVALO(TRIM$(Str$(LISTACTIVA&)))
   '\\OT-07-0424
   Call LeeLaLista(n%)
   '\\OT-07-0424
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% >= 0 Then
     Call BLOQARCH("LISTAS")
     Ban_Cancel = 1 ' para que no se cierre el formulario si esta grabando

     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     List2.Clear
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For i& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", i&)
         Ci% = CVI(Left$(NXX$, 2))
         If Ci% > 0 Then
           If Ci% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, i&)
89     Next i&
     End If
     '
     PUNCODIAC$ = String$(NUMAS%, 0): JJ& = 0
     '\\OT-07-0424
'     SQLX$ = "DELETE  FROM Preventa "
'     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & N% & " ;"
'     FLEXCONN.Execute (SQLX$)
     '
     SQLX$ = "UPDATE PREVENTA SET MARBORRA = 1 "    'MARCA MARBORRA EN 1 PARA BORRAR LUEGO DE REALIZAR EL UPDATE
     SQLX$ = SQLX$ + " WHERE Cod_Lista = " & n% & ""
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & n% & " ;"
     '
     Dim PREVVVENTA As ADODB.Recordset
     Set PREVVVENTA = New ADODB.Recordset
     PREVVVENTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     
     FIN& = ULTREG&("!LIPREVEN")
     For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", i&)
        Ci% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If Ci% > 0 Then
          If Ci% <= NUMAS% Then
            If Asc(Mid$(PUNCODIAC$, Ci%, 1)) > 0 Then
              Call MENSERR(24, "Código '" + TRIM$(CODEXT$(Ci%)) + "' Repetido en Lista.\Suprimida Segunda Aparición.")
              GoTo 94
            End If
            '
            ' QUEDA PENDIENTE VER PARA QUE ES ESTO
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
            DE$ = DESCRIT0$(Ci%)
            Call REPLA$(NXX$, DE$, 3, 42)
            Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(NXX$, MRX$, 44, 1)
            '
            With PREVVVENTA
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !Cod_Lista = n%
              !Cod_Pieza = Ci%
              !Num_Orden = i&
              !Desc_Pieza = DESCRIT0$(Ci%)
              !Denom_Pieza = DESCRIT0$(Ci%)
              !Carac_Control = Mid$(NXX$, 45, 1)
              !Num_Original = Mid$(NXX$, 46, 16)
              !Precio_Unitario = CVD(Mid$(NXX$, 62, 8))
              !Cant_Limite = CVD(Mid$(NXX$, 70, 2))
              !Porc_Limite = CVD(Mid$(NXX$, 72, 1))
              .Update
            End With
            '
            Mid$(PUNCODIAC$, Ci%, 1) = Chr$(1)
          End If
        End If
94   Next i&
     Set PREVVVENTA = Nothing
     '
     SQLX$ = "DELETE FROM PREVENTA "     'BORRRAR LOS REGISTROS MARCADOS CUYO CAMPO MARBORRA ESTA EN 1
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & n% & " AND MARBORRA = 1"
     FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     Screen.MousePointer = 1
     '
     NLIST% = LISTACTIVA&
     GENERAR_DAT (NLIST%)
     Call CARLISTA
     Ban_Cancel = 0 ' desmarca la bandera de cierre por que ya grabo
     Call LIBARCH("LISTAS")
   End If
   '
End Sub

'Sub APREPOR()
'    '
'    Dim VALVIE(32767), VALNUE(32767)
'    HOII% = HOY(HOS$)
'    '
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    If InStr(EPAC$, "DOSIVA") > 0 Then
'      If VERILLAVE% <> 1 Then Exit Sub
'    End If
'    ' SELECCIONA LA MODALIDAD DE ACTUALIZAR PRECIOS
'    OPX% = COMALTER%("Puede Actualizar a Partir\de la Lista de Precios Actualmente Vigente\o de los Costos de Reposición para un Proveedor\\Lista Vigente\Costos de Reposición\Cancelar")
'
'    Select Case OPX%
'      Case 2
'        Call APREPROV
'        Exit Sub
'      Case 3
'        Exit Sub
'    End Select
'    '
'    ' LISTACTIVA& = List1.ListIndex + 1
'    On Error Resume Next
'    LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
'    On Error GoTo 0
'    'X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
'    '
'    RANGOAC% = 0
'    If InStr(EPAC$, "DOSIVA") > 0 Then
'      TTXX$ = "Opciones de Actualización por Porcentaje:\\Sólo Bombas\Sólo Repuestos\Toda la Lista"
'      RANGOAC% = COMALTER%(TTXX$)
'      If RANGOAC% < 1 Or RANGOAC% > 3 Then Exit Sub
'      If RANGOAC% = 3 Then RANGOAC% = 0
'    End If
'    '
'    Select Case RANGOAC%
'      Case 0
'        RACTU$ = "Lista Completa"
'      Case 1
'        RACTU$ = "Bombas Nacionales (Cod. 0...)"
'      Case 2
'        RACTU$ = "Repuestos (Cod. >= 1...)"
'    End Select
'    '
'    ' SELECCIONA RANGO DE CODIGOS
'    Call DESHASCOD(COD1$, COD2$, CIKX1%, CIKX2%)
'    If CIKX1% < 1 Or CIKX2% < 1 Then Exit Sub
'    '
'    '**********
'    SQLX$ = "SELECT * FROM LISTAPRE "
'    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
'    'SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
'    'Set LISTSELEC = Ventas.OpenRecordset(SQLX$, 4)
'    Dim LISTSELEC As ADODB.Recordset
'    Set LISTSELEC = New ADODB.Recordset
'    LISTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'
'    With LISTSELEC
'      On Error Resume Next
'      .MoveFirst
'      If Err > 0 Then
'         On Error GoTo 0
'         Exit Sub
'
'      End If
'      On Error GoTo 0
'      Call REPLA(VDEF$, Chr(XVALO(!Cod_Lista)), 1, 1)
'      Call REPLA(VDEF$, !DENOM_LISTA, 2, 46)
'      Call REPLA(VDEF$, Chr(XVALO(!FORMATO)), 48, 1)
'      Call REPLA(VDEF$, Chr(XVALO(!DECIMALES)), 49, 1)
'      Call REPLA(VDEF$, Chr(XVALO(!MONEDA)), 50, 1)
'      Call REPLA(VDEF$, !DESC_ABREV, 51, 12)
'      'VDEF$ = Left$(X$, 64) + MKS$(0)
'      Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
'      'CADECI% = Asc(Mid$(VDEF$, 49, 1))
'      CADECI% = !DECIMALES
'      If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
'      Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
'    End With
'
'    ' INGRESO DEL PORCENTAJE
'5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
'    If OBX$ = "" Then
'      Exit Sub
'    End If
'    ' VALIDACIONES
'    PORAJU = CVS(Mid$(OBX$, 65, 4)) 'PORCENTAJE DE AJUSTE
''    If Abs(PORAJU) < 0.001 Then Exit Sub 'COMENTADO PARA QUE PERMITA REALIZAR AJUSTE 0
'    '
'    If PORAJU < (-90) Or PORAJU > 99.99 Then
'       Call MENSERR(24, "Porcentaje no Válido")
'       GoTo 5
'    End If
'    '
'    CADECI% = Asc(Mid$(OBX$, 49, 1))
'    If CADECI% < 0 Or CADECI% > 4 Then ' SE MODIFICO LA VALIDACION < 2 PARA QUE PERMITA REDONDEAR A ENTERO
'       Call MENSERR(24, "Decimales entre '0' y '4'")
'       GoTo 5
'    End If
'   '
'   LISNO% = Asc(Left$(OBX$, 1))
'   '
'   Screen.MousePointer = 11
'   '*************************
'   '//OT-07-0434//
'   '
'   REBLA$ = REGBLAN$("ACOSPOR")
'   ' LEEMOS LISTA DE PRECIOS ACTIVA
'   SQLX$ = "SELECT * FROM Preventa "
'   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
'   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
'   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
'   Dim ARTSELEC As ADODB.Recordset
'   Set ARTSELEC = New ADODB.Recordset
'   ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'
'   With ARTSELEC
'     On Error Resume Next
'     .MoveLast
'     FIN& = .RecordCount
'     If Err > 0 Then
'        On Error GoTo 0
'        Call MENSERR(24, "Lista Vacía")
'        Exit Sub
'     End If
'     On Error GoTo 0
'     .MoveFirst
'     U& = 0:   JJ& = 0
'     Do While Not .EOF
'            U& = U& + 1
'            Call TRACE(20, U&, FIN&)
'            TTXX$ = Space$(128)
'            CII% = XVALO(!Cod_Pieza)
'            CDXXX$ = CODEXT(XVALO(!Cod_Pieza))
'            '
'
'            PORAJX = PORAJU
'            If RANGOAC% > 0 Then
'              CDDX$ = Left$(CDXXX$, 1) '
'              If RANGOAC% = 1 Then
'                  If CDDX$ <> "0" Then PORAJX = 0
'                ElseIf RANGOAC% = 2 Then
'                  If CDDX$ = "0" Then PORAJX = 0
'              End If
'            End If
'            '
'            If UCase$(CDXXX$) < UCase$(COD1$) Or UCase$(CDXXX$) > UCase$(COD2$) Then
'              PORAJX = 0
'            End If
'            '
'            COSVIE# = XVALO(!Precio_Unitario)
'            If CADECI% > 0 Then
'              COSNUE = ROUND#(COSVIE * (100 + PORAJX) / 100, CADECI%)
'            Else
'              'CASO QUE SE PASE DECIMALES 0: POR QUE EL ROUND SI CADECI% = 0 APLICA 2 DECIMALES
'              COSNUE = XVALO(FORMATNUM$((COSVIE * (100 + PORAJX) / 100), "F16.0"))
'            End If
'            '
'            ZZ$ = REBLA$
'            Call REPLA(ZZ$, MKI$(CII%), 1, 2)
'            Call REPLA(ZZ$, MKS$(COSVIE#), 3, 4)
'            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
'
'            JJ& = JJ& + 1
'            Call GRAREG("!ACOSPOR", ZZ$, JJ&)
'        .MoveNext
'     Loop
'   End With
'   Set ARTSELEC = Nothing
'   '//OT-07-0434//
'    Call SETULTREG("!ACOSPOR", JJ&)
'    Call CIERRARCH("!ACOSPOR")
'    Screen.MousePointer = 1
'
'    ' PRESENTA PRECIOS ACTUALES Y NUEVOS
'    ATRI$ = "/NC"
'    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
'    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
'    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
'    '
'    VTB% = VENTABU%("ACOSPOR" + ATRI$)
'    If VTB% >= 0 Then
'
'      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
'         '
'         'GRABAMOS LOS CAMBIOS
'         Screen.MousePointer = 11
'         HOII% = HOY(HOS$)
'         URE& = ULTREG&("!ACOSPOR")
'             For U& = 1 To URE&
'             Call TRACE(20, U&, URE&)
'             x$ = REGLEIDO$("!ACOSPOR", U&)
'             CII% = CVI(Mid$(x$, 1, 2))
'             'Y$ = REGLEIDO$(APRECIO$, U&)
'             'MRX$ = Mid$(Y$, 45, 1)
'             PRENUEVO = CVS(Mid$(x$, 7, 4))
'             PREX$ = COSNUEX$
'             '
'             SQLX$ = "SELECT * FROM Preventa "
'             SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
'             SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
'             'Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
'             Dim ITEMSEL As ADODB.Recordset
'             Set ITEMSEL = New ADODB.Recordset
'             ITEMSEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'
'             With ITEMSEL
'              On Error Resume Next
'             .MoveFirst
'             FIN& = .RecordCount
'             If Err > 0 Then
'               On Error GoTo 0
'               Call MENSERR(24, "Lista Vacía")
'               Exit Sub
'             End If
'             On Error GoTo 0
'             PREVIE# = XVALO(!Precio_Unitario)
'             If Val(Label4) = 3 Then
'              'PREVIE# = CVD(Mid$(Y$, 62, 8))
'              If Abs(PREVIE# - PRENUEVO) < 0.00005 Or TRIM$(Mid$(y$, 46, 16)) = "" Then
'                Call REPLA(y$, HOS$, 46, 16)
'              End If
'             End If
'            '
'             '.Edit
'             !Precio_Unitario = PRENUEVO
'             .Update
'             'Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
'             'Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
'             'Call REPLA(Y$, MRX$, 44, 1)
'             'Call GRAREG(APRECIO$, Y$, U)
'             End With
'             Set ITEMSEL = Nothing
'           Next U&
'
'           Call CIERRARCH(APRECIO$)
'           Call BAJALDISCO
'           Screen.MousePointer = 1
'          '
'           Call GENERAR_DAT(LISNO%)
'           Call CARLISTA
'         '
'      End If
'    End If
'    Exit Sub
'
'
'End Sub
'
'
Sub APREPOR()
    '
    Dim VALVIE(32767), VALNUE(32767)
    HOII% = HOY(HOS$)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
    End If
    GoTo 10
    Exit Sub
    '
    ' SELECCIONA LA MODALIDAD DE ACTUALIZAR PRECIOS
    OPX% = COMALTER%("Puede Actualizar a Partir\de la Lista de Precios Actualmente Vigente\o de los Costos de Reposición para un Proveedor\\Lista Vigente\Costos de Reposición\Cancelar")
    
    Select Case OPX%
      Case 2
        Call APREPROV
        Exit Sub
      Case 3
        Exit Sub
    End Select
    '
    ' LISTACTIVA& = List1.ListIndex + 1
10  On Error Resume Next
    LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
    On Error GoTo 0
    'X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    RANGOAC% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
      TTXX$ = "Opciones de Actualización por Porcentaje:\\Sólo Bombas\Sólo Repuestos\Toda la Lista"
      RANGOAC% = COMALTER%(TTXX$)
      If RANGOAC% < 1 Or RANGOAC% > 3 Then Exit Sub
      If RANGOAC% = 3 Then RANGOAC% = 0
    End If
    '
    Select Case RANGOAC%
      Case 0
        RACTU$ = "Lista Completa"
      Case 1
        RACTU$ = "Bombas Nacionales (Cod. 0...)"
      Case 2
        RACTU$ = "Repuestos (Cod. >= 1...)"
    End Select
    '
    ' SELECCIONA RANGO DE CODIGOS
    Call DESHASCOD(COD1$, COD2$, CIKX1%, CIKX2%)
    If CIKX1% < 1 Or CIKX2% < 1 Then Exit Sub
    '
    '**********
    SQLX$ = "SELECT * FROM LISTAPRE "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
    'SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
    'Set LISTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    Dim LISTSELEC As ADODB.Recordset
    Set LISTSELEC = New ADODB.Recordset
    LISTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With LISTSELEC
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
         On Error GoTo 0
         Exit Sub
         
      End If
      On Error GoTo 0
      Call REPLA(VDEF$, Chr(XVALO(!Cod_Lista)), 1, 1)
      Call REPLA(VDEF$, !DENOM_LISTA, 2, 46)
      Call REPLA(VDEF$, Chr(XVALO(!Formato)), 48, 1)
      Call REPLA(VDEF$, Chr(XVALO(!DECIMALES)), 49, 1)
      Call REPLA(VDEF$, Chr(XVALO(!Moneda)), 50, 1)
      Call REPLA(VDEF$, !DESC_ABREV, 51, 12)
      'VDEF$ = Left$(X$, 64) + MKS$(0)
      Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
      'CADECI% = Asc(Mid$(VDEF$, 49, 1))
      CADECI% = !DECIMALES
      If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
      Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    End With
    
    ' INGRESO DEL PORCENTAJE
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    ' VALIDACIONES
    PORAJU = CVS(Mid$(OBX$, 65, 4)) 'PORCENTAJE DE AJUSTE
'    If Abs(PORAJU) < 0.001 Then Exit Sub 'COMENTADO PARA QUE PERMITA REALIZAR AJUSTE 0
    '
    If PORAJU < (-90) Or PORAJU > 99.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 0 Or CADECI% > 4 Then ' SE MODIFICO LA VALIDACION < 2 PARA QUE PERMITA REDONDEAR A ENTERO
       Call MENSERR(24, "Decimales entre '0' y '4'")
       GoTo 5
    End If
   '
   LISNO% = Asc(Left$(OBX$, 1))
   '
   Screen.MousePointer = 11
   '*************************
   '//OT-07-0434//
   '
   REBLA$ = REGBLAN$("ACOSPOR")
   ' LEEMOS LISTA DE PRECIOS ACTIVA
   SQLX$ = "SELECT * FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
   Dim ARTSELEC As ADODB.Recordset
   Set ARTSELEC = New ADODB.Recordset
   ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With ARTSELEC
     On Error Resume Next
     .MoveLast
     FIN& = .RecordCount
     If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Lista Vacía")
        Exit Sub
     End If
     On Error GoTo 0
     .MoveFirst
     U& = 0:   JJ& = 0
     Do While Not .EOF
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            TTXX$ = Space$(128)
            CII% = XVALO(!Cod_Pieza)
            CDXXX$ = CODEXT(XVALO(!Cod_Pieza))
            '
            
            PORAJX = PORAJU
            If RANGOAC% > 0 Then
              CDDX$ = Left$(CDXXX$, 1) '
              If RANGOAC% = 1 Then
                  If CDDX$ <> "0" Then PORAJX = 0
                ElseIf RANGOAC% = 2 Then
                  If CDDX$ = "0" Then PORAJX = 0
              End If
            End If
            '
            If UCase$(CDXXX$) < UCase$(COD1$) Or UCase$(CDXXX$) > UCase$(COD2$) Then
              PORAJX = 0
            End If
            '
            COSVIE# = XVALO(!Precio_Unitario)
            If CADECI% > 0 Then
              COSNUE = ROUND#(COSVIE * (100 + PORAJX) / 100, CADECI%)
            Else
              'CASO QUE SE PASE DECIMALES 0: POR QUE EL ROUND SI CADECI% = 0 APLICA 2 DECIMALES
              COSNUE = XVALO(FORMATNUM$((COSVIE * (100 + PORAJX) / 100), "F16.0"))
            End If
            '
            ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(CII%), 1, 2)
            Call REPLA(ZZ$, MKS$(COSVIE#), 3, 4)
            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
            
            JJ& = JJ& + 1
            Call GRAREG("!ACOSPOR", ZZ$, JJ&)
        .MoveNext
     Loop
   End With
   Set ARTSELEC = Nothing
   '//OT-07-0434//
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    Screen.MousePointer = 1
    
    ' PRESENTA PRECIOS ACTUALES Y NUEVOS
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + MSF1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + MSF1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         'GRABAMOS LOS CAMBIOS
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
             For U& = 1 To URE&
             Call TRACE(20, U&, URE&)
             X$ = REGLEIDO$("!ACOSPOR", U&)
             CII% = CVI(Mid$(X$, 1, 2))
             'Y$ = REGLEIDO$(APRECIO$, U&)
             'MRX$ = Mid$(Y$, 45, 1)
             PRENUEVO = CVS(Mid$(X$, 7, 4))
             PREX$ = COSNUEX$
             '
             SQLX$ = "SELECT * FROM Preventa "
             SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
             SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
             'Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
             Dim ITEMSEL As ADODB.Recordset
             Set ITEMSEL = New ADODB.Recordset
             ITEMSEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

             With ITEMSEL
              On Error Resume Next
             .MoveFirst
             FIN& = .RecordCount
             If Err > 0 Then
               On Error GoTo 0
               Call MENSERR(24, "Lista Vacía")
               Exit Sub
             End If
             On Error GoTo 0
             PREVIE# = XVALO(!Precio_Unitario)
             If Val(Label4) = 3 Then
              'PREVIE# = CVD(Mid$(Y$, 62, 8))
              If Abs(PREVIE# - PRENUEVO) < 0.00005 Or TRIM$(Mid$(Y$, 46, 16)) = "" Then
                Call REPLA(Y$, HOS$, 46, 16)
              End If
             End If
            '
             '.Edit
             !Precio_Unitario = PRENUEVO
             .Update
             'Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
             'Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
             'Call REPLA(Y$, MRX$, 44, 1)
             'Call GRAREG(APRECIO$, Y$, U)
             End With
             Set ITEMSEL = Nothing
           Next U&
           
           Call CIERRARCH(APRECIO$)
           Call BAJALDISCO
           Screen.MousePointer = 1
          '
           Call GENERAR_DAT(LISNO%)
           Call CARLISTA
         '
      End If
    End If
    Exit Sub


End Sub



Sub APREPROV()
    '//OT-07-0434//
    'LISTACTIVA& = List1.ListIndex + 1
    On Error Resume Next
    LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
    On Error GoTo 0
    X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    Call SELECHO("ACREDOR")
    VDEVU$ = VALACT1$("ACREDOR")
    NPRO% = Val(VDEVU$)
    If NPRO% < 1 Then Exit Sub
    '
    
    RACTU$ = rasocli$((-1) * NPRO%)
    SQLX$ = "SELECT * FROM LISTAPRE "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
    'Set LISTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    Dim LISTSELEC As ADODB.Recordset
    Set LISTSELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With LISTSELEC
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
         On Error GoTo 0
         Exit Sub
      End If
    
      Call REPLA(VDEF$, Chr(XVALO(!Cod_Lista)), 1, 1)
      Call REPLA(VDEF$, !DENOM_LISTA, 2, 46)
      Call REPLA(VDEF$, Chr(XVALO(!Formato)), 48, 1)
      Call REPLA(VDEF$, Chr(XVALO(!DECIMALES)), 49, 1)
      Call REPLA(VDEF$, Chr(!Moneda), 50, 1)
      Call REPLA(VDEF$, !DESC_ABREV, 51, 12)
     'VDEF$ = Left$(X$, 64) + MKS$(0)
      Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
      CADECI% = XVALO(!DECIMALES)
      On Error GoTo 0

    'VDEF$ = Left$(X$, 64) + MKS$(0)
    Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
    'CADECI% = Asc(Mid$(VDEF$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
    Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    '
    End With
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 65, 4))
    'If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 299.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
    '
'    'LISNO% = Asc(Left$(OBX$, 1))
'    'ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
'                                     ALISTA$ = "0" + ALISTA$
'                                   Wend
                                   '
'    APRECIO$ = "LIPRE" + ALISTA$
'    Call COPYSTRU("LIPREVEN", APRECIO$)
'    REBLA$ = REGBLAN$("ACOSPOR")
'    FIN& = ULTREG&(APRECIO$)
    JJ& = 0
    '
    Screen.MousePointer = 11
    PUCOD$ = ""
    'For U& = 1 To FIN&
       ' LEEMOS LISTA DE PRECIOS ACTIVA
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
     'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
     Dim ARTSELEC As ADODB.Recordset
     Set ARTSELEC = New ADODB.Recordset
     ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With ARTSELEC
       On Error Resume Next
       .MoveLast
       FIN& = .RecordCount
       If Err > 0 Then
          On Error GoTo 0
          Call MENSERR(24, "Lista Vacía")
          Exit Sub
       End If
       On Error GoTo 0
       .MoveFirst
       U& = 0:   JJ& = 0
       Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          'I0$ = REGLEIDO$(APRECIO$, U&)
          'CIXI% = CVI(Left$(I0$, 2))
          CIXI% = XVALO(!Cod_Pieza)
          CDXXX$ = CODEXT$(CIXI%)
          '
          'COSVIE = CVD(Mid$(I0$, 62, 8))
          COSVIE = XVALO(!Precio_Unitario)
          COSNUE = COSVIE
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
          Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
          Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
          JJ& = JJ& + 1
          Call GRAREG("!ACOSPOR", ZZ$, JJ&)
            PUCOD$ = PUCOD$ + MKI$(CIXI%)
          .MoveNext
       Loop
    End With
    '
    AGREGAIT% = 0
    For CIXI% = 1 To NUMAS%
      If PROVHABI%(CIXI%) = NPRO% Then
        CosRep = COSPESOS(CIXI%)
        COSNUEX$ = CSTRING$(MKS$(CosRep * (100 + PORAJU) / 100), 3, 12, CADECI%, 2)
        COSNUE = Val(COSNUEX$)
        '
        For U& = 1 To JJ&
          If CVI(Mid$(PUCOD$, 2 * U& - 1, 2)) = CIXI% Then
            ZZ$ = REGLEIDO$("!ACOSPOR", U&)
            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
            Call GRAREG("!ACOSPOR", ZZ$, U&)
            GoTo 29
          End If
        Next U&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
        Call REPLA(ZZ$, MKS$(0), 3, 4)
        Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
        JJ& = JJ& + 1
        Call GRAREG("!ACOSPOR", ZZ$, JJ&)
        PUCOD$ = PUCOD$ + MKI$(CIXI%)
        AGREGAIT% = 1
      End If
29  Next CIXI%
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    If AGREGAIT% > 0 Then
      Call COMUNI("Los Items Agregados Aparecen\al FINAL de la Lista.\ \Luego de Guardar los Nuevos Precios\Puede Realizar Ordenamiento Automatico.")
    End If
    Screen.MousePointer = 1
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + MSF1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + MSF1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
         'UREA& = ULTREG&(APRECIO$)
         'LISNO% = Asc(Left$(OBX$, 1))
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           X$ = REGLEIDO$("!ACOSPOR", U&)
           CII% = CVI(Mid$(X$, 1, 2))
           SQLX$ = "SELECT * FROM Preventa "
           SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
           SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
           'Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
           Dim ITEMSEL As ADODB.Recordset
           Set ITEMSEL = New ADODB.Recordset
           ITEMSEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

           With ITEMSEL
             On Error Resume Next
             .MoveFirst
             If Err > 0 Then
               On Error GoTo 0
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CODIEMPR = CodiEmp%
               On Error GoTo 0
               !Cod_Lista = LISTACTIVA&
               !Cod_Pieza = CII%
               !Desc_Pieza = DESCRIT0$(CII%)
               !Denom_Pieza = DESCRIT0$(CII%)
             Else
                '.Edit
             End If
             On Error GoTo 0
  '           If U& <= UREA& Then
  '               Y$ = REGLEIDO$(APRECIO$, U&)
  '             Else
  '               CIXI% = CVI(Left$(X$, 2))
  '               Y$ = REGBLAN$(APRECIO$)
  '               Call REPLA(Y$, MKI$(CIXI%), 1, 2)
  '               Call REPLA(Y$, DESCRIT0$(CIXI%), 3, 42)
  '           End If
             'MRX$ = Mid$(Y$, 45, 1)
             If Not IsNull(!Carac_Control) Then
              MRX$ = !Carac_Control
              Else: MRX$ = ""
             End If
             COSNUEX$ = CSTRING$(Mid$(X$, 7, 4), 3, 12, CADECI%, 2)
             PRENUEVO# = Val(COSNUEX$)
             PREX$ = COSNUEX$
             '
             If Val(Label4) = 3 Then
               'PREVIE# = CVD(Mid$(Y$, 62, 8))
               PREVIE# = XVALO(!Precio_Unitario)
               If Abs(PREVIE# - PRENUEVO#) < 0.00005 Or TRIM$(XVALO(!Num_Original)) = "" Then
                 'Call REPLA(Y$, HOS$, 46, 16)
                 !Num_Original = HOS$
               End If
             End If
             '
             'Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
             !Precio_Unitario = PRENUEVO#
             'Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
             'Call REPLA(Y$, MRX$, 44, 1)
             !Carac_Control = MRX$
             .Update
           End With
           'Call GRAREG(APRECIO$, Y$, U&)
         Next U&
         Set ITEMSEL = Nothing
         '
         'Call SETULTREG(APRECIO$, U&)
         'Call CIERRARCH(APRECIO$)
         'Call BAJALDISCO
         Screen.MousePointer = 1
         '
         '//OT-07-0434//
         NLIST% = LISTACTIVA&
         Call GENERAR_DAT(NLIST%)
         Call CARLISTA
         If AGREGAIT% > 0 Then
           If COMALTER%("Desea Ordenar por Codigo Creciente ?\\No, Dejar\Si, Ordenar") = 2 Then
             Call Command12_Click
           End If
         End If
         '
      End If
    End If
    '

End Sub


'
'
Sub LISTAPRE()
        '//OT-07-0434//
2000    ' ************** impresion de listas de precios
        '
        LISTACTIVA& = List1.ListIndex + 1
        X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
        TITU$ = TRIM$(Mid$(X$, 2, 46))
        '
        Call DESHASCOD(COD1$, COD2$, CIK1%, CIK2%)
        If CIK1% < 1 Or CIK2% < 1 Then Exit Sub
        '
2100    Call BLANARCH("LIPREVEN")
        Screen.MousePointer = 11
        '
        Call HEADERS("LIPREVEN", "")
        Call HEADERS("LIPREVEN", "Lista: " + TITU$)
        '
        ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        LIPREAC$ = "LIPRE" + ARCHIPRE$
        FIN& = ULTREG&(LIPREAC$)
        '
        'CONSULTA ARTICULOS SEGUN LISTA SELECCIONADA
        SQLX$ = "SELECT * FROM Preventa "
        SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & ""
        SQLX$ = SQLX$ + " ORDER BY Num_Orden;"
        'Set ARTSELE = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim ARTSELE As ADODB.Recordset
        Set ARTSELE = New ADODB.Recordset
        ARTSELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

        On Error Resume Next
        With ARTSELE
          .MoveLast
          If Err > 0 Then
             Exit Sub
          End If
          FIN& = .RecordCount
          .MoveFirst
          If FIN& > 0 Then
          'Recorremos el recordset y selecionamos solo
          'aquellos que cumplen con los parametros seleccionados
          'en la ventana  DESHASCOD donde se selecciona desde y hasta
            Do While (.EOF = False)
                CII% = XVALO(!Cod_Pieza)
                CDXX$ = CODEXT$(XVALO(!Cod_Pieza))
                If UCase(CDXX$) >= UCase(COD1$) Then
                  If UCase(CDXX$) <= UCase(COD2$) Then
                    TTXX$ = Space$(128)
                    Call REPLA(TTXX$, MKI$(CII%), 1, 2)
                    Call REPLA(TTXX$, DESCRIT$(CII%), 3, 42)
                    Call REPLA(TTXX$, !Denom_Pieza, 45, 0)
                    Call REPLA(TTXX$, !Carac_Control, 45, 1)
                    Call REPLA(TTXX$, !Num_Original, 46, 16)
                    PRUNI# = XVALO(!Precio_Unitario)
                    Call REPLA(TTXX$, MKD$(PRUNI#), 62, 8)
                    Call REPLA(TTXX$, MKS$(!Cant_Limite), 70, 2)
                    Call REPLA(TTXX$, MKD$(!Porc_Limite), 72, 1)
                    JJ& = JJ& + 1
                    Call REGAPP("LIPREVEN", TTXX$)
                  End If
                End If
                .MoveNext
            Loop
          End If
        End With
                                

              
              
        '***
        'For I& = 1 To FIN&
          'Call TRACE(20, I&, FIN&)
          'NXX$ = REGLEIDO$(LIPREAC$, I&)
          'CI% = CVI(Left$(NXX$, 2))
          'If CI% > 0 Then
            'If CI% <= NUMAS% Then
              'CDXX$ = CODEXT$(CI%)
              'If CDXX$ >= COD1$ Then
                'If CDXX$ <= COD2$ Then
                  'RELIB% = Val(Mid$(NXX$, 45, 1))
                  '
                  'For K% = 1 To RELIB%
                    'Call REGAPP("LIPREVEN", MKI$(32767) + Mid$(REGBLAN$("LIPREVEN"), 3))
                  'Next K%
                  'Call REGAPP("LIPREVEN", NXX$)
                  '
                'End If
              'End If
            'End If
          'End If
        'Next I&
        '
        'Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*LILIPRE") '
        '
        '//OT-07-0434//
End Sub

Sub SORLIPRE()
    '//OT-07-0434//
    ' NUMERO DE LA LISTA ACTIVA
    LISTACTIVA& = List1.ListIndex + 1
    'X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    'ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
    'While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
    'LIPREAC$ = "LIPRE" + ARCHIPRE$
    '
    '    Call DATLIPRE(X$)
        '
    'TOMA EL VALOR DEL ULTIMO REGISTRO DE LA LISTA ACTIVA
    SQLX$ = "SELECT * FROM Preventa "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " ;"
    'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    Dim ARTSELEC As ADODB.Recordset
    Set ARTSELEC = New ADODB.Recordset
    ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ARTSELEC
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err > 0 Then
         On Error GoTo 0
         Call MENSERR(24, "Lista Vacía")
         Exit Sub
      End If
      On Error GoTo 0
      .MoveFirst
    'FIN& = ULTREG&(LIPREAC$)
    End With
    If FIN& < 1 Then
       Exit Sub
    End If
    'Call CIERRARCH(LIPREAC$)
    'LDT$ = LUDAT$ 'ESTO QUEDA ?????
    '
    'Call REGAPP(LIPREAC$, String$(128, 0))
    'Call CIERRARCH("*.*")
    'Call TRALOCAL(LIPREAC$, "LIPREVEN")
    '
    Screen.MousePointer = 11
    SORTLIST.Clear
    '
    'LEE LA LISTA ,Y COPIA EN SORTLIST
    For i& = 1 To FIN&
      With ARTSELEC
        Call TRACE(20, i&, FIN&)
        'NXX$ = REGLEIDO$(LIPREAC$, I&)
        'CI% = CVI(Left$(NXX$, 2)) 'CAPTURO EL CODIGO DE ACCESS
        CII% = XVALO(!Cod_Pieza)
        CODE$ = CODEXT$(CII%) 'OBTENGO EL CODIGO EXTERNO
        If CODE$ <> "" Then ' =
           Z$ = Space$(24)
           Call REPLA(Z$, "A" + CODE$, 1, 18) '
           Call REPLA(Z$, TRIM$(Str$(i&)), 19, 6)
           SORTLIST.AddItem Z$
           .MoveNext
        End If
      End With
    Next i&
    '
    ARTSELEC.Close
    Set ARTSELEC = Nothing
    FIN& = SORTLIST.ListCount
    '
    'J& = 0
    'LEE DEL SORTLIST Y GRABA EN LIPREAC
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       TESOR$ = SORTLIST.List(i& - 1)
       REG& = Val(Mid$(TESOR$, 19, 6))
       CODE$ = Mid$(TESOR$, 2, 17)
       CII% = CODINT%(CODE$)
       'PARA PREVENIR EN CASO QUE EL CODIGO ESTE DE BAJA
       'Y LO DEVUELVA EN NEGATIVO
       CII% = Abs(CII%)
       SQLX$ = "SELECT * FROM Preventa "
       SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA & " "
       SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
       'Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
       Dim ITEMSEL As ADODB.Recordset
       Set ITEMSEL = New ADODB.Recordset
       ITEMSEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

       With ITEMSEL
          On Error Resume Next
          .MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 10
          End If
          On Error GoTo 0
          'If REG& > 0 Then
          '.Edit
          !Num_Orden = i&
          .Update
             'NXX$ = REGLEIDO$("!LIPREVEN", REG&)
             'J& = J& + 1
             'Call GRAREG(LIPREAC$, NXX$, J&) 'REEMPLAZO ESTO GRABANDO EN ACCESS LOS DATOS DE NXX$
       'End If                               '
       End With
       ITEMSEL.Close
       Set ITEMSEL = Nothing
10  Next i&
    '
    'Call SETULTREG&(LIPREAC$, J&)
    'Call CIERRARCH("*.*")
    NLIST% = LISTACTIVA&
    Call GENERAR_DAT(NLIST%)
    Call CARLISTA
    Screen.MousePointer = 1
    '
    '//OT-07-0434//
    Call COMUNI("Ordenamiento Completo.")
    '
End Sub

Sub CARILISTAS()
   '
   List1.Clear
 
   UMAX& = 0
   Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT * FROM LISTAPRE ORDER BY Cod_Lista ASC"
   Dim ListaPrec As ADODB.Recordset
   Set ListaPrec = New ADODB.Recordset
   ListaPrec.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   On Error Resume Next
   With ListaPrec
     .MoveLast
     If Err > 0 Then
        Exit Sub
     End If
     .MoveFirst
     For U& = 1 To .RecordCount
         Call REPLA(NXX$, Chr(!Cod_Lista), 1, 1)
         Call REPLA(NXX$, !DENOM_LISTA, 2, 46)
         Call GRAREG("LISTAS", NXX$, U&): UMAX& = U&
         CodList$ = .Fields("Cod_Lista")
         DenomLista$ = .Fields("denom_Lista")
         List1.AddItem AJUSTD$(TRIM$(CodList$), 3) + " - " + TRIM$(DenomLista$)
         .MoveNext
     Next U&
   End With
   ListaPrec.Close
   Set ListaPrec = Nothing
   '
   List1.Refresh
   Call SETULTREG("LISTAS", UMAX&)
   '
End Sub

Function VERILLAVE%()
   '
VERILLAVE% = 1
Exit Function

   VERILLAVE% = 0
10 OPX% = COMALTER%("Coloque Diskette Llave\\Continuar\Cancelar")
   If OPX% = 1 Then
      On Error GoTo 50
      SVL$ = SERIVOL$("A:")
      TECONTROL$ = ENCRIPTA$(TRIM$(UCase$(App.EXEName)) + SVL$)
      NX1% = FreeFile
      Dim TTX As String * 128
      Open "A:\" + TRIM$(UCase$(App.EXEName)) + ".KEY" For Random Access Read Write As #NX1% Len = 128
      Get #NX1%, 2, TTX$
      LTC% = Len(TECONTROL$)
VERILLAVE% = 1: GoTo 99
      If CVI(Mid$(TTX$, 18, 2)) = LTC% Then
        If Mid$(TTX$, 30, LTC%) = TECONTROL$ Then
          VERILLAVE% = 1
          GoTo 99
        End If
      End If
      Call COMUNI("Llave no Válida")
   End If
   GoTo 99
   '
50 Resume 51
51 GoTo 10
   '
99 End Function

Sub EDITLISTA(ARCHIPRE$)
   '
   'ARMA EL ENCABEZADO DE VENTABU
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + MSF1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + MSF1.Left))
   '
   'ABRE EL VENTABU
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% >= 0 Then
     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     MSF1.Rows = 1
     msf2.Rows = 1
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     ' SI  LABEL 14 (FORMATO) ES IGUAL A 3 ,SI NO GRABA DIRECTAMENTE
     'preguntar exactamente que hace
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For i& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", i&)
         Ci% = CVI(Left$(NXX$, 2))
         If Ci% > 0 Then
           If Ci% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
             RFF$ = RECFILT$(ARCHIPRE$, 1, MKI$(Ci%))
             If Len(RFF$) >= 4 Then
               REANT& = CVS(RFF$)
               If REANT& > 0 Then
                 NYY$ = REGLEIDO$(ARCHIPRE$, REANT&)
                 If Abs(CVD(Mid$(NYY$, 62, 8)) - PRELI#) < 0.00005 Then
                   If TRIM$(Mid$(NYY$, 46, 16)) <> "" Then
                     GoTo 89
                   End If
                 End If
               End If
             End If
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, i&)
89     Next i&
     End If
     '
     'RECORRE EL LIPREVEN Y GRABA EN ARCHIPRE$
     SQLX$ = "UPDATE PREVENTA SET MARBORRA = 1 "      'MARCA MARBORRA EN 1 PARA BORRAR LUEGO DE REALIZAR EL UPDATE
     SQLX$ = SQLX$ + " WHERE Cod_Lista = " & LISTACTIVA& & ""
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     Dim PREVVVENTA As ADODB.Recordset
     Set PREVVVENTA = New ADODB.Recordset
     SQLX$ = "SELECT * FROM PREVENTA"
     PREVVVENTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     '
     For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", i&)
        Ci% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If Ci% > 0 Then
          If Ci% <= NUMAS% Then
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))

            With PREVVVENTA
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !Cod_Lista = LISTACTIVA&
              !Cod_Pieza = Ci%
              !Num_Orden = i&
              !Desc_Pieza = DESCRIT0$(Ci%)
              !Denom_Pieza = DESCRIT0$(Ci%)
              !Carac_Control = Mid$(NXX$, 45, 1)
              !Num_Original = (Mid$(NXX$, 46, 16))
              !Precio_Unitario = CVD(Mid$(NXX$, 62, 8))
              !Cant_Limite = CVD(Mid$(NXX$, 70, 2))
              !Porc_Limite = CVD(Mid$(NXX$, 72, 1))
              .Update
            End With
          End If
        End If
     Next i&
     Set PREVVVENTA = Nothing
     '
     SQLX$ = "DELETE FROM PREVENTA "  ' BORRRAR LOS REGISTROS MARCADOS CUYO CAMPO MARBORRA ESTA EN 1
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " AND MARBORRA = 1"
      FLEXCONN.Execute (FILTSQL$(SQLX$))
     '
     Screen.MousePointer = 1
     '
     NLIST% = LISTACTIVA&
     Call GENERAR_DAT(NLIST%)
     Call CARLISTA
   End If

End Sub
'
Sub GENERAR_DAT(n%)
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Screen.MousePointer = 11
   ARCHIPRE$ = TRIM$(Str$(n%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   CADECI% = XVALO(Text3): If CADECI% > 4 Then CADECI% = 4
   '
   Call COPYSTRU("LIPREVEN", ARCHIPRE$)
   LF2% = LENFIELD(ARCHIPRE$, 2)
   '
   Call APERBASDAT("PEDCLI")
   CONDI$ = "Cod_Lista = " & n% & " "
   SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim ARTSELEC As ADODB.Recordset
   Set ARTSELEC = New ADODB.Recordset
   Set ARTSELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))

   FIN& = CantRegistros("PREVENTA", CONDI$)
   U& = 0
   On Error Resume Next
   With ARTSELEC
     If FIN& > 0 Then
        Do While (.EOF = False)
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            TTXX$ = REGBLAN$(ARCHIPRE$)
            CII% = XVALO(!Cod_Pieza)
            Call REPLA(TTXX$, MKI$(CII%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CII%), 3, LF2%)
               MRX$ = SAFETEXT$(!Carac_Control)
            Call REPLA(TTXX$, MRX$, 3 + LF2%, 1)
            Call REPLA(TTXX$, !Num_Original, 4 + LF2%, 16)
               PRUNI# = XVALO(!Precio_Unitario)
            Call REPLA(TTXX$, MKD$(PRUNI#), 20 + LF2%, 8)
            Call REPLA(TTXX$, MKS$(!Cant_Limite), 28 + LF2%, 2)
            Call REPLA(TTXX$, MKD$(!Porc_Limite), 30 + LF2%, 1)
            '
            PRELI# = PRUNI#
            CDC$ = TRIM$(Str$(CADECI%))
            PREX$ = " " + FORMATNUM$(PRELI#, "F12." + CDC$) + "  "
            Call REPLA(TTXX$, PREX$, 1 + LF2% - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(TTXX$, MRX$, 2 + LF2%, 1)
            '
            If INCLUALIAS% > 0 Then
              ALX$ = TRIM$(NUPLANO$(CII%))
              If ALX$ <> "" Then
                Call REPLA(TTXX$, " (" + ALX$ + ")", LF2% - 16 - Len(ALX$), 3 + Len(ALX$))
              End If
            End If
            JJ& = JJ& + 1
            Call GRAREG(ARCHIPRE$, TTXX$, JJ&)
            '
        .MoveNext
        Loop
        Call SETULTREG(ARCHIPRE$, JJ&)
        Call CIERRARCH(ARCHIPRE$)
        Call FRESHECHO(ARCHIPRE$)
      End If
   End With
   Set ARTSELEC = Nothing
   Screen.MousePointer = 1
   '
End Sub
'
'
Sub LeeLaLista(n%)
   '//OT-07-0434//
   '
   Call SETULTREG("!LIPREVEN", 0)
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT * FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & n% & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim ARTSELEC As ADODB.Recordset
   Set ARTSELEC = New ADODB.Recordset
   ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   On Error Resume Next
   With ARTSELEC
     .MoveLast
     If Err > 0 Then
        Exit Sub
     End If
     FIN& = .RecordCount
     .MoveFirst
     If FIN& > 0 Then
        Do While (.EOF = False)
            TTXX$ = Space$(128)
            CII% = XVALO(!Cod_Pieza)
            Call REPLA(TTXX$, MKI$(CII%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CII%), 3, 42)
            Call REPLA(TTXX$, !Carac_Control, 45, 1)
            Call REPLA(TTXX$, !Num_Original, 46, 16)
               PRUNI# = XVALO(!Precio_Unitario)
            Call REPLA(TTXX$, MKD$(PRUNI#), 62, 8)
            Call REPLA(TTXX$, MKS$(!Cant_Limite), 70, 2)
            Call REPLA(TTXX$, MKD$(!Porc_Limite), 72, 1)
            JJ& = JJ& + 1
            '****
            Call GRAREG("!LIPREVEN", TTXX$, JJ&)
            

        .MoveNext
        Loop
        Call SETULTREG("!LIPREVEN", JJ&)
        Call CIERRARCH("!LIPREVEN")
      End If
     End With
     Set ARTSELEC = Nothing
     '//OT-07-0434//
  
End Sub

Private Sub mnuActualizacionporcent_Click()
   Call Command11_Click
End Sub

Private Sub mnuActulistaxg_Click()
   If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") < 2 Then Exit Sub
   '
    Call VERIFICA_MODIFICACION
    
20  GRACT% = SELECGRUPOACTU
    If GRACT% < 1 Then Exit Sub
    SELFAMILI.mnuActuCostos.Visible = True
    SELFAMILI.TXTBUSCAR.Enabled = False
    SELFAMILI.CMBFILTRO.Enabled = False
    SELFAMILI.Width = 15000
    'datos para el frm de seleccion de pedidos
    Campos$ = "CODINT/A0;Código/A20;Descripción/A36;Cód.Grupo/F10;Referencia/A20;P.Venta/F12.4"
    Call CARGA_ENCABEZADO(SELFAMILI.GRID, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(SELFAMILI.MSF_2, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    SELFAMILI.Command3.Enabled = False
    '
'    'CARGO UN VECTOR CON LA REFERENCIA DE LA MONEDA
    Dim REFMON$(6)
    For U& = 1 To ULTREG&("TAMONED")
       XX$ = REGLEIDO$("TAMONED", U&)
       CTG% = U&
       REFMON$(CTG%) = TRIM$(Mid$(XX$, 2, 24))
    Next U&
    If TRIM$(REFMON$(1)) <> "" Then
      REFMON$(0) = REFMON$(1)
    Else
      REFMON$(0) = "Pesos"
    End If
    '
    Dim obj As New RECORXET
    Set RS1 = obj.RS_LISTASPRECIOS()
    Call Carga_MSF_Recordset(RS1, "Cod_Lista/F12;Referencia/A48", FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1
    FRMZELECHO.Width = 9500
    FRMZELECHO.Caption = "Seleccione Lista de Precios a Actualizar"
    FRMZELECHO.Show 1
    LP% = XVALO(RESP_ZELE_VECT(1))

    If LP% < 1 Then GoTo 20

'    Dim obj As New RECORXET
    Set RS1 = obj.RS_PREVTA(LP%, GRACT%)
    Call Carga_MSF_Recordset(RS1, Campos$, SELFAMILI.GRID, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1
    '
    'LE AGREGO LA REFERENCIA DEL GRUPO DIRECTAEMNTE A LA GRILLA.
    For i& = 1 To SELFAMILI.GRID.Rows - 1
     PRE = XVALO(SELFAMILI.GRID.TextMatrix(i&, 6))
        GRACT1% = XVALO(SELFAMILI.GRID.TextMatrix(i&, 4))
        SELFAMILI.GRID.TextMatrix(i&, 5) = ""
        If GRACT1% > 0 Then
           SELFAMILI.GRID.TextMatrix(i&, 5) = REFEGRUPOACTU$(GRACT1%)
        End If
    Next i&
    SELFAMILI.Height = 900 + SELFAMILI.GRID.Height
    SELFAMILI.mnuActuCostos.Caption = "Actualizar Lista de Precios N°: " & LP%
    SELFAMILI.Caption = "Items de Lista N°: " & LP% & " Pertenecientes al Grupo de Actualización : " & REFEGRUPOACTU$(GRACT%)
50  SELFAMILI.Show 1
    
    If SELFAMILI.APROBO_ACTUALIZACION% > 0 Then
      SELFAMILI.APROBO_ACTUALIZACION% = 0
      PCT1$ = InputBox("Ingrese Porcentaje", "Porcentaje Ingresado", "")
      PCT1texto$ = REPLACAR(PCT1$, ",", ".")
      PCT# = ROUND#(XVALO(PCT1texto$), 2)
      If Abs(PCT#) = 0 Then
        Call COMUNI("Porcentaje Ingresado No Válido")
        GoTo 50
      End If
      '
      J& = 0
      Dim MSF As msFlexGrid
      Set MSF = SELFAMILI.GRID
      OPX% = COMALTER%("Confirma Actualizar los Precios de Venta de los Items Seleccionados\\Cancelar\Actualizar Precios de Venta")
      If OPX% = 2 Then
          Screen.MousePointer = 11
          For i& = 1 To MSF.Rows - 1
            MARCA$ = TRIM$(MSF.TextMatrix(i&, 0))
            If MARCA$ = "*" Then
               CI1% = XVALO(MSF.TextMatrix(i&, 1))
               Pvta# = XVALO(MSF.TextMatrix(i&, 6))
               NUEVOPvta# = Pvta# + (Pvta# * PCT# / 100)
'               REFMONEDAX$ = MSF.TextMatrix(I&, 6) 'ACA ESTA EL TEXTO DE LA REFERENCIA DE LA MONEDA
               FCOS% = HOY(HO$)
               FECOS = CVDAT(FCOS%)
               SQLX$ = "UPDATE PREVENTA SET Precio_Unitario = '" & FORMATNUM$(NUEVOPvta#, "F16.4") & "'"
               SQLX$ = SQLX$ + " WHERE Cod_Pieza =  " & CI1%
               SQLX$ = SQLX$ + " AND  Cod_Lista =  " & LP%
               
               FLEXCONN.Execute (SQLX$)
               

'               Call ADD_NOVEDAD(CI1%, PROVHABI%(CI1%), Pvta#, NUEVOPvta#, "Cambio Masivo Por Grupo de Actualización", MON%, "COSTO")
               J& = J& + 1
            End If
          Next i&
          Screen.MousePointer = 1
          If J& > 0 Then
             Call GENERAR_DAT(LP%)
             Call COMUNI("Se Han Actualizado " & J& & " Items")
             LISTACTIVA& = XVALO(Left$(List1.TEXT, 3))
             List1.ListIndex = LP% - 1
             Call CARLISTA
          Else
             Call COMUNI("No Se Realizaron Actualizaciones!! \Debe Seleccionar Los Items a Actualizar")
          End If
      End If
      GoTo 20
    Else
      GoTo 20
    End If
   

End Sub

Private Sub mnubuscar_Click()
   Call Command20_Click
End Sub

Private Sub MnuGrabar_Click()
   Call Command3_Click
End Sub

Private Sub mnuimporta_Click()
    OPLIPREC07.Command7_Click
End Sub

Private Sub mnuListadePrecios_Click()
   Call Command8_Click
End Sub

Private Sub mnuNew_Click()
   Call Command7_Click
End Sub

Private Sub mnuNuevoItem_Click()
   Call Command21_Click
End Sub

Private Sub mnuReunir_Click()
   OPLIPREC07.Command1_Click
End Sub

Private Sub mnusort_Click()
   Call Command12_Click
End Sub

Private Sub mnuTablaAux_Click()
  Call VERIFICA_MODIFICACION
  
  If DERACCE%("TABLAUXI", "MODIFICAR TABLAS AUXILIARES") < 2 Then Exit Sub
  
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2, 1)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub mnuVerAntlp_Click()
'   1.-SELECCIONAR LISTA DE PRECIOS********************
'   2.-MOSTRAR VERSIONES DE ESA LISTA ENUMERADAS Y CON ELNRO DE MODIFCACION OCULTO
'   3.-PRESENTAR LAS MISIMAS EN UNA GRILLA GRAL
    Call VERIFICA_MODIFICACION

    Dim obj As New RECORXET
10  Set RS1 = obj.RS_LISTASPRECIOS()
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cod_Lista/F12;Referencia/A48", FRMZELECHO)
    Call Carga_MSF_Recordset(RS1, "Cod_Lista/F12;Referencia/A48", FRMZELECHO.msf2, 1)
    Set obj = Nothing
    FRMZELECHO.Width = 9500
    FRMZELECHO.Caption = "Seleccionar Lista de Precios"
    FRMZELECHO.Show 1
    LP% = XVALO(RESP_ZELE_VECT(1))
    RS1.Close
    Set RS1 = Nothing
    If LP% < 1 Then Exit Sub
    '
    Campos$ = "ID/A0;FECHA/D8;HORA/a16;IDENTIFICADOR/f12"
    i& = 0
20  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    SQLX$ = "SELECT DISTINCT  FECHMOD,IDMODIFIC,HORA FROM preventa_resp WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COD_LISTA = " & LP%
    SQLX$ = SQLX$ + " ORDER BY IDMODIFIC DESC"
    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
         i& = i& + 1
         FRMZELECHO.msf2.Rows = i& + 1
         FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!IDMODIFIC)
         FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!FECHMOD)
         FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!HORA)
         FRMZELECHO.msf2.TextMatrix(i&, 4) = i&
         .MoveNext
       Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    FRMZELECHO.Show 1
    ID_MODIF = XVALO(RESP_ZELE_VECT(1))
    FECHA$ = RESP_ZELE_VECT(2)
    HORA$ = RESP_ZELE_VECT(3)
    FECHAYHORA$ = FECHA$ & " - " & HORA$
    If ID_MODIF < 1 Then GoTo 10
    '
    Campos$ = "N.Lis/F6;Código/A16;Nro.Orig/A24;P.Vta./F14.4"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    SQLX$ = "SELECT P.Cod_Pieza,M.CODIGO,P.Num_Original, P.Precio_Unitario "
    SQLX$ = SQLX$ + " FROM  preventa_resp P WITH(NOLOCK) INNER JOIN MASTER M ON M.CODINT = P.Cod_Pieza"
    SQLX$ = SQLX$ + " WHERE COD_LISTA = " & LP%
    SQLX$ = SQLX$ + " AND IDMODIFIC = " & ID_MODIF
    
    SQLX$ = SQLX$ + " ORDER BY NUM_ORDEN "
    i& = 0
    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
         i& = i& + 1
         GRILLAGRAL.GRID.Rows = i& + 1
         GRILLAGRAL.GRID.TextMatrix(i&, 1) = LP%
         GRILLAGRAL.GRID.TextMatrix(i&, 2) = SAFETEXT$(!Codigo)
         GRILLAGRAL.GRID.TextMatrix(i&, 3) = SAFETEXT$(!Num_Original)
         GRILLAGRAL.GRID.TextMatrix(i&, 4) = ROUND(XVALO(!Precio_Unitario), 4)
         .MoveNext
       Loop
    End With
    RS.Close
    Set RS = Nothing
    GRILLAGRAL.Caption = "Lista de Precios N°: " & LP% & " (" & FECHAYHORA$ & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Visible = False
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call numerargrilla(GRILLAGRAL.GRID, 0)
    '
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "RESPLPRE", "Lista de Precios N°: " & LP% & " (" & FECHAYHORA$ & ")", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

    
    
End Sub

Private Sub MSF1_DblClick()

   FRMEDIT.TXTCODIGO.TEXT = MSF1.TextMatrix(MSF1.Row, 1)
   FRMEDIT.TXTDESCRIPCION.TEXT = MSF1.TextMatrix(MSF1.Row, 2)
   FRMEDIT.Text5.TEXT = MSF1.TextMatrix(MSF1.Row, 3)
   FRMEDIT.TXTNROORIGINAL = MSF1.TextMatrix(MSF1.Row, 4)
   FRMEDIT.txtlista.TEXT = MSF1.TextMatrix(MSF1.Row, 5)
   '
   txtlista.SetFocus
   '
   FRMEDIT.Show 1
   '
   If TRIM$(FRMEDIT.TXTCODIGO) = "" Then Exit Sub
   '
   MODIFICA% = 1
   '
   If FRMEDIT.MODO1$ = "ACTUALIZAR" Then
        If XVALO(Label4) <= 1 Then FRMEDIT.TXTNROORIGINAL = Left(TXTNROORIGINAL.TEXT, 1)
        If XVALO(Label4) = 2 Then FRMEDIT.TXTNROORIGINAL = FORMATNUM$(TXTNROORIGINAL.TEXT, "f5.2")
        If XVALO(Label4) = 3 Then FRMEDIT.TXTNROORIGINAL = FECHATEX$(HOY(HO$))

        If MSF1.Visible = True Then
           MSF1.TextMatrix(MSF1.Row, 3) = FRMEDIT.Text5.TEXT
           MSF1.TextMatrix(MSF1.Row, 4) = FRMEDIT.TXTNROORIGINAL
           MSF1.TextMatrix(MSF1.Row, 5) = FRMEDIT.txtlista.TEXT
        Else
           msf2.TextMatrix(MSF1.Row, 3) = FRMEDIT.Text5.TEXT
           msf2.TextMatrix(MSF1.Row, 3) = FRMEDIT.TXTNROORIGINAL
           msf2.TextMatrix(MSF1.Row, 3) = FRMEDIT.txtlista.TEXT
           Call ACTUALIZA_MSF1(CODINT(msf2.TextMatrix(MSF1.Row, 1)))
        End If
   ElseIf FRMEDIT.MODO1$ = "ELIMINAR" Then
        If MSF1.Visible = True Then
            Call borrafila(MSF1.Row, Me, MSF1)
        Else
            Call BORRA_MSF1(CODINT(msf2.TextMatrix(msf2.Row, 1)))
        End If
   End If
   '
   MSF1.Visible = False
   MSF1.Visible = True
   TXTCODIGO = ""
   Call LIMPIAR
   '
End Sub


Sub ACTUALIZA_MSF1(CI1%)
   If CI1% < 1 Then Exit Sub
   For J& = 1 To MSF1.Rows - 1
    CODIGO1$ = CODEXT$(CI1%)
      If TRIM$(MSF1.TextMatrix(J&, 1)) = TRIM$(CODIGO1$) Then
         MSF1.TextMatrix(J&, 3) = FRMEDIT.Text5.TEXT
         MSF1.TextMatrix(J&, 4) = FRMEDIT.TXTNROORIGINAL
         MSF1.TextMatrix(J&, 5) = FRMEDIT.txtlista.TEXT
         Exit For
      End If
    Next J&

End Sub
Sub BORRA_MSF1(CI1%)
   If CI1% < 1 Then Exit Sub
   For J& = 1 To MSF1.Rows - 1
    CODIGO1$ = CODEXT$(CI1%)
      If MSF1.TextMatrix(J&, 1) = CODIGO1$ Then
        MSF1.Row = J&
        Call borrafila(MSF1.Row, Me, MSF1)
        Exit For
      End If
    Next J&

End Sub

Private Sub MSF2_DblClick()
   
   FRMEDIT.TXTCODIGO.TEXT = msf2.TextMatrix(msf2.Row, 1)
   FRMEDIT.TXTDESCRIPCION.TEXT = msf2.TextMatrix(msf2.Row, 2)
   FRMEDIT.Text5.TEXT = msf2.TextMatrix(msf2.Row, 3)
   FRMEDIT.TXTNROORIGINAL = msf2.TextMatrix(msf2.Row, 4)
   FRMEDIT.txtlista.TEXT = msf2.TextMatrix(msf2.Row, 5)
   '
   txtlista.SetFocus
   '
   FRMEDIT.Show 1
   '
   If TRIM$(FRMEDIT.TXTCODIGO) = "" Then Exit Sub
   '
   MODIFICA% = 1

   If FRMEDIT.MODO1$ = "ACTUALIZAR" Then
        If MSF1.Visible = True Then
           MSF1.TextMatrix(MSF1.Row, 3) = FRMEDIT.Text5.TEXT
           MSF1.TextMatrix(MSF1.Row, 4) = FRMEDIT.TXTNROORIGINAL
           MSF1.TextMatrix(MSF1.Row, 5) = FRMEDIT.txtlista.TEXT
        Else
           msf2.TextMatrix(msf2.Row, 3) = FRMEDIT.Text5.TEXT
           msf2.TextMatrix(msf2.Row, 4) = FRMEDIT.TXTNROORIGINAL
           msf2.TextMatrix(msf2.Row, 5) = FRMEDIT.txtlista.TEXT
           Call ACTUALIZA_MSF1(CODINT(msf2.TextMatrix(XVALO(msf2.Row), 1)))
        End If
   ElseIf FRMEDIT.MODO1$ = "ELIMINAR" Then
        If MSF1.Visible = True Then
            Call borrafila(MSF1.Row, Me, MSF1)
        Else
            Call BORRA_MSF1(CODINT(msf2.TextMatrix(msf2.Row, 1)))
        End If
   End If
   '
   MSF1.Visible = False
   MSF1.Visible = True
   TXTCODIGO = ""
   Call LIMPIAR
   '
End Sub

Private Sub Cargar_art_seleccionado()
End Sub

Private Sub nuActulistaxg_Click()

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      TXTNROORIGINAL.SetFocus
      On Error GoTo 0
   End If
End Sub


Private Sub TXTCODIGO_Change()

    If TRIM$(TXTCODIGO) = "" Then
       TXTDESCRIPCION = ""
      GoTo 99
    Else
       TXTDESCRIPCION = ""
       Text5 = ""
       TXTNROORIGINAL = ""
       txtlista = ""
    End If
    '
    Ci% = CODINT(TXTCODIGO.TEXT)
    If Ci% > 0 Then
       TXTDESCRIPCION.TEXT = DESCRIT0$(Ci%)
    End If
    
    msf2.Redraw = False
    MSF1.Cols = 7
    msf2.Cols = 7
    COD$ = UCase$(TXTCODIGO)
    msf2.Rows = 1
    For J& = 1 To MSF1.Rows - 1
           CODIGO1$ = UCase$(MSF1.TextMatrix(J&, 1))
           If InStr(CODIGO1$, COD$) > 0 Then
             i& = i& + 1
             msf2.Rows = i& + 1
             msf2.TextMatrix(i&, 1) = MSF1.TextMatrix(J&, 1)
             msf2.TextMatrix(i&, 2) = MSF1.TextMatrix(J&, 2)
             msf2.TextMatrix(i&, 3) = MSF1.TextMatrix(J&, 3)
             msf2.TextMatrix(i&, 4) = MSF1.TextMatrix(J&, 4)
             msf2.TextMatrix(i&, 5) = MSF1.TextMatrix(J&, 5)
           End If
    Next J&
    '
    If msf2.Rows = 2 Then
       textdescripcion = DESCRIT0$(Ci%)
       Text5 = msf2.TextMatrix(1, 3)
       TXTNROORIGINAL = msf2.TextMatrix(1, 4)
       txtlista = msf2.TextMatrix(1, 5)
    End If
   
    'si el text box esta vacio lo hace invisible la grilla
99  If TXTCODIGO.TEXT <> "" Then
       msf2.Visible = True
       MSF1.Visible = False
    Else
       msf2.Visible = False
       MSF1.Visible = True
    End If
    msf2.Redraw = True
End Sub


Private Sub txtcodigo_DblClick()
    Call Command6_Click
End Sub

Private Sub TXTDESCRIPCION_Change()

    If TRIM$(TXTDESCRIPCION) = "" Then
      GoTo 99
    Else
       Text5 = ""
       TXTNROORIGINAL = ""
       txtlista = ""
       
    End If
    '
    Ci% = CODINT(TXTCODIGO.TEXT)
    If Ci% > 0 Then
       TXTCODIGO.TEXT = CODEXT$(Ci%)
    Else
       TXTCODIGO.TEXT = ""
    End If
    
    msf2.Redraw = False
    MSF1.Cols = 7
    msf2.Cols = 7
    DESCRI$ = UCase$(TXTDESCRIPCION)
    msf2.Rows = 1
    For J& = 1 To MSF1.Rows - 1
           DESCRI2$ = UCase$(MSF1.TextMatrix(J&, 2))
           If InStr(DESCRI2$, DESCRI$) > 0 Then
             i& = i& + 1
             msf2.Rows = i& + 1
             msf2.TextMatrix(i&, 1) = MSF1.TextMatrix(J&, 1)
             msf2.TextMatrix(i&, 2) = MSF1.TextMatrix(J&, 2)
             msf2.TextMatrix(i&, 3) = MSF1.TextMatrix(J&, 3)
             msf2.TextMatrix(i&, 4) = MSF1.TextMatrix(J&, 4)
             msf2.TextMatrix(i&, 5) = MSF1.TextMatrix(J&, 5)
             MSF1.TextMatrix(i&, 6) = MSF1.TextMatrix(J&, 6)
           End If
    Next J&
    '
    If msf2.Rows = 2 Then
       Text5 = msf2.TextMatrix(1, 3)
       TXTNROORIGINAL = msf2.TextMatrix(1, 4)
       txtlista = msf2.TextMatrix(1, 5)
    End If
    '
    'si el text box esta vacio lo hace invisible la grilla
99  If TXTDESCRIPCION.TEXT <> "" Then
       msf2.Visible = True
       MSF1.Visible = False
    Else
       msf2.Visible = False
       MSF1.Visible = True
    End If
    '
    msf2.Redraw = True


End Sub

Private Sub TXTLISTA_GotFocus()
'    txtlista = FORMATNUM$(XVALO(txtlista), "F14.4")
    
End Sub

Private Sub txtlista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then msf2.Visible = False: MSF1.Visible = True
        
    'VALIDO QUE HAYA UN CODIGO EN EL TEXT DE CODIGO VALIDO
    CI1% = CODINT%(TXTCODIGO)
    If CI1% < 1 Then Exit Sub
    '
    If KeyAscii = 27 Then
'    If Chr(KeyAscii) = "-" Then
        KeyAscii = 0
        Call LIMPIAR
        TXTDESCRIPCION.Enabled = True
        TXTDESCRIPCION.SetFocus
    End If
    
    If XVALO(txtlista.TEXT) > 0 Then
        If KeyAscii = 13 Then
'        If Chr(KeyAscii) = "+" Then
            KeyAscii = 0
            If VALIDA_COD_GRID(TXTCODIGO, MSF1, 1) > 0 Then GoTo 50
            cargar_datos
            TXTDESCRIPCION.SetFocus
50          Call LIMPIAR
            MODIFICA% = 1
        End If
    
    End If
End Sub
Private Sub cargar_datos()
    MSF1.Cols = 7
    MSF1.Rows = MSF1.Rows + 1
    MSF1.Row = MSF1.Rows - 1
    MSF1.TextMatrix(MSF1.Rows - 1, 1) = TXTCODIGO.TEXT
    MSF1.TextMatrix(MSF1.Rows - 1, 2) = TXTDESCRIPCION.TEXT
    MSF1.TextMatrix(MSF1.Rows - 1, 3) = Text5
    If XVALO(Label4) <= 1 Then MSF1.TextMatrix(MSF1.Rows - 1, 4) = Left(TXTNROORIGINAL.TEXT, 1)
    If XVALO(Label4) = 2 Then MSF1.TextMatrix(MSF1.Rows - 1, 4) = FORMATNUM$(TXTNROORIGINAL.TEXT, "f5.2")
    If XVALO(Label4) = 3 Then MSF1.TextMatrix(MSF1.Rows - 1, 4) = FECHATEX$(HOY(HO$))
    
    CANTIDECIM% = XVALO(Text3)
    MSF1.TextMatrix(MSF1.Rows - 1, 5) = FORMATNUM$(txtlista.TEXT, "f14." & CANTIDECIM%)
    
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.Row - VISIROWS Then
          MSF1.TopRow = MSF1.Row - VISIROWS
       End If
    End If
    
End Sub
Sub LIMPIAR()
    TXTCODIGO.TEXT = ""
    TXTDESCRIPCION.TEXT = ""
    TXTNROORIGINAL.TEXT = ""
    Text5 = ""
    txtlista = ""
End Sub

Private Sub txtlista_LostFocus()
   CANTIDECIM% = XVALO(Text3)
   txtlista = FORMATNUM$(XVALO(txtlista), "F14." & SAFETEXT$(CANTIDECIM%))
End Sub

Private Sub TXTNROORIGINAL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      txtlista.SetFocus
      On Error GoTo 0
   End If

End Sub
Sub Actualiza_DAT_X_Codigo(n%, CI1%)
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Screen.MousePointer = 11
   ARCHIPRE$ = TRIM$(Str$(n%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   CADECI% = XVALO(Text3): If CADECI% > 4 Then CADECI% = 4
   '
   Call COPYSTRU("LIPREVEN", ARCHIPRE$)
   LF2% = LENFIELD(ARCHIPRE$, 2)
   '
   Call APERBASDAT("PEDCLI")
   CONDI$ = "Cod_Lista = " & n% & " AND COD_PIEZA = " & CI1%
   SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   'Set ARTSELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim ARTSELEC As ADODB.Recordset
   Set ARTSELEC = New ADODB.Recordset
   Set ARTSELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))

   FIN& = CantRegistros("PREVENTA", CONDI$)
   U& = 0
   On Error Resume Next
   With ARTSELEC
     If FIN& > 0 Then
        Do While (.EOF = False)
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            TTXX$ = REGBLAN$(ARCHIPRE$)
            CII% = XVALO(!Cod_Pieza)
            Call REPLA(TTXX$, MKI$(CII%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CII%), 3, LF2%)
               MRX$ = SAFETEXT$(!Carac_Control)
            Call REPLA(TTXX$, MRX$, 3 + LF2%, 1)
            Call REPLA(TTXX$, !Num_Original, 4 + LF2%, 16)
               PRUNI# = XVALO(!Precio_Unitario)
            Call REPLA(TTXX$, MKD$(PRUNI#), 20 + LF2%, 8)
            Call REPLA(TTXX$, MKS$(!Cant_Limite), 28 + LF2%, 2)
            Call REPLA(TTXX$, MKD$(!Porc_Limite), 30 + LF2%, 1)
            '
            PRELI# = PRUNI#
            CDC$ = TRIM$(Str$(CADECI%))
            PREX$ = " " + FORMATNUM$(PRELI#, "F12." + CDC$) + "  "
            Call REPLA(TTXX$, PREX$, 1 + LF2% - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(TTXX$, MRX$, 2 + LF2%, 1)
            '
            If INCLUALIAS% > 0 Then
              ALX$ = TRIM$(NUPLANO$(CII%))
              If ALX$ <> "" Then
                Call REPLA(TTXX$, " (" + ALX$ + ")", LF2% - 16 - Len(ALX$), 3 + Len(ALX$))
              End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''''
            RFF$ = RECFILT$(ARCHIPRE$, 1, MKI$(CII%))
            If Len(RFF$) >= 4 Then
              REG& = CVS(Left$(RFF$, 4))
              If REG& > 0 Then
                    If REG& <= ULTREG&(ARCHIPRE$) Then
                      RELEI$ = REGLEIDO$(ARCHIPRE$, REG&)
                      REGBUS1& = REG&
                      Call GRAREG(ARCHIPRE$, TTXX$, REGBUS1&)
                    End If
               End If
            Else
                 Call REGAPP(ARCHIPRE$, TTXX$)
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''
            '
        .MoveNext
        Loop
     UREG& = ULTREG&(ARCHIPRE$)
     Call SETULTREG(ARCHIPRE$, UREG&)
     Call CIERRARCH(ARCHIPRE$)
     Call FRESHECHO("!ARCHIPRE$")
      End If
   End With
   Set ARTSELEC = Nothing
   Screen.MousePointer = 1
   '
End Sub




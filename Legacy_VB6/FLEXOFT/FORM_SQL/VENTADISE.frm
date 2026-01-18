VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VENTADISE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   18000
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   8760
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9960
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   17805
      _ExtentX        =   31406
      _ExtentY        =   17568
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ARMADO DE CONSULTAS"
      TabPicture(0)   =   "VENTADISE.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "LSTTABLAS"
      Tab(0).Control(2)=   "LSTMOSTARCMP"
      Tab(0).Control(3)=   "TXTCONSULTA"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "CMDCOPIAR"
      Tab(0).Control(6)=   "CMDPEGAR"
      Tab(0).Control(7)=   "GUARDACON"
      Tab(0).Control(8)=   "EJECUTAR"
      Tab(0).Control(9)=   "CMDBORRAR"
      Tab(0).Control(10)=   "List1"
      Tab(0).Control(11)=   "Command4"
      Tab(0).Control(12)=   "Command5"
      Tab(0).Control(13)=   "List2"
      Tab(0).Control(14)=   "MARCOBARRA"
      Tab(0).Control(15)=   "TXTCANREGI"
      Tab(0).Control(16)=   "Command8"
      Tab(0).Control(17)=   "CAMAUXI"
      Tab(0).Control(18)=   "Frame5"
      Tab(0).Control(19)=   "Combo4"
      Tab(0).Control(20)=   "LSTTABLASEC"
      Tab(0).Control(21)=   "Combo5"
      Tab(0).Control(22)=   "Command9"
      Tab(0).Control(23)=   "Frame2"
      Tab(0).Control(24)=   "Picture1"
      Tab(0).Control(25)=   "Command6"
      Tab(0).Control(26)=   "Command11"
      Tab(0).Control(27)=   "Command12"
      Tab(0).Control(28)=   "Command15"
      Tab(0).Control(29)=   "Command14"
      Tab(0).Control(30)=   "CMBJUNTAR"
      Tab(0).Control(31)=   "Frame6"
      Tab(0).Control(32)=   "Command28"
      Tab(0).Control(33)=   "Label3"
      Tab(0).Control(34)=   "Label4(1)"
      Tab(0).Control(35)=   "Label8"
      Tab(0).Control(36)=   "Label9"
      Tab(0).Control(37)=   "Label1"
      Tab(0).Control(38)=   "Label13"
      Tab(0).Control(39)=   "Label14"
      Tab(0).Control(40)=   "Label2(0)"
      Tab(0).Control(41)=   "Label10(2)"
      Tab(0).Control(42)=   "Label18"
      Tab(0).Control(43)=   "Label19"
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "VENTANAS DE CARGA O ZELECHO"
      TabPicture(1)   =   "VENTADISE.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label25(1)"
      Tab(1).Control(1)=   "Label25(0)"
      Tab(1).Control(2)=   "Label21"
      Tab(1).Control(3)=   "TEPRUEBA"
      Tab(1).Control(4)=   "Label22"
      Tab(1).Control(5)=   "Label20(0)"
      Tab(1).Control(6)=   "Label20(1)"
      Tab(1).Control(7)=   "Label20(2)"
      Tab(1).Control(8)=   "Label20(3)"
      Tab(1).Control(9)=   "Label20(4)"
      Tab(1).Control(10)=   "Label20(5)"
      Tab(1).Control(11)=   "Label23"
      Tab(1).Control(12)=   "MSFUPD"
      Tab(1).Control(13)=   "MSFESPEJO"
      Tab(1).Control(14)=   "MSF"
      Tab(1).Control(15)=   "TXTSEPARADOR"
      Tab(1).Control(16)=   "TXTFORMUEXCEL"
      Tab(1).Control(17)=   "TXTNAMETABLA"
      Tab(1).Control(18)=   "Command24"
      Tab(1).Control(19)=   "LSTMOSTARCMP2"
      Tab(1).Control(20)=   "LSTTABLAS2"
      Tab(1).Control(21)=   "CMDAPPCAMPOS"
      Tab(1).Control(22)=   "TXTCAMPOSUPDATE"
      Tab(1).Control(23)=   "Command13"
      Tab(1).Control(24)=   "Check12"
      Tab(1).Control(25)=   "Check5"
      Tab(1).Control(26)=   "TXTMOVIL"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Text8"
      Tab(1).Control(28)=   "Check6"
      Tab(1).Control(29)=   "Check7"
      Tab(1).Control(30)=   "Check8"
      Tab(1).Control(31)=   "Text9"
      Tab(1).Control(32)=   "Check9"
      Tab(1).Control(33)=   "TXTTITULO"
      Tab(1).Control(34)=   "Check10"
      Tab(1).Control(35)=   "TXTCONSU_ESPEJO"
      Tab(1).Control(36)=   "List3"
      Tab(1).Control(37)=   "Text10"
      Tab(1).Control(38)=   "Command16"
      Tab(1).Control(39)=   "Check11"
      Tab(1).Control(40)=   "Command17"
      Tab(1).Control(41)=   "TXTID"
      Tab(1).Control(42)=   "Command18"
      Tab(1).Control(43)=   "Command19"
      Tab(1).Control(44)=   "TXTCAMPOSENCABEZADO"
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "TABLAS AUXILIARES - DATASQL"
      TabPicture(2)   =   "VENTADISE.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20(9)"
      Tab(2).Control(1)=   "Label20(8)"
      Tab(2).Control(2)=   "Label20(7)"
      Tab(2).Control(3)=   "Label20(6)"
      Tab(2).Control(4)=   "Label24"
      Tab(2).Control(5)=   "MSF5"
      Tab(2).Control(6)=   "MSF4"
      Tab(2).Control(7)=   "Command27"
      Tab(2).Control(8)=   "Text15"
      Tab(2).Control(9)=   "Text14"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Command22"
      Tab(2).Control(11)=   "Check18"
      Tab(2).Control(12)=   "Text13"
      Tab(2).Control(13)=   "Text12"
      Tab(2).Control(14)=   "Check17"
      Tab(2).Control(15)=   "Check16"
      Tab(2).Control(16)=   "Check15"
      Tab(2).Control(17)=   "Check14"
      Tab(2).Control(18)=   "Check13"
      Tab(2).Control(19)=   "Text11"
      Tab(2).Control(20)=   "TXTMOVIL2"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Text20"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Command26"
      Tab(2).Control(23)=   "Command25"
      Tab(2).Control(24)=   "Command23"
      Tab(2).Control(25)=   "TXTIDLISTA"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "VENTADISE.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "VENTADISE.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "SSTab2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command50(3)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command50(1)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command50(0)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton Command50 
         Caption         =   "Cortar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   0
         Left            =   14040
         TabIndex        =   204
         Top             =   9600
         Width           =   1170
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   1
         Left            =   15195
         TabIndex        =   203
         Top             =   9600
         Width           =   1170
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Selecc Todo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   3
         Left            =   16335
         TabIndex        =   202
         Top             =   9600
         Width           =   1300
      End
      Begin VB.TextBox TXTCAMPOSENCABEZADO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   -72600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   183
         ToolTipText     =   "desde el menu Utilidades Enumerar Encabezado se Carga esta Caja"
         Top             =   900
         Visible         =   0   'False
         Width           =   5730
      End
      Begin VB.CommandButton Command19 
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   120
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Gen.Cod.Zelecho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68040
         TabIndex        =   119
         Top             =   9780
         Width           =   1935
      End
      Begin VB.TextBox TXTID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command17 
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
         Height          =   375
         Left            =   -69120
         TabIndex        =   117
         Top             =   9780
         Width           =   1000
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Copia y Escribe en otra grilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -73440
         TabIndex        =   116
         Top             =   9840
         Width           =   1500
      End
      Begin VB.CommandButton Command16 
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70080
         TabIndex        =   115
         Top             =   9780
         Width           =   1000
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   114
         Top             =   7560
         Width           =   2625
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":008C
         Left            =   -74760
         List            =   "VENTADISE.frx":008E
         TabIndex        =   113
         Top             =   8160
         Width           =   2625
      End
      Begin VB.TextBox TXTCONSU_ESPEJO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -71880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   112
         Top             =   6720
         Width           =   10890
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Levanta de Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74760
         TabIndex        =   111
         Top             =   9840
         Width           =   1200
      End
      Begin VB.TextBox TXTTITULO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   110
         Top             =   3345
         Width           =   13095
      End
      Begin VB.CheckBox Check9 
         Caption         =   "NoMostrarBotonTilde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   109
         Top             =   7080
         Width           =   2680
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -71880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   108
         Top             =   7800
         Width           =   10935
      End
      Begin VB.CheckBox Check8 
         Caption         =   "HabilitaInsertar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   6840
         Width           =   2175
      End
      Begin VB.CheckBox Check7 
         Caption         =   "HabilitaBorrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   106
         Top             =   6600
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         Caption         =   "UtilizaFormula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   105
         Top             =   6360
         Width           =   2175
      End
      Begin VB.TextBox Text8 
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
         Left            =   -74880
         TabIndex        =   104
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TXTMOVIL 
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
         Left            =   -75000
         TabIndex        =   103
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Agrega Registos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   102
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Condición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74430
         TabIndex        =   96
         Top             =   6990
         Width           =   7995
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            TabIndex        =   101
            Text            =   "Combo1"
            Top             =   315
            Width           =   2715
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
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
            Left            =   7275
            TabIndex        =   100
            ToolTipText     =   "Lista de Valores Posibles"
            Top             =   315
            Width           =   330
         End
         Begin VB.ComboBox CMBOPERADOR 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "VENTADISE.frx":0090
            Left            =   3240
            List            =   "VENTADISE.frx":0092
            TabIndex        =   99
            TabStop         =   0   'False
            Text            =   "CMBOPERADOR"
            Top             =   315
            Width           =   1080
         End
         Begin VB.ComboBox CMBCAMPOS 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   98
            Text            =   "CMBCAMPOS"
            Top             =   315
            Width           =   2790
         End
         Begin VB.CommandButton Command3 
            Caption         =   "?"
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
            Left            =   7665
            TabIndex        =   97
            ToolTipText     =   "Caracter Comodin"
            Top             =   315
            Width           =   260
         End
      End
      Begin VB.ListBox LSTTABLAS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         IntegralHeight  =   0   'False
         Left            =   -74400
         Sorted          =   -1  'True
         TabIndex        =   95
         Top             =   2265
         Width           =   1665
      End
      Begin VB.ListBox LSTMOSTARCMP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":0094
         Left            =   -70770
         List            =   "VENTADISE.frx":0096
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   94
         Top             =   2265
         Width           =   2505
      End
      Begin VB.TextBox TXTCONSULTA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -69945
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   93
         Top             =   7980
         Width           =   7050
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   -66390
         Picture         =   "VENTADISE.frx":0098
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Agregar Condición"
         Top             =   7080
         Width           =   540
      End
      Begin VB.CommandButton CMDCOPIAR 
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   -63750
         TabIndex        =   91
         Top             =   7095
         Width           =   850
      End
      Begin VB.CommandButton CMDPEGAR 
         Caption         =   "Pegar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   -63750
         TabIndex        =   90
         Top             =   7350
         Width           =   850
      End
      Begin VB.CommandButton GUARDACON 
         Caption         =   "Guardar Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   -65745
         TabIndex        =   89
         Top             =   7095
         Width           =   855
      End
      Begin VB.CommandButton EJECUTAR 
         Caption         =   "Ejecutar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   -64830
         TabIndex        =   88
         Top             =   7485
         Width           =   960
      End
      Begin VB.CommandButton CMDBORRAR 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   -63750
         TabIndex        =   87
         Top             =   7605
         Width           =   850
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2120
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":03A2
         Left            =   -68115
         List            =   "VENTADISE.frx":03A4
         TabIndex        =   86
         Top             =   2265
         Width           =   2145
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -65820
         Picture         =   "VENTADISE.frx":03A6
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2475
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -65820
         Picture         =   "VENTADISE.frx":06B0
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3630
         Width           =   330
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":09BA
         Left            =   -74445
         List            =   "VENTADISE.frx":09BC
         TabIndex        =   83
         Top             =   7980
         Width           =   4280
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   -64755
         ScaleHeight     =   75
         ScaleWidth      =   900
         TabIndex        =   81
         Top             =   7200
         Width           =   960
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   82
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.TextBox TXTCANREGI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   -69945
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Top             =   8985
         Width           =   7050
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
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
         Left            =   -74565
         TabIndex        =   79
         Top             =   1440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.ListBox CAMAUXI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":09BE
         Left            =   -70770
         List            =   "VENTADISE.frx":09C5
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   78
         Top             =   5280
         Width           =   2505
      End
      Begin VB.Frame Frame5 
         Caption         =   "Configuracion Campo Auxiliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68115
         TabIndex        =   70
         Top             =   4440
         Width           =   5220
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   1080
            Width           =   4815
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            Sorted          =   -1  'True
            TabIndex        =   73
            Text            =   "CMBCAMPOS"
            Top             =   520
            Width           =   1815
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            Sorted          =   -1  'True
            TabIndex        =   72
            Text            =   "CMBCAMPOS"
            Top             =   520
            Width           =   2415
         End
         Begin VB.CommandButton Command7 
            Caption         =   "OK"
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
            Left            =   4515
            TabIndex        =   71
            Top             =   210
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Condicion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   77
            Top             =   885
            Width           =   1275
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   76
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Tabla:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   75
            Top             =   315
            Width           =   615
         End
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74430
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   5715
         Width           =   1575
      End
      Begin VB.ListBox LSTTABLASEC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":09D2
         Left            =   -72675
         List            =   "VENTADISE.frx":09D4
         Sorted          =   -1  'True
         TabIndex        =   68
         Top             =   2265
         Width           =   1665
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72540
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   5715
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "X"
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
         Left            =   -71265
         TabIndex        =   66
         ToolTipText     =   "Caracter Comodin"
         Top             =   2040
         Width           =   260
      End
      Begin VB.Frame Frame2 
         Caption         =   "Presentacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   -65355
         TabIndex        =   42
         Top             =   1320
         Width           =   2460
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1500
            TabIndex        =   61
            Top             =   220
            Width           =   855
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   210
            TabIndex        =   60
            Top             =   550
            Width           =   2145
         End
         Begin VB.Frame Frame3 
            Caption         =   "Sort"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   105
            TabIndex        =   56
            Top             =   1425
            Width           =   975
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   225
               TabIndex        =   59
               Top             =   220
               UseMaskColor    =   -1  'True
               Value           =   -1  'True
               Width           =   645
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               Caption         =   "Asc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   225
               TabIndex        =   58
               Top             =   420
               Width           =   645
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               Caption         =   "Des"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   225
               TabIndex        =   57
               Top             =   620
               Width           =   645
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Pos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1160
            TabIndex        =   51
            Top             =   1425
            Width           =   1185
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   465
               TabIndex        =   53
               Top             =   260
               Width           =   600
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   465
               TabIndex        =   52
               Top             =   560
               Width           =   600
            End
            Begin VB.Label Label10 
               Caption         =   "H:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   55
               Top             =   285
               Width           =   795
            End
            Begin VB.Label Label11 
               Caption         =   "V:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   225
               TabIndex        =   54
               Top             =   585
               Width           =   795
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Totaliza"
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
            Left            =   210
            TabIndex        =   50
            Top             =   2385
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Agrupa"
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
            Left            =   210
            TabIndex        =   49
            Top             =   2605
            Width           =   1065
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1710
            TabIndex        =   48
            Top             =   1095
            Width           =   630
         End
         Begin VB.CommandButton Command10 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   47
            ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
            Top             =   1095
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Distintos"
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
            Left            =   210
            TabIndex        =   46
            Top             =   2825
            Width           =   1065
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Ocultar"
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
            Left            =   210
            TabIndex        =   45
            Top             =   3045
            Width           =   1065
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   210
            TabIndex        =   44
            Text            =   "Combo6"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   210
            TabIndex        =   43
            Top             =   1095
            Width           =   1170
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Form:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   660
            TabIndex        =   65
            Top             =   265
            Width           =   795
         End
         Begin VB.Label Label7 
            Caption         =   "Label:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   64
            Top             =   320
            Width           =   645
         End
         Begin VB.Label Label5 
            Caption         =   "Eco:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   63
            Top             =   870
            Width           =   645
         End
         Begin VB.Label Label12 
            Caption         =   "Form:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1605
            TabIndex        =   62
            Top             =   870
            Width           =   645
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   -65340
         ScaleHeight     =   120
         ScaleWidth      =   2385
         TabIndex        =   41
         Top             =   4680
         Width           =   2380
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -63855
         Picture         =   "VENTADISE.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3690
         Width           =   780
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   -65925
         Picture         =   "VENTADISE.frx":0CE0
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Eliminar (Borrar) Consulta Activa"
         Top             =   1425
         Width           =   500
      End
      Begin VB.CommandButton Command12 
         Caption         =   "OR"
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
         Left            =   -74670
         TabIndex        =   38
         Top             =   8760
         Width           =   210
      End
      Begin VB.CommandButton Command15 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -74670
         TabIndex        =   37
         Top             =   8040
         Width           =   210
      End
      Begin VB.CommandButton Command14 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -70140
         TabIndex        =   36
         Top             =   8040
         Width           =   210
      End
      Begin VB.ComboBox CMBJUNTAR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73470
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   5340
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Operadores Sql"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74430
         TabIndex        =   33
         Top             =   6120
         Width           =   6135
         Begin VB.CheckBox chkUnionAll 
            Caption         =   "Union All"
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
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Generar Codigo de Menus en  grilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71880
         TabIndex        =   32
         Top             =   9720
         Width           =   1665
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9375
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   17520
         Begin VB.CommandButton Command40 
            Caption         =   "Generar Ent.Request Tabla Activa (public int Nume_Cli { get; set; })"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   205
            ToolTipText     =   "Graba en BD la grilla"
            Top             =   8880
            Width           =   6135
         End
         Begin VB.TextBox TXTBUSCADOR 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5400
            TabIndex        =   188
            Top             =   7800
            Width           =   4575
         End
         Begin VB.CheckBox CheckbtnNuevo 
            Caption         =   "Botón Nuevo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   8760
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox CheckbtnGrabar 
            Caption         =   "Botón Grabar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   186
            Top             =   8520
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Generar Rutina de Carga de valores de Entidad (Segun ID)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9000
            TabIndex        =   185
            ToolTipText     =   "Graba en BD la grilla"
            Top             =   8280
            Width           =   3015
         End
         Begin VB.CommandButton Command38 
            Caption         =   "Generar Entidad Tabla Activa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   184
            ToolTipText     =   "Graba en BD la grilla"
            Top             =   8280
            Width           =   3015
         End
         Begin VB.TextBox Text25 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   182
            Text            =   "VENTADISE.frx":0FEA
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CommandButton Command37 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   17040
            TabIndex        =   181
            Top             =   4440
            Width           =   375
         End
         Begin VB.CommandButton Command36 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   17040
            TabIndex        =   180
            Top             =   4800
            Width           =   375
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Generar Código.net"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   179
            ToolTipText     =   "Graba en BD la grilla"
            Top             =   8280
            Width           =   1935
         End
         Begin VB.CommandButton Command34 
            Caption         =   "NEW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10080
            TabIndex        =   177
            Top             =   3720
            Width           =   615
         End
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7920
            TabIndex        =   176
            Top             =   4080
            Width           =   5895
         End
         Begin VB.TextBox TXTCODIGO 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7920
            TabIndex        =   175
            Top             =   3720
            Width           =   1695
         End
         Begin VB.CommandButton Command33 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9600
            TabIndex        =   174
            Top             =   3720
            Width           =   375
         End
         Begin VB.TextBox TXTNAMEFRM 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10800
            TabIndex        =   173
            Top             =   3720
            Width           =   1695
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Grabar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   172
            ToolTipText     =   "Graba en BD la grilla"
            Top             =   8280
            Width           =   1335
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   171
            Top             =   7800
            Width           =   5175
         End
         Begin VB.CommandButton cmd33 
            Caption         =   "Sel.Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   170
            ToolTipText     =   "Seleccionar Formulario a Aplicar"
            Top             =   8280
            Width           =   1335
         End
         Begin VB.CommandButton Command32 
            Caption         =   "SEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   169
            ToolTipText     =   "Selección automática"
            Top             =   4120
            Width           =   735
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Cargar Grilla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   167
            Top             =   4120
            Width           =   1575
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   1320
            Width           =   2415
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            IntegralHeight  =   0   'False
            ItemData        =   "VENTADISE.frx":1023
            Left            =   2520
            List            =   "VENTADISE.frx":102A
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   165
            Top             =   1320
            Width           =   5385
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            IntegralHeight  =   0   'False
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   164
            Top             =   1680
            Width           =   2385
         End
         Begin VB.Frame Frame11 
            Caption         =   "DatoTexto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   3000
            TabIndex        =   160
            Top             =   240
            Width           =   14415
            Begin VB.CommandButton Command31 
               Caption         =   "++"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   163
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton Command30 
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   162
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Text16 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               TabIndex        =   161
               Top             =   520
               Width           =   4575
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Dato Numerico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2775
            Begin VB.TextBox Text17 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   100
               TabIndex        =   156
               Top             =   520
               Width           =   1215
            End
            Begin VB.CommandButton Command29 
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   155
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton Command21 
               Caption         =   "++"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   154
               Top             =   240
               Width           =   615
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSF3 
            Height          =   3345
            Left            =   120
            TabIndex        =   168
            Top             =   4440
            Width           =   16875
            _ExtentX        =   29766
            _ExtentY        =   5900
            _Version        =   393216
            Cols            =   3
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
         Begin VB.Label LBLNUEVO 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nuevo Formulario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   5400
            TabIndex        =   178
            Top             =   4100
            Width           =   2415
         End
      End
      Begin VB.TextBox TXTIDLISTA 
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
         Height          =   330
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   1305
      End
      Begin VB.CommandButton Command23 
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69650
         TabIndex        =   28
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Guardar y Genera Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -63000
         TabIndex        =   27
         Top             =   8380
         Width           =   3015
      End
      Begin VB.CommandButton Command26 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70200
         TabIndex        =   26
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox Text20 
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
         Height          =   420
         Left            =   -69240
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "   "
         Top             =   900
         Width           =   2655
      End
      Begin VB.TextBox TXTMOVIL2 
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
         Left            =   -73320
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   660
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text11 
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
         Left            =   -74640
         TabIndex        =   23
         Top             =   1140
         Width           =   615
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Agrega Registos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   22
         Top             =   5340
         Width           =   1695
      End
      Begin VB.CheckBox Check14 
         Caption         =   "UtilizaFormula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   5580
         Width           =   2175
      End
      Begin VB.CheckBox Check15 
         Caption         =   "HabilitaBorrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   20
         Top             =   5820
         Width           =   2175
      End
      Begin VB.CheckBox Check16 
         Caption         =   "HabilitaInsertar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   19
         Top             =   6060
         Width           =   2175
      End
      Begin VB.CheckBox Check17 
         Caption         =   "NoMostrarBotonTilde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   18
         Top             =   6300
         Width           =   2680
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72720
         TabIndex        =   17
         Top             =   5580
         Width           =   2625
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -69945
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   5940
         Width           =   10200
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Levanta de Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74640
         TabIndex        =   15
         Top             =   6540
         Width           =   1200
      End
      Begin VB.CommandButton Command22 
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -62880
         TabIndex        =   14
         Top             =   9180
         Width           =   2895
      End
      Begin VB.CommandButton Command13 
         Caption         =   "GENERAR CODIGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63720
         TabIndex        =   13
         Top             =   9780
         Width           =   2415
      End
      Begin VB.TextBox Text14 
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
         Height          =   420
         Left            =   -66360
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "   "
         Top             =   900
         Width           =   6375
      End
      Begin VB.TextBox TXTCAMPOSUPDATE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72960
         TabIndex        =   11
         ToolTipText     =   "ACA VAN LOS CAMPOS SEPERADOS POR PUNTO Y COMA PARA ARMAR EL CODIGO UPDATE QUE SE USA PARA GRABAR UNA VENTANA "
         Top             =   4740
         Width           =   12015
      End
      Begin VB.CommandButton CMDAPPCAMPOS 
         Caption         =   "ADD CAMPOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -60720
         TabIndex        =   10
         Top             =   5340
         Width           =   1245
      End
      Begin VB.ListBox LSTTABLAS2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         IntegralHeight  =   0   'False
         Left            =   -60840
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1980
         Width           =   2145
      End
      Begin VB.ListBox LSTMOSTARCMP2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         IntegralHeight  =   0   'False
         ItemData        =   "VENTADISE.frx":103D
         Left            =   -67680
         List            =   "VENTADISE.frx":1044
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   5100
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton Command24 
         Caption         =   "GEN. COD. UPDATE"
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
         Left            =   -63360
         TabIndex        =   7
         Top             =   6180
         Width           =   2415
      End
      Begin VB.TextBox TXTNAMETABLA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -60840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1620
         Width           =   2055
      End
      Begin VB.TextBox TXTFORMUEXCEL 
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
         Left            =   -70200
         TabIndex        =   5
         Top             =   6180
         Width           =   6735
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -69960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   7050
         Width           =   10200
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Eliminar estructura Activa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -59640
         TabIndex        =   3
         ToolTipText     =   "Eliminar Estructura Activa"
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox TXTSEPARADOR 
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
         Left            =   -74040
         MaxLength       =   1
         TabIndex        =   2
         ToolTipText     =   "Defina que tipo de separador , o ; para los campos en el armado del codigo update"
         Top             =   4740
         Width           =   255
      End
      Begin VB.CommandButton Command28 
         Caption         =   "seleccionar todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70080
         TabIndex        =   1
         Top             =   2000
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2205
         Left            =   -74040
         TabIndex        =   121
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   960
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   3889
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         BackColor       =   16501405
         BackColorFixed  =   16501405
         BackColorSel    =   16501405
         ForeColorSel    =   0
         BackColorBkg    =   16501405
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
      Begin MSFlexGridLib.MSFlexGrid MSFESPEJO 
         Height          =   1005
         Left            =   -74040
         TabIndex        =   122
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   3720
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   1773
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         BackColor       =   16501405
         BackColorFixed  =   16501405
         BackColorSel    =   16501405
         BackColorBkg    =   16501405
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   2805
         Left            =   -74640
         TabIndex        =   123
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   1500
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   4948
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         BackColor       =   16777215
         BackColorFixed  =   16501405
         ForeColorFixed  =   0
         ForeColorSel    =   16777215
         GridLinesFixed  =   1
         SelectionMode   =   1
         MergeCells      =   1
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
         Height          =   765
         Left            =   -74640
         TabIndex        =   124
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   4500
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   1349
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         ForeColorFixed  =   0
         ForeColorSel    =   16777215
         GridLinesFixed  =   1
         SelectionMode   =   1
         MergeCells      =   1
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
      Begin MSFlexGridLib.MSFlexGrid MSFUPD 
         Height          =   1005
         Left            =   -74040
         TabIndex        =   125
         ToolTipText     =   $"VENTADISE.frx":1057
         Top             =   5100
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   1773
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         BackColor       =   16501405
         BackColorFixed  =   16501405
         BackColorSel    =   16501405
         BackColorBkg    =   16501405
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9135
         Left            =   120
         TabIndex        =   189
         Top             =   360
         Width           =   17685
         _ExtentX        =   31194
         _ExtentY        =   16113
         _Version        =   393216
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "FRM.Designer.CS"
         TabPicture(0)   =   "VENTADISE.frx":10E3
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lbl(0)"
         Tab(0).Control(1)=   "RichTextBox1"
         Tab(0).Control(2)=   "Text23"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Entidad"
         TabPicture(1)   =   "VENTADISE.frx":10FF
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lbl(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Text19"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Codigo Actual en frm"
         TabPicture(2)   =   "VENTADISE.frx":111B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Text22(0)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "FRM.cs"
         TabPicture(3)   =   "VENTADISE.frx":1137
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbl(3)"
         Tab(3).Control(1)=   "Text21"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Metodos, Varios"
         TabPicture(4)   =   "VENTADISE.frx":1153
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Text22(1)"
         Tab(4).Control(1)=   "lbl(4)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Anexo Diseño"
         TabPicture(5)   =   "VENTADISE.frx":116F
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lbl(5)"
         Tab(5).Control(1)=   "Text26"
         Tab(5).ControlCount=   2
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9105
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   196
            Top             =   360
            Width           =   17355
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Index           =   0
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   195
            Top             =   360
            Width           =   9240
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8625
            Index           =   1
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   194
            Top             =   360
            Width           =   17355
         End
         Begin VB.TextBox Text26 
            Height          =   8895
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   192
            Text            =   "VENTADISE.frx":118B
            Top             =   360
            Width           =   17355
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8865
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   191
            Top             =   360
            Width           =   17355
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9345
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   190
            Top             =   480
            Width           =   17400
         End
         Begin VB.PictureBox RichTextBox1 
            Height          =   255
            Left            =   -74520
            ScaleHeight     =   255
            ScaleWidth      =   30
            TabIndex        =   193
            Top             =   840
            Width           =   30
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FRM.Designer.CS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   -74400
            TabIndex        =   201
            ToolTipText     =   "Codigo de Diseño (Designer) Doble Click Ayuda"
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   200
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FRM.cs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   3
            Left            =   -65640
            TabIndex        =   199
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CargaEntidadClientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   4
            Left            =   -74880
            TabIndex        =   198
            ToolTipText     =   "Cargar en Capa Datos Clase: CD_NegocioClientes"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anexo Diseño"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   -59880
            TabIndex        =   197
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.Label Label23 
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   153
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Título de Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   152
         Top             =   7320
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "Lista de Campos Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   151
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "Código Generado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   -71880
         TabIndex        =   150
         Top             =   7620
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Consulta para Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -71895
         TabIndex        =   149
         Top             =   6495
         Width           =   2775
      End
      Begin VB.Label Label20 
         Caption         =   "Código Generado:"
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
         Left            =   -72000
         TabIndex        =   148
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Título de la Ventana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74040
         TabIndex        =   147
         Top             =   3180
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Ingresar encabezados con Formato( Codigo/A16)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70080
         TabIndex        =   146
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
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
         Left            =   -75000
         TabIndex        =   145
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Separador de campos para consulta update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   -74880
         TabIndex        =   144
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Campos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -71280
         TabIndex        =   143
         Top             =   2025
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -69060
         TabIndex        =   142
         Top             =   2025
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74880
         TabIndex        =   141
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   555
         Left            =   -71205
         TabIndex        =   140
         Top             =   1425
         Width           =   5235
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   -73710
         TabIndex        =   139
         Top             =   1680
         Width           =   2325
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72990
         TabIndex        =   138
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   1425
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Auxiliares:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70950
         TabIndex        =   137
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Tabla Principal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   -74550
         TabIndex        =   136
         Top             =   2025
         Width           =   1560
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   -72885
         TabIndex        =   135
         Top             =   5745
         Width           =   330
      End
      Begin VB.Label Label18 
         Caption         =   "Tabla Secundaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -72900
         TabIndex        =   134
         Top             =   2025
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Tipo.Join:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74430
         TabIndex        =   133
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   5400
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Cant.Campos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   132
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Título de Recordset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   -72720
         TabIndex        =   131
         Top             =   5340
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "Consulta para Recordset  (Ejemplo: Select campo1, campo2 from datasql with(nolock))"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   -69960
         TabIndex        =   130
         Top             =   5340
         Width           =   10200
      End
      Begin VB.Label Label25 
         Caption         =   "Campos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -73680
         TabIndex        =   129
         Top             =   4740
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Formula Excel ""=EXCEL;(L1/K1)*(L1/K1)*M1"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   -72480
         TabIndex        =   128
         Top             =   6080
         Width           =   2415
      End
      Begin VB.Label Label20 
         Caption         =   "Código Generado para que se cree la estrucutura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -69960
         TabIndex        =   127
         Top             =   6850
         Width           =   10200
      End
      Begin VB.Label Label20 
         Caption         =   "TABLAS AUXILIARES SOLO CONTRA TABLA DATASQL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   -69960
         TabIndex        =   126
         Top             =   5580
         Width           =   10200
      End
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   159
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   158
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   157
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReverUpd 
         Caption         =   "Revertir Update"
      End
      Begin VB.Menu mnuArmaSql 
         Caption         =   "Arma Sql"
      End
      Begin VB.Menu mnupyes 
         Caption         =   "ProcesoYes"
      End
   End
   Begin VB.Menu mnuFormatos 
      Caption         =   "Formatos"
      Begin VB.Menu mnuFormatohora 
         Caption         =   "'H' formato hora 10:00:00"
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "Uitlidades"
      Begin VB.Menu mnuArmaEnumeEncabezado 
         Caption         =   "Arma Enumeracion de Campos para Encabezado"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update Campos"
      End
      Begin VB.Menu mnuCreaEstructura 
         Caption         =   "Crear Archivo Estructura"
      End
      Begin VB.Menu mnuAyudaVentabnew 
         Caption         =   "ATAJOS EN VENTABNEW (VENTANA DE CARGA)"
      End
      Begin VB.Menu MnuAutonumTabAux 
         Caption         =   "CAMPO AUTONUMERICO EN TABLA AUXILIAR"
      End
      Begin VB.Menu mnuCeldaResultadoFormula 
         Caption         =   "Celda Resultado de Función"
      End
   End
End
Attribute VB_Name = "VENTADISE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KLIK%
Dim MODIFENCA%, ILIS1&
Dim ESTADO$
    ' variable para la clase
    Dim tTip As CLSTOOLTIP
Dim INDICEX%
Dim VectorControlAdd(512)
Dim VectorSystem_Windows_Forms(1024)
Dim VectorLLamadaZelecho(1024)
Dim VectorBlanformu(1024)
Dim VectorSoloNumeros(1024)

Dim VectorCargaValoresEntidad() ' carga en un vector estas filas segun corresponda: entidad.Nume_Cli = Convert.ToInt16(lector["Nume_Cli"]);
Dim VectorMostrarDatosenCajasTexto(1924)
Dim VectorAsignarCajasTextoaEntidad()
Public FILABUSCAR&
Sub ArmaTABLAYCONDI(TABLAYCONDI$, RET_TABLA$, RET_CONDI$)
'   Campos$ = REPLACAR(Campos$, ";", ",")
   strArray = Split(TABLAYCONDI$, ";")
   On Error Resume Next
     RET_TABLA$ = strArray(0)
     RET_CONDI$ = strArray(1)
   On Error GoTo 0

End Sub

Sub BLANFORMU()
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
        Check8.Value = 0
        Check9.Value = 0
        Check10.Value = 0
        Check11.Value = 0
        Text10 = ""
        Text11 = ""
        Text9 = ""
        Text8 = 4
        TXTCONSU_ESPEJO = ""
        TitVentana = ""
        Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(VENTADISE.MSFESPEJO)
        Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(VENTADISE.MSF)
        List3.Clear

End Sub

Sub BLANFORMUFRM()
        List5.Clear
        Text21.TEXT = ""
        Text19 = ""
        Text22(0) = ""
        Text22(1) = ""
        TXTNAMEFRM = ""
        Text24 = ""
        Text18(1) = ""
        MSF3.Rows = 1
        Text26 = ""
End Sub


Function CAJATEXTOIDREPLICA$()
    'DEVUELVE EL NOMBRE DE LA CAJA DE TEXTO GENERADO DEL CAMPO IDRELICA
    
    'SOLO DEBERIA HABER UN SOLO SI DE CAMPO IDREPLICA

   TABLA$ = Text18(0)
   For i& = 1 To MSF3.Rows - 1
      If UCase$(TRIM$(MSF3.TextMatrix(i&, 6))) = "SI" Then
         CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
         'TXTDEUDOR12_Nume_Cli
         CAJATEXTOIDREPLICA$ = "TXT" & Text18(0) & "_" & CampoName$
         Exit For
      End If
   Next i&
End Function

Sub CARGADATOSUNICOS(IDLISTA&)
  Dim obj As New RECORXET
  Set RS1 = obj.RS_TAB_COMP_ENCAPRFUN(IDLISTA&)
  Text11 = XVALO(RS1!CANTICAMPOS)
  Text12 = SAFETEXT$(RS1!TITURECORD)
  Text13 = SAFETEXT$(RS1!CONSULTASQL)
  Text14 = SAFETEXT$(RS1!Descripcion)
  Check13.Value = XVALO(RS1!ADDREGISTROS)
  Check14.Value = XVALO(RS1!USAFORMULA)
  Check15.Value = XVALO(RS1!BORRAR)
  Check16.Value = XVALO(RS1!INSERTAR)
  Check17.Value = XVALO(RS1!NOTILDE)
  Check18.Value = XVALO(RS1!USARECORD)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing

End Sub

Sub CARGAMSF4()
  MSF4.Rows = 8
  MSF4.ColWidth(0) = 4000
  MSF4.TextMatrix(0, 0) = "Encabezado + Formato (Canti/F8.1)"
  MSF4.TextMatrix(1, 0) = "Formula (4;*;6)"
  MSF4.TextMatrix(2, 0) = "CampEditables (S/N)"
  MSF4.TextMatrix(3, 0) = "ColEcoCampo (CODIGO;Descripcion;Master)"
  MSF4.TextMatrix(4, 0) = "TxtUbicadoGrillaTotal(Texto)"
  MSF4.TextMatrix(5, 0) = "Total (S/N)"
  MSF4.TextMatrix(6, 0) = "Dejar Libre (S/N)"
  MSF4.TextMatrix(7, 0) = "NUMERICO=0 - TEXTO=1"

End Sub

Sub CargarVectorBlanformu()
    Text21.TEXT = Text21.TEXT & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "  private void BLANFORMU ()" & vbCrLf
    Text21.TEXT = Text21.TEXT & "             " & "{" & vbCrLf
    For i& = 0 To UBound(VectorBlanformu)
       If TRIM$(VectorBlanformu(i&)) <> "" Then
          Text21.TEXT = Text21.TEXT & "             " & VectorBlanformu(i&) & vbCrLf
       End If
    Next i&
    Text21.TEXT = Text21.TEXT & "             " & "}" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
End Sub

Sub CargarVectorVectorSoloNumeros()
    Text21.TEXT = Text21.TEXT & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //============== CAPA VISTA FRM.CS SOLO NUMEROS PARA T.BOX NUMERICOS==========================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    For i& = 0 To UBound(VectorSoloNumeros)
       If TRIM$(VectorSoloNumeros(i&)) <> "" Then
          Text21.TEXT = Text21.TEXT & "             " & VectorSoloNumeros(i&)
       End If
    Next i&
    Text21.TEXT = Text21.TEXT & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //============== ** FIN ** CAPA VISTA FRM.CS SOLO NUMEROS PARA T.BOX NUMERICOS==========================================" & vbCrLf
    Text21.TEXT = Text21.TEXT & "       //===============================================================================================================" & vbCrLf
End Sub

Sub ArmaCodigodeAsingacionaCajasdeTexto()
    TABLA$ = Text18(0)
    T1$ = "   //============================================================================================" & vbCrLf
    T1$ = T1$ & "       // *****   ESTO VA EN CAPA VISTA DONDE SE VAN A ASIGNAR LOS VALORES A LAS CAJAS DE TEXTO ********  " & vbCrLf
    T1$ = T1$ & "   //============================================================================================" & vbCrLf
    T1$ = T1$ & "  private void MostrarDatos_" & TABLA$ & " ()" & vbCrLf
    T1$ = T1$ & "             " & "{" & vbCrLf
    For i& = 0 To UBound(VectorMostrarDatosenCajasTexto)
       If TRIM$(VectorMostrarDatosenCajasTexto(i&)) <> "" Then
          T1$ = T1$ & "             " & VectorMostrarDatosenCajasTexto(i&) & vbCrLf
       End If
    Next i&
    T1$ = T1$ & "             " & "}" & vbCrLf
    T1$ = T1$ & "       //===============================================================================================================" & vbCrLf
    T1$ = T1$ & "       //   ***** FIN DE ESTO VA EN CAPA VISTA ONDE SE VAN A ASIGNAR LOS VALORES A LAS CAJAS DE TEXTO ******** " & vbCrLf
    T1$ = T1$ & "       //===============================================================================================================" & vbCrLf
    
    Text21.TEXT = Text21.TEXT + T1$  'frm.cs

End Sub

Sub CargarVectordeCargaValoresEntidad()
    TAPRIN$ = Text18(0)
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    Text22(1) = ""
    Set Rsrecordset = FLEXCONN.Execute("sp_columns [" & TAPRIN$ & "]")
    TX1$ = "public class " & TAPRIN$ & vbCrLf

    Do Until Rsrecordset.EOF
        TF$ = ""
        cnt = cnt + 1
        NAMEX$ = Rsrecordset!column_name
        TYPEX$ = Rsrecordset!type_name
        LENGX$ = Rsrecordset!Length
        TCLARO$ = TIPOCLARO$(TYPEX$)
        TDT$ = TRIM$(UCase$(TYPEX$))
        ReDim Preserve VectorCargaValoresEntidad(cnt)
        If Left$(TDT$, 3) = "INT" Or Left$(TDT$, 8) = "SMALLINT" Or Left$(TDT$, 5) = "FLOAT" Then
           Select Case TDT$
             Case "INT"
               If XVALO(LENGX$) < 4 Then
                  TYPEX$ = "Int16"
               Else
                  TYPEX$ = "Int32"
               End If
             Case "SMALLINT": TYPEX$ = "Int16"
                   
             Case "FLOAT": TYPEX$ = "Decimal"
              
            End Select
           If Left$(UCase$(TYPEX$), 3) = "INT" Then
                T1$ = "if(lector[" & Chr(34) & NAMEX$ & Chr(34) & "] == DBNull.Value)" & vbCrLf
                T1$ = T1$ + "                {valor = 0; }" & vbCrLf
                T1$ = T1$ + "               else" & vbCrLf
                T1$ = T1$ + "                {valor = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX$ & Chr(34) & "]);}" & vbCrLf
                T1$ = T1$ + "                entidad." & NAMEX$ & " = Convert.To" & TYPEX$ & "(valor);" & vbCrLf & vbCrLf
                T1$ = T1$ + "  //-------------------------------------------------------------------------------------------" & vbCrLf

           Else  'SERIA CASO DECIMAL
                T1$ = "if(lector[" & Chr(34) & NAMEX$ & Chr(34) & "] == DBNull.Value)" & vbCrLf
                T1$ = T1$ + "                {valorDecimal = 0; }" & vbCrLf
                T1$ = T1$ + "               else" & vbCrLf
                T1$ = T1$ + "                {valorDecimal = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX$ & Chr(34) & "]);}" & vbCrLf
                T1$ = T1$ + "                entidad." & NAMEX$ & " = Convert.To" & TYPEX$ & "(valorDecimal);" & vbCrLf & vbCrLf
                T1$ = T1$ + "  //-------------------------------------------------------------------------------------------" & vbCrLf

           End If
           VectorCargaValoresEntidad(cnt) = T1$

        ElseIf Left$(TDT$, 4) = "DATE" Or Left$(TDT$, 13) = "SMALLDATETIME" Then
           'texto de fecha para concatenar mas abajo
           TYPEX$ = "DateTime"
                T1$ = "if(lector[" & Chr(34) & NAMEX$ & Chr(34) & "] == DBNull.Value)" & vbCrLf
                T1$ = T1$ + "                {" & vbCrLf
                T1$ = T1$ + "                fedate = new DateTime(1980, 01, 01);" & vbCrLf
                T1$ = T1$ + "                }" & vbCrLf
                T1$ = T1$ + "               else" & vbCrLf
                T1$ = T1$ + "                entidad." & NAMEX$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX$ & Chr(34) & "]);" & vbCrLf & vbCrLf
                T1$ = T1$ + "  //-------------------------------------------------------------------------------------------" & vbCrLf

                VectorCargaValoresEntidad(cnt) = T1$
        Else
           TYPEX$ = "String "
           T1$ = "entidad." & NAMEX$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX$ & Chr(34) & "]);" & vbCrLf & vbCrLf
           T1$ = T1$ + "  //-------------------------------------------------------------------------------------------" & vbCrLf
           VectorCargaValoresEntidad(cnt) = T1$
'           VectorCargaValoresEntidad(cnt) = "entidad." & NAMEX$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX$ & Chr(34) & "]);" & vbCrLf & vbCrLf

        End If
        
        Rsrecordset.MoveNext
     Loop
'     Text19 = TX1$
    On Error GoTo 0

End Sub
Sub EscribeEstrucSql(CODICONSULTA$, TEXTOAESCRIBIR$)
'Sub GENERA_ESTRUCTURAS()
   ILET$ = "D"
10 RUT1$ = ILET$ + ":\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT\BIB_EST_SQL\"
   RUT1$ = RUT1$ + "EscribeEstrucSql.bas"
   '
   On Error Resume Next
   If EXISTE%(RUT1$) < 1 Then
     If ILET$ = "D" Then ILET$ = "C": GoTo 10
     MsgBox "Imposible Abrir Archivo: " + RUT1$
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   CARPETOK$ = ILET$ & ":\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT\BIB_EST_SQL\"

   NARCHI1% = FILEOPEN%(CARPETOK$ + "EscribeEstrucSql.bas", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
   NARCHI2% = FILEOPEN%(CARPETOK$ + "EscribeEstrucSql2.bas", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(XX$))
       If XX1$ = "'|INICIO " & CODICONSULTA$ Then NOESCRIBIR% = 1
       If NOESCRIBIR% < 1 Then
          If UCase$(TRIM$(XX1$)) <> "END SUB" Then
            Print #NARCHI2%, XX$
          End If
       End If
    Loop
    Print #NARCHI2%, TEXTOAESCRIBIR$
    Print #NARCHI2%, "End Sub"
    Close #NARCHI1%
    Close #NARCHI2%
    
    On Error Resume Next
    Kill (CARPETOK$ + "EscribeEstrucSql.bas")
    On Error GoTo 0
    '
    'COPIO EL ARCHIVO ESCRITO SOBRE EL ORIGINAL
    XY1% = COPYFILE(CARPETOK$ + "EscribeEstrucSql2.bas", CARPETOK$ + "EscribeEstrucSql.bas")
    '

    '
99
End Sub


Sub GENERA_AGREGA_RECORDSET(CAMPOS_GRI$, Optional VENTANAX$)
    If TRIM$(VENTANAX$) = "" Then
      VENTANAX$ = "VENTABNEW.MSF"
    End If
    
    TX1$ = "  Dim obj As New RECORXET" & vbCrLf
    CAMPOS_BD$ = ""
    For i& = 0 To List3.ListCount - 1
       List3.ListIndex = i&
       CAMPOS_BD$ = CAMPOS_BD$ & List3.TEXT & ","
    Next i&
    CAMPOS_BD$ = TRIM$(CAMPOS_BD$)
    If Right(CAMPOS_BD$, 1) = "," Then CAMPOS_BD$ = Mid$(CAMPOS_BD$, 1, Len(CAMPOS_BD$) - 1)
    TX1$ = TX1$ + "  Set RS1 = obj.RS_" & TRIM$(Text10) & "()" & vbCrLf ' FUNCION QUE DEVUELVE UN RECORDSET" & vbCrLf
    
    TX1$ = TX1$ + "  Call Carga_MSF_Recordset(RS1," & TRIM$(CAMPOS_GRI$) & "," & TRIM$(VENTANAX$) & ")" & vbCrLf
    
    TX1$ = TX1$ + "  Set obj = Nothing" & vbCrLf
    TX1$ = TX1$ + "  RS1.Close" & vbCrLf
    TX1$ = TX1$ + "  Set RS1 = Nothing" & vbCrLf
    
    TX1$ = TX1$ + "'*****************************************************'" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    TX1$ = TX1$ + "ESTO VA EN LA CLASE RECORXET" & vbCrLf
    NOMB_CLASE$ = "RS_" & TRIM$(Text10)
    TX1$ = TX1$ + "Public Function " & NOMB_CLASE$ & "() As ADODB.Recordset " & vbCrLf
    TX1$ = TX1$ + "  Dim RS_FUN As ADODB.Recordset" & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = New ADODB.Recordset" & vbCrLf
    'TX1$ = TX1$ + "  If TRIM$(CAMPOS_BD$) = " & Chr$(34) & Chr$(34) & " Then CAMPOS_BD$ = " & Chr$(34) & " * " & Chr$(34) & vbCrLf
    
    TX1$ = TX1$ + "  SQLX$ = " & Chr$(34) & TXTCONSU_ESPEJO & Chr$(34) & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = READSET(SQLX$)" & vbCrLf
    TX1$ = TX1$ + "  Set " & NOMB_CLASE$ & " = RS_FUN" & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = Nothing" & vbCrLf
    TX1$ = TX1$ + "End Function" & vbCrLf
    TX1$ = TX1$ + "'*****************************************************'" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Text9 = Text9 + TX1$
    
    
 End Sub

 Function GENERA_CODI_MENUGRILLA$()
 
  TXTMG$ = "Dim GLOBALFRM_strContenidoCelda$" & vbCrLf
  TXTMG$ = TXTMG$ + "Dim GLOBALFRM_FILA&" & vbCrLf
  TXTMG$ = TXTMG$ + "Public NORFAEXTERNA$" & vbCrLf
  TXTMG$ = TXTMG$ + "Private Sub mnucopiar_Click()" & vbCrLf
  TXTMG$ = TXTMG$ + "       Clipboard.Clear" & vbCrLf
  TXTMG$ = TXTMG$ + "       Clipboard.SetText GLOBALFRM_strContenidoCelda$" & vbCrLf
  TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf
TXTMG$ = TXTMG$ + "  If Button = 2 Then" & vbCrLf
TXTMG$ = TXTMG$ + "     REG& = MSF1.MouseRow" & vbCrLf
TXTMG$ = TXTMG$ + "     COL% = MSF1.MouseCol" & vbCrLf
TXTMG$ = TXTMG$ + "     Call MENU_CELDA(MSF1, REG&, COL%)" & vbCrLf
TXTMG$ = TXTMG$ + "  End If" & vbCrLf
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
 TXTMG$ = TXTMG$ + " Sub INVICTRL()" & vbCrLf
'  '
TXTMG$ = TXTMG$ + "   mnucopiar.Visible = Not (mnucopiar.Visible)" & vbCrLf
TXTMG$ = TXTMG$ + "  mnudesap.Visible = Not (mnudesap.Visible)" & vbCrLf
TXTMG$ = TXTMG$ + "'  InformaCantidadEnProceso.Visible = Not InformaCantidadEnProceso.Visible" & vbCrLf
'  '
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)" & vbCrLf
TXTMG$ = TXTMG$ + "        If MSF.Rows <= 1 Then Exit Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA" & vbCrLf
TXTMG$ = TXTMG$ + "        Call INVICTRL"
TXTMG$ = TXTMG$ + "        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA" & vbCrLf
TXTMG$ = TXTMG$ + "        GLOBALFRM_FILA& = REG&" & vbCrLf
TXTMG$ = TXTMG$ + "        PopupMenu MNUARCHIVO" & vbCrLf
TXTMG$ = TXTMG$ + "        Call INVICTRL" & vbCrLf
TXTMG$ = TXTMG$ + "        On Error Resume Next" & vbCrLf
TXTMG$ = TXTMG$ + "        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA" & vbCrLf
TXTMG$ = TXTMG$ + "        On Error GoTo 0" & vbCrLf
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
ENCAB$ = "SE DEBERA GENEAR EN EL MENU" & vbCrLf
ENCAB$ = ENCAB$ + "'AGREGAR EN MENU'" & vbCrLf
ENCAB$ = ENCAB$ + "'CAPTION: Archivo - NO VISIBLE" & vbCrLf
ENCAB$ = ENCAB$ + "'....CAPTION -      Copiar Celda Seleccionada" & vbCrLf
ENCAB$ = ENCAB$ + "'.....NAME:  S   - VISIBLE" & vbCrLf
ENCAB$ = ENCAB$ + "'SI SE NECESITA PARA OTRA GRILLA SE DEBERA DUPLICAR LA INFORMACION RENOMBRANDO CADA METODO AGREGANDO UN 2"
GENERA_CODI_MENUGRILLA$ = ENCAB$ & TXTMG$
End Function

Sub GENERAR_CODIGO(OK%)
    If Check10.Value > 0 Then
      If TRIM$(Text10) = "" Then
        Call COMUNI("Falta Título de Recordset")
        OK% = -1
        Exit Sub
      End If
    End If
    '
    If Check10.Value > 0 Then
        POS1% = InStr(TXTCONSU_ESPEJO, "(NOLOCK)")
        If POS1% < 1 Then
           Call COMUNI("Falta With(NOLOCK) en consulta")
           OK% = -1
           Exit Sub
        End If
    End If
    '
    
    T1$ = "VENTANA.Inicializar = 0" + vbCrLf
    Campos$ = ""
    For i& = 1 To MSF.Cols - 1
       If TRIM$(MSF.TextMatrix(0, i&)) <> "" Then Campos$ = Campos$ & MSF.TextMatrix(0, i&) + ";"
    Next i&
    If Right$(Campos$, 1) = ";" Then Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    Campos$ = Chr$(34) & Campos$ & Chr$(34) & vbCrLf
    For i& = 1 To MSF.Rows - 1
       PROPIEDADX$ = MSF.TextMatrix(i&, 0)
       Select Case PROPIEDADX$
          Case "Formula (4;*;6)"
            For J% = 1 To MSF.Cols - 1
              If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 T0$ = "VENTANA.formula(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                 If TRIM$(T0$) <> "" Then If Right$(T0$, 1) = ";" Then T0$ = Mid$(T0$, Len(T0$) - 1)
                 T1$ = T1$ + T0$
              End If
            Next J%
            
          Case "CampEditables (S/N)"
              For J% = 1 To MSF.Cols - 1
               If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 'VENTANA.CampEditables(" & J% & ") = " & MSF.TextMatrix(i&, J%) + vbCrLf"
                 If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then T2$ = T2$ & J% & ";"
               End If
              Next J%
              If TRIM$(T2$) <> "" Then
                 If Right$(T2$, 1) = ";" Then T2$ = Mid$(T2$, 1, Len(T2$) - 1)
'                 T2$ = Chr$(34) & T2$ & Chr$(34)
                 T1$ = T1$ + "VENTANA.CampEditables = " & Chr$(34) & T2$ & Chr$(34) & vbCrLf
              End If
              
          Case "TxtUbicadoGrillaTotal(Texto)"
              For J% = 1 To MSF.Cols - 1
               If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 H% = J%
                 'VENTANA.TxtUbicadoGrillaTotal(J%) = MSF.TextMatrix(i&, J%)
                 T3$ = "VENTANA.TxtUbicadoGrillaTotal(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) + vbCrLf
                 If TRIM$(T3$) <> "" Then If Right$(T3$, 1) = ";" Then T3$ = Mid$(T3$, Len(T3$) - 1) + vbCrLf
                 T1$ = T1$ + T3$
               End If
              Next J%
          Case "ColEcoCampo (CODIGO;Descripcion;Master)"
              For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.ColEcoCampo(J%) = MSF.TextMatrix(i&, J%)
                  T4$ = "VENTANA.ColEcoCampo(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T4$
''                  If TRIM$(MSFESPEJO.TextMatrix(i&, J%)) <> "" Then
''                    T4$ = T4$ & ";" & TRIM$(MSFESPEJO.TextMatrix(i&, J%)) & Chr$(34) & vbCrLf
''                    T1$ = T1$ + T4$
''                  End If
                End If
              Next J%
           Case "Total (S/N)"
              For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T5$ = "VENTANA.Totales(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T5$
                End If
              Next J%
           Case "Dejar Libre (S/N)"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T6$ = "VENTANA.DejarLibre(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T6$
                End If
              Next J%
           Case "CONDICION 1"
               For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.Condicion1(J%) = MSF.TextMatrix(i&, J%)
                  T7$ = "VENTANA.Condicion1(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T7$
                End If
              Next J%
           Case "CONDICION 2"
               For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.Condicion2(J%) = MSF.TextMatrix(i&, J%)
                  T8$ = "VENTANA.Condicion2(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T8$
                End If
              Next J%
           Case "ENUMERAR"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.Enumerar(J%) = MSF.TextMatrix(i&, J%)
                  T9$ = "VENTANA.ENUMERAR(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T9$
                End If
              Next J%
         
       End Select
    Next i&
    

    T1$ = T1$ + "VENTANA.AgregaRegistro = " & Check5.Value & vbCrLf
    T1$ = T1$ + "VENTANA.UtilizaFormula = " & Check6.Value & vbCrLf
    T1$ = T1$ + "VENTANA.HabilitaBorrar = " & Check7.Value & vbCrLf
    T1$ = T1$ + "VENTANA.HabilitaInsertar = " & Check8.Value & vbCrLf
    T1$ = T1$ + "VENTANA.NoMostrarBotonTilde = " & Check9.Value & vbCrLf
    T1$ = T1$ + "VENTANA.Titulo = " & Chr$(34) & TXTTITULO.TEXT & Chr$(34) & vbCrLf
    T1$ = T1$ + "VENTANA.Inicializar = 1" & vbCrLf
    T1$ = T1$ + "VENTANA.Campox = " & Campos$ & vbCrLf
    '
    If Check11.Value > 0 Then
         T1$ = T1$ + "*** SE DEBERA CARGAR PARA EL PARAMETRO  VENTANA.Campox = LOS CAMPOS DEL MSF BASE  ****" + vbCrLf
         T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
         T1$ = T1$ + "*** SE DEBERA CARGAR PARA LOS PARAMETROS CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)   ****" + vbCrLf
         T1$ = T1$ + "*** REALIZAR LA LLAMADA " + vbCrLf
         T1$ = T1$ + "Call CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)"
    End If
    Text9 = T1$
    T1$ = T1$ & vbCrLf
    T1$ = T1$ & "'"
    T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
    Text9 = T1$
    
    If Check10.Value > 0 Then
       Call GENERA_AGREGA_RECORDSET(Campos$)
    End If
    If Check11.Value > 0 Then T1$ = T1$ & "Call T_Espera(1 / 10)" & vbCrLf
    
    Text9 = Text9 & "VENTABNEW.SHOW 1" + vbCrLf
    
    If Check11.Value > 0 Then
      Text9 = Text9 & "If VENTABNEW.APROBADO% > 0 Then" + vbCrLf
      Text9 = Text9 + "   Call CopiarMsfAOtroMsf(VENTADISE.MSF,GRILLABASE )" + vbCrLf
      Text9 = Text9 & "End If" + vbCrLf
    End If
    '
    GECOMENUGRI$ = ""
    If Check12.Value = 1 Then
       GECOMENUGRI$ = GENERA_CODI_MENUGRILLA$
       Text9 = Text9 & GECOMENUGRI$
    End If

    Text9 = Text9 & "Set VENTANA = Nothing"
    '
End Sub

Function GenerarBotonGrabar$(ind, tp, lf)
             INDICEX% = ind + 1
             TOPX = tp + 20
             LEFTX = lf
             
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & vbCrLf
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "//" & vbCrLf
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "// BTNGRABAR" & vbCrLf
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "//" & vbCrLf
'             this.BTNGRABARDocumentoOrigen.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf

             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Location = new System.Drawing.Point(" & LEFTX & "," & TOPX & ");" & vbCrLf
'                  this.button1.Name = "button1";
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Name = " & Chr$(34) & "BTNGRABAR" & Chr$(34) & ";" & vbCrLf
'            this.button1.Size = new System.Drawing.Size(20, 20);
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Size = new System.Drawing.Size(100, 20);" & vbCrLf
'                 this.button1.TabIndex = 6;
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".TabIndex = " & INDICEX% & " ;" & vbCrLf
'                 this.button1.Text = "button1";
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Text = " & Chr$(34) & "." & Chr$(34) & ";" & vbCrLf
             
             'this.BTNECOPEDENCA_NroClie.Click += new System.EventHandler(this.BTNECOPEDENCA_NroClie_Click);
             GenerarBotonGrabar$ = GenerarBotonGrabar$ & "this.BTNGRABAR" & ".Click += new System.EventHandler(this.BTNGRABAR" & "_Click);" & vbCrLf
             
             VectorControlAdd(INDICEX) = "this.Controls.Add(this.BTNGRABAR" & ");"
             VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.Button BTNGRABAR" & ";"

             INDICEX% = INDICEX% + 1

End Function

Function GenerarBotonNuevo$(ind, tp, lf)
             LEFTX = lf
             TOPX = tp + 20
             INDICEX% = ind + 1
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & vbCrLf
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "//" & vbCrLf
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "// BTNNUEVO" & vbCrLf
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "//" & vbCrLf
'             this.BTNNUEVODocumentoOrigen.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf

             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Location = new System.Drawing.Point(" & LEFTX & "," & TOPX & ");" & vbCrLf
'                  this.button1.Name = "button1";
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Name = " & Chr$(34) & "BTNNUEVO" & Chr$(34) & ";" & vbCrLf
'            this.button1.Size = new System.Drawing.Size(20, 20);
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Size = new System.Drawing.Size(100, 20);" & vbCrLf
'                 this.button1.TabIndex = 6;
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".TabIndex = " & INDICEX% & " ;" & vbCrLf
'                 this.button1.Text = "button1";
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Text = " & Chr$(34) & "." & Chr$(34) & ";" & vbCrLf
             
             'this.BTNECOPEDENCA_NroClie.Click += new System.EventHandler(this.BTNECOPEDENCA_NroClie_Click);
             GenerarBotonNuevo$ = GenerarBotonNuevo$ & "this.BTNNUEVO" & ".Click += new System.EventHandler(this.BTNNUEVO_Click);" & vbCrLf
             
             VectorControlAdd(INDICEX) = "this.Controls.Add(this.BTNNUEVO" & ");"
             VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.Button BTNNUEVO" & ";"

             INDICEX% = INDICEX% + 1

End Function
Function GenerarCodigoVectorLLamadaZelecho$(BTNNAME$, TABLA$, Campos$, CONDI$, FILAMSF3&)
    
    Campos$ = REPLACAR(Campos$, ";", ",")
    CmpoIdRepli$ = CampoIdReplica$
    LONGIT% = Len(CmpoIdRepli$)
    
    LBLECO$ = REPLACAR(BTNNAME$, "BTNECO", "LBLECO")
    If Left$(UCase$(TRIM$(Campos$)), LONGIT%) = TRIM$(UCase$(CmpoIdRepli$)) Then
       LBLECO$ = REPLACAR(BTNNAME$, "BTNECO", "TXTECO")
    End If
    TXT1$ = REPLACAR(BTNNAME$, "BTNECO", "TXT")
    
    X1$ = "private void " & BTNNAME$ & "_Click (object sender, EventArgs e)" & vbCrLf
    X1$ = X1$ & "             {" & vbCrLf
    X1$ = X1$ & "            " & BTNNAME$ & ".Enabled = false;" & vbCrLf
    X1$ = X1$ & "            XZELECHO.XZELECHO FZ = new XZELECHO.XZELECHO();" & vbCrLf
    X1$ = X1$ & "            FZ.pasado += (new XZELECHO.XZELECHO.pasar(Mostrar" & BTNNAME$ & "));//ASIGNO AL EVENTO PASADO EL DELEGADO PASAR PARA QUE CUMPLA LA ACCION MostrarBTNECONroPed" & vbCrLf
    X1$ = X1$ & "            try" & vbCrLf
    X1$ = X1$ & "                {" & vbCrLf
    X1$ = X1$ & "                //CARGA EL CONTROL DE USUARIO EN UN FRM" & vbCrLf
    X1$ = X1$ & "                f.Controls.Add(FZ);" & vbCrLf
    X1$ = X1$ & "                f.Size = FZ.Size;" & vbCrLf
    X1$ = X1$ & "                f.ShowInTaskbar = false;" & vbCrLf
    X1$ = X1$ & "                FZ.Tabla = " & Chr(34) & TABLA$ & Chr(34) & ";" & vbCrLf
    X1$ = X1$ & "                FZ.Campostabla  = " & Chr(34) & Campos$ & Chr(34) & ";" & vbCrLf
    X1$ = X1$ & "                FZ.Condi = " & Chr(34) & CONDI$ & Chr(34) & ";" & vbCrLf
    X1$ = X1$ & "                FZ.StringCamposGrid = " & Chr(34) & StringrHeader_Formato$(FILAMSF3&) & Chr(34) & ";" & vbCrLf 'Funcion que crea un string con los formatos y el rotulo Ej.  Nro Cliente/F8;Razon Social/A32'
    X1$ = X1$ & "                FZ.CargarXZelecho();" & vbCrLf
    
    X1$ = X1$ & "                f.ShowDialog();" & vbCrLf
    X1$ = X1$ & "                FZ.Dispose();" & vbCrLf
    X1$ = X1$ & "                }" & vbCrLf
    X1$ = X1$ & "            catch" & vbCrLf
    X1$ = X1$ & "                {// CONTROLO EL ERROR POR QUE ALTERNATIVAMENTE TIRA ERROR 'No se controlo NullReferenceException'" & vbCrLf
    X1$ = X1$ & "                //MessageBox.Show(" & Chr(34) & "Vuelva a Intentarlo" & Chr(34) & ");" & vbCrLf
     X1$ = X1$ & "                return;" & vbCrLf
    X1$ = X1$ & "                }" & vbCrLf
    X1$ = X1$ & "            " & BTNNAME$ & " .Enabled = true;" & vbCrLf
    X1$ = X1$ & "            }" & vbCrLf
    X1$ = X1$ & "        //===============================================================================================================" & vbCrLf
    X1$ = X1$ & "        //===============================================================================================================" & vbCrLf
    X1$ = X1$ & "        public void Mostrar" & BTNNAME$ & " (string[] dato)" & vbCrLf
    X1$ = X1$ & "            {" & vbCrLf
    X1$ = X1$ & ""
    X1$ = X1$ & "            " & TXT1$ & " .Text = dato[0].ToString();" & vbCrLf
    X1$ = X1$ & "            " & LBLECO$ & ".Text = dato[1].ToString();" & vbCrLf
    X1$ = X1$ & "            f.Hide();"
    X1$ = X1$ & "            }" & vbCrLf
    
    X1$ = X1$ & "        //===============================================================================================================" & vbCrLf
    X1$ = X1$ & "        //===============================================================================================================" & vbCrLf
    
    
    GenerarCodigoVectorLLamadaZelecho$ = X1$
End Function

Sub GRABASQLDAT(SQLX$, RET_CAMPO$, RET_TABLA$, RET_CONDI$)
   CONSU$ = UCase$(SQLX$)
   POS1% = InStr(CONSU$, " ")
   POS2% = InStr(CONSU$, " FROM ")
   POS3% = POS2% + 6 ' DONDE COMIENZA FROM + LONGITUD DE FROM Y 2 ESPACIOS
   POS4% = InStr(CONSU$, "WITH(NOLOCK)")
   POS5% = InStr(CONSU$, "WHERE ")
   POS6% = InStr(CONSU$, "ORDER ")
   POS5% = POS5% + 6
   RET_CAMPO$ = Mid$(CONSU$, POS1%, POS2% - POS1%)
   RET_TABLA$ = Mid$(CONSU$, POS3%, POS4% - POS3%)
   If POS6% > 0 Then
     RET_CONDI$ = Mid$(CONSU$, POS5%, POS6%)
   Else
     RET_CONDI$ = Mid$(CONSU$, POS5%)
   End If
End Sub

Function CampoIdReplica$()
     CampoIdReplica$ = ""
     For i& = 1 To MSF3.Rows - 1
      CampoIdReplica$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 6)))
      If CampoIdReplica$ = "SI" Then
         CampoIdReplica$ = TRIM$(MSF3.TextMatrix(i&, 1))
         Exit For
      End If
   Next i&
End Function

Sub LEERYESCRIBIR(RutayNombreLee$, Optional RutayNombreEscribe$, Optional OBJETOESCRITURA As TextBox)
   'Si se le pasa archivo de donde escribir escribe en este
   'si se le pasa un texto box donde imprmir tambien impreme en el textbox
   Screen.MousePointer = 11
   EXISTE_EN_LISTA% = 0
   If EXISTE%(RutayNombreLee$) < 1 Then
     Call COMUNI("Archivo de Lectura Inexistente")
     Exit Sub
   End If
   '
   If TRIM$(RutayNombreEscribe$) <> "" Then
    If EXISTE%(RutayNombreEscribe$) < 1 And OBJETOESCRITURA.Name = "" Then
      Call COMUNI("Archivo de Escritura y Textbox  Inexistente ")
      Exit Sub
    End If
   ElseIf OBJETOESCRITURA.Name = "" Then
     Call COMUNI("Archivo de Escritura y Textbox  Inexistente ")
   End If
   '
   TEXTOTEXBOX$ = ""
   EscribeArchivo% = 0: EscribeTB% = 0
   On Error Resume Next
   If EXISTE%(RutayNombreEscribe$) > 0 Then
       EscribeArchivo% = 1:
       Aescribir$ = RutayNombreEscribe$
       NARCHI3% = FILEOPEN%(Aescribir$, 3, 128)
   End If
   
   If OBJETOESCRITURA.Name <> "" Then
      EscribeTB% = 1
   End If
   Aleer$ = RutayNombreLee$
   On Error GoTo 0

   NARCHI2% = FILEOPEN%(Aleer$, 1, 128) ' ABRE EL ARCHIVO DE ESCRITURA
     Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
       If EscribeArchivo% > 0 Then
          Print #NARCHI3%, , XX$
       End If
       If EscribeTB% > 0 Then
          TEXTOTEXBOX$ = TEXTOTEXBOX$ & XX$ & vbCrLf
       End If

    Loop
    
    On Error Resume Next
    Close #NARCHI3%
    Close #NARCHI2%
    On Error GoTo 0
    
    If EscribeTB% > 0 Then OBJETOESCRITURA.TEXT = TEXTOTEXBOX$
    '
   
    Screen.MousePointer = 1
End Sub

Sub MOSTRARCONFI(IDLIS&)
'  Dim obj As New RECORXET
'  Set RS1 = obj.RS_MUESTRAENCAPRF(IDLIS&)
'  CAMPOS$ = "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10"
'
'  Call Carga_MSF_Recordset(RS1, CAMPOS$, VENTADISE.MSF4)
'  Set obj = Nothing
'  RS1.Close
'  Set RS1 = Nothing

End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Function System_Windows_Forms$()
   For i% = 1 To UBound(VectorSystem_Windows_Forms) - 1
     System_Windows_Forms = vbCrLf & VectorSystem_Windows_Forms(i%)
   Next i%


End Function

Function CargaVectorControlAdd$()
   For i% = 1 To UBound(VectorControlAdd)
     If TRIM$(VectorControlAdd(i%)) <> "" Then
        CargaVectorControlAdd$ = CargaVectorControlAdd$ & vbCrLf & VectorControlAdd(i%)
     End If
   Next i%
End Function
Function CargaVectorSystem_Windows_Forms$()
   For i% = 1 To UBound(VectorSystem_Windows_Forms)
     If TRIM$(VectorSystem_Windows_Forms(i%)) <> "" Then
        CargaVectorSystem_Windows_Forms$ = CargaVectorSystem_Windows_Forms$ & vbCrLf & VectorSystem_Windows_Forms(i%)
     End If
   Next i%
End Function

'
Function TIPOCLARO$(TYPEX$)

   Select Case TRIM$(TYPEX$)
      Case "smalldatetime"
         TIPOCLARO$ = "FECHA"
      Case "smallint"
         TIPOCLARO$ = "ENTERO CORTO"
      Case "ntext"
         TIPOCLARO$ = "TEXTOLARGO"
       Case "int"
         TIPOCLARO$ = "ENTERO"
       Case "float"
         TIPOCLARO$ = "Decimal"
       Case "nvarchar"
         TIPOCLARO$ = "TEXTO"
      Case Else
         TIPOCLARO$ = TYPEX$
   End Select
End Function

Function tipodatoSql$(TABLA$, CAMPO$)
    'DEVUELVE EL TIPO DE DATO QUE ES UN CAMPO LO USO IGUAL QUE AL ASIGNAR A LAS ENTIDADES
    If TABLA$ = "" Then Exit Function
    Set Rsrecordset = FLEXCONN.Execute("sp_columns [" & TABLA$ & "]")
    CAMPO$ = TRIM$(CAMPO$)
    Do Until Rsrecordset.EOF
        cnt = cnt + 1
        NAMEX$ = TRIM$(Rsrecordset!column_name)
        If CAMPO$ = NAMEX$ Then
                TYPEX$ = Rsrecordset!type_name
                LENGX$ = Rsrecordset!Length
                TCLARO$ = TIPOCLARO$(TYPEX$)
                TDT$ = TRIM$(UCase$(TYPEX$))
                If Left$(TDT$, 3) = "INT" Or Left$(TDT$, 8) = "SMALLINT" Or Left$(TDT$, 5) = "FLOAT" Then
                   Select Case TDT$
                     Case "INT"
                       If XVALO(LENGX$) < 4 Then
                          TYPEX$ = "Int16"
                       Else
                          TYPEX$ = "Int32"
                       End If
                     Case "SMALLINT": TYPEX$ = "Int16"
                           
                     Case "FLOAT": TYPEX$ = "Decimal"
                      
                    End Select
                ElseIf Left$(TDT$, 4) = "DATE" Or Left$(TDT$, 13) = "SMALLDATETIME" Then
                   'texto de fecha para concatenar mas abajo
                   TYPEX$ = "DateTime"
                Else
                   TYPEX$ = "String"
                End If
                tipodatoSql$ = TYPEX$
                Exit Do
         End If
         Rsrecordset.MoveNext
     Loop
    On Error GoTo 0



End Function

Function VALIDARCAMPOSYCONSULTA%(TABLA$, Campos$, CONSULTASQL$)
   TABLA$ = TRIM$(TABLA$)
   Campos$ = TRIM$(Campos$)
   Campos$ = REPLACAR(Campos$, ",", ";")
   VALIDARCAMPOSYCONSULTA% = 1

   'valido TABLA
   If EXISTE_TABLA%(TABLA$) < 1 Then Call COMUNI("Tabla Inexistente"): VALIDARCAMPOSYCONSULTA% = 0: Exit Function
   
   'VALIDO CAMPO
   strArray = Split(Campos$, ";")
   For i% = 0 To UBound(strArray)
      CAMPO$ = strArray(i%)
      If EXISTE_CAMPO%(TABLA$, CAMPO$) < 1 Then Call COMUNI("Campo " & CAMPO$ & " Inexistente"): VALIDARCAMPOSYCONSULTA% = 0: Exit Function
   Next i%
   Campos$ = REPLACAR(Campos$, ";", ",")
   'SI LLEGA ACA PASO LAS VALIDACIONES DE TABLA Y CAMPOS
   'VALIDA LA CONSULTA
   SQLX$ = TRIM$(CONSULTASQL$)
   On Error Resume Next
   POS1NOLOCK% = InStr(SQLX$, "(NOLOCK)")
   If POS1NOLOCK% > 0 Then
     DESPUESDENOLOCK$ = UCase$(TRIM$(Mid$(SQLX$, POS1NOLOCK% + 8)))
     If Left$(DESPUESDENOLOCK$, 5) <> "ORDER" Then
        SQLX$ = Mid$(SQLX$, 1, POS1NOLOCK% + 7)
        If TRIM$(DESPUESDENOLOCK$) <> "" Then
           SQLX$ = SQLX$ & " WHERE " & DESPUESDENOLOCK$
        End If
     End If
   End If
   Set RS = READSET(SQLX$, , 1)
   If Err.Number > 0 Then
       Call COMUNI("Atención!!!Se Detecta un Error en La Consulta Para el Campo: " & Campos$ & "\\Cancelar\Continuar y Corregir Posteriormente")
       VALIDARCAMPOSYCONSULTA% = 0
   End If
   On Error GoTo 0
End Function

Function StringrHeader_Formato$(FILAMSF3&)

   'crea un string con los formatos y el rotulo Ej.  Nro Cliente/F8;Razon Social/A32'
   StringrHeader_Formato$ = ""
   FORMATX$ = TRIM$(MSF3.TextMatrix(FILAMSF3&, 13))
   ID_01$ = UCase$(TRIM$(MSF3.TextMatrix(FILAMSF3&, 11)))
   If ID_01$ = "SI" And FILAMSF3& > 0 Then
          If FORMATX$ <> "" Then
            StringrHeader_Formato$ = FORMATX$
          Else
            Call COMUNI("Falta Formato o Rotulo en Fila: " & i&)
          End If
   End If
End Function


Private Sub cmd33_Click()
 RutaFormulario$ = SELECCION_Archivo$("frm", CMD1)
 If RutaFormulario$ <> "" Then
    Text18(1) = RutaFormulario$
    POSRIGHT% = InStrRev(RutaFormulario$, "\")
    TXTNAMEFRM = Right(RutaFormulario$, Len(RutaFormulario$) - POSRIGHT%)
    Call LEERYESCRIBIR(Text18(1), "", Text22(0))
 End If
End Sub

Private Sub CMDAPPCAMPOS_Click()
   CMDAPPCAMPOS.Enabled = False
   MSFUPD.Rows = 5
   SEPARADOR$ = TRIM$(Me.TXTSEPARADOR)
   Campos$ = TRIM$(Me.TXTCAMPOSUPDATE)
   Call TEXTOLOTES$(Campos$, SEPARADOR$)
   On Error Resume Next
   CANTICAMP = UBound(CADENATEX$)
   On Error GoTo 0
   If XVALO(CANTICAMP) > 0 Then
     MSFUPD.Cols = CANTICAMP + 1
     For i& = 1 To UBound(CADENATEX$)
        MSFUPD.ColWidth(i&) = 3500
        MSFUPD.TextMatrix(0, i&) = CADENATEX$(i&)
     Next i&
   
   End If
   
   CMDAPPCAMPOS.Enabled = True
End Sub

Private Sub CmdGrabar_Click()
        Dim MSF As MSFlexGrid
        Set MSF = MSF3
        '
        FRMNAMEX$ = TRIM$(TXTNAMEFRM)
        CmdGrabar.Enabled = False
        If FRMNAMEX$ = "" Then
           Call COMUNI("Falta Seleccionar Formulario")
           GoTo 99
        End If
        If MSF.Rows < 2 Then
           Call COMUNI("Sin datos en Grilla")
           GoTo 99
        End If
        
        CONDI$ = "FRMNAME = '" & FRMNAMEX$ & "'"
        
        YAESTA% = CantRegistros("CREACOD", CONDI$, "NOMESS")
        If YAESTA% > 0 Then
           OPX% = COMALTER%("Atención Ya Existe un Formulario con Mismo Nombre en la Base de Datos\\Cancelar\Grabar con Nuevo Código\Actualizar Existente")
           If OPX% < 2 Then GoTo 99
        End If
        '
        If OPX% = 2 Then
           LBLNUEVO.Visible = True
           NEWCOD = NumNuevaCreacod
        Else
           LBLNUEVO.Visible = False
           NEWCOD = XVALO(TXTCODIGO)
        End If
        '
        If LBLNUEVO.Visible = True Then
           ESNUEVO% = 1
        End If
        
        SQLX$ = "SELECT * FROM CREACOD "
        SQLX$ = SQLX$ + " WHERE CODIGO  = " & NEWCOD
        '
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        Call ACTUREGISTRO("CREACOD", "MARBORRA", "CODIGO  = " & NEWCOD, 1, REGAF&, "NOMESS")
        VUELTA% = 0
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
         For REG& = 1 To MSF.Rows - 1
            .AddNew
            !Codigo = NEWCOD
            !FECHEMISI = CVDAT(HOY(HO$))
            !CAMPO = MSF.TextMatrix(REG&, 1)
            !T_DATO = MSF.TextMatrix(REG&, 2)
            !LONGITUD = XVALO(MSF.TextMatrix(REG&, 3))
            !ROTULO = MSF.TextMatrix(REG&, 4)
            !Formato = MSF.TextMatrix(REG&, 5)
            !ID_REPLICA = MSF.TextMatrix(REG&, 6)
            !ESCAMPOECO = MSF.TextMatrix(REG&, 7)
            !Frame = MSF.TextMatrix(REG&, 8)
            !Graba = MSF.TextMatrix(REG&, 9)
            !MUESTRA_RS = MSF.TextMatrix(REG&, 10)
            !TABLA = Text18(0)
            !ID01 = Left$(MSF.TextMatrix(REG&, 11), 24) 'IDENTIFICACION DE GRABACION
            !SQLXCAMPOECO = Left$(MSF.TextMatrix(REG&, 12), 1024) 'CONSULTA DEL CAMPO ECO
            !ENCABEZADOSZELECHO = Left$(MSF.TextMatrix(REG&, 13), 256)
            !REFERENCIA = Left$(Text24, 48)
            !FRMNAME = TXTNAMEFRM
            
          .Update
         Next REG&
        End With
        '
        RS.Close
        Set RS = Nothing
        '
        CONDI$ = "CODIGO  = " & NEWCOD & " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("CREACOD", CONDI$, REGAF&, "NOMESS")
        
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION ")
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    TXTCODIGO = NEWCOD
99  CmdGrabar.Enabled = True

End Sub

Function NumNuevaCreacod()
    NumNuevaCreacod = 1 + XVALO(VALOBADA("Creacod", "TOP 1 codigo", "codigo > 0 ORDER BY codigo  DESC", "NOMESS"))
End Function




Private Sub Command16_Click()
VENTANA.Inicializar = 0
VENTANA.Formula(5) = "3;*;4"
VENTANA.CampEditables = "1;3;4"
VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
VENTANA.TOTALES(5) = "S"
VENTANA.AgregaRegistro = 1
VENTANA.UtilizaFormula = 1
VENTANA.HabilitaBorrar = 1
VENTANA.HabilitaInsertar = 1
VENTANA.NoMostrarBotonTilde = 0
VENTANA.TITULO = "PEDIDOS MUESTRA"
VENTANA.Inicializar = 1
VENTANA.Campox = "CODIGO/A18;DESCRIPCION/A32;CANTIDAD/F12.1;PRECIO/F12.2;IMPORTE/F12.2"




  Dim obj As New RECORXET
  Set RS1 = obj.RS_PEDIDOS1()
  Call Carga_MSF_Recordset(RS1, "CODIGO/A18;DESCRIPCION/A32;CANTIDAD/F12.1;PRECIO/F12.2;IMPORTE/F12.2", VENTABNEW.MSF)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing




'*****************************************************'




VENTABNEW.Show 1
If VENTABNEW.APROBADO% > 0 Then

End If
Set VENTANA = Nothing
End Sub
Private Sub Command17_Click()
    Command17.Enabled = False
    CamposRecordset$ = ""
    'lista de campos de recordset separados por el palito
    For i& = 0 To List3.ListCount - 1
       List3.ListIndex = i&
       CamposRecordset$ = CamposRecordset$ & List3.TEXT & "|"
    Next i&
    CamposRecordset$ = TRIM$(CamposRecordset$)
    If Right(CamposRecordset$, 1) = "|" Then CamposRecordset$ = Mid$(CamposRecordset$, 1, Len(CamposRecordset$) - 1)

    
    CONDI$ = "ID = " & XVALO(TXTID)
'    Call ACTUREGISTRO("GENECODI", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("DatGrilla", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    IDX& = XVALO(TXTID)
    Dim RS As New ADODB.Recordset
    SQLX$ = "SELECT *  FROM GeneCodi WHERE ID = " & XVALO(TXTID)
    RS.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS
      If XVALO(TXTID) < 1 Then
         IDX& = PROXIMO_ID("GeneCodi", "ID", "ID > 0") ' FUNCION
        .AddNew
        !ID = IDX&
      End If
      
      !TitVentana = TXTTITULO
      !ADDREGISTROS = Check5.Value
      !UtilFormula = Check6.Value
      !Borra = Check7.Value
      !Inserta = Check8.Value
      !NOTILDE = Check9.Value
      !NameRecord = TRIM$(Text10.TEXT)
      !CamposRecord = XVALO(CamposRecordset$)
      !USARECORD = Check10.Value
      !AotraGrilla = Check11.Value
      !ConsultaRec = XVALO(TXTCONSU_ESPEJO)
      !CodiGen = Text9
      !CanCampos = XVALO(Text8)
      .Update
    End With
    RS.Close: Set RS = Nothing
    '
    'ACTUALIZO EL NUEVO ID
    TXTID = IDX&

    CONDI$ = "ID = " & XVALO(TXTID)
    Call BORRAREGISTROS("DATGRILLA", CONDI$, REGAF&, "NOMESS")
    Dim RS1 As New ADODB.Recordset
    SQLX$ = "SELECT * FROM DatGrilla WHERE ID =" & XVALO(TXTID)
    RS1.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS1
        For i& = 0 To MSF.Rows - 1
          For J& = 1 To MSF.Cols - 1
             DATOX$ = TRIM$(MSF.TextMatrix(i&, J&))
             If DATOX$ <> "" Then
                .AddNew
                !ID = IDX&
                !FilaCol = SAFETEXT$(i&) & "|" & SAFETEXT$(J&)
                !DATO = DATOX$
                 .Update
              End If
          Next J&
        Next i&
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    '
99  Command17.Enabled = True
End Sub

Private Sub Command18_Click()


    Command18.Enabled = False
    If Check10.Value > 0 Then
      If TRIM$(Text10) = "" Then
        Call COMUNI("Falta Título de Recordset")
        OK% = -1
        GoTo 99
      End If
    End If
    '
    If Check10.Value > 0 Then
        POS1% = InStr(TXTCONSU_ESPEJO, "(NOLOCK)")
        If POS1% < 1 Then
           Call COMUNI("Falta With(NOLOCK) en consulta")
           OK% = -1
           GoTo 99
        End If
    End If
    '
    Campos$ = ""
    For i& = 1 To MSF.Cols - 1
       If TRIM$(MSF.TextMatrix(0, i&)) <> "" Then Campos$ = Campos$ & MSF.TextMatrix(0, i&) + ";"
    Next i&
    If Right$(Campos$, 1) = ";" Then Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    Campos$ = Chr$(34) & Campos$ & Chr$(34) & vbCrLf
    T1$ = "Call CARGA_ENCABEZADO(FRMZELECHO.msf2," & TRIM$(Campos$) & ", FRMZELECHO)" + vbCrLf
    T1$ = T1$ + "Call CARGA_ENCABEZADO(FRMZELECHO.msf1," & TRIM$(Campos$) & ", FRMZELECHO)" + vbCrLf

    For i& = 0 To MSF.Rows - 1
       PROPIEDADX$ = MSF.TextMatrix(i&, 0)
       Select Case PROPIEDADX$
           Case "Dejar Libre (S/N)"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T6$ = "VENTANA.DejarLibre(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T6$
                End If
              Next J%
       End Select
    Next i&


    T1$ = T1$ + "FRMZELECHO.CAPTION = " & Chr$(34) & TXTTITULO.TEXT & Chr$(34) & vbCrLf
    '
    Text9 = T1$
    T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
    Text9 = T1$
    If Check10.Value > 0 Then
       Call GENERA_AGREGA_RECORDSET(Campos$, "FRMZELECHO.MSF2")
    End If
    Text9 = Text9 & "FRMZELECHO.WIDTH = 8500 " + vbCrLf
    Text9 = Text9 & "FRMZELECHO.SHOW 1" + vbCrLf
    '
    Text9 = Text9 & "NRO = XVALO(RESP_ZELE_VECT(1))" + vbCrLf

    NRO = XVALO(RESP_ZELE_VECT(1))

99  Command18.Enabled = True
End Sub

Private Sub Command19_Click()
  TXTID = ""
End Sub

Sub LIMPIA_EDICION()
    Me.TXTID = ""
    Me.TXTIDLISTA.TEXT = ""
'    Me.TXTNOMBRELISTA.TEXT = ""
'    Me.TXTORDEN = ""
'    Me.TXTORDPRES.TEXT = ""
'    Me.TXTTIPOID = ""
    Me.CMBCAMPOS = ""
'    Me.TXTFORMATO.TEXT = ""
'    Me.CMBNOMBRECAMPO = ""
End Sub

Private Sub Command20_Click()
'    Unload VENTABNEW
'    Campos$ = "CAMPO/A28;" '1
'    Campos$ = Campos$ + "T.DATO/A12;" '2
'    Campos$ = Campos$ + "Len/f3;" '3
'    Campos$ = Campos$ + "ROTULO/A16;" '4
'    Campos$ = Campos$ + "Len/f3;" '5
'
'    Campos$ = Campos$ + "FORMATO/A9;" '6
'    Campos$ = Campos$ + "ID_REPLICA/A9;" '7   SI O NO
'    Campos$ = Campos$ + "ESCAMPOECO/A8;" '8  SI O NO AGREGAR PASAR A OTRA GRILLA LOS CAMPOS ECO PARA ASIGNAR LAS CONSULTASA
'    Campos$ = Campos$ + "FRAME/A7;" '9  NOMBRE DEL FRAME O NUMERO
'    Campos$ = Campos$ + "Graba/A7;" '10
'    Campos$ = Campos$ + "MUESTRA_RS/A10" '11
'    Campos$ = Campos$ + " ID01/F5" '12
'    VENTANA.Inicializar = 0
'    VENTANA.Campox = Campos$
'    'VENTANA.AgregaRegistro = 1
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaInsertar = 0
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.TITULO = "Armado Inicial"
'    VENTANA.CampEditables = "4;5;6;7;8;9;10;11;12"
'    VENTANA.Inicializar = 1
    For J& = 0 To List5.ListCount - 1
      If List5.Selected(J&) = True Then
        L& = L& + 1
        List5.ListIndex = J&
        campoleido$ = List5.TEXT
        CAMPO1$ = TRIM$(Left$(campoleido$, 32))
        campo2$ = TRIM$(Mid$(campoleido$, 34, 14))
        CAMPO3$ = TRIM$(Mid$(campoleido$, 50, 16))
        MSF3.Rows = L& + 1
        
        MSF3.TextMatrix(L&, 1) = CAMPO1$
        MSF3.TextMatrix(L&, 2) = campo2$
        MSF3.TextMatrix(L&, 3) = CAMPO3$
      End If
    Next J&
       
'    VENTABNEW.Show 1
'    If VENTABNEW.APROBADO% > 0 Then
'       Call CopiarMsfAOtroMsf(MSF3, MSF3, 1)
'    End If
'    Set VENTANA = Nothing
'    Unload VENTABNEW

End Sub

Private Sub Command21_Click()
   Text17.Width = Text17.Width + 100
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    '
    LISTAX$ = TRIM$(Text14)
    IDLIST& = Me.TXTIDLISTA
    'CARGO LOS DATOS UNICOS DE ENCABEZADOS
    Dim obj As New RECORXET
    Set RS1 = obj.RS_TAB_COMP_ENCAPRFUN(IDLIST&)
    MAXIMOCAMPOS = XVALO(RS1!CANTICAMPOS)
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = XVALO(RS1!ADDREGISTROS)
    VENTANA.UtilizaFormula = XVALO(RS1!USAFORMULA)
    VENTANA.HabilitaBorrar = XVALO(RS1!BORRAR)
    VENTANA.HabilitaInsertar = XVALO(RS1!INSERTAR)
    VENTANA.NoMostrarBotonTilde = XVALO(RS1!NOTILDE)
    VENTANA.TITULO = SAFETEXT$(RS1!TITURECORD)
    
    CONSU$ = SAFETEXT$(RS1!CONSULTASQL)
    Check18.Value = XVALO(RS1!USARECORD)
    
    
    SQLX$ = "SELECT * FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA = " & IDLIST&
    SQLX$ = SQLX$ + " ORDER BY ENCAPRF_ID ASC"
    Set RS = READSET(SQLX$)
    With RS
    J% = 0
      LONGIT% = 0
      Do While Not .EOF
        J% = J% + 1
        ENCAX1$ = TRIM$(SAFETEXT$(!Encabezado))
        POS1% = InStr(ENCAX1$, "/")
        LONGIT% = LONGIT% + XVALO(Mid$(ENCAX1$, POS1% + 2))
        
        FORMUX1 = TRIM$(SAFETEXT$(!Formula))
        CAMPEDI1$ = TRIM$(SAFETEXT$(!CAMPEDITAB))
        COLECO1$ = TRIM$(SAFETEXT$(!COLCAMECO))
        TXTTOTX1$ = TRIM$(SAFETEXT$(!TXTTOTAL))
        TOTAX1$ = TRIM$(SAFETEXT$(!Total))
        DEJLIB1$ = TRIM$(SAFETEXT$(!DejarLibre))
        If ENCAX1$ <> "" Then ENCAX$ = ENCAX$ + ENCAX1$ + ";" 'ENCABEZADO
        If FORMUX1 <> "" Then VENTANA.Formula(J%) = FORMUX1
        If TRIM(UCase(CAMPEDI1$)) = "S" Then CAMPEDI$ = CAMPEDI$ & J% & ";" 'CAMPOS EDITABLES
        If COLECO1$ <> "" Then VENTANA.ColEcoCampo(J%) = COLECO1$
        If TXTTOTX1$ <> "" Then VENTANA.TxtUbicadoGrillaTotal(J%) = TXTTOTX1$
        If TOTAX1$ <> "" Then VENTANA.TOTALES(J%) = TOTAX1$
        If DEJLIB1$ <> "" Then VENTANA.DejarLibre(J%) = DEJLIB1$
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If ENCAX$ <> "" Then
      If Right(ENCAX$, 1) = ";" Then ENCAX$ = Left$(ENCAX$, Len(ENCAX$) - 1)
      VENTANA.Campox = ENCAX$
    End If
    '
    If CAMPEDI$ <> "" Then
      If Right(CAMPEDI$, 1) = ";" Then CAMPEDI$ = Left$(CAMPEDI$, Len(CAMPEDI$) - 1)
      VENTANA.CampEditables = CAMPEDI$
    End If
    '
    
    VENTANA.Inicializar = 1
    '
    RS1.Close
    Set RS1 = Nothing
    
    SQLX$ = CONSU$ + " WHERE ID_LISTA = " & IDLIST&
    Set RS1 = READSET(SQLX$)
    Call Carga_MSF_Recordset(RS1, VENTANA.Campox, VENTABNEW.MSF)
    RS1.Close
    Set RS1 = Nothing
    
'    If VENTABNEW.MSF.Rows < 2 Then VENTABNEW.MSF.Rows = 2
'    VENTABNEW.TEPRUEBA.Caption = String$(LONGIT%, "A")
'    VENTABNEW.MSF.Width = VENTABNEW.TEPRUEBA.Width
'    VENTABNEW.Width = VENTABNEW.TEPRUEBA.Width + 2500
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        Call GRABASQLDAT(SQLX$, RET_CAMPOS$, RET_TABLA$, RET_CONDI$)
        'MARCO CON MARBORRA
        Call ACTUREGISTRO(RET_TABLA$, "MARBORRA", RET_CONDI$, 1, REGAF&)
        Call TEXTOLOTES$(RET_CAMPOS$, ",")
        On Error Resume Next
        CANTICAMP = UBound(CADENATEX$)
        
        On Error GoTo 0
        If CANTICAMP > 0 Then
          For i& = 1 To VENTABNEW.MSF.Rows - 1
             If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)) = "" Then GoTo 60
             VALOX$ = IDLIST& & ","
             VALOX$ = VALOX$ & Chr(39) & LISTAX$ & Chr(39) & ","
             For k& = 1 To VENTABNEW.MSF.Cols - 1
                DATOX$ = Chr(39) & VENTABNEW.MSF.TextMatrix(i&, k&) & Chr(39)
                VALOX$ = VALOX$ + DATOX$ + ","
             Next k&
             If Right$(VALOX$, 1) = "," Then VALOX$ = Left$(VALOX$, Len(VALOX$) - 1)
             SQLX$ = "INSERT INTO " + RET_TABLA$
               PRICAMP$ = "ID_LISTA, LISTA, "
             SQLX$ = SQLX$ + " (" + PRICAMP$ + RET_CAMPOS$ + ") "
             SQLX$ = SQLX$ + " VALUES (" & VALOX$ & ")"
             FLEXCONN.Execute SQLX$
60        Next i&
        End If
        Call BORRAREGISTROS(RET_TABLA$, RET_CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    End If
    Set VENTANA = Nothing
  
99  Command22.Enabled = True

End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   IDMAX& = 1 + MAXNUMCOD&("ENCAPRFUN", "ID_LISTA")
   Me.Text20 = ""
   Me.TXTIDLISTA = IDMAX&
   ESTADO$ = "NEW"
   Command23.Enabled = True
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Text9 = ""
    If TRIM$(TXTCONSU_ESPEJO) = "" Then
    Call COMUNI("FALTA CONDICION PARA  MARBORRA\LA MISMA VA A SER USADA PARA EL ADDNEW")
      GoTo 99
    End If
    If TRIM$(TXTNAMETABLA) = "" Then
      Call COMUNI("FALTA SELECCIONAR TABLA")
      GoTo 99
    End If
    '
'    For i& = 1 To MSFUPD.Cols - 1
'        CAMPO1$ = TRIM$(MSFUPD.TextMatrix(1, i&))
'        For J& = 0 To LSTMOSTARCMP2.ListCount
'            LSTMOSTARCMP2.ListIndex = J&
'            CAMPOLEIDO$ = LSTMOSTARCMP2.TEXT
'            If CAMPO1$ = CAMPOLEIDO$ Then
'             LSTMOSTARCMP2.Selected(J&) = True
'             Exit For
'            End If
'        Next J&
'    Next i&
    '
T1$ = "DIM MSF AS MSFLEXGRID" & vbCrLf
T1$ = T1$ + "SET MSF = MSF11" & vbCrLf & vbCrLf
T1$ = T1$ + "CONDI$ = " & Chr(34) & TXTCONSU_ESPEJO & Chr(34) & vbCrLf
T1$ = T1$ + "Call ACTUREGISTRO(" & Chr(34) & TXTNAMETABLA.TEXT & Chr(34)

T1$ = T1$ + "," & Chr(34) & "MARBORRA" & Chr(34)
T1$ = T1$ + ", CONDI$ "
T1$ = T1$ + ", 1"
T1$ = T1$ + ", REGAF "
T1$ = T1$ + "," & Chr(34) & "NOMESS" & Chr(34) & ")" & vbCrLf & vbCrLf

T1$ = T1$ + "SQLX$ = " & Chr(34) & "SELECT * FROM " & TXTNAMETABLA & Chr(34) & " " & vbCrLf
T1$ = T1$ + "SQLX$ = SQLX$ + " & Chr(34) & " WHERE " & Chr(34) & vbCrLf
T1$ = T1$ + "SQLX$ = SQLX$ + " & Chr(34) & " " & TXTCONSU_ESPEJO & Chr(34) & vbCrLf
T1$ = T1$ + "Dim RS As ADODB.Recordset" & vbCrLf
T1$ = T1$ + "Set RS = New ADODB.Recordset" & vbCrLf
T1$ = T1$ + "25      On Error Resume Next" & vbCrLf
T1$ = T1$ + "RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText" & vbCrLf
T1$ = T1$ + "If Err Then" & vbCrLf
T1$ = T1$ + " On Error GoTo 0" & vbCrLf
T1$ = T1$ + " Call CapturaErrorOpen" & vbCrLf
T1$ = T1$ + " GoTo 25" & vbCrLf
T1$ = T1$ + "End If" & vbCrLf

T1$ = T1$ + "  With RS" & vbCrLf
T1$ = T1$ + "      FOR REG& = 1 TO MSF.ROWS - 1" & vbCrLf
T1$ = T1$ + "       .AddNew" & vbCrLf
For i& = 1 To MSFUPD.Cols - 1
       FORMAT1$ = MSFUPD.TextMatrix(0, i&)
       Call TEXTOLOTES$(FORMAT1$, "/")
       On Error Resume Next
       FORMAT1$ = UCase(TRIM$(CADENATEX$(2)))
       If Left(FORMAT1$, 1) = "F" Then
          T1$ = T1$ & "          !" & MSFUPD.TextMatrix(1, i&) & "=" & "XVALO(MSF.TextMatrix(REG&, " & i& & "))" & vbCrLf
       Else
          T1$ = T1$ & "          !" & MSFUPD.TextMatrix(0, i&) & "=" & "MSF.TextMatrix(REG&, " & i& & ")" & vbCrLf
       End If
    Next i&
T1$ = T1$ + "       .UPDATE" & vbCrLf
T1$ = T1$ + "      NEXT REG&" & vbCrLf
T1$ = T1$ + "  END WITH" & vbCrLf & vbCrLf
T1$ = T1$ + "  RS.CLOSE" & vbCrLf
T1$ = T1$ + "  SET RS= NOTHING" & vbCrLf & vbCrLf

T1$ = T1$ + "CONDI$ =  CONDI$ + " & Chr(34) & " AND  MARBORRA = 1  " & Chr(34) & vbCrLf
T1$ = T1$ + "Call BORRAREGISTROS(" & Chr(34) & TXTNAMETABLA.TEXT & Chr(34) & ", CONDI$," & "REGAF&," & Chr(34) & "NOMESS" & Chr(34) & ")"


Text9 = T1$
            
99   Command24.Enabled = True
End Sub

Private Sub Command24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       Call AYUDA("GEN.COD.Update")
     End If
   End If

End Sub

Sub AYUDA(DATO$)
   Select Case DATO$
     Case "GEN.COD.Update"
       TXTO$ = "GENERA CODIGO PARA GRABAR DESDE UNA GRILLA " & vbCrLf
       TXTO$ = TXTO$ + "A LA TABLA CORRESPONDIENTE  " + vbCrLf
       TXTO$ = TXTO$ + "1.-ESCRIIR EN EL CASILLERO CAMPOS (TXTCAMPOSUPDATE)" + vbCrLf
       TXTO$ = TXTO$ + "LOS CAMPOS QUE ESTAN EN LAS PROPIEDADES DE LA GRILLA PARA CARGAR ENCBEZADO" + vbCrLf

       TXTO$ = TXTO$ + "2.-CLICK EN EL BOTON ADD CAMPOS" + vbCrLf
       TXTO$ = TXTO$ + "3.-CLICK EN CADA COLUMNA DE LA GRILLA PARA SELECCIONAR LOS CAMPOS DE LA TABLA A ENLAZAR" + vbCrLf
       TXTO$ = TXTO$ + "4.-VER EL SEPARADOR DE CAMPOS SI ES EL CORRECTO" + vbCrLf
       TXTO$ = TXTO$ + "5.-CLICK EN BOTON GEN. COD. UPDATE" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
       
     Case "Crear Archivo Estrucutura"
       TXTO$ = "GENERA ESTRUCTURA DE ARCHIVO TIPO PRF Y RFS" & vbCrLf
       TXTO$ = TXTO$ + "INGRESAR EN LA SOLAPA TAB2  " + vbCrLf
       TXTO$ = TXTO$ + "0.-CLIC EN NEW - PONERLE UN NOMBRE Y UNA DESCRIPCION A LA ESTRUCTURA" + vbCrLf
       TXTO$ = TXTO$ + "1.-INGRESAR EL ENCABEZADO Y FORMATO DE LOS CAMPOS)" + vbCrLf
       TXTO$ = TXTO$ + "2.-AGREGAR LOS CAMPOS A MOSTRAR  Y LA TABLA DATASQL " + vbCrLf
       TXTO$ = TXTO$ + "3.-EJ.: SELECT CAMPO1, CAMPO2,.... FROM DATASQL WITH(NOLOCK)" + vbCrLf
       TXTO$ = TXTO$ + " -ESTO EN EL CASILLERO ROTULADO CONSULTA PARA RECORDSET" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "4.-CAMPOS ECO" + vbCrLf
       TXTO$ = TXTO$ + " -Los campos Eco deben estar editables." + vbCrLf
       TXTO$ = TXTO$ + " -SE PUEDE INVOCAR A UNA TABLA SQL O A UNA ESTRUCTURA" + vbCrLf
       TXTO$ = TXTO$ + " -La referencia del eco debe ir donde va a ir el codigo en esa " + vbCrLf
       TXTO$ = TXTO$ + " -columna se va agregar un boton automatico, y en la columna " + vbCrLf
       TXTO$ = TXTO$ + " -siguiente va a ir la descripcion" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + " -en el caso de invocar como eco a una estructura se debe" + vbCrLf
       TXTO$ = TXTO$ + " -poner el siguiente codigo " + vbCrLf
       TXTO$ = TXTO$ + " -CAMPO1;CAMPO2;DATASQL;LISTA = 'CODIMOTI'" + vbCrLf
       TXTO$ = TXTO$ + " -con el mismo criterio que en el paso anterior." + vbCrLf
       TXTO$ = TXTO$ + " -poner el siguiente codigo " + vbCrLf
       MsgBox TXTO$
       
     Case "ATAJOS EN VENTABNEW (VENTANA DE CARGA)"
       TXTO$ = "ATAJOS EN VENTABNEW (VENTANA DE CARGA)" & vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI SE OPRIME F2 EN UN CAJA DE TEXTO (CELDA EDITABLE CON ECO) " + vbCrLf
       TXTO$ = TXTO$ + "     DESPLEGA LA LISTA DE SELECCION " + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI SE OPRIME F3 SOBRE UNA CELDA REPITE EL VALOR INMEDIATO SUPERIOR" + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI LA CELDA ES EDITABLE Y TIENE FORMATO D8 " + vbCrLf
       TXTO$ = TXTO$ + "   CUANDO ESTA LA CAJA DE TEXTO VISIBLE SI SE REALIZA DOBLE CLIC " + vbCrLf
       TXTO$ = TXTO$ + "   MUESTRA LA LISTA DE SELECCION DE FECHA" + vbCrLf
       MsgBox TXTO$
       
     Case "Celda Resultado de Función"
       TXTO$ = "PARA CONFIGURAR QUE UNA CELDA SEA RESULTADO DE UNA FUNCION EXISTENTE:" & vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "Ejemplo: " + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "VENTANA.formula(7) = ImporteConDescuento#;5;6" + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "DONDE 7 ES EL NUMERO DE COLULMNA DEL RESULTADO" + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "DONDE ImporteConDescuento# ES LA DENOMINACION DE LA FUNCION" + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "DONDE 5;6 SON LOS PARAMETROS QUE NECESITA LA FUNCION " + vbCrLf
       MsgBox TXTO$
       
   End Select
End Sub
Private Sub Command25_Click()
   Command25.Enabled = False
    
   IDLISTAX& = XVALO(TXTIDLISTA)
   If IDLISTAX& < 1 Then
     Call COMUNI("Falta Seleccionar Id Lista")
     GoTo 99
   End If
   '
   If XVALO(Text11) < 1 Then
      Call COMUNI("Falta Ingresar Cantidad de Campos")
      Call PINTATEXT(Text11)
      GoTo 99
   End If
   
   If ESTADO$ = "NEW" Then
      IDLISTAX& = 1 + MAXNUMCOD&("ENCAPRFUN", "ID_LISTA")
   End If
   '
   CODICON$ = TRIM$(Text20)
   If CODICON$ = "" Then
      Call COMUNI("Falta Ingresar Codigo Alfa de Consulta")
      Text20.Locked = False
      Text20.BackColor = vbWhite
      Call PINTATEXT(Text20)
      GoTo 99
   End If
   
   REFECON$ = TRIM$(Text14)
   If REFECON$ = "" Then
      Call COMUNI("Falta Ingresar Referencia de Consulta")
      Text14.Locked = False
      Text14.BackColor = vbWhite
      Call PINTATEXT(Text14)
      GoTo 99
   End If
   '
   For i& = 0 To MSF4.Cols - 1
       ENK$ = TRIM$(MSF4.TextMatrix(0, i&))
       If ENK$ = "" Then
          CANTIENCA% = i&
          Text11 = i&
          MSF4.Cols = i&
          Call COMUNI("Atención !!! Se ha Ajustado la Cantidad de Columnas Por Que se Detectaron Columnas Vacias")
          Exit For
       End If
   Next i&
   '
   
   If TRIM(Text13) = "" Then
     Call COMUNI("FALTA CONDICION DE TABLA")
     Call PINTATEXT(Text13)
     GoTo 99
   End If
   
   'Left(CAMPEDI$, 1) < 2
   If MSF4.TextMatrix(2, 1) <> "S" Then
       Call COMUNI("ATENCIÓN CUANDO EL CAMPO 1 NO ESTA EDITABLE EL SISTEMA AUTONUMERA\ESTE CAMPO PONERLO COMO NUMERICO F...\CONFIGURAR QUE NO SE PUEDA BORRAR NI INSERTAR\PARA QUE NO SE AUTONMERE Y CAMBIE LOS CODIGOS ORIGINALES")
   End If
   SQLX$ = "SELECT * FROM  ENCAPRFUN  "
   SQLX$ = SQLX$ + " WHERE ID_LISTA = " & IDLISTAX&
   SQLX$ = SQLX$ + " AND (MARBORRA = 0 OR MARBORRA IS NULL)"
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
        If .EOF Then
          .AddNew
        End If
        !ID_LISTA = IDLISTAX&
        !LISTA = Left$(Text20, 64)
        !Descripcion = Left$(Text14, 64)
        !ADDREGISTROS = Check13.Value
        !USAFORMULA = Check14.Value
        !BORRAR = Check15.Value
        !INSERTAR = Check16.Value
        !NOTILDE = Check17.Value
        !CANTICAMPOS = XVALO(Text11)
        !CodiEmpr = CodiEmp%
        !USARECORD = Check18.Value
        !FECHEMISI = FETEXCOR1$(HOY(HO$))
        !TITURECORD = Text12
        !CONSULTASQL = Text13
        .Update
   End With
   RS.Close
   Set RS = Nothing
   Call ACTUREGISTRO("ENCAPRF", "MARBORRA", "ID_LISTA = " & IDLISTAX&, 1, REGAF&, "NOMESS")
   '
   SQLX$ = "SELECT * FROM  ENCAPRF  "
   SQLX$ = SQLX$ + " WHERE ENCAPRF_ID = " & IDLISTAX&
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS1
    For i& = 1 To MSF4.Cols - 1
       .AddNew
       !ID_LISTA = IDLISTAX&
       !CodiEmpr = CodiEmp%
       !MARBORRA = 0
       !Encabezado = MSF4.TextMatrix(0, i&)
       !Formula = MSF4.TextMatrix(1, i&)
       !CAMPEDITAB = MSF4.TextMatrix(2, i&)
       !COLCAMECO = MSF4.TextMatrix(3, i&)
       !TXTTOTAL = MSF4.TextMatrix(4, i&)
       !Total = MSF4.TextMatrix(5, i&)
       !DejarLibre = MSF4.TextMatrix(6, i&)
    Next i&
    .Update
             
   End With
   RS1.Close
   Set RS1 = Nothing
   Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX& & " AND MARBORRA = 1", REGAF&, "NOMESS")
   ESTADO$ = "EDIT"
   '

    Call GENERACODIGO_ARCHIVO_ESTRUCTURA(Text20)

99 Command25.Enabled = True

End Sub
Sub GENERACODIGO_ARCHIVO_ESTRUCTURA(CODICONSULTA$)
   '
   'VALIDO SI EXISTE EN ALGUNA TABLA DE CUALQUIER SISTEMA
   EXISTE_EN_LISTA% = 0
   If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
     ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + "CONTROLESTSQL.txt"
     NARCHI2% = FILEOPEN%(ASALI$, 1, 128) ' ABRE EL ARCHIVO DE ESCRITURA
     Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(Mid$(XX$, 1, 24)))
       
       If XX1$ = TRIM$(UCase$(Text20)) Then
          Descripcion$ = TRIM$(UCase$(Mid$(XX$, 26, 32)))
          EMPRESA$ = TRIM$(UCase$(Mid$(XX$, 60, 24)))
          TXT1$ = "Estructura: " & Text20 & " Referencia: " & Descripcion$ & " Empresa: " & EMPRESA$ & "\Anteriormente Creada"
          OPX% = COMALTER%(TXT1$ & "\\Cancelar\Generar Código Para Escribir en Módulo EscribeEstrucSql.bas")
          EXISTE_EN_LISTA% = 1
          If OPX% < 2 Then Exit Sub
       End If
    Loop
    '
    If EXISTE_EN_LISTA% < 1 Then
          Call REPLA(TX$, TRIM$(UCase$(Text20)), 1, 24)
          Call REPLA(TX$, TRIM$(UCase$(Text14)), 26, 32)
          Call REPLA(TX$, TRIM$(UCase$(Left$(EMPREAC$, 24))), 60, 24)
          Call REPLA(TX$, TRIM$(UCase$(Left$(USERTER$, 24))), 86, 24)
          
          Close #NARCHI2%
          NARCHI3% = FILEOPEN%(ASALI$, 3, 128) 'ABRE PARA ESCRIBIR AL FINAL
          Print #NARCHI3%, , TX$
          Close #NARCHI3%
    End If
    '
    On Error Resume Next
    Close #NARCHI2%
    On Error GoTo 0
   End If
   
   Text15 = "'|INICIO " & CODICONSULTA$ & vbCrLf
   Text15 = Text15 + "'" & vbCrLf
   Text15 = Text15 + "CONDI$ = " & Chr$(34) & " LISTA =   '" & CODICONSULTA$ & "'" & vbCrLf
   Text15 = Text15 + "EXILISTA% = CantRegistros&("
   Text15 = Text15 & Chr$(34) & "ENCAPRFUN" & Chr$(34)
   Text15 = Text15 + ", CONDI$ ," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "If EXILISTA% > 0 Then" & vbCrLf
   Text15 = Text15 + "  FE% = " & HOY(HO$) & vbCrLf
   Text15 = Text15 + "  FEBD% = CVINT(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "FECHEMISI" & Chr$(34) & ", CONDI$, " & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "  GENERAREST% = 0" & vbCrLf
   Text15 = Text15 + "  If FEBD% < FE% Then" & vbCrLf
   Text15 = Text15 + "    IDLISTAX& = XVALO(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "ID_LISTA" & Chr$(34) & ", CONDI$," & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "    Call BORRAREGISTROS(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "ID_LISTA = " & Chr$(34) & " & IDLISTAX&,REGAF&," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "    Call BORRAREGISTROS(" & Chr$(34) & "ENCAPRF" & Chr$(34) & "," & Chr$(34) & "ID_LISTA = " & Chr$(34) & " & IDLISTAX&,REGAF&," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "    GENERAREST% = 1" & vbCrLf
   Text15 = Text15 + "  END IF  " & vbCrLf
   Text15 = Text15 + "ELSE" & vbCrLf
   Text15 = Text15 + "  GENERAREST% = 1" & vbCrLf
   Text15 = Text15 + "  CONDI$ =  " & Chr$(34) & " ID_LISTA  > 0 " & Chr$(34) & vbCrLf
   Text15 = Text15 + "  IDLISTAX& = 1 + XVALO(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "MAX(ID_LISTA)" & Chr$(34) & ",CONDI$," & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "END IF  " & vbCrLf
   Text15 = Text15 + "If GENERAREST% > 0 Then  " & vbCrLf
   
   '***********************************************************************************************************************************

   Text15 = Text15 + "  SQLX$ = " & Chr$(34) & "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE " & Chr$(34) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)" & Chr$(34) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "VALUES ( " & Chr$(34) & " & IDLISTAX&" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ",'" & Left$(Text20, 64) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ",'" & Left$(Text14, 64) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "," & Check13.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "," & Check14.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check15.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check16.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check17.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & XVALO(Text11) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & CodiEmp% & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check18.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & CVDAT(HOY(HO$)) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & Text12 & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & Text13 & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "  FLEXCONN.Execute (SQLX$)" & vbCrLf
      '
   For i& = 1 To MSF4.Cols - 1
        Text15 = Text15 + "    SQLX$ = " & Chr$(34) & "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) " & Chr$(34) & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " VALUES ( " & Chr$(34) & " & IDLISTAX&" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " ," & CodiEmp% & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " ,0" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(0, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(1, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(2, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(3, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(4, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(5, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(6, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " )" & vbCrLf
        Text15 = Text15 + "    FLEXCONN.Execute (SQLX$)" & vbCrLf & vbCrLf
        Text15 = Text15 + Chr$(39) & vbCrLf
   Next i&
   Text15 = Text15 + "End If " & vbCrLf
   Text15 = Text15 + "'|FIN " & CODICONSULTA$ & vbCrLf
   '

   OPX% = COMALTER%("Desea Escribir el Modulo 'EstrucSQL'\\No, No Escribir\Si, Escribir (Cierre Los Fuentes)")
   If OPX% = 2 Then
      Call EscribeEstrucSql(CODICONSULTA$, Text15)
   End If
   
   
End Sub
   '
Private Sub Command26_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  
  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  NRO = XVALO(RESP_ZELE_VECT(1))
  
  If NRO > 0 Then
    ESTADO$ = "EDIT"
    TXTIDLISTA = NRO
    Text20 = RESP_ZELE_VECT(2)
  End If
End Sub

Private Sub Command27_Click()
   IDLISTAX& = XVALO(Me.TXTIDLISTA)
   If IDLISTAX& > 0 Then
     OPX% = COMALTER%("Confirma Eliminar Definitivamente esta Estructura\\Cancelar\Eliminar")
     If OPX% = 2 Then
        CONDI$ = " ID_LISTA = " & IDLISTAX&
        Call BORRAREGISTROS("ENCAPRFUN", CONDI$, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", CONDI$, REGAF&, "NOMESS")
     End If
   End If
End Sub

Private Sub Command28_Click()
   For U& = 1 To LSTMOSTARCMP.ListCount
     If LSTMOSTARCMP.Selected(U& - 1) = True Then
         LSTMOSTARCMP.Selected(U& - 1) = False
       Else
         LSTMOSTARCMP.Selected(U& - 1) = True
     End If
   Next U&
End Sub

Private Sub Command29_Click()
   Text17.Width = Text17.Width - 100
End Sub

Private Sub Command30_Click()
Text16.Width = Text16.Width - 100
End Sub

Private Sub Command31_Click()
Text16.Width = Text16.Width + 100
End Sub



Private Sub Command32_Click()
        For J& = 0 To List5.ListCount - 1
            List5.ListIndex = J&
            If List5.Selected(J&) = False Then
               List5.Selected(J&) = True
            Else
               List5.Selected(J&) = False
            End If
        Next J&

End Sub

Private Sub Command33_Click()
   NUM = SELFRM
   If NUM > 0 Then
     TXTCODIGO = NUM
     LBLNUEVO.Visible = False
   End If

End Sub
 Function SELFRM() As Long
    Erase RESP_ZELE_VECT
    Campos$ = "Código/F9;Formulario/a14;Referencia/a48;ID/a0"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "codigo,FRMNAME,REFERENCIA " 'OJO TIENE GROUP BY
    TABLA$ = "CREACOD"
    CONDI$ = "CREACOD_ID > 0  "
    '
    CONDI$ = CONDI$ + "GROUP BY  codigo,FRMNAME,REFERENCIA ORDER BY FRMNAME"
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then ' 7 ES EL ID QUE SI SELECCIONO ES > 0 POR QUE ES LA CONDICION
       SELFRM = XVALO(RESP_ZELE_VECT(1))
    End If


End Function



Private Sub Command34_Click()
   TXTCODIGO = NumNuevaCreacod
   Me.LBLNUEVO.Visible = True
End Sub

Private Sub Command35_Click()
Command35.Enabled = False
'********   Form1.Designer.cs    *******

'           private void InitializeComponent ()
'            {
'            this.components = new System.ComponentModel.Container();
'            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
'            this.Text = "Form4";
            
'            this.textBox1 = new System.Windows.Forms.TextBox();
'            this.textBox2 = new System.Windows.Forms.TextBox();
'            this.textBox3 = new System.Windows.Forms.TextBox();
'            this.label1 = new System.Windows.Forms.Label();
'            this.label2 = new System.Windows.Forms.Label();
'            this.label3 = new System.Windows.Forms.Label();
'            this.SuspendLayout();
'            //
'            // textBox1
'            //
'            this.textBox1.Location = new System.Drawing.Point(133, 50);
'            this.textBox1.Name = "textBox1";
'            this.textBox1.Size = new System.Drawing.Size(105, 20);
'            this.textBox1.TabIndex = 0;
'            //
'            // textBox2
'            //
'            this.textBox2.Location = new System.Drawing.Point(133, 76);
'            this.textBox2.Name = "textBox2";
'            this.textBox2.Size = new System.Drawing.Size(105, 20);
'            this.textBox2.TabIndex = 1;
'            //
'            // textBox3
'            //
'            this.textBox3.Location = new System.Drawing.Point(133, 102);
'            this.textBox3.Name = "textBox3";
'            this.textBox3.Size = new System.Drawing.Size(105, 20);
'            this.textBox3.TabIndex = 2;
'            //
'            // button1
'            //
'            this.button1.Location = new System.Drawing.Point(239, 60);
'            this.button1.Name = "button1";
'            this.button1.Size = new System.Drawing.Size(20, 20);
'            this.button1.TabIndex = 6;
'            this.button1.Text = "button1";
'            this.button1.UseVisualStyleBackColor = true;
'            //
'            // label1
'            //
'            this.label1.AutoSize = true;
'            this.label1.Location = new System.Drawing.Point(27, 53);
'            this.label1.Name = "label1";
'            this.label1.Size = new System.Drawing.Size(35, 13);
'            this.label1.TabIndex = 3;
'            this.label1.Text = "label1";
'            //
'            // label2
'            //
'            this.label2.AutoSize = true;
'            this.label2.Location = new System.Drawing.Point(27, 79);
'            this.label2.Name = "label2";
'            this.label2.Size = new System.Drawing.Size(35, 13);
'            this.label2.TabIndex = 4;
'            this.label2.Text = "label2";
'            //
'            // label3
'            //
'            this.label3.AutoSize = true;
'            this.label3.Location = new System.Drawing.Point(27, 105);
'            this.label3.Name = "label3";
'            this.label3.Size = new System.Drawing.Size(35, 13);
'            this.label3.TabIndex = 5;
'            this.label3.Text = "label3";
'            //
'            // Form2
'            //
'''''            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
'''''            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
'''''            this.ClientSize = new System.Drawing.Size(930, 418);
'            this.Controls.Add(this.label3);
'            this.Controls.Add(this.label2);
'            this.Controls.Add(this.label1);
'            this.Controls.Add(this.textBox1);
'            this.Controls.Add(this.textBox2);
'            this.Controls.Add(this.textBox3);

'            this.Name = "Form2";
'            this.Text = "Form2";
'            this.ResumeLayout(false);
'            this.PerformLayout();
'
'            }
'
'        #endregion
'
'        private System.Windows.Forms.TextBox textBox1;
'        private System.Windows.Forms.TextBox textBox2;
'        private System.Windows.Forms.TextBox textBox3;
'        private System.Windows.Forms.Label label1;
'        private System.Windows.Forms.Label label2;
'        private System.Windows.Forms.Label label3;
'        this.Controls.Add(this.button1);
'        }
'   TEXTOPARRAFOPANEL$ = PARRAFOPANEL$      'this.panel1 = new System.Windows.Forms.Panel();     (1)
'   TEXTOLAYOUTPANEL$ = PARRAFOLAYOUTPANEL$ 'this.panel1.SuspendLayout();                        (5)
   INDICEX = 0:
   Erase VectorControlAdd
   Erase VectorSystem_Windows_Forms
   Text21.TEXT = ""
   Text21.TEXT = "//Ingresar todo debajo de : 'public partial class Form1 : Form'" & vbCrLf
   Text21.TEXT = Text21.TEXT & "Form f = new Form(); //FORMULARIO CONTENEDEOR DE CONTROL DE USUARIO XZELECHO" & vbCrLf
   Text21.TEXT = Text21.TEXT & "public Form1 () //Suponiendo que el formulario se llame Form1" & vbCrLf
   Text21.TEXT = Text21.TEXT & "//using HerrNumericas; // ****  agregar la referencia de dll en capa vista ****" & vbCrLf
   Text21.TEXT = Text21.TEXT & vbCrLf & "{" & vbCrLf
   
   Text21.TEXT = Text21.TEXT & "InitializeComponent();" & vbCrLf
   Text21.TEXT = Text21.TEXT & vbCrLf & "}" & vbCrLf
   
   
   TX23$ = "private void InitializeComponent ()"
   TX23$ = TX23$ & vbCrLf & "{"
   TX23$ = TX23$ & vbCrLf & "this.components = new System.ComponentModel.Container();"
   TX23$ = TX23$ & vbCrLf & "this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;"
   TX23$ = TX23$ & vbCrLf & "this.Text = " & Chr(34) & TXTNAMEFRM & Chr(34) & ";"
   
   TX23$ = TX23$ & vbCrLf & PARRAFOROTULO$           'this.lblcuenta = new System.Windows.Forms.Label();  (2)
   TX23$ = TX23$ & vbCrLf & PARRAFOBUTTON$ '(OJO-NO REPETIR NOMBRE DE CAMPO'this.button1 = new System.Windows.Forms.Button();   (3)
   TX23$ = TX23$ & vbCrLf & PARRAFOTEXTBOX$
   TX23$ = TX23$ & vbCrLf & "this.SuspendLayout();"
   TX23$ = TX23$ & vbCrLf & PARRAFOROTULO_2$
   TX23$ = TX23$ & vbCrLf & PARRAFOTEXTBOX_2$
   TX23$ = TX23$ & vbCrLf & PARRAFOBUTTON_2$
   '
   TX23$ = TX23$ & vbCrLf & "this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);"
   TX23$ = TX23$ & vbCrLf & "this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;"
   TX23$ = TX23$ & vbCrLf & "this.ClientSize = new System.Drawing.Size(930, 530);"
            
   TX23$ = TX23$ & vbCrLf & CargaVectorControlAdd$
   TX23$ = TX23$ & vbCrLf & "this.Name = " & Chr(34) & TXTNAMEFRM & Chr(34) & ";"
   TX23$ = TX23$ & vbCrLf & "this.Text =  " & Chr(34) & TXTNAMEFRM & Chr(34) & ";"
   TX23$ = TX23$ & vbCrLf & "this.ResumeLayout(false);"
   TX23$ = TX23$ & vbCrLf & "this.PerformLayout();"
   TX23$ = TX23$ & vbCrLf & "}"
   
   TX23$ = TX23$ & vbCrLf & "#endregion"
   TX23$ = TX23$ & vbCrLf & CargaVectorSystem_Windows_Forms$
   'ESTO ES PARA NO EXCEDER LA CAPACIDAD DE TEXTBOX
   lineas = Split(TX23$, vbNewLine)
   Me.Text23.TEXT = ""
   If UBound(lineas) > 1000 Then
      For i = 0 To UBound(lineas)
         TL$ = lineas(i)
         If i <= 1000 Then
            Text23.TEXT = Text23.TEXT & TL$ & vbCrLf
         Else
           If i = 1001 Then
            Text23.TEXT = Text23.TEXT & vbCrLf & "======= *********  TIENE ANEXO   **********  =========="
            Text26.TEXT = ""
           End If
           Text26.TEXT = Text26.TEXT & TL$ & vbCrLf
         End If
      Next i
   End If

   Call CargarVectorBlanformu
   Call CargarVectorVectorSoloNumeros
   Call ArmaCodigodeAsingacionaCajasdeTexto
   SSTab2.Tab = 0
   Command35.Enabled = True
'   TEXTOLAYOUTPANEL$ = "}"
'FIN PRIMER PARRAFO
End Sub
Function PARRAFOTEXTBOX$()
'   this.textBox1 = new System.Windows.Forms.TextBox();
   CmpoIdRepli$ = CampoIdReplica$
   For i& = 1 To MSF3.Rows - 1
      ROTULO$ = TRIM$(MSF3.TextMatrix(i&, 4))
      If ROTULO$ <> "" Then
         CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
         If CampoName$ = CmpoIdRepli$ Then GoTo 30 'saltea la caja de texto de caja de texto  campoidreplikca
         If PARRAFOTEXTBOX$ = "" Then
            PARRAFOTEXTBOX$ = "this.TXT" & Text18(0) & "_" & CampoName$ & "= new System.Windows.Forms.TextBox();"
         Else
            PARRAFOTEXTBOX$ = PARRAFOTEXTBOX$ & vbCrLf & "this.TXT" & Text18(0) & "_" & CampoName$ & "= new System.Windows.Forms.TextBox();"
         End If
      End If
30 Next i&

End Function
Function PARRAFOROTULO$()
   'this.lblcuenta = new System.Windows.Forms.Label();
   For i& = 1 To MSF3.Rows - 1
      ROTULO$ = TRIM$(MSF3.TextMatrix(i&, 4))
      CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
      If ROTULO$ <> "" Then
         If PARRAFOROTULO$ = "" Then
            PARRAFOROTULO$ = "this.LBL" & Text18(0) & "_" & CampoName$ & "= new System.Windows.Forms.Label();"
         Else
            PARRAFOROTULO$ = PARRAFOROTULO$ & vbCrLf & "this.LBL" & Text18(0) & "_" & CampoName$ & "= new System.Windows.Forms.Label();"
         End If
      End If
   Next i&
End Function
Function AsignarAEntidadCajasDeTexto()
   'this.lblcuenta = new System.Windows.Forms.Label();
   TABLA$ = Text18(0)
   CmpoIdRepli$ = CampoIdReplica$
'   Erase VectorAsignarCajasTextoaEntidad
   TX12$ = ""
   For i& = 1 To MSF3.Rows - 1
      ROTULO$ = TRIM$(MSF3.TextMatrix(i&, 4))
      CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
      If CmpoIdRepli$ <> CampoName$ Then 'SI ES EL CAMPO IDREPLICA NO SE AGREGA
        TD$ = tipodatoSql$(TABLA$, CampoName$)
        If ROTULO$ <> "" Then
           TX12$ = "entidad" & TABLA$ & "." & CampoName$ & " = Convert.To" & TD$ & "(" & "TXT" & Text18(0) & "_" & CampoName$ & ".Text);" & vbCrLf
           ReDim Preserve VectorAsignarCajasTextoaEntidad(i&)
           VectorAsignarCajasTextoaEntidad(i&) = TX12$
        End If
      End If
   Next i&
End Function
Function PARRAFOROTULO_2$()

'            //
'            // label1
'            //
'            this.label1.AutoSize = true;
'            this.label1.Location = new System.Drawing.Point(27, 53);
'            this.label1.Name = "label1";
'            this.label1.Size = new System.Drawing.Size(35, 13);
'            this.label1.TabIndex = 3;
'            this.label1.Text = "label1";
'  ver el tema de los espacios del rotulo para la asignacion del nombre, modificar por guion bajo y limitar largo
   PARRAFOROTULO_2$ = ""
   LEFTX = 30
   TOPX = 20
   For i& = 1 To MSF3.Rows - 1
      CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
      ROTULO$ = TRIM$(MSF3.TextMatrix(i&, 4))
      If ROTULO$ <> "" Then
            INDICEX% = INDICEX% + 1
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "//" & vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "// LBL" & Text18(0) & "_" & CampoName$ & vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "//" & vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".AutoSize = true;" & vbCrLf
           ' this.LBLDocumentoOrigen.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".Location = new System.Drawing.Point(" & LEFTX & "," & TOPX & ");" & vbCrLf
            'this.label1.Name = "label1";
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".Name = " & Chr$(34) & "LBL" & Text18(0) & "_" & CampoName$ & Chr$(34) & ";" & vbCrLf
            'this.label1.Size = new System.Drawing.Size(35, 13);
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".Size = new System.Drawing.Size(60, 15);" & vbCrLf  'este dato probar de dejar fijo
            'this.label1.TabIndex = 3;& vbCrLf
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".TabIndex = " & INDICEX% & ";" & vbCrLf
            'this.label1.Text = "label1"
            PARRAFOROTULO_2$ = PARRAFOROTULO_2$ & "this.LBL" & Text18(0) & "_" & CampoName$ & ".Text = " & Chr$(34) & ROTULO$ & Chr$(34) & ";" & vbCrLf
            TOPX = TOPX + 20
            If TOPX > 700 Then
               TOPX = 80
               LEFTX = 230
            End If
            VectorControlAdd(INDICEX) = "this.Controls.Add(this.LBL" & Text18(0) & "_" & CampoName$ & ");"
            VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.Label LBL" & Text18(0) & "_" & CampoName$ & ";"
      End If
   Next i&
End Function
Function PARRAFOTEXTBOX_2$()

'            //
'            // textBox1
'            //
'            this.TXTDocumentoOrigen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
'            this.textBox1.Location = new System.Drawing.Point(133, 50);
'            this.textBox1.Name = "textBox1";
'            this.textBox1.Size = new System.Drawing.Size(105, 20);
'            this.textBox1.TabIndex = 0;

   PARRAFOTEXTBOX_2$ = ""
   LEFTX = 133
   TOPX = 20
   TABLA$ = Text18(0)
   
   For i& = 1 To MSF3.Rows - 1
      CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
      LENX = XVALO(MSF3.TextMatrix(i&, 3))
      ROTULO$ = TRIM$(MSF3.TextMatrix(i&, 4))
      TIDATO$ = UCase$(Left$(TRIM$(MSF3.TextMatrix(i&, 2)), 5))
      ALINEA$ = "Left"
      If TIDATO$ <> "TEXTO" And TIDATO$ <> "FECHA" Then
        ALINEA$ = "Right"
      End If
      If ROTULO$ <> "" Then
            INDICEX% = INDICEX% + 1
            
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & vbCrLf
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "//" & vbCrLf
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "// TXT" & Text18(0) & "_" & CampoName$ & vbCrLf
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "//" & vbCrLf
            'this.TXTDocumentoOrigen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;" & vbCrLf
'            this.TXTDocumentoOrigen.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf
            'this.TXTDocumentoOrigen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".TextAlign = System.Windows.Forms.HorizontalAlignment." & ALINEA$ & ";" & vbCrLf

'            this.textBox1.Location = new System.Drawing.Point(133, 50);
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".Location = new System.Drawing.Point(" & LEFTX & "," & TOPX & ");" & vbCrLf
            'this.textBox2.MaxLength = 64;
            If TIDATO$ = "TEXTO" Then ' SOLO SI ES TEXTO AGREGA LONGITUD DE MaxLength
                TIDATOCOMPLETO$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 2)))
                If TIDATOCOMPLETO$ = "TEXTO" Then 'PARA QUE NO INGRESE CUANDO ES TEXTO LARGO (MEMO)
                  MAXLENGT = LENX
                  PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".MaxLength =" & LENX & ";" & vbCrLf
                ElseIf TIDATOCOMPLETO$ = "TEXTOLARGO" Then
                  'this.textBox2.Multiline = true;
                  PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".Multiline = true;" & vbCrLf
                  'this.textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                  PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".ScrollBars = System.Windows.Forms.ScrollBars.Vertical;" & vbCrLf
                End If
            End If

            'this.label1.Name = "label1";
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".Name = " & Chr$(34) & "TXT" & Text18(0) & "_" & CampoName$ & Chr$(34) & ";" & vbCrLf
            'this.label1.Size = new System.Drawing.Size(35, 13);
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".Size = new System.Drawing.Size(60, 15);" & vbCrLf  'este dato probar de dejar fijo
            'this.label1.TabIndex = 3;& vbCrLf
            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".TabIndex = " & INDICEX% & ";" & vbCrLf
'            'this.label1.Text = "label1"
'            PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this." & Text18(0)  & "_" & CampoName$ & ".Text = " & Chr$(34) & ROTULO$ & Chr$(34) & ";" & vbCrLf
            TOPX = TOPX + 20

            If TOPX > 700 Then
               TOPX = 80
               LEFTX = 333
            End If
            VectorControlAdd(INDICEX) = "this.Controls.Add(this.TXT" & Text18(0) & "_" & CampoName$ & ");"
            VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.TextBox TXT" & Text18(0) & "_" & CampoName$ & ";"
            VectorBlanformu(INDICEX%) = "this.TXT" & Text18(0) & "_" & CampoName$ & ".Text =" & Chr$(34) & Chr$(34) & ";" & vbCrLf
            TipoDatoCampoId$ = tipodatoSql$(TABLA$, CampoName$)
            '
            'CARGAR EN EL VECTOR DE LA RUTINA MUESTRADATOS
            If TipoDatoCampoId$ = "String" Then
             TTM$ = "(string.IsNullOrEmpty(entidad" & TABLA$ & "." & CampoName$ & ")) ? " & Chr$(34) & Chr$(34) & ": entidad" & TABLA$ & "." & CampoName$ & ".ToString();" & vbCrLf
            ElseIf TipoDatoCampoId$ = "DateTime" Then
              TTM$ = "(entidad" & TABLA$ & "." & CampoName$ & " == null) ?  " & Chr$(34) & Chr$(34) & ": entidad" & TABLA$ & "." & CampoName$ & ".ToString() ;" & vbCrLf
            Else
              TTM$ = "(entidad" & TABLA$ & "." & CampoName$ & " == null) ?  " & Chr$(34) & Chr$(34) & ": entidad" & TABLA$ & "." & CampoName$ & ".ToString() ;" & vbCrLf
              'this.TXTFTECNICACAMB_ANCHO_CM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TXTFTECNICACAMB_ANCHO_CM_KeyPress);
              'si es numerico agrego el evento keypress a las cajas de texto por que tambien se va a generar codigo de solonumeros para los txt
              PARRAFOTEXTBOX_2$ = PARRAFOTEXTBOX_2$ & "this.TXT" & Text18(0) & "_" & CampoName$ & ".KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TXT" & Text18(0) & "_" & CampoName$ & "_KeyPress" & ");" & vbCrLf
              '==============================================================================================================="
              '   funcion solo numeros a cajas de texto numericas
              '==============================================================================================================="
'              'VectorSoloNumeros= private void TXTFTECNICACAMB_ANCHO_CM_KeyPress (object sender, KeyPressEventArgs e
'                {
'                TextBox textbox = (TextBox)sender; // Convierto el sender a TextBox
'                fn.solo_numeros(ref textbox, e); // Llamamos a nuestro método
'                }
              VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "       //===============================================================================================================" & vbCrLf
              VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "private void TXT" & Text18(0) & "_" & CampoName$ & "_KeyPress (object sender, KeyPressEventArgs e)" & vbCrLf
              VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "{" & vbCrLf
              VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "TextBox textbox = (TextBox)sender; // Convierto el sender a TextBox" & vbCrLf
              If Left$(TipoDatoCampoId$, 3) <> "Int" Then
                 VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "fn.solo_numeros(ref textbox, e); // Llamamos a nuestro método" & vbCrLf
              Else
                 VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "fn.solo_numeros_enteros(ref textbox, e); // Llamamos a nuestro método" & vbCrLf
              End If
              VectorSoloNumeros(INDICEX%) = VectorSoloNumeros(INDICEX%) & "             " & "}" & vbCrLf
              '==============================================================================================================="
              '   FIN   ****   funcion solo numeros a cajas de texto numericas
              '==============================================================================================================="

              
            End If
            VectorMostrarDatosenCajasTexto(INDICEX%) = "this.TXT" & TABLA$ & "_" & CampoName$ & ".Text = " & TTM$
      End If
   Next i&

End Function
Function PARRAFOBUTTON$()
   'this.button1 = new System.Windows.Forms.Button();
   CmpoIdRepli$ = CampoIdReplica$
   For i& = 1 To MSF3.Rows - 1
      ROTULO$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 4)))
      ID_01$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 11)))
      CamposID$ = TRIM$(MSF3.TextMatrix(i&, 1))
      Campos$ = TRIM$(MSF3.TextMatrix(i&, 7))
      TABLAYCONDI$ = TRIM$(MSF3.TextMatrix(i&, 12))
      TIPDATO$ = TRIM$(MSF3.TextMatrix(i&, 2))
      
      'ACA VALIDACIONES
      If ID_01$ = "SI" Then
         If Campos$ <> "" Then
             'SI LA CONSULTA DE LA GRILLA ESTA VACIA ARMA LA CONSULTA CON LOS CAMPOS Y LA TABLA ORDENADA POR EL CAMPO ID
             If TABLAYCONDI$ = "" Then
               TABLA$ = Text18(0)
               If TIPDATO$ = "ENTERO" Then
                  CONDI$ = CamposID$ & " > 0 ORDER BY " & CamposID$
               ElseIf Left$(TIPDATO$, 5) = "TEXTO" Then
                  CONDI$ = CamposID$ & " <> '' OR " & CamposID$ & " IS NOT NULL ORDER BY " & CamposID$
               End If
               CONSULTACOMPLETASQL$ = ARMACONSULTADEGRILLA$(TABLA$, Campos$, CONDI$)
             Else
               If Campos$ = "" Then
                 Call COMUNI("Faltan Ingfresear Campos o Se Ha Configurado Mal en Fila: " & i&)
                 Exit Function
               End If
               Call ArmaTABLAYCONDI(TABLAYCONDI$, RET_TABLA$, RET_CONDI$)
               TABLA$ = RET_TABLA$
               CONDI$ = RET_CONDI$
               CONSULTACOMPLETASQL$ = ARMACONSULTADEGRILLA$(TABLA$, Campos$, CONDI$)
             End If
             
             'SI LA CONSULTASQL TIENE CONTENIDO TOMA EL MISMO, CUALQIERA DE LOS 2 SE VALIDAN
             RESULTADOX% = VALIDARCAMPOSYCONSULTA%(TABLA$, Campos$, CONSULTACOMPLETASQL$)
             If RESULTADOX% > 0 Then
               If PARRAFOBUTTON$ = "" Then
                  PARRAFOBUTTON$ = "this.BTNECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.Button();" & vbCrLf
                  If CamposID$ = CmpoIdRepli$ Then
                     PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.TXTECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.TextBox();" & vbCrLf
                  Else
                     PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.LBLECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.Label();" & vbCrLf
                  End If
               Else
                  PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.BTNECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.Button();" & vbCrLf
                  'AGREGA EL LABEL DEL ECO
                  If CamposID$ = CmpoIdRepli$ Then
                    PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.TXTECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.TextBox();" & vbCrLf
                  Else
                    PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.LBLECO" & Text18(0) & "_" & CamposID$ & " = new System.Windows.Forms.Label();" & vbCrLf
                  End If
               End If
               Text21.TEXT = Text21.TEXT & GenerarCodigoVectorLLamadaZelecho$("BTNECO" & Text18(0) & "_" & CamposID$, TABLA$, Campos$, CONDI$, i&) & vbCrLf
             End If
             
         End If
      End If
      
   Next i&
   If Me.CheckbtnGrabar.Value > 0 Then
      PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.BTNGRABAR = new System.Windows.Forms.Button();" & vbCrLf
   End If
   
   If Me.CheckbtnGrabar.Value > 0 Then
      PARRAFOBUTTON$ = PARRAFOBUTTON$ & vbCrLf & "this.BTNNUEVO = new System.Windows.Forms.Button();" & vbCrLf
   End If

End Function


Function PARRAFOBUTTON_2$()

'ESTA FUNCION CREA LOS BOTONES PARA LOS ECO Y TAMBIEN LOS LABEL DEL ECO
'            //
'            // button1
'            //
'            this.button1.Location = new System.Drawing.Point(239, 60);

'            this.button1.Name = "button1";
'            this.button1.Size = new System.Drawing.Size(20, 20);
'            this.button1.TabIndex = 6;
'            this.button1.Text = "button1";
'            this.button1.UseVisualStyleBackColor = true;

   TOPX = 20
   LEFTX = 195
   CmpoIdRepli$ = CampoIdReplica$
   For i& = 1 To MSF3.Rows - 1
      ROTULO$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 4)))
      ID_01$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 11)))
      CampoName$ = TRIM$(MSF3.TextMatrix(i&, 1))
      Campos$ = TRIM$(MSF3.TextMatrix(i&, 7))
      CONSULTASQL$ = TRIM$(MSF3.TextMatrix(i&, 12))
      TDATO$ = UCase$(TRIM$(MSF3.TextMatrix(i&, 2)))
      'ACA VALIDACIONES
      If ID_01$ = "SI" Then
         If Campos$ <> "" Then
             INDICEX% = INDICEX% + 1
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "//" & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "// BTNECO" & Text18(0) & "_" & CampoName$ & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "//" & vbCrLf
'             this.BTNECODocumentoOrigen.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf

             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Location = new System.Drawing.Point(" & LEFTX & "," & TOPX & ");" & vbCrLf
'                  this.button1.Name = "button1";
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Name = " & Chr$(34) & "BTNECO" & Text18(0) & "_" & CampoName$ & Chr$(34) & ";" & vbCrLf
'            this.button1.Size = new System.Drawing.Size(20, 20);
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Size = new System.Drawing.Size(20, 20);" & vbCrLf
'                 this.button1.TabIndex = 6;
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".TabIndex = " & INDICEX% & " ;" & vbCrLf
'                 this.button1.Text = "button1";
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Text = " & Chr$(34) & "." & Chr$(34) & ";" & vbCrLf
             
             'this.BTNECOPEDENCA_NroClie.Click += new System.EventHandler(this.BTNECOPEDENCA_NroClie_Click);
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "this.BTNECO" & Text18(0) & "_" & CampoName$ & ".Click += new System.EventHandler(this.BTNECO" & Text18(0) & "_" & CampoName$ & "_Click);" & vbCrLf
             
             VectorControlAdd(INDICEX) = "this.Controls.Add(this.BTNECO" & Text18(0) & "_" & CampoName$ & ");"
             VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.Button BTNECO" & Text18(0) & "_" & CampoName$ & ";"

             INDICEX% = INDICEX% + 1
'             LEFTX = 217
             WIDTHX = 180
             HEIGHTX = 20
             '
             variableThis_lbl_txt$ = "this.LBLECO" 'POR DEFECTO LA VARIABLE ES PARA UN LABEL
             variableThis_lbl_txtSinThis$ = "LBLECO"
             If CmpoIdRepli$ = CampoName$ Then
                variableThis_lbl_txt$ = "this.TXTECO" 'SI ES EL CAMPO IDREPLICA ES UN TEXTBOX
                variableThis_lbl_txtSinThis$ = "TXTECO"
             End If
             
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "//" & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "// LBLECO" & Text18(0) & "_" & CampoName$ & vbCrLf
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & "//" & vbCrLf
'             this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;" & vbCrLf
             
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Font = new System.Drawing.Font(" & Chr$(34) & "Arial" & Chr$(34) & ", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));" & vbCrLf
             
             'this.label1.Location = new System.Drawing.Point(221, 63);
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Location = new System.Drawing.Point(" & LEFTX + 22 & "," & TOPX & ");" & vbCrLf
'            this.label1.Name = "label1";
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Name = " & Chr$(34) & variableThis_lbl_txtSinThis$ & Text18(0) & "_" & CampoName$ & Chr$(34) & ";" & vbCrLf
             'this.label1.Size = new System.Drawing.Size(79, 15);
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Size = new System.Drawing.Size(" & WIDTHX & "," & HEIGHTX & ");" & vbCrLf
             'this.label1.TabIndex = 9;
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".TabIndex = " & INDICEX% & " ;" & vbCrLf
             'this.label1.Text = "";
             PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Text = " & Chr$(34) & Chr$(34) & ";" & vbCrLf
             If CmpoIdRepli$ = CampoName$ Then
               VectorControlAdd(INDICEX) = "this.Controls.Add(this.TXTECO" & Text18(0) & "_" & CampoName$ & ");"
               VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.TextBox TXTECO" & Text18(0) & "_" & CampoName$ & ";"
             Else
               VectorControlAdd(INDICEX) = "this.Controls.Add(this.LBLECO" & Text18(0) & "_" & CampoName$ & ");"
               VectorSystem_Windows_Forms(INDICEX) = "private System.Windows.Forms.Label LBLECO" & Text18(0) & "_" & CampoName$ & ";"
             
             End If
'             VectorBlanformu(INDICEX%) = variableThis_lbl_txt$ & Text18(0) & "_" & CampoName$ & ".Text =" & Chr$(34) & Chr$(34) & ";" & vbCrLf
             
         End If
      End If
      TOPX = TOPX + 20
      If TOPX > 700 Then
       TOPX = 80
       LEFTX = 395
      End If
   Next i&
   
   If Me.CheckbtnGrabar.Value > 0 Then
      PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & GenerarBotonGrabar(INDICEX, TOPX, LEFTX)
      INDICEX = INDICEX + 1
      TOPX = TOPX + 20
   End If
   
   If Me.CheckbtnGrabar.Value > 0 Then
      PARRAFOBUTTON_2$ = PARRAFOBUTTON_2$ & GenerarBotonNuevo(INDICEX, TOPX, LEFTX)
      INDICEX = INDICEX + 1
   End If
End Function
Function ARMACONSULTADEGRILLA$(TABLA$, CAMPOSSELECT$, Optional CONDI$)
     Campos$ = REPLACAR(CAMPOSSELECT$, ";", ",") 'REEMPLAZO POR QUE LOS CAMPOS SELECT VIENE SEPARADOS POR ;
     SQLX$ = "SELECT " & Campos$
     SQLX$ = SQLX$ + " FROM " & TABLA$ & " WITH(NOLOCK) "
     SQLX$ = SQLX$ + CONDI$
     ARMACONSULTADEGRILLA$ = SQLX$
End Function
Function PARRAFOLAYOUTPANEL$()
   For i& = 1 To MSF3.Rows - 1
      Panel$ = TRIM$(MSF3.TextMatrix(i&, 8))
      If Panel$ <> "" And Panel$ <> Panel_ANT$ Then
         If PARRAFOLAYOUTPANEL$ = "" Then
            PARRAFOLAYOUTPANEL$ = "this." & Panel$ & ".SuspendLayout();"
         Else
            PARRAFOLAYOUTPANEL$ = PARRAFOLAYOUTPANEL$ & vbCrLf & "this." & Panel$ & ".SuspendLayout();"
         End If
      End If
      Panel_ANT$ = Panel$
   Next i&
End Function
Function PARRAFOPANEL$()
   For i& = 1 To MSF3.Rows - 1
      Panel$ = TRIM$(MSF3.TextMatrix(i&, 8))
      If Panel$ <> "" And Panel$ <> Panel_ANT$ Then
         If PARRAFOPANEL$ = "" Then
            PARRAFOPANEL$ = "this." & Panel$ & "= new System.Windows.Forms.Panel();"
         Else
            PARRAFOPANEL$ = PARRAFOPANEL$ & vbCrLf & "this." & Panel$ & "= new System.Windows.Forms.Panel();"
         End If
      End If
      Panel_ANT$ = Panel$
   Next i&
End Function

Private Sub Command36_Click()
   REG& = MSF3.RowSel
   If REG& = 0 Then Exit Sub
   If REG& > 0 Then Call SELECCIONARFILA(MSF3, REG&)
   Call INSERTYDESPLAZAR(REG&, MSF3, 1)

End Sub

Private Sub Command37_Click()
   REG& = MSF3.RowSel
   If REG& = 0 Then Exit Sub
   If REG& > 0 Then Call SELECCIONARFILA(MSF3, REG&)
   Call INSERTYDESPLAZAR(REG&, MSF3, -1)

End Sub

Private Sub Command38_Click()
    TAPRIN$ = Text18(0)
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    Text19 = ""
    Set Rsrecordset = FLEXCONN.Execute("sp_columns [" & TAPRIN$ & "]")
    TX1$ = "public class E_" & TAPRIN$ & vbCrLf
    TX1$ = TX1$ & " {"
    Do Until Rsrecordset.EOF
        TF$ = ""
        cnt = cnt + 1
        NAMEX$ = TRIM$(LCase$(Rsrecordset!column_name))
        primerletra$ = UCase$(Left(NAMEX$, 1))
        restoletras$ = Right(NAMEX$, Len(NAMEX$) - 1)
        NAMEX_1_Mayuscula$ = primerletra$ & restoletras$
        TYPEX$ = Rsrecordset!type_name
        LENGX$ = Rsrecordset!Length
        TCLARO$ = TIPOCLARO$(TYPEX$)
        TDT$ = TRIM$(UCase$(TYPEX$))
        ReDim Preserve VectorCargaValoresEntidad(cnt)
        If Left$(TDT$, 3) = "INT" Or Left$(TDT$, 8) = "SMALLINT" Or Left$(TDT$, 5) = "FLOAT" Then
           Select Case TDT$
             Case "INT": TYPEX$ = "Int32"
             Case "SMALLINT": TYPEX$ = "Int16"
             Case "FLOAT": TYPEX$ = "Decimal"
            End Select
           
           TX1$ = TX1$ & "      private " & TYPEX$ & "? " & NAMEX$ & " = null ;" & vbCrLf     'private Int16? _NTrans_Cli = null ;
           TX1$ = TX1$ & "      public " & TYPEX$ & "? " & NAMEX_1_Mayuscula$ & vbCrLf 'public Int16? NTrans_Cli
           'entidad.Nume_Cli = Convert.ToInt16(lector["Nume_Cli"]);

           VectorCargaValoresEntidad(cnt) = "entidad." & NAMEX_1_Mayuscula$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX_1_Mayuscula$ & Chr(34) & "]);"
        ElseIf Left$(TDT$, 4) = "DATE" Or Left$(TDT$, 13) = "SMALLDATETIME" Then
           TX1$ = TX1$ & "      private DateTime " & "? " & NAMEX$ & ";" & vbCrLf  'private DateTime?  _FechAlta ;
           TX1$ = TX1$ & "      public DateTime " & "? " & NAMEX_1_Mayuscula$ & vbCrLf  'public DateTime? FechAlta
           'texto de fecha para concatenar mas abajo
           TF$ = "//SI TIENE VALOR NULL INICIALIZA" & vbCrLf
           TF$ = TF$ & "        if(" & NAMEX$ & " == null) {" & vbCrLf 'if(_FechAlta == null) {
           TF$ = TF$ & "        DateTime fecha = new DateTime(1980, 01, 01);" & vbCrLf 'DateTime fecha = new DateTime(1980, 01, 01);
           TF$ = TF$ & "         " & NAMEX$ & " = fecha;" & vbCrLf ' _FechAlta = fecha;
           TF$ = TF$ & "        }" & vbCrLf '}
           TYPEX$ = "DateTime"
           VectorCargaValoresEntidad(cnt) = "entidad." & NAMEX_1_Mayuscula$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX_1_Mayuscula$ & Chr(34) & "]);"

        Else
           TYPEX$ = "string "
           TX1$ = TX1$ & "      private  " & TYPEX$ & NAMEX$ & ";" & vbCrLf  'private string _Lentre0_Cli;
           TX1$ = TX1$ & "      public  " & TYPEX$ & NAMEX_1_Mayuscula$ & vbCrLf  'public string Lentre0_Cli
           TYPEX$ = "String "
           VectorCargaValoresEntidad(cnt) = "entidad." & NAMEX_1_Mayuscula$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX_1_Mayuscula$ & Chr(34) & "]);"
        End If
        
          
        TX1$ = TX1$ & "        {" & vbCrLf
        TX1$ = TX1$ & "        get" & vbCrLf
        TX1$ = TX1$ & "           {" & vbCrLf
        TX1$ = TX1$ & TF$ ' VA  A TENER TEXTO SI ES FECHA DATE
        TX1$ = TX1$ & "           return " & " " & NAMEX$ & ";" & vbCrLf
        TX1$ = TX1$ & "           }" & vbCrLf
        TX1$ = TX1$ & "        set" & vbCrLf
        TX1$ = TX1$ & "           {" & vbCrLf
        TX1$ = TX1$ & "           " & " " & NAMEX$ & " = value;" & vbCrLf
        TX1$ = TX1$ & "           }" & vbCrLf
        TX1$ = TX1$ & "        }" & vbCrLf
        TX1$ = TX1$ & "//==============================================================================================================" & vbCrLf
        Rsrecordset.MoveNext
     Loop
     TX1$ = TX1$ & " }" & vbCrLf

     Text19 = TX1$
    On Error GoTo 0
End Sub


Private Sub Command39_Click()
Text22(1) = ""
CmpoIdCondi$ = CampoIdReplica$
TABLA$ = Text18(0)
TipoDatoCampoId$ = tipodatoSql$(TABLA$, CmpoIdCondi$)
'TipoDatoCampoId$ = TipoDatoCampo(TABLA$, CmpoIdCondi$)
TEXTBOXIDREPLICA$ = CAJATEXTOIDREPLICA ' FUNCION
Call AsignarAEntidadCajasDeTexto 'CARGAR EL VECTOR DE ASIGNACION DE CAJAS DE TEXTO A LA ENTIDAD

If CmpoIdCondi$ = "" Then
   Call COMUNI("Imposible Generar Rutinas, Falta CampoIdReplica")
   Exit Sub
End If

Call CargarVectordeCargaValoresEntidad
T1$ = "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  ESTO EN CAPA VISTA PARA QUE QUEDE A MODO GENERAL EN EL FORMULARIO **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "            E_" & TABLA$ & " entidad" & TABLA$ & " = new E_" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "            CN_Negocio" & TABLA$ & "  CN" & TABLA$ & " = new  CN_Negocio" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  FIN ESTO EN CAPA VISTA PARA QUE QUEDE A MODO GENERAL EN EL FORMULARIO **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  ESTO EN CAPA VISTA PARA RECARGAR LA ENTIDAD LUEGO IRAN LAS CAJAS DE TEXTO PARA ASIGNR LOS VALORES DE LA ENTIDAD **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + " private void " & TEXTBOXIDREPLICA$ & "_TextChanged (object sender, EventArgs e)" & vbCrLf
T1$ = T1$ + "        {" & vbCrLf
T1$ = T1$ + "        try" & vbCrLf
T1$ = T1$ + "           {" & vbCrLf
T1$ = T1$ + "           " & TipoDatoCampoId$ & " " & CmpoIdCondi$ & " =0;" & vbCrLf
T1$ = T1$ + "           " & CmpoIdCondi$ & " = Convert.To" & TipoDatoCampoId$ & "(" & TEXTBOXIDREPLICA$ & ".Text);" & vbCrLf
T1$ = T1$ + "           if(" & CmpoIdCondi$ & " > 0)" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            entidad" & TABLA$ & "  = CN" & TABLA$ & ".ActualizaEntidad" & TABLA$ & "ConFiltro(" & CmpoIdCondi$ & ");" & vbCrLf
T1$ = T1$ + "            MostrarDatos_" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "           else" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            BLANFORMU();" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "           }" & vbCrLf

T1$ = T1$ + "       catch" & vbCrLf
T1$ = T1$ + "           {" & vbCrLf
T1$ = T1$ + "           return;" & vbCrLf
T1$ = T1$ + "           }" & vbCrLf
T1$ = T1$ + "        }" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  FIN ESTO EN CAPA VISTA PARA RECARGAR LA ENTIDAD LUEGO IRAN LAS CAJAS DE TEXTO PARA ASIGNR LOS VALORES DE LA ENTIDAD **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf


T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  ESTO EN CAPA VISTA Generar Codigo del Boton GrabarD **********" & vbCrLf
T1$ = T1$ + "  // **  Carga la asignacion a la entidad leyendo las cajas de texto" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "        private void BTNGRABAR_Click (object sender, EventArgs e)" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf

T1$ = T1$ + "            {if(BTNGRABAR.Text != " & Chr(34) & "Grabar" & Chr(34) & ")" & vbCrLf
T1$ = T1$ + "                {" & vbCrLf
T1$ = T1$ + "                BTNGRABAR.Text = " & Chr(34) & "Grabar" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "                " & TEXTBOXIDREPLICA$ & ".Enabled = true;" & vbCrLf
T1$ = T1$ + "                BTNNUEVO.Text = " & Chr(34) & "Nuevo" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "                }" & vbCrLf
T1$ = T1$ + "            else " & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            try" & vbCrLf
T1$ = T1$ + "                {" & vbCrLf
             For i& = 1 To UBound(VectorAsignarCajasTextoaEntidad)
                 T1$ = T1$ + "      " & VectorAsignarCajasTextoaEntidad(i&)
             Next i&
'                entidad.Raso_Cli = Convert.ToString(textBox2.Text);
'                entidad.Stat_Cli = Convert.ToInt16(textBox3.Text);
T1$ = T1$ + "                CN" & TABLA$ & ".Grabar" & TABLA$ & "(entidad" & TABLA$ & " , Convert.To" & TipoDatoCampoId$ & "(" & TEXTBOXIDREPLICA$ & ".Text));" & vbCrLf
T1$ = T1$ + "                }" & vbCrLf
T1$ = T1$ + "            catch" & vbCrLf
T1$ = T1$ + "                {" & vbCrLf
T1$ = T1$ + "                MessageBox.Show(" & Chr(34) & "Imosible Grabar Se Detectaron Datos Mal Ingresados" & Chr(34) & ");" & vbCrLf
T1$ = T1$ + "                }" & vbCrLf
T1$ = T1$ + "              }" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf


T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  FIN ESTO EN CAPA VISTA Generar Codigo del Boton GrabarD **********" & vbCrLf
T1$ = T1$ + "  // **  FIN Carga la asignacion a la entidad leyendo las cajas de texto" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  ESTO EN CAPA NEGOCIOS **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "  // dentro de el siguiente codigo" & vbCrLf


T1$ = T1$ + "  //   public class CN_Negocios" & vbCrLf
T1$ = T1$ + "  //    {" & vbCrLf
T1$ = T1$ + "  //     aqui pegar codigo .............." & vbCrLf
T1$ = T1$ + "         CD_Negocio" & TABLA$ & "  CD" & TABLA$ & " = new  CD_Negocio" & TABLA$ & "();" & vbCrLf

T1$ = T1$ + "          public E_" & TABLA$ & " ActualizaEntidad" & TABLA$ & "ConFiltro (" & TipoDatoCampoId$ & " " & CmpoIdCondi$ & ")" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            E_" & TABLA$ & " entidad = new E_" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "            entidad = CD" & TABLA$ & ".CargarEntidad" & TABLA$ & "(" & CmpoIdCondi$ & ");" & vbCrLf
T1$ = T1$ + "            return entidad;" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN ESTO EN CAPA NEGOCIOS **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  " & vbCrLf & vbCrLf
 
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  EN CAPA DATOS NEGOCIOS CREAR LA CLASE  CD_Negocio  & TABLA$  **********"
T1$ = T1$ + "  //============================================================================================" & vbCrLf
  T1$ = T1$ + "  public class CD_Negocio" & TABLA$ & vbCrLf
T1$ = T1$ + "     //CARGA LAS ENTIDADES SEGUN EL NRO DE ID" & vbCrLf
T1$ = T1$ + "     {" & vbCrLf
T1$ = T1$ + "     Conexion.Conexion CON = new Conexion.Conexion(); //INSTANCIO LA CLASE(LIBRERIA)" & vbCrLf
T1$ = T1$ + "     SqlConnection flexconn = new SqlConnection();" & vbCrLf
'public   E_Clientes CargarEntidadClientes (Int16 numcliente)
T1$ = T1$ + "     public   E_" & TABLA$ & "  CargarEntidad" & TABLA$ & " (" & TipoDatoCampoId$ & " ID)" & vbCrLf
T1$ = T1$ + "     {" & vbCrLf
T1$ = T1$ + "     using(var flexconn = new System.Data.SqlClient.SqlConnection(CON.conexionC()))" & vbCrLf
T1$ = T1$ + "         {" & vbCrLf
T1$ = T1$ + "         flexconn.Open();" & vbCrLf
T1$ = T1$ + "         string SQLX = " & Chr(34) & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "         SQLX = " & Chr(34) & "SELECT  * FROM " & TABLA$ & " WITH(NOLOCK) WHERE " & CmpoIdCondi$ & " = " & Chr(34) & " + ID + " & Chr(34) & " ORDER BY " & CmpoIdCondi$ & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "         SqlCommand cmd = new SqlCommand(SQLX, flexconn);" & vbCrLf
T1$ = T1$ + "         try" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            cmd.Connection = flexconn;" & vbCrLf
T1$ = T1$ + "            cmd.CommandType = CommandType.Text;" & vbCrLf
T1$ = T1$ + "            cmd.CommandText = SQLX;" & vbCrLf
T1$ = T1$ + "            var entidad = new E_" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "            int valor=0;" & vbCrLf
T1$ = T1$ + "            decimal valorDecimal=0;" & vbCrLf
T1$ = T1$ + "            DateTime fedate;" & vbCrLf & vbCrLf


T1$ = T1$ + "            SqlDataReader lector = cmd.ExecuteReader();" & vbCrLf
T1$ = T1$ + "            while (lector.Read())" & vbCrLf
T1$ = T1$ + "              {" & vbCrLf
XX = UBound(VectorCargaValoresEntidad)
For i& = 1 To UBound(VectorCargaValoresEntidad) 'ASIGNA EL VECTOR
   T1$ = T1$ + "               " & VectorCargaValoresEntidad(i&) & vbCrLf
'EJEMPLO: entidad.Nume_Cli = Convert.ToInt32(lector["Nume_Cli"]);
'EJEMPLO: entidad.Raso_Cli = Convert.ToString(lector["Raso_Cli"]);
Next i&

T1$ = T1$ + "              }" & vbCrLf
T1$ = T1$ + "            return entidad;" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "            catch(Exception ex)" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf

T1$ = T1$ + "            throw new Exception(ex.Message);" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "          }" & vbCrLf
T1$ = T1$ + "        }" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN EN CAPA DATOS NEGOCIOS CREAR LA CLASE  CD_Negocio  & TABLA$  **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf & vbCrLf & vbCrLf


T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  ESTO EN CAPA NEGOCIOS PARA GRABAR ENTIDAD SEGUN ID**********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "          public void Grabar" & TABLA$ & "(E_" & TABLA$ & " entidad, " & TipoDatoCampoId$ & " ID )" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            CD" & TABLA$ & ".CD_GRABAR" & TABLA$ & "( entidad,  ID);" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN ESTO EN CAPA NEGOCIOS PARA GRABAR ENTIDAD SEGUN ID **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf & vbCrLf & vbCrLf
            
            
T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********   ESTO EN CAPA DATOS EN CD_NEGOCIODEUDOR12 PARA GRABAR ENTIDAD DONDE ARMA LA CONSULTA **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "        public void CD_GRABAR" & TABLA$ & "(E_" & TABLA$ & " Ent" & TABLA$ & "," & TipoDatoCampoId$ & " ID)" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf

T1$ = T1$ + "            Transacciones.Transacciones armaConsulta = new Transacciones.Transacciones();" & vbCrLf
T1$ = T1$ + "            armaConsulta.UpdateConConsulta(Ent" & TABLA$ & " ,  " & Chr(34) & TABLA$ & Chr(34) & ", " & Chr(34) & CmpoIdCondi$ & " = " & Chr(34) & " + ID);" & vbCrLf
T1$ = T1$ + "            }"
T1$ = T1$ + "        //-------------------------------------------------------------------------------------------" & vbCrLf
T1$ = T1$ + "        //-------------------------------------------------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
T1$ = T1$ + "        }"
T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN ESTO EN CAPA DATOS EN CD_NEGOCIODEUDOR12 PARA GRABAR ENTIDAD DONDE ARMA LA CONSULTA  **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf




T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  ESTO EN CAPA VISTA Generar Codigo del Boton NUEVO **********" & vbCrLf
T1$ = T1$ + "  // **  Carga la asignacion a la entidad leyendo las cajas de texto"
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "        private void BTNNUEVO_Click (object sender, EventArgs e)" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            {if(BTNNUEVO.Text == " & Chr(34) & "Nuevo" & Chr(34) & ")" & vbCrLf
T1$ = T1$ + "                {" & vbCrLf
T1$ = T1$ + "                BTNNUEVO.Text = " & Chr(34) & "Grabar Nuevo" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "                BTNGRABAR.Text = " & Chr(34) & "Salir Modo Alta" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "                " & TEXTBOXIDREPLICA$ & ".Enabled = false;" & vbCrLf
T1$ = T1$ + "                BLANFORMU(); " & vbCrLf

T1$ = T1$ + "                }" & vbCrLf
T1$ = T1$ + "            else " & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "           //validar datos mandatarios" & vbCrLf
T1$ = T1$ + "           E_" & TABLA$ & " entidad" & TABLA$ & " = new E_" & TABLA$ & "();//instancio nuevamente la entidad para limpiarla" & vbCrLf
T1$ = T1$ + "          //asignar las cajas de texto a la entidad" & vbCrLf
                        For i& = 1 To UBound(VectorAsignarCajasTextoaEntidad)
                            T1$ = T1$ + "          " & VectorAsignarCajasTextoaEntidad(i&)
                        Next i&

T1$ = T1$ + "            CN_Negocio" & TABLA$ & "  CN" & TABLA$ & " = new  CN_Negocio" & TABLA$ & "();" & vbCrLf
T1$ = T1$ + "          " & TipoDatoCampoId$ & " resul = CN" & TABLA$ & ".RN_INSERT" & TABLA$ & " (entidad" & TABLA$ & ");" & vbCrLf
T1$ = T1$ + "          " & TEXTBOXIDREPLICA$ & ".Text = resul.ToString();" & vbCrLf
T1$ = T1$ + "           BTNNUEVO.Text = " & Chr(34) & "Nuevo" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "           BTNGRABAR.Text = " & Chr(34) & "Grabar" & Chr(34) & ";" & vbCrLf
T1$ = T1$ + "           " & TEXTBOXIDREPLICA$ & ".Enabled = true;" & vbCrLf
T1$ = T1$ + "              }" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // **  FIN ESTO EN CAPA VISTA Generar Codigo del Boton NUEVO **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf & vbCrLf & vbCrLf


T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  ESTO EN CAPA NEGOCIOS CODIGO INSERTAR NUEVO**********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf
'          public  Int32 RN_INSERTDEUDOR12 (E_DEUDOR12 EntDeudor12)
'              {
'              return CD.CD_INSERTDEUDOR12(EntDeudor12);
'              }
T1$ = T1$ + "         public " & TipoDatoCampoId$ & " RN_INSERT" & TABLA$ & " (E_" & TABLA$ & " Ent" & TABLA$ & ")" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "            return CD" & TABLA$ & ".CD_INSERT" & TABLA$ & " (Ent" & TABLA$ & ");"
T1$ = T1$ + "            }" & vbCrLf

T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN ESTO EN CAPA NEGOCIOS CODIGO INSERTAR NUEVO* **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf & vbCrLf & vbCrLf


T1$ = T1$ + "   //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  ESTO EN CAPA DATOS NEGOCIOS CODIGO INSERTAR POR STORED PROCEDURE**********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf

T1$ = T1$ + "         public " & TipoDatoCampoId$ & " CD_INSERT" & TABLA$ & " (E_" & TABLA$ & " Ent" & TABLA$ & ")" & vbCrLf
T1$ = T1$ + "            {" & vbCrLf
T1$ = T1$ + "          " & TipoDatoCampoId$ & " ID=0 ;" & vbCrLf
T1$ = T1$ + "            Transacciones.Transacciones insert" & TABLA$ & " = new Transacciones.Transacciones();" & vbCrLf
T1$ = T1$ + "            ID= insert" & TABLA$ & ".InsertarConSP(Ent" & TABLA$ & "," & Chr(34) & "ALTA" & TABLA$ & Chr(34) & " );" & vbCrLf
T1$ = T1$ + "            return ID;" & vbCrLf
T1$ = T1$ + "            }" & vbCrLf

T1$ = T1$ + "  //============================================================================================" & vbCrLf
T1$ = T1$ + "  // ***********  FIN ESTO EN CAPA DATOS NEGOCIOS CPOR STORED PROCEDURE* **********" & vbCrLf
T1$ = T1$ + "  //============================================================================================" & vbCrLf & vbCrLf & vbCrLf


Text22(1) = T1$


Call ArmaCodigodeAsingacionaCajasdeTexto
End Sub


Private Sub Command40_Click()
    TAPRIN$ = Text18(0)
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    Text19 = ""
    Set Rsrecordset = FLEXCONN.Execute("sp_columns [" & TAPRIN$ & "]")
    Do Until Rsrecordset.EOF
        TF$ = ""
        cnt = cnt + 1
        NAMEX$ = TRIM$(Rsrecordset!column_name)
'        primerletra$ = UCase$(Left(NAMEX$, 1))
'        restoletras$ = Right(NAMEX$, Len(NAMEX$) - 1)
        NAMEX_1_Mayuscula$ = NAMEX$
        TYPEX$ = Rsrecordset!type_name
        LENGX$ = Rsrecordset!Length
        TCLARO$ = TIPOCLARO$(TYPEX$)
        TDT$ = TRIM$(UCase$(TYPEX$))
        ReDim Preserve VectorCargaValoresEntidad(cnt)
        If Left$(TDT$, 3) = "INT" Or Left$(TDT$, 8) = "SMALLINT" Or Left$(TDT$, 5) = "FLOAT" Then
           Select Case TDT$
             Case "INT": TYPEX$ = "Int32"
             Case "SMALLINT": TYPEX$ = "Int16"
             Case "FLOAT": TYPEX$ = "Decimal"
            End Select
           
           TX1$ = TX1$ & "      public " & TYPEX$ & " " & NAMEX_1_Mayuscula$ & " { get; set; }" & vbCrLf

           VectorCargaValoresEntidad(cnt) = "entidad." & NAMEX_1_Mayuscula$ & " = Convert.To" & TYPEX$ & "(lector[" & Chr(34) & NAMEX_1_Mayuscula$ & Chr(34) & "]);"
        ElseIf Left$(TDT$, 4) = "DATE" Or Left$(TDT$, 13) = "SMALLDATETIME" Then
           TX1$ = TX1$ & "      public DateTime " & " " & NAMEX_1_Mayuscula$ & " { get; set; }" & vbCrLf

        Else
           TYPEX$ = "string "
           TX1$ = TX1$ & "      public  " & TYPEX$ & NAMEX_1_Mayuscula$ & " { get; set; }" & vbCrLf   'public int Nume_Cli { get; set; }
           TYPEX$ = "String "
        End If
        TX1$ = TX1$ & "//==============================================================================================================" & vbCrLf
        Rsrecordset.MoveNext
     Loop

     Text19 = TX1$
    On Error GoTo 0

End Sub

Private Sub Command50_Click(Index As Integer)
   '
   Dim CAJATEXTO As TextBox
   Select Case SSTab2.Tab
      Case 0
       Set CAJATEXTO = Text23
      Case 1
       Set CAJATEXTO = Text19
      Case 2
       Set CAJATEXTO = Text22(0)
      Case 3
       Set CAJATEXTO = Text21
      Case 4
       Set CAJATEXTO = Text22(1)
      Case 5
       Set CAJATEXTO = Text26
      
   End Select
   
   Select Case Index
   Case 0
     COMIE = CAJATEXTO.SelStart
     LOSEL = CAJATEXTO.SelLength
     CAJATEXTO.TEXT = Left$(CAJATEXTO.TEXT, COMIE) + Mid$(CAJATEXTO.TEXT, COMIE + LOSEL + 1)
   Case 1
     Clipboard.SetText Mid$(CAJATEXTO.TEXT, CAJATEXTO.SelStart + 1, CAJATEXTO.SelLength)
   Case 2
     COMIE = CAJATEXTO.SelStart
     LOSEL = CAJATEXTO.SelLength
     CAJATEXTO.TEXT = Left$(CAJATEXTO.TEXT, COMIE) + Clipboard.GetText + Mid$(CAJATEXTO.TEXT, COMIE + LOSEL + 1)
   Case 3
     CAJATEXTO.SelStart = 0
     CAJATEXTO.SelLength = Len(CAJATEXTO.TEXT)
     CAJATEXTO.SetFocus

   
   End Select
End Sub

'Function PARRAFOLAYOUTPANEL$()
'   'this.panel1.SuspendLayout();
'   For i& = 1 To MSF3.Rows - 1
'      Panel$ = TRIM$(MSF3.TextMatrix(i&, 8))
'      If Panel$ <> "" And Panel$ <> Panel_ANT$ Then
'         If PARRAFOLAYOUTPANEL$ = "" Then
'            PARRAFOLAYOUTPANEL$ = "this." & Panel$ & ".SuspendLayout();"
'         Else
'            PARRAFOLAYOUTPANEL$ = PARRAFOLAYOUTPANEL$ & vbCrLf & "this." & Panel$ & ".SuspendLayout();"
'         End If
'      End If
'      Panel_ANT$ = Panel$
'   Next i&
'End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")
End Sub

Private Sub lbl_DblClick(Index As Integer)

    TXTX$ = "PINTAR EL SIGUIENTE CODIGO  Y PEGAR EL CODIGO DE DISEÑO GENERADO" & vbCrLf
    TXTX$ = TXTX$ & "private void InitializeComponent ()" & vbCrLf
    TXTX$ = TXTX$ & "{" & vbCrLf
    TXTX$ = TXTX$ & "this.components = new System.ComponentModel.Container();" & vbCrLf
    TXTX$ = TXTX$ & "this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;" & vbCrLf
    TXTX$ = TXTX$ & "this.Text = 'Form4';" & vbCrLf
    TXTX$ = TXTX$ & "}" & vbCrLf & vbCrLf

    TXTX$ = TXTX$ & "#endregion"
    MsgBox TXTX$
End Sub


Private Sub List4_Click()
    TAPRIN$ = List4.TEXT
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    List5.Clear
    Me.Text18(0) = TAPRIN$
    '
    List5.Clear
    Set Rsrecordset = FLEXCONN.Execute("sp_columns [" & TAPRIN$ & "]")
    Do Until Rsrecordset.EOF
        cnt = cnt + 1
        NAMEX$ = Rsrecordset!column_name
        TYPEX$ = Rsrecordset!type_name
        LENGX$ = Rsrecordset!Length
        Call REPLA(TX$, NAMEX$, 1, 32)
        TCLARO$ = TIPOCLARO$(TYPEX$)
        Call REPLA(TX$, TCLARO$, 34, 14)
        Call REPLA(TX$, LENGX$, 50, 16)
        List5.AddItem PREFPRIN$ + TX$
        Rsrecordset.MoveNext
     Loop
    On Error GoTo 0
End Sub

Private Sub List6_Click()

End Sub

Private Sub List6_DblClick()
  
   List6.Visible = False
   
End Sub


Private Sub LSTMOSTARCMP2_Click()
     Call TEXTOLOTES$(LSTMOSTARCMP2.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     Me.MSFUPD.TextMatrix(1, COLU1&) = LSTMOSTARCMP2.TEXT
     LSTMOSTARCMP2.Visible = False

End Sub

Private Sub LSTMOSTARCMP2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       LSTMOSTARCMP2.Visible = False
    End If
End Sub


Private Sub LSTTABLAS2_Click()
    TAPRIN$ = LSTTABLAS2.TEXT
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    LSTMOSTARCMP2.Clear
    LSTMOSTARCMP2.AddItem ""
    Me.TXTNAMETABLA = TAPRIN$
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           'CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           LSTMOSTARCMP2.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           'Combo4.AddItem AUX.Fields(i& - 1).Name
        End If
    Next i&

End Sub

Private Sub mnuArmaEnumeEncabezado_Click()
10     textox$ = InputBox$("Leyenda (Prompt)", "Ingresar Texto", textox$)
       TXTCAMPOSENCABEZADO = ""
       Dim strArray() As String
       strArray = Split(textox$, ";")
       For i% = 0 To UBound(strArray)
           J% = i% + 1
           LEIDO$ = strArray(i%)
           LEIDO2$ = "'  " & J% & vbCrLf
           Call REPLA(TX$, LEIDO$, 1, 28)
           Call REPLA(TX$, LEIDO2$, 30, 8)
           TEXTOCOMPLETO$ = TEXTOCOMPLETO$ & TX$
       Next i%
       
       TXTCAMPOSENCABEZADO.Visible = True
       TXTCAMPOSENCABEZADO = TEXTOCOMPLETO$
       SSTab1.Tab = 1

 
End Sub

Private Sub mnuArmaSql_Click()
  ARMASQLX.Show 1
  
End Sub

Private Sub MnuAutonumTabAux_Click()
       Call COMUNI("CUANDO EL CAMPO 1 NO ESTA EDITABLE EL SISTEMA AUTONUMERA\ESTE CAMPO PONERLO COMO NUMERICO F...\CONFIGURAR QUE NO SE PUEDA BORRAR NI INSERTAR\PARA QUE NO SE AUTONMERE Y CAMBIE LOS CODIGOS ORIGINALES")

End Sub

Private Sub mnuAyudaVentabnew_Click()
   Call AYUDA("ATAJOS EN VENTABNEW (VENTANA DE CARGA)")
End Sub

Private Sub mnuCeldaResultadoFormula_Click()
   Call AYUDA(mnuCeldaResultadoFormula.Caption)
End Sub

Private Sub mnuCreaEstructura_Click()
    Call AYUDA("Crear Archivo Estrucutura")
   
End Sub

Private Sub mnupyes_Click()
Pyes.Show 1
End Sub

Private Sub mnuReverUpd_Click()
    FRMREVUPD.Show 1
End Sub

Private Sub mnuTablasAuxiliares_Click()
  'Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)
  
  '


End Sub

Private Sub mnuUpdate_Click()
   Call AYUDA("GEN.COD.Update")
End Sub

Private Sub MSF_RowColChange()
    MSFESPEJO.Rows = MSF.Rows
    MSFESPEJO.Cols = MSF.Cols
End Sub

Private Sub MSF_Scroll()
   If TXTMOVIL.Visible = True Then
     Call TXTMOVIL_LostFocus
     TXTMOVIL.Visible = False
   End If
End Sub

Private Sub CAMAUXI_Click()
  Call ARMASELFIELD
End Sub

Private Sub chkUnionAll_Click()

    LSTTABLASEC.ListIndex = -1
    UNIONALL$ = ""
    CMBJUNTAR.TEXT = ""
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    
    Call ARMA_CONSULTX
    
End Sub

Private Sub CMBJUNTAR_Click()
   Call ARMA_CONSULTX
End Sub


'
Private Sub CMDBORRAR_Click()
    TXTCONSULTA.TEXT = ""
End Sub

Private Sub Combo2_Click()
    Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & Combo2.TEXT))
    For i& = 1 To AUX.Fields.Count
        Combo3.AddItem AUX.Fields(i& - 1).Name
    Next i&
End Sub

Private Sub Combo6_Click()

    Text5 = Combo6.TEXT
    
End Sub

Private Sub Command10_Click()
'    Static CTX%
'    If CTX% < 1 Then
'       Call ARMALISTARCH(LUDAT$, "*.RFS")
'       CTX% = 1
'    End If
'    Call SELECHO("INDIRFS")   ' , "DIRECTORIO RFS")
'    VDEVU$ = TRIM$(VALACT1$("INDIRFS"))
'    If VDEVU$ <> "" Then
'       Text5 = VDEVU$
'    End If

End Sub

Private Sub Command11_Click()
    If Label1 <> "" Then
       If COMALTER%("¿ Confirma Eliminar la Consulta '" & TRIM$(Label1) & "' ?\  \(Esta Transaccion no Podrá Revertirse)\\No, Conservarla\Si, Eliminarla") = 2 Then
          CONDI$ = "COD_QUERY = '" & Label1.Caption & "'"
          Call BORRAREGISTROS("SQL_QUERIES", CONDI$, REGAF&, "NOMESS")
          
          Call ABRETACON
          STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
          Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
          If Not (MIRS.EOF And MIRS.BOF) Then
             MIRS.Delete
             DEFTREP07.Label4 = ""
             DEFTREP07.Label9 = ""
             Call DEFTREP07.VerConsul
             MIRS.Close
             Unload Me
          End If
       End If
    End If
End Sub

Private Sub Command12_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 3) <> "OR " Then
          List2.List(i%) = "OR " & List2.TEXT
         End If
         Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call GENERAR_CODIGO(OK%)
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Right(TRIM$(List2.TEXT), 1) <> ")" Then
          List2.List(i%) = List2.TEXT & ")"
        End If
        Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command15_Click()
    'INTERVENCION OR

    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 1) <> "(" Then
          List2.List(i%) = "(" & List2.TEXT
        End If
        Call ARMA_CONSULTX
    End If
End Sub



Private Sub Command7_Click()
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
      If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
        TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
        CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
        CONDIVALO$ = AJUSTI$(Text7, 512)
        'If TRIM$(TAVALO$) <> "" Then
        '  If TRIM$(CAVALO$) <> "" Then
            If TRIM$(CONDIVALO$) <> "" Then
              CAMPOSE$ = AJUSTI$(List1.TEXT, 168) + TAVALO$ + CAVALO$ + CONDIVALO$
              List1.RemoveItem (i&)
              List1.AddItem CAMPOSE$, i&
              List1.Selected(i&) = True
            End If
        '  End If
        'End If
      End If
    End If
    '
    For i& = 1 To LSTMOSTARCMP.ListCount
       If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          LSTMOSTARCMP.RemoveItem (i& - 1)
          LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
          LSTMOSTARCMP.Selected(i& - 1) = True
       End If
    Next i&
    '
    For i& = 1 To CAMAUXI.ListCount
       If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          CAMAUXI.RemoveItem (i& - 1)
          CAMAUXI.AddItem CAMPOSE$, i& - 1
          CAMAUXI.Selected(i& - 1) = True
       End If
    Next i&
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command9_Click()
    '
    LSTTABLASEC.ListIndex = (-1)
    Combo4.Clear
    Combo5.Clear
    TAPRIN$ = LSTTABLAS.TEXT
    If TAPRIN$ = "" Then Exit Sub
    '
10  For U& = 1 To LSTMOSTARCMP.ListCount
       If Left$(LSTMOSTARCMP.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          LSTMOSTARCMP.RemoveItem (U& - 1)
          GoTo 10
       End If
    Next U&
    '
20  For U& = 1 To List1.ListCount
       If Left$(List1.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          List1.RemoveItem (U& - 1)
          GoTo 20
       End If
    Next U&
    '
    Call LSTTABLAS_Click     ' PARA REFRESCAR EL COMBO4
    '
End Sub

Private Sub label1_Change()
    Dim FISELE$(128)
    List2.Clear
    
    CANFI% = 0
    ESTAENSQL% = 0
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK) where COD_QUERY = '" & Label1.Caption & "'"
    Set RS = READSET(SQLX$)
    If Not (RS.EOF And RS.BOF) Then
       Label9 = SAFETEXT$(RS!DES_QUERY)
       TTXX$ = SAFETEXT$(RS!PAR_QUERY)
       ESTAENSQL% = 1
    End If
    If ESTAENSQL% = 1 Then
      RS.Close
      Set RS = Nothing
      GoTo 30
    End If
    
    STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_QUERY
       TTXX$ = MIRS!PAR_QUERY
       '
30     While Len(TTXX$) > 0
          PPXX% = InStr(TTXX$, Chr$(124))
          If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
          COMAI$ = Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 1)
          '
          If Left$(COMAI$, 10) = "DATABASE =" Then
              Call CMBBASE_CLICK
              CMBBASE = UCase$(TRIM$(Mid$(COMAI$, 11)))
              GoTo 19
            ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
              TASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLAS.ListCount
                 If UCase$(LSTTABLAS.List(KKI% - 1)) = TASELE Then
                   LSTTABLAS.Selected(KKI% - 1) = True
                   Call LSTTABLAS_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "SELTASEC =" Then
              TASELSEC$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLASEC.ListCount
                 If UCase$(LSTTABLASEC.List(KKI% - 1)) = TASELSEC Then
                   LSTTABLASEC.Selected(KKI% - 1) = True
                   Call LSTTABLASEC_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "MATCHFI1 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo4.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "MATCHFI2 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo5.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "SELFIELD =" Then
              CANFI% = CANFI% + 1
              FISELE$(CANFI%) = TRIM$(Mid$(COMAI$, 11))
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
              For KKI% = 1 To LSTMOSTARCMP.ListCount
                 If UCase$(TRIM$(Left$(LSTMOSTARCMP.List(KKI% - 1), 64))) = CASELE$ Then
                   LSTMOSTARCMP.RemoveItem KKI% - 1
                   LSTMOSTARCMP.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   LSTMOSTARCMP.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
              For KKI% = 1 To CAMAUXI.ListCount
                 If UCase$(TRIM$(Left$(CAMAUXI.List(KKI% - 1), 64))) = CASELE$ Then
                   CAMAUXI.RemoveItem KKI% - 1
                   CAMAUXI.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   CAMAUXI.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
              List2.AddItem TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "TIPOJOIN =" Then
              CMBJUNTAR.TEXT = TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "UNIONALL =" Then
              chkUnionAll.Value = 1
              SegundaTabla$ = TRIM$(Mid$(COMAI$, 11))
              For idTbSec% = 0 To LSTTABLASEC.ListCount - 1
                    If TRIM$(LSTTABLASEC.List(idTbSec%)) = SegundaTabla$ Then
                        LSTTABLASEC.ListIndex = idTbSec
                        UNIONALL$ = SegundaTabla$
                        TablaUnionOk% = 1
                        Exit For
                    End If
              Next idTbSec%
              If TablaUnionOk% = 0 Then
                chkUnionAll.Value = 0
                UNIONALL$ = ""
              End If
            Else
          End If
19     Wend
       '
       List1.Clear
       For KKI% = 1 To CANFI%
          List1.AddItem FISELE$(KKI%)
       Next KKI%
       '
       Call ARMA_CONSULTX
       Exit Sub
       '
    End If
    '
    'SI NO ESTABA EN SQL ENTRO EN MDB Y CIERRA EL RECORDSET
    If ESTAENSQL% = 0 Then
      MIRS.Close
      Set MIRS = Nothing
    End If

    Call LSTTABLAS_Click
    '
End Sub

Private Sub Command1_Click()
   If CMBCAMPOS.TEXT = "" Or CMBOPERADOR = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   If Len(TRIM$(CMBCAMPOS.TEXT)) > 24 Then
      MsgBox "Nombre de Campo Demasiado Largo"
      Exit Sub
   End If
   SQLY$ = AJUSTI$(CMBCAMPOS.TEXT, 24) + " " + AJUSTI$(CMBOPERADOR.TEXT, 8)
   VALOCO$ = TRIM$(Combo1.TEXT)
   '
   TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
   If (TPC = 52 Or TPC = 62) And VALOCO$ = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   '
   ' ESTO ES PARA QUE PERMITA COMPARAR VALORES DEL RS ENTRE SÍ
   For KKU& = 1 To CMBCAMPOS.ListCount
      If UCase$(CMBCAMPOS.List(KKU& - 1)) = UCase$(VALOCO$) Then GoTo 10
   Next KKU&
   '
   If TPC <> 52 And TPC <> 62 Then VALOCO$ = "'" + VALOCO$ + "'"
   '
10 SQLY$ = SQLY$ + " " + VALOCO$

   List2.AddItem SQLY$
   Call ARMA_CONSULTX
   Exit Sub
   
      If KLIK% > 0 Then GoTo 25
      SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " Where " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      KLIK% = KLIK% + 1
      TXTCONSULTA.TEXT = SQLX$
      Exit Sub
      
25    SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " AND " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      TXTCONSULTA.TEXT = SQLX$
      
End Sub

Function TIPODECAMPO(NTABLA$, NCAMPO$)
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    STRSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO = MIRS!xtype
    End If

End Function

Private Sub CMDCOPIAR_Click()
   Clipboard.SetText Mid$(TXTCONSULTA.TEXT, TXTCONSULTA.SelStart + 1, Text1.SelLength)
End Sub

Private Sub CORTAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

Private Sub Command2_Click()
   If CMBCAMPOS.TEXT <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CMBCAMPOS.TEXT & " AS VALOR FROM " & LSTTABLAS.TEXT & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CMBCAMPOS.TEXT & " ASC"
      '
      Combo1.Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      On Error Resume Next
      Set VAPOSI = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Ejecutar Consulta - Guarde y Re-Intente")
         Exit Sub
      End If
      On Error GoTo 0
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1.AddItem !valor
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
      '
   End If
   
End Sub

Private Sub Command3_Click()
10  COMO$ = InputBox$("Leyenda (Prompt)", "Configuración Consultas SQL", CMBCAMPOS.TEXT)
    If COMO$ = "" Then GoTo 10
    Combo1.Clear
    Combo1.AddItem ("¿" + COMO$ + "?")
    Combo1.TEXT = "¿" + COMO$ + "?"
End Sub

Private Sub Command4_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& > 1 Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i& - 2
         List1.Selected(i& - 2) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
End Sub

Private Sub Command5_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& < List1.ListCount Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i&
         List1.Selected(i&) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
   Call ARMA_CONSULTX
End Sub

Private Sub Command6_Click()
    '
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
       TOTALI$ = "0": If Check1.Value = 1 Then TOTALI$ = "1"
       AGRUPA$ = "0": If Check2.Value = 1 Then AGRUPA$ = "1"
       OCULTA$ = "0": If Check4.Value = 1 Then OCULTA$ = "1"
       MODOSOR$ = "    "
       If Option2.Value = True Then MODOSOR$ = "ASC "
       If Option3.Value = True Then MODOSOR$ = "DESC"
       CAMPOSE$ = AJUSTI$(List1.List(i&), 168)
       While Len(CAMPOSE$) < 168: CAMPOSE$ = CAMPOSE$ + " ": Wend
       Mid$(CAMPOSE$, 65, 8) = AJUSTI$(Text1, 8)
       Mid$(CAMPOSE$, 73, 1) = OCULTA$
       Mid$(CAMPOSE$, 75, 1) = AGRUPA$
       Mid$(CAMPOSE$, 76, 1) = TOTALI$
       Mid$(CAMPOSE$, 77, 4) = MODOSOR$
       Mid$(CAMPOSE$, 81, 64) = AJUSTI$(Text2, 64)
       Mid$(CAMPOSE$, 145, 16) = AJUSTI$(Text5, 16)
       Mid$(CAMPOSE$, 161, 8) = AJUSTI$(Text6, 8)
       '
       If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
         TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
         CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
         CONDIVALO$ = AJUSTI$(Text7, 512)
         If TRIM$(CONDIVALO$) <> "" Then
            While Len(CAMPOSE$) < (168 + 64 + 64 + 512): CAMPOSE$ = CAMPOSE$ + " ": Wend
            Mid$(CAMPOSE$, 169, 64) = TAVALO$
            Mid$(CAMPOSE$, 233, 64) = CAVALO$
            Mid$(CAMPOSE$, 297, 512) = CONDIVALO$
         End If
       End If
       List1.RemoveItem (i&)
       List1.AddItem CAMPOSE$, i&
       List1.Selected(i&) = True
       '
       For i& = 1 To LSTMOSTARCMP.ListCount
          If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             LSTMOSTARCMP.RemoveItem (i& - 1)
             LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
             LSTMOSTARCMP.Selected(i& - 1) = True
          End If
       Next i&
       '
       For i& = 1 To CAMAUXI.ListCount
          If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             CAMAUXI.RemoveItem (i& - 1)
             CAMAUXI.AddItem CAMPOSE$, i& - 1
             CAMAUXI.Selected(i& - 1) = True
          End If
       Next i&
       '
    End If
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command8_Click()
   'ADAT$ = Combo2.TEXT
   ADAT$ = Label1.Caption
   Call CARGADEBASE(ADAT$, TEBUSCA$, OKXX%)
   
End Sub

Private Sub EJECUTAR_Click()
    SHOWREP07.Label2 = ""
    If TRIM$(Label9) <> "" Then
          SHOWREP07.Label3 = Label9        ' Descripcion de la consulta
       Else
          SHOWREP07.Label3 = "Consulta en Desarrollo"
    End If
    TEXMO$ = ARMA_CONSULTB
    Call DECOD_CONSULTA(TEXMO$)
    'If TRIM$(Combo2.TEXT) <> "" Then
    If TRIM$(Label1.Caption) <> "" Then
         SHOWREP07.Label2 = ""
         'SHOWREP07.Label2 = Combo2.TEXT
         SHOWREP07.Label2 = Label1.Caption
      Else
         SHOWREP07.Label2 = "_TEXT_"
    End If
    SHOWREP07.Show 1
    MODICONSU% = 0
End Sub

Private Sub GUARDACON_Click()
    '
    FGUAREP07.Text1 = TRIM$(Label1.Caption)
    CODMO$ = TRIM$(Label1.Caption)
    FGUAREP07.Text2 = Label9
    FGUAREP07.Text3 = ""
    FGUAREP07.Label11 = ""
    FGUAREP07.Text4 = ""
    FGUAREP07.Label18 = ""
    FGUAREP07.Show 1
    If FGUAREP07.Label1 = "OK" Then
10     CODMO$ = TRIM$(FGUAREP07.Text1)
       DESMO$ = TRIM$(FGUAREP07.Text2)
       TEXMO$ = ARMA_CONSULTB
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If
       
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_QUERY = CODMO$
          !DES_QUERY = DESMO$
          !TIP_QUERY = 0        ' REPORTE TABULADO
          !PAR_QUERY = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       '
30     If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
           ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + CODMO$ + ".SQF"
           Call SAVEFILE(ASALI$, TEXMO$)
       End If
       '
       If FGUAREP07.Text3 <> "" Then
          If FGUAREP07.Label11 <> "" Then
             ASALI$ = FGUAREP07.Label11 + "\" + FGUAREP07.Text3 + ".SQF"
             Call SAVEFILE(ASALI$, TEXMO$)
          End If
       End If
       '
    End If
    '
    Label1.Caption = CODMO$
    '
End Sub




Private Sub Label8_DblClick()
    Unload Me
End Sub

Private Sub List1_Click()
   If MODIFENCA% = 2 Then Exit Sub
   If MODIFENCA% = 1 Then
      If COMALTER%("¿ Descarta Cambios ?\\No, Conservar\Si, Descartar") <> 2 Then
         MODIFENCA% = 2
         On Error Resume Next
         List1.ListIndex = ILIS1&
         On Error GoTo 0
         MODIFENCA% = 1
         Exit Sub
      End If
   End If
   '
   CAMPOSE$ = List1.TEXT
   Text1 = TRIM$(Mid$(CAMPOSE$, 65, 8))
   Check3.Value = 0: If Mid$(CAMPOSE$, 74, 1) = "1" Then Check3.Value = 1
   Check4.Value = 0: If Mid$(CAMPOSE$, 73, 1) = "1" Then Check4.Value = 1    ' OCULTA
   Check1.Value = 0: If Mid$(CAMPOSE$, 76, 1) = "1" Then Check1.Value = 1
   Check2.Value = 0: If Mid$(CAMPOSE$, 75, 1) = "1" Then Check2.Value = 1
   Text2 = TRIM$(Mid$(CAMPOSE$, 81, 64))
   Text5 = TRIM$(Mid$(CAMPOSE$, 145, 16))
   Text6 = TRIM$(Mid$(CAMPOSE$, 161, 8))
   '
   TAUXI$ = Mid$(CAMPOSE$, 169, 640)
   Combo2.TEXT = TRIM$(Mid$(TAUXI$, 1, 64))
   Combo3.TEXT = TRIM$(Mid$(TAUXI$, 65, 64))
   Text7.TEXT = TRIM$(Mid$(TAUXI$, 129, 512))
   '
   ORDEX$ = TRIM$(UCase$(Mid$(CAMPOSE$, 77, 4)))
   Option1.Value = True
   If ORDEX$ = "ASC" Then Option2.Value = True
   If ORDEX$ = "DESC" Then Option3.Value = True
   MODIFENCA% = 0
   ILIS1& = List1.ListIndex
   
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   i& = List2.ListIndex
   If i& >= 0 Then
      CMBCAMPOS.TEXT = TRIM$(Left$(TTXX$, 24))
      CMBOPERADOR.TEXT = TRIM$(Mid$(TTXX$, 26, 8))
        VALORX$ = TRIM$(REPLACAR$(Mid$(TTXX$, 35), "'", ""))
      Combo1.TEXT = VALORX$
      List2.RemoveItem i&
   End If
   Call ARMA_CONSULTX
End Sub

Private Sub LSTMOSTARCMP_Click()
  Call ARMASELFIELD
End Sub

Sub ARMASELFIELD()
    '
    For i& = 0 To LSTMOSTARCMP.ListCount - 1
      ESTADO = LSTMOSTARCMP.Selected(i&)
      CAMPO$ = TRIM$(LSTMOSTARCMP.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 20
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
10         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 10
              End If
           Next J&
      End If
20  Next i&
    '
    For i& = 0 To CAMAUXI.ListCount - 1
      ESTADO = CAMAUXI.Selected(i&)
      CAMPO$ = TRIM$(CAMAUXI.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 120
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
110         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 110
              End If
           Next J&
      End If
120 Next i&
    '
   List3.Clear
   For i& = 0 To List1.ListCount - 1
      List1.ListIndex = i&
      List3.AddItem List1.TEXT
   Next i&

    Call ARMA_CONSULTX
    '
End Sub
'
Private Sub LSTTABLAS_Click()
    '
    Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Sub ARMALISTAS()

    TXTCONSULTA.TEXT = ""
    CMBCAMPOS.Clear
    CMBOPERADOR.TEXT = ""
    Combo1.Clear
    Combo4.Clear
    Combo5.Clear
    LSTMOSTARCMP.Clear
    CAMAUXI.Clear
    List1.Clear
    '
    TAPRIN$ = LSTTABLAS.TEXT
    PREFPRIN$ = ""
    TASEC$ = LSTTABLASEC.TEXT
    If TASEC$ <> "" Then PREFPRIN$ = TAPRIN$ + "."
    If TAPRIN$ = "" Then Exit Sub
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           LSTMOSTARCMP.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           Combo4.AddItem AUX.Fields(i& - 1).Name
        End If
    Next i&
    '
    If TASEC$ <> "" Then
       PREFSEC$ = TASEC$ + "."
       Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TASEC$))
       For i& = 1 To AUX.Fields.Count
          If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
            CMBCAMPOS.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            LSTMOSTARCMP.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            Combo5.AddItem AUX.Fields(i& - 1).Name
          End If
       Next i&
    End If
    '
    For i& = 1 To 9
        CAMAUXI.AddItem "ValoBada_" & i&
    Next i&
    '
End Sub

Private Sub CMBBASE_CLICK()
    LSTTABLAS.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Private Sub Form_Load()
   If App.PrevInstance = True Then
      Close: End
   End If
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   'Call CargaFuentesAnchoFijo
   '
   If APLINVO$ <> "" Then Height = 3570
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call ABRECONEXION
   KLIK% = 0
   '
   Call ABRETACON     ' crea la tabla de consultas
'   Call VerConsul
   Call CREATABLAS_SQLCTRL
   Call CREAR_TABLA_CREACOD
  Call ABRECONEXION
  KLIK% = 0
  CMBOPERADOR.Clear
  
  CMBOPERADOR.AddItem " =  "
  CMBOPERADOR.AddItem " <  "
  CMBOPERADOR.AddItem " <= "
  CMBOPERADOR.AddItem " >  "
  CMBOPERADOR.AddItem " >= "
  CMBOPERADOR.AddItem " <> "
  CMBOPERADOR.AddItem "LIKE"
  '
  CMBJUNTAR.AddItem "INNER"
  CMBJUNTAR.AddItem "LEFT"
  CMBJUNTAR.AddItem "RIGHT"
  '
  UNIONALL$ = ""
  '
  Combo6.AddItem "DEUDOR"
  Combo6.AddItem "PROVED"
  Combo6.AddItem "MASTER"
  Combo6.AddItem "USERS"
  Combo6.AddItem "SUTOT"
 
  Call MUESTRATABLAS

  TEPRUEBA.Font = MSF.Font
  TEPRUEBA.FontSize = MSF.Font.Size
  TEPRUEBA.FontBold = MSF.Font.Bold
  MSF.Rows = 10
  MSF.ColWidth(0) = 4000
  MSF.TextMatrix(0, 0) = "Encabezado + Formato (Canti/F8.1)"
  MSF.TextMatrix(1, 0) = "Formula (4;*;6)"
  MSF.TextMatrix(2, 0) = "CampEditables (S/N)"
  MSF.TextMatrix(3, 0) = "ColEcoCampo (CODIGO;Descripcion;Master)"
  MSF.TextMatrix(4, 0) = "TxtUbicadoGrillaTotal(Texto)"
  MSF.TextMatrix(5, 0) = "Total (S/N)"
  MSF.TextMatrix(6, 0) = "Dejar Libre (S/N)"
  MSF.TextMatrix(7, 0) = "CONDICION 1"
  MSF.TextMatrix(8, 0) = "CONDICION 2"
  MSF.TextMatrix(9, 0) = "ENUMERAR"
  '
  Call CARGAMSF4
  Call ABRECONEXION
  Call CREACAMPO("", "GeneCodi", "id", 4, 0, 1)
  Call CREACAMPO("", "GeneCodi", "AddRegistros", 3, 0)
  Call CREACAMPO("", "GeneCodi", "UtilFormula", 3, 0)
  Call CREACAMPO("", "GeneCodi", "Borra", 3, 0)
  Call CREACAMPO("", "GeneCodi", "Inserta", 3, 0)
  Call CREACAMPO("", "GeneCodi", "NoTilde", 3, 0)
  Call CREACAMPO("", "GeneCodi", "NameRecord", 10, 256)
  Call CREACAMPO("", "GeneCodi", "CamposRecord", 12, 0)
  Call CREACAMPO("", "GeneCodi", "UsaRecord", 3, 0)
  Call CREACAMPO("", "GeneCodi", "AotraGrilla", 3, 0)
  Call CREACAMPO("", "GeneCodi", "ConsultaRec", 10, 1056)
  Call CREACAMPO("", "GeneCodi", "UsaRecord", 3, 0)
  Call CREACAMPO("", "GeneCodi", "CodiGen", 12, 0) 'el codigo generado
  Call CREACAMPO("", "GeneCodi", "TitVentana", 10, 256) 'titulo de la ventana
  Call CREACAMPO("", "GeneCodi", "CanCampos", 3, 1) 'Cantidad de Campos que tiene la grilla
  Call CREACAMPO("", "GeneCodi", "MARBORRA", 3, 1)
  '
  Call CREACAMPO("", "DatGrilla", "id", 4, 0, 1)
  Call CREACAMPO("", "DatGrilla", "FilaCol", 10, 32)
  Call CREACAMPO("", "DatGrilla", "Dato", 10, 512)
  Call CREACAMPO("", "DatGrilla", "MARBORRA", 3, 0)
  '

'  Call CARGA_ENCABEZADO(MSF4, "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10", VENTADISE)
'  MSF4.Rows = 2

  '
  '
  Call CREATABLA_ENCAPRF
  SSTab1.Tab = 3
  Text13 = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)  *** CORROBORAR ***"
  Campos$ = "CAMPO/A28;" '1
  Campos$ = Campos$ + "T.DATO/A12;" '2
  Campos$ = Campos$ + "Len/f3;" '3
  Campos$ = Campos$ + "ROTULO/A16;" '4
    
  Campos$ = Campos$ + "FORMATO/A9;" '5
  Campos$ = Campos$ + "ID_REPLICA/A9;" '6   SI O NO
  Campos$ = Campos$ + "CAMPO_ECO/A48;" '7  EL CAMO ECO QUE SE VA A MOSTRAR O SEPARADOS POR ;
  Campos$ = Campos$ + "FRAME/A7;" '8  NOMBRE DEL FRAME O NUMERO
  Campos$ = Campos$ + "Graba/A7;" '9
  Campos$ = Campos$ + "MUESTRA_RS/A10;" '10
  Campos$ = Campos$ + "ID01/A5;" '11 SI ES ID DE PRINCIPAL PARA REALIZAR EL UPDATE
  Campos$ = Campos$ + "CONDICION/A128;" '12 'SI NO INGRESA EN ESTE ESPACIO RESOLVER EN CODIGO.
  Campos$ = Campos$ + "ENCABEZADO_ZELECHO/A128" '13 'EJ. Cta Cliente/f9;Razon Social/A32
  Call CARGA_ENCABEZADO(MSF3, Campos$, Me) ' SI
  LBLNUEVO.Visible = True
  TXTCODIGO = NumNuevaCreacod
'    TEXT21.TEXT.Height = 1785
'    TEXT21.TEXT.Left = 8100
'    TEXT21.TEXT.Top = 2280
'    TEXT21.TEXT.Width = 9360

    Call COMPLETAR_TABLAS
End Sub

Sub MUESTRATABLAS()
    LSTTABLAS.Clear: LSTTABLASEC.Clear
    Combo2.Clear: Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
       LSTTABLASEC.AddItem UCase$(AUX!Name)
       Combo2.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Sub ARMA_CONSULTX()
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = TRIM$(Left$(List1.List(i&), 64))
       '
       If UCase$(Left$(CAMPO$, 8)) = "VALOBADA" Then GoTo 19
       '
       ALIASCAMPO$ = CAMPO$
       If Mid$(List1.List(i&), 75, 1) = "1" Then
          CAMPO1$ = CAMPO$
          If InStr(CAMPO1$, ".") > 0 Then
            PPXYZ% = InStr(CAMPO1$, ".")
            CAMPO1$ = Mid$(CAMPO1$, PPXYZ% + 1)
          End If
          ALIASCAMPO$ = "Sx" + CAMPO1$
          CAMPO$ = "SUM(" + CAMPO$ + ") AS " + ALIASCAMPO$
          AGRUPA% = 1
       End If
       If VUELTA& = 0 Then
           SQLX$ = SQLX$ + CAMPO$
          Else
           SQLX$ = SQLX$ + ", " + CAMPO$
       End If
       VUELTA& = VUELTA& + 1
       '
       ORDEX$ = UCase$(TRIM$(Mid$(List1.List(i&), 77, 4)))
       If ORDEX$ <> "" Then
          If VUELTB& > O Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = Mid$(List1.List(i&), 75, 1)
       If AGRUX$ <> "1" Then
          If VUELTC& > O Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If

19  Next i&
    '
    If VUELTA = 0 Then SQLX$ = "SELECT *"
    '
    TAPRIN$ = LSTTABLAS.TEXT
    SQLY$ = " FROM " + TAPRIN$
    '
    TASEC$ = LSTTABLASEC.TEXT
    
    If TASEC$ <> "" Then
       CAMPRIN$ = TRIM$(Combo4.TEXT)
       If CAMPRIN$ <> "" Then
          CAMSEC$ = TRIM$(Combo5.TEXT)
          If CAMSEC$ <> "" Then
             TIPOJOIN$ = TRIM$(CMBJUNTAR.TEXT)
             If TIPOJOIN$ = "" Then TIPOJOIN$ = "INNER"
             SQLY$ = SQLY$ + " " & TIPOJOIN$ & " JOIN"
             SQLY$ = SQLY$ + " " + TASEC$
             
             SQLY$ = SQLY$ + " ON " + TAPRIN$ + "." + CAMPRIN$
             SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
          End If
       End If
    End If
    '
    If List2.ListCount > 0 Then
       JJ& = 0
       For i& = 0 To List2.ListCount - 1
          SQLY$ = SQLY$ '+ vbCrLf
          'INTERVENCION OR
          If JJ& < 1 Then
              SQLY$ = SQLY$ + " WHERE "
          ElseIf UCase$((Left$(TRIM$(List2.List(i&)), 3))) = "OR " Then
              SQLY$ = SQLY$ + " "
          Else
               SQLY$ = SQLY$ + " AND "
          End If
          SQLY$ = SQLY$ + TRIM$(Left$(List2.List(i&), 24)) + " " + TRIM$(Mid$(List2.List(i&), 26, 8)) + " " + TRIM$(Mid$(List2.List(i&), 35))
          JJ& = JJ& + 1
       Next i&
    End If
    '
    TXTCONSULTA.TEXT = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " ORDER BY " + SQLZ$
    '
    '
    If chkUnionAll.Value = 1 Then
        If LSTTABLASEC.ListIndex > 0 Then
            PrimeraTabla$ = TRIM$(LSTTABLAS.List(LSTTABLAS.ListIndex))
            SegundaTabla$ = TRIM$(LSTTABLASEC.List(LSTTABLASEC.ListIndex))
            CONSULTASQL$ = TXTCONSULTA.TEXT
            POSI% = InStr(1, CONSULTASQL$, PrimeraTabla$)
            If POSI% > 0 Then
                CONSULTASQL$ = Mid$(CONSULTASQL$, 1, POSI% - 1) & SegundaTabla$ & Mid$(CONSULTASQL$, POSI% + Len(PrimeraTabla$))
                TXTCONSULTA.TEXT = SQLX$ + SQLY$ & " UNION ALL " & CONSULTASQL$
            End If
            
            UNIONALL$ = SegundaTabla$
        End If
    End If
       
    TXTCANREGI.TEXT = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
End Sub

Function ARMA_CONSULTB()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(FGUAREP07.Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(FGUAREP07.Text2) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   0" + Chr$(124)
    CONSUB$ = CONSUB$ + "SELTABLE = " + LSTTABLAS.TEXT + Chr$(124)
    '
      TASECU$ = TRIM$(LSTTABLASEC.TEXT)
    If TASECU$ <> "" And chkUnionAll.Value = 0 Then CONSUB$ = CONSUB$ + "SELTASEC = " + TASECU$ + Chr$(124)
      CAMPOMATCH1$ = TRIM$(Combo4.TEXT)
    If CAMPOMATCH1$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI1 = " + CAMPOMATCH1$ + Chr$(124)
      CAMPOMATCH2$ = TRIM$(Combo5.TEXT)
    If CAMPOMATCH2$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI2 = " + CAMPOMATCH2$ + Chr$(124)
    '
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = List1.List(i&)
       CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    Next i&
    '
    For i& = 0 To List2.ListCount - 1
       CONDI$ = List2.List(i&)
       CONSUB$ = CONSUB$ + "CONDITIO = " + CONDI$ + Chr$(124)
    Next i&
    '
    TIPODEJOIN$ = TRIM$(CMBJUNTAR.TEXT)
    CONSUB$ = CONSUB$ + "TIPOJOIN = " + TIPODEJOIN$ + Chr$(124)
    '
    If chkUnionAll.Value = 1 And LSTTABLASEC.ListIndex > 0 Then
        SegundaTabla$ = LSTTABLASEC.List(LSTTABLASEC.ListIndex)
        CONSUB$ = CONSUB$ + "UNIONALL = " + SegundaTabla$ + Chr$(124)
    End If
    
    ARMA_CONSULTB = CONSUB$
End Function

Private Sub CMDPEGAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Clipboard.GetText + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

'QUEDA PENIDIENTE SI ES UN TEXTO  EL QUE VA A COMPARAR VAYA ENTRE _
COMILLAS SIMPLES (HACERLO CON ISNUMERIC ) _
SI EL OPERADOR ES LIKE DEBERIA LLEVAR EL SIMBOLO DEL PORCENTAJE _
FALTA TAMBIEN EL TEMA DEL EDITOR DE TEXTO _
VALIDAR LOS CAMPOS DE CONDICION _
CAMBIAR LAS PROPIEDADES DE LOS COMBO Y TEXT DE LA CONDICION PAARA QUE _
NO SEAN EDITABLES _
Y EJECUTAR LA CONSULTA

Private Sub LSTTABLASEC_Click()
    '
    If chkUnionAll.Value = 0 And UNIONALL$ = "" Then Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSF3_DblClick()
    Unload VENTABNEW
    
  Text13 = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)  *** CORROBORAR ***"
  Campos$ = "CAMPO/A28;" '1
  Campos$ = Campos$ + "T.DATO/A12;" '2
  Campos$ = Campos$ + "Len/f3;" '3
  Campos$ = Campos$ + "ROTULO/A16;" '4
    
  Campos$ = Campos$ + "FORMATO/A9;" '5
  Campos$ = Campos$ + "ID_REPLICA/A9;" '6   SI O NO
  Campos$ = Campos$ + "CAMPO_ECO/A48;" '7  LOS CAMO ECO QUE SE VA A MOSTRAR O SEPARADOS POR ;
  Campos$ = Campos$ + "FRAME/A7;" '8  NOMBRE DEL FRAME O NUMERO
  Campos$ = Campos$ + "Graba/A7;" '9
  Campos$ = Campos$ + "MUESTRA_RS/A10;" '10
  Campos$ = Campos$ + " ID01/A5;" '11 SI ES ID DE PRINCIPAL PARA REALIZAR EL UPDATE
  Campos$ = Campos$ + " CONDICION/A128;" '12 'SI NO INGRESA EN ESTE ESPACIO RESOLVER EN CODIGO.
  Campos$ = Campos$ + " ENCABEZADO_ZELECHO/A128" '13 'SI NO INGRESA EN ESTE ESPACIO RESOLVER EN CODIGO.

    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    'VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.HabilitaBorrar = 1
    VENTANA.TITULO = "Armado Inicial"
    VENTANA.CampEditables = "4;5;6;7;8;9;10;11;12;13"
    VENTANA.VentVectorToolTipText(4) = "Nombre del Label Identificador de la Caja de Texto"
    VENTANA.VentVectorToolTipText(6) = "Si o No (Identifica si es el Id para realizar el update y que va a replicar el resto de los campos"
    VENTANA.VentVectorToolTipText(7) = "Nombre Real del/los Campos a mostrar en Zelecho  (separados x ;) CONDICION ANGREGAR EN CAMPO CONDICION, SI CONDICION = "" TOMA TABLA BASE "
    VENTANA.VentVectorToolTipText(8) = "Nombre Real del Panel Contenedor (Se crea Código para que se genere automaticamente)"
    VENTANA.VentVectorToolTipText(9) = "Si o No identificando si se va a Grabar este dato (Si en la lista hay alguno Genera Boton y Codigo Grabar"
    VENTANA.VentVectorToolTipText(10) = "Si o No Por ahora No Tiene Uso Definido, Pensado para que si es ID01 = SI , Muestre una grilla tipo detalle"
    VENTANA.VentVectorToolTipText(11) = "Si o No (Identifica si este campo va a ser id de campo Eco Zelecho)"
    VENTANA.VentVectorToolTipText(12) = "Consulta sql para el zelecho (Depende del 6 y 7) Si se deja vacia se arma con 2,6 Y 7 (Ej. de Condi: Tabla$; CONDI$"
    VENTANA.VentVectorToolTipText(13) = "Si Columna 11 = Si Completar este indicando Encabezados en Zelecho'EJ. Cta Cliente/f9;Razon Social/A32"
    VENTANA.Inicializar = 1
    
    Call CopiarMsfAOtroMsf(Me.MSF3, VENTABNEW.MSF, 1)
    '************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, Me.MSF3, 1)
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub

Private Sub MSFUPD_Click()
    COL% = MSFUPD.MouseCol
    REG& = MSFUPD.MouseRow
    If REG& > 0 Then Exit Sub
    If COL% < 1 Then Exit Sub
    LSTMOSTARCMP2.ToolTipText = REG& & "," & COL%
    LSTMOSTARCMP2.TEXT = MSFUPD.TextMatrix(REG&, COL%)

    MSFUPD.COL = COL%

    With MSFUPD
        .COL = COL%
        Me.LSTMOSTARCMP2.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8
        Me.LSTMOSTARCMP2.Visible = True
        Me.LSTMOSTARCMP2.ZOrder 0
        FilaActiva = .Row
    End With
End Sub
Private Sub MSFUPD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       Me.TXTFORMUEXCEL = Me.TXTFORMUEXCEL + UBICACIONX1$(MSFUPD.MouseRow, MSFUPD.MouseCol)
    End If
    
End Sub


'Private Sub MSF3_DblClick()
'    VENTANA.Inicializar = 0
'    VENTANA.AgregaRegistro = 1
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaBorrar = 0
'    VENTANA.HabilitaInsertar = 1
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.Titulo = "EDIT"
'    VENTANA.CampEditables = "1;2;3;4;5;6"
'    VENTANA.Campox = "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10"
'    VENTANA.Inicializar = 1
'    Call CopiarMsfAOtroMsf(MSF3, VENTABNEW.MSF, 1)
'    '****************************************************************
'
'    VENTABNEW.Show 1
'    If VENTABNEW.APROBADO% > 0 Then
'       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF3, 1)
'    End If
'    Set VENTANA = Nothing
' End Sub
    

'Private Sub MSF_RowColChange()
''    For i& = 1 To MSF.Cols - 1
''       If msf2.Cols < MSF.Cols Then msf2.Cols = MSF.Cols
''       msf2.TextMatrix(0, i&) = MSF.TextMatrix(0, i&)
''       msf2.COL = i&
''       TEPRUEBA.Caption = MSF.TextMatrix(0, i&)
''       msf2.ColWidth(i&) = TEPRUEBA.Width + 300
''
''    Next i&
'
'End Sub

Private Sub Option1_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option2_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option3_Click()
   MODIFENCA% = 1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   A = PreviousTab
   B = SSTab1.Tab
   If LSTTABLAS2.ListCount > 0 Then Exit Sub
   If B = 1 Then
        LSTTABLAS2.Clear
        Me.TXTNAMETABLA = ""
        Dim AUX As ADODB.Recordset
        Set AUX = New ADODB.Recordset
22      SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
        Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Do While Not AUX.EOF
           If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
           LSTTABLAS2.AddItem UCase$(AUX!Name)
29      AUX.MoveNext
        Loop
   ElseIf B = 3 Then
        Call COMPLETAR_TABLAS
   End If
End Sub

Sub COMPLETAR_TABLAS()
        List4.Clear
        Me.Text18(0) = ""
        Set AUX = New ADODB.Recordset
        SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
        Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Do While Not AUX.EOF
           If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 30
           List4.AddItem UCase$(AUX!Name)
30      AUX.MoveNext
        Loop
        AUX.Close
        Set AUX = Nothing
End Sub

Private Sub Text1_Change()
   MODIFENCA% = 1
End Sub

Private Sub Text11_Change()
   A = Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       MSF4.Cols = XVALO(Text11)
       Call POSICIONAR_TXTMOVIL(VENTADISE.TXTMOVIL2, VENTADISE.MSF4, 0, 1)
    End If

End Sub


Private Sub Text15_DblClick()
   If Text15.Left <> MSF5.Left Then
        Text15.Left = MSF5.Left
        Text15.Top = MSF5.Top
        Text15.Height = 4000
        Text15.Width = MSF5.Width
   Else
        Text15.Left = 5020
        Text15.Top = 6750
        Text15.Height = 1305
        Text15.Width = 10200
   
   End If
End Sub


Private Sub Text18_Change(Index As Integer)
   If Index = 0 Then
      TABLA$ = TRIM$(Text18(Index))

      If TABLA$ <> "" Then
         For i& = 0 To List4.ListCount - 1
            If TRIM$(List4.List(i&)) = TABLA$ Then
                List4.Selected(i&) = True
            End If
         Next i&
      Else
         List4.Clear
      End If
   End If
End Sub

Private Sub Text19_DblClick()
      Text25.Visible = True
      Text25.TEXT = Text19.TEXT
      Text25.Height = 7785
      Text25.Left = 240
      Text25.Top = 360
      Text25.Width = 17300

End Sub


Private Sub Text2_Change()
   MODIFENCA% = 1
End Sub

'==============================================================================================================
Private Sub MSF_Click()
          COLUM = MSF.MouseCol
          REG& = MSF.MouseRow
          If COLUM < 1 Then Exit Sub
          TXTMOVIL.ToolTipText = REG& & "," & COLUM
          TXTMOVIL.TEXT = MSF.TextMatrix(REG&, COLUM)

          MSF.COL = COLUM
          
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            .Visible = True
            .ToolTipText = REG& & "," & COLUM
            .SetFocus
          End With

End Sub
Private Sub MSF4_Click()
          COLUM = MSF4.MouseCol
          REG& = MSF4.MouseRow
          If COLUM < 1 Then Exit Sub
          TXTMOVIL2.ToolTipText = REG& & "," & COLUM
          TXTMOVIL2.TEXT = MSF4.TextMatrix(REG&, COLUM)

          MSF4.COL = COLUM
          
          With TXTMOVIL2
            .ZOrder 0
            .Width = MSF4.CellWidth
            .Left = MSF4.CellLeft + MSF4.Left
            .Top = MSF4.CellTop + MSF4.Top
            .Visible = True
            .ToolTipText = REG& & "," & COLUM
            .SetFocus
          End With

End Sub

Private Sub Text20_DblClick()
   If Text20.Locked = True Then
      Text20.Locked = False
      Text20.BackColor = vbWhite
      Call PINTATEXT(Text20)
   Else
      Text20.BackColor = &HE0E0E0
      Text20.Locked = True
   End If

End Sub
Private Sub Text14_DblClick()
   If Text14.Locked = True Then
      Text14.Locked = False
      Text14.BackColor = vbWhite
      Call PINTATEXT(Text14)
   Else
      Text14.BackColor = &HE0E0E0
      Text14.Locked = True
   End If

End Sub


Private Sub Text21_DblClick()
      Text25.Visible = True
      Text25.TEXT = Text21.TEXT
      Text25.Height = 7785
      Text25.Left = 240
      Text25.Top = 360
      Text25.Width = 17300
End Sub


Private Sub Text22_DblClick(Index As Integer)
      Text25.Visible = True
      Text25.TEXT = Text22(Index).TEXT
      Text25.Height = 7785
      Text25.Left = 240
      Text25.Top = 360
      Text25.Width = 17300
End Sub


Private Sub Text23_DblClick()
      Text25.Visible = True
      Text25.TEXT = Text23.TEXT
      Text25.Height = 7785
      Text25.Left = 180
      Text25.Top = 360
      Text25.Width = 17300
End Sub


Private Sub Text25_DblClick()
   Text25.Visible = False
End Sub


Private Sub Text26_DblClick()
      Text25.Visible = True
      Text25.TEXT = Text26.TEXT
      Text25.Height = 7785
      Text25.Left = 240
      Text25.Top = 360
      Text25.Width = 17300
End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       MSF.Cols = XVALO(Text8)
       Call POSICIONAR_TXTMOVIL(VENTADISE.TXTMOVIL, VENTADISE.MSF, 0, 1)
    End If
End Sub


Private Sub Timer1_Timer()
   Call PINTATEXT(TXTMOVIL)
   Timer1.Enabled = False
End Sub

 
Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      GRID1.Visible = False
      GRID2.Visible = True
      Exit Sub
    End If
    '
    GRID1.Rows = 2
    For i& = 1 To GRID1.Cols - 1
       GRID1.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    GRID1.Visible = True
    GRID1.ZOrder 0
    GRID1.Cols = GRID2.Cols
    GRID1.Redraw = False
    GRID2.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID2.Rows - 1
      For J& = 1 To GRID2.Cols - 1
        If (InStr(UCase$(GRID2.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID2.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            GRID1.Rows = REG& + 1
            For H& = 1 To GRID1.Cols - 1
              GRID1.TextMatrix(REG&, H&) = GRID2.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    GRID1.Redraw = True
    
    '
End Sub


Private Sub TXTBUSCADOR_Change()
   If TRIM$(TXTBUSCADOR) = "" Then
      Exit Sub
    Else
       FILABUSCAR& = 1
    End If
    '
    Call BuscarTexto(MSF3, VENTADISE.FILABUSCAR&, 1, TXTBUSCADOR.TEXT, Me)
    
    '
End Sub


Private Sub MSF3_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If TRIM$(Me.TXTBUSCADOR) <> "" Then
          Call BuscarTexto(MSF3, FILABUSCAR&, 1, TXTBUSCADOR, Me)
       End If
   End If
End Sub

Private Sub TXTBUSCADOR_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    MSF3.SetFocus
  End If
End Sub

Private Sub TXTCAMPOSENCABEZADO_DblClick()
       TXTCAMPOSENCABEZADO.Visible = False
End Sub

Private Sub TXTCODIGO_Change()
    Call BLANFORMUFRM
'    If XVALO(TXTCODIGO) > 0 Then
'      If CantRegistros&("CREACOD", "CODIGO = " & XVALO(TXTCODIGO)) < 1 Then
'         ESNUEVO% = 1
'         Exit Sub
'      End If
'   Else
'      Exit Sub
'   End If
   REG& = 0
   Dim MSF As MSFlexGrid
   Set MSF = MSF3
   
   SQLX$ = "SELECT * FROM CREACOD  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIGO = " & XVALO(TXTCODIGO)
   SQLX$ = SQLX$ + " order by CREACOD_ID"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        If REG& = 1 Then
          TXTNAMEFRM = SAFETEXT$(!FRMNAME)
          Text24 = SAFETEXT$(!REFERENCIA)
          Text18(0) = SAFETEXT$(!TABLA)
        End If
        TXTFECHA = !FECHEMISI
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!CAMPO)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!T_DATO)
        MSF.TextMatrix(REG&, 3) = XVALO(!LONGITUD)
        MSF.TextMatrix(REG&, 4) = SAFETEXT$(!ROTULO)
        MSF.TextMatrix(REG&, 5) = SAFETEXT$(!Formato)
        MSF.TextMatrix(REG&, 6) = SAFETEXT$(!ID_REPLICA)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!ESCAMPOECO)
        MSF.TextMatrix(REG&, 8) = SAFETEXT$(!Frame)
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!Graba)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!MUESTRA_RS)
        MSF.TextMatrix(REG&, 11) = SAFETEXT$(!ID01)
        MSF.TextMatrix(REG&, 12) = SAFETEXT$(!SQLXCAMPOECO)
        MSF.TextMatrix(REG&, 13) = SAFETEXT$(!ENCABEZADOSZELECHO)
       
        .MoveNext
     Loop
   End With
   Call numerargrilla(MSF, 0)
End Sub

Private Sub TXTCONSULTA_Change()
   TXTCONSU_ESPEJO = TXTCONSULTA
End Sub

Private Sub TXTFORMUEXCEL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
'    Set tTip = New CLSTOOLTIP
'             ' almacena la fila y columna actual
'
'
'   ' Establece El tipo ( balloon o normal )
'   tTip.Estilo = TTBalloon
'   ' Indica el icono a utilizar ( info, Warning , error etc..)
'   tTip.Icono = TTIconInfo
'   tTip.Delay = 50 ' Tiempo de duración
'  Static f As Long
'  Static c As Long
'  With TXTFORMUEXCEL
'   If f <> .Left Or c <> .Top Then
'             ' almacena la fila y columna actual
'             f = .Left
'             c = .Top
'             ' texto y titulo
''             tTip.Titulo = "AYUDA   "
'             TT1$ = "Texto de la celda : " & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'              tTip.Texto = TT1$
'             tTip.Crear TXTFORMUEXCEL.hwnd ' crea el Tips
'    Else
'        ' Lo destruye
'        tTip.Destroy
'    End If
'End With
'
''  Static f As Long
''  Static c As Long
'
''  With MSFlexGrid1
''    ' verifica que el mouse está en una columna o fila
''    If .MouseRow > 0 And .MouseCol >= 0 Then
''
''          If f <> .MouseRow Or c <> .MouseCol Then
''             ' almacena la fila y columna actual
''             f = .MouseRow
''             c = .MouseCol
''
''             ' texto y titulo
''             tTip.Titulo = "Ejemplo"
''             tTip.Texto = "Texto de la celda :" & _
''                         .TextMatrix(.MouseRow, .MouseCol)
''
''             tTip.Crear .hwnd ' crea el Tips
''           End If
''    Else
''        ' Lo destruye
''        tTip.Destroy
''    End If
''  End With
  
End Sub

Private Sub TXTID_Change()
    Call BLANFORMU
    IDX& = XVALO(TXTID)
    SQLX$ = "SELECT *  FROM GeneCodi WITH(NOLOCK) WHERE ID = " & IDX&
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
            TXTTITULO = SAFETEXT$(!TitVentana)
            Check5.Value = XVALO(!ADDREGISTROS)
            Check6.Value = XVALO(!UtilFormula)
            Check7.Value = XVALO(!Borra)
            Check8.Value = XVALO(!Inserta)
            Check9.Value = XVALO(!NOTILDE)
            Text10.TEXT = SAFETEXT$(!NameRecord)
            Check10.Value = XVALO(!USARECORD)
            Check11.Value = XVALO(!AotraGrilla)
            TXTCONSU_ESPEJO = SAFETEXT$(!ConsultaRec)
            CamposRecordset$ = SAFETEXT$(!CamposRecord)
            Text9 = SAFETEXT$(!CodiGen)
            Text8 = XVALO(!CanCampos)
            .MoveNext
        Loop
    End With
    Call Text8_KeyPress(13)
    If TRIM$(CamposRecordset$) <> "" Then
       Call TEXTOLOTES$(CamposRecordset$, "|")
       On Error Resume Next
       For i& = 1 To UBound(CADENATEX$)
          List3.AddItem CADENATEX$(i&)
       Next i&
       On Error GoTo 0
    End If
    '
    RS.Close: Set RS = Nothing
    '
    '
    
    SQLX$ = "SELECT * FROM DatGrilla WITH(NOLOCK) WHERE ID = " & IDX&
    Set RS1 = READSET(SQLX$)
    With RS1
       Do While Not .EOF
         FICO$ = TRIM$(SAFETEXT$(!FilaCol))
         If FICO$ <> "" Then
           On Error Resume Next
           Call TEXTOLOTES$(FICO$, "|")
           REG1& = XVALO(CADENATEX$(1))
           FIL1& = XVALO(CADENATEX$(2))
           MSF.TextMatrix(REG1&, FIL1&) = SAFETEXT$(!DATO)
           On Error GoTo 0
         End If
         .MoveNext
        Loop
    End With
    RS1.Close: Set RS1 = Nothing
    TXTMOVIL.Visible = False ' SI NO CUANDO SE POSICIONABA BORRABA LA CELDA
    Call MSF_RowColChange
End Sub

Private Sub TXTID_DblClick()
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID/F8;TITULO VENTANA/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID/F8;TITULO VENTANA/A64", FRMZELECHO)
    FRMZELECHO.Caption = "LISTA DE CODIGOS GENERADOS"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_CODIGENE()
    Call Carga_MSF_Recordset(RS1, "ID/F8;TITULO VENTANA/A64", FRMZELECHO.msf2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    
    IDX& = XVALO(RESP_ZELE_VECT(1))
    
    TXTID = IDX&
 
End Sub
Private Sub TXTIDLISTA_Change()
  IDLIST& = XVALO(TXTIDLISTA)
  MSF4.Rows = 0
  'Text11 = IDLIST&
  Call CARGAMSF4
  
  MAXIMOCAMPOS = XVALO(Text11)
  
  MSF4.Cols = MAXIMOCAMPOS + 1
  Call CARGADATOSUNICOS(IDLIST&)
  SQLX$ = "SELECT * FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA = " & IDLIST&
  SQLX$ = SQLX$ + " ORDER BY ENCAPRF_ID ASC"
  Set RS = READSET(SQLX$)
  With RS
  J& = 0
    Do While Not .EOF
      J& = J& + 1
      MSF4.TextMatrix(0, J&) = SAFETEXT$(!Encabezado)
      MSF4.TextMatrix(1, J&) = SAFETEXT$(!Formula)
      MSF4.TextMatrix(2, J&) = SAFETEXT$(!CAMPEDITAB)
      MSF4.TextMatrix(3, J&) = SAFETEXT$(!COLCAMECO)
      MSF4.TextMatrix(4, J&) = SAFETEXT$(!TXTTOTAL)
      MSF4.TextMatrix(5, J&) = SAFETEXT$(!Total)
      MSF4.TextMatrix(6, J&) = SAFETEXT$(!DejarLibre)
      'VER SEGUN LA CANTIDAD DE CAMPOS VALIDARLOS SEGUN EL CAMPO DE CANTIDAD DE CAMPOS QUE SE GUARDA EN EL ENCABEZADO
      'FALTA PRESENTAR LO GUARDADO EN EL ENCABEZADO Y HACER LA RUTINA QUE LEVANTE DESDE LA BASE DE DATOS.
      .MoveNext
    Loop
  End With
  RS.Close
  Set RS = Nothing
  MSF4.Cols = J& + 1
End Sub

Private Sub TXTMOVIL_Change()
'     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
End Sub

Private Sub TXTMOVIL_DblClick()
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COL1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If REG1& = 3 Or REG1& = 7 Or REG1& = 8 Then ' CAMPO ECO O DE CONDICION 1 O CONDICION 2
        ARMACONDI.APROBO% = 0
        'LEVANTO LA CONDICION SI YA ESTABA GUARDADA.
        ARMACONDI.Text7 = MSFESPEJO.TextMatrix(REG1&, COL1&)
        ARMACONDI.Show 1
'        MSFESPEJO.TextMatrix(REG1&, COL1&) = ""
        If ARMACONDI.APROBO% > 0 Then
          'COPIA LA CONDICION EN UNA GRILLA OCULTA
           If REG1& = 3 Then MSFESPEJO.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
           If REG1& = 7 Or REG1& = 8 Then MSF.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
           
        End If
     End If

End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     COLU1& = XVALO(CADENATEX$(2))
     REG1& = XVALO(CADENATEX$(1))
     
     On Error GoTo 0
'     T_Espera ((0.2))
End Sub


'
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
'     If REG& = 1 Then REG& = 0
     MSF.TextMatrix(REG&, COLU1&) = TXTMOVIL.TEXT
     MSF.COL = COLU1&
     TEPRUEBA.Caption = TXTMOVIL.TEXT
     If TEPRUEBA.Width > MSF.ColWidth(COLU1&) Then
       MSF.ColWidth(COLU1&) = TEPRUEBA.Width + 300
     End If
     On Error GoTo 0
     '
     
     If MSF.Cols - 1 > COLU1& Then
       COLU1& = COLU1& + 1
     Else
       COLU1& = 1
     End If
50   MSF.COL = COLU1&
     With TXTMOVIL
        .ZOrder 0
        .Width = MSF.CellWidth
        .Left = MSF.CellLeft + MSF.Left
        .Top = MSF.CellTop + MSF.Top
        .Visible = True
        .ToolTipText = MSF.Row & "," & COLU1&
        
        'ESTO SE PARA QUE CUANDO VAYA MOVIENDOSE TOME EL TEXTO INGRESADO EN EL ENCABEZADO
         Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
         On Error Resume Next
         REG& = XVALO(CADENATEX$(1))
         COLU1& = XVALO(CADENATEX$(2))
         TXTMOVIL.TEXT = MSF.TextMatrix(REG&, COLU1&)
        .SetFocus
     End With
 
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
     Dim TXTMOVILX As TextBox
     Set TXTMOVILX = TXTMOVIL
     If SSTab1.Tab = 1 Then
        Set TXTMOVILX = TXTMOVIL
        Set MSFX = MSF
     ElseIf SSTab1.Tab = 2 Then
        Set TXTMOVILX = TXTMOVIL2
        Set MSFX = MSF4
     End If
    If TXTMOVILX.Visible = False Then Exit Sub
    
    Call TEXTOLOTES$(TXTMOVILX.ToolTipText, ",")
    On Error Resume Next
    REG1& = XVALO(CADENATEX$(1))
    COLU1& = XVALO(CADENATEX$(2))
    
    Select Case KeyCode
        Case 38 '(flecha arriba)
            On Error Resume Next
            MSF.TEXT = TXTMOVILX.TEXT
            If MSFX.Row > 1 Then
              TXTMOVILX.ToolTipText = REG1& - 1 & "," & COLU1&
            End If
            If MSFX.Row > 1 Then MSFX.Row = MSFX.Row - 1
'          MSFX.COL = 11
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
            
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
            .SetFocus
          End With
            On Error GoTo 0
        Case 40 '(flecha abajo)
            On Error Resume Next
            MSFX.TEXT = TXTMOVILX.TEXT
            If MSFX.Row < MSFX.Rows - 1 Then
              TXTMOVILX.ToolTipText = REG1& + 1 & "," & COLU1&
            End If
            If MSFX.Row < MSFX.Rows Then MSFX.Row = MSFX.Row + 1
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
            .SetFocus
          End With
          On Error GoTo 0
          
           Case 39 '(39derecha)
            On Error Resume Next
            'MSFX.SetFocus
            MSFX.TEXT = TXTMOVILX.TEXT
            If MSFX.Cols - 1 > MSFX.COL Then
              TXTMOVILX.ToolTipText = REG1& & "," & COLU1& + 1
'              MSFX.COL = MSFX.COL + 1
            End If

            If MSFX.COL < MSFX.Cols Then MSFX.COL = MSFX.COL + 1

          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
'            .ToolTipText = MSFX.Row
            .SetFocus
          End With
          On Error GoTo 0
           
           Case 37 '(37izquierda)
            On Error Resume Next
            'MSFX.SetFocus
            MSFX.TEXT = TXTMOVILX.TEXT

            If MSFX.COL > 1 Then
              TXTMOVILX.ToolTipText = REG1& & "," & COLU1& - 1
              MSFX.COL = MSFX.COL - 1
            End If
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
'            .ToolTipText = MSFX.Row
            .SetFocus
          End With
          On Error GoTo 0

    End Select
    TXTSEPARADOR = ";"
    '



End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
        Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
        On Error Resume Next
        REG& = XVALO(CADENATEX$(1))
        COLU1& = XVALO(CADENATEX$(2))
        MSF.TextMatrix(REG&, COLU1&) = TXTMOVIL
        TXTMOVIL.Visible = False
     End If
End Sub

'Private Sub TXTNOMBRELISTA_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXTCMB(Me.CMBNOMBRECAMPO)
'  End If
'
'End Sub


'Private Sub TXTORDPRES_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXT(Me.TXTFORMATO)
'  End If
'
'End Sub


'Private Sub TXTTIPOID_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXT(TXTNOMBRELISTA)
'  End If
'
'End Sub


Private Sub TXTMOVIL2_Change()
'     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0

End Sub


Private Sub TXTMOVIL2_DblClick()
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COL1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If REG1& = 3 Then ' CAMPO ECO
        ARMACONDI.APROBO% = 0
        'LEVANTO LA CONDICION SI YA ESTABA GUARDADA.
        ARMACONDI.Text7 = MSF4.TextMatrix(REG1&, COL1&)
        ARMACONDI.Show 1
'        MSFESPEJO.TextMatrix(REG1&, COL1&) = ""
        If ARMACONDI.APROBO% > 0 Then
          'COPIA LA CONDICION EN UNA GRILLA OCULTA
           MSF5.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
        End If
     End If

End Sub


Private Sub TXTMOVIL2_GotFocus()
     Call PINTATEXT(TXTMOVIL2)
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     COLU1& = XVALO(CADENATEX$(2))
     REG1& = XVALO(CADENATEX$(1))
     On Error GoTo 0

End Sub


Private Sub TXTMOVIL2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
'     If REG& = 1 Then REG& = 0
     MSF4.TextMatrix(REG&, COLU1&) = TXTMOVIL2.TEXT
     MSF4.COL = COLU1&
     TEPRUEBA.Caption = TXTMOVIL2.TEXT
     If TEPRUEBA.Width > MSF4.ColWidth(COLU1&) Then
       MSF4.ColWidth(COLU1&) = TEPRUEBA.Width + 300
     End If
     On Error GoTo 0
     '
     
     If MSF4.Cols - 1 > COLU1& Then
       COLU1& = COLU1& + 1
     Else
       COLU1& = 1
     End If
50   MSF4.COL = COLU1&
     With TXTMOVIL2
        .ZOrder 0
        .Width = MSF4.CellWidth
        .Left = MSF4.CellLeft + MSF4.Left
        .Top = MSF4.CellTop + MSF4.Top
        .Visible = True
        .ToolTipText = MSF4.Row & "," & COLU1&
        
        'ESTO SE PARA QUE CUANDO VAYA MOVIENDOSE TOME EL TEXTO INGRESADO EN EL ENCABEZADO
         Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
         On Error Resume Next
         REG& = XVALO(CADENATEX$(1))
         COLU1& = XVALO(CADENATEX$(2))
         TXTMOVIL2.TEXT = MSF4.TextMatrix(REG&, COLU1&)
        .SetFocus
     End With
 
   End If

End Sub


Private Sub TXTMOVIL2_LostFocus()
    If TXTMOVIL2.Visible = True Then
        Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
        On Error Resume Next
        REG1& = XVALO(CADENATEX$(1))
        COLU1& = XVALO(CADENATEX$(2))
        On Error GoTo 0
        MSF4.TextMatrix(REG1&, COLU1&) = TXTMOVIL2
        TXTMOVIL2.Visible = False
    End If
End Sub



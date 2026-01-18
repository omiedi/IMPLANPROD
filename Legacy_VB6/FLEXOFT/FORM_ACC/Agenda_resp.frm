VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Agenda 
   Caption         =   "Agenda de Trabajos"
   ClientHeight    =   9765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   20400
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   20205
      _ExtentX        =   35639
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Agenda Original"
      TabPicture(0)   =   "Agenda.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Informe de Tareas"
      TabPicture(1)   =   "Agenda.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   8895
         Index           =   0
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   19815
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   6645
            Left            =   14280
            ScaleHeight     =   6585
            ScaleWidth      =   5025
            TabIndex        =   79
            Top             =   2160
            Width           =   5085
            Begin VB.Image Image1 
               Height          =   4935
               Left            =   480
               Picture         =   "Agenda.frx":0038
               Stretch         =   -1  'True
               Top             =   2040
               Width           =   4215
            End
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Borrar"
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
            Left            =   15960
            Picture         =   "Agenda.frx":6425
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Baja Lógica de la Cuenta"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox DTHORARIO 
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
            Height          =   330
            Left            =   7830
            TabIndex        =   73
            Top             =   1590
            Width           =   1785
         End
         Begin VB.ComboBox CMBMES 
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
            Left            =   11880
            TabIndex        =   72
            Text            =   "CMBMES"
            Top             =   1620
            Width           =   2010
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   11880
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Text            =   "A Partir de ???"
            ToolTipText     =   "A Partir del Mes ???"
            Top             =   1320
            Width           =   2025
         End
         Begin VB.TextBox TXTORDEN 
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   1590
            Width           =   465
         End
         Begin VB.CommandButton Command5 
            Height          =   885
            Left            =   14280
            Picture         =   "Agenda.frx":672F
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Text            =   "Cod.Trabajo"
            Top             =   1320
            Width           =   1155
         End
         Begin VB.TextBox TXTTELEFONO 
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
            Left            =   9975
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   240
            Width           =   4000
         End
         Begin VB.TextBox TXTCELULAR 
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
            Left            =   9975
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   525
            Width           =   4000
         End
         Begin VB.TextBox TXTMAIL 
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
            Left            =   9975
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   810
            Width           =   4000
         End
         Begin VB.TextBox TXTLOCALIDAD 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   810
            Width           =   5325
         End
         Begin VB.TextBox TXTDIRECCION 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   525
            Width           =   5325
         End
         Begin VB.TextBox TXTNOMBRE 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   240
            Width           =   5325
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
            Left            =   240
            TabIndex        =   61
            Top             =   240
            Width           =   975
         End
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
            Left            =   1230
            TabIndex        =   60
            ToolTipText     =   "Seleccion de Cliente"
            Top             =   255
            Width           =   250
         End
         Begin VB.TextBox Text40 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   7830
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Text            =   "Horario"
            Top             =   1320
            Width           =   1785
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "Semana"
            ToolTipText     =   "Orden de la Semana"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox TXTCANTICUOTAS 
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
            Height          =   330
            Left            =   10560
            TabIndex        =   57
            Top             =   1590
            Width           =   1305
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   10560
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "Cant. Cuotas "
            Top             =   1320
            Width           =   1305
         End
         Begin VB.TextBox TXTREPITE 
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
            Height          =   330
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1590
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
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
            Height          =   330
            Left            =   600
            TabIndex        =   54
            Top             =   1590
            Width           =   1155
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "Día"
            Top             =   1320
            Width           =   1290
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Text            =   "Rep.Cada"
            ToolTipText     =   "Repetir Trabajo Cada (Meses)"
            Top             =   1320
            Width           =   975
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
            Height          =   330
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   1590
            Width           =   3735
         End
         Begin VB.ComboBox CMBORDEN 
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
            Left            =   7020
            TabIndex        =   49
            Text            =   "1"
            Top             =   1620
            Width           =   790
         End
         Begin VB.ComboBox cmbdias 
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
            Left            =   5720
            TabIndex        =   48
            Top             =   1620
            Width           =   1290
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   1
            Left            =   1760
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "Nombre:"
            Top             =   240
            Width           =   1400
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   1
            Left            =   1760
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "Dirección:"
            Top             =   525
            Width           =   1400
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   1
            Left            =   1760
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "Localidad:"
            Top             =   810
            Width           =   1400
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   2
            Left            =   8595
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "Teléfono:"
            Top             =   240
            Width           =   1400
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   2
            Left            =   8595
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "Celular:"
            Top             =   525
            Width           =   1400
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
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
            Height          =   285
            Index           =   2
            Left            =   8595
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "Mail:"
            Top             =   810
            Width           =   1400
         End
         Begin VB.CommandButton Command1 
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
            Height          =   600
            Left            =   1755
            TabIndex        =   41
            ToolTipText     =   "Seleccion de Tipo de Trabajo"
            Top             =   1320
            Width           =   250
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Grabar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   17520
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Guarda los Cambios Realizados"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
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
            Height          =   380
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "Ord."
            Top             =   1320
            Width           =   465
         End
         Begin VB.TextBox Text20 
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
            Height          =   330
            Left            =   11880
            TabIndex        =   38
            Top             =   1560
            Width           =   2025
         End
         Begin MSFlexGridLib.MSFlexGrid MSF1 
            Height          =   6675
            Left            =   120
            TabIndex        =   75
            Top             =   2160
            Width           =   13800
            _ExtentX        =   24342
            _ExtentY        =   11774
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
            Height          =   195
            Left            =   14400
            TabIndex        =   78
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label TEPRUEBA 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            Height          =   225
            Left            =   14400
            TabIndex        =   77
            Top             =   360
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label26 
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
            Height          =   300
            Left            =   240
            TabIndex        =   76
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8895
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   19995
         Begin VB.CommandButton Command13 
            Caption         =   "Generar Lista de Trabajos a Facturar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   18000
            TabIndex        =   80
            Top             =   960
            Width           =   1815
         End
         Begin VB.Frame Frame5 
            Height          =   855
            Left            =   14520
            TabIndex        =   33
            Top             =   720
            Width           =   1815
            Begin VB.CheckBox CHKSINEJECUTAR 
               Caption         =   "Pendiente de Ejecución"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Actualizar Grilla_"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   14520
            TabIndex        =   32
            Top             =   1630
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   31
            Text            =   "AGENDA DE TRABAJO"
            Top             =   240
            Width           =   14295
         End
         Begin VB.Frame Frame2 
            Caption         =   "Filtrar Por: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1400
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   6375
            Begin VB.TextBox Text10 
               Appearance      =   0  'Flat
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
               Height          =   330
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "Cliente:"
               Top             =   240
               Width           =   1395
            End
            Begin VB.TextBox TXTCUENTA 
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
               Height          =   330
               Left            =   1560
               TabIndex        =   29
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox TXTNOMBRE2 
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
               Left            =   2775
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   240
               Width           =   3495
            End
            Begin VB.CommandButton Command3 
               Caption         =   "."
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
               Left            =   2500
               TabIndex        =   27
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   240
               Width           =   250
            End
            Begin VB.TextBox Text14 
               Appearance      =   0  'Flat
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
               Height          =   330
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   26
               TabStop         =   0   'False
               Text            =   "Trabajo:"
               Top             =   600
               Width           =   1395
            End
            Begin VB.TextBox TXTCODTRABAJO 
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
               Height          =   330
               Left            =   1560
               TabIndex        =   25
               Top             =   600
               Width           =   915
            End
            Begin VB.TextBox TXTDESTRABAJO 
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
               Left            =   2775
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   600
               Width           =   3495
            End
            Begin VB.CommandButton Command6 
               Caption         =   "."
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
               Left            =   2505
               TabIndex        =   23
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   600
               Width           =   250
            End
            Begin VB.TextBox Text21 
               Appearance      =   0  'Flat
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
               Height          =   330
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "Responsable:"
               Top             =   960
               Width           =   1395
            End
            Begin VB.TextBox TXTCODRESPONSABLE 
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
               Height          =   330
               Left            =   1560
               TabIndex        =   21
               Top             =   960
               Width           =   915
            End
            Begin VB.TextBox TXTDESRESPONSABLE 
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
               Left            =   2775
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   960
               Width           =   3495
            End
            Begin VB.CommandButton Command7 
               Caption         =   "."
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
               Left            =   2505
               TabIndex        =   19
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   960
               Width           =   250
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Filtros x Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1400
            Left            =   6600
            TabIndex        =   7
            Top             =   720
            Width           =   4695
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
               Left            =   840
               TabIndex        =   17
               ToolTipText     =   "Click Para Seleccionar Sección"
               Top             =   240
               Width           =   175
            End
            Begin VB.TextBox TXTDESDE 
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
               Height          =   330
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   240
               Width           =   1290
            End
            Begin VB.CommandButton Command9 
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
               Left            =   840
               TabIndex        =   15
               ToolTipText     =   "Click Para Seleccionar Sección"
               Top             =   600
               Width           =   175
            End
            Begin VB.TextBox TXTHASTA 
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
               Height          =   330
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   600
               Width           =   1290
            End
            Begin VB.TextBox Text15 
               Appearance      =   0  'Flat
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
               Height          =   330
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "Hasta:"
               Top             =   600
               Width           =   675
            End
            Begin VB.TextBox Text18 
               Appearance      =   0  'Flat
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
               Height          =   330
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   12
               TabStop         =   0   'False
               Text            =   "Desde:"
               Top             =   240
               Width           =   675
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Fecha Orignal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   0
               Left            =   2640
               TabIndex        =   11
               Top             =   120
               Width           =   1935
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Fecha Organizada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   1
               Left            =   2640
               TabIndex        =   10
               Top             =   400
               Width           =   1935
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Fecha Realizado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   2
               Left            =   2640
               TabIndex        =   9
               Top             =   680
               Width           =   1815
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Quitar Selección"
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
               Left            =   2700
               TabIndex        =   8
               Top             =   1020
               Width           =   1935
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Otros Filtros:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1400
            Left            =   11400
            TabIndex        =   3
            Top             =   720
            Width           =   3015
            Begin VB.OptionButton OPTSOLOFACTURADO 
               Caption         =   "Solo Facturado"
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
               Left            =   240
               TabIndex        =   6
               Top             =   240
               Width           =   2415
            End
            Begin VB.OptionButton OPTPENDIENTEFACTURA 
               Caption         =   "Pendiente de Facturación"
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
               Left            =   240
               TabIndex        =   5
               Top             =   600
               Width           =   2700
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Quitar Selección"
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
               Left            =   240
               TabIndex        =   4
               Top             =   1020
               Width           =   2535
            End
         End
         Begin VB.CommandButton Command10 
            Caption         =   "FILTRAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   14520
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid MSF2 
            Height          =   6435
            Left            =   120
            TabIndex        =   35
            Top             =   2280
            Width           =   19680
            _ExtentX        =   34713
            _ExtentY        =   11351
            _Version        =   393216
            BackColorFixed  =   16501405
            SelectionMode   =   1
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   45
         End
      End
   End
   Begin VB.Menu Arcvivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu config 
      Caption         =   "Configuración"
      Begin VB.Menu TsgTrabajo 
         Caption         =   "Tabla de Tiempos según Trabajo"
      End
      Begin VB.Menu tpersonal 
         Caption         =   "Tabla de Personal"
      End
      Begin VB.Menu ttrabajo 
         Caption         =   "Tablas de Tipo de Trabajo"
      End
      Begin VB.Menu configuracion 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaestrodeClientes 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu masteritem 
         Caption         =   "Maestro de Tipo de Trabajos"
      End
      Begin VB.Menu Ventas 
         Caption         =   "Gestión Ventas y Deudores"
         Begin VB.Menu Fac_ven 
            Caption         =   "Facturacion en Cuenta Corriente"
         End
         Begin VB.Menu Recibo1 
            Caption         =   "Recibo de Cobranza"
         End
      End
      Begin VB.Menu ped_clientes 
         Caption         =   "Sub-Sistema de Pedidos de Clientes"
      End
   End
End
Attribute VB_Name = "Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cadena = Format("06/08/78", "General Date") ' Devuelve: "06/08/1978"
'
'cadena = Format("19/08/79", "Long Date") ' Devuelve : "Jueves 19 de Agosto de 1979".
'
'cadena = Format("19/8/79", "Medium Date") ' Devuelve: "19-Ago-1979"
'
'cadena = Format("17:08", "Short Time") ' "05:08"
'
'cadena = Format("17:08", "Medium Time") ' "05:08 PM"
'
'cadena = Format("17:08", "Long Time") ' "05:08:00 PM"
'http://www.recursosvisualbasic.com.ar/htm/tutoriales/tutorial-basico5.htm

Function ANOXCON20$(MESANOX1$)
    ANOXCON20$ = MESANOX1$
    MESANOX$ = MESANOX1$
    POS1% = InStr(TRIM$(MESANOX$), " ")
    If POS1% > 0 Then
      MESANOX$ = Mid$(MESANOX$, POS1% + 1)
      If XVALO(MESANOX$) > 0 And XVALO(MESANOX$) < 100 Then
        MESANOX$ = TRIM$(Str$(2000 + XVALO(MESANOX$)))
      End If
    End If
    ANOXCON20$ = MESANOX$
End Function

Sub CARGAMES()
    CMBMES.Clear
    CMBMES.AddItem "ENERO"
    CMBMES.AddItem "FEBRERO"
    CMBMES.AddItem "MARZO"
    CMBMES.AddItem "ABRIL"
    CMBMES.AddItem "MAYO"
    CMBMES.AddItem "JUNIO"
    CMBMES.AddItem "JULIO"
    CMBMES.AddItem "AGOSTO"
    CMBMES.AddItem "SEPTIEMBRE"
    CMBMES.AddItem "OCTUBRE"
    CMBMES.AddItem "NOVIEMBRE"
    CMBMES.AddItem "DICIEMBRE"
    
    CMBMES.TEXT = METEXTO$(HOY(HO$))
End Sub
Sub CARGAR_GRILLA_TRABAJO()

    ORD% = 0
    MSF2.Rows = 1
    Dim rs1 As DAO.Recordset
    SQLX$ = " SELECT * FROM AGENDA "
    SQLX$ = SQLX$ + " where facturado <> '3' "
    
    SQLX$ = SQLX$ + " ORDER BY Dia_Organizado "
    Set rs1 = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    With rs1
      Do While Not .EOF
        ORD% = ORD% + 1
        MSF2.Rows = ORD% + 1
        MSF2.TextMatrix(ORD%, 1) = XVALO(!NUME_CLI)
        MSF2.TextMatrix(ORD%, 2) = TRIM$(Left$(RASOCLI$(XVALO(!NUME_CLI)), 24)) + " - (" + TRIM$(Left$(DOMICLI$(XVALO(!NUME_CLI)), 24)) + ")"
        MSF2.TextMatrix(ORD%, 3) = CODEXT$(XVALO(!CODTRABAJO))
        MSF2.TextMatrix(ORD%, 4) = DESCRIT0$(XVALO(!CODTRABAJO))
        MSF2.TextMatrix(ORD%, 5) = FECHATEX$(CVINT(!Fecha_Original))
        MSF2.TextMatrix(ORD%, 6) = !OrdenSemana
            HORORIG$ = TRIM$(!Horario_Orig)
            If HORORIG$ = "12:00:00 p.m." Then HORORIG$ = ""
        MSF2.TextMatrix(ORD%, 7) = HORORIG$
        MSF2.TextMatrix(ORD%, 8) = !CONTINUIDAD
        MSF2.TextMatrix(ORD%, 9) = !CUOTAS
          DIAORGAN$ = TRIM$(FECHATEX$(CVINT(!Dia_Organizado)))
          If DIAORGAN$ = "" Then
             DIAORGAN$ = FECHATEX$(CVINT(!Fecha_Original))
          End If
        MSF2.TextMatrix(ORD%, 10) = DIAORGAN$
          HOROORGAN$ = TRIM$(!Horario_Organizado)
          If HOROORGAN$ = "12:00:00 p.m." Then
             HOROORGAN$ = ""
          ElseIf HOROORGAN$ = "" Then
             HOROORGAN$ = HORORIG$
          End If
        MSF2.TextMatrix(ORD%, 11) = HOROORGAN$
        MSF2.TextMatrix(ORD%, 12) = FECHATEX$(CVINT(!Dia_Realizado))
        MSF2.TextMatrix(ORD%, 13) = ECOARCH$("PERSONAL", MKS$(XVALO(!OPER_REALIZO)))
        FACTU$ = TRIM$(!FACTURADO)
        If FACTU$ <> "SÍ" Then
            If XVALO(FACTU$) = 0 Then
               FACTU$ = ""
            Else
               FACTU$ = "SÍ"
            End If
        End If
        MSF2.TextMatrix(ORD%, 14) = FACTU$
        MSF2.TextMatrix(ORD%, 15) = !Mesano_Orig
        MSF2.TextMatrix(ORD%, 16) = !ID
        MSF2.TextMatrix(ORD%, 17) = XVALO(!OPER_REALIZO)
        .MoveNext
      Loop
    End With

End Sub

Sub FILTRAR_GRILLA_TRABAJO()

    ORD% = 0
    MSF2.Rows = 1
    Dim rs1 As DAO.Recordset
    SQLX$ = " SELECT * FROM AGENDA "
    SQLX$ = SQLX$ + " WHERE NUME_CLI > 0 "
    '
    'FILTROS POR CLIENTE, TIPOTRABAJO Y RESPONSABLE.
    
    If XVALO(TXTCUENTA) > 0 Then SQLX$ = SQLX$ + " AND NUME_CLI = " & XVALO(TXTCUENTA)
    If CODINT%(TXTCODTRABAJO) > 0 Then SQLX$ = SQLX$ + " AND CodTrabajo = " & CODINT%(TXTCODTRABAJO)
    If XVALO(TXTCODRESPONSABLE) > 0 Then SQLX$ = SQLX$ + " AND OPER_REALIZO = " & XVALO(TXTCODRESPONSABLE)
    '
    'TRATAMIENTO DE FILTRO POR FECHAS VALIDACIONES PARA VER POR QUE FECHA SE FILTRA SI POR LA FECHA ORIGINAL , ORGANIZADA O REALIZADO
    'SI TIENE FECHA DESDE Y NO HASTA REALIZA EL DESDE HASTA EL FINAL, SI TIENE HASTA Y NO DESDE REALIZA DESDE EL COMIENZO HASTA EL HASTA.
    If CVINT(TXTDESDE) > 0 Or CVINT(TXTHASTA) > 0 Then
      Dim TILDE As Boolean 'valida si hay filtro por fecha que haya elegido una opcion.
      For I% = 0 To Option1.UBound
        TILDE = Option1(I%).Value
        If TILDE = True Then Exit For
      Next I%
      If TILDE = False Then
        Call COMUNI("Para Filtrar Por Fecha debe Seleccionar alguna de las 3 Opciones")
      Else
        If CVINT(TXTDESDE) > 0 Then
          FECHADESDE = CDbl(CVDAT(CVINT(TXTDESDE)))
          If Option1(0) = True Then SQLX$ = SQLX$ + " AND Fecha_Original >= " & FECHADESDE & ""
          If Option1(1) = True Then SQLX$ = SQLX$ + " AND Dia_Organizado >= " & FECHADESDE & ""
          If Option1(2) = True Then SQLX$ = SQLX$ + " AND Dia_Realizado >= " & FECHADESDE & ""
        End If
        If CVINT(TXTHASTA) > 0 Then
          FECHAHASTA = CDbl(CVDAT(CVINT(TXTHASTA)))
          If Option1(0) = True Then SQLX$ = SQLX$ + " AND Fecha_Original <= " & FECHAHASTA & ""
          If Option1(1) = True Then SQLX$ = SQLX$ + " AND Dia_Organizado <=" & FECHAHASTA & ""
          If Option1(2) = True Then SQLX$ = SQLX$ + " AND Dia_Realizado <= " & FECHAHASTA & ""
        End If
      End If
    End If
    '
    'OTROS FILTROS
    If Me.OPTPENDIENTEFACTURA = True Then SQLX$ = SQLX$ + " AND FACTURADO = ''"
    '
    If Me.CHKSINEJECUTAR = 1 Then
      FECHAFACTU = CDbl(CVDAT(CVINT("01/01/14")))
      SQLX$ = SQLX$ + " AND Dia_Realizado    >= " & FECHAFACTU & ""
    End If
    '
    If Me.OPTSOLOFACTURADO = True Then
       SQLX$ = SQLX$ + " AND FACTURADO = '3'"
    Else
       SQLX$ = SQLX$ + " AND FACTURADO <> '3'"
    End If
    
    
    SQLX$ = SQLX$ + " ORDER BY Dia_Organizado "
    Set rs1 = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    With rs1
      Do While Not .EOF
        ORD% = ORD% + 1
        MSF2.Rows = ORD% + 1
        MSF2.TextMatrix(ORD%, 1) = XVALO(!NUME_CLI)
        MSF2.TextMatrix(ORD%, 2) = TRIM$(Left$(RASOCLI$(XVALO(!NUME_CLI)), 24)) + "(" + TRIM$(Left$(DOMICLI$(XVALO(!NUME_CLI)), 24)) + ")"
        MSF2.TextMatrix(ORD%, 3) = CODEXT$(XVALO(!CODTRABAJO))
        MSF2.TextMatrix(ORD%, 4) = DESCRIT0$(XVALO(!CODTRABAJO))
        MSF2.TextMatrix(ORD%, 5) = FECHATEX$(CVINT(!Fecha_Original))
        MSF2.TextMatrix(ORD%, 6) = !OrdenSemana
            HORORIG$ = TRIM$(!Horario_Orig)
            If HORORIG$ = "12:00:00 p.m." Then HORORIG$ = ""
        MSF2.TextMatrix(ORD%, 7) = HORORIG$
        MSF2.TextMatrix(ORD%, 8) = !CONTINUIDAD
        MSF2.TextMatrix(ORD%, 9) = !CUOTAS
          DIAORGAN$ = TRIM$(FECHATEX$(CVINT(!Dia_Organizado)))
          If DIAORGAN$ = "" Then
             DIAORGAN$ = FECHATEX$(CVINT(!Fecha_Original))
          End If
        MSF2.TextMatrix(ORD%, 10) = DIAORGAN$
          HOROORGAN$ = TRIM$(!Horario_Organizado)
          If HOROORGAN$ = "12:00:00 p.m." Then
             HOROORGAN$ = ""
          ElseIf HOROORGAN$ = "" Then
             HOROORGAN$ = HORORIG$
          End If
        MSF2.TextMatrix(ORD%, 11) = HOROORGAN$
        MSF2.TextMatrix(ORD%, 12) = FECHATEX$(CVINT(!Dia_Realizado))
        MSF2.TextMatrix(ORD%, 13) = ECOARCH$("PERSONAL", MKS$(XVALO(!OPER_REALIZO)))
        FACTU$ = TRIM$(!FACTURADO)
        If FACTU$ <> "SÍ" Then
            If XVALO(FACTU$) = 0 Then
               FACTU$ = ""
            Else
               FACTU$ = "En Espera"
            End If
        End If
        '
        MSF2.TextMatrix(ORD%, 14) = FACTU$
        MSF2.TextMatrix(ORD%, 15) = !Mesano_Orig
        MSF2.TextMatrix(ORD%, 16) = !ID
'        MSF2.TextMatrix(ORD%, 17) = IDDIA%
        .MoveNext
      Loop
    End With

End Sub


Function CONTINUIDAD%(CODTRABAJO$)
   CI11% = CODINT%(CODTRABAJO$)
   CONTINUIDAD% = 1
   For I& = 1 To ULTREG&("CONTINUI")
      x$ = REGLEIDO$("CONTINUI", I&)
      CIXI2% = CVI(Mid(x$, 1, 2))
      If CI11% = CIXI2% Then
         CONTINUIDAD% = CVI(Mid(x$, 35, 2))
         Exit For
      End If
   Next I&
End Function

Function MESNUM$(MESANO1$)
   'DEVUELVE SI SE LE PASA EL MES EL NUMERO DE MES EN FORMATO MM
   'SE HACE EL INSTRING POR SI SE LE PASA EN FORMATO ABRIL 15
   POS1% = InStr(TRIM$(MESANO1$), " ")
   MESX$ = TRIM$(Mid$(MESANO1$, 1, POS1% - 1))
   If MESX$ = "" Then Exit Function
   '
   Select Case MESX$
    Case "ENERO"
      MESNUM$ = "01"
    Case "FEBRERO"
      MESNUM$ = "02"
    Case "MARZO"
      MESNUM$ = "03"
    Case "ABRIL"
      MESNUM$ = "04"
    Case "MAYO"
      MESNUM$ = "05"
    Case "JUNIO"
      MESNUM$ = "06"
    Case "JULIO"
      MESNUM$ = "07"
    Case "AGOSTO"
      MESNUM$ = "08"
    Case "SETIEMBRE"
      MESNUM$ = "09"
    Case "OCTUBRE"
      MESNUM$ = "10"
    Case "NOVIEMBRE"
      MESNUM$ = "11"
    Case "DICIEMBRE"
      MESNUM$ = "12"
   End Select
   
   
End Function

Sub ORDENA_COLUMNA_FECHA(MSF As MSFlexGrid, COLUM&, MODO As Boolean)
    '1 PASA DE FECHA A FORMATO NUMERICO PARA QUE SE PUEDA ORDENAR
    For I& = 1 To MSF.Rows - 1
       MSF.TextMatrix(I&, COLUM&) = CVINT(MSF.TextMatrix(I&, COLUM&))
    Next I&
    '
    If MODO Then
        MSF2.COL = COLUM&
        MSF2.Sort = 2
    ' Ordena en forma descendente
    Else
        MSF2.COL = COLUM&
        MSF2.Sort = 1
    End If

    '2 PASA DE FORMATO NUMERO FECHA PARA QUE QUEDE COMO ESTABA
    For I& = 1 To MSF.Rows - 1
       MSF.TextMatrix(I&, COLUM&) = FECHATEX$(XVALO((MSF.TextMatrix(I&, COLUM&))))
    Next I&

End Sub

Private Sub cmbdias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub CMBMES_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub CMBORDEN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
'   Call SELECHO("TIPOTRAB")
'   CODTRA% = XVALO(VALACT1$("TIPOTRAB"))
'   If CODTRA% < 1 Then GoTo 99
'   '
'   txtcodigo = CODTRA%
   Call SELECHO("MASTER")
   CODTRA% = CODINT%(VALACT1$("MASTER"))
   If CODTRA% < 1 Then GoTo 99
   '
   txtcodigo = VALACT1$("MASTER")
   
99 Command1.Enabled = True

End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call FILTRAR_GRILLA_TRABAJO
99 Command10.Enabled = True

End Sub

Private Sub Command11_Click()
  Option1(0).Value = False
  Option1(1).Value = False
  Option1(2).Value = False
End Sub

Private Sub Command12_Click()
   Me.OPTSOLOFACTURADO.Value = False
   Me.OPTPENDIENTEFACTURA.Value = False
   
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False

20  CAMPOS$ = "F.Realiz./D8;Cta./F6;Nombre-Domicilio/A50;codigo TRabajo/f0;Desc.Trabajo/A32;Cod.Operario/f0;Realizo/A32;Cuotas/F8;ID/F0"
    Call CARGA_ENCABEZADO(MUESTRA04.GRID, CAMPOS$, MUESTRA04, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(MUESTRA04.MSF_2, CAMPOS$, MUESTRA04, 1, 0, ANTOT, , 1)
    MUESTRA04.Command3.Caption = "Selección de Trabajos a Facturar"
    SQLX$ = " SELECT * FROM AGENDA "
    SQLX$ = SQLX$ + " WHERE NUME_CLI > 0 "
    SQLX$ = SQLX$ + " and  FACTURADO = '1'" 'facturado van a aser las que se desean facturar"
    '
    Dim RS As DAO.Recordset

    Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)

 
    j& = 0
    With RS
      Do While Not .EOF
        j& = j& + 1
        Call TRACE(20, j&, FIN&)
        FE% = CVINT(!Dia_Realizado)
        NC = XVALO(!NUME_CLI)
        RASYDOMI$ = TRIM$(Left$(RASOCLI$(NC), 24)) + " (" + TRIM$(Left$(DOMICLI$(NC), 24)) + ")"
        CODITRABAJO = XVALO(!CODTRABAJO)
        'DESCRITRABAJO$ = TRIM$(ECOARCH$("TIPOTRAB", MKI$(XVALO(!CODTRABAJO))))
        DESCRITRABAJO$ = CODEXT$(XVALO(!CODTRABAJO))
        OPERARIO% = XVALO(!OPER_REALIZO)
        NAMEOPERARIO$ = ECOARCH$("PERSONAL", MKS$(XVALO(!OPER_REALIZO)))
        
        MUESTRA04.GRID.Rows = j& + 1
        MUESTRA04.GRID.TextMatrix(j&, 1) = FECHATEX$(FE%)
        MUESTRA04.GRID.TextMatrix(j&, 2) = NC
        MUESTRA04.GRID.TextMatrix(j&, 3) = RASYDOMI$
        MUESTRA04.GRID.TextMatrix(j&, 4) = CODITRABAJO
        MUESTRA04.GRID.TextMatrix(j&, 5) = DESCRITRABAJO$
        MUESTRA04.GRID.TextMatrix(j&, 6) = OPERARIO%
        MUESTRA04.GRID.TextMatrix(j&, 7) = NAMEOPERARIO$
        MUESTRA04.GRID.TextMatrix(j&, 8) = XVALO(!CUOTAS)
        MUESTRA04.GRID.TextMatrix(j&, 9) = XVALO(!ID)
        
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing

    MUESTRA04.Show 1
    Command13.Enabled = True

End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call SELECHO("DEUDOR1")
   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE < 1 Then
     GoTo 99
   End If
   '
   Text1.TEXT = TRIM$(Str$(NCLIE))
99 Screen.MousePointer = 1
   Command19.Enabled = True
End Sub



Private Sub Command2_Click()
   Command2.Enabled = False
   NCLIE = XVALO(Text1)
   '
    Dim RS As DAO.Recordset
    For I& = 1 To MSF1.Rows - 1
      If XVALO(MSF1.TextMatrix(I&, 10)) < 1 Then ' si ya existe lo edita si no addnew
        SQLX$ = "SELECT * FROM  AGENDAORIGINAL  "
        SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NCLIE

        Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
        RS.AddNew
      End If
      If XVALO(MSF1.TextMatrix(I&, 10)) > 0 Then '
        SQLX$ = "SELECT * FROM  AGENDAORIGINAL  "
        SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NCLIE

        SQLX$ = SQLX$ + " AND ID = " & XVALO(MSF1.TextMatrix(I&, 10))
        Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
        
        RS.Edit
      End If
      With RS
          AD = !ID
         !NUME_CLI = XVALO(Text1)
         !ORDEN = MSF1.TextMatrix(I&, 1)
    '     !CODTRABAJO = XVALO(MSF1.TextMatrix(I&, 2))
         !CODTRABAJO = CODINT%(MSF1.TextMatrix(I&, 2))
         !DIA = MSF1.TextMatrix(I&, 4)
         !ID_DIA = IDDIA%(MSF1.TextMatrix(I&, 4))
         !OrdenSemana = XVALO(MSF1.TextMatrix(I&, 5))
         !Horario = MSF1.TextMatrix(I&, 6)
         !CONTINUIDAD = XVALO(MSF1.TextMatrix(I&, 7))
         !CUOTAS = XVALO(MSF1.TextMatrix(I&, 8))
         !MARBORRA = 0
         !NUME_CLI = NCLIE
         !APARTIRDE = MSF1.TextMatrix(I&, 9)
         .Update
     End With
     RS.Close: Set RS = Nothing
    Next I&
   
'   SQLX$ = "DELETE * FROM AGENDAORIGINAL  WHERE MARBORRA = 1 AND  Nume_Cli = " & NCLIE
'   Badaclie.Execute (SQLX$)

99 Command2.Enabled = True

End Sub
Function IDDIA%(DIAX$)
   DIAX1$ = TRIM$(UCase$(DIAX$))
   Select Case DIAX1$
     Case "LUNES"
       IDDIA% = 1
     Case "MARTES"
       IDDIA% = 2
     Case "MIERCOLES"
       IDDIA% = 3
     Case "MIÉRCOLES"
       IDDIA% = 3
     Case "JUEVES"
       IDDIA% = 4
     Case "VIERNES"
       IDDIA% = 5
     Case "SABADO"
       IDDIA% = 6
     Case "SÁBADO"
       IDDIA% = 6
     Case "DOMINGO"
       IDDIA% = 0
     Case Else
       Call COMUNI("Día No Válido (" & DIAX1$ & ")")
   End Select
   
End Function


Private Sub Command24_Click()
    Command24.Enabled = False
    If MSF1.Rows < 2 Then GoTo 99
    OPX% = COMALTER%("Realmente Desea Eliminar Este Trabajo de la Agenda Original\\Cancelar\Sí, Eliminar")
    'BORRO
    If OPX% = 2 Then
        SQLX$ = "DELETE * FROM AgendaOriginal  WHERE ID = " & XVALO(Label2.Caption)
        Badaclie.Execute (SQLX$)
        Call borrafila(XVALO(TXTORDEN), Agenda, MSF1)
    End If
    'REFRESCO LA GRILLA
    Agenda.CARGAR_GRILLA_TRABAJO
99  Command24.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("DEUDOR1")
   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE < 1 Then
     GoTo 99
   End If
   TXTCUENTA = NCLIE
99 Command3.Enabled = True
   
''    Command3.Enabled = False
''    '
''    If XVALO(txtcodigo2) < 1 Then
''      Call COMUNI("Falta Código de Trabajo")
''      GoTo 99
''    End If
''    '
''    If TRIM$(TXTDESCRIPCION2) = "" Then
''      Call COMUNI("Código de Trabajo No Válido")
''      GoTo 99
''    End If
''    '
''    If TRIM$(cmbdias2) = "" Then
''      Call COMUNI("Falta Seleccionar el Día")
''      GoTo 99
''    End If
''
''    If TRIM$(CMBORDEN2) = "" Then
''      Call COMUNI("Falta Seleccionar Orden del Día")
''      GoTo 99
''    End If
''
''    If TRIM$(DTHORARIO2) = "" Then
''      Call COMUNI("Falta Seleccionar Horario")
''      GoTo 99
''    End If
''
''    If XVALO(TXTREPITE2) = 0 Then
''      Call COMUNI("Falta Seleccionar la Continuidad")
''      GoTo 99
''    End If
''
''    If XVALO(TXTCANTICUOTAS2) = 0 Then
''      Call COMUNI("Falta Seleccionar la Cantidad de Cuotas")
''      GoTo 99
''    End If
''
''    ORD% = XVALO(TXTORDEN2)
''    If ORD% > 0 Then
''       ORD% = ORD%
''    Else
''       ORD% = MSF2.Rows
''       MSF2.Rows = MSF2.Rows + 1
''    End If
''    MSF2.TextMatrix(ORD%, 1) = ORD%
''    MSF2.TextMatrix(ORD%, 2) = txtcodigo.TEXT
''    MSF2.TextMatrix(ORD%, 3) = TXTDESCRIPCION.TEXT
''    MSF2.TextMatrix(ORD%, 4) = cmbdias.TEXT
''    MSF2.TextMatrix(ORD%, 5) = CMBORDEN.TEXT
''    MSF2.TextMatrix(ORD%, 6) = DTHORARIO.TEXT
''    MSF2.TextMatrix(ORD%, 7) = TXTREPITE
''    MSF2.TextMatrix(ORD%, 8) = TXTCANTICUOTAS
''
''    Call LIMPIA_EDICION
''
''99  Command3.Enabled = True
End Sub


Private Sub Command4_Click()
    Command4.Enabled = False
10  Call CARSELMESA2(12)
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
    Else
      GoTo 99
    End If
    '
    MESNUMERO_SELECCIONADO% = NUMMES%(MESASEL$)
    SQLX$ = "SELECT  COUNT(*) AS CANTI FROM AGENDA"
    SQLX$ = SQLX$ + " WHERE MESANO_ORIG = '" & MESASEL$ & "'"
    SQLX$ = SQLX$ + " AND FACTURADO <>'3'"
    Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    If RS!CANTI > 0 Then
       OPX% = COMALTER%("Atención!!\Se Ha Detectado que yá se han Importado Datos a la Agenda Activa\\Cancelar\Actualizar Agenda Activa")
       If OPX% <= 1 Then GoTo 10
    End If
    RS.Close
    Set RS = Nothing
    '
    'RECORRO LA TABLA AGENDAORIGINAL Y VEO SI EXISTE EL MISMO EN LA TABLA AGENDA
    'TENER EN  CUENTA SI EXISTE FILA CON INFO DE LA ORIGINAL EN LA TABLA AGENDA NO SE COPIA NI SE ACTUALIZA
    'SI NO EXISTE AGREGA EL REGISTRO
    'SI EL REGISTRO EN LA ORIGINAL SE ELIMINO Y SE VOLVIO A CARGAR CON LA MISMA INFO LO VA A AGREGAR EN LA TABLA AGENDA
    'POR QUE VA A CAMBIAR EL IDENTIFICADOR.
    
    SQLX$ = " SELECT * FROM AGENDAORIGINAL "
    SQLX$ = SQLX$ + " ORDER BY ID_DIA"
    Set rs1 = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    With rs1
      Do While Not .EOF
        IDX% = XVALO(!ID)
        CONTIN% = XVALO(!CONTINUIDAD)
        APART% = NUMMES%(SAFETEXT$(!APARTIRDE))
        'ENTRA EN ESTOS 3 CASOS
        'SI LA CONTIN% = 1 POR QUE ES PARA TODOS LOS MESES
        'SI EL MES GRABADO 'A PARTIR DE' ES IGUAL AL MES SELECCIONADO
        'SI EL MES SELECCIONADO ES IGUAL AL MES 'A PARTIR DE' + LA CONTINUIDAD (SI SELECCIONA FEBRERO Y APARTIR DE ES DICIEMBRE Y LA CONTINUIDAD ES 2 INGRESA)
        '(SI SELECCIONA FEBRERO Y APARTIR DE ES DICIEMBRE Y LA CONTINUIDAD ES 2 INGRESA)
        APARTMASCONTIN% = APART% + CONTIN% 'CALCULO SI LA SUMA NO SUPERA EL MES 12 CASO NOVIEMBRE = 11 + CONTINUIDAD = 6 RESULTADO 17 THEN -12 = 5
        If APARTMASCONTIN% > 12 Then APARTMASCONTIN% = APARTMASCONTIN% - 12
        If CONTIN% = 1 Or APART% = MESNUMERO_SELECCIONADO% Or APARTMASCONTIN% = MESNUMERO_SELECCIONADO% Then
            CONDI$ = "ID_ORIGINAL = " & IDX%
            CONDI$ = CONDI$ + " AND MESANO_ORIG = '" & MESASEL$ & "'"
            SQLX$ = "SELECT  COUNT(*) AS CANTI FROM AGENDA"
            SQLX$ = SQLX$ + " WHERE " & CONDI$
            Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
            If RS!CANTI < 1 Then
    
                SQLX$ = " SELECT * FROM AGENDA "
                SQLX$ = SQLX$ + "WHERE CodTrabajo < 1 "
                Set rs2 = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
                With rs2
                    .AddNew
                    !CODTRABAJO = rs1!CODTRABAJO
                    AA = FECHATEX$(UBICAR_DIA_SEMANA%(rs1!OrdenSemana, rs1!DIA, MESNUM$(MESASEL$), ANOXCON20$(MESASEL$)))
                    !Fecha_Original = FECHATEX$(UBICAR_DIA_SEMANA%(rs1!OrdenSemana, rs1!DIA, MESNUM$(MESASEL$), ANOXCON20$(MESASEL$)))
                    !OrdenSemana = rs1!OrdenSemana
                    !Horario_Orig = rs1!Horario
                    !CONTINUIDAD = rs1!CONTINUIDAD
                    !CUOTAS = rs1!CUOTAS
                    !NUME_CLI = rs1!NUME_CLI
                    !MARBORRA = 0
                    !CODTRABAJO = rs1!CODTRABAJO
                    !Dia_Organizado = CVDAT(0)
                    !Horario_Organizado = ""
                    !Dia_Realizado = CVDAT(0)
                    !Horario_Realizado = ""
                    !Mesano_Orig = MESASEL$
                    !ID_ORIGINAL = IDX%
                    !FACTURADO = ""
'                    !OPER_REALIZO = txtcodusuario
'                    !REALIZADO = Me.TXT
'                    !REALIZADO = CVINT(TXTREALIZADO)
                    .Update
                End With
            End If
        End If
        .MoveNext
        
      Loop
    End With
    Call CARGAR_GRILLA_TRABAJO
    
99  Command4.Enabled = True
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "REFRESCAR LA AGENDA " & vbCrLf
       TXTO$ = TXTO$ + "LA MISMA ACTUALIZA EL MES A SELECCIONAR" + vbCrLf
       TXTO$ = TXTO$ + "SI YA HA SIDO MIGRADO ALGUN TRABAJO ESTE SE VA A MANTENER" + vbCrLf
       TXTO$ = TXTO$ + "(NO ASI SI SE CREO MANUALMENTE)" + vbCrLf
       TXTO$ = TXTO$ + "ESTO QUIERE DECIR QUE SI SE CREO MANUALMENTE EL MISMO SE VA A DUPLICAR" + vbCrLf
       TXTO$ = TXTO$ + "EN TODO CASO BORRAR MANUALMENTE LA SEGUNDA." + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub Command5_Click()
    Command5.Enabled = False
    '
    If XVALO(txtcodigo) < 1 Then
      Call COMUNI("Falta Código de Trabajo")
      GoTo 99
    End If
    '
    If TRIM$(TXTDESCRIPCION) = "" Then
      Call COMUNI("Código de Trabajo No Válido")
      GoTo 99
    End If
    '
    If TRIM$(cmbdias) = "" Then
      Call COMUNI("Falta Seleccionar el Día")
      GoTo 99
    End If
    
    If TRIM$(CMBORDEN) = "" Then
      Call COMUNI("Falta Seleccionar Orden del Día")
      GoTo 99
    End If
    
    If TRIM$(DTHORARIO) = "" Then
      Call COMUNI("Falta Seleccionar Horario")
      GoTo 99
    End If
    
    If XVALO(TXTREPITE) = 0 Then
      Call COMUNI("Falta Seleccionar la Continuidad")
      GoTo 99
    End If
    
    If XVALO(TXTCANTICUOTAS) = 0 Then
      Call COMUNI("Falta Seleccionar la Cantidad de Cuotas")
      GoTo 99
    End If
    If XVALO(TXTREPITE) > 1 Then
      If TRIM$(CMBMES.TEXT) = "" Then
        Call COMUNI("Debe Seleccionar el 'Mes A Partir de'")
        CMBMES.SetFocus
        GoTo 99
      End If
    End If

    ORD% = XVALO(TXTORDEN)
    If ORD% > 0 Then 'si es edicion
       ORD% = ORD%
    Else
       ORD% = MSF1.Rows 'si esta agregando
       MSF1.Rows = MSF1.Rows + 1
    End If
    MSF1.TextMatrix(ORD%, 1) = ORD%
    MSF1.TextMatrix(ORD%, 2) = txtcodigo.TEXT
    MSF1.TextMatrix(ORD%, 3) = TXTDESCRIPCION.TEXT
    MSF1.TextMatrix(ORD%, 4) = cmbdias.TEXT
    MSF1.TextMatrix(ORD%, 5) = CMBORDEN.TEXT
    MSF1.TextMatrix(ORD%, 6) = DTHORARIO.TEXT
    MSF1.TextMatrix(ORD%, 7) = TXTREPITE
    MSF1.TextMatrix(ORD%, 8) = TXTCANTICUOTAS
    MSF1.TextMatrix(ORD%, 9) = CMBMES.TEXT
    MSF1.TextMatrix(I&, 10) = XVALO(Label2.Caption)
    Call LIMPIA_EDICION
   
99  Command5.Enabled = True
End Sub
Sub LIMPIA_EDICION()
    TXTORDEN.TEXT = ""
    txtcodigo.TEXT = ""
    TXTDESCRIPCION.TEXT = ""
    cmbdias.TEXT = ""
    CMBORDEN.TEXT = ""
    DTHORARIO = ""
    TXTREPITE.TEXT = ""
    TXTCANTICUOTAS.TEXT = ""
    Label2 = ""
End Sub

Private Sub Command5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtcodigo.SetFocus
    End If

End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
'   Call SELECHO("TIPOTRAB")
'   CODTRA% = XVALO(VALACT1$("TIPOTRAB"))
'   If CODTRA% < 1 Then GoTo 99
'   '
'   TXTCODTRABAJO = CODTRA%
   Call SELECHO("MASTER")
   CODTRA% = XVALO(VALACT1$("MASTER"))
   If CODTRA% < 1 Then GoTo 99
   TXTCODTRABAJO = CODTRA%
   '
99 Command6.Enabled = True

End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
    Me.TXTDESDE = VDEV$
   End If
99 Command8.Enabled = True

End Sub


Private Sub Command7_Click()
   Command7.Enabled = False
   Call SELECHO("PERSONAL")
   CODPER% = XVALO(VALACT1$("PERSONAL"))
   If CODPER% < 1 Then GoTo 99
   '
   Me.TXTCODRESPONSABLE = CODPER%
   
99 Command7.Enabled = True

End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
    Me.TXTHASTA = VDEV$
   End If
99 Command9.Enabled = True

End Sub

Private Sub configuracion_Click()
   FSETUPNEW.Show 1
   Call FINAL("")
End Sub

Private Sub DTHORARIO_Change()
  X1 = ERRTIEMPO%(DTHORARIO)
  
End Sub

Private Sub DTHORARIO_GotFocus()
  DTHORARIO.SelStart = 0
  DTHORARIO.SelLength = Len(DTHORARIO)
End Sub


Private Sub DTHORARIO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       On Error Resume Next
       SendKeys "{tab}"
       On Error GoTo 0
    End If

End Sub

Private Sub DTHORARIO_LostFocus()
   DTHORARIO = HORATEX$(HORANUM(DTHORARIO))
   Exit Sub
End Sub


Private Sub DTPicker1_Click()
NUMX% = CVINT(DTPicker1.Value)
TXTCANTICUOTAS2 = FECHATEX$(NUMX%)
End Sub

Private Sub Fac_ven_Click()
   Call CONECRUN("VENGES04/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)

   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   'VALIDACION PARA CIELO DESTAPACIONES POR SI NO PAGA
    FF% = FECHANUM("30/08/15")
    If FF% <= HOY(HO$) Then
      Call Form_Load
    End If

   openforms = DoEvents
   AA = DIASEM%(HOY(HO$))
   VV = METEXTO$(HOY(HO$)) ' ABRIL 2015
   BB = DSEM$(HOY(HO$)) ' LUNES
   'XX = UBICAR_DIA_SEMANA%(5, "MIERCOLES", "06", "2015")
   AA = FECHATEX$(XX)
   Call CARGACOMBO_DIAS
   Call CARGACOMBO_ORDEN
   Call CARGAMES
   Call CARGA_ENCABEZADO(Me.MSF1, "Ord/F5;Cod.Trab/F8;Descripción/A24;Día/A12;Semana/A8;Horario/A12;Repite Cada/A10;Cuotas/F7;A Partir de/A12;ID/F0", Agenda, 0, 0, 0, 0)
   
   Call CARGA_ENCABEZADO(Me.MSF2, "Cta./F5;R.Social-Domicilio/A48;Cod./F5;Descripción/A22;F.Orig./A10;N°/A0;Hor.Orig./A10;Cont.Mens./A0;Cuotas/F0;F.Organ./A10;Hor.Organ./A10;Realizado/A10;Realiza;Facturar ?/A8;Mesano_Orig/A0;ID/F0:IDDIA/F0;UsuarioRealizo/F0", Agenda, 0, 0, 0, 0)
'   1 Cta./F5;
'   2 R.Social/A20;
'   3 Cod./F5;
'   4 Descripción/A22;
'   5 F.Orig./A10;
'   6 N°/A0;
'   7 Hor.Orig./A10;
'   8 Cont.Mens./A0;
'   9 Cuotas/F0;
'   10F.Organ./A10;
'   11 Hor.Organ./A10;
'   12 Realizado/A10;
'   13 Realizó;
'   14 Fact./A8;
'   15 Mesano_Orig/A0;
'   16 ID/F0
'   17 IDDIA/F0
   Call ABRECLIEN
   Call GENERBASE("BADACLIEN")
End Sub
Sub CARGACOMBO_DIAS()
   cmbdias.AddItem "Lunes"
   cmbdias.AddItem "Martes"
   cmbdias.AddItem "Miércoles"
   cmbdias.AddItem "Jueves"
   cmbdias.AddItem "Viernes"
   cmbdias.AddItem "Sábado"
   cmbdias.AddItem "Domingo"
   
End Sub
Sub CARGACOMBO_ORDEN()
   CMBORDEN.AddItem "1°"
   CMBORDEN.AddItem "2°"
   CMBORDEN.AddItem "3°"
   CMBORDEN.AddItem "4°"
   CMBORDEN.AddItem "Último"
End Sub
Sub CARGACOMBO_HORARIO()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")
End Sub

Private Sub MaestrodeClientes_Click()
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
    Call FRESHECHO("DEUDOR1")
End Sub

Private Sub masteritem_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub mnufacturacion_Click()
   Call CONECRUN("VENGES04/FACTURA", "Gestión de Ventas y Deudores")
End Sub

Private Sub MSF1_DblClick()
    I& = MSF1.MouseRow
    j& = MSF1.MouseCol
    If I& = 0 Or I& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = I&
    MSF1.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If
    '

    TXTORDEN.TEXT = MSF1.TextMatrix(REG&, 1)
    txtcodigo = MSF1.TextMatrix(REG&, 2)
    TXTDESCRIPCION = MSF1.TextMatrix(REG&, 3)
    cmbdias = MSF1.TextMatrix(REG&, 4)
    CMBORDEN = MSF1.TextMatrix(REG&, 5)
    DTHORARIO = MSF1.TextMatrix(REG&, 6)
    TXTREPITE = MSF1.TextMatrix(REG&, 7)
    TXTCANTICUOTAS = MSF1.TextMatrix(REG&, 8)
    CMBMES.TEXT = MSF1.TextMatrix(REG&, 9)
    Label2 = MSF1.TextMatrix(REG&, 10)
    '
    Call SELECCIONARFILA(MSF1, REG&)
    '
End Sub

Private Sub MSF2_Click()
    Static MODO  As Boolean
    If (MSF2.MouseRow = 0) Then
        Select Case MSF2.MouseCol
        Case 5, 10, 12
          Call ORDENA_COLUMNA_FECHA(MSF2, MSF2.MouseCol, MODO)
          If MODO Then
            MODO = False
          Else
            MODO = True
          End If
          Exit Sub
        End Select
        ' Ordena en forma ascendente
        If MODO Then
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_DblClick()
    I& = MSF2.MouseRow
    j& = MSF2.MouseCol
    If I& = 0 Or I& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = I&
    MSF2.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If
    '
   With EDITAGENDA
        .TXTCUENTA = MSF2.TextMatrix(REG&, 1)
        .TXTRSOCIAL = MSF2.TextMatrix(REG&, 2)
        .txtcodigo2 = MSF2.TextMatrix(REG&, 3)
        .TXTDESCRIPCION = MSF2.TextMatrix(REG&, 4)
        .TXTFECHAORIGINAL = MSF2.TextMatrix(REG&, 5)
        .TXTHorOrig = MSF2.TextMatrix(REG&, 7)
        .TXTSEMANA = MSF2.TextMatrix(REG&, 6)
        .TXTCONTINUIDAD = MSF2.TextMatrix(REG&, 8)
        .TXTCUOTAS = MSF2.TextMatrix(REG&, 9)
        '.TXTFORGANIZADA = MSF1.TextMatrix(REG&, 8)
        .TXTFORGANIZADA = MSF2.TextMatrix(REG&, 10)
        .TXTHORORGANIZADO = MSF2.TextMatrix(REG&, 11)
        .TXTREALIZADO = MSF2.TextMatrix(REG&, 12)
        .TXTREALIZO = MSF2.TextMatrix(REG&, 13)
        .TXTFACTURADO = MSF2.TextMatrix(REG&, 14)
        .TXTID = XVALO(MSF2.TextMatrix(REG&, 16))
        .txtcodusuario = MSF2.TextMatrix(REG&, 17)

        .Show 1
       
   End With
'   1 Cta./F5;
'   2 R.Social/A20;
'   3 Cod./F5;
'   4 Descripción/A22;
'   5 F.Orig./A10;
'   6 N°/A0;
'   7 Hor.Orig./A10;
'   8 Semana/A4;
'   8 Cont.Mens./A0;
'   9 Cuotas/F0;
'   10 F.Organ./A10;
'   11 Hor.Organ./A10;
'   12 Realizado/A10;
'   13 Realizó;
'   14 Fact./A8;
'   15 Mesano_Orig/A0;
'   16 ID/F0
'   17 IDDIA/F0
End Sub




Private Sub ped_clientes_Click()

  If EXISTE%("PEDIDO04.EXE") > 0 Then
      Call CONECRUN("PEDIDO04", "Gestión de Pedidos de Clientes")
      Exit Sub
   End If

End Sub

Private Sub Recibo_Click()
   If EXISTE%("VENGES04.EXE") > 0 Then
     ACONEC$ = "VENGES04/RECIBO"
     Call CONECRUN(ACONEC$, "Recibo de Cobranza")
   Else
     Call COMUNI("Falta Módulo VENGES04")
   End If
End Sub
Private Sub Recibo1_Click()
   ' Realiza una llamada a la opcion de menu
   ' Tesoreria/Cobranzas y Pagos/Recibo de Cobranza
   Call Recibo_Click
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab <> PreviousTab Then
  If SSTab1.Caption = "Agenda Activa" Then
    CARGAR_GRILLA_TRABAJO
  End If
End If

End Sub

Private Sub Text1_Change()
   Call BLANFORMU
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE = XVALO(Text1.TEXT)
       If STATCLI(NCLIE) > 0 Then
         Call cardatclie
         Call CARGA_AGENDAORIGINAL
       End If
     End If
   End If
   Screen.MousePointer = 1
End Sub
    '
Sub CARGA_AGENDAORIGINAL()
     NCLIE = XVALO(Text1)
     Call ABRECLIEN
     SQLX$ = "SELECT * FROM AGENDAORIGINAL"
     SQLX$ = SQLX$ + " WHERE NUME_CLI = " & NCLIE
     Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
     I& = 0
     With RS
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
                I& = I& + 1
                MSF1.Rows = I& + 1
                MSF1.TextMatrix(I&, 1) = !ORDEN
                MSF1.TextMatrix(I&, 2) = CODEXT$(XVALO(!CODTRABAJO))
                MSF1.TextMatrix(I&, 3) = DESCRIT0$(XVALO(!CODTRABAJO))
                MSF1.TextMatrix(I&, 4) = !DIA
                MSF1.TextMatrix(I&, 5) = !OrdenSemana
                MSF1.TextMatrix(I&, 6) = !Horario
                  CONTIN% = XVALO(!CONTINUIDAD)
                  CONTINUIX$ = ""
                  If CONTIN% > 1 Then CONTINUIX$ = SAFETEXT$(CONTIN%) & " Meses"
                  If CONTIN% = 1 Then CONTINUIX$ = SAFETEXT$(CONTIN%) & " Mes"
                MSF1.TextMatrix(I&, 7) = CONTINUIX$
                MSF1.TextMatrix(I&, 8) = XVALO(!CUOTAS)
                MSF1.TextMatrix(I&, 9) = !APARTIRDE
                MSF1.TextMatrix(I&, 10) = XVALO(!ID)
           .MoveNext
         Loop
       End If
     End With
     RS.Close
     Set RS = Nothing
End Sub

Sub cardatclie()
    Screen.MousePointer = 11
    NCLIE = XVALO(Text1)
    TXTNOMBRE.TEXT = RASOCLI$(NCLIE)
    TXTDIRECCION.TEXT = DOMICLI$(NCLIE)
    TXTLOCALIDAD.TEXT = LOCACLI$(NCLIE)
    TXTTELEFONO.TEXT = TEL1CLI$(NCLIE)
    'TXTDIRECCION.TEXT = TEL2CLI$(NCLIE)
    TXTMAIL.TEXT = EMAILCLI$(NCLIE)
    
End Sub
Sub BLANFORMU()
   MSF1.Rows = 1
   Label26 = ""
    TXTNOMBRE.TEXT = ""
    TXTDIRECCION.TEXT = ""
    TXTLOCALIDAD.TEXT = ""
    TXTTELEFONO.TEXT = ""
    TXTCELULAR.TEXT = ""
    TXTMAIL.TEXT = ""
    
    txtcodigo.TEXT = ""
    TXTDESCRIPCION.TEXT = ""
    TXTREPITE.TEXT = ""
    TXTCANTICUOTAS.TEXT = ""
    DTHORARIO.TEXT = ""
    CMBORDEN = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       On Error Resume Next
       SendKeys "{tab}"
       On Error GoTo 0
    End If

End Sub

Private Sub Text10_DblClick()
  If TRIM$(Text10) = "" Then
     Text10 = "SÍ"
  Else
     Text10 = ""
  End If
     
  
  
End Sub


Private Sub Text27_Change()

End Sub


Private Sub Text13_Change()

End Sub

Private Sub tpersonal_Click()
    Call TRALOCAL("Personal", "")
10  VTB% = VENTABU%("Personal")
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!Personal")
        JJ& = JJ& + 1
        x$ = REGLEIDO$("!Personal", U&)
        CODI& = CVS(Left$(x$, 4))
        DESCRI$ = TRIM$(Mid$(x$, 5, 40))
        If CODI& > 0 Then
          If DESCRI$ = "" Then
            Call COMUNI("Falta Descripción")
            GoTo 10
          End If
        Else
          If DESCRI$ <> "" Then
            Call COMUNI("Falta Código")
            GoTo 10
          End If
        End If
      Next U&
      Call SETULTREG("Personal", JJ&)
    End If
    Call TRACENTRAL("Personal", "")
    Call CIERRARCH("Personal")
    Call CIERRARCH("!Personal")

End Sub


Private Sub TsgTrabajo_Click()
    Call TRALOCAL("CONTINUI", "")
10  VTB% = VENTABU%("CONTINUI")
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!CONTINUI")
        JJ& = JJ& + 1
        x$ = REGLEIDO$("!CONTINUI", U&)
        CODI& = CVI(Left$(x$, 2))
        MESES1% = CVI(Mid$(x$, 35, 2))
        AA = CVI(Mid(x$, 35, 2))
        If CODI& > 0 Then
          If MESES1% < 1 Then
            Call COMUNI("Falta Cantidad de Meses")
            GoTo 10
          End If
        Else
          If MESES1% > 0 Then
            Call COMUNI("Falta Código")
            GoTo 10
          End If
        End If
      Next U&
      Call SETULTREG("CONTINUI", JJ&)
    End If
    Call TRACENTRAL("CONTINUI", "")
    Call CIERRARCH("CONTINUI")
    Call CIERRARCH("!CONTINUI")

End Sub


Private Sub ttrabajo_Click()
    Call TRALOCAL("TIPOTRAB", "")
10  VTB% = VENTABU%("TIPOTRAB")
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!TIPOTRAB")
        JJ& = JJ& + 1
        x$ = REGLEIDO$("!TIPOTRAB", U&)
        CODI% = CVI(Left$(x$, 2))
        DESCRI$ = TRIM$(Mid$(x$, 3, 48))
        If CODI% > 0 Then
          If DESCRI$ = "" Then
            Call COMUNI("Falta Descripción")
            GoTo 10
          End If
        Else
          If DESCRI$ <> "" Then
            Call COMUNI("Falta Código")
            GoTo 10
          End If
        End If
      Next U&
      Call SETULTREG("TIPOTRAB", JJ&)
    End If
    Call TRACENTRAL("TIPOTRAB", "")
    
    Call CIERRARCH("TIPOTRAB")
    Call CIERRARCH("!TIPOTRAB")

End Sub


Private Sub TXTCANTICUOTAS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       On Error Resume Next
       SendKeys "{tab}"
       On Error GoTo 0
    End If

End Sub


Private Sub txtcodigo_Change()
   TXTDESCRIPCION = ""
   TXTDESCRIPCION = DESCRIT0$(CODINT%(txtcodigo))
   CONTINU% = CONTINUIDAD%(txtcodigo)
   'ESCRIBE MES SI ES 1 , MESES SI SON MAS DE 1 O NADA SI ES < 1
   If CONTINU% = 1 Then
     MESX$ = "Mes"
   ElseIf CONTINU% > 1 Then
     MESX$ = "Meses"
   Else
     MESX$ = ""
   End If
   
   If MESX$ <> "" Then
     TXTREPITE = CONTINU% & " " & MESX$
   Else
     TXTREPITE = ""
   End If
End Sub


Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys "{tab}"
       cmbdias.SetFocus
    End If
End Sub


'Private Sub txtcodigo2_Change()
'   TXTDESCRIPCION2 = ""
'   TXTDESCRIPCION2 = DESCRIT0$(CODINT%(txtcodigo2))
'   CONTINU% = CONTINUIDAD%(CODINT%(txtcodigo2))
'   'ESCRIBE MES SI ES 1 , MESES SI SON MAS DE 1 O NADA SI ES < 1
'   If CONTINU% = 1 Then
'     MESX$ = "Mes"
'   ElseIf CONTINU% > 1 Then
'     MESX$ = "Meses"
'   Else
'     MESX$ = ""
'   End If
'
'   If MESX$ <> "" Then
'     TXTREPITE2 = CONTINU% & " " & MESX$
'   Else
'     TXTREPITE2 = ""
'   End If
'
'End Sub


Private Sub txtcodusuario_Change()
   TXTNOMBREUSUARIO = ""
   TXTNOMBREUSUARIO = ECOARCH$("PERSONAL", MKS$(XVALO(txtcodusuario)))
End Sub


Public Function ERRTIEMPO%(TXT$)
    '
    ERRTIEMPO% = 0
    If XVALO(TXT$) > 24 Then
      Call COMUNI("Horas No Puede ser Mayor a 24")
      ERRTIEMPO% = 1
      Exit Function
    End If
    '
    INGRESO$ = TRIM$(TXT$)
    INGRESO$ = REPLACAR(INGRESO$, ",", ":")
    INGRESO$ = REPLACAR(INGRESO$, ".", ":")
    '
    For I% = 1 To Len(TRIM$(INGRESO$))
       POS1% = InStr(INGRESO$, ":")
       
       If I% = 1 Then MAXI1% = 24
       If I% = 2 Then MAXI1% = 59: If CONTROL("PADIAMOD", "CONFHORA") = "MINUCENT" Then MAXI1% = 99 ' configuracion minutos centecimal
       If I% = 3 Then MAXI1% = 99
       '
       If POS1% > 0 Then
            HOR1 = XVALO(Mid$(INGRESO$, 1, POS1% - 1))
            If HOR1 > MAXI1% Then
              Call COMUNI("Horario Ingresado Fuera de Rango - Verifique Formato por Defecto")
              ERRTIEMPO% = 1
              Exit For
            End If
            INGRESO$ = Mid$(INGRESO$, POS1% + 1)
       Else
            HOR1 = XVALO(INGRESO$)
            If HOR1 > MAXI1% Then
              Call COMUNI("Horario Ingresado Fuera de Rango - Verifique Formato por Defecto")
              ERRTIEMPO% = 1
              Exit For
            End If
       End If
       '
    Next I%
    '
End Function

Private Sub TXTCODRESPONSABLE_Change()
   TXTDESRESPONSABLE = ""
   TXTDESRESPONSABLE = ECOARCH$("PERSONAL", MKS$(XVALO(TXTCODRESPONSABLE)))


End Sub

Private Sub TXTCODTRABAJO_Change()
   TXTDESTRABAJO = ""
   TXTDESTRABAJO = CODEXT$(XVALO(TXTCODTRABAJO))

End Sub

Private Sub TXTCUENTA_Change()
    NCLIE = XVALO(TXTCUENTA)
    TXTNOMBRE2.TEXT = RASOCLI$(NCLIE)
End Sub

Private Sub TXTDESDE_DblClick()
   TXTDESDE = ""

End Sub


Private Sub TXTHASTA_DblClick()
       TXTHASTA = ""

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   I& = MSF2.MouseRow
   j& = MSF2.MouseCol
   If I& = 0 Or I& > MSF2.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       MSF2.Row = I&
       MSF2.COL = j&
       MSF2.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
       OPX% = COMALTER%("Desea Generar Ahora La Impresión del Aviso de Trabajo\\Cancelar\Imprimir Ahora")
       If OPX% <= 1 Then GoTo 30
       NC = MSF2.TextMatrix(I&, 1)
       DIAX1$ = MSF2.TextMatrix(I&, 10)
       HORAX$ = MSF2.TextMatrix(I&, 11)
       FechaTrabajo$ = DIAX1$ & " a las " & Left$(HORAX$, 5)
       Call ImprimirAviso(NC, HOY(HO$), FechaTrabajo$)
30    MSF2.CellBackColor = RGB(255, 255, 255)
   End If
   
End Sub

Sub ImprimirAviso(NC, FE%, FechaTrabajo$)
   FORIPRE$ = "AVIS-DDC"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        Exit Sub
      Else
        Screen.MousePointer = 11
        FECHDOC$ = FE%
        NUMEDOC$ = ""
        TIPODOC$ = "Aviso Mantenimiento"
        '
         
        CODPARAM$(1) = "LAR-CUTE": DOCPARAM$(1) = FECHATEXLAR$(FE%)
        CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = RASOCLI$(NC)
        CODPARAM$(3) = "FECH-VEN": DOCPARAM$(3) = FechaTrabajo$ '(FechaTrabajo)
        CAPARDOC% = 3
        '
        Call PRINTDOC("@" & FORIPRE$)
        Screen.MousePointer = 1
        '
   End If
End Sub

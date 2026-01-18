VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Agenda17 
   Caption         =   "Agenda de Trabajos"
   ClientHeight    =   9765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   0
      TabIndex        =   8
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
      TabPicture(0)   =   "Agenda17.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Informe de Tareas"
      TabPicture(1)   =   "Agenda17.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   8895
         Index           =   0
         Left            =   -74880
         TabIndex        =   45
         Top             =   480
         Width           =   19815
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
            Left            =   5700
            TabIndex        =   2
            Top             =   1620
            Width           =   1280
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   6405
            Left            =   14280
            ScaleHeight     =   6345
            ScaleWidth      =   5025
            TabIndex        =   79
            Top             =   2400
            Width           =   5085
            Begin VB.Image Image1 
               Height          =   6400
               Left            =   0
               Picture         =   "Agenda17.frx":0038
               Stretch         =   -1  'True
               Top             =   0
               Width           =   4965
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
            Picture         =   "Agenda17.frx":25616
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
            TabIndex        =   4
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
            TabIndex        =   6
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
            TabIndex        =   73
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
            TabIndex        =   72
            Top             =   1590
            Width           =   465
         End
         Begin VB.CommandButton Command5 
            Height          =   885
            Left            =   14280
            Picture         =   "Agenda17.frx":25920
            Style           =   1  'Graphical
            TabIndex        =   7
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   0
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
            TabIndex        =   64
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   5
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   1
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   3
            Top             =   1620
            Width           =   790
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Label3"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   960
            Visible         =   0   'False
            Width           =   720
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
         TabIndex        =   9
         Top             =   360
         Width           =   19995
         Begin VB.CommandButton Command14 
            Caption         =   "Exportar Todo a Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   16800
            TabIndex        =   88
            Top             =   240
            Width           =   3015
         End
         Begin MSFlexGridLib.MSFlexGrid MSF3 
            Height          =   2775
            Left            =   3240
            TabIndex        =   86
            Top             =   3720
            Visible         =   0   'False
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4895
            _Version        =   393216
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Exportar a Excel Trabajos"
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
            Left            =   16800
            TabIndex        =   85
            Top             =   840
            Visible         =   0   'False
            Width           =   3015
         End
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
            Height          =   735
            Left            =   16800
            TabIndex        =   80
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Frame Frame5 
            Height          =   855
            Left            =   14520
            TabIndex        =   41
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
               TabIndex        =   42
               ToolTipText     =   "Que No Tiene Fecha de Realización"
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
            TabIndex        =   40
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
            Left            =   6600
            TabIndex        =   39
            Text            =   "AGENDA DE TRABAJO"
            Top             =   240
            Width           =   7815
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
            Height          =   1875
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   6375
            Begin VB.CommandButton Command15 
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
               TabIndex        =   84
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   360
               Width           =   250
            End
            Begin VB.TextBox TXTREFGRUPO 
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
               TabIndex        =   83
               Top             =   360
               Width           =   3495
            End
            Begin VB.TextBox TXTGRUPO 
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
               TabIndex        =   82
               Top             =   360
               Width           =   915
            End
            Begin VB.TextBox Text13 
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
               TabIndex        =   81
               TabStop         =   0   'False
               Text            =   "Administración:"
               Top             =   360
               Width           =   1395
            End
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
               TabIndex        =   38
               TabStop         =   0   'False
               Text            =   "Cliente:"
               Top             =   720
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
               TabIndex        =   37
               Top             =   720
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
               TabIndex        =   36
               Top             =   720
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
               TabIndex        =   35
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   720
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
               TabIndex        =   34
               TabStop         =   0   'False
               Text            =   "Trabajo:"
               Top             =   1080
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
               TabIndex        =   33
               Top             =   1080
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
               TabIndex        =   32
               Top             =   1080
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
               TabIndex        =   31
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   1080
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
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "Responsable:"
               Top             =   1440
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
               TabIndex        =   29
               Top             =   1440
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
               TabIndex        =   28
               Top             =   1440
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
               TabIndex        =   27
               ToolTipText     =   "Seleccion de Tipo de Trabajo"
               Top             =   1440
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
            TabIndex        =   15
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
               TabIndex        =   25
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
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
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
               TabIndex        =   18
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
               TabIndex        =   17
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
               TabIndex        =   16
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
            TabIndex        =   11
            Top             =   720
            Width           =   3015
            Begin VB.OptionButton OPTSOLOFACTURADO 
               Caption         =   "Solo Aprobado P/Facturar"
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
               TabIndex        =   14
               Top             =   240
               Width           =   2655
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
               TabIndex        =   13
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
               TabIndex        =   12
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
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid MSF2 
            Height          =   6435
            Left            =   120
            TabIndex        =   43
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
            TabIndex        =   44
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
      Begin VB.Menu Amapro 
         Caption         =   "Maestro de Proveedores"
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
      Begin VB.Menu mnuLp 
         Caption         =   "Lista de Precios"
      End
      Begin VB.Menu mnupresup 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu mnuproveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu Fac_provee 
            Caption         =   "Carga de Factura de Proveedores"
         End
         Begin VB.Menu Orpago 
            Caption         =   "Orden de Pago"
         End
         Begin VB.Menu subdiarioyestadisticas 
            Caption         =   "Subdiario y Estadísticas"
         End
      End
      Begin VB.Menu Dfrm 
         Caption         =   "Diseñador de Formularios"
      End
      Begin VB.Menu Genrep 
         Caption         =   "Generador de Reportes"
      End
   End
End
Attribute VB_Name = "Agenda17"
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
Private Sub Amapro_Click()
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     ACONEC$ = "ARCHIG07/AMAPRO"
     Call CONECRUN(ACONEC$, "Maestro de Proveedores y Sub-Contratistas")
   Else
     Call COMUNI("Falta Módulo ARCHIG07")
   End If
'   Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
End Sub


Private Sub Command16_Click()
   Command6.Enabled = False
   A = SSTab1.Caption
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call ARMAR_MSF3
      Call GENEXCEL2(AREXPO$, Me.MSF3)
   End If
99 Command6.Enabled = True
End Sub
Sub ARMAR_MSF3()

'   Call CARGA_ENCABEZADO(Me.MSF2, "Cta./F5 1 ;R.Social-Domicilio/A48 2 ;Cod./F5 3 ;Descripción/A22 4 ;F.Orig./A10 5 ;N°/A0 6 ;Hor.Orig./A10 7 ;Cont.Mens./A0 8 ;Cuotas/F0 9 ;F.Organ./A10 10 ;Hor.Organ./A10 11 ;Realizado/A10 12 ;Realiza 13 ;Facturar ?/A8 14 ;Mesano_Orig/A0 15 ;ID/F0 16 :IDDIA/F0 17 ;UsuarioRealizo/F0 18", Agenda, 0, 0, 0, 0)
'F.Organ (lunes  2-NOV  07:45(horario organizado) ) la fecha 10
'tanque      4
'DIRECCION DE LA TAREA  domicli(1)
'PABLO   13

'PASO LA INFO A UN MSF TEMPORAL
   MSF3.Cols = 7
   For i& = 1 To MSF2.Rows - 1
      MSF3.Rows = i& + 1
      FE% = FECHANUM(MSF2.TextMatrix(i&, 10))
      MSF3.TextMatrix(i&, 1) = DSEM(FE%)
        '16 de AGOSTO de 2013
        DIAX$ = SAFETEXT$(XVALO(Mid$(MSF2.TextMatrix(i&, 10), 1, 2)))
        MEX$ = Left$(METEXTO$(FE%), 3)
        DIAXMEX$ = DIAX$ & "-" & MEX$
        HOX2$ = MSF2.TextMatrix(i&, 11)
          POS1% = InStrRev(HOX2$, ":")
          
      'MSF3.TextMatrix(I&, 1) = "=TEXTO(B" & I & ";" & Chr(34) & "dddd" & Chr(34) & ")" '=TEXTO(B2;"dddd")
      MSF3.TextMatrix(i&, 2) = MSF2.TextMatrix(i&, 10) 'DIAXMEX$ 'FECHA ORGANIZ
      MSF3.TextMatrix(i&, 3) = Mid$(HOX2$, 1, POS1% - 1) ' MSF2.TextMatrix(I&, 11) 'HORARIO HORGANIZADO
      MSF3.TextMatrix(i&, 4) = MSF2.TextMatrix(i&, 4) 'tanque      4
      MSF3.TextMatrix(i&, 5) = DOMICLI$(XVALO(MSF2.TextMatrix(i&, 1))) 'DIRECCION DE LA TAREA  domicli(1)
      MSF3.TextMatrix(i&, 6) = MSF2.TextMatrix(i&, 13) 'REALIZA
   Next i&
   'FALTA LA COLUMNA 1 Y LA DE LOS USUARIOS Y OTRA MAS.
End Sub

Private Sub Dfrm_Click()
   Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
End Sub

Private Sub Fac_provee_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES07/FACTURA"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
   End If
'   Call CONECRUN("ACRE_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Genrep_Click()
    Call CONECRUN("GENREP07", "Diseñador de Formularios")
End Sub

Private Sub Image2_Click()

End Sub

Private Sub mnuLp_Click()
    Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub mnupresup_Click()
   Call CONECRUN("PRESUP07", "Sistemas de Presupuestos")
End Sub


Private Sub Orpago_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES07/ORPAGO"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
   End If
End Sub




Private Sub subdiarioyestadisticas_Click()
   If EXISTE%("ACRGES07.EXE") > 0 Then
     ACONEC$ = "ACRGES07/SUBDESTA"
     Call CONECRUN(ACONEC$, "Sub-Diario y Estadísticas de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
   End If

End Sub

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
    SQLX$ = " SELECT * FROM AGENDA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where facturado <> '3' "
    
    SQLX$ = SQLX$ + " ORDER BY Dia_Organizado "
    Set RS1 = READSET(SQLX$)
    With RS1
      Do While Not .EOF
        ORD% = ORD% + 1
        MSF2.Rows = ORD% + 1
        MSF2.TextMatrix(ORD%, 1) = XVALO(!nume_cli)
        MSF2.TextMatrix(ORD%, 2) = TRIM$(Left$(rasocli$(XVALO(!nume_cli)), 24)) + " - (" + TRIM$(Left$(DOMICLI$(XVALO(!nume_cli)), 24)) + ")"
'        MSF2.TextMatrix(ORD%, 3) = CODEXT$(XVALO(!CODTRABAJO))
'        MSF2.TextMatrix(ORD%, 4) = DESCRIT0$(XVALO(!CODTRABAJO))
        MSF2.TextMatrix(ORD%, 3) = SAFETEXT$(!CodexTRab)
        MSF2.TextMatrix(ORD%, 4) = SAFETEXT$(!DescTrab)
        MSF2.TextMatrix(ORD%, 5) = FECHATEX$(FECHANUM(!Fecha_Original))
        MSF2.TextMatrix(ORD%, 6) = !OrdenSemana
            HORORIG$ = TRIM$(!Horario_Orig)
            If HORORIG$ = "12:00:00 p.m." Then HORORIG$ = ""
        MSF2.TextMatrix(ORD%, 7) = HORORIG$
        MSF2.TextMatrix(ORD%, 8) = !CONTINUIDAD
        MSF2.TextMatrix(ORD%, 9) = !CUOTAS
          DIAORGAN$ = !Dia_Organizado
          If DIAORGAN$ = "" Then
             DIAORGAN$ = !Fecha_Original
          End If
        If Right(DIAORGAN$, 2) = 80 Then DIAORGAN$ = ""
   
        MSF2.TextMatrix(ORD%, 10) = DIAORGAN$
          HOROORGAN$ = TRIM$(!Horario_Organizado)
          If HOROORGAN$ = "12:00:00 p.m." Then
             HOROORGAN$ = ""
          ElseIf HOROORGAN$ = "" Then
             HOROORGAN$ = HORORIG$
          End If
        MSF2.TextMatrix(ORD%, 11) = HOROORGAN$
        FEHOR = FECHATEX$(FECHANUM(!Dia_Realizado)): If Right(FEHOR, 2) = 80 Then FEHOR = ""
        MSF2.TextMatrix(ORD%, 12) = FEHOR
        MSF2.TextMatrix(ORD%, 13) = ECOARCH$("PERSONAL", MKS$(XVALO(!OPER_REALIZO)))
        Factu$ = TRIM$(!FACTURADO)
        If Factu$ <> "SÍ" Then
            If XVALO(Factu$) = 0 Then
               Factu$ = ""
            Else
               Factu$ = "SÍ"
            End If
        End If
        MSF2.TextMatrix(ORD%, 14) = Factu$
        MSF2.TextMatrix(ORD%, 15) = !MESANO_ORIG
        MSF2.TextMatrix(ORD%, 16) = !ID
        MSF2.TextMatrix(ORD%, 17) = XVALO(!OPER_REALIZO)
        .MoveNext
      Loop
    End With

End Sub

Sub FILTRAR_GRILLA_TRABAJO()

    ORD% = 0
    MSF2.Rows = 1
    SQLX$ = " SELECT * FROM AGENDA WITH(NOLOCK)"
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
      For i% = 0 To Option1.UBound
        TILDE = Option1(i%).Value
        If TILDE = True Then Exit For
      Next i%
      If TILDE = False Then
        Call COMUNI("Para Filtrar Por Fecha debe Seleccionar alguna de las 3 Opciones")
        GoTo 99
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
    If Me.OPTPENDIENTEFACTURA = True Then SQLX$ = SQLX$ + " AND FACTURADO < 1 OR FACTURADO IS NULL"
    '
    If Me.CHKSINEJECUTAR = 1 Then
      FECHAFACTU = CDbl(CVDAT(CVINT("01/01/14")))
      SQLX$ = SQLX$ + " AND Dia_Realizado    <= " & FECHAFACTU & ""
    End If
    '
    If Me.OPTSOLOFACTURADO = True Then
       SQLX$ = SQLX$ + " AND FACTURADO ='3'"
    End If
    GRUCLI% = XVALO(Me.TXTGRUPO)
    
    SQLX$ = SQLX$ + " ORDER BY Dia_Organizado  "
    Set RS1 = READSET(SQLX$)
    With RS1
      Do While Not .EOF
        NC% = XVALO(!nume_cli)
        If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
            ORD% = ORD% + 1
            MSF2.Rows = ORD% + 1
            MSF2.TextMatrix(ORD%, 1) = NC
            MSF2.TextMatrix(ORD%, 2) = TRIM$(Left$(rasocli$(NC%), 24)) + "(" + TRIM$(Left$(DOMICLI$(NC%), 24)) + ")"
            MSF2.TextMatrix(ORD%, 3) = SAFETEXT$(!CodexTRab)
            MSF2.TextMatrix(ORD%, 4) = SAFETEXT$(!DescTrab)
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
            Factu$ = TRIM$(!FACTURADO)
            If Factu$ <> "SÍ" Then
                If XVALO(Factu$) = 0 Then
                   Factu$ = ""
                Else
                   Factu$ = "Aprobado"
                End If
            End If
            '
            MSF2.TextMatrix(ORD%, 14) = Factu$
            MSF2.TextMatrix(ORD%, 15) = !MESANO_ORIG
            MSF2.TextMatrix(ORD%, 16) = !ID
    '        MSF2.TextMatrix(ORD%, 17) = IDDIA%
        End If
85      .MoveNext
      Loop
    End With
99  Screen.MousePointer = 1
End Sub


Function CONTINUIDAD%(CODTRABAJO$)
   CI11% = CODINT%(CODTRABAJO$)
   CONTINUIDAD% = 1
   For i& = 1 To ULTREG&("CONTINUI")
      X$ = REGLEIDO$("CONTINUI", i&)
      CIXI2% = CVI(Mid(X$, 1, 2))
      If CI11% = CIXI2% Then
         CONTINUIDAD% = CVI(Mid(X$, 35, 2))
         Exit For
      End If
   Next i&
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
    For i& = 1 To MSF.Rows - 1
       MSF.TextMatrix(i&, COLUM&) = CVINT(MSF.TextMatrix(i&, COLUM&))
    Next i&
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
    For i& = 1 To MSF.Rows - 1
       MSF.TextMatrix(i&, COLUM&) = FECHATEX$(XVALO((MSF.TextMatrix(i&, COLUM&))))
    Next i&

End Sub

Private Sub cmbdias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub CMBMES_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Command5.SetFocus
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

Sub Command10_Click()
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
    Call CARGA_ENCABEZADO(MUESTRA17.GRID, CAMPOS$, MUESTRA17, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(MUESTRA17.MSF_2, CAMPOS$, MUESTRA17, 1, 0, ANTOT, , 1)
    MUESTRA17.Command3.Caption = "Selección de Trabajos a Facturar"
    SQLX$ = " SELECT * FROM AGENDA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUME_CLI > 0 "
    SQLX$ = SQLX$ + " and  FACTURADO = '1'" 'facturado van a aser las que se desean facturar"
    '
'    Dim RS As DAO.Recordset
'
'    Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
      Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        FE% = CVINT(!Dia_Realizado)
        NC = XVALO(!nume_cli)
        RASYDOMI$ = TRIM$(Left$(rasocli$(NC), 24)) + " (" + TRIM$(Left$(DOMICLI$(NC), 24)) + ")"
        CODITRABAJO = XVALO(!CODTRABAJO)
        'DESCRITRABAJO$ = TRIM$(ECOARCH$("TIPOTRAB", MKI$(XVALO(!CODTRABAJO))))
        DESCRITRABAJO$ = DESCRIT0$(XVALO(!CODTRABAJO))
        OPERARIO% = XVALO(!OPER_REALIZO)
        NAMEOPERARIO$ = ECOARCH$("PERSONAL", MKS$(XVALO(!OPER_REALIZO)))
        If FE% < 1 Then
           Call COMUNI("Falta Fecha Realizado Para el Cliente: " & RASYDOMI$)
           MUESTRA17.GRID.Clear
           GoTo 99
        End If
        MUESTRA17.GRID.Rows = J& + 1
        MUESTRA17.GRID.TextMatrix(J&, 1) = FECHATEX$(FE%)
        MUESTRA17.GRID.TextMatrix(J&, 2) = NC
        MUESTRA17.GRID.TextMatrix(J&, 3) = RASYDOMI$
        MUESTRA17.GRID.TextMatrix(J&, 4) = CODITRABAJO
        MUESTRA17.GRID.TextMatrix(J&, 5) = DESCRITRABAJO$
        MUESTRA17.GRID.TextMatrix(J&, 6) = OPERARIO%
        MUESTRA17.GRID.TextMatrix(J&, 7) = NAMEOPERARIO$
        MUESTRA17.GRID.TextMatrix(J&, 8) = XVALO(!CUOTAS)
        MUESTRA17.GRID.TextMatrix(J&, 9) = XVALO(!ID)
        
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing

    MUESTRA17.Show 1
99  Command13.Enabled = True

End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   A = SSTab1.Caption
   
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
     If SSTab1.Caption = "Informe de Tareas" Then
      Call GENEXCEL(AREXPO$, Me.MSF2)
     Else
      Call GENEXCEL(AREXPO$, Me.MSF1) 'agenda original
     End If
   End If
   '
99 Command14.Enabled = True

End Sub
Sub GENEXCEL2(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
'objetoExcel.selection.NumberFormat = "dd/MM/yyyy"
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols - 1
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1&)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             'VALORX = FECHATEX$(Int(FECHANUM(FEXU$)))
             VALORX = FETEXCOR2$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1& - 1), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '

          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          Select Case COLUM&
             Case 1
              XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
             Case 2
              XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "DD-MMM"
             Case 3
              XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "hh-mm"
             Case 4, 5, 6
              XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          End Select
'          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
'          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
'          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
'          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
'
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             'VALORX = FECHATEX$(Int(FECHANUM(FEXU$)))
             VALORX = FETEXCOR2$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub


Private Sub Command15_Click()
   Command15.Enabled = False
   Call SELECHO("GRUCOCLI")
   CUECO% = XVALO(VALACT1$("GRUCOCLI"))
   If CUECO% < 1 Then GoTo 99
   Me.TXTGRUPO = CUECO%
   
   
99 Command15.Enabled = True
   
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
   Set RS = New ADODB.Recordset
'    Dim RS As DAO.Recordset
    For i& = 1 To MSF1.Rows - 1
      If XVALO(MSF1.TextMatrix(i&, 10)) < 1 Then ' si ya existe lo edita si no addnew
        Set RS = New ADODB.Recordset
        SQLX$ = "SELECT TOP 1 * FROM  AGENDAORIGINAL  "
        SQLX$ = SQLX$ + " ORDER BY ID DESC"

        'Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
25      On Error Resume Next
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
           NUEVOID = 1 + XVALO(RS!ID)
           RS.AddNew
      End If
      
      If XVALO(MSF1.TextMatrix(i&, 10)) > 0 Then '
        Set RS = New ADODB.Recordset
        SQLX$ = "SELECT * FROM  AGENDAORIGINAL  "
        SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NCLIE

        SQLX$ = SQLX$ + " AND ID = " & XVALO(MSF1.TextMatrix(i&, 10))
'        Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
35      On Error Resume Next
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 35
           End If
           On Error GoTo 0
      End If
      With RS
          AD = XVALO(!ID)
          If AD < 1 Then AD = NUEVOID 'SI AD ESTA EN 0 ES POR QUE NO ESTA REGISTRADO ENTONCES LE CARGA EL NUEVO ID
         !nume_cli = XVALO(Text1)
         !ORDEN = MSF1.TextMatrix(i&, 1)
    '     !CODTRABAJO = XVALO(MSF1.TextMatrix(I&, 2))
         !CODTRABAJO = CODINT%(MSF1.TextMatrix(i&, 2))
         !CodexTRab = Left$(MSF1.TextMatrix(i&, 2), 24)
         !DescTrab = Left$(MSF1.TextMatrix(i&, 3), 64)
         
         !DIA = MSF1.TextMatrix(i&, 4)
         !ID_DIA = IDDIA%(MSF1.TextMatrix(i&, 4))
         !OrdenSemana = XVALO(MSF1.TextMatrix(i&, 5))
         !Horario = MSF1.TextMatrix(i&, 6)
         !CONTINUIDAD = XVALO(MSF1.TextMatrix(i&, 7))
         !CUOTAS = XVALO(MSF1.TextMatrix(i&, 8))
         !MarBorra = 0
         !nume_cli = NCLIE
         !APARTIRDE = MSF1.TextMatrix(i&, 9)
         .Update
     End With
     MSF1.TextMatrix(i&, 10) = AD
     RS.Close: Set RS = Nothing
    Next i&
    
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
        SQLX$ = "DELETE  FROM AgendaOriginal  WHERE ID = " & XVALO(Label2.Caption)
        FLEXCONN.Execute (SQLX$)
        Call cmdEliminaItem(MSF1, XVALO(Label3))
        Call LIMPIA_EDICION
        
        'Call borrafila(XVALO(TXTORDEN), Agenda, MSF1)
    End If
    'REFRESCO LA GRILLA
    
    'Agenda.CARGAR_GRILLA_TRABAJO ' comentada por que demora mucho
    'Directamente la limpio
    MSF2.Rows = 1
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
    Screen.MousePointer = 11
    MESNUMERO_SELECCIONADO% = NUMMES%(MESASEL$)
    SQLX$ = "SELECT  COUNT(*) AS CANTI FROM AGENDA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE MESANO_ORIG = '" & MESASEL$ & "'"
    SQLX$ = SQLX$ + " AND FACTURADO <>'3'"
    'Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RS = READSET(SQLX$)
    If RS!CANTI > 0 Then
       OPX% = COMALTER%("Atención!!\Se Ha Detectado que yá se han Importado Datos a la Agenda Activa\\Cancelar\Actualizar Agenda Activa")
       If OPX% <= 1 Then
         Screen.MousePointer = 1
         GoTo 10
       End If
    End If
    RS.Close
    Set RS = Nothing
    '
    'RECORRO LA TABLA AGENDAORIGINAL Y VEO SI EXISTE EL MISMO EN LA TABLA AGENDA
    'TENER EN  CUENTA SI EXISTE FILA CON INFO DE LA ORIGINAL EN LA TABLA AGENDA NO SE COPIA NI SE ACTUALIZA
    'SI NO EXISTE AGREGA EL REGISTRO
    'SI EL REGISTRO EN LA ORIGINAL SE ELIMINO Y SE VOLVIO A CARGAR CON LA MISMA INFO LO VA A AGREGAR EN LA TABLA AGENDA
    'POR QUE VA A CAMBIAR EL IDENTIFICADOR.
    Set RS = New ADODB.Recordset
    SQLX$ = " SELECT * FROM AGENDAORIGINAL WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY ID_DIA"
    'Set RS1 = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RS1 = READSET(SQLX$)

    With RS1
      Do While Not .EOF
        IDX% = XVALO(!ID)
        If IDX% = 476 Then
         G = 1
        End If
        CONTIN% = XVALO(!CONTINUIDAD)
        MMMM = SAFETEXT$(!APARTIRDE)
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
            SQLX$ = "SELECT  COUNT(*) AS CANTI FROM AGENDA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$
            'Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
            Set RS = READSET(SQLX$)
            'SI YA ESTA EN LA GRILLA LO DEJA COMO ESTA, SI EL ORIGINAL NO ESTA Y SE AGREGO MANULMENTE LO VA A DUPLICAR.
            'SI SE QUIERE REFRESCAR SE DEBERA ELIMINAR DE LA GRILLA MANUALMENTE.
            If RS!CANTI < 1 Then
    
               SQLX$ = " SELECT * FROM AGENDA "
               SQLX$ = SQLX$ + "WHERE CodTrabajo < 1 "
'              Set RS2 = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
               Set RS2 = New ADODB.Recordset
25             On Error Resume Next
                 RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                 If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                    On Error GoTo 0
                    Call CapturaErrorOpen
                    GoTo 25
                 End If
                 On Error GoTo 0
                With RS2
                    .AddNew
                    !CODTRABAJO = RS1!CODTRABAJO
                    !CodexTRab = Left$(SAFETEXT$(RS1!CodexTRab), 24)
                    !DescTrab = Left$(SAFETEXT$(RS1!DescTrab), 64)
'                    AA = FECHATEX$(UBICAR_DIA_SEMANA%(rs1!OrdenSemana, rs1!DIA, MESNUM$(MESASEL$), ANOXCON20$(MESASEL$)))
                    !Fecha_Original = CVDAT(UBICAR_DIA_SEMANA%(RS1!OrdenSemana, RS1!DIA, MESNUM$(MESASEL$), ANOXCON20$(MESASEL$)))
                    !OrdenSemana = RS1!OrdenSemana
                    !Horario_Orig = RS1!Horario
                    !CONTINUIDAD = RS1!CONTINUIDAD
                    !CUOTAS = RS1!CUOTAS
                    !nume_cli = RS1!nume_cli
                    !MarBorra = 0
'                    !CODTRABAJO = rs1!CODTRABAJO
                    !Dia_Organizado = CVDAT(0)
                    !Horario_Organizado = ""
                    !Dia_Realizado = CVDAT(0)
'                    !Horario_Realizado = ""
                    !MESANO_ORIG = MESASEL$
                    !ID_ORIGINAL = IDX%
                    !FACTURADO = ""
                    .Update
                End With
            End If
        End If
        .MoveNext
        
      Loop
    End With
    Call CARGAR_GRILLA_TRABAJO
    
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "REFRESCAR LA AGENDA " & vbCrLf
       TXTO$ = TXTO$ + "LA MISMA ACTUALIZA EL MES A SELECCIONAR" + vbCrLf
       TXTO$ = TXTO$ + "SI YA HA SIDO MIGRADO ALGUN TRABAJO ESTE SE VA A MANTENER" + vbCrLf
       TXTO$ = TXTO$ + "(NO ASI SI SE CREO MANUALMENTE)" + vbCrLf
       TXTO$ = TXTO$ + "ESTO QUIERE DECIR QUE SI SE CREO MANUALMENTE EL MISMO SE VA A DUPLICAR" + vbCrLf
       TXTO$ = TXTO$ + "EN TODO CASO BORRAR MANUALMENTE LA SEGUNDA." + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI DE ALGUNO DE LOS TRABAJOS SE CAMBIO ALGUN DATO EN LA AGENDA ORIGINAL" + vbCrLf
       TXTO$ = TXTO$ + "SE DEBERA ELIMINAR MANUALMENTE Y LUEGO ACTUALIZAR GRILLA" + vbCrLf
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
    
    If TRIM$(CMBORDEN) = "" Or XVALO(CMBORDEN) < 1 Then
      Call COMUNI("Falta Seleccionar Orden del Día")
      GoTo 99
    End If
    
    If TRIM$(DTHORARIO) = "" Or TRIM$(DTHORARIO) = "0:00:00" Then
      Call COMUNI("Falta Seleccionar Horario")
      GoTo 99
    End If
    
    If XVALO(TXTREPITE) = 0 Then
      Call COMUNI("Falta Seleccionar la Continuidad")
      GoTo 99
    End If
    If TRIM$(CMBMES) = "" Then
      Call COMUNI("Falta Seleccionar A Partir de ???")
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
    MSF1.TextMatrix(i&, 10) = XVALO(Label2.Caption)
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
    Label3 = ""
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
   CODTRA$ = VALACT1$("MASTER")
   If CODTRA$ = "" Then GoTo 99
   TXTCODTRABAJO = CODTRA$
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
   Call CONECRUN("VENGES07/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Form_Load()

    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
'    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
'    If TERMINA% > 0 Then Call FINAL("")
    '
'    CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
'    If CMO$ <> "" Then Label10.Caption = CMO$

    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call GRACONTROL("NEWASPEC", "CAMBIOAS", "SI")
   If OKEY% < 1 Then Call FINAL("")
   '
   If ControlMenu%(Me.Name, Me) < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   'VALIDACION PARA del CIELO DESTAPACIONES POR SI NO PAGA
'    FF% = FECHANUM("30/04/16")
'    If FF% <= HOY(HO$) Then
'      Call Form_Load
'    End If

   Openforms = DoEvents
   AA = DIASEM%(HOY(HO$))
   VV = METEXTO$(HOY(HO$)) ' ABRIL 2015
   BB = DSEM$(HOY(HO$)) ' LUNES
   'XX = UBICAR_DIA_SEMANA%(5, "MIERCOLES", "06", "2015")
   AA = FECHATEX$(XX)
   Call CARGACOMBO_DIAS
   Call CARGACOMBO_ORDEN
   Call CARGAMES
   Call CARGA_ENCABEZADO(Me.MSF1, "Ord/F5;Cod.Trab/F8;Descripción/A24;Día/A12;Semana/A8;Horario/A12;Repite Cada/A10;Cuotas/F7;A Partir de/A12;ID/F0", Agenda17, 0, 0, 0, 0)
   Call CARGA_ENCABEZADO(Me.MSF2, "Cta./F5;R.Social-Domicilio/A48;Cod./F5;Descripción/A22;F.Orig./A10;N°/A0;Hor.Orig./A10;Cont.Mens./A0;Cuotas/F0;F.Organ./A10;Hor.Organ./A10;Realizado/A10;Realiza;Facturar ?/A8;Mesano_Orig/A0;ID/F0:IDDIA/F0;UsuarioRealizo/F0", Agenda17, 0, 0, 0, 0)
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
   Call ABRESTOCKS
   Call CREATABLA_STOCKS
         
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
   Call CONECRUN("ARCHIG07", "Archivos Maestros")

End Sub

Private Sub masteritem_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub mnufacturacion_Click()
   Call CONECRUN("VENGES04/FACTURA", "Gestión de Ventas y Deudores")
End Sub

Private Sub mnuproveedores_Click()
  If EXISTE%("ACREGES04.EXE") > 0 Then
      Call CONECRUN("ACREGES04", "Gestión de Pedidos de Clientes")
      Exit Sub
   End If

End Sub

Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '
    Label3 = REG&
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

Private Sub msf2_Click()
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
    i& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If MSF2.Rows < 2 Then Exit Sub
    If J& = 0 Then
       CALENDAR.LABEL1 = MSF2.TextMatrix(i&, 5)
       CALENDAR.Show 1
       Exit Sub
    End If
    
    If i& = 0 Or i& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = i&
    MSF2.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '
   With EDITAGEN17
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
        .TXTFACTURADO = MSF2.TextMatrix(REG&, 14): If .TXTFACTURADO = "Aprobado" Then .TXTFACTURADO.Enabled = False
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

  If EXISTE%("PEDIDO07.EXE") > 0 Then
      Call CONECRUN("PEDIDO07", "Gestión de Pedidos de Clientes")
      Exit Sub
   End If

End Sub

Private Sub Recibo_Click()
   If EXISTE%("VENGES07.EXE") > 0 Then
     ACONEC$ = "VENGES07/RECIBO"
     Call CONECRUN(ACONEC$, "Recibo de Cobranza")
   Else
     Call COMUNI("Falta Módulo VENGES07")
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
     SQLX$ = "SELECT * FROM AGENDAORIGINAL WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NUME_CLI = " & NCLIE
'     Set RS = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
     Set RS = READSET(SQLX$)
     i& = 0
     With RS
       On Error Resume Next
         Do While Not .EOF
                i& = i& + 1
                MSF1.Rows = i& + 1
                MSF1.TextMatrix(i&, 1) = !ORDEN
                MSF1.TextMatrix(i&, 2) = CODEXT$(XVALO(!CODTRABAJO))
                MSF1.TextMatrix(i&, 3) = DESCRIT0$(XVALO(!CODTRABAJO))
                MSF1.TextMatrix(i&, 4) = !DIA
                MSF1.TextMatrix(i&, 5) = !OrdenSemana
                MSF1.TextMatrix(i&, 6) = !Horario
                  CONTIN% = XVALO(!CONTINUIDAD)
                  CONTINUIX$ = ""
                  If CONTIN% > 1 Then CONTINUIX$ = SAFETEXT$(CONTIN%) & " Meses"
                  If CONTIN% = 1 Then CONTINUIX$ = SAFETEXT$(CONTIN%) & " Mes"
                MSF1.TextMatrix(i&, 7) = CONTINUIX$
                MSF1.TextMatrix(i&, 8) = XVALO(!CUOTAS)
                MSF1.TextMatrix(i&, 9) = !APARTIRDE
                MSF1.TextMatrix(i&, 10) = XVALO(!ID)
           .MoveNext
         Loop
     End With
     RS.Close
     Set RS = Nothing
End Sub

Sub cardatclie()
    Screen.MousePointer = 11
    NCLIE = XVALO(Text1)
    TXTNOMBRE.TEXT = rasocli$(NCLIE)
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

Private Sub Text27_Change()

End Sub


Private Sub tpersonal_Click()
    Call TRALOCAL("Personal", "")
10  VTB% = VENTABU%("Personal")
    If VTB% >= 0 Then
      jj& = 0
      For U& = 1 To ULTREG&("!Personal")
        jj& = jj& + 1
        X$ = REGLEIDO$("!Personal", U&)
        CODI& = CVS(Left$(X$, 4))
        DESCRI$ = TRIM$(Mid$(X$, 5, 40))
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
      Call SETULTREG("Personal", jj&)
    End If
    Call TRACENTRAL("Personal", "")
    Call CIERRARCH("Personal")
    Call CIERRARCH("!Personal")

End Sub


Private Sub TsgTrabajo_Click()
    Call TRALOCAL("CONTINUI", "")
10  VTB% = VENTABU%("CONTINUI")
    If VTB% >= 0 Then
      jj& = 0
      For U& = 1 To ULTREG&("!CONTINUI")
        jj& = jj& + 1
        X$ = REGLEIDO$("!CONTINUI", U&)
        CODI& = CVI(Left$(X$, 2))
        MESES1% = CVI(Mid$(X$, 35, 2))
        AA = CVI(Mid(X$, 35, 2))
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
      Call SETULTREG("CONTINUI", jj&)
    End If
    Call TRACENTRAL("CONTINUI", "")
    Call CIERRARCH("CONTINUI")
    Call CIERRARCH("!CONTINUI")

End Sub


Private Sub ttrabajo_Click()
    Call TRALOCAL("TIPOTRAB", "")
10  VTB% = VENTABU%("TIPOTRAB")
    If VTB% >= 0 Then
      jj& = 0
      For U& = 1 To ULTREG&("!TIPOTRAB")
        jj& = jj& + 1
        X$ = REGLEIDO$("!TIPOTRAB", U&)
        CODI% = CVI(Left$(X$, 2))
        DESCRI$ = TRIM$(Mid$(X$, 3, 48))
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
      Call SETULTREG("TIPOTRAB", jj&)
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


Private Sub TXTCODIGO_Change()
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
    For i% = 1 To Len(TRIM$(INGRESO$))
       POS1% = InStr(INGRESO$, ":")
       
       If i% = 1 Then MAXI1% = 24
       If i% = 2 Then MAXI1% = 59: If CONTROL("PADIAMOD", "CONFHORA") = "MINUCENT" Then MAXI1% = 99 ' configuracion minutos centecimal
       If i% = 3 Then MAXI1% = 99
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
    Next i%
    '
End Function

Private Sub TXTCODRESPONSABLE_Change()
   TXTDESRESPONSABLE = ""
   TXTDESRESPONSABLE = ECOARCH$("PERSONAL", MKS$(XVALO(TXTCODRESPONSABLE)))


End Sub

Private Sub TXTCODTRABAJO_Change()
   TXTDESTRABAJO = ""
   AA = CODINT%(TXTCODTRABAJO)
   TXTDESTRABAJO = DESCRIT0$(CODINT%(TXTCODTRABAJO))

End Sub

Private Sub TXTCUENTA_Change()
    NCLIE = XVALO(TXTCUENTA)
    TXTNOMBRE2.TEXT = rasocli$(NCLIE)
End Sub

Private Sub TXTDESDE_DblClick()
   TXTDESDE = ""

End Sub


Private Sub TXTGRUPO_Change()
      CUECO% = XVALO(Me.TXTGRUPO)
      If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
        CUECO% = 0
      End If
      Me.TXTREFGRUPO = ECOARCH$("GRUCOCLI", MKI$(CUECO%))
End Sub

Private Sub TXTHASTA_DblClick()
       TXTHASTA = ""

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   i& = MSF2.MouseRow
   J& = MSF2.MouseCol
   If i& = 0 Or i& > MSF2.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       MSF2.Row = i&
       MSF2.COL = J&
       MSF2.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
       OPX% = COMALTER%("Desea Generar Ahora La Impresión del Aviso de Trabajo\\Cancelar\Imprimir Ahora")
       If OPX% <= 1 Then GoTo 30
       NC = MSF2.TextMatrix(i&, 1)
       TITRABAJO% = XVALO(MSF2.TextMatrix(i&, 3))
       DIAX1$ = MSF2.TextMatrix(i&, 10)
       HORAX$ = MSF2.TextMatrix(i&, 11)
       FechaTrabajo$ = DIAX1$ & " a las " & Left$(HORAX$, 5)
       Call ImprimirAviso(NC, HOY(HO$), FechaTrabajo$, TITRABAJO%, HORAX$)
30     MSF2.CellBackColor = RGB(255, 255, 255)
       Screen.MousePointer = 1
   End If
   
End Sub


Sub ImprimirAviso(NC, FE%, FechaTrabajo$, TIPO_TRABAJO%, HORAX$)
   'PASAR LOS VALRES DESDE LA LLAMADA IMPRIMIR AVISO DE LOS VALRES QUE HAGAN FALTA SEGUN CONTINUACION
   
   'Junio de 2015. FECHA DE HOY
   'THORNE 382   DIRECCION DEL CONSORCIO  MSF2.TextMatrix(ORD%, 1) = XVALO(!NUME_CLI) -DOMICLI$(XVALO(!NUME_CLI)
   'miércoles 17-jun a las  08:45 hs.   FECHA DE PROYECTADA DE REALIZACION  DSEM$ METEXTO$  + HORA DE PANTALLA
   'MSF2.TextMatrix(ORD%, 10) = DIAORGAN$  - MSF2.TextMatrix(ORD%, 11) = HOROORGAN$
   DOMICILI$ = DOMICLI$(NC)
   FEHOY% = HOY(HO$)
   POS1% = InStr(METEXTO$(FEHOY%), " ")
   FechaDeHoy$ = Mid$(METEXTO$(FEHOY%), 1, POS1% - 1) & " de " & Right$(FECHATEXLAR$(FEHOY%), 4) 'Junio de 2015. FECHA DE HOY
   FechaDeHoy$ = Mid$(FechaDeHoy$, 1, 1) & Mid$(LCase(FechaDeHoy$), 2) 'para que quede en mayuscula la primer letra
   FeTrabajo% = FECHANUM(Left$(FechaTrabajo$, 8))
   '
   'ARMO EL TEXTO DE DIA-MES Y HORA DE TRABAJO: miércoles 17-jun a las  08:45 hs.
   'VER CON QUE FORMATO VIENE LA FECHA DE TRABAJO.
   DIANUMERO$ = Left$(FechaTrabajo$, 2)
   DIATEX$ = DSEM$(FeTrabajo%)
   MESTXT$ = METEXTO$(FeTrabajo%)
   HORAX$ = Left$(HORAX$, Len(HORAX$) - 3) 'BORRO LOS SEGUNDOS
   FechaDeRealizacion$ = DIATEX$ & " " & DIANUMERO$ & "-" & MESTXT$ & " a las " & HORAX$ & "."
   '
   OPX% = COMALTER%("Definir Tipo de Aviso a Imprimir\\Aviso de Desinsectación\Aviso de Limpieza de Tanque")
   If OPX% < 1 Then Exit Sub
   FILTX$ = FechaDeHoy$ & ";" & FechaDeRealizacion$
   If OPX% = 1 Then
     Call STDFORM("INS-CLO", FILTX$)
   ElseIf OPX% = 2 Then
     Call STDFORM("TQUE-CLO", FILTX$)
   End If

'   Call IMPRIME_AVISO_TRABAJO(TIPO_TRABAJO%, DOMICILI$, FechaDeHoy$, FechaDeRealizacion$)
   '
End Sub

Sub IMPRIME_AVISO_TRABAJO(TIPO_TRABAJO%, DOMICILIX$, FechaDeHoy$, FechaDeRealizacion$)
   '
   '**************************************************************************************************************
   'NO OLVIDAR DE AGREGAR LA REFERENCIA MICROSFT WORD 11.0 POR QUE SI NO NO DA ERROR PERO NO REEMPLAZA EL TEXTO.
   '**************************************************************************************************************
   '
   OPX% = COMALTER%("Definir Tipo de Aviso a Imprimir\\Aviso de Desinsectación\Aviso de Limpieza de Tanque")
   If OPX% < 1 Then Exit Sub
   'PASAR NOMBRE DEL ARCHIVO
   If OPX% = 1 Then
     ARCH_SELECCIONADO$ = "INSECTOS"
   ElseIf OPX% = 2 Then
     ARCH_SELECCIONADO$ = "TANQUES"
   End If
   '
   'EL ARCHIVO VA A ESTAR EN LUDAT
   ARCH_WORD_ORIGINAL$ = LUDAT$ + ARCH_SELECCIONADO$
   
   'VALIDAR SI EXISTE, SI NO CREAR LA CARPETA DONDE SE VAN A GENERAR LOS WORD MODIFICADOS
   If Dir(LUDAT$ + "AVISOS_GENERADOS", vbDirectory) = vbNullString Then
     'Crea la carpeta
      Call MkDir(LUDAT$ + "AVISOS_GENERADOS")
   End If
   '
   ARCH_WORD_DESTINO$ = LUDAT$ + "AVISOS_GENERADOS\" + ARCH_SELECCIONADO$
   '
   Screen.MousePointer = 11
   OK% = 0
   '************************************************************
   'ARCH_WORD_ORIGINAL$ = ARCHIVO_WORD$(CODGALERIA&) 'AQUI AGREGAR RUTA DEL ARCHIVO ORIGINAL
   'POS1% = InStrRev(ARCH_WORD_ORIGINAL$, ".")
   '
'   'INDICO DONDE GUARDAR EL ARCHIVO - CREO LA CARPETA CONTRATOS_GENERADOS EN LUDAT$
'   If Dir(LUDAT$ + "CARPETA WORD AVISOS", vbDirectory) = vbNullString Then
'     'Crea la carpeta
'      Call MkDir(LUDAT$ + "CARPETA WORD AVISOS")
'   End If
   '
'   'AL NOMBRE DEL ARCHIVO LE CONCATENO EL NUMERO DE CONTRATO Y CON EL MISMO NOMBRE DEL ORIGINAL
'   POS2% = InStrRev(ARCH_WORD_ORIGINAL$, "\")
'   ARCH_WORD_DESTINO$ = LUDAT$ + "CARPETA WORD AVISOS\" + TRIM$(Label18) + "-" & Mid$(ARCH_WORD_ORIGINAL$, (POS2% + 1))
   '
   'COPIO EL ARCHIVO WORD ORIGINAL
   On Error Resume Next
   FileCopy ARCH_WORD_ORIGINAL$ + ".doc", ARCH_WORD_DESTINO$ + ".doc"
   On Error GoTo 0
   '
   'AHORA SE DEBERIA REEMPLAZAR LOS DATOS VER SI SE PUEDE USAR REEMPLA_DATOS$ Y DENTRO DE ESTA UTILIZAR & _
   FUNCION DE REEMPLAZO DE DATOS EN WORD.
   '
   'REEMPLAZO LOS DATOS VARIABLES
   '*****************************
   Call REEMPLA_DATOS2$(DOMICILIX$, FechaDeHoy$, FechaDeRealizacion$, ARCH_WORD_DESTINO$ + ".doc")
   '
   OK% = 1
'"X:\Del Cielo\Flexoft\DATOS\AVISOS_GENERADOS\TANQUES"
End Sub




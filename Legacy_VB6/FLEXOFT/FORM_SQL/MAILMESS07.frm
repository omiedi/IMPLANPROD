VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Begin VB.Form MAILMESS07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar mensaje de Correo Electrónico"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command12 
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
      Left            =   11210
      Picture         =   "MAILMESS07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Configuracion Envio de Mails por FLEXWSK"
      Top             =   650
      Width           =   545
   End
   Begin VB.Frame Frame7 
      Caption         =   "Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   8560
      TabIndex        =   31
      Top             =   5600
      Width           =   3250
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         ItemData        =   "MAILMESS07.frx":0442
         Left            =   105
         List            =   "MAILMESS07.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   330
         Width           =   3040
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   105
         TabIndex        =   33
         Top             =   1420
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1680
         TabIndex        =   32
         Top             =   1420
         Width           =   1520
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Asuntos a Tratar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      TabIndex        =   52
      Top             =   7450
      Width           =   7260
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   6720
         Picture         =   "MAILMESS07.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Guarda Actividad Programada"
         Top             =   110
         Width           =   520
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   53
         Top             =   260
         Width           =   6555
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Con:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1800
      TabIndex        =   51
      Top             =   7450
      Width           =   2655
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "MAILMESS07.frx":0888
         Left            =   120
         List            =   "MAILMESS07.frx":088A
         TabIndex        =   54
         Text            =   "Combo1"
         Top             =   260
         Width           =   2415
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Contacto el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   28
      Top             =   7450
      Width           =   1575
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   30
         Top             =   260
         Width           =   1030
      End
      Begin VB.CommandButton Command15 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         TabIndex        =   29
         Top             =   260
         Width           =   150
      End
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5520
      Picture         =   "MAILMESS07.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Agrupar para Próximo Envío"
      Top             =   5160
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Frame Frame9 
      Caption         =   "Adjuntar Preferidos"
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
      Left            =   8560
      TabIndex        =   18
      Top             =   2200
      Width           =   3250
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1630
         TabIndex        =   21
         Top             =   2985
         Width           =   1520
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   120
         TabIndex        =   20
         Top             =   2985
         Width           =   1500
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   315
         Width           =   3040
      End
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   5040
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Addr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8640
      TabIndex        =   36
      ToolTipText     =   "Leer e Importar Direcciones de Mails"
      Top             =   2280
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   11210
      Picture         =   "MAILMESS07.frx":0B96
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   545
   End
   Begin VB.Frame Frame10 
      Caption         =   "Programar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   8560
      TabIndex        =   22
      Top             =   1200
      Width           =   1320
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1080
         TabIndex        =   50
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   315
         Width           =   150
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Left            =   780
         TabIndex        =   25
         Top             =   610
         Width           =   430
      End
      Begin VB.TextBox Text1 
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
         Left            =   105
         TabIndex        =   23
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   105
         TabIndex        =   26
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Hs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   500
         TabIndex        =   24
         Top             =   630
         Width           =   225
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton COMMAND16 
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
         Height          =   260
         Left            =   1560
         TabIndex        =   47
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox Text7 
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
         Left            =   600
         TabIndex        =   45
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox Text6 
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
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   360
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Repetir"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   46
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "ds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1470
         TabIndex        =   44
         Top             =   270
         Width           =   225
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1600
      Left            =   100
      TabIndex        =   11
      Top             =   5835
      Width           =   8340
      Begin VB.Frame Frame5 
         Caption         =   "Frame4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   7100
         TabIndex        =   12
         Top             =   50
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   1020
         TabIndex        =   13
         Top             =   50
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   8320
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   8280
         TabIndex        =   15
         Top             =   0
         Width           =   8340
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   75
            TabIndex        =   16
            Top             =   0
            Width           =   8100
         End
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   10575
         X2              =   10575
         Y1              =   0
         Y2              =   6225
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   8560
      TabIndex        =   9
      Top             =   130
      Width           =   2550
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "MAILMESS07.frx":0CE0
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Guardar y Enviar Inmediatamente"
         Top             =   240
         Width           =   520
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Solicita Confirmar Recibido"
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
         Left            =   1320
         TabIndex        =   41
         Top             =   180
         Value           =   1  'Checked
         Width           =   1120
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   670
         Picture         =   "MAILMESS07.frx":1122
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Guardar para Envio Posterior"
         Top             =   240
         Width           =   520
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destinatario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   8350
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Text            =   " "
         Top             =   960
         Width           =   7050
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   " "
         Top             =   1260
         Width           =   7050
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   " "
         Top             =   300
         Width           =   5250
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   " "
         Top             =   615
         Width           =   7050
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Op:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6360
         TabIndex        =   59
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   58
         Top             =   615
         Width           =   1245
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   56
         Top             =   1260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7020
         TabIndex        =   49
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Cte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6360
         TabIndex        =   48
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "C/Copia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   320
         TabIndex        =   39
         Top             =   1030
         Width           =   750
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Asunto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   6
         Top             =   720
         Width           =   750
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   735
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Mensaje Personalizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   105
      TabIndex        =   0
      Top             =   1850
      Width           =   8340
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "MAILMESS07.frx":1564
         Top             =   315
         Width           =   8100
      End
   End
   Begin VB.Line Line3 
      X1              =   8445
      X2              =   8445
      Y1              =   5800
      Y2              =   7400
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   27
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "MAILMESS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ENVIAINMEDIATO%

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        
Private Sub Command1_Click()
   '
   Command1.Enabled = False
   '
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
      If FISECOM09.VALIPROCON%(Text12, Text8) < 1 Then GoTo 99
   End If
   '
   ' crea la tabla de mails a enviar
   Call CREACAMPO("", "BANDENVI", "NUME_MAIL", 4, 4, 1)
   Call CREACAMPO("", "BANDENVI", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("", "BANDENVI", "PARA_MAIL", 10, 256)
   Call CREACAMPO("", "BANDENVI", "ASUN_MAIL", 10, 256)
   Call CREACAMPO("", "BANDENVI", "DREL_MAIL", 10, 256, 1)
   Call CREACAMPO("", "BANDENVI", "CCOP_MAIL", 12, 0)
   Call CREACAMPO("", "BANDENVI", "TEXT_MAIL", 12, 0)
   Call CREACAMPO("", "BANDENVI", "SOLI_CONF", 3, 2)
   Call CREACAMPO("", "BANDENVI", "CONF_RECEP", 8, 0, 1)       ' FECHA Y HORA CONFIRMACION DE RECEPCION
   Call CREACAMPO("", "BANDENVI", "FHOR_PREPA", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
   Call CREACAMPO("", "BANDENVI", "USUA_PREPA", 3, 2, 1)       ' USUARIO QUE LO PREPARÓ
   Call CREACAMPO("", "BANDENVI", "FHOR_SOLENVI", 8, 0, 1)     ' FECHA Y HORA PROXIMO ENVIO PROGRAMADO
   Call CREACAMPO("", "BANDENVI", "FREC_REPI", 3, 0, 1)        ' FRECUENCIA DE REPETICION (DIAS)
   Call CREACAMPO("", "BANDENVI", "FEFIN_REPI", 8, 0, 1)       ' FECHA FIN DE REITERACION
   Call CREACAMPO("", "BANDENVI", "FHOR_ULTENVI", 8, 0, 0)     ' FECHA Y HORA ULTIMO ENVIO
   Call CREACAMPO("", "BANDENVI", "CANT_ENVI", 4, 4, 0)        ' CANTIDAD DE VECES ENVIADO
   '
   Call CREACAMPO("", "ADJUNMAI", "NUME_MAIL", 4, 4, 1)
   Call CREACAMPO("", "ADJUNMAI", "NORD_MAIL", 3, 2, 1)
   Call CREACAMPO("", "ADJUNMAI", "FILE_NAME", 10, 256, 0)
   Call CREACAMPO("", "ADJUNMAI", "FILE_DATA", 12, 0)
   '
   NUPROMAI& = 1 + XVALO(VALOBADA("BANDENVI", "MAX(NUME_MAIL)", "NUME_MAIL > 0", "NOMESS"))
   CONDI$ = "NUME_MAIL = " & NUPROMAI&
   Call CreaRegistro("BANDENVI", "NUME_MAIL", NUPROMAI&, "NOMESS")
   Label13 = TRIM$(Str$(NUPROMAI&))
   '
   Call ACTUREGISTRO("BANDENVI", "NUME_MAIL", CONDI$, NUPROMAI&, RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "Codi_Cli", CONDI$, Left$(Label26, 12), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "PARA_MAIL", CONDI$, Left$(Text4, 256), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "ASUN_MAIL", CONDI$, Left$(Text20, 256), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "DREL_MAIL", CONDI$, Left$(Text5, 256), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "CCOP_MAIL", CONDI$, Text3, RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "TEXT_MAIL", CONDI$, Text2, RAFE&, "NOMESS")
      SOLICO% = Check1.Value
   Call ACTUREGISTRO("BANDENVI", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "CONF_RECEP", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
      SOLENVI = FETEXCOR1$(HOY(HOS$)) + " " + Time$
   Call ACTUREGISTRO("BANDENVI", "FHOR_PREPA", CONDI$, SOLENVI, RAFE&, "NOMESS")         ' F-HOR PREPARADO
   Call ACTUREGISTRO("BANDENVI", "USUA_PREPA", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")  ' USUARIO PREPARÓ
      If Label4 <> "" Then SOLENVI = FETEXCOR1(FECHANUM(Label4)) & " " & Text1 & ":00"
   Call ACTUREGISTRO("BANDENVI", "FHOR_SOLENVI", CONDI$, SOLENVI, RAFE&, "NOMESS")       ' PROXIMO ENVIO
   Call ACTUREGISTRO("BANDENVI", "FREC_REPI", CONDI$, XVALO(Text6), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "FEFIN_REPI", CONDI$, FETEXCOR1$(FECHANUM(Text7)), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "FHOR_ULTENVI", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "CANT_ENVI", CONDI$, 0, RAFE&, "NOMESS")
   '
   TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
   QUEDAPEN% = 0: If TRIM$(Mid$(TTXX$, 553, 8)) = "CLIENT" Then QUEDAPEN% = 1
   Call CREACAMPO("", "BANDENVI", "PENDIENV", 3, 2, 1)
   Call ACTUREGISTRO("BANDENVI", "PENDIENV", CONDI$, QUEDAPEN%, RAFE&, "NOMESS")
   '
   ' AGREGAR ADJUNTOS PREFERIDOS
   For KKI% = 1 To List2.ListCount
      If List2.Selected(KKI% - 1) = True Then
         ANAM$ = TRIM$(Mid$(List2.List(KKI% - 1), 65, 128))
         ARUT$ = TRIM$(Mid$(List2.List(KKI% - 1), 193, 128))
         If EXISTE%(ARUT$ + "\" + ANAM$) > 0 Then
            Call ADJUNARCH(NUPROMAI&, ANAM$, ARUT$)
         End If
      End If
   Next KKI%
   ' AGREGAR ADJUNTOS VARIOS
   For KKII% = 1 To List3.ListCount
         TLIS$ = List3.List(KKII% - 1)
         FINAME$ = TRIM$(Left$(TLIS$, 128))
         PPXX% = InStr(129, TLIS$, FINAME$)
         FICOMPLE$ = Mid$(TLIS$, 129, PPXX% - 129)
         If PPXX% > 0 Then
            Call ADJUNARCH(NUPROMAI&, FINAME$, FICOMPLE$)
         End If
   Next KKII%
   '
   ' ANOTA EN HISTORIAL DE SEGUIMIENTOS
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label10)
   If COCLI$ <> "" Then
      FEC% = HOY(HO$)
      TTTT$ = "Preparación de Mail (#" & TRIM$(Str$(NUPROMAI&)) & "#):" + CRLF$ + CRLF$
      TTTT$ = TTTT$ + " - Para: " + TRIM$(Text4) + CRLF$
      TTTT$ = TTTT$ + " - Asunto: " + TRIM$(Text20) + CRLF$ + CRLF$
      If ENVIAINMEDIATO% > 0 Then
            TTTT$ = TTTT$ + "Envío Inmediato - " & Format(Now, "dd/mm/yy - hh:mm") + CRLF$
         Else
            TTTT$ = TTTT$ + "Preparado: " & Format(Now, "dd/mm/yy - hh:mm") + CRLF$
      End If

      Call FISECOM09.ANOTAHISTORIAL(TTTT$, COCLI$)
        PERSCON$ = Combo1.TEXT
        REFACT$ = TRIM$(Text8)
        FEPROG% = FECHANUM(Text12)
      Call FISECOM09.PROGRACTI(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%)     ' ANOTA EN PROXIMO CONTACTO
      '
   End If
   '
   If SHINMAIL09.Visible = True Then
      NORIMAI& = XVALO(SHINMAIL09.Label13)
      If NORIMAI& > 0 Then
        CONDI$ = "NUME_MAIL = " & NORIMAI&
        ' marca como respondido
        Call ACTUREGISTRO("BANDENTRA", "CONF_RECEP", CONDI$, AHORA, RAFE&, "NOMESS")
        ' pone el numero de mail de respuesta
        Call ACTUREGISTRO("BANDENTRA", "NUMAIRESPU", CONDI$, NUPROMAI&, RAFE&, "NOMESS")
        SHINMAIL09.Label139 = TRIM$(Str$(NUPROMAI&))
         ' saca de la lista
        If CONACT09.Visible = True Then
           ILIS& = CONACT09.List3.ListIndex
           If XVALO(Left$(CONACT09.List3.TEXT, 5)) = NORIMAI& Then
                CONACT09.List3.RemoveItem (ILIS&)
                If ILIS& >= 0 Then
                  If ILIS& < CONACT09.List3.ListCount Then
                    CONACT09.List3.Selected(ILIS&) = True
                  End If
                End If
              Else
                Call CONACT09.Command1_Click
           End If
        End If
      End If
   End If
   '
   Call CIERRARCH("*.*")
   Command1.Enabled = True
   '
   If ENVIAINMEDIATO% < 1 Then
      Unload Me
   End If
   '
99 Command1.Enabled = True
   '
End Sub

Sub ADJUNARCH(NROMAI&, FINAME$, RUTARCH$)
   XXX$ = LOADFILE$(RUTARCH$ + "\" + FINAME$)
   If XXX$ <> "" Then
10    NUMAPRO& = XVALO(RANDSTRING$(-6))
      If CantRegistros("ADJUNMAI", "NUME_MAIL = " & NUMAPRO&, "NOMESS") > 0 Then GoTo 10
      Call CreaRegistro("ADJUNMAI", "NUME_MAIL", NUMAPRO&, "NOMESS")
      '
      CONDI$ = "NUME_MAIL = " & NROMAI&
      NORDE& = 1 + XVALO(VALOBADA("ADJUNMAI", "MAX(NORD_MAIL)", CONDI$, "NOMESS"))
      Call ACTUREGISTRO("ADJUNMAI", "NORD_MAIL", "NUME_MAIL = " & NUMAPRO&, NORDE&, RAFE&, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "FILE_NAME", "NUME_MAIL = " & NUMAPRO&, FINAME$, RAFE, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "FILE_DATA", "NUME_MAIL = " & NUMAPRO&, XXX$, RAFE&, "NOMESS")
      Call ACTUREGISTRO("ADJUNMAI", "NUME_MAIL", "NUME_MAIL = " & NUMAPRO&, NROMAI&, RAFE&, "NOMESS")
   End If
   '
End Sub
'
Sub ENVIAMAIL()
      
   MAI_DESTI$ = TRIM$(Text5)
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = TRIM$(Text20)
   Call FRATEXTO("." & Text2, 88, 1, 1)
   TTXX$ = ""
   For KKI% = 1 To CARETEX%
     TTXX$ = TTXX$ + RETEX$(KKI%) + CRLF$
   Next KKI%
   TTYY$ = REPLACAR$(TTXX$, Chr$(13), "~")
   TTYY$ = REPLACAR$(TTXX$, Chr$(10), "")
   TTYY$ = REPLACAR$(TTXX$, "(", "-")
   TTYY$ = REPLACAR$(TTXX$, ")", "-")
   MAI_TEXTO$ = TTYY$
   '
   For i& = 1 To 256: MAI_ADJUN$(i&) = "": Next i&
   JJ& = 0
   On Error Resume Next
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       AA1% = EXISTE%(TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128)))
       If AA1% < 1 Then
           MsgBox "No Existe " + DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
         Else
           JJ& = JJ& + 1
           MAI_ADJUN$(JJ&) = DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
       End If
     End If
   Next U&
   '
   Call GENERAMAIL
   '
   HOII% = HOY(HOS$)
   COCLI$ = TRIM$(FISECOM07.Label26)
   Call FRATEXTO(Text2, 56)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(76)
     Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
     If KKI% = 1 Then
       Call REPLA(TTXX$, HOS$, 1, 8)
       Call REPLA(TTXX$, NOMTER$, 69, 24)
     End If
     TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
   Next KKI%
   '
   YADJU% = 0
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       If YADJU% = 0 Then
            Call REPLA(TTXX$, Space(56), 11, 56)
            Call REPLA(TTXX$, "Adj: " + Mid$(TLXX$, 1, 51), 11, 56)
            YADJU% = 1
          Else
            Call REPLA(TTXX$, "     " + Mid$(TLXX$, 1, 51), 11, 56)
       End If
       TTYY1$ = TTYY1$ + TTXX$ + Chr$(10) + Chr$(13)
     End If
   Next U&
   '
   SQLX$ = "SELECT HIS_CONTAC FROM BSEGCOM"
   SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
   Dim ANOTMP As New ADODB.Recordset
   Set ANOTMP = New ADODB.Recordset
   ANOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With ANOTMP
      If Not (.BOF And .EOF) Then
         ANOTA$ = SAFETEXT$(!HIS_CONTAC)
         ANOTA$ = ANOTA$ + TTXX1$ + TTYY1$
         ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
         !HIS_CONTAC = ANOTA$
         .Update
       End If
   End With
   ANOTMP.Close
   Set ANOTMP = Nothing
   '
   ' marcar prioridad y fecha proximo contacto
   FEPROCON% = DIPOST(DIPOST(DIPOST(HOY(HOS$))))
   CONDICION$ = "CODI_CLI = '" & COCLI$ & "'"
   RCONTA$ = SAFETEXT(VALOBADA("BSEGCOM", "Raso_cli", CONDICION$))
   If Mid$(RCONTA$, 60, 3) <> "!!!" Then
      If COMALTER%("¿ Desea Marcar Contacto con Prioridad ?\\No, Dejar Como Está\Si, Marcarlo") = 2 Then
        RCONTA$ = AJUSTI$(RCONTA$, 59) + "!!!"
        SQLX$ = "SELECT * FROM BSEGCOM"
        SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
        Dim FichaTmp As ADODB.Recordset
        Set FichaTmp = New ADODB.Recordset
        FichaTmp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With FichaTmp
          !raso_cli = RCONTA$
          !FechProxCto = CVDAT(FEPROCON%)
          .Update
        End With
        FISECOM07.Text15 = "!!!"
        FISECOM07.Text12 = FECHATEX(FEPROCON%)
      End If
   End If
   '
   Exit Sub
   '
   ' ACA EMPIEZA EL CODIGO ANTERIOR
   GoTo 10

' esta parte fue hecha provisoriammente para trabajar con flxmail.exe
' hasta encontrar que producia el error
   DRZZ$ = Left$(LUDAT$, 2)
   For KKI% = 0 To 256: MAI_ADJUN$(KKI%) = "": Next KKI%
   KKI% = 0
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
'       DRXX$ = DRZZ$
'       DRYY$ = UCase$(Mid$(TLXX$, 191, 2))
'       'If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
'       If Mid$(DRYY$, 2, 1) = ":" Then DRXX$ = DRYY$
       '
       KKI% = KKI% + 1
       MAI_ADJUN$(KKI%) = TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
       '
     End If
   Next U&
   '
   For U& = 1 To List3.ListCount
     TLXX$ = List3.List(U& - 1)
     DRXX$ = DRZZ$
     DRYY$ = UCase$(Mid$(TLXX$, 129, 2))
     If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
     '
     KKI% = KKI% + 1
     MAI_ADJUN$(KKI%) = DRXX$ + TRIM$(Mid$(TLXX$, 131))
     '
   Next U&
   '
   MAI_DESTI$ = Text4
   MAI_DIREL$ = Text5
   MAI_ASUNT$ = Text20
   Call FRATEXTO(Text2, 96)
   TTXX$ = ""
   For KKI% = 1 To CARETEX%
     TTXX$ = TTXX$ + RETEX$(KKI%) + CRLF$
   Next KKI%
   MAI_TEXTO$ = TTXX$
   '
   Call GENERAMAIL
   GoTo 70
   '
Exit Sub
' ***************************** fin de parte provisoria ****************
   
'///OT-8-999
10 GoSub POROUTLOOK
   If InStr(UCase$(App.EXEName), "TELEF") > 0 Then
      Exit Sub
   End If
   '
   GoTo 70
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.MsgIndex = -1
   MAPIMessages1.MsgSubject = Text20
   MAPIMessages1.RecipAddress = Text5
   MAPIMessages1.MsgNoteText = Text2
   '
   DRZZ$ = Left$(LUDAT$, 2)
   JJ& = 0
   On Error Resume Next
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       'DRXX$ = DRZZ$
       'DRYY$ = TRIM$(Mid$(TLXX$, 193, 2))
       'If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
       'If Mid$(DRYY$, 2, 1) = ":" Then DRXX$ = DRYY$
       '
       AA1% = EXISTE%(TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128)))
       If AA1% < 1 Then
           MsgBox "No Existe " + DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
         Else
           JJ& = JJ& + 1
           MAPIMessages1.AttachmentIndex = JJ& - 1
           MAPIMessages1.AttachmentPathName = DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
           MAPIMessages1.AttachmentName = TRIM$(Mid$(TLXX$, 65, 128))
           If Err Then
             Resume
             Call COMUNI("Imposible Adjuntar Archivo\" + REPLACAR$(TRIM$(Mid$(TLXX$, 65, 128)), "\", "/"))
             MAPISession1.Action = 2
             Exit Sub
           End If
       End If
       '
     End If
   Next U&
   '
   For U& = 1 To List3.ListCount
     TLXX$ = List3.List(U& - 1)
     DRXX$ = DRZZ$
     DRYY$ = UCase$(Mid$(TLXX$, 129, 2))
     If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
       AA1% = EXISTE%(DRXX$ + TRIM$(Mid$(TLXX$, 131)))
       If AA1% < 1 Then
           MsgBox "No Existe " + DRXX$ + TRIM$(Mid$(TLXX$, 131))
         Else
           JJ& = JJ& + 1
           MAPIMessages1.AttachmentIndex = JJ& - 1
           MAPIMessages1.AttachmentPathName = DRXX$ + TRIM$(Mid$(TLXX$, 131))
           MAPIMessages1.AttachmentName = TRIM$(FILESHORTNA$(DRXX$ + TRIM$(Mid$(TLXX$, 131))))
           If Err Then
             Resume
             Call COMUNI("Imposible Adjuntar Archivo\" + REPLACAR$(TRIM$(Mid$(TLXX$, 1, 128)), "\", "/"))
             MAPISession1.Action = 2
             Exit Sub
           End If
       End If
   Next U&
   On Error GoTo 0
   '
   On Error Resume Next
   MAPIMessages1.Action = 3   ' vbMessageSend
   If Err Then
     Call COMUNI("Error de Envío de Correo (" + TRIM$(Str$(Err)) + ")")
     Resume
   End If
   MAPISession1.Action = 2
   '
70 Call GRAPROCON
   '
   Call FRATEXTO(Text2, 56)
   For KKI% = 1 To CARETEX%
     If TRIM$(RETEX$(KKI%)) = "" Then
       INIKK% = KKI% + 1
       GoTo 80
     End If
   Next KKI%
   INIKK% = 1
   '
80 HOII% = HOY(HOS$)
   NPX& = Val(FISECOM07.Label26)
   TTXX$ = Space$(76)
   Call REPLA(TTXX$, HOS$, 1, 8)
   Call REPLA(TTXX$, "Enviamos Mail:", 11, 56)
   Call REPLA(TTXX$, NOMTER$, 69, 24)
   Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
   TTXX$ = Space$(76)
   Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
   '
   TERMINA% = 0
   For KKI% = INIKK% To CARETEX%
     If InStr(UCase$(RETEX$(KKI%)), "BAUER") > 0 Then TERMINA% = 1
     If InStr(UCase$(RETEX$(KKI%)), "SORRENTINO") > 0 Then TERMINA% = 1
     TTXX$ = Space$(76)
     Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
     Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
     If TERMINA% = 1 Then
       Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + Space$(92), 96))
       GoTo 90
     End If
   Next KKI%
   '
90 YADJU% = 0
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       If YADJU% = 0 Then
            TLXX$ = "Adj: " + TLXX$
            YADJU% = 1
          Else
            TLXX$ = "     " + TLXX$
       End If
       TTXX$ = Space$(76)
       Call REPLA(TTXX$, TLXX$, 11, 56)
       Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
     End If
   Next U&
   '
   For U& = 1 To List3.ListCount
     TLXX$ = List3.List(U& - 1)
     If TRIM$(TLXX$) <> "" Then
       If YADJU% = 0 Then
            TLXX$ = "Adj: " + TLXX$
            YADJU% = 1
          Else
            TLXX$ = ".    " + TLXX$
       End If
       TTXX$ = Space$(76)
       Call REPLA(TTXX$, TLXX$, 11, 56)
       Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
     End If
   Next U&
   '
   Call REGAPP("ANOTASEG", MKS$(NPX&) + String$(88, "="))
   Call CIERRARCH("ANOTASEG")
   '
   ' marcar prioridad y fecha proximo contacto
   NPY% = NPX&
   FEPROCON% = DIPOST(DIPOST(DIPOST(HOY(HOS$))))
   XXX$ = FILTERGETRECORD$("BSEGCOM7", 1, MKI$(NPY%))
   RCONTA$ = AJUSTI$(Mid$(XXX$, 3, 62), 62)
   If Mid$(RCONTA$, 60, 3) <> "!!!" Then
      If COMALTER%("¿ Desea Marcar Contacto con Prioridad ?\\No, Dejar Como Está\Si, Marcarlo") = 2 Then
         Call REPLA(RCONTA$, "!!!", 60, 3)
      End If
      Call REPLA(XXX$, RCONTA$, 3, 62)
      FISECOM07.Text15 = "!!!"
   End If
   Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
   Call FILTERPUTRECORD("BSEGCOM7", 1, MKI$(NPY%), XXX$)
   Call CIERRARCH("BSEGCOM7")
   '
Exit Sub
   '
POROUTLOOK:
     '
110  Call CIERRAPLI("Bandeja de Entrada")
     Call CIERRAPLI("Bandeja de salida")
     Call CIERRAPLI("Elementos enviados")
     Call CIERRAPLI("Elementos eliminados")
     '
120  RETVAL = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
     On Error Resume Next
     AppActivate RETVAL
     If Err > 0 Then
        If COMALTER%("Se ha Producido Error " + TRIM$(Str$(Err)) + "\\Continuar\Terminar") = 1 Then
          On Error GoTo 0
          GoTo 120
        End If
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     DIREL$ = TRIM$(Text5)
     For KKII% = 1 To 3:      SendKeys "{ESC}", True: Next KKII%
     MUESTRA$ = TRIM$(Text20)
       Call FRATEXTO(Text2, 88, 1, 1)
       TTXX$ = ""
       For KKI% = 1 To CARETEX%
         TTXX$ = TTXX$ + RETEX$(KKI%) + CRLF$
       Next KKI%
     TTXY$ = REPLACAR$(TTXX$, Chr$(13), "~")
     TTXX$ = REPLACAR$(TTXY$, Chr$(10), "")
     TTXY$ = REPLACAR$(TTXX$, "(", "-")
     TTXX$ = REPLACAR$(TTXY$, ")", "-")
     SK1$ = "%ANM" + DIREL$ + "{TAB}{TAB}" + MUESTRA$ + "{TAB}"
     SK1$ = DIREL$ + "{TAB}{TAB}" + MUESTRA$ + "{TAB}"
     SK2$ = TTXX$
     SK3$ = "%AE" + "%" + "{F4}"
     SK4$ = "^Y" + "{B 2}"
     SendKeys SK1$ + SK2$ + SK3$ + SK4$, True
     '
Return
'
End Sub

Sub CIERRAPLI(TAPLI$)
    '
    VUELTA% = 0
    '
10  On Error Resume Next
    AppActivate TAPLI$
    If Err < 1 Then
       On Error GoTo 0
       SendKeys "%{F4}", True
       VUELTA% = VUELTA% + 1
       If VUELTA% < 3 Then GoTo 10
    End If
    On Error GoTo 0
    '
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   '
   If FISECOM09.VALIPROCON%(Text12, Text8) > 0 Then
      ENVIAINMEDIATO% = 1
      Label13 = ""
      Call Command1_Click
      '
      TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
      If TRIM$(Mid$(TTXX$, 553, 8)) <> "CLIENT" Then
         NUMAI$ = TRIM$(Label13)
         If NUMAI$ <> "" Then
            SHOWMAIL09.Label13 = NUMAI$
            Call SHOWMAIL09.Command1_Click          ' PARA ENVIARLO
         End If
      End If
      Label13 = ""
      ENVIAINMEDIATO% = 0
      '
      Call COMUNI("E-Mail Enviado a '" + REPLACAR$(Text4, "\", "/") + "'.")
      Unload Me                       '  ' PARA IMPIAR LA PANTALLA
   End If
   '
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label10)
   If COCLI$ <> "" Then
      If FISECOM09.VALIPROCON%(Text12, Text8) > 0 Then
         PERSCON$ = Combo1.TEXT
         REFACT$ = TRIM$(Text8)
         FEPROG% = FECHANUM(Text12)
         Call FISECOM09.PROGRACTI(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%)
      End If
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
    If DERACCE%("CFGEMAIL", "Configurar Cuenta de Correo Electronico") < 2 Then
     GoTo 99
    End If
    '
    STPFMAIL.Show 1
    '
99  Command12.Enabled = True
    '
End Sub

'
Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Command16_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text7.TEXT = VDEV$
    End If
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   IXS& = List3.ListIndex
   If IXS& >= 0 Then
     List3.RemoveItem ISX&
   End If
   Command2.Enabled = True
End Sub

Private Sub Command28_Click()
   '
   List4.Clear
   List4.Top = 0
   List4.Left = 0
   List4.Width = 11000
   List4.Height = 8500
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   For U& = 1 To CAMENSA&
     MAPIMessages1.MsgIndex = U& - 1
     AA1$ = MAPIMessages1.MsgOrigDisplayName
     aa2$ = AJUSTI$(MAPIMessages1.MsgOrigAddress, 48)
     TTXX$ = LCase$(aa2$ + AA1$)
     For KKU& = 1 To List4.ListCount
       If List4.List(KKU& - 1) = TTXX$ Then GoTo 20
     Next KKU&
     List4.AddItem TTXX$
     '
20   CARECIP& = MAPIMessages1.RecipCount
     For UU& = 1 To CARECIP&
       MAPIMessages1.RecipIndex = UU& - 1
       aa3$ = MAPIMessages1.RecipDisplayName
       AA4$ = AJUSTI$(MAPIMessages1.RecipAddress, 48)
     TTXX$ = LCase$(AA4$ + aa3$)
     For KKU& = 1 To List4.ListCount
       If List4.List(KKU& - 1) = TTXX$ Then GoTo 30
     Next KKU&
     List4.AddItem TTXX$
     '
       AAA = BBB
30   Next UU&
   Next U&
   MsgBox List4.ListCount
   List4.Refresh
   'MAPIMessages1.MsgSubject = Text20
   'MAPIMessages1.RecipAddress = Text5
   'MAPIMessages1.MsgNoteText = Text2
   '

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   DOCPREFE.Drive1.Drive = Left$(LUDAT$, 2)
   DOCPREFE.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   IXS& = List2.ListIndex
   If IXS& >= 0 Then
     TBUS$ = TRIM$(List2.TEXT)
     JJ& = 0
     For U& = 1 To ULTREG&("DOCPREFE")
       XX$ = REGLEIDO$("DOCPREFE", U&)
       If TRIM$(Left$(XX$, 320)) <> TBUS$ Then
         JJ& = JJ& + 1
         Call GRAREG("DOCPREFE", XX$, JJ&)
       End If
     Next U&
     Call SETULTREG("DOCPREFE", JJ&)
     Call CIERRARCH("DOCPREFE")
     Call BAJALDISCO
     '
     Call CARLISPREF
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   Call MENSERR(24, "Opción No Disponible")
GoTo 99
   '
   XX$ = REGLEIDO$("MAILLST", ULTREG&("MAILLST"))
   MSGNBR& = 1 + CVS(Left$(XX$, 4))
   '
   HOII% = HOY(HOS$)
   HORAI% = Int(HORANUM(Time$))
   REFX$ = HOS$ + " - " + TRIM$(Text4) + " - " + TRIM$(Text20)
   '
   TTXX$ = REGBLAN$("MAILLST")
   Call REPLA(TTXX$, MKS$(MSGNBR&), 1, 4)
   Call REPLA(TTXX$, REFX$, 5, 80)
   Call REPLA(TTXX$, MKS$(Val(Label26)), 85, 4)
   Call REPLA(TTXX$, MKI$(HOII%), 89, 2)
   Call REPLA(TTXX$, MKI$(USNBR%), 91, 2)
   Call REPLA(TTXX$, MKI$(HOII%), 93, 2)
   Call REPLA(TTXX$, MKI$(0), 95, 2)
   Call REPLA(TTXX$, MKI$(-1), 97, 2)
   Call REPLA(TTXX$, MKI$(0), 99, 2)
   Call REGAPP("MAILLST", TTXX$)
   '
   TTXX0$ = REGBLAN$("MAILSPL")
   Call REPLA(TTXX0$, MKS$(MSGNBR&), 1, 4)
   '
   TTXX$ = TTXX0$
   Call REPLA(TTXX$, "SUBJECT", 5, 8)
   Call REPLA(TTXX$, TRIM$(Text20), 13, 244)
   Call REGAPP("MAILSPL", TTXX$)
   '
   TTXX$ = TTXX0$
   Call REPLA(TTXX$, "ADDRESS", 5, 8)
   Call REPLA(TTXX$, TRIM$(Text5), 13, 244)
   Call REGAPP("MAILSPL", TTXX$)
   '
   MSGG$ = Text2
   While Len(MSGG$) > 0
     TTXX$ = TTXX0$
     Call REPLA(TTXX$, "MESSAGE", 5, 8)
     Call REPLA(TTXX$, Left$(MSGG$, 244), 13, 244)
     Call REGAPP("MAILSPL", TTXX$)
     Msg$ = Mid$(Msg$, 245)
   Wend
   '
   DRZZ$ = Left$(LUDAT$, 2)
   TTXX0$ = REGBLAN$("MAILSPL")
   Call REPLA(TTXX0$, MKS$(MSGNBR&), 1, 4)
   Call REPLA(TTXX0$, "ATTACH", 5, 8)
   '
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       DRXX$ = DRZZ$
       DRYY$ = UCase$(Mid$(TLXX$, 191, 2))
       If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
       TTXX$ = TTXX0$
       Call REPLA(TTXX$, DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128)), 13, 244)
       Call REGAPP("MAILSPL", TTXX$)
     End If
   Next U&
   '
   For U& = 1 To List3.ListCount
     TLXX$ = List3.List(U& - 1)
     DRXX$ = DRZZ$
     DRYY$ = UCase$(Mid$(TLXX$, 129, 2))
     If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
     TTXX$ = TTXX0$
     Call REPLA(TTXX$, DRXX$ + TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128)), 13, 244)
     Call REGAPP("MAILSPL", TTXX$)
   Next U&
   '
   Call CIERRARCH("MAILLST")
   Call CIERRARCH("MAILSPL")
   '
   Call GRAPROCON
   '
99 Command5.Enabled = True
End Sub

Private Sub Command7_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   DOCADJUN09.Show 1
   For U& = 1 To DOCADJUN09.List1.ListCount
      fifuna$ = DOCADJUN09.List1.List(U& - 1)
      If EXISTE%(TRIM$(fifuna$)) > 0 Then
         FISHNA$ = fifuna$
         While InStr(FISHNA$, "\") > 0
            FISHNA$ = Mid$(FISHNA$, 2)
         Wend
         If TRIM$(FISHNA$) <> "" Then
            List3.AddItem AJUSTI$(FISHNA$, 128) + fifuna$
         End If
      End If
   Next U&
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
       Label4.Caption = VDEV$
   End If
'   Command9.Enabled = False
'   INSTRUIN$ = " " + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN + "Para la instalación proceder como sigue:" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "1) Abrir este mensaje." + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "2) Realizar doble click sobre el ícono correspondiente al archivo ejecutable adjunto." + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "3) Cliquear sobre botón 'Si'." + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "4) Cliquear sobre el botón 'Unzip'." + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "5) Cliquear sobre el botón 'Aceptar'." + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "6) Cliquear sobre el botón 'Close'." + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
'   INSTRUIN$ = INSTRUIN$ + "Ante cualquier duda favor comunicarse con nuestra mesa de ayuda." + Chr$(13) + Chr$(10)
'   '
'   Text2 = Left$(Text2, Text2.SelStart) + INSTRUIN$ + Mid$(Text2, Text2.SelStart + 1)
'   '
'   Command9.Enabled = True
End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     Call CENTRAFORM(Me)
     ATENTI% = 0
     Text2.SetFocus
   End If
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   Call CARLISPREF
   FEX = DIPOST(DIPOST(HOY(HOS$)))
   Text12 = FECHATEX$(FEX)
   '
End Sub
'
Sub CARLISPREF()
   '
   List2.Clear
   For U& = 1 To ULTREG&("DOCPREFE")
     XX$ = REGLEIDO$("DOCPREFE", U&)
     YY$ = Space$(320)
     Call REPLA(YY$, XX$, 1, 320)
     List2.AddItem YY$
   Next U&
   List2.Refresh
   DoEvents
   '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
End Sub

Private Sub Label26_Change()
    Combo1.Clear
    CODICLIX$ = TRIM$(Label26)
    CONDI$ = "CodiCli = '" & CODICLIX$ & "'"
    '
    SQLX$ = "SELECT * FROM CONTACSCOM WHERE " + CONDI$
    Dim Fichanew As ADODB.Recordset
    Set Fichanew = New ADODB.Recordset
    Set Fichanew = FLEXCONN.Execute(FILTSQL$(SQLX$, "NOMESS"))
    With Fichanew
       Do While Not .EOF
       NOCON$ = TRIM$(SAFETEXT$(!NOME_CTAC))
       If NOCON$ <> "" Then
          Combo1.AddItem NOCON$
       End If
       .MoveNext
       Loop
    End With
    Fichanew.Close
    Set Fichanew = Nothing
    '
    If Combo1.ListCount < 1 Then
       FISECOM09.Label26 = Label26
       For KKI% = 1 To FISECOM09.Text6.Count
          NOCON$ = TRIM$(FISECOM09.Text6(KKI% - 1))
          If NOCON$ <> "" Then
             Combo1.AddItem NOCON$
          End If
       Next KKI%
    End If
    '
    If Combo1.ListCount > 0 Then Combo1.TEXT = Combo1.List(0)
    For KKI% = 1 To Combo1.ListCount
      NOCON$ = Combo1.List(KKI% - 1)
      If Left$(Text4, Len(NOCON$)) = NOCON$ Then
         Combo1.TEXT = NOCON$
      End If
   Next KKI%
   '
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub

Private Sub List1_DblClick()
  '
  TTXX$ = List1.TEXT
  PPXX% = InStr(TTXX$, "(#")
  If PPXX% > 0 Then
     PPYY% = InStr(PPXX%, TTXX$, "#)")
     If PPYY% > PPXX% + 2 Then
        NUMAI$ = TRIM$(Mid$(TTXX$, PPXX% + 2, PPYY% - PPXX% - 2))
        SHOWMAIL09.Label13 = NUMAI$
        SHOWMAIL09.Show 1
        Exit Sub
     End If
  End If
  '
  PPXX% = InStr(TTXX$, "($")
  If PPXX% > 0 Then
     PPYY% = InStr(PPXX%, TTXX$, "$)")
     If PPYY% > PPXX% + 2 Then
        If SHINMAIL09.Visible = True Then
           Call MENSERR(24, "Imposible Mostrar el Mail en Este Contexto.")
           Exit Sub
        End If
        NUMAI$ = TRIM$(Mid$(TTXX$, PPXX% + 2, PPYY% - PPXX% - 2))
        SHINMAIL09.Label13 = NUMAI$
        SHINMAIL09.Show 1
        Exit Sub
     End If
  End If
  '
End Sub

Private Sub List2_DblClick()
   TLXX$ = List2.TEXT
   DRXX$ = Left$(LUDAT$, 2)
   DRYY$ = UCase$(Mid$(TLXX$, 191, 2))
   If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
   On Error Resume Next
   DOCPREFE.Drive1.Drive = DRXX$
   DOCPREFE.Dir1.Path = TRIM$(Mid$(TLXX$, 193, 128))
   If Err Then
      On Error GoTo 0
      List2.Selected(List2.ListIndex) = False
      Call MENSERR(24, "Archivo no Encontrado")
      Exit Sub
   End If
   On Error GoTo 0
   DOCPREFE.Label11 = TRIM$(Mid$(TLXX$, 193, 128))
   DOCPREFE.Label1 = TRIM$(Mid$(TLXX$, 65, 128))
   DOCPREFE.Text1 = TRIM$(Left$(TLXX$, 64))
   '
   For U& = 1 To DOCPREFE.File1.ListCount
     If DOCPREFE.File1.List(U& - 1) = DOCPREFE.Label1 Then
       DOCPREFE.File1.Selected(U& - 1) = True
       Exit For
     End If
   Next U&
   '
   Call DOCPREFE.Command3_Click
End Sub

Private Sub LIST3_DblClick()
   DOCUACT$ = TRIM$(Mid$(List3.TEXT, 129))
   Call MUESTRADOC4(DOCUACT$)
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Sub GRAPROCON()
    '
    NPX& = Val(FISECOM07.Label26)
    NPX1% = NPX&
    RFF$ = RECFILT$("BSEGCOM7", 1, MKI$(NPX1%))
    If Len(RFF$) >= 4 Then
      NPY& = CVS(Left$(RFF$, 4))
      If NPY& <= ULTREG&("BSEGCOM7") Then
        HOII% = HOY(HOS$)
        FEPROCON% = FECHANUM(Text12)
        If FEPROCON% < HOII% Then FEPROCON% = HOII%
        '
        XXX$ = REGLEIDO$("BSEGCOM7", NPY&)
        If CVI(Left$(XXX$, 2)) = NPX1% Then
          Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
          Call GRAREG("BSEGCOM7", XXX$, NPY&)
          Call CIERRARCH("BSEGCOM7")
        End If
        Call COPYSTRU("BSEGCOM7", "BSEGALE")
        JJ& = 0
        FIN& = ULTREG&("BSEGCOM7")
        For U& = 1 To FIN&
           XX$ = REGLEIDO$("BSEGCOM7", U&)
           If CVI(Mid$(XX$, 971, 2)) = 3 Then
              JJ& = JJ& + 1
              Call GRAREG("BSEGALE", XX$, JJ&)
           End If
        Next U&
        Call SETULTREG("BSEGALE", JJ&)
        Call FRESHECHO("BSEGALE")
        '
      End If
    End If
    '
End Sub

Sub MUESTRADOC4(DOCUACT$)
    '
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
    '
End Sub

Private Sub Text2_Change()
    UTIMETEC = Timer
End Sub

Private Sub Text8_Change()
    UTIMETEC = Timer
End Sub

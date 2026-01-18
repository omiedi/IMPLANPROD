VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form MAILMESS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar mensaje de Correo Electrónico"
   ClientHeight    =   8355
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
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame12 
      Caption         =   "Repetir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1280
      Left            =   9900
      TabIndex        =   42
      Top             =   1130
      Width           =   1920
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
         Height          =   250
         Left            =   1635
         TabIndex        =   50
         Top             =   900
         Width           =   150
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         TabIndex        =   49
         Top             =   900
         Width           =   1030
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   2
         Left            =   1080
         TabIndex        =   47
         Top             =   645
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   46
         Top             =   435
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dias"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   0
         Left            =   1080
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   43
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta"
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
         Left            =   50
         TabIndex        =   48
         Top             =   930
         Width           =   585
      End
      Begin VB.Label Label6 
         Caption         =   "Cada"
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
         Left            =   120
         TabIndex        =   44
         Top             =   440
         Width           =   465
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
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   2520
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
      Height          =   270
      Left            =   11180
      TabIndex        =   40
      ToolTipText     =   "Leer e Importar Direcciones de Mails"
      Top             =   820
      Width           =   645
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   11180
      Picture         =   "MAILMESS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   645
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
      Height          =   1520
      Left            =   8560
      TabIndex        =   33
      Top             =   6200
      Width           =   3250
      Begin VB.CommandButton Command9 
         Caption         =   "Instrucc."
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
         Left            =   2100
         TabIndex        =   39
         Top             =   1155
         Width           =   1050
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   315
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
         TabIndex        =   35
         Top             =   1155
         Width           =   975
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
         Left            =   1050
         TabIndex        =   34
         Top             =   1155
         Width           =   1050
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Nuevo Contacto el:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8560
      TabIndex        =   30
      Top             =   7830
      Width           =   3250
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
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
         Left            =   1950
         TabIndex        =   32
         Top             =   80
         Width           =   1030
      End
      Begin VB.CommandButton Command15 
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
         Height          =   250
         Left            =   2980
         TabIndex        =   31
         Top             =   95
         Width           =   150
      End
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
      Left            =   9765
      TabIndex        =   24
      Top             =   130
      Width           =   1320
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         Height          =   250
         Left            =   105
         TabIndex        =   28
         Top             =   315
         Width           =   1110
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
         TabIndex        =   26
         Top             =   630
         Width           =   225
      End
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
      Height          =   2955
      Left            =   8560
      TabIndex        =   19
      Top             =   3130
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
         TabIndex        =   22
         Top             =   2625
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
         TabIndex        =   21
         Top             =   2625
         Width           =   1500
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
         Height          =   2310
         Left            =   105
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   315
         Width           =   3040
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
      Height          =   1280
      Left            =   8560
      TabIndex        =   17
      Top             =   1130
      Width           =   1575
      Begin VB.CheckBox Check1 
         Caption         =   "Confirma Recep."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   210
         TabIndex        =   18
         Top             =   315
         Width           =   1155
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
      Height          =   1730
      Left            =   100
      TabIndex        =   11
      Top             =   6550
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
         Height          =   4005
         Left            =   8085
         TabIndex        =   12
         Top             =   150
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
         Height          =   4110
         Left            =   1155
         TabIndex        =   13
         Top             =   150
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   0
         TabIndex        =   14
         Top             =   285
         Width           =   8300
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
         Y1              =   1700
         Y2              =   1700
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
      Width           =   1170
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         Picture         =   "MAILMESS.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Agrupar para Próximo Envío"
         Top             =   315
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   105
         Picture         =   "MAILMESS.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Enviar Inmediatamente"
         Top             =   315
         Width           =   465
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
      Begin VB.TextBox Text20 
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
         Left            =   1155
         TabIndex        =   5
         Text            =   " "
         Top             =   1260
         Width           =   7050
      End
      Begin VB.TextBox Text4 
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
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   7050
      End
      Begin VB.TextBox Text5 
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
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   735
         Width           =   7050
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
         Top             =   525
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
         Top             =   840
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
      Height          =   4530
      Left            =   105
      TabIndex        =   0
      Top             =   1890
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
         Height          =   4110
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "MAILMESS.frx":075E
         Top             =   315
         Width           =   8100
      End
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line3 
      X1              =   8440
      X2              =   8440
      Y1              =   6540
      Y2              =   8295
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
      TabIndex        =   29
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "MAILMESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
'  GoTo 10   ' para trabajar con EL OBJETO mapi
   '
   DRZZ$ = Left$(LUDAT$, 2)
   For KKI% = 0 To 256: MAI_ADJUN$(KKI%) = "": Next KKI%
   KKI% = 0
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       TLXX$ = List2.List(U& - 1)
       '
       AADJUN$ = TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
       If EXISTE%(AADJUN$) > 0 Then
            KKI% = KKI% + 1
            MAI_ADJUN$(KKI%) = TRIM$(Mid$(TLXX$, 193, 128)) + "\" + TRIM$(Mid$(TLXX$, 65, 128))
          Else
            MsgBox "Imposible Adjuntar " + AADJUN$
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
     '
     KKI% = KKI% + 1
     MAI_ADJUN$(KKI%) = DRXX$ + TRIM$(Mid$(TLXX$, 131))
     '
   Next U&
   '
   MAI_DESTI$ = Text4
   MAI_DIREL$ = Text5
   MAI_ASUNT$ = Text20

   Call GENHTML(Text2)
   MAI_TEXTO$ = LOADFILE$("F:\FLEXOFT\MAILCTRL\HTMAIL.HTM")
   MAI_ADJUN$(0) = "F:\FLEXOFT\MAILCTRL\COMPONENTES\ENCABEZADO_MAIL_FLEXOFT.JPG"
   MAI_FTEXT$ = "HTML"
   MAI_COLEC$ = "": If Check1.Value = 1 Then MAI_COLEC = "SI"
   '
   MAI_REPIT$ = ""
   If XVALO(Text3) > 0 Then
      FREC$ = "D"
      If Option1(1).Value = True Then FREC$ = "H"
      If Option1(2).Value = True Then FREC$ = "M"
      MAI_REPIT$ = Text3 & ";" & FREC$ & ";" & Text6
   End If
   '
   Call GENERAMAIL    ' ("/AUTO")
   GoTo 70
   '
   '
10 GoSub POROUTLOOK
   If InStr(UCase$(App.EXEName), "TELEF") > 0 Then
      Exit Sub
   End If
   '
   GoTo 70
   '
   '
   '
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
   NPX& = Val(FISECOM.Label26)
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
   XXX$ = FILTERGETRECORD$("BSEGCOM", 1, MKI$(NPY%))
   RCONTA$ = AJUSTI$(Mid$(XXX$, 3, 62), 62)
   If Mid$(RCONTA$, 60, 3) <> "!!!" Then
      If COMALTER%("¿ Desea Marcar Contacto con Prioridad ?\\No, Dejar Como Está\Si, Marcarlo") = 2 Then
         Call REPLA(RCONTA$, "!!!", 60, 3)
      End If
      Call REPLA(XXX$, RCONTA$, 3, 62)
      FISECOM.Text15 = "!!!"
   End If
   Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
   Call FILTERPUTRECORD("BSEGCOM", 1, MKI$(NPY%), XXX$)
   Call CIERRARCH("BSEGCOM")
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
     SK2$ = TTXX$
     SK3$ = "%AE" + "%" + "{F4}"
     SK4$ = "^Y" + "{B 2}"
     SendKeys SK1$ + SK2$ + SK3$ + SK4$, True
     '
Return
'
End Sub

Sub GENHTML(TTXX$)
    Dim RENGLO$(1024)
    CAREN% = 0
    FORMUPLA.AutoAnch.FontName = "tahoma"
    FORMUPLA.AutoAnch.FontSize = 9
    Call FRATEXTO(TTXX$, 82, 1, 1)
    AMASCA$ = "F:\FLEXOFT\MAILCTRL\Mascara_mail_flexoft.htm"
    N1% = FILEOPEN%(AMASCA$, 1, 128)
    Do While Not EOF(N1%)
       Line Input #N1%, AX$
       PPXX% = InStr(UCase$(AX$), "COMIENZO>>>>")
       If PPXX% > 0 Then
'         IZQUI$ = Left$(AX$, PPXX% - 1)
'         DERE$ = "</span></p>"
'         DEREFIN$ = "<o:p>"
'         CAREN% = CAREN% + 1
'         RENGLO$(CAREN%) = "<PRE>"
         For J% = 1 To CARETEX
            RENGX$ = RETEX$(J%)
'            RETRI$ = TRIM$(RETEX$(J%))
'            If InStr(RENGX$, RETRI$) > (Len(RENGX$) - Len(RETRI$)) / 2 Then
'               CAREN% = CAREN% + 1
'               RENGLO$(CAREN%) = "<p align=right"
'            End If
            CAREN% = CAREN% + 1
'            RENGLO$(CAREN%) = IZQUI$ + RETEX$(J%) + DERE$
            RENGLO$(CAREN%) = RETEX$(J%)
'            If J% = CARETEX Then
'               RENGLO$(CAREN%) = IZQUI$ + RETEX$(J%) + DEREFIN$
'            End If
         Next J%
         Do While Not EOF(N1%)
            Line Input #N1%, AX$
            If InStr(UCase$(AX$), ">>>>FINAL") > 0 Then
'               CAREN% = CAREN% + 1
'               RENGLO$(CAREN%) = "</PRE>"
               GoTo 19
            End If
         Loop
       End If
       CAREN% = CAREN% + 1
       RENGLO$(CAREN%) = AX$
19  Loop
    Close #N1%
    '
    N2% = FILEOPEN("F:\FLEXOFT\MAILCTRL\HTMAIL.HTM", 2, 128)
    For U% = 1 To CAREN%
       Print #N2%, RENGLO$(U%)
    Next U%
    Close #N2%
'    Stop
         
End Sub
'
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
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text6.TEXT = VDEV$
    End If
End Sub

'
Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
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
     AA2$ = AJUSTI$(MAPIMessages1.MsgOrigAddress, 48)
     TTXX$ = LCase$(AA2$ + AA1$)
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
MsgBox "No Disponible"
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
     MSG$ = Mid$(MSG$, 245)
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
   DOCADJUN.Show 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   INSTRUIN$ = " " + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN + "Para la instalación proceder como sigue:" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "1) Abrir este mensaje." + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "2) Realizar doble click sobre el ícono correspondiente al archivo ejecutable adjunto." + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "3) Cliquear sobre botón 'Si'." + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "4) Cliquear sobre el botón 'Unzip'." + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "5) Cliquear sobre el botón 'Aceptar'." + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "6) Cliquear sobre el botón 'Close'." + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
   INSTRUIN$ = INSTRUIN$ + "Ante cualquier duda favor comunicarse con nuestra mesa de ayuda." + Chr$(13) + Chr$(10)
   '
   Text2 = Left$(Text2, Text2.SelStart) + INSTRUIN$ + Mid$(Text2, Text2.SelStart + 1)
   '
   Command9.Enabled = True
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

Private Sub List2_DblClick()
   TLXX$ = List2.TEXT
   DRXX$ = Left$(LUDAT$, 2)
   DRYY$ = UCase$(Mid$(TLXX$, 191, 2))
   If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
   DOCPREFE.Drive1.Drive = DRXX$
   DOCPREFE.Dir1.Path = DRXX$ + TRIM$(Mid$(TLXX$, 193, 128))
   DOCPREFE.Label11 = DRXX$ + TRIM$(Mid$(TLXX$, 193, 128))
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
   DOCPREFE.Show 1
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
    NPX& = Val(FISECOM.Label26)
    NPX1% = NPX&
    RFF$ = RECFILT$("BSEGCOM", 1, MKI$(NPX1%))
    If Len(RFF$) >= 4 Then
      NPY& = CVS(Left$(RFF$, 4))
      If NPY& <= ULTREG&("BSEGCOM") Then
        HOII% = HOY(HOS$)
        FEPROCON% = FECHANUM(Text12)
        If FEPROCON% < HOII% Then FEPROCON% = HOII%
        '
        XXX$ = REGLEIDO$("BSEGCOM", NPY&)
        If CVI(Left$(XXX$, 2)) = NPX1% Then
          Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
          Call GRAREG("BSEGCOM", XXX$, NPY&)
          Call CIERRARCH("BSEGCOM")
        End If
        Call COPYSTRU("BSEGCOM", "BSEGALE")
        JJ& = 0
        FIN& = ULTREG&("BSEGCOM")
        For U& = 1 To FIN&
           XX$ = REGLEIDO$("BSEGCOM", U&)
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
    On Error Resume Next
    WORDDOCS.OLE1.SourceDoc = DOCUACT$
    TIDOCU$ = "WORD"
    If UCase$(Right$(DOCUACT$, 4)) = ".EML" Then
        TIDOCU$ = "MAIL"
      ElseIf UCase$(Right$(DOCUACT$, 4)) = ".XLS" Then
        TIDOCU$ = "EXCL"
      ElseIf UCase$(Right$(DOCUACT$, 4)) = ".ZIP" Then
        TIDOCU$ = "ZIPF"
    End If
    '
    If TIDOCU$ = "ZIPF" Then
      RETURNVALUE = Shell("C:\Archivos de programa\WinZip\WINZIP32.EXE", 4)
      AppActivate RETURNVALUE: AppActivate RETURNVALUE
      SendKeys "^O", True
      SendKeys DOCUACT$ + Chr$(13), True
      GoTo 99
    End If
      '
    UNILOG$ = "": DOCUNE$ = WORDDOCS.OLE1.SourceDoc
    If Mid$(DOCUNE$, 2, 1) = ":" Then
      UNILOG$ = Left$(DOCUNE$, 2)
      DOCUNE$ = Mid$(DOCUNE$, 3)
    End If
    '
    WORDDOCS.OLE1.Action = 1
    '
    WORDDOCS.OLE1.Action = 6
    WORDDOCS.VScroll1.Min = 0
    WORDDOCS.VScroll1.Value = 0
    ABARRA = WORDDOCS.OLE1.Height - WORDDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
    WORDDOCS.VScroll1.Max = ABARRA
    WORDDOCS.VScroll1.SmallChange = ABARRA / 100
    WORDDOCS.VScroll1.LargeChange = ABARRA / 10
    WORDDOCS.OLE1.Top = 0
    WORDDOCS.OLE1.DoVerb (-1)
    If TIDOCU$ = "WORD" Then
        SendKeys "%{F10}", True  ' esto para documentos en word
      ElseIf TIDOCU$ = "MAIL" Then
        SendKeys Chr$(13) + "%{F10}", True ' esto para mails
      ElseIf TIDOCU$ = "EXCL" Then
        SendKeys Chr$(13) & "%V" & (DOWN8) & "P", True ' ESTO PARA EXCEL
    End If
    '
    WORDDOCS.Show 1
    '
99  On Error GoTo 0
    '
End Sub




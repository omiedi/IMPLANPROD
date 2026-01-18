VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PAD_MAQ04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empleados - A-B-M-L Base de Datos de Empleados"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11640
   Icon            =   "PAD_MAQ04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   55
      Top             =   240
      Width           =   10380
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
         Height          =   315
         Left            =   3480
         Picture         =   "PAD_MAQ04.frx":0442
         TabIndex        =   58
         Top             =   465
         Width           =   175
      End
      Begin VB.CommandButton Command9 
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
         Height          =   330
         Left            =   1935
         Picture         =   "PAD_MAQ04.frx":074C
         TabIndex        =   66
         Top             =   465
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         MaxLength       =   5
         TabIndex        =   65
         Text            =   " "
         Top             =   465
         Width           =   780
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   64
         Top             =   885
         Width           =   7050
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   8325
         TabIndex        =   60
         Top             =   360
         Width           =   1940
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   685
            Picture         =   "PAD_MAQ04.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
            Top             =   250
            Width           =   550
         End
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   125
            Picture         =   "PAD_MAQ04.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Guarda los Cambios Realizados"
            Top             =   250
            Width           =   550
         End
         Begin VB.CommandButton Command20 
            Height          =   500
            Left            =   1260
            Picture         =   "PAD_MAQ04.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Búsqueda por  Domicilio, Nombre, etc"
            Top             =   250
            Width           =   550
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   59
         Top             =   465
         Width           =   405
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6495
         TabIndex        =   57
         Top             =   465
         Width           =   405
      End
      Begin VB.CommandButton Command15 
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
         Height          =   300
         Left            =   6855
         Picture         =   "PAD_MAQ04.frx":23D0
         TabIndex        =   56
         Top             =   465
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Legajo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -105
         TabIndex        =   72
         Top             =   390
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -315
         TabIndex        =   71
         Top             =   825
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
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
         Left            =   1965
         TabIndex        =   70
         Top             =   570
         Width           =   1110
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   5340
         TabIndex        =   69
         Top             =   570
         Width           =   1110
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   68
         Top             =   465
         Width           =   1995
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7110
         TabIndex        =   67
         Top             =   465
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1785
      Width           =   10380
      Begin VB.Frame Frame6 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Datos Gráficos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   5630
         TabIndex        =   73
         Top             =   360
         Width           =   4575
         Begin VB.Frame Frame1 
            BackColor       =   &H00D8D8C0&
            Caption         =   "Fotografía"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2650
            Index           =   2
            Left            =   630
            TabIndex        =   74
            Top             =   450
            Width           =   3300
            Begin VB.PictureBox Picture1 
               AutoRedraw      =   -1  'True
               Height          =   1695
               Index           =   2
               Left            =   210
               ScaleHeight     =   1635
               ScaleWidth      =   2850
               TabIndex        =   77
               ToolTipText     =   "Click para Maximizar"
               Top             =   735
               Width           =   2905
            End
            Begin VB.CommandButton Command53 
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
               Height          =   285
               Index           =   2
               Left            =   2950
               TabIndex        =   76
               Top             =   315
               Width           =   175
            End
            Begin VB.FileListBox File1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1845
               Index           =   2
               Left            =   210
               Pattern         =   "*.JPG"
               TabIndex        =   75
               Top             =   630
               Visible         =   0   'False
               Width           =   2755
            End
            Begin VB.Label Label34 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   210
               TabIndex        =   78
               Top             =   315
               Width           =   2755
            End
            Begin VB.Image Image1 
               Height          =   19995
               Index           =   2
               Left            =   0
               Stretch         =   -1  'True
               Top             =   105
               Visible         =   0   'False
               Width           =   19995
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   5630
         TabIndex        =   50
         Top             =   3810
         Width           =   4575
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            TabIndex        =   52
            Text            =   " "
            Top             =   1200
            Width           =   3480
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   51
            Text            =   " "
            Top             =   630
            Width           =   3480
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N.Afil.:"
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
            Left            =   -1080
            TabIndex        =   54
            Top             =   1300
            Width           =   1905
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Jub.:"
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
            Left            =   -1080
            TabIndex        =   53
            Top             =   700
            Width           =   1905
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Datos Laborales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   210
         TabIndex        =   37
         Top             =   3810
         Width           =   5295
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3900
            TabIndex        =   44
            Top             =   345
            Width           =   1125
         End
         Begin VB.CommandButton Command18 
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
            Height          =   300
            Left            =   5010
            Picture         =   "PAD_MAQ04.frx":26DA
            TabIndex        =   43
            Top             =   345
            Width           =   150
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1485
            TabIndex        =   42
            Top             =   360
            Width           =   1125
         End
         Begin VB.CommandButton Command11 
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
            Height          =   300
            Left            =   2595
            Picture         =   "PAD_MAQ04.frx":29E4
            TabIndex        =   41
            Top             =   360
            Width           =   150
         End
         Begin VB.CommandButton Command19 
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
            Height          =   315
            Left            =   1890
            Picture         =   "PAD_MAQ04.frx":2CEE
            TabIndex        =   40
            Top             =   1200
            Width           =   175
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1485
            TabIndex        =   39
            Top             =   1200
            Width           =   405
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1485
            TabIndex        =   38
            Text            =   " "
            Top             =   780
            Width           =   3690
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Egreso:"
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
            TabIndex        =   49
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Ingreso:"
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
            Left            =   480
            TabIndex        =   48
            Top             =   455
            Width           =   930
         End
         Begin VB.Label Label23 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2190
            TabIndex        =   47
            Top             =   1200
            Width           =   2985
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
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
            Left            =   315
            TabIndex        =   46
            Top             =   1305
            Width           =   1110
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tarea:"
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
            Left            =   -480
            TabIndex        =   45
            Top             =   885
            Width           =   1905
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Datos Personales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   210
         TabIndex        =   17
         Top             =   360
         Width           =   5295
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   27
            Top             =   1200
            Width           =   1125
         End
         Begin VB.CommandButton Command16 
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
            Height          =   300
            Left            =   2580
            Picture         =   "PAD_MAQ04.frx":2FF8
            TabIndex        =   26
            Top             =   1200
            Width           =   150
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   25
            Text            =   " "
            Top             =   2880
            Width           =   3690
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   24
            Text            =   " "
            Top             =   2460
            Width           =   3690
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   23
            Text            =   " "
            Top             =   2040
            Width           =   345
         End
         Begin VB.TextBox Text25 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   22
            Top             =   360
            Width           =   3660
         End
         Begin VB.TextBox Text26 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   21
            Text            =   " "
            Top             =   780
            Width           =   330
         End
         Begin VB.TextBox Text29 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   20
            Text            =   " "
            Top             =   1605
            Width           =   3690
         End
         Begin VB.CommandButton Command23 
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
            Height          =   300
            Left            =   1800
            Picture         =   "PAD_MAQ04.frx":3302
            TabIndex        =   19
            Top             =   780
            Width           =   150
         End
         Begin VB.CommandButton Command34 
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
            Height          =   300
            Left            =   1800
            Picture         =   "PAD_MAQ04.frx":360C
            TabIndex        =   18
            Top             =   2040
            Width           =   150
         End
         Begin VB.Label Label19 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2145
            TabIndex        =   36
            Top             =   780
            Width           =   2985
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio:"
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
            Left            =   -495
            TabIndex        =   35
            Top             =   2565
            Width           =   1905
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "F. Nacim.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   105
            TabIndex        =   34
            Top             =   1320
            Width           =   1290
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad:"
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
            Left            =   -495
            TabIndex        =   33
            Top             =   3000
            Width           =   1905
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil:"
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
            Left            =   -495
            TabIndex        =   32
            Top             =   2145
            Width           =   1905
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Documento:"
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
            Left            =   -525
            TabIndex        =   31
            Top             =   465
            Width           =   1905
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo:"
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
            Left            =   -525
            TabIndex        =   30
            Top             =   885
            Width           =   1905
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nacionalidad:"
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
            Left            =   -495
            TabIndex        =   29
            Top             =   1725
            Width           =   1905
         End
         Begin VB.Label Label22 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   28
            Top             =   2040
            Width           =   2985
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D8D8C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   12
      Top             =   8200
      Width           =   10215
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   975
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton Command12 
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
         Height          =   300
         Left            =   2190
         Picture         =   "PAD_MAQ04.frx":3916
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cat.:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00E0E0E0&
         X1              =   5280
         X2              =   5280
         Y1              =   240
         Y2              =   3600
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   5280
         Y1              =   240
         Y2              =   3600
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Index           =   0
      Left            =   10710
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":3C20
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1995
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   6825
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":3F2A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprime Ficha del Empleado"
         Top             =   5010
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":4594
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Emisión de Listados"
         Top             =   4320
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":489E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5880
         Width           =   650
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":4CE0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":4FEA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Baja/Restauración Lógica del Legajo"
         Top             =   3570
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":52F4
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":543E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1050
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PAD_MAQ04.frx":5748
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta de Nuevo Legajo"
         Top             =   2880
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "PAD_MAQ04.frx":5892
         Stretch         =   -1  'True
         Top             =   7100
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   11760
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NuevEmp 
         Caption         =   "Nuevo Legajo"
      End
      Begin VB.Menu EditEmp 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaEmp 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Baj_Emp 
         Caption         =   "Baja/Restauración Lógica del Legajo"
      End
      Begin VB.Menu PLh2 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiEmp 
      Caption         =   "Edición"
      Begin VB.Menu DeHaCaEmp 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu php5a 
         Caption         =   "-"
      End
      Begin VB.Menu BusEmp 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu ConLiEmp 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu ConfEmp 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenEmp 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu PHP35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrupEmp 
         Caption         =   "Grupos de Empleados"
      End
      Begin VB.Menu mnuCatego 
         Caption         =   "Categorias de Empleados"
      End
      Begin VB.Menu mnuSalNomBas 
         Caption         =   "Salarios Nominales Basicos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParInd 
         Caption         =   "Parametros Individuales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemOrd 
         Caption         =   "Remuneraciones Ordinarias"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRetLegApo 
         Caption         =   "Retenciones Legales y Aportes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAsiSubFam 
         Caption         =   "Asignaciones y Subsidios Familiares"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "PAD_MAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODIFIC% 'MARCA DE MODIFICACION
Public MODIFICNOT% 'MARCA PARA LAS NOTAS DE EMPLEADOS
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)

Private Sub AyuCli_Click()
  '
  Call Command4_Click
  '
End Sub

Private Sub Baj_Emp_Click()
  '
  Call Command24_Click
  '
End Sub

Private Sub BusEmp_Click()
  '
  Call Command20_Click
  '
End Sub

Private Sub Command1_Click()
  '
  If Val(Text1) > 0 Then
    Call GRAEMP(Val(Text1))
    Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
    Call GRAFOTO 'GRABA FOTO DEL EMPLEADO
    Call FRESHECHO("PADREMP")
    MODIFIC% = 0
    MODIFICNOT% = 0
  End If
  '
End Sub

Private Sub Command10_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("DATGREM1")
     VDEVU$ = TRIM$(VALACT1$("DATGREM1"))
     If VDEVU$ <> "" Then
        Text6.TEXT = VALACT1$("DATGREM1")
     End If
  End If
End Sub

Private Sub Command11_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text9.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command12_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command13_Click()
  '
  Command13.Enabled = False
  Call BLANARCH("LIPADRON")
  '
  For J& = 1 To ULTREG&("PADRON")
    X$ = REGLEIDO$("PADRON", J&)
    Call REGAPP("LIPADRON", X$)
  Next J&
  '
  Call SETULTREG("LIPADRON", J&)
  Call CIERRARCH("LIPADRON")
  '
  Call FINAL("*LIPADRON")
  Command13.Enabled = True
  '
End Sub

Private Sub Command14_Click()
  '
  OPADMAQ04.Show 1
  '
End Sub

Private Sub Command16_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text10.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command18_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
     End If
  End If
End Sub

Private Sub Command19_Click()
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("CATEGO")
     VDEVU$ = TRIM$(VALACT1$("CATEGO"))
     If VDEVU$ <> "" Then
        Text16.TEXT = VALACT1$("CATEGO")
     End If
  End If
End Sub

Private Sub Command2_Click()
  '
  If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
     Exit Sub
  End If
  '
  Command2.Enabled = False
    A$ = TRIM$(Text1.TEXT)
10  Text1.TEXT = ""
    Text1.TEXT = A$
    If TRIM$(Text3.TEXT) = "" Then GoTo 10
  Command2.Enabled = True
  '
End Sub

Private Sub Command20_Click()
  '
  Static ARGU$
  Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Empleados por Texto", ARGU$)))
   If ARGU$ <> "" Then
      Screen.MousePointer = 11
      JJ& = 0
      Call COPYSTRU("PADREMP", "SELPAD")
      '
      JJ& = 0
      XX$ = Space(256)
      For J& = 1 To ULTREG&("PADREMP")
        VATEX$ = REGLEIDO$("PADREMP", J&)
         If InStr(VATEX$, ARGU$) > 0 Then
            Call REPLA(XX$, VATEX$, 1, 256)
            JJ& = JJ& + 1
            Call GRAREG("!SELPAD", XX$, JJ&)
          End If
      Next J&
    '
    Call SETULTREG("!SELPAD", JJ&)
    Call CIERRARCH("!SELPAD")
    '
    Screen.MousePointer = 1
    If JJ& < 1 Then
        Call MENSERR(24, "No se Encontraron Coincidencias")
        GoTo 10
      ElseIf JJ& = 1 Then
        CI2% = TRIM$(Str$(CVI(Mid$(REGLEIDO$("!SELPAD", 1), 1, 2))))
        Text1.TEXT = Str$(CI2%)
      Else
        Call FRESHECHO("!SELPAD")
        Call SELECHO("!SELPAD")
        NCLIX$ = VALACT1$("!SELPAD")
        If NCLIX$ <> "" Then
          Text1 = NCLIX$
        End If
    End If
  End If
  Command20.Enabled = True
  '
End Sub

Private Sub Command21_Click()
  '
  Command21.Enabled = False
  
  If DERACCE%("NUOPER", "Alta de Nuevo Operario") < 1 Then GoTo 99
     
  'VDEF$ TIENE QUE DAR EL PROX. NUM DE LEGAJO
  VDEF$ = Space(32)
  For JJ& = 1 To ULTREG("PADREMP")
  XX$ = REGLEIDO$("PADREMP", JJ&)
    NLEG% = CVI(Mid$(XX$, 1, 2))
    If NLEG% > PRONU% Then
      PRONU% = NLEG%
    End If
  Next JJ&
  Call REPLA(VDEF$, MKI$(PRONU% + 1), 1, 2)
9 OBJ$ = OBJPLA$("ALTEMP", VDEF$)
  If OBJ$ = "" Then GoTo 99
  VDEF$ = OBJ$
  '
  NULEG% = CVI(Mid$(OBJ$, 1, 2))
  DESCRI$ = TRIM$(Mid$(OBJ$, 3, 30))
  NUGRU% = CVI(Mid$(OBJ$, 33, 2))
  '//////////////VALIDACIONES////////////
  If Val(NULEG%) <= 0 Then
    Call COMUNI("Número de Legajo no válido.")
    GoTo 9
  End If
  '
  If TRIM(DESCRI$) = "" Then
    Call COMUNI("Apellido y Nombre no válido.")
    GoTo 9
  End If
  '
  For J& = 1 To ULTREG("PADREMP")
    XX$ = REGLEIDO$("PADREMP", J&)
    LEG% = CVI(Mid$(XX$, 1, 2))
    NOAP$ = TRIM(Mid$(XX$, 3, 30))
    If NULEG% = LEG% Then
      Call COMUNI("Número de Legajo no válido o Pre-Existente.")
      GoTo 9
    End If
    If NOAP$ = DESCRI$ Then
      Call COMUNI("Empleado Pre-Existente.")
      GoTo 9
    End If
  Next J&
  '
  If NUGRU% <= 0 Then
    Call COMUNI("Grupo de Empleado no Válido.")
    GoTo 9
  End If
  'LIMPIO LOS TEXT
  Text1 = ""
  'GRABO EL REGISTRO
  X$ = REGBLAN("PADREMP")
  ULTRE& = ULTREG&("PADREMP")
  Call REPLA(X$, MKI$(NULEG%), 1, 2)
  Call REPLA(X$, DESCRI$, 3, 30)
  Call REPLA(X$, Chr$(NUGRU%), 49, 1)
  Call REPLA(X$, Chr$(1), 48, 1) 'ESTADO ACTIVO
  Call GRAREG("PADREMP", X$, ULTRE& + 1)
  Call SETULTREG("PADREMP", ULTRE& + 1)
  'GRABO EN ARCHIVOS AUXILIARES
  PADA1$ = Space(64)
  Call REPLA(PADA1$, X$, 1, 64)
  Call GRAREG("PADRON", PADA1$, ULTRE& + 1)
  PADA2$ = Space(96)
  Call GRAREG("PADAUX", PADA2$, ULTRE& + 1)
  PADA3$ = Space(48)
  Call GRAREG("PADAUX2", PADA3$, ULTRE& + 1)
  '
  Text1 = NULEG%
  '
99 Command21.Enabled = True
End Sub
Private Sub Command22_Click()
  VTB% = VENTABU%("CONCOPCI")
End Sub

Private Sub Command23_Click()
   '
   If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
   If TRIM(Text1) <> "" Then
     Call SELECHO("SEXTAB")
     NSEX$ = TRIM$(VALACT1$("SEXTAB"))
     If TRIM$(NSEX$) = "" Then Exit Sub
     Text26.TEXT = NSEX$
   End If
   '
End Sub

Private Sub Command24_Click()
  '
  If DERACCE%("BAJAOPE", "Baja de Operario") < 1 Then Exit Sub
  '
  If Val(Text1) > 0 Then
   If Val(Text19) = 1 Then
    If COMALTER%("¿Esta Seguro de Dar de Baja el Empleado: '" + TRIM$(Text3.TEXT) + "'?\\Dar de Baja\Cancelar") = 1 Then
     Screen.MousePointer = 11
     MODIFIC% = 0
     Call BAJEMP(Val(Text1)) 'BAJA DE EMPLEADO
     Screen.MousePointer = 1
     Exit Sub
    End If
   Else
    If COMALTER%("¿Esta Seguro de Restaurar el Empleado: '" + TRIM$(Text3.TEXT) + "'?\\Restaurar\Cancelar") = 1 Then
     Screen.MousePointer = 11
     MODIFIC% = 0
     Call ACTEMP(Val(Text1)) 'ALTA DE EMPLEADO
     Screen.MousePointer = 1
     Exit Sub
    End If
   End If
  End If
  '
End Sub

Private Sub Command25_Click()
  '
  PAD_MAQ04.PRINTFORM
  '
End Sub

Private Sub Command3_Click()
  '
  If Val(Text1) > 32767 Then
    GoTo 10
  End If
  If TRIM(RECFILT("PADREMP", 1, MKI$(Val(Text1)))) <> "" Then
    If MODIFIC% = 1 Then
      If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
        Screen.MousePointer = 11
        MODIFIC% = 0
       Call GRAEMP(Val(Text1))
       Screen.MousePointer = 1
      End If
    End If
    If MODIFICNOT% = 1 Then
      Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
    End If
  End If
  '
10 Call CIERRARCH("*.*")
   Unload Me
  '
End Sub

Private Sub Command34_Click()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If TRIM(Text1) <> "" Then
     Call SELECHO("ESTCIV")
     VDEVU$ = TRIM$(VALACT1$("ESTCIV"))
     If VDEVU$ <> "" Then
        Text2.TEXT = VALACT1$("ESTCIV")
     End If
  End If
  '
End Sub

Private Sub Command4_Click()
  '
  Call HELPONLINE("PADMAQ")
  '
End Sub

Private Sub Command53_Click(Index As Integer)
   '
   NLEG% = Val(TRIM$(Text1.TEXT))
   If NLEG% < 1 Then Exit Sub
   If TRIM(RECFILT$("PADREMP", 1, MKI$(NLEG%))) = "" Then Exit Sub
   '
   Command53(Index).Enabled = False
   '
10 CommonDialog1.InitDir = ULODATOS$ & "\"
   CommonDialog1.DialogTitle = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Archivo de Imagen(JPG; BMP; WMF)|*.JPG; *.BMP; *.WMF|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   '
   RUELE$ = Left$(CommonDialog1.FileName, 2)
   If RUELE$ <> ULODATOS$ Then
      Call MENSERR(24, "Ruta de Archivo no Válida.\ Los Archivos Gráficos Deben \ Alojarse en el Disco de Trabajo del Sistema.")
      CommonDialog1.InitDir = ULODATOS$ & "\"
      GoTo 10
   End If
   '
   Label34(Index) = UCase$(CommonDialog1.FileName)
   '
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command53(Index).Enabled = True
   On Error GoTo 0
   '
End Sub

Private Sub Command9_Click()
   '
   If MODIFIC% = 1 And Val(Text1) > 0 And Val(Text1) < 32767 Then
    If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
      Screen.MousePointer = 11
      MODIFIC% = 0
      Call GRAEMP(Val(Text1))
      Screen.MousePointer = 1
    End If
   End If
   '
   If MODIFICNOT% = 1 Then
    Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
   End If
   '
   Call SELECHO("PADREMP")
   NLEG% = Val(TRIM$(VALACT1$("PADREMP")))
   If NLEG% < 1 Then Exit Sub
   '
   NC% = NLEG%
   Text1.TEXT = TRIM$(Str$(NC%))
   Call CARDATEMP(NLEG%) 'CARGA DATOS DEL EMPLEADO
   'Call CARFAMCAR(NLEG%) 'CARGA FAMILIARES A CARGO
'   Call CARCONCEOP(NLEG%) 'CARGA DE CONCEPTOS OPCIONALES
   'Call CARNOTAEMP 'CARGA ANOTACIONES DEL EMPLEADO
   MODIFIC% = 0
   MODIFICNOT% = 0
   '
End Sub
'Sub CARFAMCAR(LEGA%) 'CARGA DE FAMILIARES X Nº LEG.
'  '
'  List3.Clear
'  '
'  Call TRALOFIL("FAMCAR", 1, MKI$(LEGA%))
'  For J& = 1 To ULTREG&("!FAMCAR")
'      XX$ = REGLEIDO$("!FAMCAR", J&)
'      NOAP$ = Mid$(XX$, 3, 25)
'      FENA = CVI(Mid$(XX$, 28, 2))
'      PAREN$ = Mid$(XX$, 30, 10)
'      DOCUFA$ = Mid$(XX$, 40, 11)
'      YY$ = Space(68)
'      Call REPLA(YY$, NOAP$, 1, 27)
'      Call REPLA(YY$, FECHATEX$(FENA), 30, 10)
'      Call REPLA(YY$, PAREN$, 41, 10)
'      Call REPLA(YY$, DOCUFA$, 53, 11)
'      List3.AddItem YY$
'  Next J&
'  '
'End Sub
Sub CARDATEMP(LEGA%) 'CARGA DE EMPLEADOS
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(LEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  NOAP$ = TRIM$(Mid$(XX$, 3, 30))
  TIDO$ = Mid$(XX$, 33, 13)
  FEIN = CVI(Mid$(XX$, 46, 2))
  STA = Asc(Mid$(XX$, 48, 1))
  GRU = Asc(Mid$(XX$, 49, 1))
  CATE$ = TRIM$(Mid$(XX$, 54, 9))
  FENAC = CVI(Mid$(XX$, 63, 2))
    DOMI$ = Mid$(XX$, 65, 30) 'DOMICILIO
  LOCA$ = Mid$(XX$, 95, 30) 'LOCALIDAD
  NACIO$ = Mid$(XX$, 125, 10) 'NACIONALIDAD
  ECIV$ = Mid$(XX$, 135, 10) 'EST.CIVIL
  FEGRE = CVI(Mid$(XX$, 147, 2)) 'F.EGRESO
  SEX$ = Mid$(XX$, 149, 1) 'SEXO
    TARE$ = Mid$(XX$, 161, 20) 'TAREA
  CAJU$ = Mid$(XX$, 181, 16) 'CAJA DE JUBILACION
  NUAFI$ = Mid$(XX$, 197, 12) 'NUMERO DE AFILIADO
  '
  Text3 = NOAP$
  Text25 = TIDO$
  Text9 = FECHATEX(FEIN)
  Text6 = CStr(GRU)
  Text10 = FECHATEX(FENAC)
  Text16 = CATE$
  Text19 = CStr(STA)
    Text4 = DOMI$
  Text7 = LOCA$
  Text29 = NACIO$
  Text2 = ECIV$
  Text12 = FECHATEX(FEGRE)
  Text26 = SEX$
    Text13 = CAJU$
  Text14 = NUAFI$
  Text8 = TARE$
  'CARGA FOTO DEL EMPLEADO
  YY1$ = RECFILT$("FOTOEMP", 1, MKI$(LEGA%))
  J2& = CVS(YY1$)
  XZ$ = REGLEIDO$("FOTOEMP", J2&)
  FOTO2$ = TRIM$(Mid$(XZ$, 3, 510))
  Label34(2) = TRIM$(FOTO2$)
'   For J& = 1 To ULTREG("PADREMP")
'     XX$ = REGLEIDO$("PADREMP", J&)
'     If TRIM$(XX$) = "" Then
'      Call BLANFORMU
'      Exit Sub
'     End If
'     NLEG1% = CVI(Mid$(XX$, 1, 2))
'     If NLEG1% = LEGA% Then
'       NOAP$ = Mid$(XX$, 3, 30)
'       TIDO$ = Mid$(XX$, 33, 13)
'       FEIN = CVI(Mid$(XX$, 46, 2))
'       STA = Asc(Mid$(XX$, 48, 1))
'       GRU = Asc(Mid$(XX$, 49, 1))
'       CATE$ = TRIM$(Mid$(XX$, 54, 9))
'       FENAC = CVI(Mid$(XX$, 63, 2))
'       '
'       Text3 = NOAP$
'       Text25 = TIDO$
'       Text9 = FECHATEX(FEIN)
'       Text6 = CStr(GRU)
'       Text10 = FECHATEX(FENAC)
'       Text16 = CATE$
'       Text19 = CStr(STA)
'       '
'       GoTo 10
'       FOUND1% = 1
'     End If
'   Next J&
'   '
'   If FOUND1% = 0 Then
'    Call BLANFORMU
'    Exit Sub
'   End If
'  '
'10 YY$ = REGLEIDO$("PADAUX", J&)
'  '
'  DOMI$ = Mid$(YY$, 1, 30) 'DOMICILIO
'  LOCA$ = Mid$(YY$, 31, 30) 'LOCALIDAD
'  NACIO$ = Mid$(YY$, 61, 10) 'NACIONALIDAD
'  ECIV$ = Mid$(YY$, 71, 10) 'EST.CIVIL
'  FEGRE = CVI(Mid$(YY$, 83, 2)) 'F.EGRESO
'  SEX$ = Mid$(YY$, 85, 1) 'SEXO
'  '
'  Text4 = DOMI$
'  Text7 = LOCA$
'  Text29 = NACIO$
'  Text2 = ECIV$
'  Text12 = FECHATEX(FEGRE)
'  Text26 = SEX$
'  '
'  GoTo 20
'  '
'20 ZZ$ = REGLEIDO$("PADAUX2", J&)
'  '
'  TARE$ = Mid$(ZZ$, 1, 20) 'TAREA
'  CAJU$ = Mid$(ZZ$, 21, 16) 'CAJA DE JUBILACION
'  NUAFI$ = Mid$(ZZ$, 37, 12) 'NUMERO DE AFILIADO
'  '
'  Text13 = CAJU$
'  Text14 = NUAFI$
'  Text8 = TARE$
'  '
End Sub
Sub GRAEMP(NLEGA%) 'GRABA EMPLEADO
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    NLEG% = Val(Text1)
    NOAP$ = Text3
    If TRIM$(NOAP$) = "" Then
      Call MENSERR(24, "Apellido y Nombre no Válido.")
      Exit Sub
    End If
    TIDO$ = Text25
    'validación del numero de documento
    FEIN% = FECHANUM(TRIM(Text9))
    GRU% = Val(Text6)
    If TRIM$(ECOARCH$("DATGREM1", MKI$(GRU%))) = "" Then
      Call MENSERR(24, "Grupo de Empleados no Válido.")
      Exit Sub
    End If
    CATE$ = TRIM(Text16)
    If TRIM$(ECOARCH$("CATEGO", CATE$)) = "" Then
      Call MENSERR(24, "Categoria no Válida.")
      Exit Sub
    End If
    FENA% = FECHANUM(TRIM(Text10))
    DOMI$ = TRIM$(Text4)
    LOCA$ = TRIM$(Text7)
    NACIO$ = TRIM$(Text29)
    ECIV$ = TRIM$(Text2)
    If TRIM$(ECOARCH("ESTCIV", MKI$(Val(ECIV$)))) = "" Then
      Call MENSERR(24, "Estado Civil no Válido.")
      Exit Sub
    End If
    FEGRE% = FECHANUM(TRIM(Text12))
    SEX$ = TRIM$(Text26)
    If TRIM$(ECOARCH("SEXTAB", SEX$)) = "" Then
      Call MENSERR(24, "Sexo no Válido.")
      Exit Sub
    End If
    TARE$ = TRIM$(Text8)
    CAJUB$ = TRIM$(Text13)
    NUAFI$ = TRIM$(Text14)
    '
    Call REPLA(XX$, MKI$(NLEG%), 1, 2)
    Call REPLA(XX$, NOAP$, 3, 30)
    Call REPLA(XX$, TIDO$, 33, 13)
    Call REPLA(XX$, MKI$(FEIN%), 46, 2)
    Call REPLA(XX$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    Call REPLA(XX$, Chr$(GRU%), 49, 1)
    Call REPLA(XX$, CATE$, 54, 9)
    Call REPLA(XX$, MKI$(FENA%), 63, 2)
    Call REPLA(XX$, DOMI$, 65, 30)
    Call REPLA(XX$, LOCA$, 95, 30)
    Call REPLA(XX$, NACIO$, 125, 10)
    Call REPLA(XX$, ECIV$, 135, 10)
    Call REPLA(XX$, MKI$(FEIN%), 145, 2)
    Call REPLA(XX$, MKI$(FEGRE%), 147, 2)
    Call REPLA(XX$, SEX$, 149, 1)
    Call REPLA(XX$, TARE$, 161, 20)
    Call REPLA(XX$, CAJUB$, 181, 16)
    Call REPLA(XX$, NUAFI$, 197, 12)
    'GRABO EN PADREMP
    Call GRAREG("PADREMP", XX$, J1&)
    'GRABO EN ARCHIVOS AUXILIARES
    PADA1$ = Mid$(XX$, 1, 64)
    Call GRAREG("PADRON", PADA1$, J1&)
    PADA2$ = Mid$(XX$, 65, 96)
    ECOECIV$ = TRIM(Label22)
    Call REPLA(PADA2$, ECOECIV$, 71, 10)
    Call GRAREG("PADAUX", PADA2$, J1&)
    PADA3$ = Mid$(XX$, 161, 48)
    Call GRAREG("PADAUX2", PADA3$, J1&)
    '
  End If
  '
End Sub
Sub ACTEMP(NLEGA%) 'ESTADO ACTIVO AL  EMPLEADO
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    Call REPLA(XX$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    '
    Call GRAREG("PADREMP", XX$, J1&)
    Text19 = 1
    'GRABO ESTADO EN ARCHIVO AUXILIAR
    XX1$ = REGLEIDO$("PADRON", J1&)
    Call REPLA(XX1$, Chr$(1), 48, 1) 'ESTADO ACTIVO
    Call GRAREG("PADRON", XX1$, J1&)
    '
  End If
  '
End Sub
Sub BAJEMP(NLEGA%) 'BAJA DE EMPLEADO
  '
  XX1$ = RECFILT$("PADREMP", 1, MKI$(NLEGA%))
  J1& = CVS(XX1$)
  '
  XX$ = REGLEIDO$("PADREMP", J1&)
  If XX$ <> "" Then
    '
    Call REPLA(XX$, Chr$(0), 48, 1) 'ESTADO BAJA
    '
    Call GRAREG("PADREMP", XX$, J1&)
    Text19 = 0
    'GRABO ESTADO EN ARCHIVO AUXILIAR
    XX1$ = REGLEIDO$("PADRON", J1&)
    Call REPLA(XX1$, Chr$(0), 48, 1) 'ESTADO BAJA
    Call GRAREG("PADRON", XX1$, J1&)
    '
  End If
  '
End Sub

Private Sub ConGenEmp_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub ConLiEmp_Click()
  '
  Call Command13_Click
  '
End Sub

Private Sub DeHaCaEmp_Click()
  '
  Call Command2_Click
  '
End Sub

Private Sub EditEmp_Click()
  '
  Call Command9_Click
  '
End Sub

Private Sub Exit_Click()
  '
  Call Command3_Click
  '
End Sub
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
    Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  '
  Call BLANFORMU
  '
  Call CREATAB
  Call HACECONCEP 'ARMA ECO DE CONCEPTOS
  Call APACTPADRO 'APAREA PADRON DE EMPLEADOS CON ACTIVA.DAT
  Call APAPADREMP 'APAREA PADREMP.DAT CON PADRON.DAT, PADAUX.DAT, PADAUX2.DAT.
  Call APAREMBA   'APARECA PADRON DE EMPLEADOS CON REMBAS.DAT
  '
  TOPEINI(2) = Picture1(2).Top
  IZQUINI(2) = Picture1(2).Left
  ANCHINI(2) = Picture1(2).Width
  ALTUINI(2) = Picture1(2).Height
  '
End Sub
'DESDE ACA COPIADO - ¿PARA QUE SIRVE?
Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub


'HASTA ACA
'Private Sub Form_Unload(Cancel As Integer)
'  '
'  If TRIM(RECFILT("PADREMP", 1, MKI$(Val(Text1)))) <> "" Then
'    If MODIFIC% = 1 Then
'      If COMALTER%("Se han Modificado Datos del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") = 1 Then
'        Screen.MousePointer = 11
'        MODIFIC% = 0
'       Call GRAEMP(Val(Text1))
'       Screen.MousePointer = 1
'      End If
'    End If
'    If MODIFICNOT% = 1 Then
'      Call GRANOTEMP 'GRABA ANOTACIONES DE EMPLEADOS
'    End If
'  End If
'  '
'  Call CIERRARCH("*.*")
'  '
'End Sub
Private Sub GuaEmp_Click()
  '
  Call Command1_Click
  '
End Sub

Private Sub Label2_dblClick()
  '
  NELGUI% = Val(Text1)
  If NELGUI% <= 0 Then Exit Sub
  '
  NR$ = RECFILT$("PADREMP", 1, MKI$(NELGUI%))
  NR1& = CVS(NR$) 'Nº DE REGISTRO DE EMPLEADO
  '
  X$ = REGLEIDO$("ACTIVA", NR1&)
  Call BLANARCH("!CONCOPCI")
  For J& = 1 To 96 Step 2
    C1% = CVI(Mid$(X$, J&, 2)) 'Nº DE CODIGO
    If C1% > 0 Then
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(C1%), 1, 2)
      Call REGAPP("!CONCOPCI", Y$)
    End If
  Next J&
  '
  For J1& = 97 To 128 'LEO LA CANTIDAD
    CAN1 = Asc(Mid$(X$, J1&, 1))
    If CAN1 > 0 Then
      NUASI% = Int(400 + J1& - 96)
      Y$ = REGBLAN$("!CONCOPCI")
      Call REPLA(Y$, MKI$(NUASI%), 1, 2) 'CODIGO
      Call REPLA(Y$, MKI$(Int(CAN1)), 7, 2) 'CANTIDAD
      Call REGAPP("!CONCOPCI", Y$)
    End If
  Next J1&
  '
15 VTB% = VENTABU%("CONCOPCI")
  If VTB% >= 0 Then
    RENUE$ = String$(128, Chr$(0))
    For J& = 1 To ULTREG&("!CONCOPCI")
      X1$ = REGLEIDO$("!CONCOPCI", J&)
      '
      NUCO% = CVI(Mid$(X1$, 1, 2))
      CANTI1 = CVI(Mid$(X1$, 7, 2))
      'ADVERTENCIA
      If NUCO% < 400 And CANTI1 > 0 Then
        Call MENSERR(24, "Solo es Posible Ingresar Cantidad a las Asignaciones Familiares.")
        GoTo 15
      End If
      '
      If NUCO% < 400 Then
       If J& = 1 Then
        Call REPLA(RENUE$, MKI$(NUCO%), 1, 2)
       Else
        Call REPLA(RENUE$, MKI$(NUCO%), Int(2 * J& - 1), 2)
       End If
      End If
      If CANTI1 > 0 And NUCO% > 400 Then
        Call REPLA(RENUE$, Chr$(CANTI1), 97 + (NUCO% - 401), 1)
      End If
      '
    Next J&
    Call GRAREG("ACTIVA", RENUE$, NR1&) 'RENGLON SEGUN EMPLEADO
'    Call CARCONCEOP(NELGUI%) 'REFRESCO LIST1
  End If
  '
End Sub

Private Sub Label37_DBLClick()
  '
  NELGUI% = Val(Text1)
  If NELGUI% <= 0 Then Exit Sub
  '
  Call TRALOFIL("FAMCAR", 1, MKI$(NELGUI%))
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left + 1200))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + SSTab1.Top + 1800))
  VTB% = VENTABU%("FAMCAR/" + ATRI$)
  If VTB% >= 0 Then
        List3.Clear
        '
        For J& = 1 To ULTREG&("!FAMCAR")
            XX$ = REGLEIDO$("!FAMCAR", J&)
            NOAP$ = Mid$(XX$, 3, 25)
            FENA = CVI(Mid$(XX$, 28, 2))
            PAREN$ = Mid$(XX$, 30, 10)
            DOCUFA$ = Mid$(XX$, 40, 11)
            YY$ = Space(68)
            Call REPLA(YY$, NOAP$, 1, 27)
            Call REPLA(YY$, FECHATEX$(FENA), 30, 10)
            Call REPLA(YY$, PAREN$, 41, 10)
            Call REPLA(YY$, DOCUFA$, 53, 11)
            List3.AddItem YY$
        Next J&
        '
        Call TRACENFIL("FAMCAR", 1, MKI$(NELGUI%))
  End If
  '
End Sub



Private Sub LiCuInEmp_Click()
  '
  Call COMUNI("Listados. No Disponible por el Momento.")
  '
End Sub

Private Sub LiGeClaEmp_Click()
  '
  Call COMUNI("Listados. No Disponible por el Momento.")
  '
End Sub

Private Sub List_Click()
  '
  Call Command13_Click
  '
End Sub

Private Sub List1_DblClick()
  Call Label2_dblClick
End Sub

Private Sub List3_DblClick()
  Call Label37_DBLClick
End Sub

Private Sub Label34_change(Index As Integer)
   '
   Screen.MousePointer = 11
   '
   If INTERVER < 1 Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
        RUTARGRA$ = RUTARGRA$ + "\"
     End If
     '
   End If
   '
   Picture1(Index).Visible = True
   Label34(Index).Visible = True
   Refresh
   '
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   'COPLANO$ = RUTARGRA$ + TRIM$(Label34(Index).Caption)
   COPLANO$ = Label34(Index)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         On Error GoTo 90
         'Picture1(Index).Visible = False
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Top = TOPEINI(Index)
         Picture1(Index).Left = IZQUINI(Index)
         Picture1(Index).Width = ANCHINI(Index)
         Picture1(Index).Height = ALTUINI(Index)
         Picture1(Index).Picture = LoadPicture("")
         Picture1(Index).Refresh
         '
         HIMAGEN = Image1(Index).Picture.Width
         VIMAGEN = Image1(Index).Picture.Height
         
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1(Index).Height = VCUADRO * REDUV / REDUC
         Picture1(Index).Width = HCUADRO * REDUH / REDUC
         Picture1(Index).Refresh
         '
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         Picture1(Index).Left = (Frame1(Index).Width - HCUADRO) / 2
         '
         Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
         Picture1(Index).Visible = True
         Picture1(Index).Refresh
         MODIFIC% = 1
       Else
         'Call MENSERR(24, "No Existe Archivo '" + COPLANO$ + "'.")
      End If
   End If
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
End Sub

Private Sub Label34_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   Label34(Index).ToolTipText = Label34(Index)
   '
End Sub

Private Sub mnuAsiSubFam_Click()
With OPSUEL05
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuCatego_Click()
With OPADMAQ04
  .Option1(3).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuGrupEmp_Click()
With OPADMAQ04
  .Option1(0).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuParInd_Click()
With OPSUEL05
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuRemOrd_Click()
With OPSUEL05
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuRetLegApo_Click()
With OPSUEL05
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub mnuSalNomBas_Click()
With OPSUEL05
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub NuevEmp_Click()
  '
  Call Command21_Click
  '
End Sub

Private Sub Print_Click()
  '
  Call Command25_Click
  '
End Sub


Private Sub Text1_Change()
  '
  Call BLANFORMU
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  NLEG% = Val(Text1)
  Call CARDATEMP(NLEG%) 'CARGA DATOS DEL EMPLEADO
  'Call CARFAMCAR(NLEG%) 'CARGA FAMILIARES A CARGO
'  Call CARCONCEOP(NLEG%) 'CARGA DE CONCEPTOS OPCIONALES
'  Call CARNOTAEMP 'CARGA ANOTACIONES DEL EMPLEADO
  MODIFIC% = 0
  MODIFICNOT% = 0
  '
End Sub

Private Sub Text1_DblClick()
  '
  Call Command9_Click
  '
End Sub

Private Sub Text10_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text11_Change()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text12_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text13_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text14_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  MODIFIC% = 1
  '
End Sub

Private Sub Text16_Change()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label23 = ECOARCH$("CATEGO", Text16)
  MODIFIC% = 1
  '
End Sub

Private Sub Text19_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  If Val(Text19) = 1 Then
    Label21 = "ACTIVO"
  Else
    Label21 = "BAJA"
  End If
  '
End Sub

Private Sub Text19_gotfocus()
  '
  Text1.SetFocus
  '
End Sub

Private Sub Text2_change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label22 = ECOARCH("ESTCIV", MKI$(Val(Text2)))
  MODIFIC% = 1
  '
End Sub

Private Sub Text25_Change()
  '
  On Error Resume Next
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  MODIFIC% = 1
  '
End Sub

Private Sub Text26_Change()
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label19 = ECOARCH$("SEXTAB", TRIM$(Text26))
  MODIFIC% = 1
  '
End Sub
Sub text26_lostfocus()
  '
  XX$ = Text26
  Text26 = UCase(XX$)
  '
End Sub
Sub BLANFORMU()
  '
  Text25 = ""
  Text6 = ""
  Text19 = ""
  Text3 = ""
  Text16 = ""
  Text26 = ""
  Text10 = ""
  Text29 = ""
  Label19 = ""
  Label20 = ""
  Label21 = ""
  Label22 = ""
  Label23 = ""
  Text2 = ""
  Text4 = ""
  Text7 = ""
  Text8 = ""
  Text13 = ""
  Text14 = ""
  Text9 = ""
  Text11 = ""
  Text12 = ""
  Text6 = ""
  Text19 = ""
  Text28 = ""
  '
End Sub
Sub CREATAB()
  'TABLA DE TIPO DE ACCESO
  If ULTREG&("TABACCE") < 1 Then
    Call BLANARCH("TABACCE")
    Call REGAPP("TABACCE", Chr$(0) + "Libre")
    Call REGAPP("TABACCE", Chr$(1) + "Reservado")
    Call CIERRARCH("TABACCE")
  End If
  'TABLA  DE FRECUENCIA
  If ULTREG&("TABFREQ") < 1 Then
    Call BLANARCH("TABFREQ")
    Call REGAPP("TABFREQ", Chr$(1) + "Mensual")
    Call REGAPP("TABFREQ", Chr$(2) + "Quincenal")
    Call CIERRARCH("TABFREQ")
  End If
  'TABLA DE MODO
  If ULTREG&("TABMODO") < 1 Then
    Call BLANARCH("TABMODO")
    Call REGAPP("TABMODO", Chr$(0) + "Directo")
    Call REGAPP("TABMODO", Chr$(1) + "Formula")
    Call CIERRARCH("TABMODO")
  End If
  'TABLA DE TABLA CARÁCTER
  If ULTREG&("TACARAC") < 1 Then
    Call BLANARCH("TACARAC")
    Call REGAPP("TACARAC", Chr$(0) + "Opcional")
    Call REGAPP("TACARAC", Chr$(1) + "Mandatorio Mensuales")
    Call REGAPP("TACARAC", Chr$(2) + "Mandatorio Quincenales")
    Call REGAPP("TACARAC", Chr$(3) + "Mandatorio General")
    Call CIERRARCH("TACARAC")
  End If
  'TABLA DE SUJETO A RETENCION
  If ULTREG&("TASURET") < 1 Then
    Call BLANARCH("TASURET")
    Call REGAPP("TASURET", Chr$(0) + "Exento")
    Call REGAPP("TASURET", Chr$(1) + "Gravado")
    Call CIERRARCH("TASURET")
  End If
  'TABLA DE QUINCENA DE LIQUIDACION
  If ULTREG&("TAQUINLI") < 1 Then
    Call BLANARCH("TAQUINLI")
    Call REGAPP("TAQUINLI", Chr$(0) + "Todas las Liquidaciones")
    Call REGAPP("TAQUINLI", Chr$(1) + "Solo 1ra Quincena")
    Call REGAPP("TAQUINLI", Chr$(2) + "Solo 2da Quincena")
    Call CIERRARCH("TAQUINLI")
  End If
  'TABLA DE ESTADOS
  If ULTREG&("TASTATUS") < 1 Then
    Call BLANARCH("TASTATUS")
    Call REGAPP("TASTATUS", "A" + "Activo")
    Call REGAPP("TASTATUS", "S" + "Suspendido")
    Call REGAPP("TASTATUS", "F" + "Fantasma")
    Call CIERRARCH("TASTATUS")
  End If
  'TABLA DE TIPOS DE REMUNERACION
  If ULTREG&("TATIREM") < 1 Then
    Call BLANARCH("TATIREM")
    Call REGAPP("TATIREM", Chr$(1) + "Normal")
    Call REGAPP("TATIREM", Chr$(2) + "Gratificacion")
    Call CIERRARCH("TATIREM")
  End If
  'TABLA DE CALCULO DE REMUNERACIONES
  If ULTREG&("TAMOCREM") < 1 Then
    Call BLANARCH("TAMOCREM")
    Call REGAPP("TAMOCREM", Chr$(0) + "Calculo Interno")
    Call REGAPP("TAMOCREM", Chr$(1) + "Informe Importe")
    Call REGAPP("TAMOCREM", Chr$(2) + "Informe Horas")
    Call REGAPP("TAMOCREM", Chr$(3) + "Cantidad")
    Call REGAPP("TAMOCREM", Chr$(4) + "Escalonamiento")
    Call CIERRARCH("TAMOCREM")
  End If
  'TABLA DE RECALCULO OFF / ON
  If ULTREG&("TARECALC") < 1 Then
    Call BLANARCH("TARECALC")
    Call REGAPP("TARECALC", Chr$(0) + "Recalculo off")
    Call REGAPP("TARECALC", Chr$(1) + "Recalculo on")
    Call CIERRARCH("TARECALC")
  End If
  'TABLA DE MODO DE CALCULO DE RETENCIONES
  If ULTREG&("TAMOCRET") < 1 Then
    Call BLANARCH("TAMOCRET")
    Call REGAPP("TAMOCRET", Chr$(0) + "Por Formula")
    Call REGAPP("TAMOCRET", Chr$(1) + "Informe Importe")
    Call REGAPP("TAMOCRET", Chr$(2) + "Porcentajes")
    Call REGAPP("TAMOCRET", Chr$(3) + "Sumas Fijas")
    Call CIERRARCH("TAMOCRET")
  End If
  'TABLA DE TIPO DE ACTUALIZACION (SALARIOS NOMINALES)
  If ULTREG&("TIACTSAL") < 1 Then
    Call BLANARCH("TIACTSAL")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(1), 1, 2)
    Call REPLA(Y$, "Actualizacion", 3, 30)
    Call REGAPP("TIACTSAL", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(2), 1, 2)
    Call REPLA(Y$, "Listado", 3, 30)
    Call REGAPP("TIACTSAL", Y$)
    Call CIERRARCH("TIACTSAL")
  End If
  'TABLA DE SEXO
  If ULTREG&("SEXTAB") < 1 Then
    Call BLANARCH("SEXTAB")
    Y$ = Space(32)
    Call REPLA(Y$, "F", 1, 1)
    Call REPLA(Y$, "FEMENINO", 2, 31)
    Call REGAPP("SEXTAB", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, "M", 1, 1)
    Call REPLA(Y$, "MASCULINO", 2, 31)
    Call REGAPP("SEXTAB", Y$)
    Call CIERRARCH("SEXTAB")
  End If
  'TABLA DE ESTADO CIVIL
  If ULTREG&("ESTCIV") < 1 Then
    Call BLANARCH("ESTCIV")
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(1), 1, 2)
    Call REPLA(Y$, "SOLTERO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(2), 1, 2)
    Call REPLA(Y$, "CASADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(3), 1, 2)
    Call REPLA(Y$, "SEPARADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(4), 1, 2)
    Call REPLA(Y$, "DIVORCIADO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Y$ = Space(32)
    Call REPLA(Y$, MKI$(5), 1, 2)
    Call REPLA(Y$, "VIUDO/A", 3, 30)
    Call REGAPP("ESTCIV", Y$)
    Call CIERRARCH("ESTCIV")
  End If
  '\\\LIQUIDACIONES///
  'TABLA DE TIPOS DE LIQUIDACION
  If ULTREG&("TATIPLIQ") < 1 Then
    Call BLANARCH("TATIPLIQ")
    Call REGAPP("TATIPLIQ", Chr$(1) + "Primera Quincena")
    Call REGAPP("TATIPLIQ", Chr$(2) + "Segunda Quincena")
    Call REGAPP("TATIPLIQ", Chr$(3) + "Retroactivo")
    Call CIERRARCH("TATIPLIQ")
  End If
  'TABLA DE ASIGNACIONES FLIARES.
  If ULTREG&("TASIGFAM") < 1 Then
    Call BLANARCH("TASIGFAM")
    Call REGAPP("TASIGFAM", Chr$(0) + "No se Liquida")
    Call REGAPP("TASIGFAM", Chr$(1) + "Si se Liquida")
    Call CIERRARCH("TASIGFAM")
  End If
  'TABLA DE MESES
  If ULTREG&("TAMESE") < 1 Then
    Call BLANARCH("TAMESE")
    Call REGAPP("TAMESE", Chr$(1) + "Enero")
    Call REGAPP("TAMESE", Chr$(2) + "Febrero")
    Call REGAPP("TAMESE", Chr$(3) + "Marzo")
    Call REGAPP("TAMESE", Chr$(4) + "Abril")
    Call REGAPP("TAMESE", Chr$(5) + "Mayo")
    Call REGAPP("TAMESE", Chr$(6) + "Junio")
    Call REGAPP("TAMESE", Chr$(7) + "Julio")
    Call REGAPP("TAMESE", Chr$(8) + "Agosto")
    Call REGAPP("TAMESE", Chr$(9) + "Septiembre")
    Call REGAPP("TAMESE", Chr$(10) + "Octubre")
    Call REGAPP("TAMESE", Chr$(11) + "Noviembre")
    Call REGAPP("TAMESE", Chr$(12) + "Diciembre")
    Call CIERRARCH("TAMESE")
  End If
End Sub

Private Sub Text28_Change()
  '
  MODIFICNOT% = 1
  '
End Sub

Private Sub Text29_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text3_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text4_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text6_Change()
  '
  If Val(Text1) > 32767 Then
    Call COMUNI("Número de Legajo no Válido.")
    Text1.SelStart = 0
    Text1.SelLength = (Len(Text1))
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  '
  Label20 = ECOARCH("DATGREM1", MKI$(Val(Text6)))
  MODIFIC% = 1
  '
End Sub

Private Sub text7_change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text8_Change()
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text9_Change()
  '
  MODIFIC% = 1
  '
End Sub
Sub HACECONCEP() 'CARGA ECO DE CONCEPTOS
    '
    Call BLANARCH("ECONCEP")
    '
    For J& = 1 To ULTREG("CONCEP")
      X$ = REGLEIDO$("CONCEP", J&)
      NUCOD% = Int(J&)
      DESCRI1$ = TRIM$(Mid$(X$, 1, 30))
      EST% = Asc(Mid$(X$, 31, 1))
      If EST% <> 1 Then GoTo 10 'SI ACTIVO
      If TRIM(DESCRI1$) = "" Then GoTo 10 'SI DESCRIPCION
      Y$ = Space$(32)
      Call REPLA(Y$, MKI$(NUCOD%), 1, 2)
      Call REPLA(Y$, DESCRI1$, 3, 30)
      Call REGAPP("ECONCEP", Y$)
10  Next J&
    '
    For J& = 1 To ULTREG("RETLEG")
      X1$ = REGLEIDO$("RETLEG", J&)
      NUCOD% = Int(J& + 200)
      DESCRI1$ = TRIM$(Mid$(X1$, 1, 30))
      EST% = Asc(Mid$(X1$, 31, 1))
      If EST% <> 1 Then GoTo 20
      If TRIM(DESCRI1$) = "" Then GoTo 20
      Y1$ = Space$(32)
      Call REPLA(Y1$, MKI$(NUCOD%), 1, 2)
      Call REPLA(Y1$, DESCRI1$, 3, 30)
      Call REGAPP("ECONCEP", Y1$)
20  Next J&
    '
    For J& = 1 To ULTREG("ASIGNA")
      X2$ = REGLEIDO$("ASIGNA", J&)
      NUCOD% = Int(J& + 400)
      DESCRI1$ = TRIM$(Mid$(X2$, 1, 30))
      EST% = Asc(Mid$(X2$, 43, 1))
      If EST% <> 1 Then GoTo 30
      If TRIM(DESCRI1$) = "" Then GoTo 30
      Y2$ = Space$(32)
      Call REPLA(Y2$, MKI$(NUCOD%), 1, 2)
      Call REPLA(Y2$, DESCRI1$, 3, 30)
      Call REGAPP("ECONCEP", Y2$)
30  Next J&
    '
    Call CIERRARCH("ECONCEP")
    '
End Sub
'Sub CARCONCEOP(NELG%)
'  '
'  If NELG% <= 0 Then Exit Sub
'  '
'  NR$ = RECFILT$("PADREMP", 1, MKI$(NELG%))
'  NR1& = CVS(NR$) 'Nº DE REGISTRO DE EMPLEADO
'  '
'  X$ = REGLEIDO$("ACTIVA", NR1&)
'  Call BLANARCH("!CONCOPCI")
'  For J& = 1 To 96 Step 2
'    C1% = CVI(Mid$(X$, J&, 2)) 'Nº DE CODIGO
'    If C1% > 0 Then
'      Y$ = REGBLAN$("!CONCOPCI")
'      Call REPLA(Y$, MKI$(C1%), 1, 2)
'      Call REGAPP("!CONCOPCI", Y$)
'      '
'    End If
'  Next J&
'  '
'  For j1& = 97 To 128 'LEO LA CANTIDAD
'    CAN1 = Asc(Mid$(X$, j1&, 1))
'    If CAN1 > 0 Then
'      NUASI% = Int(400 + j1& - 96)
'      Y$ = REGBLAN$("!CONCOPCI")
'      Call REPLA(Y$, MKI$(NUASI%), 1, 2) 'CODIGO
'      Call REPLA(Y$, MKI$(Int(CAN1)), 7, 2) 'CANTIDAD
'      Call REGAPP("!CONCOPCI", Y$)
'    End If
'  Next j1&
'  'AGREGO A LA LISTA
'  List1.Clear
'  For J& = 1 To ULTREG&("!CONCOPCI")
'    XX$ = REGLEIDO$("!CONCOPCI", J&)
'    Z$ = Space$(64)
'    cod2% = CVI(Mid$(XX$, 1, 2))
'    DESCRI$ = ECOARCH("ECONCEP", MKI$(cod2%))
'    UM$ = Mid$(XX$, 3, 4)
'    CANT2% = CVI(Mid$(XX$, 7, 2))
'    FORMU$ = Mid$(XX$, 9, 56)
'    '
'    Call REPLA(Z$, FORMATNUM(cod2%, "F5"), 1, 5)
'    Call REPLA(Z$, DESCRI$, 7, 30)
'    Call REPLA(Z$, UM$, 39, 4)
'    Call REPLA(Z$, FORMATNUM(CANT2%, "F4"), 45, 4)
'    Call REPLA(Z$, FORMU$, 52, 56)
'    '
'    List1.AddItem Z$
'  Next J&
'  '
'End Sub
Sub APACTPADRO() 'APAREA PADREMP = ACTIVA
  '
  PANU& = ULTREG("PADREMP")
  CONU& = ULTREG("ACTIVA")
  DIF& = PANU& - CONU&
  If DIF& > 0 Then
    For J& = 1 To DIF&
      YYY$ = String$(128, Chr$(0))
      'PARA Q' NO SEA VÁLIDO EL NUMERO DE CONCEPTO OPCIONAL
      Call REPLA(YYY$, MKI$(-1), 1, 2)
      Call REGAPP("ACTIVA", YYY$)
    Next J&
  End If
  '
  Call CIERRARCH("PADREMP")
  Call CIERRARCH("ACTIVA")
  '
End Sub
Sub APAREMBA() 'APAREA PADREMP = REMBAS
  '
  PANU& = ULTREG("PADREMP")
  CONU& = ULTREG("REMBAS")
  DIF& = PANU& - CONU&
  If DIF& > 0 Then
    For J& = 1 To DIF&
      YYY$ = String$(128, Chr$(0))
      Call REPLA(YYY$, Space(48), 1, 48)
      Call REGAPP("REMBAS", YYY$)
    Next J&
  End If
  '
  Call CIERRARCH("PADREMP")
  Call CIERRARCH("REMBAS")
  '
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
Private Sub Command5_Click()
   '
   If DERACCE%("MODIOPE", "Modificacion Datos del Operario") > 1 Then
      Call Command1_Click
   End If
   '
End Sub
'
Sub GRAFOTO() 'GRABA FOTO DEL EMPLEADO
   '
   FOTO1$ = Label34(2)
   If TRIM(FOTO1$) = "" Then Exit Sub
   NLEG% = Val(TRIM$(Text1.TEXT))
   If NLEG% < 1 Then Exit Sub
   If TRIM(RECFILT$("PADREMP", 1, MKI$(NLEG%))) = "" Then Exit Sub
   '
   If TRIM(RECFILT$("FOTOEMP", 1, MKI$(NLEG%))) = "" Then
      Y$ = REGBLAN$("FOTOEMP")
      Call REPLA(Y$, MKI$(NLEG%), 1, 2)
      Call REPLA(Y$, FOTO1$, 3, 510)
      Call REGAPP("FOTOEMP", Y$)
   Else
      FIN1& = ULTREG&("FOTOEMP")
      YYY$ = RECFILT$("FOTOEMP", 1, MKI$(NLEG%))
      J2& = CVS(YYY$)
      XZ$ = REGLEIDO$("FOTOEMP", J2&)
      Call REPLA(XZ$, FOTO1$, 3, 510)
      Call GRAREG("FOTOEMP", XZ$, J2&)
      Call SETULTREG("FOTOEMP", FIN1&)
      Call CIERRARCH("FOTOEMP")
   End If
   '
End Sub
Sub GRANOTEMP() 'GRABA ANOTACIONES DE EMPLEADOS
   '
   NLEG% = Val(TRIM$(Text1.TEXT))
   If NLEG% < 1 Then Exit Sub
   If TRIM(RECFILT$("PADREMP", 1, MKI$(NLEG%))) = "" Then Exit Sub
   '
   If NLEG% > 0 Then
         If MODIFICNOT% < 1 Then
           Exit Sub
         End If
         '
         If DERACCE%("ANOTACLI", "Anotaciones de Empleado") < 2 Then
            Exit Sub
         End If
         '
         If MODIFICNOT% = 1 Then
          If COMALTER%("Se han Modificado las Notas del Empleado\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
'            Call CARNOTAEMP   'CARGA NOTAS DE EMPLEADO
             Exit Sub
          End If
         End If
         '
         Screen.MousePointer = 11
         URN& = ULTREG&("NOTAEMP")
         RNC$ = RECFILT$("NOTAEMP", 1, MKI$(NLEG%))
         KK& = (-3)
         '
         TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 78)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(NLEG%) + RETEX$(LL%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("NOTAEMP", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTAEMP", MKI$(0) + Space$(78), RENOTA&)
         Wend
         '
         Call CIERRARCH("NOTAEMP")
         MODIFICNOT% = 0
         Screen.MousePointer = 1
   End If
End Sub
'Sub CARNOTAEMP() 'CARGA NOTAS DEL EMPLEADO
'
'   NLEG% = Val(Text1.TEXT)
'   Text28.TEXT = ""
'
'   If NLEG% > 0 Then
'      If TRIM$(RECFILT$("PADREMP", 1, MKI$(NLEG%))) <> "" Then
'         Screen.MousePointer = 11
'         RNC$ = RECFILT$("NOTAEMP", 1, MKI$(NLEG%))
'         URN& = ULTREG&("NOTAEMP")
'
'         For U& = 1 To Len(RNC$) Step 4
'            RENOTA& = CVS(Mid$(RNC$, U&, 4))
'            If RENOTA& > 0 Then
'               If RENOTA& <= URN& Then
'                  TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTAEMP", RENOTA&), 3))
'                  If Text28.TEXT <> "" Then
'                     Text28.TEXT = Text28.TEXT + Chr$(13) + Chr$(10)
'                  End If
'                  Text28.TEXT = Text28.TEXT + RTrim$(TTYY$)
'               End If
'            End If
'         Next U&
'         Screen.MousePointer = 1
'      End If
'   End If
'   Text28.Refresh
'
'End Sub
Sub APAPADREMP()
  'APAREA SI NO EXISTE PADREMP O NO TIENE REGISTROS
  If ULTREG&("PADREMP") <= 0 Then
    Call BLANARCH("PADREMP")
    For J& = 1 To ULTREG("PADRON")
      PA1$ = REGLEIDO$("PADRON", J&)
      PA2$ = REGLEIDO$("PADAUX", J&)
      PA3$ = REGLEIDO$("PADAUX2", J&)
      NULE% = CVI(Mid$(PA1$, 1, 2))
      If NULE% > 0 Then 'PARA QUE NO GRABE REGISTROS EN BLANCO
        PADREM$ = Space(256)
        Call REPLA(PADREM$, PA1$, 1, 64)
        'CATEGO1$ = Mid$(PA1$, 54, 9)
        Call REPLA(PADREM$, PA2$, 65, 96)
        Call REPLA(PADREM$, PA3$, 161, 48)
        Call REGAPP("PADREMP", PADREM$)
      End If
    Next J&
  End If
  '
End Sub

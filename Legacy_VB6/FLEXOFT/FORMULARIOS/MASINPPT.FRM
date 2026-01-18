VERSION 5.00
Begin VB.Form MASINPPT 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   75
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   10920
      Picture         =   "MASINPPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   105
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   10920
      Picture         =   "MASINPPT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7750
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9120
      Left            =   735
      TabIndex        =   27
      Top             =   0
      Width           =   9900
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   735
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   5145
         Width           =   8940
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control de Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   0
         Left            =   765
         TabIndex        =   0
         Top             =   2835
         Width           =   2955
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordenes de Fabricación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   765
         TabIndex        =   1
         Top             =   3150
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestión de Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   765
         TabIndex        =   2
         Top             =   3465
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seguimiento de Manufactura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   3
         Left            =   765
         TabIndex        =   3
         Top             =   3780
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Programación de Materiales MRP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   4
         Left            =   5250
         TabIndex        =   4
         Top             =   2835
         Width           =   3900
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Carga de Máquinas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   5265
         TabIndex        =   5
         Top             =   3150
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Análisis de Costos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   11
         Left            =   5265
         TabIndex        =   6
         Top             =   3465
         Width           =   3165
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestión de la Calidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   6
         Left            =   5265
         TabIndex        =   7
         Top             =   3780
         Width           =   2850
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestión de Ventas y Deudores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   7
         Left            =   765
         TabIndex        =   8
         Top             =   4305
         Width           =   3585
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acreedores Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   8
         Left            =   765
         TabIndex        =   9
         Top             =   4620
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contabilidad General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   9
         Left            =   5265
         TabIndex        =   10
         Top             =   4305
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Administración de Personal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   10
         Left            =   5250
         TabIndex        =   11
         Top             =   4620
         Width           =   3270
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remitir información a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2640
         Left            =   0
         TabIndex        =   28
         Top             =   5880
         Width           =   9675
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   7
            Left            =   6195
            TabIndex        =   17
            Text            =   " "
            Top             =   315
            Width           =   3270
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   6
            Left            =   6195
            TabIndex        =   18
            Text            =   " "
            Top             =   630
            Width           =   3270
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   5
            Left            =   6195
            TabIndex        =   19
            Text            =   " "
            Top             =   945
            Width           =   3270
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   4
            Left            =   6195
            TabIndex        =   20
            Text            =   " "
            Top             =   1260
            Width           =   3270
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1275
            Index           =   1
            Left            =   4830
            TabIndex        =   41
            Top             =   360
            Width           =   1300
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Att Sr/Sra:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Teléfono:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   0
               TabIndex        =   44
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   0
               TabIndex        =   43
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "e-mail:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   0
               TabIndex        =   42
               Top             =   945
               Width           =   1275
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1275
            Index           =   0
            Left            =   105
            TabIndex        =   36
            Top             =   360
            Width           =   1300
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Pcia - País:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   0
               TabIndex        =   40
               Top             =   945
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "CP -Localidad:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   0
               TabIndex        =   39
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Domicilio:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   0
               TabIndex        =   38
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Empresa:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   3
            Left            =   1470
            TabIndex        =   16
            Text            =   " "
            Top             =   1260
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   1470
            TabIndex        =   15
            Text            =   " "
            Top             =   945
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   1470
            TabIndex        =   14
            Text            =   " "
            Top             =   630
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   1470
            TabIndex        =   13
            Text            =   " "
            Top             =   315
            Width           =   3165
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Idioma Preferido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   210
            TabIndex        =   29
            Top             =   1785
            Width           =   9255
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Francais"
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
               Index           =   4
               Left            =   7560
               TabIndex        =   25
               Top             =   210
               Width           =   1170
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Deutsch"
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
               Index           =   3
               Left            =   5880
               TabIndex        =   24
               Top             =   210
               Width           =   1065
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "English"
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
               Index           =   2
               Left            =   4095
               TabIndex        =   23
               Top             =   210
               Width           =   1380
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Portugués"
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
               Index           =   1
               Left            =   2100
               TabIndex        =   22
               Top             =   210
               Width           =   1485
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Español"
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
               Index           =   0
               Left            =   315
               TabIndex        =   21
               Top             =   210
               Value           =   -1  'True
               Width           =   1380
            End
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "           4524-0796 / 4522-2048 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   52
         Top             =   1890
         Width           =   3900
      End
      Begin VB.Line Line2 
         X1              =   6195
         X2              =   9450
         Y1              =   9080
         Y2              =   9080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Otros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   0
         TabIndex        =   50
         Top             =   5145
         Width           =   645
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   195
         Left            =   1995
         TabIndex        =   49
         Top             =   315
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9870
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia - País"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7350
         TabIndex        =   48
         Top             =   840
         Width           =   2400
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "CP - Localidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   7350
         TabIndex        =   47
         Top             =   630
         Width           =   2400
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7245
         TabIndex        =   46
         Top             =   420
         Width           =   2505
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Su Empresa"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   7350
         TabIndex        =   35
         Top             =   0
         Width           =   2340
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deseo recibir más información, sin cargo ni compromiso , acerca de los sistemas de:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   34
         Top             =   2310
         Width           =   8940
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sres."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   33
         Top             =   1260
         Width           =   3060
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "FLEXOFT Soluciones Informáticas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   32
         Top             =   1470
         Width           =   3690
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "TeleFax: (00-54-11)  4524-1284"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   31
         Top             =   1680
         Width           =   3900
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Buenos Aires, 25 de marzo de 2002"
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
         Left            =   3255
         TabIndex        =   30
         Top             =   1260
         Width           =   6420
      End
   End
   Begin VB.Line Line3 
      X1              =   10635
      X2              =   10635
      Y1              =   0
      Y2              =   9450
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   10630
      Stretch         =   -1  'True
      Top             =   5150
      Width           =   3800
   End
End
Attribute VB_Name = "MASINPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      For KKI% = 0 To 7
        If Trim$(Text1(KKI%).Text) = "" Then
          MsgBox "   Por favor, ingrese sus datos" + Chr$(13) + Chr$(10) + "para el envío de la información."
          On Error Resume Next
          Text1(KKI%).SetFocus
          On Error GoTo 0
          Exit Sub
        End If
      Next KKI%
      '
10    AFOR = MASINPPT.Width
      IFRA = Frame1.Left
      APRI = Printer.ScaleWidth - 400
      BACOL = MASINPPT.BackColor
      APRIN = MASINPPT.Width
      If MASINPPT.Width > APRI Then
        APRIN = APRI
      End If
      Frame1.Left = APRIN - Frame1.Width
      Command1(0).Visible = False
      Command1(1).Visible = False
      Image1.Visible = False
      Line3.Visible = False
      MASINPPT.BackColor = &HFFFFFF
      Frame1.BackColor = &HFFFFFF
      Frame2.BackColor = &HFFFFFF
      Frame3.BackColor = &HFFFFFF
      Frame4(0).BackColor = &HFFFFFF
      Frame4(1).BackColor = &HFFFFFF
      For KKI% = 0 To 11: Check1(KKI%).BackColor = &HFFFFFF: Next KKI%
      For KKI% = 0 To 4: Option1(KKI%).BackColor = &HFFFFFF: Next KKI%
      Refresh
      '
      PrintForm
      '
      MASINPPT.Width = AFOR
      Frame1.Left = IFRA
      Command1(0).Visible = True
      Command1(1).Visible = True
      Image1.Visible = True
      Line3.Visible = True
      MASINPPT.BackColor = BACOL
      Frame1.BackColor = BACOL
      Frame2.BackColor = BACOL
      Frame3.BackColor = BACOL
      Frame4(0).BackColor = BACOL
      Frame4(1).BackColor = BACOL
      For KKI% = 0 To 11: Check1(KKI%).BackColor = BACOL: Next KKI%
      For KKI% = 0 To 4: Option1(KKI%).BackColor = BACOL: Next KKI%
      '
      Refresh
      Text2.Refresh
      For KKI% = 0 To 7: Text1(KKI%).Refresh: Next KKI%
      Refresh
      OPENFORMS = DoEvents
      '
      MsgBox "           Favor enviarnos por fax" + Chr$(10) + Chr$(13) + "su pedido de información impreso."
   End If
   '
   Screen.MousePointer = 11
   Close
   Hide
   OPENFORMS = DoEvents
   If MINIDEM% = 1 Then
      On Error Resume Next
      Unload VECPANTA
      Unload SECUENDE
      Unload PROPINTEM
      Unload INFODEM
      Unload FORDEMO
      Unload FINSECUEN
      Unload ENCAPIPPT
      Unload CONFIPAN
      Unload BLACKSCREEN
      Screen.MousePointer = 1
      Unload Me
      End
    Else
      DEMOSPPT.Command1.Visible = True
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Label13 = FETEXTO$
   If MINIDEM% = 1 Then
       Image1.Picture = PROPINTEM.Picture1.Picture
     Else
       Image1.Picture = PROPIPPT.Picture1.Picture
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   If Index <= 3 Then
      AMAXI = 0
      For KKI% = 0 To 3
         TEPRUEBA.Font = Label14(KKI%).Font
         TEPRUEBA.FontSize = Label14(KKI%).FontSize
         TEPRUEBA.FontBold = Label14(KKI%).FontBold
         TEPRUEBA.Caption = Trim$(Text1(KKI%).Text)
         ATEXTO = TEPRUEBA.Width
         If ATEXTO > AMAXI Then AMAXI = ATEXTO
      Next KKI%
      '
      For KKI% = 0 To 3
         Label14(KKI%).Left = Frame1.Width - AMAXI - 300
         Label14(KKI%).Width = AMAXI
         Label14(KKI%).Caption = Trim$(Text1(KKI%).Text)
      Next KKI%
      OPENFORMS = DoEvents
   End If
End Sub

Function FETEXTO$()
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    '
    FE$ = Format$(Now, "dd/mm/yy")
    MX% = Val(Mid$(FE$, 4, 2))
    If MX% < 1 Or MX% > 12 Then
       FETEXTO$ = FE$
       Exit Function
    End If
    MXS$ = MES$(MX%)
    '
    FT$ = Trim$(Str$(Val(Left$(FE$, 2)))) + " de " + MXS$ + " de "
    If Val(Mid$(FE$, 7, 2)) >= 80 Then
           FT$ = FT$ + "19"
        Else
           FT$ = FT$ + "20"
    End If
    '
    FT$ = FT$ + Mid$(FE$, 7, 2) + ".-"
    '
    FETEXTO$ = FT$
    '
End Function
'


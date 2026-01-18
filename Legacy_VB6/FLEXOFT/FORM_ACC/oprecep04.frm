VERSION 5.00
Begin VB.Form OPRECEP04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formularios de Recepción y Rechazos"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture15 
      Height          =   4900
      Left            =   10395
      ScaleHeight     =   4845
      ScaleWidth      =   1035
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   5000
         Left            =   0
         Picture         =   "oprecep04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5250
      TabIndex        =   53
      Top             =   210
      Width           =   4950
      Begin VB.CheckBox Check10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Número de Control de Vigilancia"
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
         Left            =   735
         TabIndex        =   55
         Top             =   300
         Width           =   3375
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Certificado de Calidad Proveedor"
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
         Left            =   735
         TabIndex        =   54
         Top             =   555
         Width           =   3585
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Comprobante Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5250
      TabIndex        =   40
      Top             =   3500
      Width           =   4950
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sticker"
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
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox Text10 
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
         Left            =   2415
         TabIndex        =   44
         Text            =   " "
         Top             =   420
         Width           =   1380
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
         Height          =   285
         Left            =   3780
         TabIndex        =   43
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
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
         Left            =   210
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   " "
         Top             =   760
         Width           =   3750
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   4095
         Picture         =   "oprecep04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form:"
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
         Left            =   0
         TabIndex        =   45
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Remito por Devolución (Rechazos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5250
      TabIndex        =   17
      Top             =   1365
      Width           =   4950
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3885
         TabIndex        =   25
         Top             =   1155
         Width           =   855
         Begin VB.TextBox Text6 
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
            Left            =   190
            TabIndex        =   26
            Text            =   "RD"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   22
         Top             =   1155
         Width           =   3540
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Automatica"
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
            Left            =   1995
            TabIndex        =   24
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Manual"
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
            Left            =   420
            TabIndex        =   23
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4095
         Picture         =   "oprecep04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   420
         Width           =   645
      End
      Begin VB.TextBox Text4 
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
         Left            =   210
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
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
         Left            =   3780
         TabIndex        =   19
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2415
         TabIndex        =   18
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Formato Listador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2660
      TabIndex        =   14
      Top             =   3750
      Width           =   2490
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ficha p/Item"
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
         Height          =   240
         Left            =   315
         TabIndex        =   16
         Top             =   630
         Width           =   1800
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Form. Tabulado"
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
         Left            =   315
         TabIndex        =   15
         Top             =   340
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   11
      Top             =   3750
      Width           =   2355
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bloques Espumas"
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
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   13
         Top             =   630
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mat. General"
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
         Index           =   0
         Left            =   315
         TabIndex        =   12
         Top             =   340
         Value           =   1  'Checked
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Boleta de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4950
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Otras Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   210
         TabIndex        =   31
         Top             =   1995
         Width           =   4530
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Suprime Validación de Remito"
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
            Left            =   210
            TabIndex        =   52
            Top             =   450
            Width           =   3585
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
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
            Left            =   3400
            TabIndex        =   48
            Text            =   "0.0"
            Top             =   870
            Width           =   660
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Bloquea Sobre-Embarque más de              %"
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
            Left            =   210
            TabIndex        =   47
            Top             =   980
            Width           =   4250
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Anula Cantidades Presentadas"
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
            Left            =   210
            TabIndex        =   33
            Top             =   720
            Width           =   3060
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "No lmpresión Comprobante"
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
            Left            =   210
            TabIndex        =   32
            Top             =   200
            Width           =   2850
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Anticipación"
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
         Left            =   1995
         TabIndex        =   28
         Top             =   1050
         Width           =   1800
         Begin VB.TextBox Text7 
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
            Left            =   1050
            TabIndex        =   29
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dias F.Sol.Ent:"
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
            Left            =   -60
            TabIndex        =   30
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   9
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4095
         Picture         =   "oprecep04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   645
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Numeración"
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
         Left            =   210
         TabIndex        =   3
         Top             =   1050
         Width           =   1695
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Manual"
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
            TabIndex        =   5
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Automatica"
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
            TabIndex        =   4
            Top             =   525
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Código"
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
         Left            =   3885
         TabIndex        =   1
         Top             =   1050
         Width           =   855
         Begin VB.TextBox Text5 
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
            Left            =   190
            TabIndex        =   2
            Text            =   "BR"
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   6
         Top             =   375
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7140
      Left            =   10395
      ScaleHeight     =   7080
      ScaleWidth      =   1425
      TabIndex        =   34
      Top             =   -150
      Width           =   1485
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         Picture         =   "oprecep04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Actualizar Tabla de Motivos de Rechazo."
         Top             =   1810
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   38
         Top             =   4000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   640
         Left            =   105
         Picture         =   "oprecep04.frx":2E00
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Salir"
         Top             =   420
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "oprecep04.frx":2F4A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Guardar"
         Top             =   3150
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "oprecep04.frx":3254
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Help"
         Top             =   1060
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         Picture         =   "oprecep04.frx":355E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Actualizar Tabla de Destinos de Rechazo."
         Top             =   2200
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         Picture         =   "oprecep04.frx":36A8
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Autorización de Recepciones sin O/C"
         Top             =   2600
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "oprecep04.frx":37F2
         Stretch         =   -1  'True
         Top             =   4410
         Width           =   1410
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tablas 
      Caption         =   "Tablas"
      Begin VB.Menu actualizatablamotivorechazo 
         Caption         =   "Actualizar Tabla de Motivos de Rechazo."
      End
      Begin VB.Menu Actualizatabladestinorechazo 
         Caption         =   "Actualizar Tabla de Destinos de Rechazo."
      End
      Begin VB.Menu autorizacionderecepcionsinordendecompra 
         Caption         =   "Autorización de Recepciones sin O/C."
      End
   End
End
Attribute VB_Name = "OPRECEP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizatabladestinorechazo_Click()
    Call Command10_Click
End Sub

Private Sub actualizatablamotivorechazo_Click()
   Call Command11_Click
End Sub

Private Sub autorizacionderecepcionsinordendecompra_Click()
    Call Command14_Click
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Text10 = "STICKER"
        Text9 = ""
      ElseIf Check4.Value = 0 Then
        If Text10 = "STICKER" Then
          Text10 = ""
          Text9 = ""
        End If
    End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'\\\OT-05-0723-WAR-22/11/05///
Private Sub Command10_Click()
  '
  'TABLA DE DESTINO DE RECHAZO
  If ULTREG&("TADESTI") < 1 Then
    Call BLANARCH("TADESTI")
    Call REGAPP("TADESTI", Chr$(0) + "Retrabajo")
    Call REGAPP("TADESTI", Chr$(1) + "Devolución")
    Call REGAPP("TADESTI", Chr$(2) + "Baja")
    Call CIERRARCH("TADESTI")
  End If
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TADESTI", "TADESTI")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TADESTI/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TADESTI")
         X$ = REGLEIDO$("!TADESTI", U&)
         JJ& = JJ& + 1
         Call GRAREG("TADESTI", X$, JJ&)
      Next U&
      Call SETULTREG("TADESTI", JJ&)
      Call CIERRARCH("TADESTI")
    End If
  End If
  '
End Sub
'\\\OT-05-0723-WAR-FIN///
'\\\OT-05-0723-WAR-23/11/05///
Private Sub Command11_Click()
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TAMOTIV", "TAMOTIV")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TAMOTIV/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TAMOTIV")
         X$ = REGLEIDO$("!TAMOTIV", U&)
         JJ& = JJ& + 1
         Call GRAREG("TAMOTIV", X$, JJ&)
      Next U&
      Call SETULTREG("TAMOTIV", JJ&)
      Call CIERRARCH("TAMOTIV")
    End If
  End If
  '
End Sub
'\\\OT-05-0723-WAR-FIN///
Private Sub Command14_Click()
  '
  '\\\OT-05-0694-WAR-25/10/05///
  If DERACCE%("STAMACO", "Autorización Recepciones sin Orden de Compra") < 2 Then
    Exit Sub
  End If
  Command14.Enabled = False
  '
  Call TRALOCAL("TAMATCON", "TAMATCON")
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 2600))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("TAMATCON/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!TAMATCON")
        X$ = REGLEIDO$("!TAMATCON", U&)
        'LE PONGO EL USUARIO QUE AUTORIZO
        US$ = Mid$(X$, 5, 24)
        If TRIM(US$) = "" Then
          Call REPLA(X$, USERNAME$, 5, 24)
        End If
        F1% = CVI(Mid$(X$, 3, 2))
        If F1% = 0 Then
          HOI% = HOY(HO$)
          Call REPLA(X$, MKI$(HOI%), 3, 2)
        End If
        JJ& = JJ& + 1
        Call GRAREG("TAMATCON", X$, JJ&)
     Next U&
     Call SETULTREG("TAMATCON", JJ&)
     Call CIERRARCH("TAMATCON")
  End If
  '
  Call FILESORT("TAMATCON", "TAMATCON", 1, 1, "ASC")
  '
  RTA% = COMALTER%("Seleccione la Siguiente Tarea:\\Emitir Listado\Terminar")
  If RTA% = 1 Then
    Call FINAL("*LIMATCON")
  End If
  Command14.Enabled = True
  '\\\OT-05-0694-WAR-FIN///
  '
End Sub

Private Sub Command2_Click()
    '
    FORIPRE$ = TRIM$(UCase$(Text1.TEXT))
    Call GRACONTROL("", "FORMBREC", Text1.TEXT)
    Call GRACONTROL("", "CODMBREC", Text5.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEBREC", MONUME$)
    '
    Call GRACONTROL("", "REMIRECH", Text3.TEXT)
    Call GRACONTROL("", "CODRRECH", Text6.TEXT)
    '
    Call GRACONTROL("", "STIKAPRO", Text10.TEXT)
    '
    MONUME$ = "AUTO": If Option5.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMERECH", MONUME$)
    '
    MONUME$ = "NO": If Check1(1).Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("", "RECESPUM", MONUME$)
    '
    MONUME$ = "FORTABU": If Option4.Value = True Then MONUME$ = "FICHA"
    Call GRACONTROL("", "COMPBREC", MONUME$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("BOLRECEP", "SUPRINT", SUPRINT$)
    '
    ANUCAN$ = "NO"
    If Check3.Value = 1 Then ANUCAN$ = "SI"
    Call GRACONTROL("BOLRECEP", "ANUCANTI", ANUCAN$)
    '
    '\\\OT-06-0515-RG-28/11/06///
    MONUME$ = "NO"
    If Check6.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("BOLRECEP", "SUPREMI", MONUME$)
    '\\\OT-06-0515-RG-FIN///
    '
    BLOSOBRE$ = "NO"
    If Check5.Value = 1 Then BLOSOBRE$ = "SI"
    Call GRACONTROL("BOLRECEP", "BLOSOBRE", BLOSOBRE$)
    Call GRACONTROL("BOLRECEP", "PORSOBRE", Text8)
    '
    DIANREC$ = TRIM$(Text7)
    Call GRACONTROL("BOLRECEP", "DIANREC", DIANREC$)
    '
    NUCONVIG$ = "NO": If Check10.Value = 1 Then NUCONVIG$ = "SI"
    Call GRACONTROL("BOLRECEP", "NUCONVIG", NUCONVIG$)
    '
    CERCAPRO$ = "NO": If Check7.Value = 1 Then CERCAPRO$ = "SI"
    Call GRACONTROL("BOLRECEP", "CERCAPRO", CERCAPRO$)
    '
    X$ = REGBLAN$("TATIREC")
    Call REPLA(X$, Chr$(1), 1, 1)
    Call REPLA(X$, "Material General", 2, 31)
    Call REPLA(X$, "BOLREGEN", 49, 16)
    Call REPLA(X$, FORIPRE$, 65, 16)
    Call GRAREG("TATIREC", X$, 1)
    '
    If Check1(1).Value = 1 Then
          X$ = REGBLAN$("TATIREC")
          Call REPLA(X$, Chr$(2), 1, 1)
          Call REPLA(X$, "Bloques de Espuma", 2, 31)
          Call REPLA(X$, "RECEP-ESPUMA", 33, 16)
          Call REPLA(X$, "BOLREGEN", 49, 16)
          Call REPLA(X$, FORIPRE$, 65, 16)
       Else
          X$ = String$(80, 0)
    End If
    Call GRAREG("TATIREC", X$, 2)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Check1_CLICK(Index As Integer)
    If Index = 0 Then
        Check1(Index).Value = 1
    End If
End Sub

Private Sub Command8_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command9_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture15.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      
      Me.Check1(0).BackColor = &H8000000F
      Me.Check1(1).BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      Me.Check6.BackColor = &H8000000F
      Me.Check7.BackColor = &H8000000F
      Me.Check10.BackColor = &H8000000F
      
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      MARCOBARRA.Top = 4600
      Set MARCOBARRA.Container = Me.Picture15
      
    End If
    
    Text1.TEXT = CONTROL$("", "FORMBREC")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
    If CONTROL$("", "NUMEBREC") = "MANUAL" Then Option1.Value = True
    '
    CMO$ = TRIM$(CONTROL$("", "CODMBREC")): If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    Text3.TEXT = CONTROL$("", "REMIRECH")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    Text10.TEXT = CONTROL$("", "STIKAPRO")
    Text9.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    If Left$(UCase$(Text10), 3) = "STI" Then
      Check4.Value = 1
    End If
    '
    If CONTROL$("", "NUMERECH") = "MANUAL" Then Option5.Value = True
    '
    CMO$ = TRIM$(CONTROL$("", "CODRRECH")): If CMO$ <> "" Then Text6.TEXT = CMO$
    '
    Check1(1).Value = 0
    If TRIM$(CONTROL$("", "RECESPUM")) = "SI" Then
        Check1(1).Value = 1
    End If
    '
    If CONTROL$("BOLRECEP", "SUPRINT") = "SI" Then Check2.Value = 1
    '
    Option3.Value = True
    If TRIM$(CONTROL$("", "COMPBREC")) = "FICHA" Then
        Option4.Value = True
    End If
    '
    Text7 = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
    If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
      Check3.Value = 1
    End If
    '
    '\\\OT-06-0515-RG-28/11/06///
    If CONTROL$("BOLRECEP", "SUPREMI") = "SI" Then Check6.Value = 1
    '\\\OT-06-0515-RG-FIN///
    '
    If CONTROL$("BOLRECEP", "BLOSOBRE") = "SI" Then
      Check5.Value = 1
      Text8 = FORMATNUM(Val(CONTROL("BOLRECEP", "PORSOBRE")), "F5.1")
    End If
    '
    If CONTROL$("BOLRECEP", "NUCONVIG") = "SI" Then Check10.Value = 1
    If CONTROL$("BOLRECEP", "CERCAPRO") = "SI" Then Check7.Value = 1

End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub salir_Click()
  Call Command1_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text10_Change()
   If Left$(UCase$(Text10), 3) <> "STI" Then
      Check4.Value = 0
   End If
End Sub

Private Sub Text10_DblClick()
   Call Command9_Click
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_LostFocus()
   Text8 = FORMATNUM$(Val(Text8), "f5.1")
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

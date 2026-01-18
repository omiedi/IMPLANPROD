VERSION 5.00
Begin VB.Form CFGRECEP 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formularios"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11145
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
   ScaleHeight     =   5730
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Configuracion Deposito de Transito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5250
      TabIndex        =   56
      Top             =   4905
      Width           =   4950
      Begin VB.CommandButton Command13 
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
         Left            =   2250
         TabIndex        =   58
         Top             =   280
         Width           =   175
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
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
         Left            =   1860
         TabIndex        =   57
         Top             =   280
         Width           =   405
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito Transito:"
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
         TabIndex        =   60
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label Label11 
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
         Left            =   2520
         TabIndex        =   59
         Top             =   280
         Width           =   2205
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Etiqueta de Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   5235
      TabIndex        =   35
      Top             =   1815
      Width           =   4950
      Begin VB.TextBox Text19 
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
         Left            =   1440
         TabIndex        =   61
         Text            =   " "
         Top             =   1680
         Width           =   3300
      End
      Begin VB.TextBox Text16 
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
         Left            =   1440
         TabIndex        =   50
         Text            =   " "
         Top             =   1320
         Width           =   3300
      End
      Begin VB.TextBox Text15 
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
         Left            =   1440
         TabIndex        =   48
         Text            =   " "
         Top             =   960
         Width           =   3300
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fecha Mandatorias"
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
         Left            =   120
         TabIndex        =   41
         Top             =   2145
         Width           =   4560
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Fecha de Venc."
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
            Left            =   2580
            TabIndex        =   43
            Top             =   300
            Width           =   1680
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Fecha de Retest"
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
            Left            =   270
            TabIndex        =   42
            Top             =   300
            Width           =   1770
         End
      End
      Begin VB.CommandButton Command12 
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
         Height          =   650
         Left            =   4095
         Picture         =   "CFGRECEP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   270
         Width           =   645
      End
      Begin VB.TextBox Text12 
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
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3750
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
         Left            =   3795
         TabIndex        =   37
         Top             =   270
         Width           =   180
      End
      Begin VB.TextBox Text11 
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
         Left            =   2415
         TabIndex        =   36
         Text            =   " "
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Calidad  "
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
         TabIndex        =   62
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Compras"
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
         TabIndex        =   51
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Program."
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
         TabIndex        =   49
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label Label6 
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
         TabIndex        =   40
         Top             =   330
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresión de Vale de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   195
      TabIndex        =   21
      Top             =   4545
      Width           =   4950
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
         Height          =   300
         Left            =   2415
         TabIndex        =   25
         Text            =   " "
         Top             =   330
         Width           =   1380
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
         Height          =   300
         Left            =   3780
         TabIndex        =   24
         Top             =   330
         Width           =   180
      End
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
         Height          =   285
         Left            =   210
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
         Width           =   3750
      End
      Begin VB.CommandButton Command7 
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
         Height          =   650
         Left            =   4095
         Picture         =   "CFGRECEP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label7 
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
         TabIndex        =   26
         Top             =   390
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresión de Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   5250
      TabIndex        =   15
      Top             =   120
      Width           =   4950
      Begin VB.TextBox Text17 
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
         Left            =   1440
         TabIndex        =   54
         Text            =   " "
         Top             =   1250
         Width           =   3330
      End
      Begin VB.TextBox Text14 
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
         Left            =   1440
         TabIndex        =   46
         Text            =   " "
         Top             =   900
         Width           =   3330
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
         Height          =   300
         Left            =   2415
         TabIndex        =   19
         Text            =   " "
         Top             =   210
         Width           =   1380
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
         Height          =   300
         Left            =   3780
         TabIndex        =   18
         Top             =   210
         Width           =   180
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
         Left            =   210
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   570
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
         Height          =   650
         Left            =   4095
         Picture         =   "CFGRECEP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Com.Ext."
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
         Left            =   45
         TabIndex        =   55
         Top             =   1280
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Adminst."
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
         Left            =   45
         TabIndex        =   47
         Top             =   950
         Width           =   1365
      End
      Begin VB.Label Label3 
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
         TabIndex        =   20
         Top             =   270
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresión de Etiqueta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   210
      TabIndex        =   9
      Top             =   2640
      Width           =   4950
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Puerto de Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   210
         TabIndex        =   52
         Top             =   960
         Width           =   4540
         Begin VB.ComboBox Combo1 
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
            Left            =   360
            TabIndex        =   53
            Text            =   "Combo1"
            Top             =   280
            Width           =   3795
         End
      End
      Begin VB.CommandButton Command4 
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
         Height          =   650
         Left            =   4095
         Picture         =   "CFGRECEP.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   645
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   570
         Width           =   3750
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
         Height          =   300
         Left            =   3780
         TabIndex        =   11
         Top             =   210
         Width           =   180
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
         Height          =   300
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   210
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
         TabIndex        =   14
         Top             =   270
         Width           =   2325
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresión de Boleta de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   210
      TabIndex        =   3
      Top             =   120
      Width           =   4950
      Begin VB.TextBox Text13 
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
         Height          =   315
         Left            =   3360
         TabIndex        =   45
         Text            =   "0.0"
         Top             =   1020
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
         TabIndex        =   44
         Top             =   1050
         Width           =   4320
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
         Height          =   885
         Left            =   3870
         TabIndex        =   33
         Top             =   1440
         Width           =   855
         Begin VB.TextBox Text8 
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
            TabIndex        =   34
            Text            =   "BR"
            Top             =   420
            Width           =   450
         End
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
         Height          =   885
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
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
            TabIndex        =   32
            Top             =   525
            Value           =   -1  'True
            Width           =   1380
         End
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
            TabIndex        =   31
            Top             =   315
            Width           =   1380
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
         Height          =   885
         Left            =   1980
         TabIndex        =   27
         Top             =   1440
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
            TabIndex        =   28
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label2 
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
            TabIndex        =   29
            Top             =   315
            Width           =   1065
         End
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
         Height          =   300
         Left            =   2475
         TabIndex        =   7
         Text            =   " "
         Top             =   250
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
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Top             =   250
         Width           =   180
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
         Height          =   315
         Left            =   195
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   3840
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
         Height          =   650
         Left            =   4095
         Picture         =   "CFGRECEP.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   250
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Left            =   75
         TabIndex        =   8
         Top             =   330
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7140
      Left            =   10260
      ScaleHeight     =   7080
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   -30
      Width           =   1200
      Begin VB.CommandButton command1 
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
         Left            =   105
         Picture         =   "CFGRECEP.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   105
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
         Height          =   500
         Left            =   105
         Picture         =   "CFGRECEP.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guardar"
         Top             =   615
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CFGRECEP.frx":1386
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4560
         Width           =   1515
      End
   End
End
Attribute VB_Name = "CFGRECEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'\\\OT-05-0723-WAR-22/11/05///

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command11_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text11.TEXT = VDEVU$
        Text12.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command12_Click()
   Call CONECRUN("FLFORMS/" + TRIM$(Text11.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command13_Click()
  Command13.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  DEPO$ = VALACT1$("TADEPOSI")
    
  If DEPO$ <> "" Then
      Text18 = DEPO$
  End If
  '
  Command13.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEBREC", MONUME$)
    Call GRACONTROL("BOLRECEP", "DIANREC", TRIM$(Text7))
    Call GRACONTROL("", "CODMBREC", Text8.TEXT)
    Call GRACONTROL("", "FORMBREC", Text10.TEXT)
    '
    Call GRACONTROL("", "STIKAPRO", Text1.TEXT)
    Call GRACONTROL("", "FOREMTRA", Text4.TEXT)
    Call GRACONTROL("", "FORTRSTM", Text6.TEXT)
    Call GRACONTROL(IDENTER$, "PORTSTIK", Combo1.TEXT)
    Call GRACONTROL("*", "PORTSTIK", Combo1.TEXT)
    Call GRACONTROL("BOLRECEP", "FORAPRO", Text11.TEXT)
    '
    'Fechas mandatorias
    Call GRACONTROL("BOLRECEP", "IFEVEN", Check1.Value)
    Call GRACONTROL("BOLRECEP", "IFERTEST", Check2.Value)
    '
    'Sobre-Embarque
    BLOSOBRE$ = "NO"
    If Check5.Value = 1 Then BLOSOBRE$ = "SI"
    Call GRACONTROL("BOLRECEP", "BLOSOBRE", BLOSOBRE$)
    Call GRACONTROL("BOLRECEP", "PORSOBRE", Text13)
    '
    'E-mail para copia de remito de traslado del piking
    Call GRACONTROL("BOLRECEP", "REMIMAIL", Text14)
    'E-mail para administracion por rechazo de recepcion
    Call GRACONTROL("BOLRECEP", "PROGMAIL", Text15)
    'E-mail para compras por rechazo de recepciopn
    Call GRACONTROL("BOLRECEP", "COMPMAIL", Text16)
    'E-mail para Comercio Exterior.
    Call GRACONTROL("BOLRECEP", "COEXMAIL", Text17)
    'E-mail para Gestion de Calidad.
    Call GRACONTROL("BOLRECEP", "GECAMAIL", Text19)
    
    
    'Guarda el deposito para el material en transito
    If TRIM$(Text18) <> "" Then
        If ECOARCH$("TADEPOSI", Chr$(XVALO(Text18))) <> "" Then
            Call GRACONTROL("BOLRECEP", "DEPTRANS", Text18)
        Else
            Call MENSERR(24, "Depósito de Transito no Válido")
            Exit Sub
        End If
    End If
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

Private Sub Command4_Click()
 Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command5_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text3.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command7_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
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
    MONUME$ = CONTROL("", "NUMEBREC")
    If MONUME$ = "AUTO" Then
       Option1.Value = True
       Option2.Value = False
      Else
       Option1.Value = False
       Option2.Value = True
    End If
    Text7 = CONTROL("BOLRECEP", "DIANREC")
    Text10.TEXT = CONTROL$("", "FORMBREC")
    '
    Text9.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    Text1.TEXT = CONTROL$("", "STIKAPRO")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text4.TEXT = CONTROL$("", "FOREMTRA")
    Text3.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    Text6.TEXT = CONTROL$("", "FORTRSTM")
    Text5.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
   
    Text11.TEXT = CONTROL$("BOLRECEP", "FORAPRO")
    Text12.TEXT = ECOARCH$("INDIFRM", Text11.TEXT)
    '
    Check1.Value = XVALO(CONTROL("BOLRECEP", "IFEVEN"))
    Check2.Value = XVALO(CONTROL("BOLRECEP", "IFERTEST"))
    '
    'Sobre-Embarque
    If CONTROL$("BOLRECEP", "BLOSOBRE") = "SI" Then
       Check5.Value = 1
       Text13 = FORMATNUM(Val(CONTROL("BOLRECEP", "PORSOBRE")), "F5.1")
       Else
       Check5.Value = 0
    End If
    '
    'E-mail para copia de remito de traslado del piking
    Text14 = CONTROL$("BOLRECEP", "REMIMAIL")
    'E-mail para administracion por rechazo de recepcion
    Text15 = CONTROL$("BOLRECEP", "PROGMAIL")
    'E-mail para compras por rechazo de recepciopn
    Text16 = CONTROL$("BOLRECEP", "COMPMAIL")
    'E-mail para comercio exterior.
    Text17 = CONTROL$("BOLRECEP", "COEXMAIL")
    'E-mail para Gestion de Calidad.
    Text19 = CONTROL("BOLRECEP", "GECAMAIL")
    'Deposito configurado para material en transito
    Text18 = CONTROL$("BOLRECEP", "DEPTRANS")
    '
    COMP$ = CONTROL("", "CODMBREC")
    If TRIM$(COMP$) = "" Then COMP$ = "BR"
    Text8.TEXT = COMP$
    '
    PFAC$ = CONTROL$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTSTIK")
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo1.TEXT = PFAC$
      End If
    Next
    PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    If PFAC$ <> "" Then Combo1.TEXT = PFAC$
    '
End Sub

Private Sub Text13_LostFocus()
  Text13 = FORMATNUM$(Val(Text13), "f5.1")
End Sub

Private Sub Text18_Change()
    Label11 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text18)))
End Sub

Private Sub Text18_DblClick()
    Call Command13_Click
End Sub

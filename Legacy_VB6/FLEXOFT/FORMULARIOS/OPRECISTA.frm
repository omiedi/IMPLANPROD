VERSION 5.00
Begin VB.Form OPRECISTA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Recibo Standard de Cobranza"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10890
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   7560
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame14 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8400
      TabIndex        =   50
      Top             =   1560
      Width           =   2295
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresca Nume- ro de Cuenta"
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
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   1800
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Retiene el Anterior"
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
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5520
      TabIndex        =   47
      Top             =   1560
      Width           =   2760
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Nota de Credito"
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
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1260
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Suprime Impresión  al  Generar Documento"
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
         Left            =   240
         TabIndex        =   48
         Top             =   960
         Width           =   2220
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Acceso a Calculadora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   5520
      TabIndex        =   40
      Top             =   3240
      Width           =   5160
      Begin VB.CommandButton Command11 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         Picture         =   "OPRECISTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   420
         Width           =   540
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   210
         TabIndex        =   43
         Top             =   900
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   360
         Left            =   210
         TabIndex        =   42
         Top             =   420
         Width           =   1275
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   2940
         Pattern         =   "*.exe"
         TabIndex        =   41
         Top             =   420
         Width           =   2000
      End
      Begin VB.Line Line3 
         X1              =   2740
         X2              =   2740
         Y1              =   900
         Y2              =   3250
      End
      Begin VB.Line Line2 
         X1              =   225
         X2              =   2745
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label11 
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
         Height          =   300
         Left            =   240
         TabIndex        =   46
         Top             =   3660
         Width           =   4680
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Acceso:"
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
         Left            =   120
         TabIndex        =   45
         Top             =   3390
         Width           =   1590
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tablas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   5520
      TabIndex        =   37
      Top             =   270
      Width           =   1500
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
         Height          =   600
         Left            =   780
         Picture         =   "OPRECISTA.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Actualizar Tabla de Talonarios de Recibo y Cobradores"
         Top             =   240
         Width           =   555
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
         Height          =   600
         Left            =   160
         Picture         =   "OPRECISTA.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Actualizar Tabla de Valores y Medios de Cobranza"
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ajuste de Saldos Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   210
      TabIndex        =   31
      Top             =   5880
      Width           =   5160
      Begin VB.CommandButton Command9 
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPRECISTA.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   420
         Width           =   855
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
         Height          =   360
         Left            =   210
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
      Begin VB.CommandButton Command8 
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
         Left            =   3780
         TabIndex        =   33
         Top             =   420
         Width           =   175
      End
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
         Height          =   360
         Left            =   2415
         TabIndex        =   32
         Text            =   " "
         Top             =   420
         Width           =   1380
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
         Left            =   0
         TabIndex        =   36
         Top             =   600
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recibo de Cobranza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   210
      TabIndex        =   3
      Top             =   240
      Width           =   5160
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
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
         TabIndex        =   30
         Top             =   1470
         Width           =   4740
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   27
         Top             =   4200
         Width           =   1650
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Vencim."
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
            TabIndex        =   29
            Top             =   735
            Width           =   1275
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Emision"
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
            TabIndex        =   28
            Top             =   420
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prefijo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   3675
         TabIndex        =   19
         Top             =   2940
         Width           =   1275
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
            Height          =   360
            Left            =   210
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   525
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   18
         Top             =   2940
         Width           =   3330
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   380
            Left            =   315
            TabIndex        =   24
            Top             =   735
            Width           =   2850
            Begin VB.OptionButton Option6 
               BackColor       =   &H00C0E0FF&
               Caption         =   "NewPage"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   26
               Top             =   0
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00C0E0FF&
               Caption         =   "EndDoc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   435
            Left            =   315
            TabIndex        =   21
            Top             =   315
            Width           =   2850
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Formulario Continuo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Hoja Suelta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1575
               TabIndex        =   22
               Top             =   0
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Talonario en Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1995
         TabIndex        =   14
         Top             =   4200
         Width           =   2955
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
            Height          =   360
            Left            =   2580
            TabIndex        =   16
            Top             =   525
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
            Height          =   360
            Left            =   1425
            TabIndex        =   15
            Text            =   " "
            Top             =   525
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración"
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
            Left            =   165
            TabIndex        =   17
            Top             =   480
            Width           =   1170
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
         Height          =   360
         Left            =   2415
         TabIndex        =   13
         Text            =   " "
         Top             =   420
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
         Height          =   360
         Left            =   3780
         TabIndex        =   12
         Top             =   420
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
         Height          =   360
         Left            =   210
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPRECISTA.frx":0A18
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   6
         Top             =   2100
         Width           =   3750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pre-Impresa"
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
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Auto-Numeración"
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
            Left            =   1890
            TabIndex        =   7
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
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
         Left            =   4095
         TabIndex        =   4
         Top             =   2100
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
            Height          =   360
            Left            =   190
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "RB"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5145
         Y1              =   1890
         Y2              =   1890
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
         TabIndex        =   9
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
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
      Left            =   8280
      Picture         =   "OPRECISTA.frx":0D22
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   9105
      Picture         =   "OPRECISTA.frx":102C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   735
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
      Height          =   855
      Left            =   9930
      Picture         =   "OPRECISTA.frx":1336
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6670
      Picture         =   "OPRECISTA.frx":1480
      Top             =   690
      Width           =   2010
   End
End
Attribute VB_Name = "OPRECISTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CARTACOBRA
    Command10.Enabled = True
End Sub

Private Sub Command13_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
        FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text3.SetFocus
        Exit Sub
    End If
    '
    Call GRACONTROL("", "FORECIST", Text1.TEXT)
    Call GRACONTROL("", "FORAJUSA", Text8.TEXT)
    Call GRACONTROL("", "CODMRECI", "RB")    ' Text5.TEXT)
    '
    MONUME$ = "AUTONUME": If Option1.Value = True Then MONUME$ = "PREIMPRE"
    Call GRACONTROL("", "NURECIST", MONUME$)
    '
    TIPAPEL$ = "FORMUCON": If Option4.Value = True Then TIPAPEL$ = "HOJASUEL"
    Call GRACONTROL("", "TIPARECI", TIPAPEL$)
    '
    AVAHOJA$ = "ENDDOC": If Option6.Value = True Then AVAHOJA$ = "NEWPAGE"
    Call GRACONTROL("RECISTA", "AVAHOJA", AVAHOJA$)
    '
    MODORDE$ = "EMISION": MODORECI% = 1
    If Option8.Value = True Then
        MODORDE$ = "VENCIM"
        MODORECI% = 2
    End If
    Call GRACONTROL("RECISTA", "MODORDE", MODORDE$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("RECISTA", "SUPRINT", SUPRINT$)
    '
    DESCUEX$ = ""
    If Check3.Value = 1 Then DESCUEX$ = "CREDESCU"
    Call GRACONTROL("RECISTA", "DESCAJA", DESCUEX$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check1.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("CREDESCU", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = ""
    If Option10.Value = True Then SUPRINT$ = "REFRESH"
    Call GRACONTROL("RECISTA", "CLIENTE", SUPRINT$)
    '
    Call GRACONTROL("", "FENURECI", FENUE$)
    '
    PREFIREM$ = TRIM$(Text4.TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
    If PREFIREM$ <> "" Then
       If Len(PREFIREM$) <> 4 Then
          Call COMUNI("Prefijo no Válido")
          Text4.SetFocus
          Exit Sub
       End If
       For U& = 1 To 4
          UI% = Asc(Mid$(PREFIREM$, U&))
          If UI% < 48 Or UI% > 57 Then
             Call COMUNI("Prefijo no Válido")
             Text4.SetFocus
             Exit Sub
          End If
       Next U&
    End If
    Call GRACONTROL("", "PREFRECI", PREFIREM$)
    Call GRACONTROL(IDENTER$, "RUT-CALC", Dir1.Path)
    Call GRACONTROL(IDENTER$, "EXE-CALC", File1.filename)
    '
    Call CIERRARCH("*.*")
    '
    FORECISTA.List1.Clear
    FORECISTA.List2.Clear
    FORECISTA.Label4.Caption = ""
    FORECISTA.Label8.Caption = ""
    FORECISTA.Label10.Caption = ""
    FORECISTA.Label14(0).Caption = ""
    FORECISTA.Label18(0).Caption = ""
    FORECISTA.Label20(0).Caption = ""
    FORECISTA.BOTNEXT.Enabled = True
    FORECISTA.Command5.Enabled = False
    '
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
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CARTAMECO
    Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command9_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FORECIST")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = CONTROL$("", "FENURECI")
    Text8.TEXT = CONTROL$("", "FORAJUSA")
    Text9.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    '
    If CONTROL$("", "NURECIST") = "AUTONUME" Then Option2.Value = True
    If CONTROL$("", "TIPARECI") = "HOJASUEL" Then Option4.Value = True
    If CONTROL$("RECISTA", "AVAHOJA") = "NEWPAGE" Then Option6.Value = True
    If CONTROL$("RECISTA", "SUPRINT") = "SI" Then Check2.Value = 1
    If CONTROL$("RECISTA", "DESCAJA") = "CREDESCU" Then Check3.Value = 1
    If CONTROL$("CREDESCU", "SUPRINT") = "SI" Then Check1.Value = 1
    If CONTROL$("RECISTA", "CLIENTE") = "REFRESH" Then Option10.Value = True
    MODORECI% = 1: If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then Option8.Value = True: MODORECI% = 2
    '
    PREFIREM$ = TRIM$(CONTROL$("", "PREFRECI"))
    Text4.TEXT = PREFIREM$
    Text5.TEXT = "RB"
    '
    Dir1.Path = CONTROL(IDENTER$, "RUT-CALC")
    FISEL$ = CONTROL(IDENTER$, "EXE-CALC")
    For KKI% = 1 To File1.ListCount
      If File1.List(KKI% - 1) = FISEL$ Then
        File1.ListIndex = KKI% - 1
      End If
    Next KKI%
    Label11.Caption = Dir1.Path + "\" + FISEL$
    Label11.Refresh
    '
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   If Err Then
      Drive1.Drive = Left$(CurDir, 2)
      Exit Sub
   End If
End Sub

Private Sub File1_Click()
   Label11.Caption = Dir1.Path + "\" + File1.filename
   Label11.Refresh
End Sub

Private Sub Option11_Click()
   If Option11.Value = True Then
     Call MENSERR(24, "No Disponible en la Presente Version")
     Option9.Value = True
   End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Text5_Click()
    Command2.SetFocus
End Sub

Private Sub Text6_Click()
    Text7.SetFocus
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text9_Click()
    Text8.SetFocus
End Sub

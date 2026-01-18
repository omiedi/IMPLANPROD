VERSION 5.00
Begin VB.Form DOCUMI03 
   BackColor       =   &H00F0EFF0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos de Ingeniería"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800080&
      Height          =   8895
      Left            =   8820
      ScaleHeight     =   8835
      ScaleWidth      =   9675
      TabIndex        =   68
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton Command39 
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   5850
         Width           =   650
      End
      Begin VB.CommandButton Command37 
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5160
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Copia Procedimientos Asociados desde otro Código"
         Top             =   3690
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Alta de Nuevo Procedimiento / Instrucción"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Height          =   690
         Left            =   120
         Picture         =   "DOCUMI03.frx":0D40
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1620
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   930
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
         Left            =   120
         Picture         =   "DOCUMI03.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "DOCUMI03.frx":149E
         Top             =   7200
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0EFF0&
      Caption         =   "Plan de Inspección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   210
      TabIndex        =   28
      Top             =   5565
      Width           =   8310
      Begin VB.CommandButton Command2 
         Height          =   645
         Left            =   7350
         Picture         =   "DOCUMI03.frx":33C0
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1365
         Width           =   750
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0EFF0&
         Caption         =   "Puntos de Detención Obligatoria"
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
         Index           =   3
         Left            =   315
         TabIndex        =   35
         Top             =   1260
         Width           =   6840
         Begin VB.Label Label4 
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
            Height          =   300
            Index           =   3
            Left            =   420
            TabIndex        =   36
            Top             =   315
            Width           =   6210
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0EFF0&
         Caption         =   "Frecuencia de Inspección"
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
         Index           =   2
         Left            =   2835
         TabIndex        =   33
         Top             =   420
         Width           =   5265
         Begin VB.Label Label4 
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
            Height          =   300
            Index           =   2
            Left            =   315
            TabIndex        =   34
            Top             =   315
            Width           =   4740
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0EFF0&
         Caption         =   "Nivel"
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
         Index           =   1
         Left            =   1470
         TabIndex        =   31
         Top             =   420
         Width           =   1275
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   945
            TabIndex        =   37
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label4 
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
            Height          =   300
            Index           =   1
            Left            =   315
            TabIndex        =   32
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0EFF0&
         Caption         =   "AQL"
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
         Index           =   0
         Left            =   315
         TabIndex        =   29
         Top             =   420
         Width           =   1065
         Begin VB.Label Label4 
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
            Height          =   300
            Index           =   0
            Left            =   420
            TabIndex        =   30
            Top             =   315
            Width           =   435
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F0EFF0&
      Caption         =   "Documentos Anexos - Instrucciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   210
      TabIndex        =   11
      Top             =   1575
      Width           =   8310
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   7200
         TabIndex        =   67
         ToolTipText     =   "Desconecta Documento"
         Top             =   2940
         Width           =   230
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   7200
         TabIndex        =   66
         ToolTipText     =   "Desconecta Documento"
         Top             =   2070
         Width           =   230
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   7200
         TabIndex        =   65
         ToolTipText     =   "Desconecta Documento"
         Top             =   1230
         Width           =   230
      End
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   7200
         TabIndex        =   64
         ToolTipText     =   "Desconecta Documento"
         Top             =   390
         Width           =   230
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         Index           =   3
         Left            =   7530
         Picture         =   "DOCUMI03.frx":36CA
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Edición de Documento"
         Top             =   2940
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         Index           =   2
         Left            =   7530
         Picture         =   "DOCUMI03.frx":39D4
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Edición de Documento"
         Top             =   2070
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         Index           =   1
         Left            =   7530
         Picture         =   "DOCUMI03.frx":3CDE
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Edición de Documento"
         Top             =   1230
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         Index           =   0
         Left            =   7530
         Picture         =   "DOCUMI03.frx":3FE8
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   390
         Width           =   540
      End
      Begin VB.CommandButton Command31 
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
         Index           =   3
         Left            =   3675
         TabIndex        =   21
         Top             =   2940
         Width           =   175
      End
      Begin VB.CommandButton Command31 
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
         Index           =   2
         Left            =   3675
         TabIndex        =   18
         Top             =   2100
         Width           =   175
      End
      Begin VB.CommandButton Command31 
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
         Index           =   1
         Left            =   3675
         TabIndex        =   15
         Top             =   1260
         Width           =   175
      End
      Begin VB.CommandButton Command31 
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
         Index           =   0
         Left            =   3675
         TabIndex        =   12
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label9 
         Height          =   225
         Index           =   3
         Left            =   5250
         TabIndex        =   63
         Top             =   2835
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label9 
         Height          =   225
         Index           =   2
         Left            =   5250
         TabIndex        =   62
         Top             =   1995
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label9 
         Height          =   225
         Index           =   1
         Left            =   5250
         TabIndex        =   61
         Top             =   1155
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label9 
         Height          =   225
         Index           =   0
         Left            =   5250
         TabIndex        =   60
         Top             =   210
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Index           =   3
         Left            =   6090
         TabIndex        =   59
         Top             =   2940
         Width           =   1000
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   11
         Left            =   5250
         TabIndex        =   58
         Top             =   3045
         Width           =   750
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Index           =   3
         Left            =   4515
         TabIndex        =   57
         Top             =   2940
         Width           =   645
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Index           =   10
         Left            =   3885
         TabIndex        =   56
         Top             =   3045
         Width           =   540
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Index           =   2
         Left            =   6090
         TabIndex        =   54
         Top             =   2100
         Width           =   1000
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   9
         Left            =   5250
         TabIndex        =   53
         Top             =   2205
         Width           =   750
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Index           =   2
         Left            =   4515
         TabIndex        =   52
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Index           =   8
         Left            =   3885
         TabIndex        =   51
         Top             =   2205
         Width           =   540
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Index           =   1
         Left            =   6090
         TabIndex        =   49
         Top             =   1260
         Width           =   1000
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   7
         Left            =   5250
         TabIndex        =   48
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Index           =   1
         Left            =   4515
         TabIndex        =   47
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Index           =   6
         Left            =   3885
         TabIndex        =   46
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Index           =   0
         Left            =   6090
         TabIndex        =   44
         Top             =   420
         Width           =   1000
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   5
         Left            =   5250
         TabIndex        =   43
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Index           =   0
         Left            =   4515
         TabIndex        =   42
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Index           =   4
         Left            =   3885
         TabIndex        =   41
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Index           =   3
         Left            =   1890
         TabIndex        =   27
         Top             =   2940
         Width           =   1800
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Index           =   2
         Left            =   1890
         TabIndex        =   26
         Top             =   2100
         Width           =   1800
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Index           =   1
         Left            =   1890
         TabIndex        =   25
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Index           =   0
         Left            =   1890
         TabIndex        =   24
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label40 
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
         Height          =   300
         Index           =   3
         Left            =   1890
         TabIndex        =   23
         Top             =   3300
         Width           =   6200
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones de Ensayos:"
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
         Index           =   3
         Left            =   420
         TabIndex        =   22
         Top             =   2835
         Width           =   1380
      End
      Begin VB.Label Label40 
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
         Height          =   300
         Index           =   2
         Left            =   1890
         TabIndex        =   20
         Top             =   2460
         Width           =   6200
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Reposición:"
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
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   1995
         Width           =   1800
      End
      Begin VB.Label Label40 
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
         Height          =   300
         Index           =   1
         Left            =   1890
         TabIndex        =   17
         Top             =   1620
         Width           =   6200
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compras y Abastecimientos:"
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
         Index           =   1
         Left            =   -420
         TabIndex        =   16
         Top             =   1155
         Width           =   2220
      End
      Begin VB.Label Label40 
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
         Height          =   300
         Index           =   0
         Left            =   1890
         TabIndex        =   14
         Top             =   780
         Width           =   6200
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Almacenamiento:"
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
         Index           =   0
         Left            =   -420
         TabIndex        =   13
         Top             =   525
         Width           =   2220
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1110
      Left            =   8295
      ScaleHeight     =   1050
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   315
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12000
      End
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
      Left            =   1575
      MaxLength       =   15
      TabIndex        =   0
      Top             =   315
      Width           =   2040
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
      Left            =   3600
      Picture         =   "DOCUMI03.frx":42F2
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.Label Label6 
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
      Height          =   540
      Left            =   1575
      TabIndex        =   39
      Top             =   840
      Width           =   6525
   End
   Begin VB.Label Label58 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   210
      TabIndex        =   10
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label37 
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
      Height          =   330
      Left            =   5295
      TabIndex        =   6
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   6195
      TabIndex        =   5
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   7350
      TabIndex        =   4
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   105
      TabIndex        =   3
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   375
      TabIndex        =   2
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "DOCUMI03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIDOC%
Dim GRABATODO%

Private Sub Command1_Click()
   Command1.Enabled = False
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If VALCODART% < 1 Then GoTo 99
   '
   Call GRADATITEM
   PLAINSPEC.Label28 = Label28
   PLAINSPEC.Show 1
   '
   Call CARDATITEM
99 Command2.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Call GRADATITEM
   Call NUEINSTRU
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If VALCODART% > 0 Then
     Call SELECHO("MASTER/Seleccione ORIGEN de Procedimientos Anexos")
     COD$ = TRIM$(VALACT1$("MASTER"))
     If COD$ <> "" Then
       CODIORIG% = CODINT%(COD$)
       If CODIORIG% > 0 Then
         Call COPYANEXOS(CODIORIG%)
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Call GRADATITEM
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub COMMAND31_CLICK(Index As Integer)
   If VALCODART% < 1 Then Exit Sub
   XX$ = REGSEL$("INSTRUCI")
   If Len(XX$) > 4 Then
     CODINSTRU$ = Left$(XX$, 16)
     DESINSTRU$ = Mid$(XX$, 17, 48)
     REVINSTRU$ = Mid$(XX$, 65, 4)
     CIIINSTRU% = CVI(Mid$(XX$, 69, 2))
     FEAPRO = CVI(Mid$(XX$, 121, 2))
     '
     Label2(Index) = CODINSTRU$
     Label40(Index) = DESINSTRU$
     Label7(Index) = REVINSTRU$
     Label8(Index) = FECHATEX$(FEAPRO)
     Label9(Index) = CIIINSTRU%
   End If
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   FORDOCINGE.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   PRINTFORM
   Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Call MENSERR(24, "Ayuda no Disponible")
End Sub

Private Sub Command5_Click(Index As Integer)
   If VALCODART% < 1 Then Exit Sub
   CIIINSTRU% = Val(Label9(Index))
   If CIIINSTRU% < 1 Or CIIINSTRU% > ULTREG&("INSTRUCI") Then
      Call MENSERR(24, "Falta Seleccionar Procedimiento")
      Exit Sub
   End If
   '
   INSTRUING.Label28 = TRIM$(Str$(CIIINSTRU%))
   INSTRUING.Show 1
   '
   Call CARDATITEM
   '
End Sub

Private Sub Command6_Click(Index As Integer)
   Label2(Index) = ""
   Label40(Index) = ""
   Label7(Index) = ""
   Label8(Index) = FECHATEX$(0)
   Label9(Index).Caption = ""
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label2_Change(Index As Integer)
   MODIDOC% = 1
End Sub

Private Sub Label2_DblClick(Index As Integer)
   If VALCODART% < 1 Then Exit Sub
   XX$ = REGSEL$("INSTRUCI")
   If Len(XX$) > 4 Then
     CODINSTRU$ = Left$(XX$, 16)
     DESINSTRU$ = Mid$(XX$, 17, 48)
     REVINSTRU$ = Mid$(XX$, 65, 4)
     CIIINSTRU% = CVI(Mid$(XX$, 69, 2))
     FEAPRO = CVI(Mid$(XX$, 121.2))
     '
     Label2(Index) = CODINSTRU$
     Label40(Index) = DESINSTRU$
     Label7(Index) = REVINSTRU$
     Label8(Index) = FECHATEX$(FEAPRO)
     Label9(Index) = CIIINSTRU%
   End If
End Sub

Private Sub Text1_Change()
   '
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRADATITEM
   '
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
          Call MENSERR(24, "Código Inexistente o no Válido")
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Label6 = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    '
    RFF$ = RECFILT$("CODIPROC", 1, MKI$(CI%))
    If Len(RFF$) >= 4 Then
      RECODIPRO& = CVS(Left$(RFF$, 4))
      If RECODIPRO& > 0 Then
        If RECODIPRO& <= ULTREG&("CODIPROC") Then
          YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
          Label4(0) = TRIM$(Str$(CVI(Mid$(YY$, 17, 2))))
          Label4(1) = TRIM$(Str$(CVI(Mid$(YY$, 19, 2))))
          Label4(2) = Mid$(YY$, 21, 32)
          Label4(3) = Mid$(YY$, 53, 48)
          For KKI% = 0 To 3
            CODINSTRU% = CVI(Mid$(YY$, 3 + 2 * KKI%, 2))
            If CODINSTRU% > 0 Then
              RECODIN& = CODINSTRU%
              If RECODIN& <= ULTREG&("INSTRUCI") Then
                XX$ = REGLEIDO$("INSTRUCI", RECODIN&)
                Label2(KKI%) = Left$(XX$, 16)
                Label40(KKI%) = Mid$(XX$, 17, 48)
                Label7(KKI%) = Mid$(XX$, 65, 4)
                FEAPRO = CVI(Mid$(XX$, 121, 2))
                Label8(KKI%) = FECHATEX$(FEAPRO)
                Label9(KKI%) = CVI(Mid$(XX$, 69, 2))
              End If
            End If
          Next KKI%
        End If
      End If
    End If
    '
    MODIDOC% = 0
    Screen.MousePointer = 1
    '
End Sub

Sub BLANFORMU()
   '
   Label6.Caption = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
   For KKI% = 0 To 3
     Label2(KKI%) = ""
     Label4(KKI%) = ""
     Label7(KKI%) = ""
     Label8(KKI%) = ""
     Label9(KKI%) = ""
     Label40(KKI%) = ""
   Next KKI%
   OPENFORMS = DoEvents
   '
End Sub

Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   If MODIDOC% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado los Documentos Anexos del Artículo\'" + TRIM$(Label37) + " - " + TRIM$(Label6) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       Call CARDATITEM
       MODIDOC% = 0
       Exit Sub
     End If
   End If
   '
   Call BLOQARCH("CODIPROC")
   URECODI& = ULTREG&("CODIPROC")
   RECODIPRO& = 0
   '
   RFF$ = RECFILT$("CODIPROC", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
       RECODIPRO& = CVS(Left$(RFF$, 4))
   End If
   If RECODIPRO& < 1 Or RECODIPRO& > URECODI& Then
         RECODIPRO& = 1 + URECODI&
         YY$ = String$(256, 0)
      Else
         YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
   End If
   '
   Call REPLA(YY$, MKI$(CI%), 1, 2)
   For KKI% = 0 To 3
     CODINSTRU% = Val(Label9(KKI%))
     Call REPLA(YY$, MKI$(CODINSTRU%), 3 + 2 * KKI%, 2)
   Next KKI%
   Call GRAREG("CODIPROC", YY$, RECODIPRO&)
   '
   Call CIERRARCH("CODIPROC")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub

Function VALCODART%()
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
End Function
'
Sub COPYANEXOS(CODORIG%)
    '
    If CODORIG% < 1 Then
       Call MENSERR(24, "Código Inexistente o no Válido")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    RFF$ = RECFILT$("CODIPROC", 1, MKI$(CODORIG%))
    If Len(RFF$) >= 4 Then
      RECODIPRO& = CVS(Left$(RFF$, 4))
      If RECODIPRO& > 0 Then
        If RECODIPRO& <= ULTREG&("CODIPROC") Then
          YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
          For KKI% = 0 To 3
            CODINSTRU% = CVI(Mid$(YY$, 3 + 2 * KKI%, 2))
            Label2(KKI%) = ""
            Label40(KKI%) = ""
            Label7(KKI%) = ""
            Label8(KKI%) = ""
            Label9(KKI%) = ""
            If CODINSTRU% > 0 Then
              RECODIN& = CODINSTRU%
              If RECODIN& <= ULTREG&("INSTRUCI") Then
                XX$ = REGLEIDO$("INSTRUCI", RECODIN&)
                Label2(KKI%) = Left$(XX$, 16)
                Label40(KKI%) = Mid$(XX$, 17, 48)
                Label7(KKI%) = Mid$(XX$, 65, 4)
                FEAPRO = CVI(Mid$(XX$, 121, 2))
                Label8(KKI%) = FECHATEX$(FEAPRO)
                Label9(KKI%) = CVI(Mid$(XX$, 69, 2))
              End If
            End If
          Next KKI%
        End If
      End If
    End If
    '
    Screen.MousePointer = 1
    '
End Sub

'
Sub NUEINSTRU()
   '
   HOII% = HOY(HO$)
   VDEF$ = REGBLAN$("INSTRUCI")
   Call REPLA(VDEF$, USERNAME$, 71, 24)
   Call REPLA(VDEF$, MKI$(HOII%), 95, 2)
   INDINS% = 1 + ULTREG&("INSTRUCI")
   Call REPLA(VDEF$, MKI$(INDINS%), 69, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEINSTRUC", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 16)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Código de Instrucción")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 16)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("INSTRUCI")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("INSTRUCI", UI&)
      If UCase$(TRIM$(Left$(XX$, 16))) = CODINS$ Then
         Call MENSERR(24, "Código Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   DESINS$ = TRIM$(Mid$(OBX$, 17, 48))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Descripción de la Instrucción")
      GoTo 10
   End If
   '
   REVINS$ = TRIM$(Mid$(OBX$, 65, 4))
   If REVINS$ = "" Then
      Call MENSERR(24, "Falta Revisión de la Instrucción")
      GoTo 10
   End If
   '
   PREPINS$ = TRIM$(Mid$(OBX$, 71, 24))
   If PREPINS$ = "" Then
      Call MENSERR(24, "Falta 'Preparó' de la Instrucción")
      GoTo 10
   End If
   '
   FECHINS% = CVI(Mid$(OBX$, 95, 2))
   If FECHINS% < 1 Then
      Call MENSERR(24, "Falta Fecha de Preparación\de la Instrucción")
      GoTo 10
   End If
   '
   Call CIERRARCH("INSTRUCI")
   Call BLOQARCH("INSTRUCI")
   INDINS% = 1 + ULTREG&("INSTRUCI")
   Call REPLA(OBX$, MKI$(INDINS%), 69, 2)
   Call REGAPP("INSTRUCI", OBX$)
   Call CIERRARCH("INSTRUCI")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   INSTRUING.Label28 = TRIM$(Str$(INDINS%))
   On Error Resume Next
   INSTRUING.Show 1
   On Error GoTo 0
   '
End Sub



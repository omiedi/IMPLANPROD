VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPFACES00 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formularios de Nota de Débito y Crédito"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5775
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
   ScaleHeight     =   7875
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "NOTA DE DÉBITO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3200
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   4635
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
         Left            =   3465
         TabIndex        =   43
         Top             =   2430
         Width           =   175
      End
      Begin VB.CommandButton Command28 
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2420
         Width           =   650
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   41
         Top             =   2420
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   32
         Top             =   945
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   29
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command21 
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   945
         Width           =   650
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
         Left            =   3465
         TabIndex        =   33
         Top             =   945
         Width           =   175
      End
      Begin VB.CommandButton Command20 
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   210
         Width           =   650
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
         Left            =   3465
         TabIndex        =   30
         Top             =   210
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   28
         Top             =   1680
         Width           =   1065
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
         Left            =   3465
         TabIndex        =   27
         Top             =   1680
         Width           =   175
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
         Width           =   650
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Fact. Crédito Tipo 'A'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   0
         TabIndex        =   45
         Top             =   2490
         Width           =   2325
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   44
         Top             =   2790
         Width           =   3480
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   40
         Top             =   1320
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'B'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   39
         Top             =   1040
         Width           =   2325
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   38
         Top             =   585
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'A'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   37
         Top             =   295
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'E'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   0
         TabIndex        =   36
         Top             =   1780
         Width           =   2325
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   35
         Top             =   2055
         Width           =   3480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "NOTA DE CRÉDITO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3200
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   4635
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2420
         Width           =   650
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
         Left            =   3465
         TabIndex        =   47
         Top             =   2430
         Width           =   175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   46
         Top             =   2420
         Width           =   1065
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   10
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   18
         Top             =   210
         Width           =   1065
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
         Left            =   3465
         TabIndex        =   17
         Top             =   210
         Width           =   175
      End
      Begin VB.CommandButton Command22 
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Width           =   650
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   2415
         TabIndex        =   15
         Top             =   945
         Width           =   1065
      End
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
         Height          =   300
         Index           =   3
         Left            =   3465
         TabIndex        =   14
         Top             =   945
         Width           =   175
      End
      Begin VB.CommandButton Command23 
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   945
         Width           =   650
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
         Height          =   660
         Left            =   3780
         Picture         =   "OPFACES00.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   3465
         TabIndex        =   11
         Top             =   1680
         Width           =   175
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   50
         Top             =   2795
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Fact. Crédito Tipo 'A'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   0
         TabIndex        =   49
         Top             =   2490
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'A'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   23
         Top             =   585
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'B'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   22
         Top             =   1040
         Width           =   2325
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   1320
         Width           =   3480
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   2055
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Tipo 'E'"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   0
         TabIndex        =   19
         Top             =   1770
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3675
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "IMPUTACIÓN POR DEFAULT DE CRÉDITOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   6680
      Width           =   4635
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cheques Rechazados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Cheques Rechazados"
         Top             =   360
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Diferencia de Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   1500
         TabIndex        =   5
         ToolTipText     =   "Diferencia de Cambio"
         Top             =   360
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Otros Débitos / Créditos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   7
         Left            =   3000
         TabIndex        =   4
         ToolTipText     =   "Varios"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   4920
      ScaleHeight     =   8490
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACES00.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPFACES00.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   930
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPFACES00.frx":1CA4
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPFACES00.frx":1D1E
      Top             =   1200
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPFACES00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Label1.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command11_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text2.TEXT = VDEVU$
        Label2.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command12_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Label3.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command13_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Label5.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command14_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Label7.Caption = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command18_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text18.TEXT = VDEVU$
        Label18.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORNDE-A", Text1.TEXT)
    Call GRACONTROL("", "FORNDE-B", Text2.TEXT)
    Call GRACONTROL("", "FORNCR-A", Text3.TEXT)
    Call GRACONTROL("", "FORNCR-B", Text4.TEXT)
    Call GRACONTROL("", "NDBCRE-A", Text18.TEXT)
    Call GRACONTROL("", "NCRCRE-A", Text7.TEXT)
    '
    'EXPORTACION
    Call GRACONTROL("", "FORNDE-E", Text5.TEXT)
    Call GRACONTROL("", "FORNCR-E", Text6.TEXT)

    If Option1(5).Value = True Then
      Call GRACONTROL("", "IMPUNCRE", "CHEQUESRECHAZADOS")
    ElseIf Option1(6).Value = True Then
      Call GRACONTROL("", "IMPUNCRE", "DIFERENCIACAMBIO")
    ElseIf Option1(7).Value = True Then
      Call GRACONTROL("", "IMPUNCRE", "VARIOS")
    End If
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command28_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text18.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Label6 = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
  Call CONECRUN(ACONEC$ + "/" + TRIM$(Text5.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command20_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command21_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text2.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command22_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label8.Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command8_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPVALEMAT.Show 1
   Command9.Enabled = True
End Sub

Private Sub Command23_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
'    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Picture16.Visible = True
'      Picture16.ZOrder 0
'      Me.BackColor = &HFCD7B6
'      Me.Frame1.BackColor = &HFCD7B6
'      Me.Frame2.BackColor = &HFCD7B6
'      Me.Frame3.BackColor = &HFCD7B6
'      Me.Frame4.BackColor = &HFCD7B6
'      Me.Option1(5).BackColor = &HFCD7B6
'      Me.Option1(6).BackColor = &HFCD7B6
'      Me.Option1(7).BackColor = &HFCD7B6
'
'      MARCOBARRA.Top = 6380
'
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
'      ' &H8000000F
'    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
   
    Text1.TEXT = Control$("", "FORNDE-A")
    Label1.Caption = ECOARCH$("INDIFRM", Text1.TEXT)
    Text2.TEXT = Control$("", "FORNDE-B")
    Label2.Caption = ECOARCH$("INDIFRM", Text2.TEXT)
    Text3.TEXT = Control$("", "FORNCR-A")
    Label3.Caption = ECOARCH$("INDIFRM", Text3.TEXT)
    Text4.TEXT = Control$("", "FORNCR-B")
    Label5.Caption = ECOARCH$("INDIFRM", Text4.TEXT)
    Text18.TEXT = Control$("", "NDBCRE-A")
    Label18.Caption = ECOARCH$("INDIFRM", Text18.TEXT)
    Text7.TEXT = Control$("", "NCRCRE-A")
    Label8.Caption = ECOARCH$("INDIFRM", Text7.TEXT)
    '
    'EXPORTACION
    Text5.TEXT = Control$("", "FORNDE-E")
    Label6.Caption = ECOARCH$("INDIFRM", Text5.TEXT)
    Text6.TEXT = Control$("", "FORNCR-E")
    Label7.Caption = ECOARCH$("INDIFRM", Text6.TEXT)
    '
    Option1(5).Value = True
    If Control$("", "IMPUNCRE") = "DIFERENCIACAMBIO" Then
      Option1(6).Value = True
    ElseIf Control$("", "IMPUNCRE") = "VARIOS" Then
      Option1(7).Value = True
    End If
        
    '
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text1_DblClick()
     Call Command10_Click
      'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text1.TEXT = VDEVU$
'        Label1.Caption = VALACT2$("INDIFRM")
'    End If
End Sub

Private Sub Text18_DblClick()
    Call Command18_Click
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
    'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text2.TEXT = VDEVU$
'        Label2.Caption = VALACT2$("INDIFRM")
'    End If
End Sub

Private Sub Text4_DblClick()
    Call Command13_Click(3)
     'COMENTADO PARA NO REPLICAR CODIGO
     'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text4.TEXT = VDEVU$
'        Label5.Caption = VALACT2$("INDIFRM")
'    End If
End Sub

Private Sub Text3_DBLCLICK()
   Call Command12_Click
    'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text3.TEXT = VDEVU$
'        Label3.Caption = VALACT2$("INDIFRM")
'    End If
End Sub


Private Sub Text5_DblClick()
    Call Command3_Click
    'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text5.TEXT = VDEVU$
'        Label6 = VALACT2$("INDIFRM")
'    End If
End Sub

Private Sub Text6_DBLCLICK()
    Call Command14_Click
     'COMENTADO PARA NO REPLICAR CODIGO
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text6.TEXT = VDEVU$
'        Label7.Caption = VALACT2$("INDIFRM")
'    End If
End Sub

Private Sub Text7_DblClick()
    Call Command6_Click
End Sub

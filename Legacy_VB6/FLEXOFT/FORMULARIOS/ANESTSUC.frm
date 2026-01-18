VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form ANESTSUC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Estadistico - Consumos de Stock"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ultimo Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   4830
      TabIndex        =   65
      Top             =   315
      Width           =   5895
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Período"
         Height          =   750
         Left            =   210
         TabIndex        =   69
         Top             =   735
         Width           =   4425
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Des:"
            Height          =   225
            Left            =   -210
            TabIndex        =   73
            Top             =   375
            Width           =   750
         End
         Begin VB.Label label1 
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
            Index           =   11
            Left            =   575
            TabIndex        =   72
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Has:"
            Height          =   225
            Left            =   1850
            TabIndex        =   71
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label15 
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
            Left            =   2625
            TabIndex        =   70
            Top             =   315
            Width           =   1620
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Meses"
         Height          =   1170
         Left            =   4750
         TabIndex        =   67
         Top             =   315
         Width           =   900
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   105
            TabIndex        =   68
            Top             =   420
            Width           =   675
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   66
         Top             =   2310
         Width           =   4425
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   1155
         TabIndex        =   78
         Top             =   360
         Width           =   3450
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cálc:"
         Height          =   225
         Index           =   7
         Left            =   0
         TabIndex        =   77
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursales:"
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   76
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label Label31 
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
         Height          =   495
         Left            =   1260
         TabIndex        =   75
         Top             =   1680
         Width           =   4395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes:"
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   74
         Top             =   2310
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stock"
      Height          =   750
      Left            =   10920
      Picture         =   "ANESTSUC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
      Top             =   2790
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kardex"
      Height          =   750
      Left            =   10920
      Picture         =   "ANESTSUC.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
      Top             =   2000
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valores Sugeridos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   0
      TabIndex        =   39
      Top             =   2520
      Width           =   4635
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   79
         ToolTipText     =   "Doble Click para Registrar Valor como Lote Económico"
         Top             =   525
         Width           =   990
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   62
         Top             =   525
         Width           =   990
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L.Repos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   2205
         TabIndex        =   80
         Top             =   555
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "St.Minimoo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   315
         TabIndex        =   42
         Top             =   560
         Width           =   855
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1155
         TabIndex        =   41
         Top             =   2295
         Width           =   750
      End
      Begin VB.Label Label13 
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
         Left            =   1260
         TabIndex        =   40
         Top             =   2535
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3675
      TabIndex        =   24
      Top             =   315
      Width           =   960
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
         Height          =   285
         Left            =   315
         TabIndex        =   25
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   105
      TabIndex        =   7
      Top             =   1260
      Width           =   4530
      Begin VB.Label MECONSUM 
         Caption         =   "MECONSUM"
         Height          =   225
         Left            =   3150
         TabIndex        =   64
         Top             =   3780
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Venta"
         Height          =   285
         Index           =   4
         Left            =   420
         TabIndex        =   28
         Top             =   3405
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fábrica"
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   27
         Top             =   3405
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   285
         Index           =   4
         Left            =   1890
         TabIndex        =   26
         Top             =   3405
         Width           =   1095
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2835
         TabIndex        =   23
         Top             =   380
         Width           =   960
      End
      Begin VB.Label Label5 
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
         Index           =   3
         Left            =   1995
         TabIndex        =   22
         Top             =   4770
         Width           =   1095
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
         Height          =   285
         Index           =   3
         Left            =   945
         TabIndex        =   21
         Top             =   4770
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Index           =   3
         Left            =   3255
         TabIndex        =   20
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Prom:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   19
         Top             =   380
         Width           =   855
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Máx:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   380
         Width           =   750
      End
      Begin VB.Label Label5 
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
         Index           =   2
         Left            =   1995
         TabIndex        =   17
         Top             =   4200
         Width           =   1095
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
         Height          =   285
         Index           =   2
         Left            =   945
         TabIndex        =   16
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   420
         TabIndex        =   15
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Index           =   1
         Left            =   1995
         TabIndex        =   14
         Top             =   3915
         Width           =   1095
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
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   13
         Top             =   3915
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   1890
         TabIndex        =   12
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   1995
         TabIndex        =   11
         Top             =   3630
         Width           =   1095
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
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   10
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   525
         TabIndex        =   9
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mínimo:"
         Height          =   225
         Index           =   0
         Left            =   -315
         TabIndex        =   8
         Top             =   3690
         Width           =   750
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   750
      Left            =   10920
      Picture         =   "ANESTSUC.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime Ficha"
      Top             =   1210
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   750
      Left            =   10920
      Picture         =   "ANESTSUC.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   1
      Top             =   315
      Width           =   3480
      Begin VB.CommandButton Command5 
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
         Height          =   280
         Left            =   3150
         TabIndex        =   2
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label MESAN 
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
         Left            =   1155
         TabIndex        =   3
         Top             =   420
         Width           =   2040
      End
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4425
      Left            =   105
      OleObjectBlob   =   "ANESTSUC.frx":0DC8
      TabIndex        =   0
      Top             =   3675
      Width           =   11700
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Disponibilidad de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1540
      Left            =   735
      TabIndex        =   44
      Top             =   3885
      Width           =   4635
      Begin VB.Line Line5 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   4620
         Y1              =   980
         Y2              =   980
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reservas:"
         Height          =   225
         Left            =   2310
         TabIndex        =   61
         Top             =   1120
         Width           =   1065
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos:"
         Height          =   225
         Left            =   210
         TabIndex        =   60
         Top             =   1120
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
         Height          =   225
         Left            =   2310
         TabIndex        =   59
         Top             =   695
         Width           =   1065
      End
      Begin VB.Label Label25 
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
         Left            =   3465
         TabIndex        =   58
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label24 
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
         Left            =   1155
         TabIndex        =   57
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label23 
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
         Left            =   3465
         TabIndex        =   56
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label22 
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
         Left            =   1155
         TabIndex        =   55
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label20 
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
         Left            =   3465
         TabIndex        =   54
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label18 
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
         Left            =   1155
         TabIndex        =   53
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compras:"
         Height          =   225
         Left            =   210
         TabIndex        =   51
         Top             =   695
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción:"
         Height          =   225
         Left            =   2310
         TabIndex        =   46
         Top             =   380
         Width           =   1065
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         Height          =   225
         Left            =   210
         TabIndex        =   45
         Top             =   380
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lead Time Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1550
      Left            =   4935
      TabIndex        =   29
      Top             =   3990
      Width           =   4635
      Begin VB.Label Label12 
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
         Left            =   1680
         TabIndex        =   38
         Top             =   1095
         Width           =   2850
      End
      Begin VB.Label Label11 
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
         Left            =   1680
         TabIndex        =   37
         Top             =   705
         Width           =   2850
      End
      Begin VB.Label Label10 
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
         Left            =   1680
         TabIndex        =   36
         Top             =   420
         Width           =   2850
      End
      Begin VB.Label Label9 
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
         Left            =   1050
         TabIndex        =   35
         Top             =   1090
         Width           =   570
      End
      Begin VB.Label Label8 
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
         Left            =   1050
         TabIndex        =   34
         Top             =   700
         Width           =   570
      End
      Begin VB.Label Label6 
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
         Left            =   1050
         TabIndex        =   33
         Top             =   420
         Width           =   570
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mínimo:"
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   32
         Top             =   480
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Máximo:"
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   31
         Top             =   765
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ultimo:"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   30
         Top             =   1155
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L.Econom."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1550
      Left            =   9240
      TabIndex        =   47
      Top             =   4200
      Width           =   1380
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   63
         ToolTipText     =   "Doble Click para Registrar Valor como Lote Económico"
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label Label16 
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
         Left            =   210
         TabIndex        =   50
         Top             =   535
         Width           =   990
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   105
         TabIndex        =   49
         Top             =   300
         Width           =   750
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sugerido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   105
         TabIndex        =   48
         Top             =   870
         Width           =   855
      End
   End
End
Attribute VB_Name = "ANESTSUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim CORRESP$, ABREMES$, CORRESQ$
Dim DESFE$, HASFE$
Dim CODIMES%
Dim CAMESA%
'VEPED$, VECON$
'

Private Sub Command1_Click()
   Command1.Enabled = False
   Clipboard.SetText TRIM$(MESAN)
   Call CONECRUN("CONSTO02", "Control y Movimientos de Stock")
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Printer.TrackDefault = False
   Printer.Orientation = 2
   PRINTFORM
   Printer.Orientation = 1
   Printer.TrackDefault = True
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX/" + MESAN, "Ficha Kardex de Movimientos")
      Else
        Call GENKARDEX(MESAN)
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Y$ = REGSEL$("STATSUCU")
   If Len(Y$) > 4 Then
     RESTAT$ = Y$
     MESAN = Left$(Y$, 16)
   End If
   Command5.Enabled = True
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   '
   Call CENTRAFORM(Me)
   '
   MSChart1.ColumnCount = 1
   CAMESA% = Val(FESTSUCU.Label12)
   MSChart1.RowCount = CAMESA%
   '
   For UJ% = 1 To CAMESA%
     MSChart1.Row = UJ%
     For UK% = 1 To 1
       MSChart1.Column = UK%
       MSChart1.Data = 0
     Next UK%
   Next UJ%
   '
   Call PREMESES
   '
   Label3(5) = FESTSUCU.Label3
   label1(11) = FESTSUCU.label1
   Label15 = FESTSUCU.Label9
   Label14 = FESTSUCU.Label12
   Label31 = FESTSUCU.Label11
   '
   List1.Clear
   For U& = 1 To FESTSUCU.List1.ListCount
     List1.AddItem Left$(FESTSUCU.List1.List(U& - 1), 48)
   Next U&
   '
   Refresh
   DoEvents
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub MESAN_Change()
   CIXI% = CODINT%(MESAN)
   If CIXI% > 0 Then
     MSChart1.Title = TRIM$(MESAN) + " - " + DESCRIT0$(CIXI%)
     Label7 = UNIMED$(CIXI%)
     '
     Call GENGRAFIC
     '
     VEPED$ = Mid$(RESTAT$, 129, 96)
     VECON$ = Mid$(RESTAT$, 257, 96)
     RVEN$ = Mid$(RESTAT$, 225, 16)
     RFAB$ = Mid$(RESTAT$, 353, 16)
     RTOT$ = Mid$(RESTAT$, 65, 16)
     '

     Label3(3) = Format(CVS(Mid$(RTOT$, 1, 4)), "####0.0")   ' MINIMO
     Label3(1) = Format(CVS(Mid$(RTOT$, 5, 4)), "####0.0")   ' MAXIMO
     Label3(2) = Format(CVS(Mid$(RTOT$, 9, 4)), "####0.0")   ' PROMEDIO
     '
     STOMIRE = 0
     If Val(Label3(2)) > 0 Then STOMIRE = 1
     If Val(Label3(2)) > 1 Then STOMIRE = 1 + Int(Val(Label3(2)) - 0.05)
     Text1 = AJUSTD$(Format(STOMIRE, "####0.0"), 7)
     '
     LOTEREP = STOMIRE
     If Val(Label3(1)) > LOTEREP Then LOTEREP = 1 + Int(Val(Label3(1)) - 0.05)
     Text3 = AJUSTD$(Format(LOTEREP, "####0.0"), 7)
     '
   End If
End Sub


Private Sub MESAN_DblClick()
   '
   Y$ = REGSEL$("STATSUCU")
   If Len(Y$) > 4 Then
     RESTAT$ = Y$
     MESAN = Left$(Y$, 16)
   End If
   '
End Sub

Private Sub MSChart1_DblClick()
   If MSChart1.Height < 7900 Then
      MSChart1.Top = 100
      MSChart1.Height = 4000
      MSChart1.Refresh
      MSChart2.Top = 4200
      MSChart2.Height = 2000
      MSChart2.Refresh
      MSChart3.Top = 6300
      MSChart3.Height = 2000
      MSChart3.Refresh
   End If
End Sub


Sub GENGRAFIC()

   'Call ARMAMES
   'If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   '
   Call PREMESES
   '
   VEPED$ = Mid$(RESTAT$, 129, 96)
   VECON$ = Mid$(RESTAT$, 257, 96)
   RVEN$ = Mid$(RESTAT$, 225, 16)
   RFAB$ = Mid$(RESTAT$, 353, 16)
   RTOT$ = Mid$(RESTAT$, 65, 16)
   '
   CAMESA% = Val(FESTSUCU.Label12)
'   If ATENTI% < 1 Then
'      COLMAX% = 2
'      MSChart1.ColumnCount = 1 + COLMAX%
      '
      MSChart1.RowCount = CAMESA%
      '
      For UJ% = 1 To CAMESA%
        MSChart1.Row = UJ%
        '
        MSChart1.Column = 1
        COVEN = CVS(Mid$(VEPED$, 4 * UJ% - 3, 4)) + CVS(Mid$(VECON$, 4 * UJ% - 3, 4))
        MSChart1.Data = COVEN
        '
'        MSChart1.Column = 2
'        COFAB = CVS(Mid$(VECON$, 4 * UJ% - 3, 4))
'        MSChart1.Data = COFAB
        '
'        MSChart1.Column = 3
'        MSChart1.Data = COVEN + COFAB
      Next UJ%
'   End If
   '
   MECONSU% = 0
   For UJ% = 1 To CAMESA%
     COVEN = CVS(Mid$(VEPED$, 4 * UJ% - 3, 4))
     COFAB = CVS(Mid$(VECON$, 4 * UJ% - 3, 4))
     If COVEN + COFAB > 0.5 Then MECONSU% = MECONSU% + 1
   Next UJ%
   MECONSUM = Str$(MECONSU%)
   '
   For UK% = 0 To 3
     Label3(UK%) = Format(CVS(Mid$(RVEN$, 1 + 4 * UK%, 4)), "#####0.0")
     Label4(UK%) = Format(CVS(Mid$(RFAB$, 1 + 4 * UK%, 4)), "#####0.0")
     Label5(UK%) = Format(CVS(Mid$(RTOT$, 1 + 4 * UK%, 4)), "#####0.0")
   Next UK%
   '
90 If ATENTI% < 1 Then Screen.MousePointer = 1
   '
End Sub

Sub ARMAMES()
   '
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(MESAN)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Call MENSERR(24, "Falta Indicar Período Mensual")
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   CODIMES% = 12 * ANIO% + MESO%
   '
End Sub


Sub PREMESES()
   '
   CAMESA% = Val(FESTSUCU.Label12)
   CORRESPO$ = FESTSUCU.Label9
   CORRESQ0$ = MESANO$(CORRESPO, DESFEX%, HASFEX%)
   FEX = HASFEX%
   FEXU$ = FECHATEX$(FEX)
   ANIO% = Val(Mid$(FEXU$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(FEXU$, 4, 2))
   CODIMES% = 12 * ANIO% + MESO%
   
   HACODI% = CODIMES%
   DECODI% = HACODI% - CAMESA% + 1
   '
   MSChart1.Repaint = True
   MSChart1.RowCount = CAMESA%
   For KKI% = DECODI% To HACODI%
     MSChart1.Row = 1 + KKI% - DECODI%
     MSChart1.RowLabel = ABREMESAN$(KKI%, 0)
   Next KKI%
   '
End Sub


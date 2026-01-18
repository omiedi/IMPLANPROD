VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ANESTK04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   3795
      Left            =   11025
      ScaleHeight     =   3735
      ScaleWidth      =   1005
      TabIndex        =   61
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   690
         Left            =   90
         Picture         =   "ANESTK04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   690
         Left            =   90
         Picture         =   "ANESTK04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Imprime Ficha"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Kardex"
         Height          =   690
         Left            =   90
         Picture         =   "ANESTK04.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stock"
         Height          =   690
         Left            =   90
         Picture         =   "ANESTK04.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2370
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ANESTK04.frx":0DC8
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3195
         Width           =   1515
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
      Left            =   9450
      TabIndex        =   44
      Top             =   1995
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
         TabIndex        =   59
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
         TabIndex        =   47
         Top             =   535
         Width           =   990
      End
      Begin VB.Label Label1 
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
         TabIndex        =   46
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label1 
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
         TabIndex        =   45
         Top             =   870
         Width           =   855
      End
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
      Left            =   4725
      TabIndex        =   41
      Top             =   1995
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
         TabIndex        =   57
         Top             =   1120
         Width           =   1065
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos:"
         Height          =   225
         Left            =   210
         TabIndex        =   56
         Top             =   1120
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
         Height          =   225
         Left            =   2310
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compras:"
         Height          =   225
         Left            =   210
         TabIndex        =   48
         Top             =   695
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción:"
         Height          =   225
         Left            =   2310
         TabIndex        =   43
         Top             =   380
         Width           =   1065
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         Height          =   225
         Left            =   210
         TabIndex        =   42
         Top             =   380
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "St.Mínimo"
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
      Left            =   9450
      TabIndex        =   37
      Top             =   210
      Width           =   1380
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
         Left            =   210
         TabIndex        =   58
         ToolTipText     =   "Doble Click para Registrar Valor como Stock Mínimo"
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   105
         TabIndex        =   40
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   105
         TabIndex        =   39
         Top             =   300
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
         Left            =   210
         TabIndex        =   38
         Top             =   535
         Width           =   990
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
      Left            =   4725
      TabIndex        =   27
      Top             =   210
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mínimo:"
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   30
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Máximo:"
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   29
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ultimo:"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   28
         Top             =   1155
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   22
      Top             =   210
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
         TabIndex        =   23
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   2280
      Left            =   105
      TabIndex        =   5
      Top             =   1260
      Width           =   4530
      Begin VB.Label MECONSUM 
         Caption         =   "MECONSUM"
         Height          =   225
         Left            =   3255
         TabIndex        =   60
         Top             =   210
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
         Left            =   1050
         TabIndex        =   26
         Top             =   460
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fábrica"
         Height          =   285
         Index           =   4
         Left            =   2100
         TabIndex        =   25
         Top             =   460
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   285
         Index           =   4
         Left            =   3150
         TabIndex        =   24
         Top             =   460
         Width           =   1095
      End
      Begin VB.Line Line4 
         X1              =   4305
         X2              =   4305
         Y1              =   1825
         Y2              =   2115
      End
      Begin VB.Line Line3 
         X1              =   1155
         X2              =   4305
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line2 
         X1              =   4305
         X2              =   4305
         Y1              =   695
         Y2              =   1540
      End
      Begin VB.Line Line1 
         X1              =   1155
         X2              =   4305
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Estimado:"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   21
         Top             =   1880
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
         Left            =   3255
         TabIndex        =   20
         Top             =   1825
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
         Left            =   2205
         TabIndex        =   19
         Top             =   1825
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
         Left            =   1155
         TabIndex        =   18
         Top             =   1825
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Promedio:"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Máximo:"
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   16
         Top             =   1035
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
         Left            =   3255
         TabIndex        =   15
         Top             =   1260
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
         Left            =   2205
         TabIndex        =   14
         Top             =   1260
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
         Left            =   1155
         TabIndex        =   13
         Top             =   1260
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
         Left            =   3255
         TabIndex        =   12
         Top             =   975
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
         Left            =   2205
         TabIndex        =   11
         Top             =   975
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
         Left            =   1155
         TabIndex        =   10
         Top             =   975
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
         Left            =   3255
         TabIndex        =   9
         Top             =   690
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
         Left            =   2205
         TabIndex        =   8
         Top             =   690
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
         Left            =   1155
         TabIndex        =   7
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mínimo:"
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   6
         Top             =   750
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
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
      Top             =   210
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
      OleObjectBlob   =   "ANESTK04.frx":2CEA
      TabIndex        =   0
      Top             =   3675
      Width           =   11700
   End
End
Attribute VB_Name = "ANESTK04"
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
   Call CONECRUN("CONSTO04", "Control y Movimientos de Stock")
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
   Y$ = REGSEL$("STATSTOK")
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
   MSChart1.ColumnCount = 3
   CAMESA% = Val(FESTOK04.Label12)
   MSChart1.RowCount = CAMESA%
   '
   For UJ% = 1 To CAMESA%
     MSChart1.Row = UJ%
     For UK% = 1 To 3
       MSChart1.Column = UK%
       MSChart1.Data = 0
     Next UK%
   Next UJ%
   '
   Call PREMESES
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
     Call LITACOM
     If ATENTI% < 1 Then Call MUEDIS(CIXI%)
     '
     Label13 = Format(STOMIN(CIXI%), "####0.0")
     Label16 = Format(LOTESTAN(CIXI%), "####0.0")
     '
     PLAREPO = (Val(Label6) + Val(Label8)) / 2
     If PLAREPO < Val(Label9) Then PLAREPO = Val(Label9)
     If TRIM$(Label9) = "" Then PLAREPO = PLAREPO + 1
     '
     STOMIRE = Int((1 + PLAREPO) * Val(Label5(3)) + 0.5)
     If Val(MECONSUM) <= 1 Then STOMIRE = 0
     Text1 = AJUSTD$(Format(STOMIRE, "####0.0"), 7)
     '
     LOTEREP = STOMIRE
     If COSUNI(CIXI%) > 0 Then
       COSUNRE = COSUNI(CIXI%) * TICAMONE(MONECOSTO%(CIXI%))
       If COSUNRE > 0 Then
         If LOTEREP * COSUNRE < 500 Then
           LOTEREP = 500 / COSUNRE
         End If
       End If
     End If
     If LOTEREP > 12 * Val(Label5(1)) Then
       LOTEREP = 12 * Val(Label5(1)) ' límite=consumo máximo de 12 meses
     End If
     '
     LOTEREP = Int(LOTEREP + 0.5)
     Text2 = AJUSTD$(Format(LOTEREP, "####0.0"), 7)
     '
   End If
End Sub

Sub LITACOM()
   '
   CIXI% = CODINT%(MESAN)
   RFF$ = RECFILT$("COPLAENT", 1, MKI$(CIXI%))
   Label6 = ""
   Label8 = ""
   Label9 = ""
   Label10 = ""
   Label11 = ""
   Label12 = ""
   If Len(RFF$) >= 4 Then
       LTMIN = 999999: NPRMIN% = 0
       LTMAX = 0: NPRMAX% = 0
       For U& = 1 To Len(RFF$) Step 4
         RECO& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("COPLAENT", RECO&)
         PLAEN = CVS(Mid$(XX$, 17, 4))
         NPRX% = CVI(Mid$(XX$, 9, 2))
         If PLAEN < LTMIN Then LTMIN = PLAEN: NPRMIN% = NPRX%
         If PLAEN > LTMAX Then LTMAX = PLAEN: NPRMAX% = NPRX%
       Next U&
       Label6 = Format(LTMIN, "##0.0")
       Label8 = Format(LTMAX, "##0.0")
       Label9 = Format(PLAEN, "##0.0")
       If NPRMIN% > 0 Then Label10 = RASOCLI$((-1) * NPRMIN%)
       If NPRMAX% > 0 Then Label11 = RASOCLI$((-1) * NPRMAX%)
       If NPRX% > 0 Then Label12 = RASOCLI$((-1) * NPRX%)
   End If
   '
End Sub
Private Sub MESAN_DblClick()
   '
   Y$ = REGSEL$("STATSTOK")
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
   CAMESA% = Val(FESTOK04.Label12)
   If ATENTI% < 1 Then
      COLMAX% = 2
      MSChart1.ColumnCount = 1 + COLMAX%
      '
      MSChart1.RowCount = CAMESA%
      '
      For UJ% = 1 To CAMESA%
        MSChart1.Row = UJ%
        '
        MSChart1.Column = 1
        COVEN = CVS(Mid$(VEPED$, 4 * UJ% - 3, 4))
        MSChart1.Data = COVEN
        '
        MSChart1.Column = 2
        COFAB = CVS(Mid$(VECON$, 4 * UJ% - 3, 4))
        MSChart1.Data = COFAB
        '
        MSChart1.Column = 3
        MSChart1.Data = COVEN + COFAB
      Next UJ%
   End If
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
   CAMESA% = Val(FESTOK04.Label12)
   CORRESPO$ = FESTOK04.Label9
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

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$
     '
     STOPROP = 0
     TORECEP = 0
     TOPEDID = 0
     TOSUMIN = 0
     TOFABRI = 0
     TORESER = 0
     GoSub PRESEDIS
     '
     Screen.MousePointer = 11
     Call CARDEPOS
     Call VERISAL(CI%)
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     REG& = CI% + 1
     If REG& > 0 Then
       If REG& <= NUMAS% + 1 Then
         SALCA$ = REGLEIDO$("SALSTOCK", REG&) + Mid$(REGLEIDO$("SALDEPOS", REG&), 41)
         For U& = 1 To NUSU%
           XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
           If Abs(CVS(XXZ$)) >= 0.05 Then
             If DISTO%(U&) >= 0 Then
               STOPROP = STOPROP + CVS(XXZ$)
             End If
           End If
         Next U&
         GoSub PRESEDIS
       End If
     End If
     '
12   TORECEP = 0
     RFF$ = RECFILT$("BOLREPEN", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("BOLREPEN", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("BOLREPEN", 1)
           Call GRAREG("BOLREPEN", X$, 1)
           Call CIERRARCH("BOLREPEN")
           GoTo 12
        End If
        '
        CARECEP = CVD(Mid$(X$, 89, 8))
        If CARECEP > 0 Then
           TORECEP = TORECEP + CARECEP
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
     ABUSQUE$ = "PEDDETP"
14   UREPE& = ULTREG&(ABUSQUE$)
     RGXX$ = RECFILT$(ABUSQUE$, 12, MKI$(CI%))
     For U& = 1 To Len(RGXX$) Step 4
       REG& = CVS(Mid$(RGXX$, U&, 4))
       If REG& > 0 Then
         If REG& <= UREPE& Then
           X$ = REGLEIDO$(ABUSQUE$, REG&)
           '
           If ABUSQUE$ = "PEDDETP" Then
             REGPEDA& = CVS(Mid$(X$, 109, 4))
             If Left$(REGLEIDO$("PEDDETA", REGPEDA&), 108) <> Left$(X$, 108) Then
               ABUSQUE$ = "PEDDETA"
               GoTo 14
             End If
           End If
           '
           If Asc(Mid$(X$, 27, 1)) <= 2 Then
             If Mid$(X$, 51, 1) <> "R" Then ' no es reparacion
               NroPed& = CVS(Left$(X$, 4))
               NCI% = CVI(Mid$(X$, 7, 2))
               CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
               CAPEN = CAPEN - CVS(Mid$(X$, 75, 4)) 'descuenta asignados
               FEX = CVI(Mid$(X$, 25, 2))
               If CAPEN < 0 Then CAPEN = 0
               '
               If CAPEN > 0 Then
                 TOPEDID = TOPEDID + CAPEN
               End If
             End If
           End If
         End If
       End If
     Next U&
     '
     GoSub PRESEDIS
     '
18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        CIIX% = CVI(Mid$(X$, 9, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("OCOMDETA", 1)
           Call GRAREG("OCOMDETA", X$, 1)
           Call CIERRARCH("OCOMDETA")
           GoTo 18
        End If
        STAT% = Asc(Mid$(X$, 124, 1))
        If STAT% < 8 Then      ' no esta anulada NI CUMPLIDA
          NUORSE& = CVS(Mid$(X$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(X$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              TOSUMIN = TOSUMIN + CAPEN
            End If
          End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
25   RFF$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("ORDEFABR", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("ORDEFABR", 1)
           Call GRAREG("ORDEFABR", X$, 1)
           Call CIERRARCH("ORDEFABR")
           GoTo 25
        End If
        STAT% = Asc(Mid$(X$, 108, 1))
        If STAT% <= 2 Then
           CACOM = CVD(Mid$(X$, 89, 8))
           CAREC = CVD(Mid$(X$, 119, 8))
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
              TOFABRI = TOFABRI + CAPEN
           End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
35   RFF$ = RECFILT$("RESERVA", 1, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("RESERVA", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("RESERVA", 1)
           Call GRAREG("RESERVA", X$, 1)
           Call CIERRARCH("RESERVA")
           GoTo 35
        End If
        TORE# = CVD(Mid$(X$, 19, 8))
        CACO# = CVD(Mid$(X$, 27, 8))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
           TORE# = TORE# / 1000
           CACO# = CACO# / 1000
        End If
        '
        SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
        SARE = TORE# - CACO#
        If SARE * SIGNO% < 0 Then
          SARE = 0
        End If
        CAPEN = SARE
        If Abs(CAPEN) > 0 Then
           TORESER = TORESER + CAPEN
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
135  RFF$ = RECFILT$("REPUREPA", 3, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("REPUREPA", U&)
        TORE# = CVS(Mid$(X$, 21, 4))
        CACO# = CVS(Mid$(X$, 25, 4))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
           TORE# = TORE# / 1000
           CACO# = CACO# / 1000
        End If
        '
        SARE = TORE# - CACO#
        If SARE < 0 Then SARE = 0
        CAPEN = SARE
        If Abs(CAPEN) > 0 Then
           TORESER = TORESER + CAPEN
        End If
139  Next UK&
     '
     GoSub PRESEDIS
     '
     Screen.MousePointer = 1
     Exit Sub
     '
PRESEDIS:
     '
     Label18 = TRIM$(CSTRING$(MKS$(STOPROP), 4, 12, 1, 2))
     Label20 = TRIM$(CSTRING$(MKS$(TORECEP), 4, 12, 1, 2))
     '
     Label24 = TRIM$(CSTRING$(MKS$(TOPEDID), 4, 12, 1, 2))
     Label25 = TRIM$(CSTRING$(MKS$(TORESER), 4, 12, 1, 2))
     '
     Label22 = TRIM$(CSTRING$(MKS$(TOSUMIN), 4, 12, 1, 2))
     Label23 = TRIM$(CSTRING$(MKS$(TOFABRI), 4, 12, 1, 2))
     '
Return
'
End Sub

Private Sub Text1_DblClick()
   STMI = Val(Text1)
   If STMI > 30000 Then
     STMI = (-1) * Abs(Int(STMI / 100) + 0.5)
   End If
   STMI1% = STMI
   CIXI% = CODINT%(MESAN)
   REMAL& = 1 + CIXI%
   XX$ = REGLEIDO$("MASTER", REMAL&)
   Call REPLA(XX$, MKI$(STMI1%), 105, 2)
   Call GRAREG("MASTER", XX$, REMAL&)
   Label13 = Text1
End Sub

Private Sub Text2_DblClick()
   STMI = Val(Text2)
   If STMI > 30000 Then
     STMI = (-1) * Abs(Int(STMI / 100) + 0.5)
   End If
   STMI1% = STMI
   CIXI% = CODINT%(MESAN)
   REMAL& = 1 + CIXI%
   XX$ = REGLEIDO$("MASTER", REMAL&)
   Call REPLA(XX$, MKI$(STMI1%), 107, 2)
   Call GRAREG("MASTER", XX$, REMAL&)
   Label16 = Text2

End Sub

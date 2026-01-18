VERSION 5.00
Begin VB.Form PREFINAN 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto Financiero"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4620
      TabIndex        =   96
      Top             =   105
      Width           =   2535
      Begin VB.Label Label9 
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
         Height          =   300
         Left            =   210
         TabIndex        =   97
         Top             =   420
         Width           =   2115
      End
   End
   Begin VB.CommandButton Command4 
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
      Height          =   855
      Left            =   9975
      Picture         =   "PREFINAN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Listado del Presupuesto Mensual"
      Top             =   3255
      Width           =   855
   End
   Begin VB.CommandButton Command3 
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
      Height          =   855
      Left            =   9975
      Picture         =   "PREFINAN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime Ficha de Presupuesto Mensual"
      Top             =   2310
      Width           =   855
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
      Left            =   9030
      Picture         =   "PREFINAN.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Guarda Presupuesto Financiero Mensual"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calc."
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
      Left            =   9975
      Picture         =   "PREFINAN.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Re-Calcula Presupuesto Financiero Mensual"
      Top             =   1365
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Flujo de Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   315
      TabIndex        =   64
      Top             =   6615
      Width           =   9465
      Begin VB.Label CAFLOSEM 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   86
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label CAFLOACU 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   85
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label CAFLOACU 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   75
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label CAFLOSEM 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   74
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label CAFLOACU 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   73
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label CAFLOSEM 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   72
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label CAFLOACU 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   71
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label CAFLOSEM 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   70
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label CAFLOACU 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   69
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label CAFLOSEM 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   68
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Semana 1        Semana 2         Semana 3         Semana 4                   Total"
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
         TabIndex        =   67
         Top             =   315
         Width           =   6705
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cash-Flow Acumulado:"
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
         Index           =   4
         Left            =   105
         TabIndex        =   66
         Top             =   1155
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cash-Flow Semanal:"
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
         Left            =   105
         TabIndex        =   65
         Top             =   735
         Width           =   2010
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presupuesto de Egresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   315
      TabIndex        =   45
      Top             =   3780
      Width           =   9465
      Begin VB.TextBox PADIFER 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   10
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox PADIFER 
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
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox PADIFER 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   13
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox PADIFER 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   12
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox PADIFER 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   11
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox COMPRACO 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   14
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox COMPRACO 
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
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox COMPRACO 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   17
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox COMPRACO 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   16
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox COMPRACO 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   15
         Top             =   1260
         Width           =   1350
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pagos a Diferir:"
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
         Left            =   105
         TabIndex        =   93
         Top             =   1050
         Width           =   2010
      End
      Begin VB.Label TOGASVAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   91
         ToolTipText     =   "Doble Click para Actualizar Gastos Varios"
         Top             =   1890
         Width           =   1485
      End
      Begin VB.Label TOGASVAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   90
         ToolTipText     =   "Doble Click para Actualizar Gastos Varios"
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label TOGASVAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   89
         ToolTipText     =   "Doble Click para Actualizar Gastos Varios"
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label TOGASVAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   88
         ToolTipText     =   "Doble Click para Actualizar Gastos Varios"
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label DEUPROVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   83
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label PAVALDIF 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   82
         Top             =   1575
         Width           =   1485
      End
      Begin VB.Label TOTEGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   81
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Est.Compras Contado:"
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
         Index           =   5
         Left            =   105
         TabIndex        =   76
         Top             =   1365
         Width           =   2010
      End
      Begin VB.Label TOTEGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   63
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label PAVALDIF 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   62
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUPROVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   61
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTEGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   60
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label PAVALDIF 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   59
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUPROVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   58
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTEGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   57
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label PAVALDIF 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   56
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUPROVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   55
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTEGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   54
         Top             =   2310
         Width           =   1350
      End
      Begin VB.Label TOGASVAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   53
         ToolTipText     =   "Doble Click para Actualizar Gastos Varios"
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label PAVALDIF 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   52
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUPROVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   51
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Deuda c/ Proveedores:"
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
         Left            =   105
         TabIndex        =   50
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pago Valores Diferidos:"
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
         Left            =   105
         TabIndex        =   49
         Top             =   1680
         Width           =   2010
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Egresos:"
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
         Left            =   105
         TabIndex        =   48
         Top             =   2415
         Width           =   2010
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Gastos Varios:"
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
         Left            =   105
         TabIndex        =   47
         Top             =   1995
         Width           =   2010
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Semana 1        Semana 2         Semana 3         Semana 4                   Total"
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
         TabIndex        =   46
         Top             =   315
         Width           =   6705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presupuesto de Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   315
      TabIndex        =   27
      Top             =   1260
      Width           =   9465
      Begin VB.TextBox INCOBRA 
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
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox INCOBRA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   5
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox INCOBRA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   4
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox INCOBRA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   3
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox INCOBRA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   2
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox INGREVAR 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   6
         Text            =   " "
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox INGREVAR 
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
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox INGREVAR 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   9
         Text            =   " "
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox INGREVAR 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   8
         Text            =   " "
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox INGREVAR 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   7
         Text            =   " "
         Top             =   1260
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Estimado Incobrables:"
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
         Left            =   105
         TabIndex        =   92
         Top             =   1050
         Width           =   2010
      End
      Begin VB.Label DEUDOVEN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   80
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label CHETERCE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   79
         Top             =   1575
         Width           =   1485
      End
      Begin VB.Label TOTINGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7770
         TabIndex        =   78
         Top             =   1995
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Semana 1        Semana 2         Semana 3         Semana 4                   Total"
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
         TabIndex        =   44
         Top             =   315
         Width           =   6705
      End
      Begin VB.Label TOTINGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   43
         Top             =   1995
         Width           =   1350
      End
      Begin VB.Label CHETERCE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   42
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUDOVEN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6300
         TabIndex        =   41
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTINGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   40
         Top             =   1995
         Width           =   1350
      End
      Begin VB.Label CHETERCE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   39
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUDOVEN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   4935
         TabIndex        =   38
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTINGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   37
         Top             =   1995
         Width           =   1350
      End
      Begin VB.Label CHETERCE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   36
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label DEUDOVEN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   35
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label TOTINGRE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   34
         Top             =   1995
         Width           =   1350
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Est.Ingresos Contado:"
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
         Left            =   105
         TabIndex        =   33
         Top             =   1365
         Width           =   2010
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Ingresos:"
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
         Left            =   105
         TabIndex        =   32
         Top             =   2100
         Width           =   2010
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cheques de Terceros:"
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
         Left            =   105
         TabIndex        =   31
         Top             =   1680
         Width           =   2010
      End
      Begin VB.Label CHETERCE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   30
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Deudores por Ventas:"
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
         Left            =   105
         TabIndex        =   29
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label DEUDOVEN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2205
         TabIndex        =   28
         Top             =   630
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Período Mensual Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   315
      TabIndex        =   25
      Top             =   105
      Width           =   4215
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
         TabIndex        =   0
         Top             =   420
         Width           =   175
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
         Height          =   300
         Left            =   1575
         TabIndex        =   87
         Top             =   420
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
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
         Left            =   -735
         TabIndex        =   26
         Top             =   525
         Width           =   2220
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   870
      Left            =   7770
      ScaleHeight     =   810
      ScaleWidth      =   60
      TabIndex        =   23
      Top             =   210
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command24 
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
      Left            =   9975
      Picture         =   "PREFINAN.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Left            =   8085
      Picture         =   "PREFINAN.frx":10D2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "PREFINAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
Dim CODIPERIO%
'

Private Sub CAFLOSEM_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(CAFLOSEM(KKI%).Caption)
     Next KKI%
     CAFLOSEM(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     CAFLOSEM(5).Refresh
     '
     For KKI% = 1 To 4
       SUTOIN# = SUTOIN# + Val(CAFLOSEM(KKI%).Caption)
       CAFLOACU(KKI%).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
     Next KKI%
   End If
End Sub

Private Sub CHETERCE_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(CHETERCE(KKI%).Caption)
     Next KKI%
     CHETERCE(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     CHETERCE(5).Refresh
     '
     SUTOIN# = Val(INGREVAR(Index).Text) - Val(INCOBRA(Index).Text) + Val(DEUDOVEN(Index).Caption) + Val(CHETERCE(Index).Caption)
     TOTINGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       GoTo 99
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       GoTo 99
    End If
    '
    If DES% <= HOY(HOS$) Then
       TTXX$ = "Ese Período Mensual ya está Iniciado !!!\Realmente Desea Re-Calcularlo ?"
       If COMALTER%(TTXX$ + "\\No, Cancelar Cálculo\Si, Re-Calcular") <> 2 Then GoTo 99
    End If
    '
    Screen.MousePointer = 11
    For KKI% = 1 To 4
      DEUDOVEN(KKI%).Caption = ""
      CHETERCE(KKI%).Caption = ""
      DEUPROVE(KKI%).Caption = ""
      PAVALDIF(KKI%).Caption = ""
      TOGASVAR(KKI%).Caption = ""
    Next KKI%
    Call CALDEUVEN
    Call CALCHETER
    Call CALDEUPRO
    Call CALCHEMI
    Call ACPREGASTOS(0)
    '
    FEX = HOY(HOS$)
    Call GRACONTROL("PRFI" + TRIM$(Str$(CODIPERIO%)), "FECALPRE", HOS$ + " - " + Left$(Time$, 5) + " hs")
    Label9.Caption = HOS$ + " - " + Left$(Time$, 5) + " hs"
    '
    Screen.MousePointer = 1
99  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Dim REPRESU$(16)
    Command2.Enabled = False
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    CODIPERIO% = 100 * Val(Mid$(DESX$, 7, 2)) + Val(Mid$(DESX$, 4, 2))
    '
    J& = 0
    For U& = 1 To ULTREG&("!PREGASTO")
      X$ = REGLEIDO$("!PREGASTO", U&)
      CODICON% = Abs(CVI(Mid$(X$, 15, 2)))
      If CODICON% < 10000 Then
        J& = J& + 1
        Call GRAREG("!PREGASTO", X$, J&)
      End If
    Next U&
    Call SETULTREG("!PREGASTO", J&)
    Call CIERRARCH("!PREGASTO")
    '
    REBLA$ = REGBLAN$("PREGASTO")
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(10001)
    Mid$(REPRE$, 17, 48) = "Deudores por Ventas"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(DEUDOVEN(KKI%).Caption))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(10002)
    Mid$(REPRE$, 17, 48) = "Estimado Incobrables"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(INCOBRA(KKI%).Text))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(10003)
    Mid$(REPRE$, 17, 48) = "Estimado Ingresos Contado"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(INGREVAR(KKI%).Text))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(10004)
    Mid$(REPRE$, 17, 48) = "Cheques de Terceros"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(CHETERCE(KKI%).Caption))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(-10001)
    Mid$(REPRE$, 17, 48) = "Deuda con Proveedores"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(DEUPROVE(KKI%).Caption))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(-10002)
    Mid$(REPRE$, 17, 48) = "Pagos a Diferir"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(PADIFER(KKI%).Text))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(-10003)
    Mid$(REPRE$, 17, 48) = "Estimado Compras Contado"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(COMPRACO(KKI%).Text))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    REPRE$ = REBLA$
    Mid$(REPRE$, 15, 2) = MKI$(-10004)
    Mid$(REPRE$, 17, 48) = "Pago Valores Diferidos"
    For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(PAVALDIF(KKI%).Caption))
    Next KKI%
    Call REGAPP("!PREGASTO", REPRE$)
    '
    '
    RFF$ = RECFILT$("PREGASTO", 1, MKI$(CODIPERIO%))
AAA = ULTREG&("!PREGASTO")
BBB = Len(RFF$) / 4
    PRL& = (-3)
    For U& = 1 To ULTREG&("!PREGASTO")
      X$ = REGLEIDO$("!PREGASTO", U&)
      Call REPLA(X$, MKI$(CODIPERIO%), 1, 2)
30    PRL& = PRL& + 4
      If PRL& <= Len(RFF$) - 3 Then
          REGU& = CVS(Mid$(RFF$, PRL&, 4))
          Call GRAREG("PREGASTO", X$, REGU&)
        Else
AAA = CVI(Mid$(X$, 15, 2))
          Call REGAPP("PREGASTO", X$)
      End If
    Next U&
    '
    While PRL& < Len(RFF$) - 3
      PRL& = PRL& + 4
      REGU& = CVS(Mid$(RFF$, PRL&, 4))
      Call GRAREG("PREGASTO", String$(128, 0), REGU&)
    Wend
    '
    Call CIERRARCH("PREGASTO")
    Call CIERRARCH("!PREGASTO")
    Screen.MousePointer = 1
    '
99  Command2.Enabled = True
    'Call FINAL("")
End Sub

Private Sub Command10_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.Caption = MESASEL$
        Call CARGAPRESU
        Call FRESHALL
    End If
    Call FRESHALL
End Sub

Private Sub Command24_Click()
    Call FINAL("")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("PRESUFIN")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   PrintForm
99 Command3.Enabled = True
End Sub


Private Sub Command8_Click()
   Call REMENGAS
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
       GoTo 99
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
       GoTo 99
    End If
    '
    Call LIPREFIN
    '
99  Command4.Enabled = True
End Sub

Private Sub COMPRACO_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(COMPRACO(KKI%).Text)
     Next KKI%
     COMPRACO(5).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     COMPRACO(5).Refresh
     '
     SUTOIN# = Val(COMPRACO(Index).Text) - Val(PADIFER(Index).Text) + Val(DEUPROVE(Index).Caption) + Val(PAVALDIF(Index).Caption) + Val(TOGASVAR(Index).Caption)
     TOTEGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub COMPRACO_GotFocus(Index As Integer)
    '
    If Index > 4 Then
       COMPRACO(4).SetFocus
       Exit Sub
    End If
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       On Error Resume Next
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error GoTo 0
       Command10.SetFocus
       Exit Sub
    End If
    '
End Sub

Private Sub COMPRACO_LostFocus(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = Val(COMPRACO(Index).Text)
     COMPRACO(Index).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 10, 2, 2)
     COMPRACO(Index).Refresh
   End If
End Sub

Private Sub DEUDOVEN_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(DEUDOVEN(KKI%).Caption)
     Next KKI%
     DEUDOVEN(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     DEUDOVEN(5).Refresh
     '
     SUTOIN# = Val(INGREVAR(Index).Text) - Val(INCOBRA(Index).Text) + Val(DEUDOVEN(Index).Caption) + Val(CHETERCE(Index).Caption)
     TOTINGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub DEUPROVE_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(DEUPROVE(KKI%).Caption)
     Next KKI%
     DEUPROVE(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     DEUPROVE(5).Refresh
     '
     SUTOIN# = Val(COMPRACO(Index).Text) - Val(PADIFER(Index).Text) + Val(DEUPROVE(Index).Caption) + Val(PAVALDIF(Index).Caption) + Val(TOGASVAR(Index).Caption)
     TOTEGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    TABLA$(0) = "TABTIMA"       ' tabla de tipos de material
    TABLA$(1) = "UMEDIDA"
    TABLA$(2) = "TADEPOSI"
    TABLA$(3) = ""       ' grupos contables por articulo
    TABLA$(4) = "*VENDEDOR"
    TABLA$(5) = "TAPROVIN"
    TABLA$(6) = "CONPAGO"
    TABLA$(7) = "ADMIPRE"
    TABLA$(8) = "CUEINGAS"
    TABLA$(9) = "*CUECOVE"
    '
End Sub

Private Sub INCOBRA_Change(Index As Integer)
   If Index < 5 Then
     SUINCOBRA# = 0
     For KKI% = 1 To 4
       SUINCOBRA# = SUINCOBRA# + Val(INCOBRA(KKI%).Text)
     Next KKI%
     INCOBRA(5).Text = CSTRING$(MKD$(SUINCOBRA#), 4, 11, 2, 2)
     INCOBRA(5).Refresh
     '
     SUTOIN# = Val(INGREVAR(Index).Text) - Val(INCOBRA(Index).Text) + Val(DEUDOVEN(Index).Caption) + Val(CHETERCE(Index).Caption)
     TOTINGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub INCOBRA_GotFocus(Index As Integer)
    '
    If Index > 4 Then
       INCOBRA(4).SetFocus
       Exit Sub
    End If
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
End Sub

Private Sub INCOBRA_LostFocus(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = Val(INCOBRA(Index).Text)
     INCOBRA(Index).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 10, 2, 2)
     INCOBRA(Index).Refresh
   End If
End Sub


Private Sub INGREVAR_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(INGREVAR(KKI%).Text)
     Next KKI%
     INGREVAR(5).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     INGREVAR(5).Refresh
     '
     SUTOIN# = Val(INGREVAR(Index).Text) - Val(INCOBRA(Index).Text) + Val(DEUDOVEN(Index).Caption) + Val(CHETERCE(Index).Caption)
     TOTINGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub INGREVAR_GotFocus(Index As Integer)
    '
    If Index > 4 Then
       INGREVAR(4).SetFocus
       Exit Sub
    End If
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
End Sub

Private Sub INGREVAR_LostFocus(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = Val(INGREVAR(Index).Text)
     INGREVAR(Index).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 10, 2, 2)
     INGREVAR(Index).Refresh
   End If
End Sub

Private Sub PADIFER_Change(Index As Integer)
   If Index < 5 Then
     SUPADIFER# = 0
     For KKI% = 1 To 4
       SUPADIFER# = SUPADIFER# + Val(PADIFER(KKI%).Text)
     Next KKI%
     PADIFER(5).Text = CSTRING$(MKD$(SUPADIFER#), 4, 11, 2, 2)
     PADIFER(5).Refresh
     '
     SUTOIN# = Val(COMPRACO(Index).Text) - Val(PADIFER(Index).Text) + Val(DEUPROVE(Index).Caption) + Val(PAVALDIF(Index).Caption) + Val(TOGASVAR(Index).Caption)
     TOTEGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub PADIFER_GotFocus(Index As Integer)
    '
    If Index > 4 Then
       PADIFER(4).SetFocus
       Exit Sub
    End If
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       On Error Resume Next
       Command10.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
End Sub

Private Sub PADIFER_LostFocus(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = Val(PADIFER(Index).Text)
     PADIFER(Index).Text = CSTRING$(MKD$(SUMAINGRE#), 4, 10, 2, 2)
     PADIFER(Index).Refresh
   End If
End Sub

Private Sub MESAN_DblClick()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.Caption = MESASEL$
        DoEvents
        Call CARGAPRESU
        Call FRESHALL
    End If
    Call FRESHALL
End Sub

Sub CARGAPRESU()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    CODIPERIO% = 100 * Val(Mid$(DESX$, 7, 2)) + Val(Mid$(DESX$, 4, 2))
    '
    Label9.Caption = CONTROL$("PRFI" + TRIM$(Str$(CODIPERIO%)), "FECALPRE")
    Label9.Refresh
    '
15  RFF$ = RECFILT$("PREGASTO", 1, MKI$(CODIPERIO%))
    If RFF$ = "" Then
       Call SETULTREG("!PREGASTO", 0)
       Call ACPREGASTOS(0)
       GoTo 90
    End If
    '
    RFF$ = RECFILT$("PREGASTO", 1, MKI$(CODIPERIO%))
    If RFF$ = "" Then
      Exit Sub
    End If
    '
    For KKI% = 1 To 4
      DEUDOVEN(KKI%).Caption = ""
      CHETERCE(KKI%).Caption = ""
      DEUPROVE(KKI%).Caption = ""
      PAVALDIF(KKI%).Caption = ""
      TOGASVAR(KKI%).Caption = ""
      INGREVAR(KKI%).Text = ""
      INCOBRA(KKI%).Text = ""
      PADIFER(KKI%).Text = ""
      COMPRACO(KKI%).Text = ""
    Next KKI%
    '
    J& = 0
    For KKU& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, KKU&, 4))
      If U& > 0 Then
        If U& <= ULTREG&("PREGASTO") Then
          X$ = REGLEIDO$("PREGASTO", U&)
          If CVI(Left$(X$, 2)) <> CODIPERIO% Then
            X$ = REGLEIDO$("PREGASTO", 1)
            Call GRAREG("PREGASTO", X$, 1)
            Call CIERRARCH("PREGASTO")
            GoTo 15
          End If
          CODICON% = CVI(Mid$(X$, 15, 2))
          If CODICON% = 10001 Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(DEUDOVEN(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 DEUDOVEN(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = 10002 Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(INCOBRA(KKI%).Text) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 INCOBRA(KKI%).Text = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = 10003 Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(INGREVAR(KKI%).Text) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 INGREVAR(KKI%).Text = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = 10004 Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(CHETERCE(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 CHETERCE(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = (-10001) Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(DEUPROVE(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 DEUPROVE(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = (-10002) Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(PADIFER(KKI%).Text) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 PADIFER(KKI%).Text = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = (-10003) Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(COMPRACO(KKI%).Text) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 COMPRACO(KKI%).Text = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            ElseIf CODICON% = (-10004) Then
               For KKI% = 1 To 4
                 SUMACOL# = Val(PAVALDIF(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 PAVALDIF(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
            Else
               J& = J& + 1
               Call GRAREG("!PREGASTO", X$, J&)
               For KKI% = 1 To 4
                 SUMACOL# = Val(TOGASVAR(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
                 TOGASVAR(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
               Next KKI%
          End If
        End If
      End If
    Next KKU&
    Call SETULTREG("!PREGASTO", J&)
    Call CIERRARCH("!PREGASTO")
    '
    If ULTREG&("!PREGASTO") < 1 Then
      Call ACPREGASTOS(0)
    End If
    '
90  Screen.MousePointer = 1
    '
End Sub
'
Sub ACPREGASTOS(MODO%)
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    CODIPERIO% = 100 * Val(Mid$(DESX$, 7, 2)) + Val(Mid$(DESX$, 4, 2))
    '
15  RFF$ = RECFILT$("PREGASTO", 1, MKI$(CODIPERIO%))
    If RFF$ = "" Then
       Call SETULTREG("!PREGASTO", 0)
       For KKI% = 1 To 4
         DEUDOVEN(KKI%).Caption = ""
         CHETERCE(KKI%).Caption = ""
         DEUPROVE(KKI%).Caption = ""
         PAVALDIF(KKI%).Caption = ""
         TOGASVAR(KKI%).Caption = ""
         INGREVAR(KKI%).Text = ""
         INCOBRA(KKI%).Text = ""
         PADIFER(KKI%).Text = ""
         COMPRACO(KKI%).Text = ""
       Next KKI%
       Call GENPREGAS
       Call CIERRARCH("PREGASTO")
       Call CIERRARCH("!PREGASTO")
       '
       Call CALDEUVEN
       Call CALCHETER
       Call CALDEUPRO
       Call CALCHEMI
       '
       FEX = HOY(HOS$)
       Call GRACONTROL("PRFI" + TRIM$(Str$(CODIPERIO%)), "FECALPRE", HOS$ + " - " + Left$(Time$, 5) + " hs")
       Label9.Caption = HOS$ + " - " + Left$(Time$, 5) + " hs"
       '
       Screen.MousePointer = 1
       GoTo 25
    End If
    '
    J& = 0
    For KKU& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, KKU&, 4))
      If U& > 0 Then
        If U& <= ULTREG&("PREGASTO") Then
          X$ = REGLEIDO$("PREGASTO", U&)
          If CVI(Left$(X$, 2)) <> CODIPERIO% Then
            X$ = REGLEIDO$("PREGASTO", 1)
            Call GRAREG("PREGASTO", X$, 1)
            Call CIERRARCH("PREGASTO")
            GoTo 15
          End If
          rgg% = Abs(CVI(Mid$(X$, 15, 2)))
          If rgg% < 10000 Then
            J& = J& + 1
            Call GRAREG("!PREGASTO", X$, J&)
          End If
        End If
      End If
    Next KKU&
    Call SETULTREG("!PREGASTO", J&)
    '
20  Call CIERRARCH("PREGASTO")
    Call CIERRARCH("!PREGASTO")
    '
    If ULTREG&("!PREGASTO") < 1 Then
       Call GENPREGAS
       GoTo 20
    End If
    '
25  If MODO% = 1 Then
      VTB% = VENTABU%("CAPREGAS/TITUPAN=Actualizacion Presupuesto de Gastos - " + MESASEL$)
      If VTB% < 0 Then Exit Sub
    End If
    '
    For KKI% = 1 To 4
      TOGASVAR(KKI%).Caption = ""
    Next KKI%
    '
    RFF$ = RECFILT$("PREGASTO", 1, MKI$(CODIPERIO%))
    PRL& = (-3)
    For U& = 1 To ULTREG&("!PREGASTO")
      X$ = REGLEIDO$("!PREGASTO", U&)
      Call REPLA(X$, MKI$(CODIPERIO%), 1, 2)
30    PRL& = PRL& + 4
      If PRL& <= Len(RFF$) - 3 Then
          REGU& = CVS(Mid$(RFF$, PRL&, 4))
          Y$ = REGLEIDO$("PREGASTO", REGU&)
          If Abs(CVI(Mid$(Y$, 15, 2))) > 10000 Then GoTo 30
          Call GRAREG("PREGASTO", X$, REGU&)
        Else
          Call REGAPP("PREGASTO", X$)
      End If
      For KKI% = 1 To 4
        SUMACOL# = Val(TOGASVAR(KKI%).Caption) + CVD(Mid$(X$, 57 + 8 * KKI%, 8))
        TOGASVAR(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
      Next KKI%
    Next U&
    '
    While PRL& < Len(RFF$) - 3
      PRL& = PRL& + 4
      REGU& = CVS(Mid$(RFF$, PRL&, 4))
      Call GRAREG("PREGASTO", String$(128, 0), REGU&)
    Wend
    '
    Call CIERRARCH("PREGASTO")
    Call CIERRARCH("!PREGASTO")
    '
    If ULTREG&("!PREGASTO") < 1 Then
      Call GENPREGAS
      GoTo 20
    End If
    '
End Sub

Sub REMENGAS()
    '
    Call CARSELMESA
5   Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    If DES% > HOY(HOS$) Then
       Call MENSERR(24, "Período no Válido")
       GoTo 5
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    
'FEX = FECOR%
'ABC1$ = FECHATEX(FEX)
'ABC2$ = FECHATEX(FEX + 7)
'ABC3$ = FECHATEX(FEX + 14)
'ABC4$ = FECHATEX(FEX + 21)
    '
    JJ& = 0
    '
    Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    'Dim COCUEGA$(1024), IGAS#(1024, 5)
    Dim VEGASTO$(1024)
    '
    If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
    End If
    '
    CANOPFF% = 0
    FIN& = ULTREG&("CUECORR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If Mid$(CC0$, 7, 2) = "FF" Then
        If CVI(Left$(CC0$, 2)) >= DES% Then
          If CVI(Left$(CC0$, 2)) <= HAS% Then
            NOPA = Val(Mid$(CC0$, 20, 6))
            If NOPA <> NOPFF(CANOPFF%) Then
               CANOPFF% = CANOPFF% + 1
               NOPFF(CANOPFF%) = NOPA
            End If
          End If
        End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      FEOPA% = CVI(CH8$)
      If FEOPA% >= DES% Then                ' CORRESPONDE FECHA
        If FEOPA% <= HAS% Then
          IM# = CVD(CH5$)
          If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHEQUEMI", J&)
                  CH1$ = Mid$(CH0$, 1, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 41, 4)
                  If CVI(CH1$) = NP% Then
                     If CVS(CH6$) = NOPA Then
                        ITOT# = ITOT# + CVD(CH5$)
                     End If
                  End If
               Next J&
               '
               If TICUEPRO((-1) * NP%) = 1 Then ' CUENTA DE GASTOS
                  For UDP& = 1 To ULTREG&("DETPAGAS")
                     DTP$ = REGLEIDO$("DETPAGAS", UDP&)
                     If CVS(Mid$(DTP$, 3, 4)) = NOPA Then
                        CUEDP% = CVI(Mid$(DTP$, 17, 2))
                        If TRIM$(CODCUE$(CUEDP%)) <> "" Then
                           IDP# = CVD(Mid$(DTP$, 9, 8))
                           CUEGA$ = CODCUE$(CUEDP%)
                           FEGA% = FEOPA%
                           IMGA# = IDP#
                           GoSub 80
                           ITOT# = ITOT# - IDP#
                        End If
                     End If
                  Next UDP&
               End If
               '
               ' considera solo cuentas de gastos
               'If Abs(ITOT#) >= 0.005 Then
               '  CUEGA$ = CUEMADRE$((-1) * NP%)
               '  FEGA% = FEOPA%
               '  IMGA# = ITOT#
               '  GoSub 80
               'End If
               '
               NOPANT = NOPA
               NPANT% = NP%
            End If
          End If
        End If
      End If
76 Next IAX&
   '
   For JJ& = 1 To CACU%
     ZZ$ = REGBLAN$("RESGASTO")
     Call REPLA(ZZ$, Left$(VEGASTO$(JJ&), 12), 3, 12)
     SUMARE# = 0
     For COLUI% = 1 To 5
       Call REPLA(ZZ$, MKD$(Val(Mid$(VEGASTO$(JJ&), 12 * COLUI% + 1, 12))), 57 + 8 * COLUI%, 8)
       SUMARE# = SUMARE# + Val(Mid$(VEGASTO$(JJ&), 12 * COLUI% + 1, 12))
     Next COLUI%
     Call REPLA(ZZ$, MKD$(SUMARE#), 105, 8)
     Call GRAREG("RESGASTO", ZZ$, JJ&)
   Next JJ&
   '
   JJ& = CACU%
   Call SETULTREG("RESGASTO", JJ&)
   Call CIERRARCH("RESGASTO")
   Call HEADERS("RESGASTO", "")
   Call HEADERS("RESGASTO", "Corresponde a: " + PERIO$)
   '
   Screen.MousePointer = 1
   '
   Call FINAL("*LIREGAS")
   GoTo 99
   '
80 If FEGA% <= FECOR% Then
       COLUXI% = 1
     Else
       COLUXI% = 2 + Int((FEGA% - FECOR%) / 7)
   End If
   If COLUXI% > 4 Then COLUXI% = 4
   If COLUXI% > 0 Then
     If COLUXI% <= 4 Then
       CUEGA1$ = AJUSTI$(CUEGA$, 12)
       For KK% = 1 To CACU%
         If Left$(VEGASTO$(KK%), 12) = CUEGA1$ Then
           TTZ$ = VEGASTO$(KK%)
           IGASTO# = Val(Mid$(VEGASTO$(KK%), 12 * COLUXI% + 1, 12)) + IMGA#
           Call REPLA(TTZ$, CSTRING$(MKD$(IGASTO#), 4, 12, 2, 2), 12 * COLUXI% + 1, 12)
           VEGASTO$(KK%) = TTZ$
           GoTo 89
         End If
       Next KK%
       '
       TTZ$ = Space$(128)
       Call REPLA(TTZ$, CUEGA1$, 1, 12)
       Call REPLA(TTZ$, CSTRING$(MKD$(IMGA#), 4, 12, 2, 2), 12 * COLUXI% + 1, 12)
       '
       For KK% = CACU% To 1 Step -1
         If Left$(VEGASTO$(KK%), 12) <= CUEGA1$ Then
           VEGASTO$(KK% + 1) = TTZ$
           CACU% = CACU% + 1
           GoTo 89
         End If
         VEGASTO$(KK% + 1) = VEGASTO$(KK%)
       Next KK%
       VEGASTO$(1) = TTZ$
       CACU% = CACU% + 1
     End If
   End If
   '
89 Return
   '
   '
99 End Sub

       
Sub GENPREGAS()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Call MENSERR(24, "Seleccione Período Mensual")
       Command10.SetFocus
       Exit Sub
    End If
    '
    If COMALTER%("Generación Automatica Presupuesto Financiero\para " + TRIM$(MESASEL$) + ".\  \Seleccione el Período Mensual\para la Estimación Básica de Gastos.\\Continuar\Cancelar") <> 1 Then
       Call FINAL("")
    End If
    '
    Call CARSELMESA
5   Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    If DES% > HOY(HOS$) Then
       Call MENSERR(24, "Período no Válido")
       GoTo 5
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    '
    JJ& = 0
    '
    Dim NOPFF(2048)
    Dim VEGASTO$(1024)
    '
    CANOPFF% = 0
    FIN& = ULTREG&("CUECORR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If Mid$(CC0$, 7, 2) = "FF" Then
        If CVI(Left$(CC0$, 2)) >= DES% Then
          If CVI(Left$(CC0$, 2)) <= HAS% Then
            NOPA = Val(Mid$(CC0$, 20, 6))
            If NOPA <> NOPFF(CANOPFF%) Then
               CANOPFF% = CANOPFF% + 1
               NOPFF(CANOPFF%) = NOPA
            End If
          End If
        End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      FEOPA% = CVI(CH8$)
      If FEOPA% >= DES% Then                ' CORRESPONDE FECHA
        If FEOPA% <= HAS% Then
          IM# = CVD(CH5$)
          If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHEQUEMI", J&)
                  CH1$ = Mid$(CH0$, 1, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 41, 4)
                  If CVI(CH1$) = NP% Then
                     If CVS(CH6$) = NOPA Then
                        ITOT# = ITOT# + CVD(CH5$)
                     End If
                  End If
               Next J&
               '
               If TICUEPRO((-1) * NP%) = 1 Then ' CUENTA DE GASTOS
                  For UDP& = 1 To ULTREG&("DETPAGAS")
                     DTP$ = REGLEIDO$("DETPAGAS", UDP&)
                     If CVS(Mid$(DTP$, 3, 4)) = NOPA Then
                        CUEDP% = CVI(Mid$(DTP$, 17, 2))
                        If TRIM$(CODCUE$(CUEDP%)) <> "" Then
                           IDP# = CVD(Mid$(DTP$, 9, 8))
                           CUEGA$ = CODCUE$(CUEDP%)
                           FEGA% = FEOPA%
                           IMGA# = IDP#
                           GoSub 80
                           ITOT# = ITOT# - IDP#
                        End If
                     End If
                  Next UDP&
               End If
               '
               ' considera solo cuentas de gastos
               'If Abs(ITOT#) >= 0.005 Then
               '  CUEGA$ = CUEMADRE$((-1) * NP%)
               '  FEGA% = FEOPA%
               '  IMGA# = ITOT#
               '  GoSub 80
               'End If
               '
               NOPANT = NOPA
               NPANT% = NP%
            End If
          End If
        End If
      End If
76 Next IAX&
   '
   For JJ& = 1 To CACU%
     ZZ$ = REGBLAN$("PREGASTO")
     Call REPLA(ZZ$, Left$(VEGASTO$(JJ&), 12), 3, 12)
     SUMARE# = 0
     For COLUI% = 1 To 5
       Call REPLA(ZZ$, MKD$(Val(Mid$(VEGASTO$(JJ&), 12 * COLUI% + 1, 12))), 57 + 8 * COLUI%, 8)
       SUMARE# = SUMARE# + Val(Mid$(VEGASTO$(JJ&), 12 * COLUI% + 1, 12))
     Next COLUI%
     Call REPLA(ZZ$, MKD$(SUMARE#), 105, 8)
     Call GRAREG("!PREGASTO", ZZ$, JJ&)
   Next JJ&
   '
   JJ& = CACU%
   Call SETULTREG("!PREGASTO", JJ&)
   Call CIERRARCH("!PREGASTO")
   '
   Screen.MousePointer = 1
   '
   GoTo 99
   '
80 If FEGA% <= FECOR% Then
       COLUXI% = 1
     Else
       COLUXI% = 2 + Int((FEGA% - FECOR%) / 7)
   End If
   If COLUXI% > 4 Then COLUXI% = 4
   If COLUXI% > 0 Then
     If COLUXI% <= 4 Then
       CUEGA1$ = AJUSTI$(CUEGA$, 12)
       For KK% = 1 To CACU%
         If Left$(VEGASTO$(KK%), 12) = CUEGA1$ Then
           TTZ$ = VEGASTO$(KK%)
           IGASTO# = Val(Mid$(VEGASTO$(KK%), 12 * COLUXI% + 1, 12)) + IMGA#
           Call REPLA(TTZ$, CSTRING$(MKD$(IGASTO#), 4, 12, 2, 2), 12 * COLUXI% + 1, 12)
           VEGASTO$(KK%) = TTZ$
           GoTo 89
         End If
       Next KK%
       '
       TTZ$ = Space$(128)
       Call REPLA(TTZ$, CUEGA1$, 1, 12)
       Call REPLA(TTZ$, CSTRING$(MKD$(IMGA#), 4, 12, 2, 2), 12 * COLUXI% + 1, 12)
       '
       For KK% = CACU% To 1 Step -1
         If Left$(VEGASTO$(KK%), 12) <= CUEGA1$ Then
           VEGASTO$(KK% + 1) = TTZ$
           CACU% = CACU% + 1
           GoTo 89
         End If
         VEGASTO$(KK% + 1) = VEGASTO$(KK%)
       Next KK%
       VEGASTO$(1) = TTZ$
       CACU% = CACU% + 1
     End If
   End If
   '
89 Return
   '
99 End Sub

Private Sub PAVALDIF_CHANGE(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(PAVALDIF(KKI%).Caption)
     Next KKI%
     PAVALDIF(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     PAVALDIF(5).Refresh
     '
     SUTOIN# = Val(COMPRACO(Index).Text) - Val(PADIFER(Index).Text) + Val(DEUPROVE(Index).Caption) + Val(PAVALDIF(Index).Caption) + Val(TOGASVAR(Index).Caption)
     TOTEGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub TOGASVAR_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(TOGASVAR(KKI%).Caption)
     Next KKI%
     TOGASVAR(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     TOGASVAR(5).Refresh
     '
     SUTOIN# = Val(COMPRACO(Index).Text) - Val(PADIFER(Index).Text) + Val(DEUPROVE(Index).Caption) + Val(PAVALDIF(Index).Caption) + Val(TOGASVAR(Index).Caption)
     TOTEGRE(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub

Private Sub TOGASVAR_DblClick(Index As Integer)
   Call ACPREGASTOS(1)
End Sub

Private Sub TOTEGRE_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(TOTEGRE(KKI%).Caption)
     Next KKI%
     TOTEGRE(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     TOTEGRE(5).Refresh
     '
     SUTOIN# = Val(TOTINGRE(Index).Caption) - Val(TOTEGRE(Index).Caption)
     CAFLOSEM(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If

End Sub

Private Sub TOTINGRE_Change(Index As Integer)
   If Index < 5 Then
     SUMAINGRE# = 0
     For KKI% = 1 To 4
       SUMAINGRE# = SUMAINGRE# + Val(TOTINGRE(KKI%).Caption)
     Next KKI%
     TOTINGRE(5).Caption = CSTRING$(MKD$(SUMAINGRE#), 4, 11, 2, 2)
     TOTINGRE(5).Refresh
     '
     SUTOIN# = Val(TOTINGRE(Index).Caption) - Val(TOTEGRE(Index).Caption)
     CAFLOSEM(Index).Caption = CSTRING$(MKD$(SUTOIN#), 4, 10, 2, 2)
   End If
End Sub


Sub CALDEUVEN()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    'If DES% > HOY(HOS$) Then
    '   Call MENSERR(24, "Período no Válido")
    '   Exit Sub
    'End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    '
    FIN& = ULTREG&("SVD202")
    For IL& = 1 To FIN&
        SS$ = REGLEIDO$("SVD202", IL&)
        NCI% = CVI(Left$(SS$, 2))
        TI% = Asc(Mid$(SS$, 3, 1))
        If TI% > 7 Then TI% = 2
        FEGA% = CVI(Mid$(SS$, 9, 2))
        If FEGA% <= HAS% Then
          SALDEU# = CVD(Mid$(SS$, 11, 8))
          If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
          GoSub 80
        End If
809 Next IL&
    Call CIERRARCH("*.*")
    Exit Sub
    '
80 If FEGA% <= FECOR% Then
       KKI% = 1
     Else
       KKI% = 2 + Int((FEGA% - FECOR%) / 7)
   End If
   If KKI% > 4 Then KKI% = 4
   If KKI% > 0 Then
     If KKI% <= 4 Then
       SUMACOL# = Val(DEUDOVEN(KKI%).Caption) + SALDEU#
       DEUDOVEN(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
       DEUDOVEN(KKI%).Refresh
     End If
   End If
   '
89 Return
   '
End Sub

Sub CALCHETER()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    'If DES% > HOY(HOS$) Then
    '   Call MENSERR(24, "Período no Válido")
    '   Exit Sub
    'End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    '
    FIN& = ULTREG&("BADACHE")
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
        FEGA% = CVI(Mid$(X$, 59, 2))
        If FEGA% <= HAS% Then
          SALDEU# = CVD(Mid$(X$, 49, 8))
          GoSub 80
        End If
      End If
19 Next I&
   '
   Call CIERRARCH("*.*")
   Exit Sub
   '
80 If FEGA% <= FECOR% Then
       KKI% = 1
     Else
       KKI% = 2 + Int((FEGA% - FECOR%) / 7)
   End If
   If KKI% > 4 Then KKI% = 4
   If KKI% > 0 Then
     If KKI% <= 4 Then
       SUMACOL# = Val(CHETERCE(KKI%).Caption) + SALDEU#
       CHETERCE(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
       CHETERCE(KKI%).Refresh
     End If
   End If
   '
89 Return
   '
End Sub

Sub CALCHEMI()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    '
    FIN& = ULTREG&("CHEQUEMI")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("CHEQUEMI", I&)
       PROV% = CVI(Left$(X$, 2))
       NUMERO$ = Mid$(X$, 3, 12)
       BANCO$ = Mid$(X$, 15, 16)
       EMIS% = CVI(Mid$(X$, 31, 2))
       IMPO# = CVD(Mid$(X$, 33, 8))
       VENCI% = CVI(Mid$(X$, 49, 2))
       '
       If CVI(Mid$(X$, 45, 2)) <> 999 Then ' excluye cheques terceros
        If VENCI% >= DES% Then
         If TRIM$(BANCO$) <> "" Then
          If InStr(BANCO$, "EFECT") < 1 Then
           If InStr(UCase$(BANCO$), "EFECT") < 1 Then
            If InStr(UCase$(BANCO$), "DOLAR") < 1 Then
             If InStr(UCase$(BANCO$), "PATACON") < 1 Then
              FEGA% = VENCI%
              If FEGA% <= HAS% Then
                SALDEU# = IMPO#
                GoSub 80
              End If
             End If
            End If
           End If
          End If
         End If
        End If
       End If
    Next I&
    '
    Call CIERRARCH("*.*")
    Exit Sub
    '
80  If FEGA% <= FECOR% Then
       KKI% = 1
     Else
       KKI% = 2 + Int((FEGA% - FECOR%) / 7)
    End If
    If KKI% > 4 Then KKI% = 4
    If KKI% > 0 Then
      If KKI% <= 4 Then
        SUMACOL# = Val(CHETERCE(KKI%).Caption) + SALDEU#
        PAVALDIF(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
        PAVALDIF(KKI%).Refresh
      End If
    End If
    '
89  Return
   '
End Sub

Sub CALDEUPRO()
    '
    MESASEL$ = MESAN.Caption
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    'If DES% > HOY(HOS$) Then
    '   Call MENSERR(24, "Período no Válido")
    '   Exit Sub
    'End If
    '
    Screen.MousePointer = 11
10  FEX = DES%: DESX$ = FECHATEX$(FEX)
    DISEDES% = DIASEM%(FEX)
11  FEX = HAS%: HASX$ = FECHATEX$(FEX)
    PERIO$ = MESASEL$
    FIPRISEM% = 7 - DISEDES%
    If FIPRISEM% < 6 Then FIPRISEM% = FIPRISEM% + 7
    FECOR% = DES% + FIPRISEM% - 1
    '
    FIN& = ULTREG&("CUECORR")
    For IX& = 1 To FIN&
       Call TRACE(20, IX&, FIN&)
       SS$ = REGLEIDO$("CUECORR", IX&)
       SS4$ = Mid$(SS$, 5, 2)
       SD1$ = Mid$(SS$, 115, 8)
       SH1$ = Mid$(SS$, 123, 8)
       SALDEUX$ = Mid$(SS$, 139, 60)
       '
       NC% = CVI(Mid$(SS$, 5, 2))
       If NC% > 0 Then
          'If TICUE%(NC%) = 0 Then   ' saltea cuentas de gastos
             For KKX% = 1 To 6
                FEGA% = CVI(Mid$(SALDEUX$, 10 * KKX% - 9, 2))
                If FEGA% <= HAS% Then
                  SALDEU# = CVD(Mid$(SALDEUX$, 10 * KKX% - 7, 8))
                  If SALDEU# >= 0.005 Then
                    GoSub 80
                  End If
                End If
             Next KKX%
          'End If
       End If
    Next IX&
    '
    Call CIERRARCH("*.*")
    Exit Sub
    '
80 If FEGA% <= FECOR% Then
       KKI% = 1
     Else
       KKI% = 2 + Int((FEGA% - FECOR%) / 7)
   End If
   If KKI% > 4 Then KKI% = 4
   If KKI% > 0 Then
     If KKI% <= 4 Then
       SUMACOL# = Val(DEUPROVE(KKI%).Caption) + SALDEU#
       DEUPROVE(KKI%).Caption = CSTRING$(MKD$(SUMACOL#), 4, 10, 2, 2)
       DEUPROVE(KKI%).Refresh
     End If
   End If
   '
89 Return
   '

End Sub

Sub LIPREFIN()
   '
   Dim SUINGRE#(5), SUEGRE#(5), SUGASVA#(5)
   Dim CAFLOSE#(5), CAFLOACU#(5)
   '
   JJ& = 0
   REBLA$ = REGBLAN$("LIPREFIN")
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "I N G R E S O S"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "***************"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10001)
   Mid$(REPRE$, 17, 48) = "Deudores por Ventas"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(DEUDOVEN(KKI%).Caption))
      SUINGRE#(KKI%) = SUINGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10002)
   Mid$(REPRE$, 17, 48) = "Estimado Incobrables (-)"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(INCOBRA(KKI%).Text))
      SUINGRE#(KKI%) = SUINGRE#(KKI%) - CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10003)
   Mid$(REPRE$, 17, 48) = "Estimado Ingresos Contado"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(INGREVAR(KKI%).Text))
      SUINGRE#(KKI%) = SUINGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10004)
   Mid$(REPRE$, 17, 48) = "Cheques de Terceros"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(CHETERCE(KKI%).Caption))
      SUINGRE#(KKI%) = SUINGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "T O T A L   I N G R E S O S"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(SUINGRE#(KKI%))
      CAFLOSE#(KKI%) = CAFLOSE#(KKI%) + SUINGRE#(KKI%)
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(248, "*")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "E G R E S O S"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "*************"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(-10001)
   Mid$(REPRE$, 17, 48) = "Deuda con Proveedores"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(DEUPROVE(KKI%).Caption))
      SUEGRE#(KKI%) = SUEGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(-10002)
   Mid$(REPRE$, 17, 48) = "Pagos a Diferir (-)"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(PADIFER(KKI%).Text))
      SUEGRE#(KKI%) = SUEGRE#(KKI%) - CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(-10003)
   Mid$(REPRE$, 17, 48) = "Estimado Compras Contado"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(COMPRACO(KKI%).Text))
      SUEGRE#(KKI%) = SUEGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(-10004)
   Mid$(REPRE$, 17, 48) = "Pago Valores Diferidos"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(Val(PAVALDIF(KKI%).Caption))
      SUEGRE#(KKI%) = SUEGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   '
   For U& = 1 To ULTREG&("!PREGASTO")
      X$ = REGLEIDO$("!PREGASTO", U&)
      CUE% = CUEINT%(TRIM$(Mid$(X$, 3, 12)))
      REFE$ = TRIM$(DENOCUE$(CUE%)) + " (" + TRIM$(Mid$(X$, 3, 12)) + ")"
      REPRE$ = X$
      Mid$(REPRE$, 17, 48) = REFE$
      SUMEN# = 0
      For KKI% = 1 To 4
        SUMEN# = SUMEN# + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
        SUGASVA#(KKI%) = SUGASVA#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
        SUEGRE#(KKI%) = SUEGRE#(KKI%) + CVD(Mid$(REPRE$, 57 + 8 * KKI%, 8))
      Next KKI%
      Mid$(REPRE$, 97, 8) = MKD$(SUMEN#)
      SUGASVA#(5) = SUGASVA#(5) + SUMEN#
      SUEGRE#(5) = SUEGRE(5) + SUMEN#
      JJ& = JJ& + 1
      Call GRAREG("LIPREFIN", REPRE$, JJ&)
    Next U&
    '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "T O T A L   G A S T O S   V A R I O S"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(SUGASVA#(KKI%))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "T O T A L   E G R E S O S"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(SUEGRE#(KKI%))
      CAFLOSE#(KKI%) = CAFLOSE#(KKI%) - SUEGRE#(KKI%)
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(248, "*")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "F L U J O   D E   C A J A"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "*************************"
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "Cash-Flow Semanal"
   For KKI% = 1 To 5
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(CAFLOSE#(KKI%))
      CAFLOACU#(KKI%) = CAFLOACU#(KKI% - 1) + CAFLOSE#(KKI%)
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = REBLA$
   Mid$(REPRE$, 15, 2) = MKI$(10099)
   Mid$(REPRE$, 17, 48) = "Cash-Flow Acumulado"
   For KKI% = 1 To 4
      Mid$(REPRE$, 57 + 8 * KKI%, 8) = MKD$(CAFLOACU#(KKI%))
   Next KKI%
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   REPRE$ = String$(256, "-")
   JJ& = JJ& + 1
   Call GRAREG("LIPREFIN", REPRE$, JJ&)
   '
   '
   Call SETULTREG("LIPREFIN", JJ&)
   Call CIERRARCH("LIPREFIN")
   '
   Call HEADERS("LIPREFIN", "")
   Call HEADERS("LIPREFIN", "Corresponde a: " + TRIM$(MESAN.Caption))
   '
   Call CONECRUN("*LIPREFI", "Presupuesto Financiero Mensual")
   Screen.MousePointer = 1
   '
End Sub

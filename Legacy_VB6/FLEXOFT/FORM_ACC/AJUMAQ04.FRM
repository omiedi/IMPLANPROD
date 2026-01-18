VERSION 5.00
Begin VB.Form AJUMAQ04 
   BackColor       =   &H00EEEEB0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stocks a Fecha"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8940
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1180
      Left            =   4560
      ScaleHeight     =   1125
      ScaleWidth      =   3990
      TabIndex        =   43
      Top             =   240
      Width           =   4050
      Begin VB.CommandButton Command26 
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
         Picture         =   "AJUMAQ04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1400
         Picture         =   "AJUMAQ04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   210
         Width           =   650
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
         Height          =   690
         Left            =   750
         Picture         =   "AJUMAQ04.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   210
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   2760
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   44
         Top             =   240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   2330
         Picture         =   "AJUMAQ04.frx":0896
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   480
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEB0&
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
      Height          =   1850
      Left            =   210
      TabIndex        =   35
      Top             =   1875
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Motivo - Referencia"
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
         Left            =   140
         TabIndex        =   40
         Top             =   945
         Width           =   3820
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
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   315
            Width           =   3480
         End
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
         Left            =   315
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   580
         Width           =   2220
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
         Left            =   2730
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   580
         Width           =   1065
      End
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
         Height          =   285
         Left            =   3780
         TabIndex        =   36
         Top             =   580
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   210
         TabIndex        =   41
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   2625
         TabIndex        =   39
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   4515
      TabIndex        =   18
      Top             =   1800
      Width           =   4215
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   150
         TabIndex        =   34
         Top             =   2520
         Width           =   3060
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EEEEB0&
            Caption         =   "Listado Impreso"
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
            Left            =   630
            TabIndex        =   30
            Top             =   735
            Width           =   1905
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00EEEEB0&
            Caption         =   "Por Pantalla"
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
            Left            =   630
            TabIndex        =   29
            Top             =   420
            Value           =   -1  'True
            Width           =   2115
         End
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
         Height          =   1020
         Left            =   3360
         Picture         =   "AJUMAQ04.frx":27B8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2625
         Width           =   645
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Rango Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   150
         TabIndex        =   22
         Top             =   525
         Width           =   1800
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
            Height          =   285
            Left            =   525
            TabIndex        =   23
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox Text13 
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
            Left            =   525
            TabIndex        =   24
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            TabIndex        =   33
            Top             =   500
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   105
            TabIndex        =   32
            Top             =   1130
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEEEB0&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   2100
         TabIndex        =   19
         Top             =   525
         Width           =   1905
         Begin VB.TextBox Text12 
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
            Left            =   525
            TabIndex        =   27
            Text            =   " "
            Top             =   1365
            Width           =   1065
         End
         Begin VB.TextBox Text11 
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
            Left            =   525
            TabIndex        =   25
            Text            =   " "
            Top             =   735
            Width           =   1065
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            Picture         =   "AJUMAQ04.frx":2AC2
            TabIndex        =   26
            Top             =   735
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1575
            Picture         =   "AJUMAQ04.frx":2DCC
            TabIndex        =   28
            Top             =   1365
            Width           =   175
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   105
            TabIndex        =   21
            Top             =   1130
            Width           =   750
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Top             =   500
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   4110
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3675
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
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
      Left            =   1050
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   210
      TabIndex        =   2
      Top             =   1080
      Width           =   4110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEB0&
      Caption         =   "Ingreso de Ajustes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   12
      Top             =   3765
      Width           =   4110
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
         Height          =   285
         Left            =   1785
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   735
         Width           =   1170
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
         Height          =   285
         Left            =   1785
         TabIndex        =   4
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1785
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   420
         Width           =   1170
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
         Left            =   1785
         TabIndex        =   5
         Top             =   1365
         Width           =   1170
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
         Height          =   600
         Left            =   3255
         Picture         =   "AJUMAQ04.frx":30D6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   645
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3255
         Picture         =   "AJUMAQ04.frx":33E0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salida:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1050
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "StockTeórico:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Ajustado:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1365
         Width           =   1590
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   2730
      TabIndex        =   9
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   105
      TabIndex        =   8
      Top             =   255
      Width           =   855
   End
End
Attribute VB_Name = "AJUMAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODENTRAJU%

Private Sub Acdir_Click()
   Call Command27_Click
End Sub

Private Sub Amasto_Click()
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
End Sub

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        If Index = 0 Then
            Text11.TEXT = VALACT1$("SELFECHA")
          Else
            Text12.TEXT = VALACT1$("SELFECHA")
        End If
    End If
End Sub

Private Sub Command1_Click()
    PREPA% = 0
    Hide
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    QUERSTK04.Show 1
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If App.EXEName = "INVENT04" Then
      Command14.Enabled = True
      Hide
      Exit Sub
    End If
    Call CONECRUN("INVENT04", "Inventarios Físicos y Valorizados")
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    Call GENKARDEX(Text1)
    Command15.Enabled = True
End Sub

Private Sub Command2_Click()
    COD$ = ""
    CI% = 0
    Text1.TEXT = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text6.TEXT = ""
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command26_Click()
  Command26.Enabled = False
    'Call MENSERR(24, "No hay Opciones de Configuración\en la Presente Versión.")
     OPMOMAQ04.Show 1
  Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    Call ACCDIR("AJUSTOCK")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("AJUSTOCK")
End Sub

Private Sub Command3_Click()
    Call AJUSTO3
    Call AJUSTO2
    Text6.TEXT = Text9.TEXT
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command33_Click()
'\\\OT-05-0704-WAR-02/11/05///
  Command33.Enabled = False
   '
   Dim CATEO#(32767)
   '
   If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   '
10 Call SELECHO("SELFECHA")
   FA$ = VALACT1$("SELFECHA")
   FA1% = FECHANUM(FA$)
   If FA1% < 0 Then
     Call MENSERR(24, "Fecha de Ajuste no Válida")
     GoTo 10
   End If
   If FA1% = 0 Then
    GoTo 99
   End If
   '
20 Call SELECHO("TADEPOSI")
   DP1% = Val(VALACT1$("TADEPOSI"))
   If DP1% <= 0 Then
     GoTo 10
   End If
   '
   DPT1$ = FORMATNUM$(DP1%, "F3.0") + " " + ECOARCH$("TADEPOSI", Chr$(DP1%))
   If COMALTER%("Esta Seguro de Ajustar a cero\todos los Stocks del Depósito Nro. : " & DPT1$ & "\ al " & FA$ & " \\No, Cancelar\Si, Ajustar a Cero") = 2 Then
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
     SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC"
     UNUM& = 0
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     On Error Resume Next
     MoviTemp.MoveFirst
     If Err < 1 Then
       UNUM& = Val(Mid$(MoviTemp!DoPriCor, 4))
     End If
     On Error GoTo 0
     NRO& = 1 + UNUM&
     MoviTemp.Close
     '
     For KI% = 1 To NUMAS%
       CATEO#(KI%) = 0
     Next KI%
     '
     Call ABRESTOCKS
     '
     CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
     If CMOV$ = "" Then CMOV$ = "AJ"
     DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
     DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
     DOPRILA$ = DOPRICO$
     DOSELA$ = DOSECO$
     '
     For CI% = 1 To NUMAS%
       '
       Screen.MousePointer = 11
       Call VESARRAS(CI%)
       '
       SQLX$ = " SELECT * FROM MOVISTO "
       SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
       SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FA1%)) & " "
       SQLX$ = SQLX$ + "AND DepoMovi = " & DP1% & " "
       '
       Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
       With MoviTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             CANTI1# = !CantIngre - !CantSalid
             CATEO#(CI%) = CATEO#(CI%) + CANTI1#
             .MoveNext
           Loop
         End If
       End With
       '
       If Abs(CATEO#(CI%)) < 0.001 Then GoTo 30
       'NUORIT% = 0
       NUORIT% = NUORIT% + 1
       VUNI# = 0
       MONEII% = 0
       NC% = 0
       DEPO% = DP1%
       CANTI = CATEO#(CI%) * (-1)
       REFE$ = "Ajuste de Stocks a Cero"
       '
       Call MOVISTOK(FA1%, CI%, LTX$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
       '
30     IU& = Val(CI%)
       FIN& = Val(NUMAS%)
       Call TRACE(24, IU&, FIN&)
       Screen.MousePointer = 1
    Next CI%
    '
   Else
     GoTo 20
   End If
   '
   Call COMUNI("Proceso Completo")
99 Command33.Enabled = True
'\\\OT-05-0704-WAR-FIN///
End Sub

Private Sub Command4_Click()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Command6_Click()
CI% = CODINT%(Text1)
MsgBox STOCKACT(CI%)
For KKI% = 1 To 4
MsgBox STODEPOS(CI%, KKI%)
Next KKI%
End Sub

Private Sub Command7_Click()
    '
    DESNUM& = 1: If Val(Text14.TEXT) > 0 Then DESNUM& = Val(Text14.TEXT)
    DESNUX$ = "AS." + TRIM$(Str$(DESNUM&))
    HASNUM& = 999999
    If TRIM$(Text13.TEXT) <> "" Then
        If Val(Text13.TEXT) >= DESNUM& Then HASNUM& = Val(Text13.TEXT)
    End If
    HASNUX$ = "AS." + TRIM$(Str$(HASNUM&))
    DESFE% = 0: If FECHANUM(TRIM$(Text11.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text11.TEXT))
    HASFE% = 32767
    If TRIM$(Text12.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text12.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text12.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    '
    FEC1# = CDbl(CVDAT(DESFE%))
    FEC2# = CDbl(CVDAT(HASFE%))
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  'AJ' "
    SQLX$ = SQLX$ + " and DoPriCor >= '" & DESNUX$ & "' "
    SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & FEC1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & FEC2# & " "
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) ASC"
    '
    Set AJUSMOVTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With AJUSMOVTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Ajustes Registrados segun Filtro")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      Do While .EOF = False
      REFE$ = !ReferCor
      CI1% = !Cod_Inte
      DP1% = !DepoMovi
      CAN# = !CantIngre
      FF% = CVINT(!FechMov)
      NRO& = Mid(!DoPriCor, 4)
      Y$ = REGBLAN$("LISTRAJ")
      Call REPLA(Y$, MKI$(FF%), 1, 2)
      Call REPLA(Y$, REFE$, 3, 48)
      Call REPLA(Y$, MKS$(NRO&), 51, 4)
      Call REPLA(Y$, MKI$(CI1%), 55, 2)
      Call REPLA(Y$, UNIMED$(CI1%), 57, 4)
      If CAN# > 0 Then
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CantSalid), 69, 8)
        Else
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CantSalid), 69, 8)
      End If
      Call REPLA(Y$, Chr$(DP1%), 80, 1)
      Call REGAPP("LISTRAJ", Y$)
      
      .MoveNext
9      Loop
    End With
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        'Call OPNOIN("")
        Call COMUNI("Solo Disponible para Impresión. Genere el \ Listado y Exportelo a un Archivo de Texto")
        GoTo 10
      ElseIf OPCI% = 2 Then
10      Call CONECRUN("*LIAJUST", "")
    End If
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    Call TRANSLABELS(Name)
    Call AJUSTO1
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub

Private Sub Invent_Click()
   Call Command14_Click
End Sub

Private Sub Kardex_Click()
   Call Command15_Click
End Sub

Private Sub LiajuPan_Click()
   Option5.Value = True
   Call Command7_Click
End Sub

Private Sub LiajuPrin_Click()
   Option4.Value = True
   Call Command7_Click
End Sub

Private Sub List1_Click()
    '
    ILISTA& = List1.ListIndex
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    '
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If Val(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * Val(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
End Sub

Private Sub NUaju_Click()
   Call Command4_Click
End Sub

Private Sub Reports_Click()
   Call Command13_Click
End Sub

Private Sub Setup_Click()
   Call Command26_Click
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        Call AJUSTO2               ' recalcular existencias a fecha
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text10_Change()
    If Visible = True Then
        If FECHANUM(Text10.TEXT) > 0 Then
            Call AJUSTO2
        End If
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text2_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text3_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text4_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text6_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SetFocus
End Sub
'

Private Sub Text7_Change()
    If MODENTRAJU% = 7 Then
        Text8.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text8_Change()
    If MODENTRAJU% = 8 Then
        Text7.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text9_Change()
    If MODENTRAJU% = 9 Then
       DIFER# = Val(Text9.TEXT) - Val(Text6.TEXT)
       Text7.TEXT = ""
       Text8.TEXT = ""
       If DIFER# > 0 Then
           Text7.TEXT = Str$(DIFER#)
         ElseIf DIFER# < 0 Then
           Text8.TEXT = Str$(Abs(DIFER#))
       End If
    End If
End Sub
'
Private Sub Text7_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
    MODENTRAJU% = 7
End Sub

Private Sub Text8_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
    MODENTRAJU% = 8
End Sub
'

Private Sub Text9_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
    MODENTRAJU% = 9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command3.SetFocus
    End If
End Sub

Sub AJUSTO1()
    '
    HOOI% = HOY(HOS$)
    '
    List1.Clear
    Call CARDEPOS
    '
111 If NUSU% < 1 Then
        Call MENSERR(24, "Falta Definir\Depósitos y Sucursales.")
        Hide
        Call FRESHALL
        PREPA% = 0
        Exit Sub
    End If
    '
115 Text10.TEXT = HOS$
    Text5.TEXT = "Diferencia de Inventario"
    '
End Sub
'

Sub AJUSTO2()
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
        Text10.TEXT = HOS$
        Text10.SetFocus
        Exit Sub
    End If
    '
    CI% = CODINT%(Text1)
    If CI% > 0 Then
      '
      Screen.MousePointer = 11
      Call VESARRAS(CI%)
      For KI% = 1 To 64
        CATEO#(KI%) = 0
      Next KI%
      '
      Call ABRESTOCKS
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND INT(CDBL(FechMov))<= " & CDbl(CVDAT(FF%)) & " "
      '
      Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With MoviTemp
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
            CANTI# = !CantIngre - !CantSalid
            DEP% = !DepoMovi
            LOT$ = !IdenLote
            '
            COLUX% = 1
            For KI% = 1 To NUSU%
              If DEP% = COSU%(KI%) Then COLUX% = KI%
            Next KI%
            '
            CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
          .MoveNext
          Loop
        End If
      End With
    End If
    '
    List1.Clear
    For U& = 1 To NUSU%
       CAN = CATEO#(U&)
       List1.AddItem SUCU$(U&) + CSTRING$(MKS$(CAN), 4, 10, 1, 2)
    Next U&
    '
    List1.ListIndex = ILISTA&
    List1.Refresh
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    If Val(Text9.TEXT) < 0 Then
        Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * Val(Text9.TEXT)), 4, 10, 1, 2))
        Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub AJUSTO3()
    '
    Static YAFORI%, FORIPRE$   ', UREMOVI&
    '
'    If UREMOVI& < 1 Then
'       UREMOVI& = ULTREG&("MOVISTO")
'       Call CIERRARCH("MOVISTO")
'    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORAJUST"))
    YAFORI% = 1
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Ajuste de Stocks:\  \'" + RMCC$ + "'.")
        YAFORI% = (-1)
      End If
    End If
    '
    HOII% = HOY(HOS$)
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
    '\\\OT-05-0303-WAR-19/05/05///
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC"
    '\\\OT-05-0303-WAR-FIN///
    UNUM& = 0
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    MoviTemp.MoveFirst
    If Err < 1 Then
      UNUM& = Val(Mid$(MoviTemp!DoPriCor, 4))
    End If
    On Error GoTo 0
    NRO& = 1 + UNUM&
    MoviTemp.Close
    '
    CAVIE = Val(Text6.TEXT)
    CANUE = Val(Text9.TEXT)
    '
    FF% = FECHANUM(Text10.TEXT)
    If FF% < 1 Then
      Call MENSERR(24, "Fecha Incorrecta")
      Text10.SetFocus
      Exit Sub
    End If
    '
    CIII% = CODINT%(Text1.TEXT)
    If CIII% < 1 Then
      Call MENSERR(24, "Falta Código de Material")
      Text1.SetFocus
      Exit Sub
    End If
    '
    LTX$ = LOTE$
    If Left$(UCase$(LOTE$), 5) = "TODOS" Or TRIM$(LOTE$) = "" Then
      LTX$ = Space$(16)
    End If
    ILISTA& = List1.ListIndex
    DEPO% = COSU%(ILISTA& + 1)
    FECHA% = FF%
    CANTI = CANUE - CAVIE
    '
    If Abs(CANTI) >= 0.05 Then
        '
        CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
        If CMOV$ = "" Then CMOV$ = "AJ"
        DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
        NUORIT% = 0
        VUNI# = 0
        MONEII% = 0
        NC% = 0
        REFE$ = Text5.TEXT
        If TRIM$(REFE$) = "" Then
            REFE$ = "Ajuste de Inventario"
        End If
        '
        Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
        '
237     If TRIM$(FORIPRE$) <> "" Then
          If YAFORI% > 0 Then
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CIII%)), "COD-MAT", 0, 0)
               Call PRINTFORWIN(DESCRIT$(CIII%), "DES-MAT", 0, 0)
               Call PRINTFORWIN(UNIMED$(CIII%), "UNIMED", 0, 0)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DEPO%)), "DEPOSITO", 0, 0)
               Call PRINTFORWIN(LTX$, "NRO-LOTE", 0, 0)
               CANS$ = CSTRING$(MKS$(CANTI), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, 0)
               Call PRINTFORWIN(REFE$, "REF-MAT1", 0, 0)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Ajuste de Stocks" + Str$(NRO&), OKX%)
          End If
        End If
        '
        Call COMUNI("Ajuste Número" + Str$(NRO&) + "\Completo y Verificado")
        '
299 End If
    '
    Call CIERRARCH("*.*")
    '
End Sub
'



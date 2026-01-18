VERSION 5.00
Begin VB.Form ACLISPRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de Precios de Venta"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11670
   Icon            =   "ACLISPRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11670
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004040&
      Height          =   8290
      Left            =   10815
      ScaleHeight     =   8235
      ScaleWidth      =   1845
      TabIndex        =   33
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox Picture16 
         Height          =   8300
         Left            =   0
         ScaleHeight     =   8235
         ScaleWidth      =   1020
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   8300
            Left            =   0
            Picture         =   "ACLISPRE.frx":030A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command8 
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
         Left            =   113
         Picture         =   "ACLISPRE.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Emite Lista de Precios Impresa"
         Top             =   6250
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   525
         TabIndex        =   43
         Top             =   7245
         Width           =   585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00004040&
         Caption         =   "Prn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   950
         Left            =   105
         TabIndex        =   42
         Top             =   5985
         Width           =   585
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00004040&
         Caption         =   "ABM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3030
         Left            =   105
         TabIndex        =   37
         Top             =   1995
         Width           =   585
         Begin VB.CommandButton Command7 
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
            Left            =   20
            Picture         =   "ACLISPRE.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Nueva Lista de Precios"
            Top             =   1645
            Width           =   600
         End
         Begin VB.CommandButton COMMAND3 
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
            Left            =   0
            Picture         =   "ACLISPRE.frx":2CB6
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Guarda Configuración de la Lista"
            Top             =   260
            Width           =   600
         End
         Begin VB.CommandButton Command9 
            Caption         =   "More"
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
            Left            =   0
            Picture         =   "ACLISPRE.frx":2FC0
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Opciones Complementarias"
            Top             =   2345
            Width           =   600
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
            Left            =   20
            Picture         =   "ACLISPRE.frx":32CA
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Alta de Nuevo Item de Stock"
            Top             =   955
            Width           =   600
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00004040&
         Caption         =   "Ctrl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1650
         Left            =   105
         TabIndex        =   34
         Top             =   50
         Width           =   585
         Begin VB.CommandButton Command5 
            Caption         =   "Help"
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
            Left            =   20
            Picture         =   "ACLISPRE.frx":3414
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   955
            Width           =   600
         End
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
            Left            =   20
            Picture         =   "ACLISPRE.frx":371E
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Terminar - Salir de Listas de Precios de Venta"
            Top             =   265
            Width           =   600
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ACLISPRE.frx":3868
         Top             =   7770
         Width           =   1515
      End
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Utilidades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   8715
      TabIndex        =   22
      Top             =   210
      Width           =   1740
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   29
         Top             =   1890
         Width           =   1450
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "ACLISPRE.frx":578A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Control y Depuración de Lista Activa"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Control"
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
            Left            =   630
            TabIndex        =   31
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   26
         Top             =   1155
         Width           =   1450
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "ACLISPRE.frx":5A94
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Ordena por Código de Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort por Código"
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
            Left            =   630
            TabIndex        =   28
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   140
         ScaleHeight     =   585
         ScaleWidth      =   1395
         TabIndex        =   23
         Top             =   420
         Width           =   1450
         Begin VB.CommandButton Command11 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Actualización por Porcentaje"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste +/- (%)"
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
            TabIndex        =   25
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   210
      TabIndex        =   0
      Top             =   3675
      Width           =   10410
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   10365
      TabIndex        =   20
      Top             =   3255
      Width           =   10430
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"ACLISPRE.frx":5D9E
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
         Left            =   0
         TabIndex        =   21
         Top             =   60
         Width           =   10410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Lista de Precio Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   4620
      TabIndex        =   3
      Top             =   210
      Width           =   3945
      Begin VB.TextBox Text4 
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
         Left            =   2835
         MaxLength       =   8
         TabIndex        =   47
         Top             =   1605
         Width           =   960
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   2250
         Width           =   435
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
         Height          =   285
         Left            =   945
         TabIndex        =   15
         Top             =   2220
         Width           =   175
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
         Left            =   945
         TabIndex        =   11
         Top             =   1590
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3570
         TabIndex        =   9
         Top             =   990
         Width           =   175
      End
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
         Height          =   285
         Left            =   630
         TabIndex        =   6
         Top             =   1000
         Width           =   1590
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   46
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2730
         TabIndex        =   18
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Left            =   1260
         TabIndex        =   17
         Top             =   2235
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   16
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   14
         Top             =   2235
         Width           =   330
      End
      Begin VB.Label Label6 
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
         Left            =   1260
         TabIndex        =   13
         Top             =   1605
         Width           =   1380
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formato:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   12
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   10
         Top             =   1600
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   2730
         TabIndex        =   8
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vigencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2100
         TabIndex        =   7
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abrev:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -105
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Indice de Listas de Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   4260
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   3795
      End
   End
   Begin VB.Line Line3 
      X1              =   210
      X2              =   11865
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Line Line1 
      X1              =   10635
      X2              =   10635
      Y1              =   -105
      Y2              =   8295
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu altaitem 
         Caption         =   "Alta de Items de Stocks"
      End
      Begin VB.Menu newlist 
         Caption         =   "Nueva Lista de Precios"
      End
      Begin VB.Menu grabarlistaactiva 
         Caption         =   "Grabar Lista Activa"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu imprimelistaactiva 
         Caption         =   "Imprimir Lista Activa"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu reunilistadeprecios 
         Caption         =   "Re-Unir Listas de Precios "
      End
   End
End
Attribute VB_Name = "ACLISPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LISTACTIVA&

Private Sub altaitem_Click()
   Call COMMAND21_CLICK
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command10_Click()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If DERACCE%("APREPOR", "Actualización de Precios por Porcentaje") = 2 Then
      Call APREPOR
      Screen.MousePointer = 1
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SORLIPRE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then GoTo 99
   End If
   '
   If DERACCE%("DEPUPREC", "Depurar Listas de Precios de Venta") < 2 Then
     GoTo 99
   End If
   '
   'ABRE EL FORMULARIO DE SELECCION DE OPERACION
   CLEANPRE.Show 1
   DoEvents
   '
   If CLEANPRE.Check1(2).Value = Checked Then
     Call VEPRECOS
     GoTo 99
   End If
   '
   LICOBAJ% = 0: If CLEANPRE.Check1(0).Value = Checked Then LICOBAJ% = 1
   LISINST% = 0: If CLEANPRE.Check1(1).Value = Checked Then LISINST% = 1
   CLEANPRE.Check1(0).Value = Unchecked
   CLEANPRE.Check1(1).Value = Unchecked
   CLEANPRE.Check1(2).Value = Unchecked
   DoEvents
   '
   ' SI SELECCIANA CHECK1(ELIMINAR ITEMS DE BAJA)
   If LICOBAJ% > 0 Or LISINST% > 0 Then
     If LISTACTIVA& < 1 Then GoTo 99
     ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
     While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ARCHIPRE$
     '
     Screen.MousePointer = 11
     JJ& = 0
     'FIN& = ULTREG&(ARCHIPRE$)
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + " ORDER BY Num_Orden ;"
     Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
     With ARTSELEC
         On Error Resume Next
        .MoveFirst
        FIN& = .RecordCount
        If Err > 0 Then
          On Error GoTo 0
          Call MENSERR(24, "Lista Vacía")
          Exit Sub
        End If
        On Error GoTo 0
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          'X$ = REGLEIDO$(ARCHIPRE$, U&)
          'CIXI% = CVI(Left$(X$, 2))
          x$ = Space$(64)
          CIXI% = !Cod_Pieza
          cx = CODEXT$(CIXI%)
          Call REPLA(x$, MKI(CIXI%), 1, 2)
          Call REPLA(x$, DESCRIT$(CIXI%), 3, 42)
          Call REPLA(x$, !Carac_Control, 45, 1)
          Call REPLA(x$, !Num_Original, 46, 16)
             PRUNI# = !Precio_Unitario
          Call REPLA(x$, MKD$(PRUNI#), 62, 8)
          .MoveNext
            '****
          If LICOBAJ% = 1 Then
            'If Mid$(CODEXT$(CIXI%), 16, 1) = Chr$(96) Then
            ESTADO& = STATITEM%(CIXI%)
            If ESTADO& = -1 Then
              GoTo 19
            End If
          End If
          '
          If LISINST% = 1 Then
            If STOCKACT(CIXI%) < 0.05 Then
              GoTo 19
            End If
          End If
          '
          JJ& = JJ& + 1
          Call GRAREG("!LIPREVEN", x$, JJ&)
          '
19      Next U&
     End With
     '
     Call SETULTREG("!LIPREVEN", JJ&)
     Call CIERRARCH("!LIPREVEN")
     Screen.MousePointer = 1
     '
     Call EDITLISTA(ARCHIPRE$)
     
   End If
   '
99 Command13.Enabled = True
   '
End Sub

Sub VEPRECOS()
   '
   If LISTACTIVA& < 1 Then Exit Sub
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   'MUESTRA UN FORMULARIO SOLICITANDO EL RANGO DE PRECIOS
10 VDEF$ = MKS$(1.5) + MKS$(4)
   OBX$ = OBJPLA$("COEFICOSPRE", VDEF$)
   If OBX$ <> "" Then
     Screen.MousePointer = 11
     COEFIMI = CVS(Left$(OBX$, 4))
     COEFIMA = CVS(Mid$(OBX$, 5, 4))
     If COEFIMI > COEFIMA Then GoTo 10
     '
     'ARMA EL TITULO
     x$ = REGLEIDO$("LISTAS", LISTACTIVA&)
     TITU$ = TRIM$(Mid$(x$, 2, 46))
     Call HEADERS("LIPREVEN", "")
     Call HEADERS("LIPREVEN", "Lista: " + TITU$ + " - Precios a Revisar por Costos")
     '
     'LEE LOS REGISTROS DE LA LISTA SELECCIONADA
     JJ& = 0
     'FIN& = ULTREG&(ARCHIPRE$)
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + " ORDER BY Num_Orden ;"
     Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
     With ARTSELEC
      On Error Resume Next
      .MoveFirst
      FIN& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Lista Vacía")
        Exit Sub
      End If
      On Error GoTo 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        'XX$ = REGLEIDO$(ARCHIPRE$, U&)
        'CIXI% = CVI(Left$(XX$, 2))
        CIXI% = !Cod_Pieza
        'PRE# = CVD(Mid$(XX$, 62, 8))
        PRE# = !Precio_Unitario
        COCAL = COSCALCU(CIXI%, MODOCO%)
        If PRE# < COEFIMI * COCAL Or PRE# > COEFIMA * COCAL Then
          XX$ = Space$(64)
          Call REPLA(XX$, MKI$(CIXI%), 1, 2)
          Call REPLA(XX$, DESCRIT$(CIXI%), 3, 42)
          Call REPLA(XX$, !Denom_Pieza, 45, 0)
          Call REPLA(XX$, !Carac_Control, 45, 1)
          Call REPLA(XX$, !Num_Original, 46, 16)
          Call REPLA(XX$, MKD$(PRE#), 62, 8)
          Call REPLA(XX$, MKI$(!Cant_Limite), 70, 2)
          Call REPLA(XX$, MKI$(!Porc_Limite), 72, 1)
          JJ& = JJ& + 1
          Call GRAREG("LIPREVEN", XX$, JJ&)
          .MoveNext
        End If
      Next U&
     End With
     Call SETULTREG("LIPREVEN", JJ&)
     '
     If JJ& < 1 Then
          Call MENSERR(24, "Sin Valores Fuera de Rango.")
        Else
          Screen.MousePointer = 1
          Call FINAL("*LILIPRE") '
     End If
   End If
   '
End Sub
'
Private Sub Command2_Click()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label3 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub COMMAND21_CLICK()
   Command1.Enabled = False
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   '
   Command1.Enabled = True
End Sub

Private Sub Command3_Click()
   COMMAND3.Enabled = False
   LISTACTIVA& = List1.ListIndex + 1
   If LISTACTIVA < 1 Then
     Call MENSERR(24, "Imposible Procesar Transacción.\Ninguna Lista de Precios\Marcada Como Lista Activa.")
     GoTo 99
   End If
   '***********
   '//OT -07 - 0434//
   
   SQLX$ = "SELECT * FROM LISTAPRE "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
   Set LISPRE = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
   On Error Resume Next
   With LISPRE
      .MoveLast
      If Err > 0 Then
        Exit Sub
      End If
      .MoveFirst
      .Edit
      !Denom_Lista = Text1.TEXT
      !Vigencia = CVDAT(FECHANUM(Label3.Caption))
      !FORMATO = (Label4.Caption)
      If Val(Text3) < 0 Or Val(Text3) > 5 Then Text3 = "2"
      !Decimales = (Text3.TEXT)
      !MONEDA = (Label7.Caption)
      !Desc_Abrev = Text2.TEXT
      !Descu_lista = Text4.TEXT
      .Update
   End With
   
   '//OT -07 - 0434//
   'XX$ = REGLEIDO$("LISTAS", LISTACTIVA&)
   'Mid$(XX$, 2, 46) = AJUSTI$(Text1, 46)
   'Mid$(XX$, 51, 12) = AJUSTI$(Text2, 12)
   'FEXI% = FECHANUM(Label3)
   'Mid$(XX$, 63, 2) = MKI$(FEXI%)
   'Mid$(XX$, 48, 1) = Chr$(Val(Label4))
   'Mid$(XX$, 50, 1) = Chr$(Val(Label7))
   'Mid$(XX$, 49, 1) = Chr$(Val(Text3))
   'Call GRAREG("LISTAS", XX$, LISTACTIVA&)
   'Call CIERRARCH("LISTAS")
   Call CARILISTAS
99 COMMAND3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call OPNOIN("")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
        '
        Command7.Enabled = False

3000    ' ************************* nueva lista de precios
        '
        '//OT-07-0434//
        On Error Resume Next
        With ListaPrec
        .MoveLast
        If .RecordCount >= 99 Then
            Call COMUNI("Hay 99 Listas de Precios\ya Definidas\-\Capacidad del Sistema Saturada.")
            Exit Sub
        End If
        '
        NULIS% = 1 + .RecordCount
        x$ = Chr$(NULIS%) + Mid$(REGBLAN$("LISTAS"), 2)
3010    x$ = OBJPLA$("LISPRE-NUEVA", x$)
        If x$ = "" Then
            GoTo 99
        End If
        '
        If TRIM$(Mid$(x$, 2, 46)) <> "" Then
          If Asc(Mid$(x$, 48, 1)) > 0 Then
            If Asc(Mid$(x$, 48, 1)) <= 2 Then
              If Asc(Mid$(x$, 49, 1)) <= 4 Then
                If Asc(Mid$(x$, 50, 1)) > 0 Then
                  If Asc(Mid$(x$, 50, 1)) <= 2 Then
                    If TRIM$(Mid$(x$, 51, 12)) <> "" Then
                      If CVI(Mid$(x$, 63, 2)) > 0 Then
                        'Call REGAPP("LISTAS", X$) '
                        'Call CIERRARCH("LISTAS")
                        .AddNew
                        .Fields("Cod_Lista") = Asc(Mid$(x$, 1, 1))
                        .Fields("Denom_Lista") = (Mid$(x$, 2, 46))
                        .Fields("Vigencia") = FECHATEX(CVI(Mid$(x$, 63, 2)))
                        .Fields("Formato") = Asc(Mid$(x$, 48, 1))
                        .Fields("Decimales") = Asc(Mid$(x$, 49, 1))
                        .Fields("Moneda") = Asc(Mid$(x$, 50, 1))
                        .Fields("Desc_Abrev") = (Mid$(x$, 51, 12))
                        
                        .Update
                        Call CARILISTAS
                        '//OT-07-0434//
                        GoTo 99
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
        '
        Call COMUNI("Datos Incompletos\o No Válidos")
        GoTo 3010
        '
99      Command7.Enabled = True
       End With
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call LISTAPRE
  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPLPREAC.Show 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Me.Visible = False
   VERANTER$ = "ADMIPRE"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + " - " + EPAC$
   End If
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
     
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 7900
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If

   '
   Screen.MousePointer = 11
   If ULTREG&("VALIFORM") < 3 Then
      Call BLANARCH("VALIFORM")
      Call REGAPP("VALIFORM", Chr$(1) + "Nro. Original")
      Call REGAPP("VALIFORM", Chr$(2) + "Descuento por Cantidad")
      Call REGAPP("VALIFORM", Chr$(3) + "Fecha Ultima Actualizacion")
      Call CIERRARCH("VALIFORM")
   End If
   '
   If ULTREG&("VALIMONE") < 2 Then
      Call BLANARCH("VALIMONE")
      Call REGAPP("VALIMONE", Chr$(1) + "Pesos")
      Call REGAPP("VALIMONE", Chr$(2) + "Dolares U$S")
      Call CIERRARCH("VALIMONE")
   End If
   '
   CONTROLI$ = TRIM$(CONTROL$("LIPREVEN", "CONTRODE"))
   If CONTROLI$ = "" Then
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("LIPRE001")
       x$ = REGLEIDO$("LIPRE001", U&)
       CII% = CVI(Left$(x$, 2))
       DEAR$ = TRIM$(UCase$(DESCRIT0$(CII%)))
       If InStr(DEAR$, "ARTICULO") > 0 Then
         REMA& = CII% + 1
         DEAR$ = TRIM$(Mid$(x$, 3, 42))
         y$ = REGLEIDO$("MASTER", REMA&)
         Call REPLA(y$, DEAR$, 5, 64)
         Call GRAREG("MASTER", y$, REMA&)
       End If
     Next U&
     Call CIERRARCH("MASTER")
     'HOII% = HOY(HOS$)
     'Call GRACONTROL("LIPREVEN", "CONTRODE", HOS$)
     Screen.MousePointer = 1
   End If
   '
   Call CARILISTAS
   ILIOO& = 0
   If Val(APLINVO$) > 0 Then
     ILI00& = Val(APLINVO$) - 1
   End If
   LISTACTIVA& = ILI00&
   On Error Resume Next
   List1.ListIndex = ILI00&
   On Error GoTo 0
   Screen.MousePointer = 1
   '
   Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub grabarlistaactiva_Click()
    Call Command3_Click
End Sub

Private Sub imprimelistaactiva_Click()
   Call Command8_Click
End Sub

Private Sub Label15_DblClick()
   Call List2_DblClick
End Sub

Private Sub Label4_Change()
   Label6 = ECOARCH$("VALIFORM", Chr$(Val(Label4)))
   If Val(Label4) = 3 Then
       Label15 = " Código                     Descripción                                             M.     Fecha Act.                           Precio"
     Else
       Label15 = " Código                     Descripción                                             M.     Nro.Original                         Precio"
   End If
   Label15.Refresh
End Sub

Private Sub Label4_DblClick()
   Call SELECHO("VALIFORM")
   If VALACT1$("VALIFORM") <> "" Then
     Label4 = VALACT1$("VALIFORM")
   End If
End Sub

Private Sub Label7_Change()
   Label9 = ECOARCH$("VALIMONE", Chr$(Val(Label7)))
End Sub

Private Sub Label7_DblClick()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Private Sub List1_Click()
   If List1.ListIndex <> LISTACTIVA& - 1 Then
     Call CARLISTA
   End If
End Sub
   '\\OT-07-0424
Sub CARLISTA()
   '
   Screen.MousePointer = 11
   LISTACTIVA& = List1.ListIndex + 1
   APERBASDAT ("PEDCLI")
   SQLX$ = "SELECT * FROM LISTAPRE "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
   Set LISTASEL = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With LISTASEL
      .MoveLast
      If Err > 0 Then
        Exit Sub
      End If
      Text1 = !Denom_Lista
      Text2 = !Desc_Abrev
      Label3 = (Format$(!Vigencia, "dd/mm/yy"))
      Label4 = !FORMATO
      Label7 = !MONEDA
      Text3 = !Decimales
      Text4 = !Descu_lista
   End With
   
   SQLX$ = "SELECT * FROM PREVENTA "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden ;"
   Set PRODUC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With PRODUC
      .MoveLast
      If Err > 0 Then
        List2.Clear
        Screen.MousePointer = 1
        Exit Sub
      End If
      List2.Clear
      FIN& = .RecordCount
      .MoveFirst
      If FIN& > 0 Then
        Do While (.EOF <> True)
        TTXX$ = Space$(64)
        CII% = !Cod_Pieza
        PRE# = !Precio_Unitario
        Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
        Call REPLA(TTXX$, DESCRIT$(CII%), 18, 31)
        Call REPLA(TTXX$, !Carac_Control, 51, 2)
        Call REPLA(TTXX$, !Num_Original, 56, 13)
        Call REPLA(TTXX$, FORMATNUM$(PRE#, "F12,4"), 72, 12)

        List2.AddItem TTXX$
        .MoveNext
        Loop
      End If
   End With
   Screen.MousePointer = 1
  
   '\\OT-07-0424
   '
End Sub

Private Sub List2_DblClick()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
   End If
   '
   ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
   '
   n% = XVALO(TRIM$(Str$(LISTACTIVA&)))
   '\\OT-07-0424
   Call LeeLaLista(n%)
   '\\OT-07-0424
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% >= 0 Then
     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     List2.Clear
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For i& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", i&)
         CI% = CVI(Left$(NXX$, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, i&)
89     Next i&
     End If
     '
     PUNCODIAC$ = String$(NUMAS%, 0): JJ& = 0
     '\\OT-07-0424
     SQLX$ = "DELETE * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & n% & " ;"
     Ventas.Execute (SQLX$)
     '
     FIN& = ULTREG&("!LIPREVEN")
     For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", i&)
        CI% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            If Asc(Mid$(PUNCODIAC$, CI%, 1)) > 0 Then
              Call MENSERR(24, "Código '" + TRIM$(CODEXT$(CI%)) + "' Repetido en Lista.\Suprimida Segunda Aparición.")
              GoTo 94
            End If
            '
            ' QUEDA PENDIENTE VER PARA QUE ES ESTO
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
            DE$ = DESCRIT0$(CI%)
            Call REPLA$(NXX$, DE$, 3, 42)
            Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(NXX$, MRX$, 44, 1)
            '
            With Preventa
              .AddNew
              !Cod_Lista = n%
              !Cod_Pieza = CI%
              !Num_Orden = i&
              !Desc_Pieza = DESCRIT0$(CI%)
              !Denom_Pieza = DESCRIT0$(CI%)
              !Carac_Control = Mid$(NXX$, 45, 1)
              !Num_Original = (Mid$(NXX$, 46, 16))
              !Precio_Unitario = CVD(Mid$(NXX$, 62, 8))
              !Cant_Limite = CVD(Mid$(NXX$, 70, 2))
              !Porc_Limite = (Mid$(NXX$, 72, 1))
              .Update
            End With
            '
            Mid$(PUNCODIAC$, CI%, 1) = Chr$(1)
          End If
        End If
94   Next i&
     '
     Screen.MousePointer = 1
     '
     Call CARLISTA
   End If
   ''\\OT-07-0424
End Sub

Sub APREPOR()
    '
    Dim VALVIE(32767), VALNUE(32767)
    HOII% = HOY(HOS$)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      If VERILLAVE% <> 1 Then Exit Sub
    End If
    ' SELECCIONA LA MODALIDAD DE ACTUALIZAR PRECIOS
    OPX% = COMALTER%("Puede Actualizar a Partir\de la Lista de Precios Actualmente Vigente\o de los Costos de Reposición para un Proveedor\\Lista Vigente\Costos de Reposición\Cancelar")
    Select Case OPX%
      Case 2
        Call APREPROV
        Exit Sub
      Case 3
        Exit Sub
    End Select
    '
    LISTACTIVA& = List1.ListIndex + 1
    'X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    RANGOAC% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
      TTXX$ = "Opciones de Actualización por Porcentaje:\\Sólo Bombas\Sólo Repuestos\Toda la Lista"
      RANGOAC% = COMALTER%(TTXX$)
      If RANGOAC% < 1 Or RANGOAC% > 3 Then Exit Sub
      If RANGOAC% = 3 Then RANGOAC% = 0
    End If
    '
    Select Case RANGOAC%
      Case 0
        RACTU$ = "Lista Completa"
      Case 1
        RACTU$ = "Bombas Nacionales (Cod. 0...)"
      Case 2
        RACTU$ = "Repuestos (Cod. >= 1...)"
    End Select
    '
    ' SELECCIONA RANGO DE CODIGOS
    Call DESHASCOD(COD1$, COD2$, CIKX1%, CIKX2%)
    If CIKX1% < 1 Or CIKX2% < 1 Then Exit Sub
    '
    '**********
    SQLX$ = "SELECT * FROM LISTAPRE "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
    'SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
    Set LISTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    With LISTSELEC
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
         On Error GoTo 0
         Exit Sub
         
      End If
      On Error GoTo 0
    
      Call REPLA(VDEF$, Chr(!Cod_Lista), 1, 1)
      Call REPLA(VDEF$, !Denom_Lista, 2, 46)
      Call REPLA(VDEF$, Chr(!FORMATO), 48, 1)
      Call REPLA(VDEF$, Chr(!Decimales), 49, 1)
      Call REPLA(VDEF$, Chr(!MONEDA), 50, 1)
      Call REPLA(VDEF$, !Desc_Abrev, 51, 12)
      'VDEF$ = Left$(X$, 64) + MKS$(0)
      Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
      'CADECI% = Asc(Mid$(VDEF$, 49, 1))
      CADECI% = !Decimales
      If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
      Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    End With
    
    ' INGRESO DEL PORCENTAJE
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    ' VALIDACIONES
    PORAJU = CVS(Mid$(OBX$, 65, 4)) 'PORCENTAJE DE AJUSTE
    If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 99.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
   '
   LISNO% = Asc(Left$(OBX$, 1))
   '
   Screen.MousePointer = 11
   '*************************
   '//OT-07-0434//
   '
   REBLA$ = REGBLAN$("ACOSPOR")
   ' LEEMOS LISTA DE PRECIOS ACTIVA
   SQLX$ = "SELECT * FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
   With ARTSELEC
     On Error Resume Next
     .MoveLast
     FIN& = .RecordCount
     If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Lista Vacía")
        Exit Sub
     End If
     On Error GoTo 0
     .MoveFirst
     U& = 0:   JJ& = 0
     Do While Not .EOF
            U& = U& + 1
            Call TRACE(20, U&, FIN&)
            TTXX$ = Space$(128)
            CII% = !Cod_Pieza
            CDXXX$ = CODEXT(!Cod_Pieza)
            '
            
            PORAJX = PORAJU
            If RANGOAC% > 0 Then
              CDDX$ = Left$(CDXXX$, 1) '
              If RANGOAC% = 1 Then
                  If CDDX$ <> "0" Then PORAJX = 0
                ElseIf RANGOAC% = 2 Then
                  If CDDX$ = "0" Then PORAJX = 0
              End If
            End If
            '
            If UCase$(CDXXX$) < UCase$(COD1$) Or UCase$(CDXXX$) > UCase$(COD2$) Then
              PORAJX = 0
            End If
            '
            COSVIE# = !Precio_Unitario
            COSNUE = ROUND#(COSVIE * (100 + PORAJX) / 100, CADECI%)
            '
            ZZ$ = REBLA$
            Call REPLA(ZZ$, MKI$(CII%), 1, 2)
            Call REPLA(ZZ$, MKS$(COSVIE#), 3, 4)
            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
            
            JJ& = JJ& + 1
            Call GRAREG("!ACOSPOR", ZZ$, JJ&)
        .MoveNext
     Loop
   End With
   '//OT-07-0434//
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    Screen.MousePointer = 1
    
    ' PRESENTA PRECIOS ACTUALES Y NUEVOS
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         'GRABAMOS LOS CAMBIOS
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
             For U& = 1 To URE&
             Call TRACE(20, U&, URE&)
             x$ = REGLEIDO$("!ACOSPOR", U&)
             CII% = CVI(Mid$(x$, 1, 2))
             'Y$ = REGLEIDO$(APRECIO$, U&)
             'MRX$ = Mid$(Y$, 45, 1)
             PRENUEVO = CVS(Mid$(x$, 7, 4))
             PREX$ = COSNUEX$
             '
             SQLX$ = "SELECT * FROM Preventa "
             SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISNO% & " "
             SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
             Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
             With ITEMSEL
              On Error Resume Next
             .MoveFirst
             FIN& = .RecordCount
             If Err > 0 Then
               On Error GoTo 0
               Call MENSERR(24, "Lista Vacía")
               Exit Sub
             End If
             On Error GoTo 0
             PREVIE# = !Precio_Unitario
             If Val(Label4) = 3 Then
              'PREVIE# = CVD(Mid$(Y$, 62, 8))
              If Abs(PREVIE# - PRENUEVO) < 0.00005 Or TRIM$(Mid$(y$, 46, 16)) = "" Then
                Call REPLA(y$, HOS$, 46, 16)
              End If
             End If
            '
             .Edit
             !Precio_Unitario = PRENUEVO
             .Update
             'Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
             'Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
             'Call REPLA(Y$, MRX$, 44, 1)
             'Call GRAREG(APRECIO$, Y$, U)
             End With
           Next U&

           Call CIERRARCH(APRECIO$)
           Call BAJALDISCO
           Screen.MousePointer = 1
          '
           Call CARLISTA
         '
      End If
    End If
    Exit Sub


End Sub


Sub APREPROV()
    '//OT-07-0434//
    LISTACTIVA& = List1.ListIndex + 1
    x$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    '
    Call SELECHO("ACREDOR")
    VDEVU$ = VALACT1$("ACREDOR")
    NPRO% = Val(VDEVU$)
    If NPRO% < 1 Then Exit Sub
    '
    
    RACTU$ = RASOCLI$((-1) * NPRO%)
    SQLX$ = "SELECT * FROM LISTAPRE "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & "; "
    Set LISTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    With LISTSELEC
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
         On Error GoTo 0
         Exit Sub
      End If
      On Error GoTo 0
    
      Call REPLA(VDEF$, Chr(!Cod_Lista), 1, 1)
      Call REPLA(VDEF$, !Denom_Lista, 2, 46)
      Call REPLA(VDEF$, Chr(!FORMATO), 48, 1)
      Call REPLA(VDEF$, Chr(!Decimales), 49, 1)
      Call REPLA(VDEF$, Chr(!MONEDA), 50, 1)
      Call REPLA(VDEF$, !Desc_Abrev, 51, 12)
      'VDEF$ = Left$(X$, 64) + MKS$(0)
      Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
      CADECI% = !Decimales

    'VDEF$ = Left$(X$, 64) + MKS$(0)
    Call REPLA(VDEF$, MKI$(HOII%), 63, 2)
    'CADECI% = Asc(Mid$(VDEF$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
    Call REPLA(VDEF$, Chr$(CADECI%), 49, 1)
    '
    End With
5   OBX$ = OBJPLA$("APREVENPOR/Actualizacion por Porcentaje - " + RACTU$, VDEF$)
    If OBX$ = "" Then
      Exit Sub
    End If
    '
    PORAJU = CVS(Mid$(OBX$, 65, 4))
    'If Abs(PORAJU) < 0.001 Then Exit Sub
    '
    If PORAJU < (-90) Or PORAJU > 299.99 Then
       Call MENSERR(24, "Porcentaje no Válido")
       GoTo 5
    End If
    '
    CADECI% = Asc(Mid$(OBX$, 49, 1))
    If CADECI% < 2 Or CADECI% > 4 Then
       Call MENSERR(24, "Decimales entre '2' y '4'")
       GoTo 5
    End If
    '
'    'LISNO% = Asc(Left$(OBX$, 1))
'    'ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
'                                     ALISTA$ = "0" + ALISTA$
'                                   Wend
                                   '
'    APRECIO$ = "LIPRE" + ALISTA$
'    Call COPYSTRU("LIPREVEN", APRECIO$)
'    REBLA$ = REGBLAN$("ACOSPOR")
'    FIN& = ULTREG&(APRECIO$)
    JJ& = 0
    '
    Screen.MousePointer = 11
    PUCOD$ = ""
    'For U& = 1 To FIN&
       ' LEEMOS LISTA DE PRECIOS ACTIVA
     SQLX$ = "SELECT * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
     SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
     Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
     With ARTSELEC
       On Error Resume Next
       .MoveLast
       FIN& = .RecordCount
       If Err > 0 Then
          On Error GoTo 0
          Call MENSERR(24, "Lista Vacía")
          Exit Sub
       End If
       On Error GoTo 0
       .MoveFirst
       U& = 0:   JJ& = 0
       Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          'I0$ = REGLEIDO$(APRECIO$, U&)
          'CIXI% = CVI(Left$(I0$, 2))
          CIXI% = !Cod_Pieza
          CDXXX$ = CODEXT$(CIXI%)
          '
          'COSVIE = CVD(Mid$(I0$, 62, 8))
          COSVIE = !Precio_Unitario
          COSNUE = COSVIE
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
          Call REPLA(ZZ$, MKS$(COSVIE), 3, 4)
          Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
          JJ& = JJ& + 1
          Call GRAREG("!ACOSPOR", ZZ$, JJ&)
            PUCOD$ = PUCOD$ + MKI$(CIXI%)
          .MoveNext
       Loop
    End With
    '
    AGREGAIT% = 0
    For CIXI% = 1 To NUMAS%
      If PROVHABI%(CIXI%) = NPRO% Then
        COSREP = COSPESOS(CIXI%)
        COSNUEX$ = CSTRING$(MKS$(COSREP * (100 + PORAJU) / 100), 3, 12, CADECI%, 2)
        COSNUE = Val(COSNUEX$)
        '
        For U& = 1 To JJ&
          If CVI(Mid$(PUCOD$, 2 * U& - 1, 2)) = CIXI% Then
            ZZ$ = REGLEIDO$("!ACOSPOR", U&)
            Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
            Call GRAREG("!ACOSPOR", ZZ$, U&)
            GoTo 29
          End If
        Next U&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
        Call REPLA(ZZ$, MKS$(0), 3, 4)
        Call REPLA(ZZ$, MKS$(COSNUE), 7, 4)
        JJ& = JJ& + 1
        Call GRAREG("!ACOSPOR", ZZ$, JJ&)
        PUCOD$ = PUCOD$ + MKI$(CIXI%)
        AGREGAIT% = 1
      End If
29  Next CIXI%
    '
    Call SETULTREG("!ACOSPOR", JJ&)
    Call CIERRARCH("!ACOSPOR")
    If AGREGAIT% > 0 Then
      Call COMUNI("Los Items Agregados Aparecen\al FINAL de la Lista.\ \Luego de Guardar los Nuevos Precios\Puede Realizar Ordenamiento Automatico.")
    End If
    Screen.MousePointer = 1
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Actualizacion de Precios por Porcentaje - " + RACTU$
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    VTB% = VENTABU%("ACOSPOR" + ATRI$)
    If VTB% >= 0 Then
      If COMALTER%("Desea Guardar los Nuevos Precios de Venta ?\\Cancelar\Guardar") = 2 Then
         '
         Screen.MousePointer = 11
         HOII% = HOY(HOS$)
         URE& = ULTREG&("!ACOSPOR")
         'UREA& = ULTREG&(APRECIO$)
         'LISNO% = Asc(Left$(OBX$, 1))
         For U& = 1 To URE&
           Call TRACE(20, U&, URE&)
           x$ = REGLEIDO$("!ACOSPOR", U&)
           CII% = CVI(Mid$(x$, 1, 2))
           SQLX$ = "SELECT * FROM Preventa "
           SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " "
           SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
           Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
           With ITEMSEL
             On Error Resume Next
             .MoveFirst
             If Err > 0 Then
               On Error GoTo 0
               .AddNew
               !Cod_Lista = LISTACTIVA&
               !Cod_Pieza = CII%
               !Desc_Pieza = DESCRIT0$(CII%)
               !Denom_Pieza = DESCRIT0$(CII%)
             Else
                .Edit
             End If
  '           If U& <= UREA& Then
  '               Y$ = REGLEIDO$(APRECIO$, U&)
  '             Else
  '               CIXI% = CVI(Left$(X$, 2))
  '               Y$ = REGBLAN$(APRECIO$)
  '               Call REPLA(Y$, MKI$(CIXI%), 1, 2)
  '               Call REPLA(Y$, DESCRIT0$(CIXI%), 3, 42)
  '           End If
             'MRX$ = Mid$(Y$, 45, 1)
             If Not IsNull(!Carac_Control) Then
              MRX$ = !Carac_Control
              Else: MRX$ = ""
             End If
             COSNUEX$ = CSTRING$(Mid$(x$, 7, 4), 3, 12, CADECI%, 2)
             PRENUEVO# = Val(COSNUEX$)
             PREX$ = COSNUEX$
             '
             If Val(Label4) = 3 Then
               'PREVIE# = CVD(Mid$(Y$, 62, 8))
               PREVIE# = !Precio_Unitario
               If Abs(PREVIE# - PRENUEVO#) < 0.00005 Or TRIM$(!Num_Original) = "" Then
                 'Call REPLA(Y$, HOS$, 46, 16)
                 !Num_Original = HOS$
               End If
             End If
             '
             'Call REPLA(Y$, MKD$(PRENUEVO#), 62, 8)
             !Precio_Unitario = PRENUEVO#
             'Call REPLA(Y$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
             'Call REPLA(Y$, MRX$, 44, 1)
             !Carac_Control = MRX$
             .Update
           End With
           'Call GRAREG(APRECIO$, Y$, U&)
         Next U&
         '
         'Call SETULTREG(APRECIO$, U&)
         'Call CIERRARCH(APRECIO$)
         'Call BAJALDISCO
         Screen.MousePointer = 1
         '
         '//OT-07-0434//
         Call CARLISTA
         If AGREGAIT% > 0 Then
           If COMALTER%("Desea Ordenar por Codigo Creciente ?\\No, Dejar\Si, Ordenar") = 2 Then
             Call Command12_Click
           End If
         End If
         '
      End If
    End If
    '

End Sub


'
Sub ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   Call ABREMASTER
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER
   Master.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Codint = " & CODII%
     Master.FindFirst ABUS$
     If Master.NoMatch = False Then
       CODII% = CODII% + 1
       GoTo 20
     End If
   End If
   '
   With Master
      .AddNew
      !CODIGO = CODINUE$
      !DESCRIPCION = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATMASTER% = 1
   End With
   '
   REMAS& = NUMAS% + 2
   OKX% = 1
   Call GENECOMAS
   '
99 Screen.MousePointer = 1

End Sub
'
Sub LISTAPRE()
        '//OT-07-0434//
2000    ' ************** impresion de listas de precios
        '
        LISTACTIVA& = List1.ListIndex + 1
        x$ = REGLEIDO$("LISTAS", LISTACTIVA&)
        TITU$ = TRIM$(Mid$(x$, 2, 46))
        '
        Call DESHASCOD(COD1$, COD2$, CIK1%, CIK2%)
        If CIK1% < 1 Or CIK2% < 1 Then Exit Sub
        '
2100    Call BLANARCH("LIPREVEN")
        Screen.MousePointer = 11
        '
        Call HEADERS("LIPREVEN", "")
        Call HEADERS("LIPREVEN", "Lista: " + TITU$)
        '
        ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
        LIPREAC$ = "LIPRE" + ARCHIPRE$
        FIN& = ULTREG&(LIPREAC$)
        '
        'CONSULTA ARTICULOS SEGUN LISTA SELECCIONADA
        SQLX$ = "SELECT * FROM Preventa "
        SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & ""
        SQLX$ = SQLX$ + " ORDER BY Num_Orden;"
        Set ARTSELE = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        On Error Resume Next
        With ARTSELE
          .MoveLast
          If Err > 0 Then
             Exit Sub
          End If
          FIN& = .RecordCount
          .MoveFirst
          If FIN& > 0 Then
          'Recorremos el recordset y selecionamos solo
          'aquellos que cumplen con los parametros seleccionados
          'en la ventana  DESHASCOD donde se selecciona desde y hasta
            Do While (.EOF = False)
                CII% = !Cod_Pieza
                CDXX$ = CODEXT$(!Cod_Pieza)
                If UCase(CDXX$) >= UCase(COD1$) Then
                  If UCase(CDXX$) <= UCase(COD2$) Then
                    TTXX$ = Space$(128)
                    Call REPLA(TTXX$, MKI$(CII%), 1, 2)
                    Call REPLA(TTXX$, DESCRIT$(CII%), 3, 42)
                    Call REPLA(TTXX$, !Denom_Pieza, 45, 0)
                    Call REPLA(TTXX$, !Carac_Control, 45, 1)
                    Call REPLA(TTXX$, !Num_Original, 46, 16)
                    PRUNI# = !Precio_Unitario
                    Call REPLA(TTXX$, MKD$(PRUNI#), 62, 8)
                    Call REPLA(TTXX$, MKI$(!Cant_Limite), 70, 2)
                    Call REPLA(TTXX$, MKI$(!Porc_Limite), 72, 1)
                    JJ& = JJ& + 1
                    Call REGAPP("LIPREVEN", TTXX$)
                  End If
                End If
                .MoveNext
            Loop
          End If
        End With
                                

              
              
        '***
        'For I& = 1 To FIN&
          'Call TRACE(20, I&, FIN&)
          'NXX$ = REGLEIDO$(LIPREAC$, I&)
          'CI% = CVI(Left$(NXX$, 2))
          'If CI% > 0 Then
            'If CI% <= NUMAS% Then
              'CDXX$ = CODEXT$(CI%)
              'If CDXX$ >= COD1$ Then
                'If CDXX$ <= COD2$ Then
                  'RELIB% = Val(Mid$(NXX$, 45, 1))
                  '
                  'For K% = 1 To RELIB%
                    'Call REGAPP("LIPREVEN", MKI$(32767) + Mid$(REGBLAN$("LIPREVEN"), 3))
                  'Next K%
                  'Call REGAPP("LIPREVEN", NXX$)
                  '
                'End If
              'End If
            'End If
          'End If
        'Next I&
        '
        'Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*LILIPRE") '
        '
        '//OT-07-0434//
End Sub

Sub SORLIPRE()
    '//OT-07-0434//
    ' NUMERO DE LA LISTA ACTIVA
    LISTACTIVA& = List1.ListIndex + 1
    'X$ = REGLEIDO$("LISTAS", LISTACTIVA&)
    'ARCHIPRE$ = TRIM$(Str$(LISTACTIVA&))
    'While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
    'LIPREAC$ = "LIPRE" + ARCHIPRE$
    '
    '    Call DATLIPRE(X$)
        '
    'TOMA EL VALOR DEL ULTIMO REGISTRO DE LA LISTA ACTIVA
    SQLX$ = "SELECT * FROM Preventa "
    SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " ;"
    Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 4)
    With ARTSELEC
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err > 0 Then
         On Error GoTo 0
         Call MENSERR(24, "Lista Vacía")
         Exit Sub
      End If
      On Error GoTo 0
      .MoveFirst
    'FIN& = ULTREG&(LIPREAC$)
    End With
    If FIN& < 1 Then
       Exit Sub
    End If
    
    'Call CIERRARCH(LIPREAC$)
    'LDT$ = LUDAT$ 'ESTO QUEDA ?????
    '
    'Call REGAPP(LIPREAC$, String$(128, 0))
    'Call CIERRARCH("*.*")
    'Call TRALOCAL(LIPREAC$, "LIPREVEN")
    '
    Screen.MousePointer = 11
    SORTLIST.Clear
    '
    'LEE LA LISTA ,Y COPIA EN SORTLIST
    For i& = 1 To FIN&
      With ARTSELEC
        Call TRACE(20, i&, FIN&)
        'NXX$ = REGLEIDO$(LIPREAC$, I&)
        'CI% = CVI(Left$(NXX$, 2)) 'CAPTURO EL CODIGO DE ACCESS
        CII% = !Cod_Pieza
        CODE$ = CODEXT$(CII%) 'OBTENGO EL CODIGO EXTERNO
        If CODE$ <> "" Then ' =
           Z$ = Space$(24)
           Call REPLA(Z$, "A" + CODE$, 1, 18) '
           Call REPLA(Z$, TRIM$(Str$(i&)), 19, 6)
           SORTLIST.AddItem Z$
           .MoveNext
        End If
      End With
    Next i&
    '
    FIN& = SORTLIST.ListCount
    '
    'J& = 0
    'LEE DEL SORTLIST Y GRABA EN LIPREAC
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       TESOR$ = SORTLIST.List(i& - 1)
       REG& = Val(Mid$(TESOR$, 19, 6))
       CODE$ = Mid$(TESOR$, 2, 17)
       CII% = CODINT%(CODE$)
       'PARA PREVENIR EN CASO QUE EL CODIGO ESTE DE BAJA
       'Y LO DEVUELVA EN NEGATIVO
       CII% = Abs(CII%)
       SQLX$ = "SELECT * FROM Preventa "
       SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA & " "
       SQLX$ = SQLX$ + " AND  Cod_Pieza = " & CII% & " ;"
       Set ITEMSEL = Ventas.OpenRecordset(SQLX$, 2)
       With ITEMSEL
          On Error Resume Next
          .MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 10
          End If
          On Error GoTo 0
          'If REG& > 0 Then
          .Edit
          !Num_Orden = i&
          .Update
             'NXX$ = REGLEIDO$("!LIPREVEN", REG&)
             'J& = J& + 1
             'Call GRAREG(LIPREAC$, NXX$, J&) 'REEMPLAZO ESTO GRABANDO EN ACCESS LOS DATOS DE NXX$
       'End If                               '
       End With
10  Next i&
    '
    'Call SETULTREG&(LIPREAC$, J&)
    'Call CIERRARCH("*.*")
    Call CARLISTA
    Screen.MousePointer = 1
    '
    '//OT-07-0434//
    Call COMUNI("Ordenamiento Completo.")
    '
End Sub

Sub CARILISTAS()
 '
   '//OT-07-0434//
   List1.Clear
 
   APERBASDAT ("PEDCLI")
   On Error Resume Next
   With ListaPrec
     .MoveLast
     If Err > 0 Then
        Exit Sub
     End If
     .MoveFirst
     For U& = 1 To .RecordCount
         CodList$ = .Fields("Cod_Lista")
         DenomLista$ = .Fields("denom_Lista")
         List1.AddItem AJUSTD$(TRIM$(CodList$), 2) + "-." + TRIM$(DenomLista$)
         .MoveNext
     Next U&
   End With
   List1.Refresh
   '//OT-07-0434//
End Sub

Function VERILLAVE%()
   '
   VERILLAVE% = 0
10 OPX% = COMALTER%("Coloque Diskette Llave\\Continuar\Cancelar")
   If OPX% = 1 Then
      On Error GoTo 50
      SVL$ = SERIVOL$("A:")
      TECONTROL$ = ENCRIPTA$(TRIM$(UCase$(App.EXEName)) + SVL$)
      NX1% = FreeFile
      Dim TTX As String * 128
      Open "A:\" + TRIM$(UCase$(App.EXEName)) + ".KEY" For Random Access Read Write As #NX1% Len = 128
      Get #NX1%, 2, TTX$
      LTC% = Len(TECONTROL$)
VERILLAVE% = 1: GoTo 99
      If CVI(Mid$(TTX$, 18, 2)) = LTC% Then
        If Mid$(TTX$, 30, LTC%) = TECONTROL$ Then
          VERILLAVE% = 1
          GoTo 99
        End If
      End If
      Call COMUNI("Llave no Válida")
   End If
   GoTo 99
   '
50 Resume 51
51 GoTo 10
   '
99 End Function

Sub EDITLISTA(ARCHIPRE$)
   '
   'ARMA EL ENCABEZADO DE VENTABU
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=" + TRIM$(List1.TEXT)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
   '
   'ABRE EL VENTABU
   XII% = VENTABU%("LIPREVEN" + ATRI$)
   If XII% >= 0 Then
     CADECI% = Val(Text3)
     If CADECI% < 0 Or CADECI% > 5 Then CADECI% = 2
     List2.Clear
     '
     FIN& = ULTREG&("!LIPREVEN")
     Screen.MousePointer = 11
     '
     ' SI  LABEL 14 (FORMATO) ES IGUAL A 3 ,SI NO GRABA DIRECTAMENTE
     'preguntar exactamente que hace
     HOII% = HOY(HOS$)
     If Val(Label4) = 3 Then
       For i& = 1 To FIN&
         NXX$ = REGLEIDO$("!LIPREVEN", i&)
         CI% = CVI(Left$(NXX$, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             PRELI# = CVD(Mid$(NXX$, 62, 8))
             RFF$ = RECFILT$(ARCHIPRE$, 1, MKI$(CI%))
             If Len(RFF$) >= 4 Then
               REANT& = CVS(RFF$)
               If REANT& > 0 Then
                 NYY$ = REGLEIDO$(ARCHIPRE$, REANT&)
                 If Abs(CVD(Mid$(NYY$, 62, 8)) - PRELI#) < 0.00005 Then
                   If TRIM$(Mid$(NYY$, 46, 16)) <> "" Then
                     GoTo 89
                   End If
                 End If
               End If
             End If
           End If
         End If
         Call REPLA(NXX$, HOS$, 46, 16)
         Call GRAREG("!LIPREVEN", NXX$, i&)
89     Next i&
     End If
     '
     'RECORRE EL LIPREVEN Y GRABA EN ARCHIPRE$
     'VER SI SE VA A HACER UN DELETE ANTES DE GRABAR
     'O SI SE VA USAR OTRO METODO COMO CREAR OTRO CAMPO
     'DE ESTADO , MARCAR Y LUEGO BORRAR
     'BORRA ANTES DE INGRESAR LOS DATOS
     SQLX$ = "DELETE * FROM Preventa "
     SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTACTIVA& & " ;"
     Ventas.Execute (SQLX$)

     For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        NXX$ = REGLEIDO$("!LIPREVEN", i&)
        CI% = CVI(Left$(NXX$, 2))
        MRX$ = Mid$(NXX$, 45, 1)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            PRELI# = CVD(Mid$(NXX$, 62, 8))
            PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
            'DE$ = DESCRIT0$(CI%)
            'Call REPLA$(NXX$, DE$, 3, 42)
            'Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
            'Call REPLA(NXX$, MRX$, 44, 1)
            'INGRESA LOS DATOS
          With Preventa
            .AddNew
            !Cod_Lista = LISTACTIVA&
            !Cod_Pieza = CI%
            !Num_Orden = i&
            !Desc_Pieza = DESCRIT0$(CI%)
            !Denom_Pieza = DESCRIT0$(CI%)
            !Carac_Control = Mid$(NXX$, 45, 1)
            !Num_Original = (Mid$(NXX$, 46, 16))
            !Precio_Unitario = CVD(Mid$(NXX$, 62, 8))
            !Cant_Limite = CVD(Mid$(NXX$, 70, 2))
            !Porc_Limite = (Mid$(NXX$, 72, 1))
            .Update
          End With
          'Call GRAREG(ARCHIPRE$, NXX$, I&)
          End If
        End If
     Next i&
     'Call SETULTREG(ARCHIPRE$, FIN&)
     'Call CIERRARCH(ARCHIPRE$)
     Screen.MousePointer = 1
     '
     Call CARLISTA
   End If

End Sub
Sub LeeLaLista(n%)
   '//OT-07-0434//
   '
   Call SETULTREG("!LIPREVEN", 0)
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT * FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & n% & " "
   SQLX$ = SQLX$ + "ORDER BY Num_Orden;"
   Set ARTSELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With ARTSELEC
     .MoveLast
     If Err > 0 Then
        Exit Sub
     End If
     FIN& = .RecordCount
     .MoveFirst
     If FIN& > 0 Then
        Do While (.EOF = False)
            TTXX$ = Space$(128)
            CII% = !Cod_Pieza
            Call REPLA(TTXX$, MKI$(CII%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CII%), 3, 42)
            Call REPLA(TTXX$, !Carac_Control, 45, 1)
            Call REPLA(TTXX$, !Num_Original, 46, 16)
               PRUNI# = !Precio_Unitario
            Call REPLA(TTXX$, MKD$(PRUNI#), 62, 8)
            Call REPLA(TTXX$, MKI$(!Cant_Limite), 70, 2)
            Call REPLA(TTXX$, MKI$(!Porc_Limite), 72, 1)
            JJ& = JJ& + 1
            '****
            Call GRAREG("!LIPREVEN", TTXX$, JJ&)
            

        .MoveNext
        Loop
        Call SETULTREG("!LIPREVEN", JJ&)
        Call CIERRARCH("!LIPREVEN")
      End If
     End With
     '//OT-07-0434//
  
End Sub


Private Sub newlist_Click()
    Call Command7_Click
End Sub


Private Sub reunilistadeprecios_Click()
    Call OPLPREAC.Command1_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub



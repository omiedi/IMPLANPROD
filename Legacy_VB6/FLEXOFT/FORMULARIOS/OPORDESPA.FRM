VERSION 5.00
Begin VB.Form OPORDESPA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Pedidos de Clientes y Ordenes de Despacho"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
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
   ScaleHeight     =   5190
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "Cotiz."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7630
      Picture         =   "OPORDESPA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Parametros de Cálculo - Tipo de Cambio Dolar"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valorizar"
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
      Left            =   5565
      TabIndex        =   39
      Top             =   210
      Width           =   1965
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Muestra Pre- cios e Importes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   210
         TabIndex        =   40
         Top             =   315
         Width           =   1695
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
      Height          =   855
      Left            =   9975
      Picture         =   "OPORDESPA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   9200
      Picture         =   "OPORDESPA.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   750
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
      Left            =   8410
      Picture         =   "OPORDESPA.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pedido de Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Left            =   210
      TabIndex        =   17
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ventana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   3190
         TabIndex        =   35
         Top             =   2520
         Width           =   1755
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
            Left            =   140
            MaxLength       =   8
            TabIndex        =   36
            ToolTipText     =   "Introduzca el Nombre del Controlador 'PRF'"
            Top             =   420
            Width           =   1460
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   33
         Top             =   3675
         Width           =   1755
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   315
            TabIndex        =   34
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Crédito en Cuenta Corriente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2140
         TabIndex        =   31
         Top             =   3675
         Width           =   2805
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Verificar Límite de Crédito Asignado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   525
            TabIndex        =   32
            Top             =   315
            Width           =   2010
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato de Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   28
         Top             =   2520
         Width           =   2835
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Normal"
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
            Left            =   525
            TabIndex        =   30
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Itemizado (Conjuntos)"
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
            Left            =   525
            TabIndex        =   29
            Top             =   630
            Width           =   2220
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   25
         Top             =   1575
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
            Height          =   360
            Left            =   190
            TabIndex        =   26
            Text            =   "PC"
            Top             =   275
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
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
         Top             =   1575
         Width           =   3750
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
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
            Left            =   2205
            TabIndex        =   24
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
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
            Left            =   525
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPORDESPA.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   510
         Width           =   855
      End
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
         Height          =   360
         Left            =   210
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
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
         Height          =   360
         Left            =   3780
         TabIndex        =   19
         Top             =   525
         Width           =   175
      End
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
         Left            =   2415
         TabIndex        =   18
         Text            =   " "
         Top             =   525
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
         TabIndex        =   27
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   5565
      TabIndex        =   3
      Top             =   1350
      Width           =   5160
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2415
         TabIndex        =   41
         Top             =   1365
         Width           =   1545
         Begin VB.CheckBox Check6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Captura"
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
            ToolTipText     =   "Captura Cola de Impresión de Ordenes de Despacho y/o Pedidos."
            Top             =   600
            Width           =   1170
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   42
            ToolTipText     =   "Suprime Impresión - Enía Datos a Cola de Impresión."
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Generación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2730
         TabIndex        =   37
         Top             =   2415
         Width           =   2175
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Despachar Según Disponibilidad de Stocks"
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
            Left            =   210
            TabIndex        =   38
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   210
         TabIndex        =   14
         Top             =   2415
         Width           =   2325
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   1680
            TabIndex        =   15
            Text            =   " "
            Top             =   550
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Anticipo de Activación de Despacho (dias):"
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
            Left            =   -105
            TabIndex        =   16
            Top             =   315
            Width           =   1665
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
         Height          =   360
         Left            =   3780
         TabIndex        =   12
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
         Height          =   360
         Left            =   210
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
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
         Picture         =   "OPORDESPA.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   315
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
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
         Height          =   960
         Left            =   210
         TabIndex        =   6
         Top             =   1365
         Width           =   2070
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   8
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
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
            Left            =   420
            TabIndex        =   7
            Top             =   580
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
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
         Height          =   960
         Left            =   4095
         TabIndex        =   4
         Top             =   1365
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
            Left            =   210
            TabIndex        =   5
            Text            =   "OD"
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
         TabIndex        =   9
         Top             =   420
         Width           =   2325
      End
   End
End
Attribute VB_Name = "OPORDESPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORDESPA", Text1.TEXT)
    Call GRACONTROL("", "CODDESPA", Text5.TEXT)
    Call GRACONTROL("", "FORPECLI", Text6.TEXT)
    Call GRACONTROL("", "CODPECLI", Text8.TEXT)
    Call GRACONTROL("", "VENCARPE", Text4.TEXT)
    '
    MONUME$ = "AUTO"
    If Option1.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de O/D's.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMDESPA", MONUME$)
    '
    MONUME$ = TRIM$(Text3.TEXT)
    Call GRACONTROL$("", "ANTIDESP", MONUME$)
    '
    MONUME$ = "AUTO"
    If Option3.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de Pedidos de Clientes.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMPECLI", MONUME$)
    '
    MONUME$ = "NORMAL": FORCAPED% = 1
    If Option5.Value = True Then
        MONUME$ = "ITEMIZADO": FORCAPED% = 2
    End If
    Call GRACONTROL("ADMIPED", "FORMACAR", MONUME$)
    '
    MONUME$ = "NO": VERLICRE% = 0
    If Check1.Value = 1 Then
        MONUME$ = "SI": VERLICRE% = 1
    End If
    Call GRACONTROL("ADMIPED", "VERLICRE", MONUME$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("PEDCLIEN", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO"
    If Check5.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("ORDESPA", "SUPRINT", SUPRINT$)
    '
    CAPRINT$ = "NO"
    If Check6.Value = 1 Then CAPRINT$ = "SI"
    Call GRACONTROL("CAPRINT", IDENTER$, CAPRINT$)
    '
    CMO$ = ""
    If Check3.Value = 1 Then CMO$ = "REVISTOCK"
    Call GRACONTROL$("ADMIPED", "GENDESPA", CMO$)
    '
    CMO$ = "NO"
    If Check4.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL$("ADMIPED", "MUEPREC", CMO$)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
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
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") = 2 Then
     Call CARTAMONED
   End If
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FORDESPA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = CONTROL$("", "FORPECLI")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
    VENCARPE$ = CONTROL$("", "VENCARPE")
    If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
    Text4.TEXT = VENCARPE$
    '
    If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then Check1.Value = 1
    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
       Check5.Value = 1
    End If
    '
    If IDENTER$ <> "" Then
      If CONTROL$("CAPRINT", IDENTER$) = "SI" Then
        Check6.Value = 1
      End If
    End If
    '
    If CONTROL$("", "NUMDESPA") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "NUMPECLI") = "MANUAL" Then Option3.Value = True
    If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then Option5.Value = True
    Text3.TEXT = TRIM$(CONTROL$("", "ANTIDESP"))
    If Val(Text3.TEXT) < 1 Then Text3.TEXT = "3"
    '
    CMO$ = TRIM$(CONTROL$("", "CODDESPA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODPECLI")): If CMO$ <> "" Then Text8.TEXT = CMO$
    '
    If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then Check3.Value = 1
    If CONTROL$("ADMIPED", "MUEPREC") = "SI" Then Check4.Value = 1
    '
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

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub


Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
End Sub



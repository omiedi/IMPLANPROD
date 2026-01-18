VERSION 5.00
Begin VB.Form FORMULA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulas y Estructuras de Producto"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   315
      ScaleHeight     =   75
      ScaleWidth      =   1110
      TabIndex        =   39
      Top             =   1365
      Visible         =   0   'False
      Width           =   1170
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4410
      Picture         =   "FORMULA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Fórmula de Conjunto Anterior"
      Top             =   200
      Width           =   750
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Más ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Control de Recursividad y Repetición de Fórmulas"
      Top             =   6090
      Width           =   855
   End
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
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   3570
      Width           =   855
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   945
      TabIndex        =   35
      Top             =   2625
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.CommandButton Command21 
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
      Height          =   390
      Left            =   3990
      TabIndex        =   33
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command6 
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
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Copia Fórmulas de un Producto a Otro"
      Top             =   4410
      Width           =   855
   End
   Begin VB.CommandButton Command5 
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
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Emite Listados por Impresora"
      Top             =   5250
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Actualizar Fórmula / Estructura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   210
      TabIndex        =   14
      Top             =   5040
      Width           =   8100
      Begin VB.CommandButton Command8 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7035
         Picture         =   "FORMULA.frx":0F12
         TabIndex        =   30
         ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
         Top             =   1365
         Width           =   855
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
         Height          =   360
         Left            =   3045
         TabIndex        =   34
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7035
         Picture         =   "FORMULA.frx":121C
         TabIndex        =   29
         ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
         Top             =   1050
         Width           =   855
      End
      Begin VB.TextBox Text15 
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
         Left            =   6510
         TabIndex        =   19
         Text            =   " "
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7035
         Picture         =   "FORMULA.frx":1526
         TabIndex        =   21
         ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
         Top             =   735
         Width           =   855
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
         Height          =   570
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "FORMULA.frx":1830
         Top             =   1050
         Width           =   5580
      End
      Begin VB.TextBox Text13 
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
         Left            =   3885
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   435
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
         Height          =   360
         Left            =   1260
         TabIndex        =   15
         Top             =   420
         Width           =   1800
      End
      Begin VB.TextBox Text14 
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
         Left            =   4935
         TabIndex        =   18
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Append"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7035
         Picture         =   "FORMULA.frx":1832
         TabIndex        =   20
         ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "D:"
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
         Left            =   5880
         TabIndex        =   28
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         TabIndex        =   27
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label8 
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
         Left            =   2940
         TabIndex        =   26
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   210
         TabIndex        =   23
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Uso:"
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
         Left            =   3990
         TabIndex        =   22
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3000
      Left            =   210
      TabIndex        =   12
      Top             =   1845
      Width           =   8100
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2670
         Left            =   0
         TabIndex        =   13
         Top             =   315
         Width           =   8100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   360
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   9000
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            Caption         =   " Codigo                   Descripcion                                       U-M              Uso     D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   45
            Width           =   8000
         End
      End
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6720
      TabIndex        =   4
      Text            =   " "
      Top             =   210
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   0
      Top             =   210
      Width           =   2325
   End
   Begin VB.TextBox TEXT3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7665
      TabIndex        =   3
      Text            =   " "
      Top             =   1100
      Width           =   645
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
      Height          =   540
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   945
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":1E46
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":1F90
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Restaura Valores Anteriores"
      Top             =   1890
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "FORMULA.frx":229A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Basica:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5145
      TabIndex        =   11
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Producto o Conjunto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   10
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "       Unidad de Medida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6090
      TabIndex        =   9
      Top             =   1065
      Width           =   1485
   End
   Begin VB.Label Label4 
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
      TabIndex        =   8
      Top             =   945
      Width           =   1485
   End
End
Attribute VB_Name = "FORMULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI%, CABAS, CODE$, CIE%
Dim HIDOCLICK As Integer
'

Private Sub Command1_Click()
    ILISTA& = List1.ListIndex
    CAIT% = List1.ListCount
    CODICABE% = CODINT%(FORMULA.Text1.Text)
    If CODICABE% < 1 Then Exit Sub
    '
    Dim EXPLOLI1$(64)
    CAIT1% = CAIT%
    For JJI% = 1 To 64
       EXPLOLI1$(JJI%) = ""
       If JJI% <= List1.ListCount Then
          EXPLOLI1$(JJI%) = List1.List(JJI% - 1)
       End If
    Next JJI%
    '
    For I& = 1 To CAIT%
        List1.ListIndex = I& - 1
        EXPLOLIN(I&) = List1.Text
        '
        CICOMPO% = CODINT%(Left$(EXPLOLIN(I&), 16))
        Call EXPLOMUL(CICOMPO%)
        For JJI% = 1 To CONTAMUL%
          If CIK%(JJI%) = CODICABE% Then
            Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(CODICABE%)) + "\" + DESCRIT0$(CODICABE%) + ".")
            Exit Sub
          End If
        Next JJI%
    Next I&
    '
    CAIT% = CAIT1%
    For JJI% = 0 To 64
       EXPLOLIN$(JJI%) = EXPLOLI1$(JJI%)
    Next JJI%
    '
    COD$ = AJUSTI$(FORMULA.Text1.Text, 16)
    CABAS = Val(Text4.Text)
    CAIT% = List1.ListCount
    '
    If CABAS > 0.01 Then
        Call EXPLOGR(COD$, CABAS)
        MODIFI% = 0
    End If
    '
    List1.ListIndex = ILISTA&
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    OPXX% = ALTERNA%("Fórmulas de Desarmado\Repetición en Despieces")
    If OPXX% = 1 Then
         FORMUDES.Show 1
       ElseIf OPXX% = 2 Then
         Call REDESPI
       'ElseIf OPXX% = 3 Then
       '  Call RECURSI
    End If
    Command10.Enabled = True
End Sub

Private Sub COMMAND11_CLICK()
    CODXX$ = TRIM$(Text11.Text)
    If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
        Call MENSERR(24, "Código no Válido")
        Text11.SetFocus
        Exit Sub
    End If
    '
    CIY% = CODINT%(CODXX$)
    DEXX$ = DESCRIT$(CIY%)
    UNXX$ = UNIMED$(CIY%)
    USO = Val(Text14.Text)
    If Abs(USO) < 0.0001 Then
        Call MENSERR(24, "Uso no Válido")
        Text14.SetFocus
        Exit Sub
    End If
    USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
    PD = Int(Val(Text15.Text))
    If PD < 1 Then PD = 0
    If PD > 9 Then PD = 9
    PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
    '
    XXX$ = Space$(64)
    Call REPLA(XXX$, CODXX$, 1, 16)
    Call REPLA(XXX$, DEXX$, 17, 28)
    Call REPLA(XXX$, UNXX$, 47, 2)
    Call REPLA(XXX$, USOX$, 49, 12)
    Call REPLA(XXX$, PDX$, 61, 4)
    CAIT% = CAIT% + 1
    EXPLOLIN$(CAIT%) = XXX$
    '
    List1.AddItem EXPLOLIN$(CAIT%)
    List1.ListIndex = List1.ListCount - 1
    '
    Text11.Text = ""
    CODE$ = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text11.SetFocus
    MODIFI% = 1
    '
    Command4.Enabled = True
    Command8.Enabled = True
    '
End Sub

Private Sub Command12_Click()
    If List1.ListIndex < 0 Then
        Call MENSERR(24, "Opción Inserción no Válida")
        Exit Sub
    End If
    '
    CODXX$ = TRIM$(Text11.Text)
    If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
        Call MENSERR(24, "Código no Válido")
        Text11.SetFocus
        Exit Sub
    End If
    '
    CIY% = CODINT%(CODXX$)
    DEXX$ = DESCRIT$(CIY%)
    UNXX$ = UNIMED$(CIY%)
    USO = Val(Text14.Text)
    If Abs(USO) < 0.0001 Then
        Call MENSERR(24, "Uso no Válido")
        Text14.SetFocus
        Exit Sub
    End If
    USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
    PD = Int(Val(Text15.Text))
    If PD < 1 Then PD = 0
    If PD > 9 Then PD = 9
    PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
    '
    XXX$ = Space$(64)
    Call REPLA(XXX$, CODXX$, 1, 16)
    Call REPLA(XXX$, DEXX$, 17, 28)
    Call REPLA(XXX$, UNXX$, 47, 2)
    Call REPLA(XXX$, USOX$, 49, 12)
    Call REPLA(XXX$, PDX$, 61, 4)
    '
    If List1.ListIndex < 0 Then
        If List1.ListCount > 0 Then
            List1.ListIndex = 0
        End If
    End If
    '
    ILISTA& = List1.ListIndex
    List1.AddItem XXX$, ILISTA&
    If ILISTA& < 64 Then EXPLOLIN$(ILISTA& + 1) = XXX$
    List1.ListIndex = ILISTA&
    '
    Text11.Text = ""
    CODE$ = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text11.SetFocus
    MODIFI% = 1
    '
    Command4.Enabled = True
    Command8.Enabled = True
    '
End Sub

Private Sub Command13_Click()
    If IUCOD% > 1 Then
        IUCOD% = IUCOD% - 1
        SUBENI% = 1
        Text1.Text = UCOD$(IUCOD%)
        Text1.SetFocus
    End If
End Sub

Private Sub Command14_Click()
End Sub

Private Sub COMMAND4_CLICK()
    '
    If List1.ListIndex < 0 Then
        Call MENSERR(24, "Opción Borrado no Válida")
        Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    MODIFI% = 1
    '
End Sub

Private Sub Command5_Click()
    MODO% = COMALTER%("Elija el Formato de Impresión\de los Listados Solicitados:\\Estructuras y Enlaces\Listas de Materiales\Listas de Compras\Cancelar")
    On MODO% GoTo 100, 200, 300
    Exit Sub
    '
100 Call LIESTRU
    Exit Sub
    '
200 Call LIMULTI
    Exit Sub
    '
300 Call LIMULIN
    Exit Sub
    '
End Sub
'

Private Sub Command6_Click()
    '
    Text1.Text = ""
    Call COPYEXPLO
    '
End Sub
'

Private Sub Command7_Click()
    '
    CODCABE$ = Text1.Text
    Call NEWITEM
    Call FRESHECHO("MASTER")
    Text1.Text = CODCABE$
    '
End Sub

Private Sub Command8_Click()
    '
    If List1.ListIndex < 0 Then
        Call MENSERR(24, "Opción Edición no Válida")
        Exit Sub
    End If
    '
    Text11.Text = Left$(List1.Text, 16)
    CDYY% = CODINT%(Text11.Text)
    Text12.Text = DESCRIT$(CDYY%)
    Text13.Text = UNIMED$(CDYY%)
    Text14.Text = TRIM$(Mid$(List1.Text, 49, 12))
    Text15.Text = TRIM$(Mid$(List1.Text, 61, 4))
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command4.Enabled = False
    Command8.Enabled = False
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    Text14.SetFocus
    MODIFI% = 1
    '
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
    COMMA$ = Command$
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    If COMMA$ <> "" Then
        Text1.Text = COMMA$
    End If
    Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    INI& = 1 + ULTREG&("ESTRUDOS")
    FIN& = ULTREG&("ESTRUNUE")
    For U& = INI& To FIN&
        RB$ = String$(256, 0)
        Mid$(RB$, 251, 4) = MKS$(0.0001)
        Call GRAREG("ESTRUDOS", RB$, U&)
    Next U&
    '
    Call CIERRARCH("*.*")
    Call BKSTRUC
    Call FINAL("")
End Sub

Private Sub List1_DblClick()
    Text1.Text = Left$(List1.Text, 16)
    Text4.Text = TRIM$(Mid$(List1.Text, 49, 12))
    Text1.SetFocus
End Sub

Private Sub Text1_Click()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Text4.Text = 1
End Sub

Private Sub Text1_Change()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    MODIFI% = 0
    Text2.Text = ""
    TEXT3.Text = ""
    Text4.Text = ""
    List1.Clear
End Sub

Private Sub Text1_DBLCLICK()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Text1.Text = ""
    Call SELECHO("MASTER")
End Sub
'
Private Sub COMMAND21_CLICK()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Text1.Text = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub
'

Private Sub Text1_GotFocus()
    Screen.MousePointer = 11
    FORMULA.Refresh
    If Text1.Text = "" Then
        Text1.Text = VALACT1$("MASTER")
    End If
    If DESARMA% = 1 Then
       Text1.Text = ""
       Text2.Text = ""
       TEXT3.Text = ""
       Screen.MousePointer = 1
       DESARMA% = 0
       Exit Sub
    End If
    CODX$ = TRIM$(Text1.Text)
    cix% = CODINT%(CODX$)
    If cix% > 0 Then
         COD$ = CODX$: CI% = cix%
         Text1.Text = CODX$
         Text2.Text = DESCRIT$(cix%)
         TEXT3.Text = UNIMED$(CI%)
       Else
         Text1.Text = COD$
    End If
    '
    If TRIM$(Text4.Text) = "" Then Text4.Text = 1
    '
    If TRIM$(COD$) <> "" Then
        CABAS = Val(Text4.Text)
        '
        If ECOARCH$("FORMUDES", MKI$(CODINT%(COD$))) <> "" Then
           FORMUDES.Text1.Text = COD$
           'Text1.Text = ""
           COD$ = ""
           Screen.MousePointer = 1
           FORMUDES.Show 1
           Text1.Text = ""
           Text1.SetFocus
           Screen.MousePointer = 1
           Exit Sub
        End If
        '
        Call LEEEXPLO(COD$, CABAS)
        '
        If SUBENI% = 0 Then
          If IUCOD% < 128 Then
            If COD$ <> UCOD$(IUCOD%) Then
              IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = COD$
            End If
          End If
        End If
        SUBENI% = 0
        '
        List1.Clear
        For I% = 1 To CAIT%
          List1.AddItem EXPLOLIN$(I%)
        Next I%
        If List1.ListCount > 0 Then
            List1.ListIndex = 0
        End If
        MODIFI% = 0
    End If
    Screen.MousePointer = 1
End Sub
'

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CODX$ = TRIM$(Text1.Text)
        cix% = CODINT%(CODX$)
        If cix% > 0 Then
            COD$ = CODX$: CI% = cix%
            Text1.Text = CODX$
            COD$ = CODX$: CI% = cix%
            Text2.Text = DESCRIT$(CI%)
            TEXT3.Text = UNIMED$(CI%)
            Text4.SetFocus
          Else
            Call MENSERR(24, "Código no Válido")
            Text1.Text = COD$
        End If
    End If
End Sub
'

Private Sub Text14_GotFocus()
    If TRIM$(Text11.Text) = "" Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If Abs(Val(Text14.Text)) > 0 Then
        If KeyAscii = 13 Then
            Text15.SetFocus
        End If
    End If
End Sub

Private Sub Text15_GotFocus()
    If TRIM$(Text11.Text) = "" Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command11.SetFocus
    End If
End Sub

Private Sub Text2_GotFocus()
    TEXT3.SetFocus
End Sub
'

Private Sub Text3_GotFocus()
    CODX$ = TRIM$(Text1.Text)
    cix% = CODINT%(CODX$)
    If cix% > 0 Then
          COD$ = CODX$: CI% = cix%
          Text2.Text = DESCRIT$(CI%)
          TEXT3.Text = UNIMED(CI%)
          Text4.SetFocus
        Else
          Call MENSERR(24, "Código no Válido")
          Text1.Text = COD$
          Text1.SetFocus
    End If
End Sub
'

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TRIM$(TEXT3.Text) <> "" Then
            Command1.SetFocus
        End If
    End If
End Sub
'


Private Sub Command2_Click()
    Call HELPONLINE("FORMULAS")
End Sub
'

Private Sub Command3_Click()
    Text1.Text = COD$
    TEXT3.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command9_Click()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    Unload FORMULA
    '
End Sub

Private Sub Text4_Change()
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    MODIFI% = 0
    If TRIM$(COD$) <> "" Then
        CABAS = Val(Text4.Text)
        '
        If ECOARCH$("FORMUDES", MKI$(CODINT%(COD$))) <> "" Then
           FORMUDES.Text1.Text = COD$
           ' Text1.Text = ""
           COD$ = ""
           Screen.MousePointer = 1
           FORMUDES.Show 1
           Text1.Text = ""
           Text1.SetFocus
           Screen.MousePointer = 1
           Exit Sub
        End If
        '
        Call LEEEXPLO(COD$, CABAS)
        '
        If SUBENI% = 0 Then
          If IUCOD% < 128 Then
            If COD$ <> UCOD$(IUCOD%) Then
              IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = COD$
            End If
          End If
        End If
        SUBENI% = 0
        '
        List1.Clear
        For I% = 1 To CAIT%
          List1.AddItem EXPLOLIN$(I%)
        Next I%
    End If
End Sub

Private Sub Text4_GOTFOCUS()
    If Val(Text4.Text) = 0 Then
        Text4.Text = "1"
    End If
End Sub

Private Sub Text11_Click()
    Text14.Text = Str$(CABAS)
End Sub

Private Sub Text11_DBLCLICK()
    Text11.Text = ""
    HIDOCLICK% = 1
    Call SELECHO("MASTER")
End Sub
'
Private Sub COMMAND31_CLICK()
    Text11.Text = ""
    HIDOCLICK% = 1
    Call SELECHO("MASTER")
    Text11.SetFocus
End Sub
'

Private Sub Text11_GotFocus()
    If TRIM$(Text1.Text) = "" Then
        Text1.SetFocus
    End If
    '
    If Text11.Text = "" Then
        If HIDOCLICK% = 1 Then
            Text11.Text = VALACT1$("MASTER")
            HIDOCLICK% = 0
        End If
    End If
    '
    If TRIM$(Text11.Text) = TRIM$(Text1.Text) Then
        CODE$ = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
    End If
    CODY$ = TRIM$(Text11.Text)
    CIY% = CODINT%(CODY$)
    If CIY% > 0 Then
         CODE$ = CODY$: CIE% = CIY%
         Text11.Text = CODY$
         Text12.Text = DESCRIT$(CIY%)
         Text13.Text = UNIMED$(CIY%)
       Else
         Text11.Text = CODE$
    End If
    '
    If TRIM$(Text14.Text) = "" Then Text14.Text = Str$(CABAS)
    '
End Sub
'

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CODY$ = TRIM$(Text11.Text)
        CIY% = CODINT%(CODY$)
        If CIY% > 0 Then
            CODE$ = CODY$: CIE% = CIY%
            Text11.Text = CODY$
            CODE$ = CODY$: CIE% = CIY%
            Text12.Text = DESCRIT$(CIE%)
            Text13.Text = UNIMED$(CIE%)
            Text14.SetFocus
          Else
            Call MENSERR(24, "Código no Válido")
            Text11.Text = CODE$
        End If
    End If
End Sub
'

Private Sub Text12_GotFocus()
    Text13.SetFocus
End Sub
'

Private Sub Text13_GotFocus()
    CODY$ = TRIM$(Text11.Text)
    CIY% = CODINT%(CODY$)
    If CIY% > 0 Then
          CODE$ = CODY$: CIE% = CIY%
          Text2.Text = DESCRIT$(CIY%)
          TEXT3.Text = UNIMED(CIY%)
          Text14.SetFocus
        Else
          Call MENSERR(24, "Código no Válido")
          Text11.Text = CODE$
          Text11.SetFocus
    End If
End Sub
'
Sub REDESPI()
    '
    CAERR% = 0
    FIN& = ULTREG&("ESTRUNUE")
    FINDOS& = ULTREG&("ESTRUDOS")
    List2.Clear
    List2.Visible = True
    MARCOBARRA.Visible = True
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      R0$ = REGLEIDO$("ESTRUNUE", I&)
      R1$ = Left$(R0$, 2)
      DPS$ = Mid$(R0$, 3, 120)
      If I& <= FINDOS& Then
        DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", I&)
        CI1% = CVI(R1$)
        If CI1% > 0 Then
          If CI1% <= NUMAS% Then
            A$ = CODEXT$(CI1%) + " " + DESCRIT$(CI1%)
            List2.AddItem A$
            If List2.TopIndex < List2.ListCount - 6 Then
               List2.TopIndex = List2.ListCount - 6
            End If
            List2.Refresh
            '
            J% = 0
            For K% = 1 To Len(DPS$) Step 8
              DPSX$ = Mid$(DPS$, K%, 8)
              CI2% = CVI(Left$(DPSX$, 2))
              If CI2% > 0 Then
                If CI2% <= NUMAS% Then
                  J% = J% + 1
                  CII1%(J%) = CI2%
                End If
              End If
            Next K%

            For K% = 1 To J%
              If CII1%(K%) = CI1% Then
                MENSA$ = "Recursividad en Código " + TRIM$(CODEXT$(CI1%)) + "\" + DESCRIT0$(CI1%)
                If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
              End If
              '
              For K1% = K% + 1 To J%
                If CII1%(K1%) = CII1%(K%) Then
                  CAERR% = CAERR% + 1
                  cod1$ = TRIM$(CODEXT$(CI1%))
                  cod2$ = TRIM$(CODEXT$(CII1%(K%)))
                  If CODINT%(cod1$) = CI1% Then ' SUPRIME ITEMS DE BAJA
                    MENSA$ = "Repeticion del Codigo " + cod2$ + "\en Fórmula de Codigo " + cod1$
                    If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
                  End If
                  GoTo 89
                End If
              Next K1%
            Next K%
          End If
        End If
      End If
89  Next I&
    '
    If CAERR% < 1 Then
        Call COMUNI("Control O.K. - No se Detectaron Repeticiones")
      Else
        Call MENSERR(20, "Se Encontraron" + Str$(CAERR%) + " Repeticiones en Despieces")
    End If
    '
99  List2.Clear
    List2.Visible = False
    MARCOBARRA.Visible = False
    Call CIERRARCH("*.*")
    '
End Sub



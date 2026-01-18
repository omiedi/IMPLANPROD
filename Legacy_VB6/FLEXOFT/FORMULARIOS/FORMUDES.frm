VERSION 5.00
Begin VB.Form FORMUDES 
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulas de Descomposicion"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8505
      Picture         =   "FORMUDES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Impresión - Vuelco de Pantalla"
      Top             =   4490
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8505
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   31
      Top             =   5565
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Materias Primas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   25
      Top             =   945
      Width           =   8100
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
         Height          =   345
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   5
         Top             =   315
         Width           =   1170
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
         Height          =   345
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   3
         Top             =   315
         Width           =   2010
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
         Height          =   345
         Left            =   3255
         TabIndex        =   4
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1260
         TabIndex        =   30
         Top             =   840
         Width           =   6630
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
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
         Left            =   5775
         TabIndex        =   29
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4620
         TabIndex        =   28
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  U-M:"
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
         Left            =   3675
         TabIndex        =   27
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
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
         Left            =   -315
         TabIndex        =   26
         Top             =   420
         Width           =   1485
      End
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
      Left            =   8505
      Picture         =   "FORMUDES.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   2835
      Width           =   855
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
      Height          =   345
      Left            =   3150
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actualizar Productos"
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
      TabIndex        =   20
      Top             =   5460
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
         Picture         =   "FORMUDES.frx":07B4
         TabIndex        =   12
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
         TabIndex        =   7
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
         Picture         =   "FORMUDES.frx":0ABE
         TabIndex        =   11
         ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
         Top             =   1050
         Width           =   855
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
         Picture         =   "FORMUDES.frx":0DC8
         TabIndex        =   10
         ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
         Top             =   735
         Width           =   855
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
         TabIndex        =   6
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
         Left            =   5670
         TabIndex        =   8
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
         Picture         =   "FORMUDES.frx":10D2
         TabIndex        =   9
         ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   590
         Left            =   1260
         TabIndex        =   35
         Top             =   1050
         Width           =   5580
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3990
         TabIndex        =   34
         Top             =   420
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
         TabIndex        =   24
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
         Left            =   3045
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   4725
         TabIndex        =   21
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   210
      TabIndex        =   18
      Top             =   2520
      Width           =   8100
      Begin VB.ListBox List1 
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
         Left            =   210
         TabIndex        =   19
         Top             =   680
         Width           =   7680
      End
      Begin VB.Line Line1 
         X1              =   7875
         X2              =   7875
         Y1              =   360
         Y2              =   730
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo                   Descripcion                                       U-M             Cant"
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
         TabIndex        =   33
         Top             =   360
         Width           =   7680
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
      Height          =   345
      Left            =   1155
      MaxLength       =   15
      TabIndex        =   0
      Top             =   315
      Width           =   2010
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
      Left            =   3570
      TabIndex        =   2
      Top             =   315
      Width           =   4740
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
      Left            =   8505
      Picture         =   "FORMUDES.frx":13DC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1995
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
      Left            =   8505
      Picture         =   "FORMUDES.frx":16E6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   315
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
      Left            =   8505
      Picture         =   "FORMUDES.frx":1830
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   555
      Left            =   8295
      Picture         =   "FORMUDES.frx":1B3A
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   4725
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Código:"
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
      Left            =   -210
      TabIndex        =   17
      Top             =   420
      Width           =   1275
   End
End
Attribute VB_Name = "FORMUDES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
   COD$ = ""
   DESARMA% = 1
End Sub

Private Sub Text1_Change()
   CODFOR$ = Text1.Text
   CODINTFOR% = Abs(CODINT%(CODFOR$))
   If CODINTFOR% > 0 Then
      Text2.Text = DESCRIT0$(CODINTFOR%)
      Call CARFORDES
   End If
End Sub

Private Sub Text1_DBLCLICK()
    Call SELECHO("FORMUDES")
    CODFOR$ = VALACT1$("FORMUDES")
    Text1.Text = CODFOR$
End Sub
'
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CODFOR$ = TRIM$(Text1.Text)
        CODINTFOR% = Abs(CODINT%(CODFOR$))
        If CODINTFOR% > 0 Then
           Text2.Text = DESCRIT0$(CODINTFOR%)
        End If
    End If
End Sub
'
Private Sub COMMAND21_CLICK()
    Call SELECHO("FORMUDES")
    CODFOR$ = VALACT1$("FORMUDES")
    Text1.Text = CODFOR$
End Sub
'
Private Sub Text4_Change()
   CODMPRI$ = Text4.Text
   CODINTMPRI% = CODINT%(CODMPRI$)
   If CODINTMPRI% > 0 Then
      Label1.Caption = UNIMED$(CODINTMPRI%)
      Label12.Caption = DESCRIT0$(CODINTMPRI%)
   End If
End Sub

Private Sub Text4_DblClick()
    Call SELECHO("MASTER")
    CODMPRI$ = VALACT1$("MASTER")
    Text4.Text = CODMPRI$
End Sub

Private Sub Command3_Click()
    Call SELECHO("MASTER")
    CODMPRI$ = VALACT1$("MASTER")
    Text4.Text = CODMPRI$
End Sub

Private Sub Text11_Change()
   CODPRO$ = Text11.Text
   CODINTPRO% = CODINT%(CODPRO$)
   If CODINTPRO% > 0 Then
      Call CARGAPRO
   End If
End Sub
'
Private Sub Text11_DBLCLICK()
    Call SELECHO("MASTER")
    CODPRO$ = VALACT1$("MASTER")
    If CODPRO$ <> "" Then
       Text11.Text = CODPRO$
    End If
End Sub
'
Private Sub COMMAND31_CLICK()
    Call SELECHO("MASTER")
    CODPRO$ = VALACT1$("MASTER")
    If CODPRO$ <> "" Then
       Text11.Text = CODPRO$
    End If
End Sub
'
Private Sub Command1_Click()
    Call GRAFORDES
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
    USOX$ = CSTRING$(MKS$(USO), 4, 10, 5, 2)
    PD = 0
    '
    XXX$ = Space$(64)
    Call REPLA(XXX$, CODXX$, 1, 16)
    Call REPLA(XXX$, DEXX$, 17, 28)
    Call REPLA(XXX$, UNXX$, 47, 2)
    Call REPLA(XXX$, USOX$, 51, 10)
    CAIT% = CAIT% + 1
    EXPLOLIN$(CAIT%) = XXX$
    '
    List1.AddItem EXPLOLIN$(CAIT%)
    List1.ListIndex = List1.ListCount - 1
    '
    Text11.Text = ""
    CODPRO$ = ""
    Label10.Caption = ""
    Text14.Text = "1.00000"
    Label13.Caption = ""
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
    USOX$ = CSTRING$(MKS$(USO), 4, 10, 5, 2)
    PD = 0
    '
    XXX$ = Space$(64)
    Call REPLA(XXX$, CODXX$, 1, 16)
    Call REPLA(XXX$, DEXX$, 17, 28)
    Call REPLA(XXX$, UNXX$, 47, 2)
    Call REPLA(XXX$, USOX$, 51, 10)
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
    List1.ListIndex = ILISTA&
    '
    Text11.Text = ""
    CODPRO$ = ""
    Label10.Caption = ""
    Text14.Text = "1.00000"
    Label13.Caption = ""
    Text11.SetFocus
    MODIFI% = 1
    '
    Command4.Enabled = True
    Command8.Enabled = True
    '
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
   Command5.Enabled = False
   PrintForm
   Command5.Enabled = True
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
    Label13.Caption = DESCRIT0$(CDYY%)
    Label10.Caption = UNIMED$(CDYY%)
    Text14.Text = TRIM$(Mid$(List1.Text, 51, 10))
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Text14_GotFocus()
    If TRIM$(Text11.Text) = "" Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If Abs(Val(Text14.Text)) > 0 Then
        If KeyAscii = 13 Then
            Command11.SetFocus
        End If
    End If
End Sub
'
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(TEXT3.Text) > 0 Then
            Text11.SetFocus
        End If
    End If
End Sub
'
Private Sub Command2_Click()
    Call HELPONLINE("FORMULAS")
End Sub
'

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
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    COD$ = ""
    DESARMA% = 1
    Hide
    '
End Sub

Sub CARFORDES()
    CODFOR$ = Text1.Text
    If TRIM$(CODFOR$) <> "" Then
       CODINTFOR% = Abs(CODINT%(CODFOR$))
       If CODINTFOR% > 0 Then
          Call LEEEXPLO(CODFOR$, 1)
          List1.Clear
          Text4.Text = ""
          Label1.Caption = ""
          TEXT3.Text = ""
          Label12.Caption = ""
          If CAIT% >= 1 Then
             CODIMP% = CIJ%(1)
             CANTMP = USOI(1)
             Text4.Text = CODEXT$(CODIMP%)
             Label1.Caption = UNIMED$(CODIMP%)
             TEXT3.Text = CANTMP
             Label12.Caption = DESCRIT0$(CODIMP%)
             '
             For JI% = 2 To CAIT%
                If USOI(JI%) < 0 Then
                  XXX$ = Space$(64)
                  Call REPLA(XXX$, CODEXT$(CIJ%(JI%)), 1, 16)
                  Call REPLA(XXX$, DESCRIT0$(CIJ%(JI%)), 17, 28)
                  Call REPLA(XXX$, UNIMED$(CIJ%(JI%)), 47, 2)
                  Call REPLA(XXX$, CSTRING$(MKS$(Abs(USOI(JI%))), 3, 10, 5, 2), 51, 10)
                  List1.AddItem XXX$
                End If
             Next JI%
          End If
       End If
    End If
    '
End Sub
'

Sub CARGAPRO()
    CODPRO$ = Text11.Text
    If TRIM$(CODPRO$) <> "" Then
      If TRIM$(CODPRO$) <> TRIM$(Text1.Text) Then
        If TRIM$(CODPRO$) <> TRIM$(Text4.Text) Then
          CODINTPRO% = CODINT%(CODPRO$)
          If CODINTPRO% > 0 Then
            Label10.Caption = UNIMED$(CODINTPRO%)
            Label13.Caption = DESCRIT0$(CODINTPRO%)
            Exit Sub
          End If
        End If
      End If
    End If
    Text11.Text = ""
    Label10.Caption = ""
    Label13.Caption = ""
    Text14.Text = ""
End Sub

Sub GRAFORDES()
   If Len(TRIM$(Text1.Text)) >= 3 Then
      If Len(TRIM$(Text2.Text)) > 0 Then
         CODMPRI$ = TRIM$(Text4.Text)
         If CODINT%(CODMPRI$) > 0 Then
            CANMPRI = Val(TEXT3.Text)
            If CANMPRI > 0 Then
               If List1.ListCount > 0 Then
                  COD$ = AJUSTI$(TRIM$(Text1.Text), 16)
                  If CODINT%(COD$) = 0 Then
                     DE$ = TRIM$(Text2.Text)
                     UNID$ = "U."
                     TIMA% = 0
                     CRIT% = 0
                     STMIN = 0
                     STMI& = 0
                     STMII% = STMI&
                     LOREP = 0
                     LORE& = 0
                     LOREI% = LORE&
                     LOTEMA = 0
                     LOMA& = LORE&
                     LOTEMAI% = LOMA&
                     COU = 0
                     MONEI% = 0
                     FECO% = HOY(HOS$)
                     FCP% = 0
                     Call ALTAMAS
                     Call FRESHECHO("MASTER")
                     Call FRESHECHO("ECOMAST")
                  End If
                  CIIX% = CODINT%(COD$)
                  '
                  For U& = 1 To ULTREG&("FORMUDES")
                     X$ = REGLEIDO$("FORMUDES", U&)
                     If CVI(Left$(X$, 2)) = CIIX% Then
                        Call REPLA(X$, Text2.Text, 3, 62)
                        Call GRAREG("FORMUDES", X$, U&)
                        GoTo 20
                     End If
                  Next U&
                  X$ = REGBLAN$("FORMUDES")
                  Call REPLA(X$, MKI$(CIIX%), 1, 2)
                  Call REPLA(X$, Text2.Text, 3, 62)
                  Call REGAPP("FORMUDES", X$)
                  '
20                Call CIERRARCH("FORMUDES")
                  Call FRESHECHO("FORMUDES")
                  '
                  ILISTA& = List1.ListIndex
                  CAIT% = 1 + List1.ListCount
                  EXPLOLIN$(1) = Space$(64)
                  Call REPLA(EXPLOLIN$(1), TRIM$(Text4.Text), 1, 16)
                  Call REPLA(EXPLOLIN$(1), TEXT3.Text, 49, 16)
                  '
                  For I& = 1 To List1.ListCount
                     EXPLOLIN(I& + 1) = List1.List(I& - 1)
                     Call REPLA(EXPLOLIN$(I& + 1), "-", 49, 1)
                  Next I&
                  '
                  COD$ = AJUSTI$(Text1.Text, 16)
                  CABAS = 1
                  '
                  If CABAS > 0.01 Then
                     Call EXPLOGR(COD$, CABAS)
                     MODIFI% = 0
                  End If
                  '
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
   '
   Call MENSERR(24, "Imposible Grabar - Faltan Datos")
   '
End Sub

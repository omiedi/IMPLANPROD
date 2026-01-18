VERSION 5.00
Begin VB.Form HORUTA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secuencia de Operaciones Standard de Producto "
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7665
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   26
      Top             =   4515
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command8 
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
      Height          =   795
      Left            =   7665
      Picture         =   "HORUTA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton Command7 
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
      TabIndex        =   24
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3130
      Left            =   210
      TabIndex        =   19
      Top             =   1470
      Width           =   7155
      Begin VB.CommandButton Command6 
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
         Height          =   360
         Left            =   6120
         Picture         =   "HORUTA.frx":066A
         TabIndex        =   7
         Top             =   2520
         Width           =   855
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
         Left            =   1210
         TabIndex        =   2
         Text            =   " "
         Top             =   2520
         Width           =   4740
      End
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
         Left            =   170
         TabIndex        =   1
         Text            =   " "
         Top             =   2520
         Width           =   645
      End
      Begin VB.CommandButton Command5 
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
         Height          =   360
         Left            =   6120
         Picture         =   "HORUTA.frx":0974
         TabIndex        =   6
         Top             =   1995
         Width           =   855
      End
      Begin VB.CommandButton Command11 
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
         Height          =   360
         Left            =   6120
         Picture         =   "HORUTA.frx":0C7E
         TabIndex        =   3
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         Picture         =   "HORUTA.frx":0F88
         TabIndex        =   4
         Top             =   945
         Width           =   855
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
         Height          =   360
         Left            =   6120
         Picture         =   "HORUTA.frx":1292
         TabIndex        =   5
         Top             =   1470
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   170
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   420
         Width           =   5790
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   360
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   9000
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "     N-O          Descripcion "
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
            TabIndex        =   22
            Top             =   105
            Width           =   7995
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
      Left            =   5775
      TabIndex        =   14
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
      Left            =   6720
      TabIndex        =   13
      Text            =   " "
      Top             =   840
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
      TabIndex        =   12
      Top             =   735
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
      Height          =   795
      Left            =   7665
      Picture         =   "HORUTA.frx":159C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3570
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
      Height          =   795
      Left            =   7665
      Picture         =   "HORUTA.frx":18A6
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Height          =   795
      Left            =   7665
      Picture         =   "HORUTA.frx":19F0
      Style           =   1  'Graphical
      TabIndex        =   10
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
      Height          =   795
      Left            =   7665
      Picture         =   "HORUTA.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2835
      TabIndex        =   23
      Top             =   4935
      Width           =   7050
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "HORUTA.frx":2004
      Stretch         =   -1  'True
      Top             =   4830
      Width           =   11340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Std:"
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
      TabIndex        =   18
      Top             =   315
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Producto: "
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
      Left            =   420
      TabIndex        =   17
      Top             =   315
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   5775
      TabIndex        =   16
      Top             =   945
      Width           =   855
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
      TabIndex        =   15
      Top             =   735
      Width           =   1485
   End
End
Attribute VB_Name = "HORUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HIDOCLICK As Integer
'

Private Sub Command1_Click()
    '
    Call GRASECOR
    Text1.SetFocus
    '
End Sub

Private Sub Command11_Click()
    '
    If List1.ListIndex < 0 Then
        Call MENSERR(24, "Opción Editar no Válida")
        Exit Sub
    End If
    '
    Call ACTSECOR
    '
End Sub

Private Sub Command12_Click()
    '
    If List1.ListIndex < 0 Then
        Call MENSERR(24, "Opción Agregar no Válida")
        Exit Sub
    End If
    '
    NOPY% = 10 * Int(ULTOP% / 10) + 10
    Text5.Text = TRIM$(Str$(NOPY%))
    Text6.Text = ""
    Text5.SetFocus
    '
End Sub

Private Sub Command4_Click()
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
    '
End Sub
'

Private Sub Command5_Click()
    '
    If CODINT%(COD$) < 1 Then
        Call MENSERR(24, "Opción Restaurar no Válida")
        Exit Sub
    End If
    '
    Call LEESECOR
    Text1.SetFocus
    '
End Sub
'

Private Sub Command6_Click()
    '
    NOPY% = Val(Text5.Text)
    If NOPY% < 1 Or NOPY% > 9999 Then
       Call MENSERR(24, "Número de Operación no Válido")
       NOPY% = 10 * Int(ULTOP% / 10) + 10
       Text5.Text = TRIM$(Str$(NOPY%))
       Text5.SetFocus
       Exit Sub
    End If
    '
    ILISTA& = HORUTA.List1.ListIndex
    For I% = 0 To HORUTA.List1.ListCount - 1
       HORUTA.List1.ListIndex = I%
       If Val(Left$(HORUTA.List1.Text, 5)) = NOPY% Then
           HORUTA.List1.ListIndex = ILISTA&
           Call MENSERR(24, "Número de Operación no Válido")
           NOPY% = 10 * Int(ULTOP% / 10) + 10
           Text5.Text = TRIM$(Str$(NOPY%))
           Text5.SetFocus
           Exit Sub
       End If
    Next I%
    '
    If TRIM$(Text6.Text) = "" Then
       Call MENSERR(24, "Falta Denominación de Operación")
       Exit Sub
    End If
    '
    NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
    DOPE$ = AJUSTI$(HORUTA.Text6.Text, 48)
    REGO& = 0
    REGOX$ = AJUSTD$(Str$(REGO&), 6)
    SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
    SECO$ = Left$(SECO$ + Space$(192), 192)
    X$ = REGBLAN$("SECOPER")
    Call REPLA(X$, MKI$(CI%), 1, 2)
    Call REPLA(X$, MKI$(NOPY%), 3, 2)
    Call REPLA(X$, DOPE$, 5, 48)
    Call REPLA(SECO$, X$, 65, 128)
    '
    HORUTA.List1.AddItem SECO$
    HORUTA.List1.Refresh
    For I% = 0 To HORUTA.List1.ListCount - 1
       HORUTA.List1.ListIndex = I%
       If Val(Left$(HORUTA.List1.Text, 5)) = NOPY% Then
           ILISTA& = I%
       End If
    Next I%
    HORUTA.List1.ListIndex = ILISTA&
    HORUTA.Text5.Text = ""
    HORUTA.Text6.Text = ""
    HORUTA.Text5.SetFocus
    Text1.SetFocus
    Call ACTSECOR
    '
End Sub

Private Sub Command7_Click()
    Text1.Text = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call LISECOP
    Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub List1_DblClick()
    '
    Call ACTSECOR
    '
End Sub


Private Sub Text1_DBLCLICK()
    Text1.Text = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub
'

Private Sub Text1_GotFocus()
    HORUTA.Refresh
    If COMMA$ <> "" Then
        Text1.Text = COMMA$
        COMMA$ = ""
    End If
    '
    If Text1.Text = "" Then
        Text1.Text = VALACT1$("MASTER")
    End If
    COD$ = TRIM$(Text1.Text)
    CI% = CODINT%(COD$)
    If CI% > 0 Then
         Text1.Text = COD$
         Text2.Text = DESCRIT$(CI%)
         TEXT3.Text = UNIMED$(CI%)
         Text4.Text = LOTESTAN(CI%)
         Call LEESECOR
         'Text4.SetFocus
       Else
         Text1.Text = ""
         Text2.Text = ""
         TEXT3.Text = ""
         Text4.Text = ""
         CAIT% = 0
         List1.Clear
         ' Text1.SetFocus
         Exit Sub
    End If
    '
End Sub
'

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        COD$ = TRIM$(Text1.Text)
        CODX$ = COD$
        CI% = CODINT%(CODX$)
        If CI% > 0 Then
            Text1.Text = COD$
            Text2.Text = DESCRIT$(CI%)
            TEXT3.Text = UNIMED$(CI%)
            Text4.Text = LOTESTAN(CI%)
            Call LEESECOR
            Text4.SetFocus
          Else
            Call MENSERR(24, "Código no Válido")
            Text1.Text = ""
            Text2.Text = ""
            TEXT3.Text = ""
            Text4.Text = ""
            CAIT% = 0
            List1.Clear
            Text1.SetFocus
        End If
    End If
End Sub
'

Private Sub Text14_GotFocus()
    If TRIM$(Text11.Text) = "" Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text15_GotFocus()
    If TRIM$(Text11.Text) = "" Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text2_GotFocus()
    TEXT3.SetFocus
End Sub
'

Private Sub Text3_GotFocus()
    CODX$ = TRIM$(Text1.Text)
    CIX% = CODINT%(CODX$)
    If CIX% > 0 Then
          COD$ = CODX$: CI% = CIX%
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
    Call HELPONLINE("SECOPER")
End Sub
'

Private Sub Command3_Click()
    '
    If CODINT%(COD$) < 1 Then
        Call MENSERR(24, "Opción Restaurar no Válida")
        Exit Sub
    End If
    '
    Call LEESECOR
    Text1.SetFocus
    '
End Sub

Private Sub Command9_Click()
    Call CIERRARCH("*.*")
    'Call FINAL("*SORSECO")
    Hide
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
      If TRIM$(COD$) <> "" Then
        CABAS = Val(Text4.Text)
        Call GRALOTSTAN
      End If
    End If
    '
End Sub

Private Sub Text4_GOTFOCUS()
    '
    If Val(Text4.Text) < 1 Then
        Text4.Text = "1"
        CABAS = 1
    End If
    '
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

Private Sub Text4_LostFocus()
    '
    If TRIM$(COD$) <> "" Then
      CABAS = Val(Text4.Text)
      Call GRALOTSTAN
    End If
    '
End Sub

Private Sub Text5_GOTFOCUS()
    '
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    '
End Sub
'

Private Sub Text5_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       Text6.SetFocus
    End If
    '
End Sub
'

Private Sub Text6_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       Command6.SetFocus
    End If
    '
End Sub
'

Private Sub Text6_GOTFOCUS()
    '
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    '
End Sub


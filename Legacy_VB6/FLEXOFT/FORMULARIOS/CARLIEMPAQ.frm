VERSION 5.00
Begin VB.Form CARLIEMPAQ 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Listas de Empaque"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6300
      MaxLength       =   4
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "1/1"
      Top             =   2415
      Width           =   1275
   End
   Begin VB.TextBox Text11 
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
      Left            =   1785
      TabIndex        =   29
      Text            =   " "
      Top             =   2625
      Width           =   2955
   End
   Begin VB.TextBox Text10 
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
      Left            =   1785
      TabIndex        =   27
      Text            =   " "
      Top             =   2100
      Width           =   2955
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
      Height          =   2985
      Left            =   210
      TabIndex        =   0
      Top             =   5025
      Width           =   7365
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   7305
      TabIndex        =   23
      Top             =   4620
      Width           =   7365
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5590
         TabIndex        =   26
         Top             =   50
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         TabIndex        =   25
         Top             =   50
         Width           =   2010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   50
         TabIndex        =   24
         Top             =   50
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7790
      TabIndex        =   20
      Top             =   5565
      Width           =   855
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   105
         Picture         =   "CARLIEMPAQ.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1155
         Width           =   645
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   210
         MaxLength       =   1
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   400
         Width           =   435
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7790
      Picture         =   "CARLIEMPAQ.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   315
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
      Left            =   7790
      Picture         =   "CARLIEMPAQ.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1260
      Width           =   855
   End
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
      Height          =   855
      Left            =   7790
      Picture         =   "CARLIEMPAQ.frx":0556
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Clear"
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
      Left            =   7770
      Picture         =   "CARLIEMPAQ.frx":0998
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4620
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1260
      TabIndex        =   15
      Top             =   1365
      Width           =   175
   End
   Begin VB.TextBox Text8 
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
      Left            =   210
      TabIndex        =   13
      Text            =   " "
      Top             =   1365
      Width           =   1065
   End
   Begin VB.TextBox Text7 
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
      Left            =   1785
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   " "
      Top             =   315
      Width           =   5790
   End
   Begin VB.TextBox Text6 
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
      Left            =   1785
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   " "
      Top             =   840
      Width           =   5790
   End
   Begin VB.TextBox Text4 
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
      Left            =   2940
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   " "
      Top             =   1365
      Width           =   4635
   End
   Begin VB.TextBox Text5 
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
      Left            =   1785
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   " "
      Top             =   1365
      Width           =   960
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
      Height          =   360
      Left            =   3570
      TabIndex        =   4
      Top             =   3465
      Width           =   175
   End
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
      Height          =   360
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3990
      Width           =   7365
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6300
      MaxLength       =   4
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   3270
      Width           =   1275
   End
   Begin VB.TextBox Text1 
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
      Left            =   1050
      TabIndex        =   1
      Text            =   " "
      Top             =   3465
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Número:"
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
      Left            =   5355
      TabIndex        =   32
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Pedido:"
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
      Left            =   -420
      TabIndex        =   30
      Top             =   2730
      Width           =   2115
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Factura:"
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
      Left            =   -420
      TabIndex        =   28
      Top             =   2205
      Width           =   2115
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   7545
      Picture         =   "CARLIEMPAQ.frx":0CA2
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   5250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
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
      Left            =   210
      TabIndex        =   14
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label2 
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
      Left            =   3675
      TabIndex        =   8
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   3570
      Width           =   960
   End
   Begin VB.Label Label30 
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
      Height          =   225
      Left            =   5355
      TabIndex        =   6
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label UMEDI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4620
      TabIndex        =   5
      Top             =   3465
      Width           =   645
   End
End
Attribute VB_Name = "CARLIEMPAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        Text7.TEXT = RASOCLI$(NC%)
        Text6.TEXT = DOMICLI$(NC%)
        Text5.TEXT = CPOSCLI$(NC%)
        Text4.TEXT = LOCACLI$(NC%)
        Text12.TEXT = "1/1"
        Text11.TEXT = ""
        Text10.TEXT = ""
    End If
    '
End Sub

Private Sub Command16_Click()
    '
    FORIPRE$ = Control$("", "FORLIPAK")
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario\de Impresión de Listas de Empaque.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    TIPODOC$ = "Lista de Empaque"
    NUMEDOC$ = "S/N"
    CACOPDOC% = Val(Text9.TEXT)
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CACOLTAB1% = 3
    '
    DESTIDOC% = 0
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    For U& = 1 To List1.ListCount
       Z$ = CARLIEMPAQ.List1.List(U& - 1)
       CACO = Val(Mid$(Z$, 52, 6))
       If CACO >= 0.05 Then
          CODX$ = Left$(Z$, 16)
          CIIX% = CODINT%(CODX$)
          DECO$ = DESCRIT0$(CIIX%)
          '
          CAPEX$ = Str$(CACO)
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = CODX$
          TETABU1$(2, CAITAB1%) = DECO$
          TETABU1$(3, CAITAB1%) = CAPEX$
       End If
    Next U&
    '
    CODPARAM$(1) = "RASO-CLI": DOCPARAM$(1) = TRIM$(Text7.TEXT)
    CODPARAM$(2) = "DOMI-CLI": DOCPARAM$(2) = TRIM$(Text6.TEXT)
    CODPARAM$(3) = "CPOS-CLI": DOCPARAM$(3) = TRIM$(Text5.TEXT)
    CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = TRIM$(Text4.TEXT)
    CODPARAM$(5) = "NRO-REMI": DOCPARAM$(5) = TRIM$(Text12.TEXT)
    CODPARAM$(6) = "NUME-COM": DOCPARAM$(6) = TRIM$(Text10.TEXT)
    CODPARAM$(7) = "NRO-AENT": DOCPARAM$(7) = TRIM$(Text11.TEXT)
    CAPARDOC% = 7
    '
    Call PRINTDOC("FORLIPAK")        ' LISTA DE EMPAQUE
    List1.Clear
    '
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    COD$ = CODD$
    Text1.TEXT = COD$
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   OPAKLIST.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command27_Click()
   If List1.ListCount > 0 Then
      If COMALTER%("Cancela sin Imprimir ?\\No - Restaurar\Si - Cancelar") = 2 Then
         List1.Clear
      End If
   End If
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("LISEMPAQ")
End Sub

Private Sub Command3_Click()
   If List1.ListCount > 0 Then
      If COMALTER%("Cancela sin Imprimir ?\\No - Restaurar\Si - Cancelar") <> 2 Then
         Exit Sub
      End If
   End If
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Form_Load()
    Show
    Refresh
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub List1_DblClick()
   REAC = List1.ListIndex
   If REAC >= 0 Then
      If REAC < List1.ListCount Then
         List1.RemoveItem REAC
         If REAC < List1.ListCount Then
             List1.ListIndex = REAC
           Else
             List1.ListIndex = List1.ListCount - 1
         End If
      End If
   End If
End Sub

Private Sub Text1_Change()
    '
    Text3.TEXT = ""
    UMEDI.Caption = ""
    '
    If Len(Text1.TEXT) >= 8 Then
       If CODINT%(Text1.TEXT) > 0 Then
          COD$ = TRIM$(Text1.TEXT)
          CI% = CODINT%(COD$)
          UMEDI.Caption = UNIMED$(CI%)
          Text3.TEXT = DESCRIT$(CI%)
          CAN = Val(Text2.TEXT)
          For U& = 1 To List1.ListCount
             If TRIM$(Left$(List1.List(U& - 1), 16)) = COD$ Then
                TTXX$ = List1.List(U& - 1)
                CAN = CAN + Val(Mid$(List1.List(U& - 1), 52, 6))
                List1.RemoveItem (U& - 1)
                Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 52, 6)
                List1.AddItem TTXX$, U& - 1
                List1.ListIndex = U& - 1
                GoTo 20
             End If
          Next U&
          '
          TTXX$ = Space$(64)
          Call REPLA(TTXX$, AJUSTI$(COD$, 16), 1, 16)
          Call REPLA(TTXX$, DESCRIT$(CI%), 18, 33)
          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 52, 6)
          List1.AddItem TTXX$
          List1.ListIndex = List1.ListCount - 1
20        Text1.TEXT = ""
          Text2.TEXT = 1
       End If
    End If
    Text1.SetFocus
    '
End Sub
'

Private Sub Text8_DBLCLICK()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        Text7.TEXT = RASOCLI$(NC%)
        Text6.TEXT = DOMICLI$(NC%)
        Text5.TEXT = CPOSCLI$(NC%)
        Text4.TEXT = LOCACLI$(NC%)
        Text12.TEXT = "1/1"
        Text11.TEXT = ""
        Text10.TEXT = ""
    End If
    '
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii% = 13 Then
       NCXI% = Val(Text8.TEXT)
       If NCXI% > 0 Then
          NC% = NCXI%
          Text7.TEXT = RASOCLI$(NC%)
          Text6.TEXT = DOMICLI$(NC%)
          Text5.TEXT = CPOSCLI$(NC%)
          Text4.TEXT = LOCACLI$(NC%)
       End If
    End If
    '

End Sub

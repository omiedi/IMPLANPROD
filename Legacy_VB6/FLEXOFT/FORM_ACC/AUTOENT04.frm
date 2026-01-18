VERSION 5.00
Begin VB.Form AUTOENT04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Entrega "
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9540
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Auto. de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1005
      Left            =   5670
      TabIndex        =   24
      Top             =   525
      Width           =   2745
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
         Height          =   285
         Left            =   2310
         TabIndex        =   25
         Top             =   630
         Width           =   195
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1410
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Fecha de Emisión de la Autorización de Entrega"
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   -260
         TabIndex        =   29
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label18 
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
         Left            =   1515
         TabIndex        =   28
         ToolTipText     =   "Número de Autorización de Entrega"
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   -365
         TabIndex        =   27
         Top             =   735
         Width           =   1695
      End
   End
   Begin VB.ListBox LICOPEN 
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Doble-Click para Autorizar la Entrega"
      Top             =   2040
      Width           =   9285
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   9225
      TabIndex        =   6
      Top             =   1630
      Width           =   9285
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Código           Descripción                    U/S      Vig.O/C      A/E      "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   0
         TabIndex        =   7
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   1575
      Left            =   8610
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command39 
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
         Height          =   650
         Left            =   105
         Picture         =   "AUTOENT04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Imprime Autorización de Entrega"
         Top             =   840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4305
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "AUTOENT04.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AUTOENT04.frx":0974
         Top             =   4600
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9400
      Y1              =   5025
      Y2              =   5025
   End
   Begin VB.Line Line2 
      X1              =   9400
      X2              =   9400
      Y1              =   1650
      Y2              =   5000
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6615
      TabIndex        =   23
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label15 
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
      Left            =   7520
      TabIndex        =   22
      ToolTipText     =   "Fecha de Emisión de la Orden de Compra"
      Top             =   105
      Width           =   870
   End
   Begin VB.Label Label14 
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
      Left            =   6195
      TabIndex        =   21
      ToolTipText     =   "Número de Orden de Compra"
      Top             =   105
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/C:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   4420
      TabIndex        =   20
      Top             =   160
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6930
      TabIndex        =   19
      Top             =   160
      Width           =   540
   End
   Begin VB.Label Label11 
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
      Left            =   1785
      TabIndex        =   18
      Top             =   4935
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label10 
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
      Left            =   105
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   1440
      TabIndex        =   16
      Top             =   1230
      Width           =   4095
   End
   Begin VB.Label Label8 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label7 
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
      Left            =   1440
      TabIndex        =   14
      Top             =   840
      Width           =   735
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
      Left            =   1440
      TabIndex        =   13
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label4 
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
      Left            =   1440
      TabIndex        =   12
      Top             =   105
      Width           =   4095
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
      Left            =   480
      TabIndex        =   11
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   975
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Posición IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   4005
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   4005
      TabIndex        =   8
      Top             =   855
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7050
      Y1              =   4340
      Y2              =   4340
   End
End
Attribute VB_Name = "AUTOENT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HUBOAENT%
'
Private Sub BOTEXIT_Click()
   If HUBOAENT% = 1 Then
     Call Command39_Click
     HUBOAENT% = 0
   End If
   LICOPEN.Clear
   Unload Me
End Sub
'
Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command39_Click()
  '
  Command39.Enabled = False
  Screen.MousePointer = 11
'  Call SETULTREG("!OBSERVOF", 0)
'  Call FRATEXTO(Text13, 52)
'  For KKU& = 1 To CARETEX%
'    Call REGAPP("!OBSERVOF", RETEX$(KKU&))
'  Next KKU&
'  Call CIERRARCH("!OBSERVOF")
  'IMPRESION DE LA AUTORIZACION DE ENTREGA
  '
  'FORIPRE$ = CONTROL$("", "FORMOCOM")
'EN EL SETUP CONFIGURAR EL CONTROL DEL FORMULARIO DE IMPRESION
  FORIPRE$ = "AENT-IGD"
  '
  If TRIM$(FORIPRE$) <> "" Then
    '
    Screen.MousePointer = 11
    FECHDOC$ = Text14
    NUMEDOC$ = TRIM$(Label18)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = NUMEDOC$
    TIPODOC$ = "Autorización de Entrega"
    '
    CODPARAM$(1) = "N-OCOMPR"
    CODPARAM$(2) = "N-AUTENT"
    CAPARDOC% = 2
    '
    DOCPARAM$(1) = FORMATNUM$(XVALO(Label14), "F7.0")
    DOCPARAM$(2) = FORMATNUM$(XVALO(Label18), "F7.0")
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "FECH-VEN"
    CODTABU1$(5) = "REF-MAT1"
    CACOLTAB1% = 5
    '
    DESTIDOC% = (-1) * Val(Label3)
    '
    NC1% = DESTIDOC%
    Call ABRECOMPRAS
    NUMEOC1 = XVALO(Label14)
    SQLX$ = "SELECT *  FROM AUTOENT "
    SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUMEOC1 & ""
    SQLX$ = SQLX$ + " AND  Nume_Aten = " & TRIM$(Label18)
    SQLX$ = SQLX$ + " AND  STAT_ATEN = 0"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom ASC, Cod_Exte ASC, Fentre_Sol ASC"
    Set AETMP = Compras.OpenRecordset(SQLX$, 4)
    With AETMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Autorizaciones de Entrega\para la OC Seleccionada")
        On Error GoTo 0
        GoTo 999
      End If
      '
      On Error GoTo 0
      '
      CODANT% = 0
      CAITAB1% = 0
      Do While .EOF = False
        CI0% = !Cod_inte
        'Para saltear los codigos iguales.
        If CODANT% = CI0% Then
          CI0% = 0
        End If
        CODANT% = !Cod_inte
        CANTU = !Cant_Aten
        OBSX1$ = !Observa_Aten
        FEX = CVINT(!Fentre_Sol)
        '
        CAITAB1% = CAITAB1% + 1
        CODDX$ = TRIM$(CODEXT$(CI0%))
        '
        TETABU1$(1, CAITAB1%) = CODDX$ 'TRIM$(Str$(CI0%))
        TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
        If CI0% = 0 Then
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = ""
        End If
        TETABU1$(3, CAITAB1%) = TRIM$(Str$(CANTU))
        TETABU1$(4, CAITAB1%) = FECHATEX$(FEX)
        TETABU1$(5, CAITAB1%) = OBSX1$
        '
        .MoveNext
      Loop
    End With
    '
    Call PRINTDOC("@AENT-IGD")   ' Orden de Compra
    '
    For KK1% = 1 To CACOLTAB1%
      For KK2% = 1 To CAITAB1%
        TETABU1$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '
    DOCUORIG$ = "A/Ent-" + NUMEDOC$
    NUMEDOC$ = ""
    '
  End If
999 Call CIERRARCH("*.*")
  Command39.Enabled = True
  Screen.MousePointer = 1
  '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If HUBOAENT% = 1 Then
     Call Command39_Click
     HUBOAENT% = 0
   End If
End Sub

Private Sub Label14_Change()
   '
   NUOCO& = Val(Label14)
   If NUOCO& > 0 Then
     Call ABRECOMPRAS
     SQLX$ = " SELECT Npro_Ocom, Femi_Ocom, AtSr_Ocom "
     SQLX$ = SQLX$ + " FROM ENCOCABI "
     SQLX$ = SQLX$ + " WHERE Nume_Ocom =   " & NUOCO& & "     "
     Set COMPLETEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPLETEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "Imposible Presentar O/C " + Str$(NUOCO&))
         Exit Sub
       End If
       '
       NPRO% = !Npro_Ocom
       Label15 = FECHATEX$(CVINT(!Femi_Ocom))
       Label9 = !AtSr_Ocom
     End With
     Set COMPLETEMP = Nothing
     '
     Label3 = NPRO%
     NPRO1% = (-1) * NPRO%
     '
     Label4 = RASOCLI(NPRO1%)
     Label6 = DOMICLI(NPRO1%)
     Label7 = CPOSCLI(NPRO1%)
     Label8 = LOCACLI(NPRO1%)
     Label10 = PIVACLI(NPRO1%)
     Label11 = CUITCLI(NPRO1%)
     '
     SQLX$ = "SELECT * FROM DETOCABI "
     SQLX$ = SQLX$ + " WHERE  Nume_Ocom =   " & NUOCO& & " "
     SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
     Set ITETEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     JJ& = 0
     With ITETEMP
       .MoveFirst
       LICOPEN.Clear
       Do While (.EOF = False)
         CII% = !Cod_inte
         DEMAT$ = DESCRIT0$(CII%)
           If CII% <= 0 Then              ' PARA CONSIDERAR TEXTO LIBRE
             DEMAT$ = TRIM$(Left$(!Delibre, 64))
           End If
         COEFCAN = !Coef_Unids
           If COEFCAN < 0.005 Then COEFCAN = 1
         CAORI = !Cant_Ocom * COEFCAN
         CAREC = !Cant_Rec * COEFCAN
         CAPEN = CAORI - CAREC: If CAPEN < 0 Then CAPEN = 0
         UNICO$ = !Uni_Mas
         COEFUNI = !Coef_Unids
         STAT% = !Stat_Ocom
         NUORIT% = !NuOrd_Item
         FV1% = CVINT(!FVig_Ocom)
         'PUNI1 = !Pre_Unit
         '
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
         Call REPLA(TTXX$, DEMAT$, 18, 30)
         Call REPLA(TTXX$, UNICO$, 49, 4)
         Call REPLA(TTXX$, FECHATEX$(FV1%), 58, 8)
         'Call REPLA(TTXX$, FORMATNUM$(PUNI1, "F10.4"), 70, 10)
         Call REPLA(TTXX$, Str$(NUORIT%), 121, 8)  ' numero de orden del item
         '
         If STAT% = 8 Then
           Call REPLA(TTXX$, "Cumplido", 76, 10)
         End If
         JJ& = JJ& + 1
         LICOPEN.AddItem TTXX$
       .MoveNext
       Loop
     End With
     HUBOAENT% = 0
     '
   End If
   '
End Sub

'
Private Sub LICOPEN_DblClick()
  '
  'Label18 = Str$(PRONUAE&)
  '
  Call BLANARCH("!AUTOENT")
  '
  Call ABRECOMPRAS
  '
  Screen.MousePointer = 11
  NUOCO& = Val(Label14)
  TTXX$ = AUTOENT04.LICOPEN.TEXT
  NUORIT% = Val(TRIM$(Mid$(TTXX$, 121, 8)))
  '
  SQLX$ = "SELECT * FROM DETOCABI "
  SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
  SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
  SQLX$ = SQLX$ + "AND NuOrd_Item = " & NUORIT% & " "
  '
  Set MARCTEMP = Compras.OpenRecordset(SQLX$, dbOpenDynaset)
  '
  Y$ = REGBLAN$("AUTOENT")
  With MARCTEMP
    .MoveFirst
    NPRO1% = !Npro_Ocom
    Do While (.EOF = False)
       CI% = !Cod_inte
       .MoveNext
    Loop
  End With
  'CARGO ENTREGAS ANTERIORES CON SALDO PENDIENTE
  Y$ = REGBLAN$("AUTOENT")
  With AutoEnt
    SQLX$ = "Nume_Ocom = " & NUOCO& & " "
    SQLX$ = SQLX$ + " AND Stat_Aten < 9 "
    SQLX$ = SQLX$ + " AND Cod_Inte= " & CI% & ""
    SQLX$ = SQLX$ + " AND Sald_Aten > 0.005 "
    JJJ& = 0
    .FindFirst SQLX$
25  If .NoMatch = False Then
      Call REPLA(Y$, MKI$(CI%), 1, 2)
        FES1% = CVINT%(!Fentre_Sol)
      Call REPLA(Y$, MKI$(FES1%), 3, 2)
      Call REPLA(Y$, MKD$(!Sald_Aten), 5, 8)
      Call REPLA(Y$, !Observa_Aten, 13, 30)
      JJJ& = JJJ& + 1
      Call GRAREG("!AUTOENT", Y$, JJJ&)
      .FindNext SQLX$
      GoTo 25
    End If
  End With
  '
  Call SETULTREG("!AUTOENT", JJJ&)
  '
  CE1$ = CODEXT$(CI%)
  DE1$ = DESCRIT0$(CI%)
  '
20  NAENT& = Val(Label18)     ' PRONUAE&
  ATRI$ = "/NC"
  ATRI$ = ATRI$ + "/TITUPAN=Entrega Nro.: " + Str$(NAENT&) + " - Artículo: " + CE1$ + " - " + DE1$
  'ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Frame1.Left + 1900))
  'ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
  '
  VTB1% = VENTABU%("AUTOENT" + ATRI$)
  If VTB1% >= 0 Then
      For J& = 1 To ULTREG&("!AUTOENT")
         X$ = REGLEIDO$("!AUTOENT", J&)
         CI1% = CI%
         FES1% = CVI(Mid$(X$, 3, 2))
         Q1 = CVD(Mid$(X$, 5, 8))
         OBS1$ = Mid$(X$, 13, 30)
         HOII% = HOY(HOS$)
         'VALIDACIONES
         If FES1% < HOII% Then
           Call MENSERR(24, "Fecha de Entrega no Válida")
           GoTo 20
         End If
      Next J&
      '
      'MARCA COMO ANULADA LA A.ENT. ANTERIOR
      SQLX$ = "Nume_Ocom = " & NUOCO& & " "
      SQLX$ = SQLX$ + " AND Stat_Aten < 9 "
      SQLX$ = SQLX$ + " AND Cod_Inte= " & CI% & ""
      SQLX$ = SQLX$ + " AND Sald_Aten > 0.005 "
      '
      With AutoEnt
        .FindFirst SQLX$
35      If .NoMatch = False Then
          .Edit
          !Stat_Aten = 9
          .Update
          .FindNext SQLX$
          GoTo 35
        End If
      End With
      '
      '4.- Grabar Autorizacion de Entrega
      ORDIT% = 0
      For J& = 1 To ULTREG&("!AUTOENT")
         X$ = REGLEIDO$("!AUTOENT", J&)
         CI1% = CI% 'CVI(Mid$(X$, 1, 2))
         FES1% = CVI(Mid$(X$, 3, 2))
         Q1 = CVD(Mid$(X$, 5, 8))
         If Q1 < 0.005 Then GoTo 45
         OBS1$ = Mid$(X$, 13, 30)
         FEAE% = FECHANUM(Text14.TEXT)
         '
         'ACA TENDRIA QUE BUSCAR SI NO ES UNA EDICION,
         'ENTONCES HAGO UN ADDNEW
         ORDIT% = ORDIT% + 1
         With AutoEnt
           .AddNew
             !Nume_Ocom = NUOCO&
             !Nume_Aten = NAENT&
             !Femi_Aten = CVDAT(FEAE%)
             !Npro_Aten = NPRO1%
             !Raso_Aten = RASOCLI$(NPRO1% * (-1))
             !Cod_inte = CI1%
             !Cod_Exte = CODEXT$(CI1%)
             !Descritem = DESCRIT0$(CI1%)
             !Fentre_Sol = CVDAT(FES1%)
             !Cant_Aten = Q1
             !Acuen_Aten = 0
             !Sald_Aten = Q1
             !NuOrd_Item = ORDIT%
             !Stat_Aten = 0 'En curso
             !Observa_Aten = OBS1$
           .Update
         End With
         '
45    Next J&
      '
      ILISTA& = LICOPEN.ListIndex
      If ILISTA& >= 0 Then
        TTXX$ = LICOPEN.TEXT
        Call REPLA(TTXX$, FORMATNUM$(NAENT&, "F6.0"), 68, 6)
        LICOPEN.RemoveItem ILISTA&
        LICOPEN.AddItem TTXX$, ILISTA&
        LICOPEN.Refresh
        HUBOAENT% = 1
      End If
     '5.- Imprimir Autorización de Entrega
    Exit Sub
  End If
  '
  Call MENSERR(24, "Autorización Cancelada.\Se Mantiene la AE Vigente.")
  '
End Sub
Function PRONUAE&()
   '
   PRONU& = 1
   Call ABRECOMPRAS
   '
   SQLX$ = "SELECT Nume_Aten FROM AutoEnt "
   SQLX$ = SQLX$ + "ORDER BY Nume_Aten DESC;"
   Set AutoEnTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With AutoEnTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Aten
     End If
   End With
   On Error GoTo 0
   '
   PRONUAE& = PRONU&
   '
End Function

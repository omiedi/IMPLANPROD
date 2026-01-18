VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIMCLI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Maestro"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
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
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMCLI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\CLIENTES.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "'cli1#csv-vdr'$"
      Top             =   105
      Width           =   1905
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
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMCLI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIMCLI.frx":0454
      Height          =   6420
      Left            =   0
      OleObjectBlob   =   "FORIMCLI.frx":0464
      TabIndex        =   1
      Top             =   1365
      Width           =   10830
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aro Plastico"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resorte"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   6
      Top             =   2730
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   4
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   1785
      Width           =   1380
   End
End
Attribute VB_Name = "FORIMCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Close
   Hide
End Sub


Private Sub Command4_Click()
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 50) * DBGrid1.RowHeight
   '
   JJ& = 0
   Dim PUNUCLI As String * 32767
   '
   Dim NCLI%(32767)
   Dim COVENDE%(32), DEVENDE$(32)
   Dim COCONDI%(32), DECONDI$(32)
   Dim COPROVIN$(32), DEPROVIN$(32)
   '
   LU$ = LUDAT$
   '
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   YY0$ = REGBLAN$("BADACLIE")
   PIVA% = 1
   CATIB% = 2
   LIPRE% = 1
   CUEMA% = 1
   '
   Call REPLA(XX0$, Chr$(PIVA%), 136, 1)
   Call REPLA(XX0$, Chr$(LIPRE%), 151, 1)
   Call REPLA(XX0$, Chr$(COPAGI%), 150, 1)
   '
   XX$ = MKI$(1) + "Distribuidores": Call GRAREG("GRUCOCLI", XX$, 1)
   XX$ = MKI$(2) + "Mayoristas": Call GRAREG("GRUCOCLI", XX$, 2)
   Call SETULTREG("GRUCOCLI", 2)
   Call CIERRARCH("GRUCOCLI")
   '
   Call SETULTREG("PROVINC", 0)
   Call SETULTREG("PROVINB", 0)
   CAPROVI% = 0
   For U& = 1 To ULTREG&("PROVINC")
      XX$ = REGLEIDO$("PROVINC", U&)
      CAPROVI% = CAPROVI% + 1
      COPROVIN$(CAPROVI%) = Left$(XX$, 1)
      DEPROVIN$(CAPROVI%) = TRIM$(UCase$(Mid$(XX$, 2, 29)))
   Next U&
   '
   CACONDI% = 0
   If ULTREG("CONPAG") < 1 Then
      XX$ = REGBLAN$("CONPAG")
      Call REPLA(XX$, Chr$(1), 1, 1)
      Call REPLA(XX$, "CONTADO", 2, 35)
      Call REPLA(XX$, Chr$(1), 37, 1)
      Call REGAPP("CONPAG", XX$)
      Call CIERRARCH("CONPAG")
   End If
   For U& = 1 To ULTREG&("CONPAG")
      XX$ = REGLEIDO$("CONPAG", U&)
      CACONDI% = CACONDI% + 1
      COCONDI%(CACONDI%) = Asc(Left$(XX$, 1))
      DECONDI$(CACONDI%) = TRIM$(UCase$(Mid$(XX$, 2, 35)))
   Next U&
   '
   Call BLANARCH("VENDEDOR")
   CAVENDE% = 0
   For U& = 1 To ULTREG&("VENDEDOR")
      XX$ = REGLEIDO$("VENDEDOR", U&)
      CAVENDE% = CAVENDE% + 1
      COVENDE%(CAVENDE%) = Asc(Left$(XX$, 1))
      DEVENDE$(CAVENDE%) = TRIM$(UCase$(Mid$(XX$, 2, 30)))
   Next U&
   '
   REGX& = 0: REGY& = 0: NCIMAX% = 0
   '
' IMPORTACION clientes
   '
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   '
   JJL& = 0
   Call SETULTREG("DEUDOR1", JJL&)
   Call SETULTREG("DEUDOR2", JJL&)
   Call SETULTREG("BADACLIE", JJL&)
   '
   For U& = 1 To URECORD&
     DBGrid1.Row = U&
     DBGrid1.Col = 0: NCI% = Val(DBGrid1.TEXT)
        GRUCLI% = 1
     DBGrid1.Col = 1: RASO$ = DBGrid1.TEXT
       FANT$ = Left$(RASO$, 10)
     DBGrid1.Col = 8: DOMI$ = DBGrid1.TEXT
     DBGrid1.Col = 11: CPOSTX$ = DBGrid1.TEXT
       On Error Resume Next
       CPOST% = Val(CPOSTX$)
       On Error GoTo 0
     DBGrid1.Col = 9: LOCA$ = DBGrid1.TEXT
     DBGrid1.Col = 10: DEPROV$ = DBGrid1.TEXT: GoSub 60
     DBGrid1.Col = 12: TELE$ = DBGrid1.TEXT
     DBGrid1.Col = 13: FAXX$ = DBGrid1.TEXT
     DBGrid1.Col = 16: CUIT$ = DBGrid1.TEXT
     DBGrid1.Col = 31: EXPRESO$ = DBGrid1.TEXT
     DBGrid1.Col = 25: VENDE$ = DBGrid1.TEXT
     DBGrid1.Col = 41:: EMAI$ = DBGrid1.TEXT
      NUVENDE% = Val(Mid$(AA$, 189, 2)): GoSub 63
     DBGrid1.Col = 27: NUCONPA% = Val(DBGrid1.TEXT)
     DBGrid1.Col = 28: CONDIPAG$ = DBGrid1.TEXT: GoSub 64
      LIMICREDI$ = TRIM$(Mid$(AA$, 212, 8)): GoSub 62
      '
      If TRIM$(RASO$) <> "" Then
         JJL& = JJL& + 1
         GoSub GRACLIEN
      End If
      '
   Next U&
   '
   Call SETULTREG("DEUDOR1", JJL&)
   Call SETULTREG("DEUDOR2", JJL&)
   Call SETULTREG("BADACLIE", JJL&)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   Call CIERRARCH("*.*")
   '
   '
   Call COMUNI("La Importación de Clientes y Distribuidores\se ha Completado Exitosamente.")
   '
Exit Sub







20 N1% = FILEOPEN("A:PRECIOS.PRN", 1, 0)
   CI0% = NUMAS%
   Dim CODIMPOR$(16384)
   Dim DEGRU$(256)
   CAIMPOR% = 0
   '
   J& = 0
   LINACT$ = "00"
   While Not EOF(N1%)
      Line Input #N1%, AA$
      If TRIM$(AA$) <> "" Then
         COD$ = TRIM$(Left$(AA$, 10))
         If COD$ = "Línea" Then
            CAGRU% = CAGRU% + 1
            DEGRU$(CAGRU%) = TRIM$(Mid$(AA$, 11, 42))
            LINACT$ = Left$(DEGRU$(CAGRU%), 2)
            GoTo 28
         End If
         CODLIN$ = LINACT$ + "-" + COD$
         '
         For KI% = 1 To CAIMPOR%
            If CODIMPOR$(KI%) = CODLIN$ Then
               Call COMUNI("Imposible Importar Codigo " + COD$ + " - Repetido.")
               GoTo 28
            End If
         Next KI%
         '
         CAIMPOR% = CAIMPOR% + 1
         CODIMPOR$(CAIMPOR%) = CODLIN$
         '
         DE$ = DESCRIT0$(CI%)
         LIMICREDI$ = Mid$(AA$, 53, 10): GoSub 62
         PRE# = CDbl(LICRE)
         CI% = CODINT%(CODLIN$)
         If CI% > 0 Then GoTo 25
         '
         CI% = CODINT%(COD$)
         If CI% > 0 Then
            REMA& = CI% + 1
            REMAS$ = REGLEIDO$("MASTER", REMA&)
            Call REPLA(REMAS$, CODLIN$, 77, 16)
            Call GRAREG("MASTER", REMAS$, REMA&)
            GoTo 25
         End If
         '
         CI0% = CI0% + 1
         CI% = CI0%
         DE$ = Mid$(AA$, 11, 42)
         REMA& = CI% + 1
         REMAS$ = String$(128, 0)
         Call REPLA(REMAS$, MKI$(CI%), 1, 2)
         Call REPLA(REMAS$, DE$, 5, 64)
         Call REPLA(REMAS$, String$(8, 0), 69, 8)
         Call REPLA(REMAS$, CODLIN$, 77, 16)
         Call REPLA(REMAS$, MKI$(CAGRU%), 125, 2)
         Call GRAREG("MASTER", REMAS$, REMA&)
         '
25       Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, DE$, 3, 42)
         Call REPLA(Z$, MKD$(PRE#), 62, 8)
         J& = J& + 1
         Call GRAREG("LIPRE001", Z$, J&)
      End If
28 Wend
   '
29 Close #N1%
   Call SETULTREG("LIPRE001", J&)
   Call CIERRARCH("LIPRE001")
   J& = 0
   For KI% = 1 To CAGRU%
      Z$ = REGBLAN$("GRUCOVEN")
      Call REPLA(Z$, MKI$(KI%), 1, 2)
      Call REPLA(Z$, DEGRU$(KI%), 3, 30)
      J& = J& + 1
      Call GRAREG("GRUCOVEN", Z$, J&)
   Next KI%
   Call SETULTREG("GRUCOVEN", J&)
   Call CIERRARCH("GRUCOVEN")
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   '
   Call COMUNI("La Importación de Listas de Precios\se ha Completado Exitosamente.\  \Ejecute Ahora un Control del Maestro de Articulos.")
   Call FINAL("")
   '
Exit Sub
   '
GRACLIEN:
      XX$ = XX0$
      Call REPLA(XX$, MKI$(NCI%), 1, 2)
      Call REPLA(XX$, TRIM$(RASO$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, COPROVI$, 149, 1)
      Call REPLA(XX$, Chr$(CONDI%), 150, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COVEN%), 152, 1)
      Call REPLA(XX$, MKS$(LICRE), 153, 4)
      Call REPLA(XX$, MKI$(GRUCLI%), 157, 2)
      Call REPLA(XX$, MKI$(CUEMA%), 159, 2)
      '
      REGX& = REGX& + 1
      Call GRAREG("DEUDOR1", Left$(XX$, 32), REGX&)
      Call GRAREG("DEUDOR2", Mid$(XX$, 33), REGX&)
      '
      YY$ = YY0$
      Call REPLA(YY$, MKI$(NCI%), 1, 2)
      Call REPLA(YY$, TRIM$(RASO$), 3, 62)
      Call REPLA(YY$, TRIM$(FANT$), 65, 32)
      Call REPLA(YY$, TRIM$(DOMI$), 97, 64)
      Call REPLA(YY$, TRIM$(CPOSTX$), 161, 16)
      Call REPLA(YY$, TRIM$(LOCA$), 177, 48)
      Call REPLA(YY$, "ARGENTINA", 225, 32)
      Call REPLA(YY$, TRIM$(TELE$), 257, 24)
      Call REPLA(YY$, TRIM$(FAXX$), 281, 24)
      Call REPLA(YY$, TRIM$(EMAI$), 305, 48)
      Call REPLA(YY$, TRIM$(EXPRESO$), 353, 32)
      '
      REGY& = REGY& + 1
      Call GRAREG("BADACLIE", YY$, REGY&)
      '
      If NCI% > 0 Then
         Mid$(PUNUCLI$, NCI%, 1) = "1"
      End If
      '
      If NCI% > NCIMAX% Then NCIMAX% = NCI%
      '
Return
   '
   '
60  For KI% = 1 To CAPROVI%
      If DEPROVIN$(KI%) = DEPROV$ Then
         COPROVI$ = COPROVIN$(KI%)
         Return
      End If
    Next KI%
    COPROVI$ = "A"
601 For KI% = 1 To CAPROVI%
      If COPROVIN$(KI%) = COPROVI$ Then
         COPROVI$ = Chr$(1 + Asc(COPROVI$))
         GoTo 601
      End If
    Next KI%
    CAPROVI% = CAPROVI% + 1
    COPROVIN$(CAPROVI%) = COPROVI$
    DEPROVIN$(CAPROVI%) = DEPROV$
    XX$ = REGBLAN$("PROVINC")
    Call REPLA(XX$, COPROVI$, 1, 1)
    Call REPLA(XX$, DEPROV$, 2, 29)
    Call REGAPP("PROVINC", XX$)
Return
   '
   '
61 CUIT$ = Left$(CUITC$, 2) + "-" + Mid$(CUITC$, 3, 8) + "-" + Mid$(CUITC$, 11, 1)
   Return
   '
62 LCRX$ = REPLACAR$(LIMICREDI$, ".", "")
   LCRY$ = REPLACAR$(LCRX$, ",", ".")
   LICRE = Val(LCRY$)
   Return
   '
63  If VENDE$ = "ALCALAY" Then VENDE$ = "ALCALAYA"
    If TRIM$(VENDE$) = "0" Then VENDE$ = ""
    If TRIM$(VENDE$) = "" Then
       COVEN% = 0
       Return
    End If
    For KI% = 1 To CAVENDE%
      If DEVENDE$(KI%) = VENDE$ Then
         COVEN% = COVENDE%(KI%)
         Return
      End If
    Next KI%
    COVEN% = NUVENDE%
631 For KI% = 1 To CAVENDE%
      If COVENDE%(KI%) = COVEN% Then
         VENDE$ = DEVENDE$(KI%)
         Return
      End If
    Next KI%
    CAVENDE% = CAVENDE% + 1
    COVENDE%(CAVENDE%) = COVEN%
    DEVENDE$(CAVENDE%) = VENDE$
    XX$ = REGBLAN$("VENDEDOR")
    Call REPLA(XX$, Chr$(COVEN%), 1, 1)
    Call REPLA(XX$, VENDE$, 2, 30)
    Call REPLA(XX$, VENDE$, 32, 12)
    Call REGAPP("VENDEDOR", XX$)
Return
   '
   '
64  If TRIM$(CONDIPAG$) = "" Then
       CONDI% = 1
       Return
    End If
    For KI% = 1 To CACONDI%
      If DECONDI$(KI%) = CONDIPAG$ Then
         CONDI% = COCONDI%(KI%)
         Return
      End If
    Next KI%
    CONDI% = NUCONPA%
    CACONDI% = CACONDI% + 1
    COCONDI%(CACONDI%) = CONDI%
    DECONDI$(CACONDI%) = CONDIPAG$
    XX$ = REGBLAN$("CONPAG")
    Call REPLA(XX$, Chr$(CONDI%), 1, 1)
    Call REPLA(XX$, CONDIPAG$, 2, 35)
    Call REPLA(XX$, Chr$(1), 37, 1)
    Call REGAPP("CONPAG", XX$)
Return
   '
End Sub

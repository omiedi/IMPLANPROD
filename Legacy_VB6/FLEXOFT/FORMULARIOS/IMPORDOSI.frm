VERSION 5.00
Begin VB.Form IMPORDOSI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DOSIVAC - Importacion de Bases de Datos"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Control de Proceso"
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
      TabIndex        =   5
      Top             =   1470
      Width           =   6000
      Begin VB.Label Label2 
         Caption         =   "Label2"
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
         TabIndex        =   7
         Top             =   465
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   4935
         TabIndex        =   6
         Top             =   420
         Width           =   960
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C.C.Clie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   2205
      Picture         =   "IMPORDOSI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Base de Datos de Clientes"
      Top             =   420
      Width           =   850
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Piezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   4515
      Picture         =   "IMPORDOSI.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Base de Datos de Articulos de Stock"
      Top             =   420
      Width           =   850
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
      Height          =   850
      Left            =   5460
      Picture         =   "IMPORDOSI.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   420
      Width           =   850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Provedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1260
      Picture         =   "IMPORDOSI.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Base de Datos de Proveedores"
      Top             =   420
      Width           =   850
   End
   Begin VB.CommandButton COMMAND1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   315
      Picture         =   "IMPORDOSI.frx":09E0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Base de Datos de Clientes"
      Top             =   420
      Width           =   850
   End
End
Attribute VB_Name = "IMPORDOSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIENEMOVI%(32767)

Private Sub Command2_Click()
   Command2.Enabled = False
   Call IMPORPRODOSI
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call IMPORMASDOSI
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   Dim SALDO#(32767)
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\IMPORTA2\CUENTA_1.PRN", 1, 0)
   '
   JFAC& = 0: JCOB& = 0
   REBLAFAC$ = REGBLAN$("FACTUR")
   REBLACOB$ = REGBLAN$("COBRAN")
   Label2.Caption = "Importando Cuenta Corriente ..."
   Label2.Refresh
   '
   Line Input #N1%, AA$
   Screen.MousePointer = 11
   '
3  Line Input #N1%, AA$
   Label1.Caption = TRIM$(Str$(JFAC& + JCOB& + 1))
   Label1.Refresh
   If Left$(AA$, 12) = "haber_moneda" Then GoTo 99
   '
   FECHA$ = Left$(AA$, 10)
   TICO$ = Mid$(AA$, 21, 2)
   VACO$ = Mid$(AA$, 41, 1)
   NUCO$ = Mid$(AA$, 61, 7)
   IDEBE$ = Mid$(AA$, 68, 9)
   IHABE$ = Mid$(AA$, 77, 9)
   NPEDI$ = Mid$(AA$, 87, 4)
   NCLI$ = Mid$(AA$, 106, 5)
   RACLI$ = Mid$(AA$, 111, 36)
   '
   NCI% = Val(NCLI$)
   If NCI% > 10500 Then
      If NCI% <= 10999 Then
         NCI% = NCI% - 6000
      End If
   End If
   '
   FECO$ = Left$(FECHA$, 6) + Mid$(FECHA$, 9, 2)
   FF% = FECHANUM(FECO$)
   '
   TIX% = -1: SIG% = 1
   If TICO$ = "FC" Then TIX% = 2
   If TICO$ = "RC" Then TIX% = 4
   If TICO$ = "AH" Then TIX% = 0
   If TICO$ = "AD" Then TIX% = 0
   If TICO$ = "ND" Then TIX% = 3
   If TICO$ = "NC" Then TIX% = 6: SIG% = (-1)
   If TIX% < 0 Then
      Stop
   End If
   '
   VAX% = 0
   If VACO$ = "A" Then VAX% = 1
   If VACO$ = "B" Then VAX% = 2
   If VACO$ = "X" Then VAX% = 3
   '
   NRO& = Val(NUCO$)
   '
   IDEX$ = REPLACAR$(IDEBE$, ",", ".")
   IHAX$ = REPLACAR$(IHABE$, ",", ".")
   IMPO# = Val(IDEX$) - Val(IHAX$)
   If TIX% = 4 Then If IMPO# > 0 Then TIX% = 0
   If TIX% > 0 Then IMPO# = Abs(IMPO#)
   '
   If NCI% < 1 Then GoTo 90
   '
   TIENEMOVI%(NCI%) = 1
   NPEDID% = Val(NPEDI$)
   SUCUVEN% = 1
   '
   If TIX% <> 4 Then
      F$ = REBLAFAC$
      Call REPLA(F$, MKI$(NCI%), 1, 2)
      Call REPLA(F$, MKS$(NRO&), 3, 4)
      Call REPLA(F$, MKI$(FF%), 7, 2)
      Call REPLA(F$, MKI$(FF%), 9, 2)   ' fecha de vencimiento
      Call REPLA(F$, MKI$(NPEDID%), 11, 2)     ' lista de precios
      Call REPLA(F$, Chr$(TIX%) + Chr$(VAX%), 13, 2)
      '
      NEFA# = IMPO#
      NIVA# = 0
      NNIN# = 0
      TOTA# = IMPO#
      IINT# = 0
      PERC# = 0
      '
      SALDO#(NCI%) = SALDO#(NCI%) + IMPO# * SIG%
      If TIX% > 0 Then
        NEFA# = IMPO# / 1.21
        NIVA# = IMPO# - NEFA#
      End If
      '
      Call REPLA(F$, MKD$(NEFA#), 15, 8)
      Call REPLA(F$, MKD$(NIVA#), 23, 8)
      Call REPLA(F$, MKD$(NNIN#), 31, 8)
      Call REPLA(F$, MKD$(TOTA#), 39, 8)
      Call REPLA(F$, MKI$(SUCUVEN%), 47, 2)
      Call REPLA(F$, MKD$(IINT#), 49, 8)
      Call REPLA(F$, MKD$(PERC#), 57, 8)
      '
      JFAC& = JFAC& + 1
      Call GRAREG("FACTUR", F$, JFAC&)
      '
    ElseIf TIX% = 4 Then
      '
      SALDO#(NCI%) = SALDO#(NCI%) - IMPO#
      F$ = REBLACOB$
      Call REPLA(F$, MKI$(NCI%), 1, 2)
      Call REPLA(F$, MKS$(NRO&), 3, 4)
      Call REPLA(F$, MKI$(FF%), 7, 2)
      Call REPLA(F$, Chr$(TIX%) + Chr$(VA%), 9, 2)
      Call REPLA(F$, MKI$(NPEDID%), 11, 2)
      Call REPLA(F$, String$(2, 0), 13, 4)
      Call REPLA(F$, MKD$(IMPO#), 15, 8)
      Call REPLA(F$, MKI$(COBRACT%), 23, 2)
      '
      JCOB& = JCOB& + 1
      Call GRAREG("COBRAN", F$, JCOB&)
      '
   End If
   '
90 GoTo 3
   '
99 Call SETULTREG("FACTUR", JFAC&)
   Call SETULTREG("COBRAN", JCOB&)
   Call CIERRARCH("*.*")
   '
   For U& = 1 To ULTREG&("DEUDOR1")
     X$ = REGLEIDO$("DEUDOR1", U&)
     NCI% = CVI(Left$(X$, 2))
     If NCI% > 0 Then
       Y$ = REGLEIDO$("DEUDOR2", U&)
       Call REPLA(Y$, MKD$(SALDO#(NCI%)), 109, 8)
       Call GRAREG("DEUDOR2", Y$, U&)
     End If
   Next U&
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   '
   Call REIMPORCLIDOSI
   Call IMPORCOMPDEU
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
   Call VERJOBID
End Sub

Private Sub COMMAND1_Click()
   COMMAND1.Enabled = False
   Call IMPORCLIDOSI
   COMMAND1.Enabled = True
End Sub

Sub IMPORCLIDOSI()
   '
   If ULTREG&("DEUDOR1") > 0 Then
      Call MENSERR(24, "Importacion Imposible.\Ya Existe Archivo de Clientes y Distribuidores.")
      Exit Sub
   End If
   '
   Dim PUNUCLI As String * 32767
   '
   Dim NCLI%(32767)
   Dim COVENDE%(32), DEVENDE$(32)
   Dim COCONDI%(32), DECONDI$(32)
   Dim COPROVIN$(32), DEPROVIN$(32)
   Dim COGRU%(512), DEGRU$(512)
   '
   LU$ = LUDAT$
   '
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   YY0$ = REGBLAN$("BADACLIE")
   PIVA% = 1
   CATIB% = 2
   LIPRE% = 1
   CONDI% = 1
   CUEMA% = CUEINT%("113000")
   '
   Call REPLA(XX0$, Chr$(LIPRE%), 151, 1)
   Call REPLA(XX0$, Chr$(COPAGI%), 150, 1)
   '
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
      Call REPLA(XX$, "Condicion Habitual", 2, 35)
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
   CAGRUVE% = 0
   '
   REGX& = 0: REGY& = 0: REGZ& = 0: NCIMAX% = 0
   '
' IMPORTACION CLIENTES
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTE1.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTES.TXT", 0, 512)
   Dim BB As String * 512
   '
   OFSE% = 4588
   CONTA% = 0
3  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 3
   '
4  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 4
   '
   For CONTA% = 2 To OFSE%
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 350, 6))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      If NCI% < 1 Or NCI% > 9999 Then GoTo 8
      If NCI% > 0 Then
         If Mid$(PUNUCLI$, NCI%, 1) = "1" Then
            NCI% = 0
            GoTo 8
         End If
      End If
      '
      RASO$ = TRIM$(Mid$(AA$, 1, 36))
      FANT$ = Left$(RASO$, 10)
      DOMI$ = TRIM$(Mid$(AA$, 58, 48))
      CPOSTX$ = TRIM$(Mid$(AA$, 150, 4))
      CPOST% = Val(CPOSTX$)
      LOCA$ = TRIM$(Mid$(AA$, 106, 34))
      While Left$(LOCA$, 1) = "-": LOCA$ = Mid$(LOCA$, 2): Wend
      LOCA$ = TRIM$(LOCA$)
      'DEPROV$ = TRIM$(UCase$(Mid$(AA$, 106, 22))): GoSub 60
      TELE$ = TRIM$(Mid$(AA$, 154, 34))
      FAXNU$ = TRIM$(Mid$(AA$, 187, 31))
      CUITC$ = AJUSTI$(Mid$(AA$, 257, 11), 11): GoSub 61
      POSIV$ = TRIM$(Mid$(AA$, 275, 3))
      DESCU$ = TRIM$(Mid$(AA$, 304, 4))
      'EXPRESO$ = TRIM$(Mid$(AA$, 327, 28))
      'VENDE$ = TRIM$(UCase$(Mid$(AA$, 182, 7)))
      'NUVENDE% = Val(Mid$(AA$, 189, 2)): GoSub 63
      'CONDIPAG$ = TRIM$(UCase$(Mid$(AA$, 191, 21))): GoSub 64
      'LIMICREDI$ = TRIM$(Mid$(AA$, 212, 8)): GoSub 62
      '
      If TRIM$(RASO$) <> "" Then GoSub GRACLIEN
      '
   End If
   '
8   Next CONTA%
   '
9  Close #N1%: N1% = 0
   Close #N2%: N2% = 0
   '
   KK% = 0
   For U& = 1 To ULTREG&("DEUDOR1")
      XX$ = REGLEIDO$("DEUDOR1", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("DEUDOR1", XX$, U&)
      End If
   Next U&
   '
   KK% = 0
   For U& = 1 To ULTREG&("BADACLIE")
      XX$ = REGLEIDO$("BADACLIE", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("BADACLIE", XX$, U&)
         Mid$(PUNUCLI$, NCIMAX% + KK%, 1) = "1"
      End If
   Next U&
   '
   NCIMAX% = NCIMAX% + KK%
   Call SETULTREG("DEUDOR1", REGX&)
   Call SETULTREG("DEUDOR2", REGX&)
   Call SETULTREG("BADACLIE", REGY&)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   Call CIERRARCH("*.*")
   '
' IMPORTACION CLIENTE2
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTE2.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTES.TXT", 0, 512)
   '
   OFSE% = 4589
   CONTA% = 0
13 Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 13
   '
14 Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 14
   '
   For CONTA% = 2 To OFSE%
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 77, 5))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      If NCI% < 1 Or NCI% > 9999 Then GoTo 18
      '
      VENDE$ = TRIM$(UCase$(Mid$(AA$, 11, 21)))
      NUVENDE% = Val(Mid$(AA$, 1, 3)): GoSub 63
      GRUVE$ = TRIM$(Mid$(AA$, 43, 31))
      GRUCLI% = Val(Mid$(AA$, 32, 3)): GoSub 65
      DEPROV$ = TRIM$(UCase$(Mid$(AA$, 345, 22))): GoSub 60
      PAIS$ = TRIM$(Mid$(AA$, 367, 16)): If PAIS$ = "" Then PAIS$ = "ARGENTINA"
      FAXNU$ = TRIM$(Mid$(AA$, 383, 30))
      EMAIL$ = TRIM$(Mid$(AA$, 157 + 256, 28))
      CONTACTO$ = TRIM$(Mid$(AA$, 257, 34))
      '
      GoSub GRACLI2
      '
   End If
   '
18 Next CONTA%
   '
19 Close #N1%: N1% = 0
   Close #N2%: N2% = 0
   '
' IMPORTACION NOTAS DE CLIENTES
   '
20 Call BLANARCH("NOTACLI")
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLITEXTO.PRN", 1, 0)
   '
23 If EOF(N1%) Then GoTo 29
   Line Input #N1%, AA$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Left$(AA$, 6))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      If NCI% < 1 Or NCI% > 9999 Then GoTo 23
      TTXX$ = TRIM$(Mid$(AA$, 8))
      GoSub GRACLI3
   End If
   GoTo 23
   '
29 Call CIERRARCH("*.*")
   Close #N1%
   '
   Screen.MousePointer = 1
   Call COMUNI("La Importación de Clientes y Distribuidores\se ha Completado Exitosamente.")
   Exit Sub
   '
   '
30 N1% = FILEOPEN("A:PRECIOS.PRN", 1, 0)
   CI0% = NUMAS%
   Dim CODIMPOR$(16384)
   'Dim DEGRU$(256)
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
            GoTo 38
         End If
         CODLIN$ = LINACT$ + "-" + COD$
         '
         For KI% = 1 To CAIMPOR%
            If CODIMPOR$(KI%) = CODLIN$ Then
               Call COMUNI("Imposible Importar Codigo " + COD$ + " - Repetido.")
               GoTo 38
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
         If CI% > 0 Then GoTo 35
         '
         CI% = CODINT%(COD$)
         If CI% > 0 Then
            REMA& = CI% + 1
            REMAS$ = REGLEIDO$("MASTER", REMA&)
            Call REPLA(REMAS$, CODLIN$, 77, 16)
            Call GRAREG("MASTER", REMAS$, REMA&)
            GoTo 35
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
35       Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, DE$, 3, 42)
         Call REPLA(Z$, MKD$(PRE#), 62, 8)
         J& = J& + 1
         Call GRAREG("LIPRE001", Z$, J&)
      End If
38 Wend
   '
39 Close #N1%
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
      PIVA% = 0   ' consumidor final
      If POSIV$ <> "RI" Then
         AAA = BBB
      End If
      If POSIV$ = "RI" Then PIVA% = 1
      If POSIV$ = "RNI" Then PIVA% = 2
      If POSIV$ = "EXP" Then PIVA% = 5
      If POSIV$ = "REX" Then PIVA% = 4
      If POSIV$ = "EX" Then PIVA% = 4
      '
      XX$ = XX0$
      Call REPLA(XX$, MKI$(NCI%), 1, 2)
      Call REPLA(XX$, TRIM$(RASO$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, Chr$(PIVA%), 136, 1)
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
      Call REPLA(YY$, TRIM$(FAXNU$), 281, 24)
      Call REPLA(YY$, TRIM$(EXPRESO$), 353, 32)
      '
      REGY& = REGY& + 1
      Call GRAREG("BADACLIE", YY$, REGY&)
      '
      If Val(DESCU$) > 0 Then
         ZZ$ = REGBLAN$("DESCUCLI")
         Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
         Call REPLA(ZZ$, DESCU$, 3, 14)
         REGZ& = REGZ& + 1
         Call GRAREG("DESCUCLI", ZZ$, REGZ&)
      End If
      '
      If NCI% > 0 Then
         Mid$(PUNUCLI$, NCI%, 1) = "1"
      End If
      '
      If NCI% > NCIMAX% Then NCIMAX% = NCI%
      '
Return
   '
GRACLI2:
   '
   RFF$ = RECFILT$("DEUDOR1", 1, MKI$(NCI%))
   REGI& = CVS(RFF$)
   XX$ = REGLEIDO$("DEUDOR2", REGI&)
   Call REPLA(XX$, COPROVI$, 117, 1)
   Call REPLA(XX$, Chr$(NUVENDE%), 120, 1)
   Call REPLA(XX$, MKI$(GRUCLI%), 125, 2)
   Call GRAREG("DEUDOR2", XX$, REGI&)
   '
   RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCI%))
   REGI& = CVS(RFF$)
   YY$ = REGLEIDO$("BADACLIE", REGI&)
   Call REPLA(YY$, PAIS$, 225, 32)
   If TRIM$(Mid$(YY$, 281, 24)) = "" Then
      Call REPLA(YY$, TRIM$(FAXNU$), 281, 24)
   End If
   Call REPLA(YY$, EMAIL$, 305, 48)
   Call REPLA(YY$, TRIM$(CONTACTO$), 385, 32)
   Call GRAREG("BADACLIE", YY$, REGI&)
   '
Return
   '
GRACLI3:
   '
   If TRIM$(TTXX$) <> "" Then
     XX$ = REGBLAN$("NOTACLI")
     Call REPLA(XX$, MKI$(NCI%), 1, 2)
     Call REPLA(XX$, TTXX$, 3, 126)
     Call REGAPP("NOTACLI", XX$)
   End If
   '
Return

60  For KI% = 1 To CAPROVI%
      If DEPROVIN$(KI%) = DEPROV$ Then
         COPROVI$ = COPROVIN$(KI%)
         Return
      End If
    Next KI%
    COPROVI$ = " "
Return
   '
   '
61 CUIT$ = ""
   If TRIM$(CUITC$) <> "" Then
      CUIT$ = Left$(CUITC$, 2) + "-" + Mid$(CUITC$, 3, 8) + "-" + Mid$(CUITC$, 11, 1)
   End If
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
    CONDI% = 1
641 For KI% = 1 To CACONDI%
      If COCONDI%(KI%) = CONDI% Then
         CONDI% = CONDI% + 1
         GoTo 641
      End If
    Next KI%
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
65  For KI% = 1 To CAGRUVE%
      If COGRU%(KI%) = GRUCLI% Then Return
    Next KI%
    If GRUCLI% > 0 Then
       CAGRUVE% = CAGRUVE% + 1
       COGRU%(CAGRUVE%) = GRUCLI%
       DEGRU$(CAGRUVE%) = GRUVE$
       XX$ = REGBLAN$("GRUCOCLI")
       Call REPLA(XX$, MKI$(GRUCLI%), 1, 2)
       Call REPLA(XX$, GRUVE$, 3, 30)
       U& = CAGRUVE%
       Call GRAREG("GRUCOCLI", XX$, U&)
       Call SETULTREG("GRUCOCLI", U&)
       Call CIERRARCH("GRUCOCLI")
    End If
Return
   '

End Sub


Sub IMPORPRODOSI()

   If ULTREG&("PROVED1") > 0 Then
      Call MENSERR(24, "Importacion Imposible.\Ya Existe Archivo de Proveedores.")
      Exit Sub
   End If
   '
   Dim PUNUCLI As String * 32767
   '
   Dim NCLI%(32767)
   Dim COVENDE%(32), DEVENDE$(32)
   Dim COCONDI%(32), DECONDI$(32)
   Dim COPROVIN$(32), DEPROVIN$(32)
   Dim COGRU%(512), DEGRU$(512)
   '
   LU$ = LUDAT$
   '
   XX0$ = REGBLAN$("PROVED1") + REGBLAN$("PROVED2")
   YY0$ = REGBLAN$("BADAPROV")
   PIVA% = 1
   CATIB% = 2
   LIPRE% = 0    ' CUENTA CORRIENTE
   CONDI% = 1
   CUEMA% = CUEINT%("211100")
   '
   Call REPLA(XX0$, Chr$(LIPRE%), 151, 1)
   Call REPLA(XX0$, Chr$(COPAGI%), 150, 1)
   '
   REGX& = 0: REGY& = 0: REGZ& = 0: NCIMAX% = 0
   '
' IMPORTACION PROVEEDORES
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\PROVEDO1.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\PROVEDOR.TXT", 0, 512)
   Dim BB As String * 512
   '
   OFSE% = 1192
   CONTA% = 0
3  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 3
   '
4  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 4
   '
   For CONTA% = 2 To OFSE%
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 1, 5))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% < 1 Or NCI% > 9999 Then GoTo 8
      If NCI% > 0 Then
         If Mid$(PUNUCLI$, NCI%, 1) = "1" Then
            NCI% = 0
            GoTo 8
         End If
      End If
      '
      RASO$ = TRIM$(Mid$(AA$, 28, 36))
      FANT$ = Left$(RASO$, 10)
      DOMI$ = TRIM$(Mid$(AA$, 101, 56))
      CPOSTX$ = TRIM$(Mid$(AA$, 256 + 37, 4))
      CPOST% = Val(CPOSTX$)
      LOCA$ = TRIM$(Mid$(AA$, 181, 34))
      While Left$(LOCA$, 1) = "-": LOCA$ = Mid$(LOCA$, 2): Wend
      LOCA$ = TRIM$(LOCA$)
      'DEPROV$ = TRIM$(UCase$(Mid$(AA$, 106, 22))): GoSub 60
      TELE$ = TRIM$(Mid$(AA$, 64, 21))
      FAXNU$ = TRIM$(Mid$(AA$, 214, 43))
      'CUITC$ = AJUSTI$(Mid$(AA$, 257, 11), 11): GoSub 61
      'POSIV$ = TRIM$(Mid$(AA$, 275, 3))
      'DESCU$ = TRIM$(Mid$(AA$, 304, 4))
      'EXPRESO$ = TRIM$(Mid$(AA$, 327, 28))
      'VENDE$ = TRIM$(UCase$(Mid$(AA$, 182, 7)))
      'NUVENDE% = Val(Mid$(AA$, 189, 2)): GoSub 63
      'CONDIPAG$ = TRIM$(UCase$(Mid$(AA$, 191, 21))): GoSub 64
      'LIMICREDI$ = TRIM$(Mid$(AA$, 212, 8)): GoSub 62
      '
      If TRIM$(RASO$) <> "" Then GoSub GRAPROVE
      '
   End If
   '
8   Next CONTA%
   '
9  Close #N1%: N1% = 0
   Close #N2%: N2% = 0
   '
   Call SETULTREG("PROVED1", REGX&)
   Call SETULTREG("PROVED2", REGX&)
   Call SETULTREG("BADAPROV", REGY&)
   Call CIERRARCH("PROVED1")
   Call CIERRARCH("PROVED2")
   Call CIERRARCH("BADAPROV")
   Call CIERRARCH("*.*")
   '
' IMPORTACION PROVEE
   '
   Call BLANARCH("NOTAPRO")
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\PROVEE.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\PROVEDOR.TXT", 0, 512)
   '
   OFSE% = 1115
   CONTA% = 0
13 Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 13
   '
14 Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 14
   '
   For CONTA% = 2 To OFSE%
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 1, 5))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% < 1 Or NCI% > 9999 Then GoTo 18
      '
      CUITC$ = AJUSTI$(Mid$(AA$, 256 + 11, 11), 11): GoSub 61
      POSIV$ = TRIM$(Mid$(AA$, 256 + 22, 3))
      TTXX$ = TRIM$(Mid$(AA$, 163, 94))
      '
      GoSub GRAPRO2
      '
   End If
   '
18 Next CONTA%
   '
19 Close #N1%: N1% = 0
   Close #N2%: N2% = 0
   '
' IMPORTACION NOTAS DE PROVEEDORES
   '
20 N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\PROTEXTO.PRN", 1, 0)
   '
23 If EOF(N1%) Then GoTo 29
   Line Input #N1%, AA$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Left$(AA$, 6))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      If NCI% < 1 Or NCI% > 9999 Then GoTo 23
      TTXX$ = TRIM$(Mid$(AA$, 8))
      GoSub GRAPRO3
   End If
   GoTo 23
   '
29 Call CIERRARCH("*.*")
   Close #N1%
   '
   Screen.MousePointer = 1
   Call COMUNI("La Importación de Proveedores\se ha Completado Exitosamente.")
   Exit Sub
   '
   '
GRAPROVE:
      '
      XX$ = XX0$
      Call REPLA(XX$, MKI$(NCI%), 1, 2)
      Call REPLA(XX$, TRIM$(RASO$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, COPROVI$, 149, 1)
      Call REPLA(XX$, Chr$(CONDI%), 150, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COVEN%), 152, 1)
      Call REPLA(XX$, MKS$(LICRE), 153, 4)
      Call REPLA(XX$, MKI$(GRUCLI%), 157, 2)
      Call REPLA(XX$, MKI$(CUEMA%), 159, 2)
      '
      REGX& = REGX& + 1
      Call GRAREG("PROVED1", Left$(XX$, 32), REGX&)
      Call GRAREG("PROVED2", Mid$(XX$, 33), REGX&)
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
      Call REPLA(YY$, TRIM$(FAXNU$), 281, 24)
      Call REPLA(YY$, TRIM$(EXPRESO$), 353, 32)
      '
      REGY& = REGY& + 1
      Call GRAREG("BADAPROV", YY$, REGY&)
      '
      If NCI% > 0 Then
         Mid$(PUNUCLI$, NCI%, 1) = "1"
      End If
      '
      If NCI% > NCIMAX% Then NCIMAX% = NCI%
      '
Return
   '
GRAPRO2:
   '
   PIVA% = 0   ' consumidor final
   If POSIV$ = "RI" Then PIVA% = 1
   If POSIV$ = "RNI" Then PIVA% = 2
   If POSIV$ = "EXP" Then PIVA% = 5
   If POSIV$ = "REX" Then PIVA% = 4
   If POSIV$ = "EX" Then PIVA% = 4
   '
   RFF$ = RECFILT$("PROVED1", 1, MKI$(NCI%))
   REGI& = CVS(RFF$)
   XX$ = REGLEIDO$("PROVED2", REGI&)
   Call REPLA(XX$, TRIM$(CUIT$), 75, 15)
   Call REPLA(XX$, Chr$(PIVA%), 120, 1)
   Call GRAREG("PROVED2", XX$, REGI&)
   '
   If TTXX$ <> "" Then GoSub GRAPRO3
   '
Return
   '
GRAPRO3:
   '
   If TRIM$(TTXX$) <> "" Then
     XX$ = REGBLAN$("NOTAPRO")
     Call REPLA(XX$, MKI$(NCI%), 1, 2)
     Call REPLA(XX$, TTXX$, 3, 126)
     Call REGAPP("NOTAPRO", XX$)
   End If
   '
Return

   '
61 CUIT$ = ""
   If TRIM$(CUITC$) <> "" Then
      CUIT$ = Left$(CUITC$, 2) + "-" + Mid$(CUITC$, 3, 8) + "-" + Mid$(CUITC$, 11, 1)
   End If
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
    CONDI% = 1
641 For KI% = 1 To CACONDI%
      If COCONDI%(KI%) = CONDI% Then
         CONDI% = CONDI% + 1
         GoTo 641
      End If
    Next KI%
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


Sub IMPORMASDOSI()
   '
   If NUMAS% > 0 Then
      Call MENSERR(24, "Importacion Imposible.\Ya Existe Archivo de Artículos de Stock.")
      Exit Sub
   End If
   '
   Dim TXX$(128)
   Dim STC$(8192)
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\ARTICULO.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTES.TXT", 0, 512)
   Dim BB As String * 512
   '
   OFSE% = 5124
   CONTA% = 0
3  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 3
   '
4  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 4
   '
   JJ& = 1: KK& = 0
   For CONTA% = 2 To OFSE%
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   '
30 COD$ = TRIM$(Left$(AA$, 22))
   If Len(COD$) > 15 Then
      cod1$ = COD$
      COD$ = REPLACAR$(cod1$, "-", "")
   End If
   If Len(COD$) > 15 Then
      cod1$ = COD$
      COD$ = REPLACAR$(cod1$, ".", "")
   End If
   '
   If Len(COD$) > 15 Then GoTo 39
   '
   DE$ = TRIM$(Mid$(AA$, 23, 37))
   UNI$ = TRIM$(UCase$(Mid$(AA$, 76, 4)))
   If UNI$ = "C/U" Or UNI$ = "C//U" Then UNI$ = "U."
   
If UNI$ <> "PAR" And UNI$ <> "U" And UNI$ <> "U." And UNI$ <> "" And UNI$ <> "BOMB" Then
AAA = BBB
End If
   If UNI$ = "M" Then UNI$ = "M."
   If UNI$ = "MTS." Then UNI$ = "M."
   If UNI$ = "GR." Then UNI$ = "Gr"
   If UNI$ = "LT" Or UNI$ = "LTS" Then UNI$ = "Lt"
   If UNI$ = "L" Then UNI$ = "Lt"
   If UNI$ = "KG." Then UNI$ = "Kg"
   COSUX$ = TRIM$(Mid$(AA$, 134, 12))
   COU = Val(REPLACAR$(COSUX$, ",", "."))
   PREUX$ = TRIM$(Mid$(AA$, 146, 14))
   PREVEN# = Val(REPLACAR$(PREUX$, ",", "."))
   MONEI% = 1
   TIMA% = 1: If PREVEN# > 0.005 Then TIMA% = 5
   '
   STOMI = Val(Mid$(AA$, 160, 11))
   If STOMI < Val(Mid$(AA$, 194, 14)) Then STOMI = Val(Mid$(AA$, 194, 14))
   STMII% = 0: If STOMI < 32000 Then STMII% = STOMI
   LOREI% = 1
   STOMA = Val(Mid$(AA$, 171, 12))
   LOTEMAI% = STMII%
   If STOMA < 32000 Then
     If STOMA > STMII% Then
       LOTEMAI% = STOMA
     End If
   End If
   '
   FECO% = FECHANUM(Mid$(AA$, 222, 6) + Mid$(AA$, 230, 2))
   PROVEX = Val(Mid$(AA$, 256 + 32, 5))
   NPRO% = 0: If PROVEX < 32767 Then NPRO% = PROVEX
   FULT% = FECO%
   '
   If Len(COD$) >= 3 Then
      CI% = JJ&
      REMAS$ = REGBLAN$("MASTER")
      Call REPLA(REMAS$, MKI$(CI%), 1, 2)
      Call REPLA(REMAS$, DE$, 5, 64)
      Call REPLA(REMAS$, String$(8, 0), 69, 8)
      Call REPLA(REMAS$, COD$, 77, 16)
      Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
      Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
      Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
      Call REPLA(REMAS$, String$(7, 0), 96, 7)
      Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
      Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
      Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
      Call REPLA(REMAS$, UNI$, 109, 2)
      Call REPLA(REMAS$, MKS$(COU), 111, 4)
      Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
      Call REPLA(REMAS$, MKS$(PPP), 117, 4)
      Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
      Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
      Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
      Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
      '
      JJ& = JJ& + 1
      Call GRAREG("MASTER", REMAS$, JJ&)
      X$ = REGLEIDO$("MASTER", 1&)
      Call REPLA(X$, MKI$(CI%), 1, 2)
      Call GRAREG("MASTER", X$, 1&)
      '
      If PREVEN# >= 0.005 Then
         Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, DE$, 3, 42)
         Call REPLA(Z$, MKD$(PREVEN#), 62, 8)
         KK& = KK& + 1
         Call GRAREG("LIPRE001", Z$, KK&)
      End If
   End If
   '
39 Next CONTA%
   '
   Call SETULTREG("LIPRE001", KK&)
   Call CIERRARCH("*.*")
   Close

End Sub

Sub REIMPORCLIDOSI()
   '
   Dim PUNUCLI As String * 32767
   '
   Dim NCLI%(32767)
   Dim COVENDE%(32), DEVENDE$(32)
   Dim COCONDI%(32), DECONDI$(32)
   Dim COPROVIN$(32), DEPROVIN$(32)
   Dim COGRU%(512), DEGRU$(512)
   '
   LU$ = LUDAT$
   '
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   YY0$ = REGBLAN$("BADACLIE")
   PIVA% = 1
   CATIB% = 2
   LIPRE% = 1
   CONDI% = 1
   CUEMA% = CUEINT%("113000")
   '
   Call REPLA(XX0$, Chr$(LIPRE%), 151, 1)
   Call REPLA(XX0$, Chr$(COPAGI%), 150, 1)
   '
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
      Call REPLA(XX$, "Condicion Habitual", 2, 35)
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
   For U& = 1 To ULTREG&("VENDEDOR")
      XX$ = REGLEIDO$("VENDEDOR", U&)
      CAVENDE% = CAVENDE% + 1
      COVENDE%(CAVENDE%) = Asc(Left$(XX$, 1))
      DEVENDE$(CAVENDE%) = TRIM$(UCase$(Mid$(XX$, 2, 30)))
   Next U&
   '
   CAGRUVE% = 0
   '
   REGX& = 0: REGY& = 0: REGZ& = 0: NCIMAX% = 0
   '
' IMPORTACION CLIENTES
   '
   Label2.Caption = "Importando Clientes Agregados ..."
   Label2.Refresh
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\IMPORTA2\CLIENT_1.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\CLIENTES.TXT", 0, 512)
   Dim BB As String * 512
   '
   OFSE% = 4625
   CONTA% = 0
3  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Label1.Caption = TRIM$(Str$(CONTA%))
   Label1.Refresh
   '
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 3
   '
4  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 4
   '
   For CONTA% = 2 To OFSE%
   Label1.Caption = TRIM$(Str$(CONTA%))
   Label1.Refresh
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 256 + 119, 5))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      If NCI% > 0 Then
         If Mid$(PUNUCLI$, NCI%, 1) = "1" Then
            NCI% = 0
            GoTo 8
         End If
      End If
      '
      If (NCI% > 3100 And NCI% < 3300) Or TIENEMOVI%(NCI%) > 0 Then
         RASO$ = TRIM$(Mid$(AA$, 1, 46))
         FANT$ = Left$(RASO$, 10)
         DOMI$ = TRIM$(Mid$(AA$, 58, 48))
         CPOSTX$ = TRIM$(Mid$(AA$, 150, 4))
         CPOST% = Val(CPOSTX$)
         LOCA$ = TRIM$(Mid$(AA$, 106, 34))
         While Left$(LOCA$, 1) = "-": LOCA$ = Mid$(LOCA$, 2): Wend
         LOCA$ = TRIM$(LOCA$)
         'DEPROV$ = TRIM$(UCase$(Mid$(AA$, 106, 22))): GoSub 60
         TELE$ = TRIM$(Mid$(AA$, 154, 34))
         FAXNU$ = TRIM$(Mid$(AA$, 188, 34))
         CUITC$ = AJUSTI$(Mid$(AA$, 256 + 25, 11), 11): GoSub 61
         POSIV$ = TRIM$(Mid$(AA$, 256 + 43, 3))
         DESCU$ = TRIM$(Mid$(AA$, 256 + 72, 4))
         'EXPRESO$ = TRIM$(Mid$(AA$, 327, 28))
         'VENDE$ = TRIM$(UCase$(Mid$(AA$, 182, 7)))
         'NUVENDE% = Val(Mid$(AA$, 189, 2)): GoSub 63
         'CONDIPAG$ = TRIM$(UCase$(Mid$(AA$, 191, 21))): GoSub 64
         'LIMICREDI$ = TRIM$(Mid$(AA$, 212, 8)): GoSub 62
         '
         If TRIM$(RASO$) <> "" Then GoSub GRACLIEN
      End If
      '
   End If
   '
8  Next CONTA%
   '
9  Close #N1%: N1% = 0
   Close #N2%: N2% = 0
   '
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   Call CIERRARCH("*.*")
   '
Exit Sub
'
GRACLIEN:
   '
   RFF$ = RECFILT$("DEUDOR1", 1, MKI$(NCI%))
   If Len(RFF$) > 0 Then GoTo 90
   '
      PIVA% = 0   ' consumidor final
      If POSIV$ = "RI" Then PIVA% = 1
      If POSIV$ = "RNI" Then PIVA% = 2
      If POSIV$ = "EXP" Then PIVA% = 5
      If POSIV$ = "REX" Then PIVA% = 4
      If POSIV$ = "EX" Then PIVA% = 4
      '
      XX$ = XX0$
      Call REPLA(XX$, MKI$(NCI%), 1, 2)
      Call REPLA(XX$, TRIM$(RASO$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, Chr$(PIVA%), 136, 1)
      Call REPLA(XX$, COPROVI$, 149, 1)
      Call REPLA(XX$, Chr$(CONDI%), 150, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COVEN%), 152, 1)
      Call REPLA(XX$, MKS$(LICRE), 153, 4)
      Call REPLA(XX$, MKI$(GRUCLI%), 157, 2)
      Call REPLA(XX$, MKI$(CUEMA%), 159, 2)
      '
      REGX& = 1 + ULTREG&("DEUDOR1")
      Call GRAREG("DEUDOR1", Left$(XX$, 32), REGX&)
      For REDEU2& = 1 + ULTREG&("DEUDOR2") To REGX&
        Call GRAREG("DEUDOR2", " " + String$(127, 0), REDEU2&)
      Next REDEU2&
      Call GRAREG("DEUDOR2", Mid$(XX$, 33), REGX&)
      '
      REBADA& = 0
      RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCI%))
      If Len(RFF$) >= 4 Then
         REBADA& = CVS(Left$(RFF$, 4))
      End If
      If REBADA& < 1 Then REBADA& = 0
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
      Call REPLA(YY$, TRIM$(FAXNU$), 281, 24)
      Call REPLA(YY$, TRIM$(EXPRESO$), 353, 32)
      '
      If REBADA& < 1 Then REBADA& = 1 + ULTREG&("BADACLIE")
      Call GRAREG("BADACLIE", YY$, REBADA&)
      '
      If Val(DESCU$) > 0 Then
         ZZ$ = REGBLAN$("DESCUCLI")
         Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
         Call REPLA(ZZ$, DESCU$, 3, 14)
         REGZ& = REGZ& + 1
         Call GRAREG("DESCUCLI", ZZ$, REGZ&)
      End If
      '
90 Return
    '
60  For KI% = 1 To CAPROVI%
      If DEPROVIN$(KI%) = DEPROV$ Then
         COPROVI$ = COPROVIN$(KI%)
         Return
      End If
    Next KI%
    COPROVI$ = " "
Return
   '
   '
61 CUIT$ = ""
   If TRIM$(CUITC$) <> "" Then
      CUIT$ = Left$(CUITC$, 2) + "-" + Mid$(CUITC$, 3, 8) + "-" + Mid$(CUITC$, 11, 1)
   End If
   Return
   '
   '
End Sub

Sub IMPORCOMPDEU()

   ZZ0$ = REGBLAN$("SVD202")
   JSAL& = 0
   '
' IMPORTACION SALDOS DEUDORES
   '
   Label2.Caption = "Importando Saldos Deudores"
   Label2.Refresh
   '
   N1% = FILEOPEN("C:\CLIENTES\DOSIVAC\IMPORTA2\COMPOS_1.PRN", 1, 0)
   N2% = FILEOPEN("C:\CLIENTES\DOSIVAC\TRAIDOS\COMPODEU.TXT", 0, 512)
   Dim BB As String * 512
   '
   OFSE% = 3077
   CONTA% = 0
3  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Label1.Caption = TRIM$(Str$(CONTA%))
   Label1.Refresh
   '
   RTRA$ = Space$(512)
   Call REPLA(RTRA$, AA$, 1, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA%, BB$
   If CONTA% < OFSE% Then GoTo 3
   '
4  Line Input #N1%, AA$
   CONTA% = CONTA% + 1
   Get #N2%, CONTA% - OFSE%, BB$
   RTRA$ = BB$
   Call REPLA(RTRA$, AA$, 257, 256)
   LSet BB$ = RTRA$
   Put #N2%, CONTA% - OFSE%, BB$
   If CONTA% < 2 * OFSE% Then GoTo 4
   '
   For CONTA% = 2 To OFSE%
   Label1.Caption = TRIM$(Str$(CONTA%))
   Label1.Refresh
   Get #N2%, CONTA%, BB$
   AA$ = BB$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 24, 5))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      NCI% = NCX&
      If NCI% > 10500 Then
         If NCI% <= 10999 Then
            NCI% = NCI% - 6000
         End If
      End If
      '
      FECHA$ = Left$(AA$, 10)
      TICO$ = Mid$(AA$, 94, 2)
      NUCO$ = Mid$(AA$, 114, 6)
      IDEBE$ = Mid$(AA$, 120, 8)
      IHABE$ = Mid$(AA$, 128, 8)
      ISALD$ = Mid$(AA$, 141, 14)
      RACLI$ = Mid$(AA$, 29, 64)
      FEVEN$ = Mid$(AA$, 256 + 103, 10)
      NPEDI$ = Mid$(AA$, 256 + 139, 4)
      NCII% = NCI%
      '
      If NCI% <> NCIA% Then
         NCIA% = NCI%
         RASOA$ = RASOCLI$(NCI%)
         If TRIM$(RACLI$) <> TRIM$(RASOA$) Then
            REDEU& = REGICLI&(NCI%)
            If REDEU& > 0 Then
              XXX$ = REGLEIDO$("DEUDOR1", REDEU&)
              Call REPLA(XXX$, RACLI$, 3, 30)
              Call GRAREG("DEUDOR1", XXX$, REDEU&)
            End If
            RFFX$ = RECFILT$("BADACLIE", 1, MKI$(NCI%))
            If Len(RFFX$) >= 4 Then
              REDEU& = CVS(Left$(RFFX$, 4))
              If REDEU& > 0 Then
                 XXX$ = REGLEIDO$("BADACLIE", REDEU&)
                 Call REPLA(XXX$, RACLI$, 3, 62)
                 Call GRAREG("BADACLIE", XXX$, REDEU&)
              End If
            End If
         End If
      End If
      '
      NCI% = NCII%
      NRO& = Val(NUCO$)
      On Error Resume Next
      NPEDX% = Val(NPEDI$)
      On Error GoTo 0
      '
      FECO$ = Left$(FECHA$, 6) + Mid$(FECHA$, 9, 2)
      FF% = FECHANUM(FECO$)
      FECO$ = Left$(FEVEN$, 6) + Mid$(FEVEN$, 9, 2)
      FVI% = FECHANUM(FECO$)
      '
      TIX% = -1
      If TICO$ = "FC" Then TIX% = 2: VAX% = 1
      If TICO$ = "RC" Then TIX% = 4: VAX% = 0
      If TICO$ = "AH" Then TIX% = 0: VAX% = 0
      If TICO$ = "AD" Then TIX% = 0: VAX% = 0
      If TICO$ = "ND" Then TIX% = 3: VAX% = 1
      If TICO$ = "NC" Then TIX% = 6: VAX% = 1
      If TIX% < 0 Then
         'Stop
      End If
      '
      IDEX$ = REPLACAR$(IDEBE$, ",", ".")
      IHAX$ = REPLACAR$(IHABE$, ",", ".")
      ISAX$ = REPLACAR$(ISALD$, ",", ".")
      '
      TOTAVAL# = Val(IDEX$) + Val(IHAX$)
      If TIX% = 4 Then If TOTAVAL# > 0 Then TIX% = 0
      ANTI# = Val(ISAX$)
      '
      If TIX% > 0 Then
         TOTAVAL# = Abs(TOTAVAL#)
         ANTI# = Abs(ANTI#)
      End If
      '
      If NCI% < 1 Then GoTo 90
      If Abs(ANTI#) < 0.005 Then GoTo 90
      '
      Z$ = ZZ0$
      Call REPLA(Z$, MKI$(NCI%), 1, 2)
      Call REPLA(Z$, Chr$(TIX%), 3, 1)
      Call REPLA(Z$, Chr$(0), 4, 1)
      Call REPLA(Z$, MKS$(NRO&), 5, 4)
      Call REPLA(Z$, MKI$(FVI%), 9, 2)
      Call REPLA(Z$, MKD$(ANTI#), 11, 8)
      Call REPLA(Z$, MKI$(FF%), 19, 2)
      Call REPLA(Z$, MKD$(TOTAVAL#), 21, 8)
      Call REPLA(Z$, MKI$(NPEDX%), 29, 2)
      Call REPLA(Z$, String$(2, 255), 31, 2)
      '
      JSAL& = JSAL& + 1
      Call GRAREG("SVD202", Z$, JSAL&)
      '
   End If
   '
90 Next CONTA%
   '
   Call SETULTREG("SVD202", JSAL&)
   Call CIERRARCH("SVD202")
   '
End Sub

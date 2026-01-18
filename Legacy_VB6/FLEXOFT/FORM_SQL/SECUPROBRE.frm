VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SECUPROC09 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de Procesos - Batch Record"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   7320
      ScaleHeight     =   7395
      ScaleWidth      =   1320
      TabIndex        =   15
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command3 
         Height          =   600
         Left            =   105
         Picture         =   "SECUPROBRE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Refrescar Tiempos Standard para Programacion"
         Top             =   2800
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Height          =   600
         Left            =   105
         Picture         =   "SECUPROBRE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Emite Batch Record por Impresora"
         Top             =   3945
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   600
         Left            =   105
         Picture         =   "SECUPROBRE.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Refrescar Secuencia de Operaciones para la Orden de Produccion Activa"
         Top             =   2200
         Width           =   645
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   780
         Left            =   120
         Picture         =   "SECUPROBRE.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guardar Cierre de O/F y Reingreso de Materiales a O/F"
         Top             =   6420
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   6180
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   -135
            TabIndex        =   21
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command15 
         Height          =   600
         Left            =   120
         Picture         =   "SECUPROBRE.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ficha Kardex de Movimientos de Stock"
         Top             =   5040
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Height          =   600
         Left            =   120
         Picture         =   "SECUPROBRE.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SECUPROBRE.frx":1354
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5730
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Secuencia de Procesos"
      Height          =   5565
      Left            =   105
      TabIndex        =   9
      Top             =   1680
      Width           =   7110
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   5000
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   5000
         Left            =   6550
         TabIndex        =   11
         Top             =   375
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   5000
         Left            =   4320
         TabIndex        =   10
         Top             =   360
         Width           =   30
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         ItemData        =   "SECUPROBRE.frx":3276
         Left            =   210
         List            =   "SECUPROBRE.frx":3278
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   645
         Width           =   6630
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  N-O    Descripcion                                         H-C-Proc       "
         Height          =   270
         Left            =   225
         TabIndex        =   14
         Top             =   375
         Width           =   6630
      End
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
      Left            =   3225
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SECUPROBRE.frx":327A
      TabIndex        =   27
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   525
      OleObjectBlob   =   "SECUPROBRE.frx":32F4
      Top             =   420
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Proyectada:"
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   795
      Width           =   1545
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5730
      TabIndex        =   17
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   750
      Width           =   1395
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   300
      Left            =   -150
      TabIndex        =   8
      Top             =   790
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Prod. Nº:"
      Height          =   300
      Left            =   180
      TabIndex        =   7
      Top             =   355
      Width           =   1125
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   750
      Width           =   1965
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1155
      Width           =   5680
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   300
      Left            =   210
      TabIndex        =   4
      Top             =   1195
      Width           =   1125
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   2
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro S/F:"
      Height          =   300
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "SECUPROC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    NORPRO$ = Label32
    SOLFA$ = Label26
    If SOLFA$ <> "" Then
       If DERACCE%("GENSECOP", "Re-Generación Secuencia Standard de Operaciones") >= 2 Then
          OPP% = 3     ' SI NADA SELECCIONADO, BORRA SECUENCIA COMPLETA
          If List1.ListIndex >= 0 Then
             TTXX$ = "Puede Refrescar Solo la Hoja Correspondiente\"
             TTXX$ = TTXX$ + "a la Operacion Activa (" + TRIM$(Left$(List1.TEXT, 38)) + ")\"
             TTXX$ = TTXX$ + "o la Secuencia Completa del el Producto '" + TRIM$(Label28) + "'."
             OPP% = COMALTER%(TTXX$ + "\\Cancelar\Solo Activa\Secuencia Completa")
          End If
          '
          If OPP% = 2 Then
               HOCOPRO$ = TRIM$(Mid$(List1.TEXT, 39, 16))
               TTXX$ = "Confirma Refrescar la Hoja de Control de Proceso '" + HOCOPRO$ + "'\"
               TTXX$ = TTXX$ + "Correspondiente a la Operación " + TRIM$(Left$(List1.TEXT, 38)) + " ?\   \"
               TTXX$ = TTXX$ + "ATENCION: Esta Transacción Borrará Toda Información\Cargada en esa Hoja de Batch Record\y NO PODRÁ REVERTIRSE."
               If COMALTER%(TTXX$ + "\\No, Conservar Actual\Si, Refrescar") = 2 Then
                  NOPE% = XVALO(Left$(List1.TEXT, 5))
                  Call BORRAHOPRO(SOLFA$, NOPE%)
               End If
            ElseIf OPP% = 3 Then
               TTXX$ = "Confirma Refrescar la Secuencia de Operaciones\Standard para esta Orden de Producción ?\   \ATENCION: Esta Transacción Borrará Toda Información\Cargada en las Hojas de Batch Record\y las Reemplazará por la Secuencia Actualmente Vigente\para el Producto '" + TRIM$(Label28) + " - " + TRIM$(Label30) + "'"
               If COMALTER%(TTXX$ + "\\No, Conservar Actual\Si, Refrescar") = 2 Then
                 Call REGENOPER(Label26)
                 Call Label32_Change
               End If
          End If
       End If
    End If
99  Command1.Enabled = True
End Sub

Private Sub Command15_Click()
    FIKARDEX09.Text1 = Label20
    FIKARDEX09.Show 1
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    SOLFA$ = Label26
    If SOLFA$ <> "" Then
      If DERACCE%("PRIBATCH", "Impresión de Batch Record") >= 2 Then
        Call ABREORFAB
        CIXI% = CODINT%(Label28)
        If CIXI% < 1 Then GoTo 99
        '
        PRICOLANT% = PRINTERCONTROL%
        PRINTERCONTROL% = 9999                 ' PARA QUE CIERRE CON NEWPAGE
        For U& = 1 To List1.ListCount
            NOPE% = XVALO(Left$(List1.List(U& - 1), 5))
            CONDI$ = "IdSubOr = '" & SOLFA$ & "' AND NumeOper = " & NOPE%
            FILTX$ = SOLFA$ & ";" & CIXI% & ";" & NOPE%
            '
            FORIMPR$ = TRIM$(Mid$(List1.List(U& - 1), 40, 16))
            If FORIMPR$ <> "" Then
               FOGUAR$ = VALOBADA("OPERFAB", "HOBATCHREC", CONDI$)
               If TRIM$(FOGUAR$) <> "" Then
                  Call SAVEFILE(LUCOM$ + FORIMPR$ + ".SQX", FOGUAR$)
                  FORIMPR$ = "!" + FORIMPR$
                  FILTX$ = ""
               End If
               Call STDFORM(FORIMPR$, FILTX$, "PRINT")
            End If
        Next U&
        On Error Resume Next
        Printer.EndDoc
        On Error GoTo 0
        PRINTERCONTROL% = PRICOLANT%
      End If
    End If
99  Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    ' Refresca Tiempos de Operacion para Programacion
    Command3.Enabled = False
    NORPRO$ = Label32
    SOLFA$ = Label26
    If SOLFA$ <> "" Then
       ' derecho de acceso
       If DERACCE%("FRETSTAN", "Re-Frescar Tiempos Standard para Programacion") >= 2 Then
          Dim OPERFFFAB As ADODB.Recordset
          Set OPERFFFAB = New ADODB.Recordset
          SQLX$ = "SELECT * FROM OperFab WHERE IdSubOr = '" & SOLFA$ & "'"
          OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          '
          With OPERFFFAB
             Do While Not .EOF
                CIJK% = XVALO(!Cod_Inte)
                NO% = XVALO(!NUMEOPER)
                CAPROYE# = XVALO(!CANTPROG)
                CAREALI# = XVALO(!CANTREAL)
                CAPENDE# = MAXIVAL(CAPROYE# - CAREALI#, 0)
                '
                ' LEE DATOS DEL MAESTRO DE ARTICULOS
                SQLX$ = "SELECT * FROM SecOper WITH(NOLOCK) "
                SQLX$ = SQLX$ + "WHERE CodIConj=" & CIJK% & " "
                SQLX$ = SQLX$ + " AND NUMEOPER = " & NO%
                   ' ATO% = XVALO(!AltOper)
                Dim SECOTEMP As ADODB.Recordset
                Set SECOTEMP = New ADODB.Recordset
                Set SECOTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
                '
                With SECOTEMP
                    If .EOF Then
                       SECOTEMP.Close
                       Set SECOTEMP = Nothing
                       GoTo 29
                    End If
                    '
                    DOPE$ = SAFETEXT$(!DESCOPER)
                    EQ1$ = SAFETEXT$(!CodMaq0)
                    EQAUX$ = ""
                       If Not IsNull(!CodMaq1) Then EQAUX$ = AJUSTI$(!CodMaq1, 6)
                       If Not IsNull(!CodMaq2) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq2, 6)
                       If Not IsNull(!CodMaq3) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq3, 6)
                    LOTESTA = MAXIVAL(XVALO(!LOTESTAN), 1)
                    CALOTES = 1
                    CAOPS = XVALO(!CantiOps): If CAOPS < 0.01 Then CAOPS = 1
                    CICLOHE = XVALO(!PieCiclo): If CICLOHE < 1 Then CICLOHE = 1
                    MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
                        If CAREALI# > 0.05 Then MINPREP = 0
                    MINFIJO = 60 * XVALO(!HorsFijo) * CALOTES
                    MINVARI = CAPENDE# * XVALO(!MinuStan)
                    MINHERR = CAPENDE# * 60 * XVALO(!HorsMHer) / CICLOHE
                    MINDEMO = (MINFIJO + MINVARI + MINHERR) * XVALO(!PorDemor) / 100
                    MINLIMP = 60 * XVALO(!HorsLimp) * CALOTES
                    MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO + MINLIMP
                    HORAPROD = MINLOTE / 60
                    '
                    BVOF% = BINVAL%(AJUSTI$(SOLFA$, 12))
                    NUOROP% = 0
                    FEAL% = HOY(HOS$)
                    HOAL% = Int(HORANUM(Time$))
                    USUA% = USNBR% Mod 100
                    DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
                    INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
                    PROGBLA$ = Space$(6) + String$(26, 0)
                End With
                SECOTEMP.Close
                Set SECOTEMP = Nothing
                '
                !COD_EXTE = CODEXT$(CIJK%)
                !Descrip = DESCRIT0$(CIJK%)
                !DESCOPER = DOPE$
                !MAQUINA = EQ1$
                !DescMaq = ECOARCH("PADMAQ", EQ1$)
                !EquiAux = EQAUX$
                !CantiOps = CAOPS
                !CantPend = CAPENDE#
                !TiemPrep = MINPREP
                !TiemFijo = MINFIJO
                !TiemVari = MINVARI
                !TiemDemo = MINDEMO
                !TiemMHer = MINHERR
                !TiemLimp = MINLIMP
                !TiemOtro = 0
                !TiemTota = MINLOTE
                !TiemEjec = 0
                !TiemPend = MINLOTE
                .Update
                '
29           .MoveNext
             Loop
          End With
          OPERFFFAB.Close
          Set OPERFFFAB = Nothing
          '
          Call COMUNI("COMPLETADO")
       End If
    End If
    '
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command5_Click()
    '
    CONDI$ = "StatuOrf < 3 "   ' para que no este anulada ni cumplida
    'CONDI$ = CONDI$ + "AND CanApro < (CanProy - 0.005) " ' Se quito para que traiga las reabiertas que estan fabricadas
    CONDI$ = CONDI$ + " and NuOrProd <> '' "
    CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    '
    Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
    Call CommSel("ORDPEN/Indice de Ordenes de Produccion Pendientes", Command5, Label32)
    
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call APLICACIONSKINS(Me)

    Call GRATITFOR(Caption)
    '
End Sub

Private Sub Label32_Change()
    '
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
       NVM$ = TRIM$(ProNroComprobante("CF"))
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    Label51 = "CF-" + NVM$
    HOII% = HOY(HOS$): Label24 = HOS$
    '
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Text1 = "": Text2 = ""
    Label17 = "": Label28 = "": Label30 = "": Label26 = ""
    Label19 = "": Label20 = "": Label21 = ""
    Label44 = "": Label7 = "": Label9 = ""
    Label3 = "": Label4 = "": Label33 = ""
    List1.Clear
    '
    Call ABREORFAB
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         CONDI$ = "NuOrProd = '" & NORFA$ & "'"
         CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
       Label28 = CODEXT$(CIXI%)
       Label30 = DESCRIT0$(CIXI%)
         OF$ = TRIM$(SAFETEXT(VALOBADA("ORDEFABR", "IdSubOr", CONDI$, "NOMESS")))
       Label26 = OF$
       Label17 = FORMATNUM(XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS")), "F13." + CNTDC$(CIXI%))
       '
       CONDI$ = "IdSubOr = '" & OF$ & "' "
       QTR& = CantRegistros("OPERFAB", CONDI$)
       If QTR& > 0 Then
            Call CARGALISTA(List1, "OPERFAB+SECOPER/Cod_Inte = CODICONJ AND OPERFAB.NUMEOPER = SECOPER.NUMEOPER", "OPERFAB.NUMEOPER/I5; OPERFAB.DESCOPER/A32; SECOPER.HOPROFAB/A16", CONDI$ & " ORDER BY OPERFAB.NUMEOPER")
          Else
            Call CARGALISTA(List1, "SECOPER", "NUMEOPER/I5; DESCOPER/A32; HOPROFAB/A16", "CODICONJ = " & CIXI% & " ORDER BY NUMEOPER")
       End If
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub
'
Private Sub Label32_DblClick()
    Call Command5_Click
End Sub

Sub List1_DblClick()
   '
   Call ABREORFAB
   OF$ = TRIM$(Label26)
   If OF$ = "" Then Exit Sub
   CIIIX% = CODINT%(Label28)
   If CIIIX% < 1 Then Exit Sub
   '
16 NOPE% = XVALO(Left$(List1.TEXT, 5))
   CONDI$ = "IdSubOr = '" & OF$ & "' AND NumeOper = " & NOPE%
   FILTX$ = OF$ & ";" & CIIIX% & ";" & NOPE%
   '
   FOGUAR$ = VALOBADA("OPERFAB", "HOBATCHREC", CONDI$)
   TITUHOJ$ = TRIM$(Left$(List1.TEXT, 39))
   If TRIM$(FOGUAR$) <> "" Then
        CODFO$ = RANDSTRING$(-8)
        Call SAVEFILE(LUCOM$ + CODFO$ + ".SQX", FOGUAR$)
        FORIMPR$ = "!" + CODFO$
      Else
        FORIMPR$ = TRIM$(Mid$(List1.TEXT, 40, 16))
   End If
   '
20 ORDHOJA& = 1 + List1.ListIndex
   TOTHOJA& = List1.ListCount
   '
   If FORIMPR$ <> "" Then
      Call STDFORM(FORIMPR$, FILTX$, , TITUHOJ$)
      Openforms = DoEvents
      If STDSCREEN.APRO = "OK" Then
         Call GUARDAFOR(OF$, NOPE%)
      End If
   End If
   Openforms = DoEvents
   '
   If STDSCREEN.PAGIN = "NEXT" Then          ' PAGINA SIGUIENTE
      If ORDHOJA& < List1.ListCount Then
         List1.ListIndex = ORDHOJA&
         GoTo 16
      End If
   End If
   '
   If STDSCREEN.PAGIN = "PREV" Then          ' PAGINA SIGUIENTE
      If ORDHOJA& > 1 Then
         List1.ListIndex = ORDHOJA& - 2
      End If
      GoTo 16
   End If
   '
End Sub

Sub BORRAHOPRO(OF$, NOPE%)
   '
   Call ABREORFAB
   CONDI$ = "IdSubOr = '" & OF$ & "' AND NumeOper = " & NOPE%
   QTR& = CantRegistros("OPERFAB", CONDI$)
   If QTR& > 0 Then
      Call ACTUREGISTRO("OPERFAB", "HoBatchRec", CONDI$, "", RAFE&)
   End If
   '
End Sub
'
Sub GUARDAFOR(OF$, NOPE%)
   '
   Call ABREORFAB
   CONDI$ = "IdSubOr = '" & OF$ & "' AND NumeOper = " & NOPE%
   QTR& = CantRegistros("OPERFAB", CONDI$)
   If QTR& < 1 Then
10    ABUS$ = RANDSTRING$(-10)
      CONDI$ = "IdSubOr = '" & ABUS$ & "'"
      If CantRegistros("OPERFAB", CONDI$) > 0 Then GoTo 10
      Call CreaRegistro("OPERFAB", "IdSubOr", ABUS$, "NOMESS")
      Call ACTUREGISTRO("OPERFAB", "NumeOper", CONDI$, NOPE%, RAFE&, "NOMESS")
        CIXI% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & OF$ & "'", "NOMESS")
      Call ACTUREGISTRO("OPERFAB", "Cod_Inte", CONDI$, CIXI%, RAFE&, "NOMESS")
        CODE$ = CODEXT$(CIXI%)
      Call ACTUREGISTRO("OPERFAB", "Cod_Exte", CONDI$, CODE$, RAFE&, "NOMESS")
        DESCO$ = DESCRIT0$(CIXI%)
      Call ACTUREGISTRO("OPERFAB", "Descrip", CONDI$, DESCO$, RAFE&, "NOMESS")
        DESCO$ = DESCRIOP(CIXI%, NOPE%)
      Call ACTUREGISTRO("OPERFAB", "DescOper", CONDI$, DESCO$, RAFE&, "NOMESS")
      Call ACTUREGISTRO("OPERFAB", "IdSubOr", CONDI$, OF$, RAFE&, "NOMESS")
   End If
   '
   CONDI$ = "IdSubOr = '" & OF$ & "' AND NumeOper = " & NOPE%
   '
   UGRID% = 0
   CODFO$ = RANDSTRING$(-8)
   IGRAF% = 0: TGRAF$ = ""
   N1% = FILEOPEN%(LUCOM$ + CODFO$ + ".SQX", 2, 128)
   ' DATOS DE ENCABEZADO
   TTXX$ = "CONFIG"
   TTXX$ = TTXX$ + Chr$(124) + "FOR_DESC = " + FOR_DESC$
   TTXX$ = TTXX$ + Chr$(124) + "FOR_ORIE = " + FOR_ORIE$
   TTXX$ = TTXX$ + Chr$(124) + "FOR_MGIZ = " + FORMATNUM$(FOR_MGIZ, "F10.1")
   TTXX$ = TTXX$ + Chr$(124) + "FOR_MGSU = " + FORMATNUM$(FOR_MGSU, "F10.1")
   TTXX$ = TTXX$ + Chr$(124) + "AUTO_PRN = " + AUTO_PRN$
   TTXX$ = TTXX$ + Chr$(124) + "PAGI_ANT = " + PAGI_ANT$
   TTXX$ = TTXX$ + Chr$(124) + "PAGI_SIG = " + PAGI_SIG$
   Print #N1%, TTXX$ + Chr$(124)
   '
   ' DATOS DE CADA OBJETO
   For I% = 1 To CANOBJ%
     TTXX$ = "TIPO = " + TIOBJ$(I%)
     TTXX$ = TTXX$ + Chr$(124) + "NAME = " + ONAME$(I%)
     TTXX$ = TTXX$ + Chr$(124) + " H1 = " + FORMATNUM$(H1(I%), "F10.1")
     TTXX$ = TTXX$ + Chr$(124) + " V1 = " + FORMATNUM$(V1(I%), "F10.1")
     TTXX$ = TTXX$ + Chr$(124) + " H2 = " + FORMATNUM$(H2(I%), "F10.1")
     TTXX$ = TTXX$ + Chr$(124) + " V2 = " + FORMATNUM$(V2(I%), "F10.1")
       '
       If TIOBJ$(I%) = "GRAPH" Then
          FNAME$ = TEXTO(I%)
          AGRAF$ = LOADFILE$(FNAME$)
          If AGRAF$ <> "" Then
             IGRAF% = IGRAF% + 1
             NGRAF$ = TRIM$(Str$(IGRAF%))
             While Len(NGRAF$) < 4: NGRAF$ = "0" + NGRAF$: Wend
             TGRAF$ = TGRAF$ + ">>>>>>>>GRAF" + NGRAF$ + AGRAF$ + "GRAF" + NGRAF$ + "<<<<<<<<"
             TEXTO(I%) = "GRAF" + NGRAF$
          End If
       End If
       '
       TX123$ = TEXTO(I%)
       If InStr(TX123$, "COMPAR") > 0 Then
         TX123$ = TEXREPLA$(TX123$)
       End If
       '
     If TRIM$(TX123$) <> "" Then
        TX123$ = TX123$ + "/LOCKED"                        ' AQUI PONER CONDICION DE BLOQUEO
        TTXX$ = TTXX$ + Chr$(124) + " TEXT = " + TX123$
     End If
     '
     If TRIM$(FUENTE$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FONT = " + FUENTE$(I%)
     If TRIM$(TAFUEN$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " SIZE = " + TAFUEN$(I%)
     If TRIM$(FORPA$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FORM = " + FORPA$(I%)
     If MULIN(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " MULIN = 1"
     If TRIM$(ARDOCLI$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " ADOC = " + ARDOCLI$(I%)
     '
     If NEGRI(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " NEGRI = 1"
     If ITALI(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " ITALI = 1"
     If SUBRA(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " SUBRA = 1"
     If INVIS(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " INVIS = 1"
     If ENCAS(I%) = 1 Then TTXX$ = TTXX$ + Chr$(124) + " ENCAS = 1"
     '
     If TRIM$(FOREPLA$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " FOREPLA = " + FOREPLA$(I%)
     If TRIM$(TASEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TASEL = " + TASEL$(I%)
     If TRIM$(CASEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CASEL = " + CASEL$(I%)
     If TRIM$(CONDSEL$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " CONDSEL = " + CONDSEL$(I%)
     If TRIM$(LOCKFI$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " LOCKFI = " + LOCKFI$(I%)
     If TRIM$(VAMINI$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " VAMINI = " + VAMINI$(I%)
     If TRIM$(VAMAXI$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " VAMAXI = " + VAMAXI$(I%)
     If TRIM$(TABVALI$(I%)) <> "" Then TTXX$ = TTXX$ + Chr$(124) + " TABVALI = " + TABVALI$(I%)
     If OBLIGA%(I%) > 0 Then TTXX$ = TTXX$ + Chr$(124) + " OBLIGA = 1"
     '
     If TIOBJ$(I%) = "GRID" Then
        UGRID% = UGRID% + 1
        TTXX$ = TTXX$ + Chr$(124) + " TEGRID = " + TEGUAGRID$(UGRID%)
     End If
     '
     Print #N1%, TTXX$ + Chr$(124)
   Next I%
   Close #N1%
   '
   TTYY$ = LOADFILE$(LUCOM$ + CODFO$ + ".SQX")
      If TGRAF$ <> "" Then TTYY$ = TTYY$ + TGRAF$
   Call ACTUREGISTRO("OPERFAB", "HoBatchRec", CONDI$, TTYY$, RAFE&)
   Call FILEKILL(LUCOM$ + CODFO$ + ".SQX")
   '
End Sub

Function TEGUAGRID$(UGRID%)
   '
   If UGRID% > STDSCREEN.GRID99.Count - 1 Then GoTo 99
   '
   TTYY$ = ""
' POSIBLE ERROR 20
   For ROWI& = 1 To STDSCREEN.GRID99(UGRID%).Rows
      For COLI& = 1 To STDSCREEN.GRID99(UGRID%).Cols
'                       STDSCREEN.GRID99(UGRID%).Row = ROWI& - 1
'                       STDSCREEN.GRID99(UGRID%).Col = COLI& - 1
         VVXX$ = STDSCREEN.GRID99(UGRID%).TextMatrix(ROWI& - 1, COLI& - 1)
         If TRIM$(VVXX$) <> "" Then
'                       Label17.Font = STDSCREEN.GRID99(UGRID%).CellFontName
'                       Label17.FontSize = STDSCREEN.GRID99(UGRID%).CellFontSize
'                       Label17.FontBold = STDSCREEN.GRID99(UGRID%).CellFontBold
'                       Label17.FontItalic = STDSCREEN.GRID99(UGRID%).CellFontItalic
'                       Label17.FontUnderline = STDSCREEN.GRID99(UGRID%).CellFontUnderline
            TTZZ$ = "(" & ROWI& & ";" & COLI& & ";" & " VALUE = " & VVXX$ & ")"
            TTYY$ = TTYY$ + " " + TTZZ$
         End If
      Next COLI&
   Next ROWI&
   '
   TEGUAGRID$ = TTYY$
   '
99 End Function


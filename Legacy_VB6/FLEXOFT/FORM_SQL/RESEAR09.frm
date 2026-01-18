VERSION 5.00
Begin VB.Form RESEAR09 
   BackColor       =   &H00C0D0D0&
   Caption         =   "Generación de SIM por Artículo"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Caption         =   "Frame7"
      Height          =   3010
      Left            =   6600
      TabIndex        =   16
      Top             =   1080
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   3010
      Left            =   7560
      TabIndex        =   15
      Top             =   1080
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame7"
      Height          =   3010
      Left            =   3070
      TabIndex        =   14
      Top             =   1080
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame7"
      Height          =   3010
      Left            =   1410
      TabIndex        =   13
      Top             =   1080
      Width           =   15
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      ItemData        =   "RESEAR09.frx":0000
      Left            =   120
      List            =   "RESEAR09.frx":0007
      MultiSelect     =   1  'Simple
      TabIndex        =   12
      Top             =   1370
      Width           =   9315
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   9285
      TabIndex        =   10
      Top             =   1080
      Width           =   9315
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"RESEAR09.frx":0099
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   9120
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4920
      Left            =   9600
      ScaleHeight     =   4890
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   0
      Width           =   945
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RESEAR09.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Graba e Imprime Solicitud"
         Top             =   3360
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RESEAR09.frx":043C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Kdx"
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
         Left            =   120
         Picture         =   "RESEAR09.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1580
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   690
         Left            =   120
         Picture         =   "RESEAR09.frx":0A50
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command22 
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
         Height          =   690
         Left            =   120
         Picture         =   "RESEAR09.frx":0D5A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   100
         Width           =   650
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Restante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1440
      TabIndex        =   20
      Top             =   4380
      Width           =   1185
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4080
      TabIndex        =   19
      Top             =   4380
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7410
      TabIndex        =   6
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Artículo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   225
      Width           =   1905
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   9315
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   4290
      Width           =   1575
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   6480
      TabIndex        =   2
      Top             =   4350
      Width           =   1185
   End
End
Attribute VB_Name = "RESEAR09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
    Command10.Enabled = False
    CI% = CODINT(Label7)
    If CI% > 0 Then
       FIKARDEX07.Text1 = Label7
    End If
    FIKARDEX07.Show 1
    Command10.Enabled = True

End Sub

Private Sub Command14_Click()
        '
        Command14.Enabled = False
        '
        EligioOF% = 1
        For indLista& = 0 To List1.ListCount - 1
            If List1.Selected(indLista&) = True Then
                EligioOF% = 1: Exit For
            End If
        Next indLista&
        If EligioOF% = 0 Then
            Call COMUNI("Imposible Continuar\ \Debe Elegir la Reserva\de Algunas de las OF en Pantalla")
            GoTo 99
        End If
        '
        Call ABRECOMPRAS
        CONDI$ = "NUME_SOLINT > 0 "
        NUMSOL& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", CONDI$, "NOMESS"))
        HOII% = HOY(HO$)
        CI% = CODINT(Label7)
        'Graba en Encasolint
'        Call CreaRegistro("ENCASOLINT", "NUME_SOLINT", NUMSOL&)
'        REFE$ = "Solicitud de Materiales por Artículo: " & CODEXT$(CI%)
'        CONDI$ = "NUME_SOLINT = " & NUMSOL&
'        Call ACTUREGISTRO("ENCASOLINT", "FEMI_SOLINT", CONDI$, FETEXCOR1$(HOII%), REGAF&)
'        Call ACTUREGISTRO("ENCASOLINT", "DESCRI_SOLINT", CONDI$, REFE$, REGAF&)
        REFE$ = "Solicitud de Materiales por Artículo: " & CODEXT$(CI%)
        Dim RSTENCA As ADODB.Recordset
        Set RSTENCA = New ADODB.Recordset
        SQLX$ = "SELECT * FROM ENCASOLINT WHERE NUME_SOLINT < 1"
        RSTENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        HOII% = HOY(HO$)
        With RSTENCA
            .AddNew
            !NUME_SOLINT = NUMSOL&
            !FEMI_SOLINT = CVDAT(HOII%)
            !DESCRI_SOLINT = REFE$
            .Update
        End With
        RSTENCA.Close
        Set RSTENCA = Nothing
        
        'Grabo en Detasolint
'        Call CreaRegistro("DETASOLINT", "NUME_SOLINT", NUMSOL&)
'        Call ACTUREGISTRO("DETASOLINT", "NUORD_ITEM", CONDI$, 1, REGAF&)
'        Call ACTUREGISTRO("DETASOLINT", "COD_INTE", CONDI$, CI%, REGAF&)
'        Call ACTUREGISTRO("DETASOLINT", "CANT_SOLINT", CONDI$, Label2, REGAF&)
'        Call ACTUREGISTRO("DETASOLINT", "Uni_Mas", CONDI$, UNIMED$(CI%), REGAF&)
'        Call ACTUREGISTRO("DETASOLINT", "FENTRE_SOL", CONDI$, FETEXCOR1$(HOII%), REGAF&)
'        Call ACTUREGISTRO("DETASOLINT", "DESCRITEM", CONDI$, DESCRIT$(CI%), REGAF&)
        Dim RSTDETA As ADODB.Recordset
        Set RSTDETA = New ADODB.Recordset
        SQLX$ = "SELECT * FROM DETASOLINT WHERE Status < 0"
        RSTDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        HOII% = HOY(HO$)
        With RSTDETA
            .AddNew
            !NUME_SOLINT = NUMSOL&
            !NUORD_ITEM = 1
            !COD_INTE = CI%
            '!CANT_SOLINT = XVALO(Label2)
            !CANT_SOLINT = XVALO(Label3)
            !UNI_MAS = Unimed$(CI%)
            !FENTRE_SOL = FETEXCOR1$(HOII%)
            !DESCRITEM = DESCRIT$(CI%)
            .Update
        End With
        RSTDETA.Close
        Set RSTDETA = Nothing
        
        'ACTUALIZO EN EL RESERVA GRABANDO EL NUMERO DE SOLICITUD
        'CONDI$ = "CANTRES > 0.05 AND (NUME_SOLINT IS NULL OR NUME_SOLINT < 1) AND COD_INTE = " & CI%
        'Call ACTUREGISTRO("RESERVA", "NUME_SOLINT", CONDI$, NUMSOL&, REGAF&, "NOMESS")
        CondicionSql$ = "CANTRES > 0.05 AND (NUME_SOLINT IS NULL OR NUME_SOLINT < 1) AND COD_INTE = " & CStr(CI%) & " AND ("
        For indLista& = 0 To List1.ListCount - 1
            If List1.Selected(indLista&) = True Then
                OrFab$ = TRIM$(Mid$(List1.List(indLista&), 1, 12))
                CondicionSql$ = CondicionSql$ & "IdSubOr='" & OrFab$ & "' OR "
            End If
        Next indLista&
        CondicionSql$ = Mid$(CondicionSql$, 1, Len(CondicionSql$) - 4) & ")"
        Call ACTUREGISTRO("RESERVA", "NUME_SOLINT", CondicionSql$, NUMSOL&, REGAF&, "NOMESS")
        '
        Call PRINSOLCOD(CI%, NUMSOL&) ' IMPRIME
        Call PEDCOT07.GENISXOF ' ACTUALIZA EL INDICE
        
        Call COMUNI("Solicitud Generada")
99      Command14.Enabled = True
        Call Command22_Click

End Sub

Private Sub Command22_Click()
   
   Call CIERRARCH("*.*")
   Unload Me
   
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    
10  If ULTREG&("INDICOT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
     Call SELECHO("INDICOT/Indice General de Concurso de Precios")
    Else
     GoTo 99
    End If
    '
    NUDOCU$ = VALACT1$("INDICOT") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    TIDOCUM$ = Left$(VALACT2("INDICOT"), 11)
    pos% = InStr(VALACT2("INDICOT"), "-") ' POSICION DEL PRIMER GUION ANTES DEL VENDED
    PROV$ = Mid$(VALACT2("INDICOT"), pos% + 1, 18)
    If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
           PPXX% = InStr(XX$, "Nro.")
           If PPXX% > 0 Then
             XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
           End If
           If InStr(XX$, NUDOCU$) > 4 Then
             PPX1% = InStr(XX$, PROV$) ' VERIFICA QUE TAMBIEN COINCIDA EL PROVEEDOR
             If PPX1% > 0 Then
               DOCUNU& = CVS(Left$(XX$, 4))
               GoTo 5
             End If
           End If
         End If
    Next UI&
       '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
    GoTo 10
       '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
   GoTo 10

99  Command4.Enabled = True

End Sub

Private Sub Frame4_DragDrop(Source As CONTROL, x As Single, y As Single)

End Sub

Private Sub Label7_Change()
   CI% = CODINT(Label7)
   Label1 = DESCRIT$(CI%) ' DESCRIPCION DEL CODIGO
   'CARGO EL LIST
   List1.Clear
   CONDI$ = "CANTRES > 0.05 AND COD_INTE = " & CI% & " AND (NUME_SOLINT IS NULL OR NUME_SOLINT < 1) ORDER BY COD_INTE"
   Call CARGALISTA(List1, "RESERVA", "IDSUBOR/A69<;CantRes/F14.1" & CNTDC$, CONDI$)
   
   For i& = 0 To List1.ListCount - 1
      TX$ = List1.List(i&)
      CANTI = CANTI + XVALO(Mid$(TX$, 71, 14))
      '
      NORFA$ = TRIM$(Left$(TX$, 12))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
      CODIEXTE$ = VALOBADA("ORDEFABR", "COD_EXTE", CONDI$, "NOMESS")
      DESCRI$ = VALOBADA("ORDEFABR", "DESCRIP", CONDI$, "NOMESS")
      CANTIORF = VALOBADA("ORDEFABR", "CanProy", CONDI$, "NOMESS")
      Call REPLA(TX$, CODIEXTE$, 13, 16)
      Call REPLA(TX$, DESCRI$, 29, 32)
      Call REPLA(TX$, FORMATNUM$(CANTIORF, "F10.1"), 61, 10)
      List1.List(i&) = TX$
   Next i&
   Label2 = FORMATNUM(CANTI, "F14." & CNTDC$)
   Label3 = ""
End Sub

Sub PRINSOLCOD(CI%, NUMSOL&)

   FORIPRE$ = TRIM$(CONTROL$("", "SOLIARTI")) ' FALTA DESIGNAR DONDE SE VA A CAMBIAR
   '
   If FORIPRE$ = "" Then
        Call MENSERR(24, "Falta Definir Formulario de Impresión")
        Exit Sub
   End If
   '
   
   If EXISTE%(FORIPRE$ & "SQR") > 0 Then
       FILTX$ = TRIM$(Str$(NUMSOL&))
       Call STDFORM(FORIPRE$, FILTX$, "PRINT")
       Exit Sub
   End If
   
   TIPODOC$ = "Solicitud de Materiales por Artículo" ' guardar el documento immpreso
   NUMEDOC$ = TRIM$(Str$(NUMSOL&))
   '
   HOII% = HOY(HO$)
   FEC = FECHATEX$(HOII%)
   '
   CODPARAM$(1) = "COD-MPRI": DOCPARAM$(1) = CODEXT$(CI%) '  ARITCULO
   CODPARAM$(2) = "DES-MPRI": DOCPARAM$(2) = DESCRIT$(CI%)
   CODPARAM$(3) = "FECH-EMI": DOCPARAM$(3) = FEC
   CODPARAM$(4) = "CAN-CONJ": DOCPARAM$(4) = Label3 ' TOTAL
   CODPARAM$(5) = "UNI-COMP": DOCPARAM$(5) = Unimed$(CI%)
   CODPARAM$(6) = "REF-MAT2": DOCPARAM$(6) = Label4
   
   CAPARDOC% = 6
   '
   CODTABU1$(1) = "REF-MAT1"
   CODTABU1$(2) = "COD-CONJ"
   CODTABU1$(3) = "DES-MAT"
   CODTABU1$(4) = "CAN-MPRI"
   CODTABU1$(5) = "CANTIDAD"
   CACOLTAB1% = 5
       '
   CAITAB1% = 0

   For J& = 0 To List1.ListCount - 1
        If List1.Selected(J&) = True Then
            CAITAB1% = CAITAB1% + 1
            x$ = List1.List(J&)
            NORFA$ = TRIM$(Mid$(x$, 1, 12))
            CODIEXT$ = TRIM$(Mid$(x$, 13, 16))
            DESCRI$ = TRIM$(Mid$(x$, 29, 32))
            CANTIORF$ = TRIM$(Mid$(x$, 61, 10))
            CANTI$ = TRIM$(Mid$(x$, 71, 14))
            '
            TETABU1$(1, CAITAB1%) = NORFA$
            TETABU1$(2, CAITAB1%) = CODIEXT$
            TETABU1$(3, CAITAB1%) = DESCRI$
            TETABU1$(4, CAITAB1%) = CANTI$
            TETABU1$(5, CAITAB1%) = CANTIORF$
         End If
    Next J&

   Call PRINTDOC("@" + FORIPRE$)

End Sub

Private Sub List1_Click()

    CantidadReservada# = 0
    For indLista& = 0 To List1.ListCount - 1
        If List1.Selected(indLista&) = True Then
            filaLista$ = List1.List(indLista&)
            CantidadReservada# = CantidadReservada# + XVALO(Mid$(filaLista$, 71, 14))
        End If
    Next indLista&
    
    CantidadTotal# = XVALO(Label2)
    Label3.Caption = TRIM$(FORMATNUM$(CantidadReservada#, "F12." & CNTDC$))
    Label4.Caption = TRIM$(FORMATNUM$(CantidadTotal# - CantidadReservada#, "F12." & CNTDC$))

End Sub



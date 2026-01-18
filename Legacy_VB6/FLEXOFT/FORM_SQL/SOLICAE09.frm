VERSION 5.00
Begin VB.Form SOLICAE09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00902090&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cae"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   4320
      ScaleHeight     =   4515
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "SOLICAE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3120
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "SOLICAE09.frx":014A
         Top             =   3400
         Width           =   1515
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "SOLICAE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IDENTIFICAE$
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   TESTCAE09.Show 1

End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   HOII% = HOY(HOS$)
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   '
  ' Timer1.Enabled = True
   
   
99 End Sub

Private Sub Timer1_Timer()
   '
   '
   Call ABRECABAN
   SQLX$ = "SELECT IDCAE,NUMETALO FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NUMECOMP < 0 "
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'FVEN'"
   SQLX$ = SQLX$ + "ORDER BY FEHOREAL"
   Dim CAJATMP As New ADODB.Recordset
   Set CAJATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CAJATMP
      Do While Not .EOF
         IDENTCAE$ = SAFETEXT$(!IDCAE)
         If TRIM$(IDENTCAE$) <> "" Then
            TESTCAE09.Text4 = XVALO(!NUMETALO)
            TESTCAE09.Text5 = IDENTCAE$
            TESTCAE09.Command5_Click
         End If
      .MoveNext
      Loop
   End With
   CAJATMP.Close
   Set CAJATMP = Nothing
   '
   '
End Sub
Sub SOLICITA_CAE_ID(IDENTCAE$)

    FLAG% = 0
    cont% = 0
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    '
    If EXISTE%(LUDAT$ + "consulta_cae_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_original.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    
    ARCHIV$ = CARPETOK$ + "CONSULTA_CAE_" + IDENTCAE$ + ".TXT"
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(ARCHIV$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
    '
    'OBTENGO DATOS
    CONDI$ = "IDCAE = '" & IDENTCAE$ & "'"
    NC% = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI$, "NOMESS"))
    NUFACTU$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
    NUMEROFAC = XVALO(VALOBADA("CAJABANC", "NUMECOMP", CONDI$, "NOMESS"))
    FECHEM$ = FECHATEX$(CVINT(VALOBADA("CAJABANC", "FECHEMISI", CONDI$, "NOMESS")))
    '
    CONDI1$ = CONDI$ + " AND TipoMovim = 'IBRU' AND OpciMovim = 'IB-B' AND VaMovim = 'PERC'"
    PERCEPIB# = XVALO(VALOBADA("CAJABANC", "ImpoDeb", CONDI1$, "NOMESS"))
    '
    CONDI1$ = CONDI$ + " AND TipoMovim = 'IVA' AND OpciMovim = 'DeFi' AND VaMovim = 'PER-IVA'"
    PERCEPIVA# = XVALO(VALOBADA("CAJABANC", "ImpoDeb", CONDI1$, "NOMESS"))
    '
    CONDI1$ = CONDI$ + " AND TIPOMOVIM = 'VENTA'"
    IMPNONETO# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI1$, "NOMESS"))
    '
    CONDI1$ = CONDI$ + " AND TIPOMOVIM = 'FVEN'"
    IMPOTOT# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CONDI1$, "NOMESS"))
    '
    CONDI1$ = CONDI$ + "AND TIPOMOVIM = 'IVA'"
    IMPIVA# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI1$, "NOMESS"))
    '
    CONDI1$ = CONDI$ + "AND TIPOMOVIM = 'FVEN'"
    PREFI$ = VALOBADA("CAJABANC", "OPCIMOVIM", CONDI1$, "NOMESS")
    LETRA$ = VALOBADA("CAJABANC", "VARIMOVIM", CONDI1$, "NOMESS")
    PVTA$ = Str(XVALO(Mid$(NUFACTU$, 5, 3)))
    '
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           'Case 4: DATO$ = TRIM$(FORFACPRINT.Label18)          ' Cantidad de registros del detalle de la factura
           Case 4: DATO$ = "1"            ' corregido porque debe ser siempre 1
           Case 5: DATO$ = "1"            ' Tipo de Presentacion dada (0 servicios , 1 Bienes de Capital)
           Case 9: DATO$ = TIPO_DOC$(NC%) ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 10: DATO$ = NRO_DOC$(NC%)      ' Numero Dni o Cuit
           Case 11: DATO$ = TIPO_CBTE$(PREFI$, LETRA$)       ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 12: DATO$ = PVTA$ ' Punto de Venta
           Case 13: DATO$ = TRIM$(Str(NUMEROFAC))         ' Numero de comprobante inicial)
           Case 14: DATO$ = TRIM$(Str(NUMEROFAC))               ' Numero de comprobante hasta
           Case 15: DATO$ = TRIM$(FORMATNUM$(IMPOTOT#, "F12.2"))         ' Importe Total
           Case 16: DATO$ = TRIM$(IMP_TOT_CONC$(PERCEPIVA#, PERCEPIB#, IMPNONETO#))          ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
           Case 17: DATO$ = TRIM$(FORMATNUM$(IMPNONETO#, "F12.2"))               ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 18: DATO$ = TRIM$(FORMATNUM$(IMPIVA#, "F12.2"))           ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 19: DATO$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 20: DATO$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 21, 22, 23, 24:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FECHEM$))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    DATO$ = ANIO$ + MES$ + DIA$
          End Select
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Close #NARCHI1%
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE" + IDENTCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    If EXISTE%(CARPETOK$ + "CAE" + IDENTCAE$ + ".TXT") = 0 Then
      If EXISTE%(CARPETOK$ + "wsfe_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfe_cliente.exe AUTO " & IDENTCAE$, 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
         GoTo 99
      End If
    End If
    '
99 End Sub

Sub BUSCACAE(IDENTCAE$)
   '
   CARPETOK$ = "C:\FLEXOFT\FACELEC\"
   If EXISTE%(CARPETOK$ + "CAE_" + IDENTCAE$) > 0 Then
     IDENTIFICAE$ = IDENTCAE$
       'VALIDAR SE TIENE EL CAE
        Timer2.Enabled = True
        'SI: RUTINA DE GRABACION DE FACTURA
        'NO: IMPRIMIR UN COMPROBANTE CON LOS DATOS DE LA FACTURA Y EL MOTIVO RECHAZO
    Else
      'CASO CONTRARIO (NO EXISTE)
        'RUTINA QUE CARGA EL DETALLE DE LA FACTURA Y PASAR LA INFORMACION AL ARCHVIVO
        'EN EL CONSULTA_CAE_NNNN.TXT
        Call SOLICITA_CAE_ID(IDENTCAE$)
        'EJECUTAR EL .NET Y PASAR LA LEYENDA AUTO MAS EL NUMERO DEL ID
        'ESPERAR LA DEVOLUCION DEL CAE MEDIANTE UN TIMER LEECAE
        ' SI NO VUELVE EL CAE QUE CONTINUE CON EL PROXIMO
        
      
   End If
   
End Sub

Function TIPO_DOC$(NC%)
    ' TIPO DE DOCUMENTO
    'NC% = XVALO(FORFACPRINT.Text1)
    POSDEIVA% = PIVACLI%(NC%)
    Select Case POSDEIVA%
        Case 0: VALOR$ = 96
        Case 1: VALOR$ = 80
        Case 2: VALOR$ = 80
        Case 3: VALOR$ = 80
        Case 4: VALOR$ = 80
        Case 5: VALOR$ = 80
        Case 6: VALOR$ = 80
        Case 7: VALOR$ = 80
    End Select
    TIPO_DOC$ = VALOR$
End Function


Function NRO_DOC$(NC%)
    NRO_DOC$ = ""
    POXIVA% = PIVACLI%(NC%)
    '
    NUMERO_CUIT_DOC1$ = CUITCLI$(NC%)
    NUMERO_CUIT_DOC$ = REPLACAR(NUMERO_CUIT_DOC1$, "-", "")
    '
    If POXIVA% < 1 And TRIM$(NUMERO_CUIT_DOC$) = "" Then NUMERO_CUIT_DOC$ = "0"
    '
    If Not IsNumeric(NUMERO_CUIT_DOC$) Then
      Call MENSERR(24, "Documento Nro.: " & NUMERO_CUIT_DOC$ & " No Válido\El Dato debe ser Numérico\Modifique y Vuelva a Cargar la Factura ")
      Call FINAL("")
    End If
    '
    NRO_DOC$ = NUMERO_CUIT_DOC$
    '
End Function


Function IMP_TOT_CONC$(PERCEPIVA#, PERCEPIB#, IMPNONETO#)

   IMP_TOT_CONC$ = "" ' SUMA DE IVA NO GRAVADO + PERCEPCION DE IVA + PERCEPCION DE INGRESOS BRUTOS
   IMPTOTACONC$ = FORMATNUM$(XVALO(IMPNONETO#) + XVALO(PERCEPIVA#) + XVALO(PERCEPIB#), "F12.2")
   If TRIM$(IMPTOTACONC$) = "" Then IMPTOTACONC$ = "0"
   IMP_TOT_CONC$ = TRIM$(IMPTOTACONC$)
   
End Function

Function IMPTO_LIQ$()
  IMPTO_LIQ$ = ""
  IMPTOLIQUID$ = FORMATNUM$(XVALO(FORFACPRINT.Total(3)) + XVALO(FORFACPRINT.Total(4)) + XVALO(FORFACPRINT.Total(5)), "F12.2")
  IMPTO_LIQ$ = TRIM$(IMPTOLIQUID$)
End Function

Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPO_CBTE$(PREFI$, LETRA$)
    
    TIPOCOMP$ = UCase$(PREFI$ & LETRA$)
    Select Case TIPOCOMP$
        Case "FCA": VALOR$ = "1" 'FACTURA A
        Case "NDA": VALOR$ = "2" 'N DEBITO A
        Case "NCA": VALOR$ = "3" 'N CREDITO 1
        Case "FCB": VALOR$ = "6" 'FACTURA B
        Case "NDB": VALOR$ = "7" 'N DEBITO B
        Case "NCB": VALOR$ = "8" 'FACTURA B
    End Select
    '
    TIPO_CBTE$ = VALOR$
    '
End Function


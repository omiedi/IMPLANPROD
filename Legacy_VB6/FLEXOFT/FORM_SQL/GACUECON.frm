VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GACUECON 
   BackColor       =   &H00C7D9CC&
   Caption         =   "Gastos Mensuales Detallados"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   9210
      Left            =   11760
      ScaleHeight     =   9180
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command23 
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
         Left            =   105
         Picture         =   "GACUECON.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir "
         Top             =   7455
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   690
         Left            =   105
         Picture         =   "GACUECON.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   6615
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "GACUECON.frx":19D0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir de  la Aplicación"
         Top             =   105
         Width           =   650
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8190
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   14446
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   14737632
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "Label2"
      Height          =   225
      Left            =   10815
      TabIndex        =   9
      Top             =   105
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "302.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9660
      TabIndex        =   8
      Top             =   8775
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "Gastos Mensuales Detallados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   10725
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "GACUECON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Private Sub Command2_Click()
   
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To MSF.Rows - 1
      For i& = 0 To MSF.Cols - 1
       If InStr(UCase$(MSF.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        MSF.Row = U&: MSF.COL = i&
        If MSF.TopRow > MSF.Row Then MSF.TopRow = MSF.Row
        On Error Resume Next
        VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If MSF.TopRow < MSF.Row - VISIROWS Then
              MSF.TopRow = MSF.Row - VISIROWS
            End If
         End If
        MSF.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   If XVALO(Label2.Caption) > 0 Then
       'POR CUENTA CONTABLE
       Call FINAL("?GASCUECON2(" & TRIM$(Label2) & ")")
   Else
       'REPORTE CON TODOS
       Call FINAL("?GASCUECON1")
   End If
   Command23.Enabled = True
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   '
   'CARGO LA GRILLA
   'FILTRO POR CUENTA CONTABLE
   Screen.MousePointer = 1
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   Call SELECHO("CUECON")
   
   Screen.MousePointer = 11
   CTA% = XVALO(VALACT1$("CUECON"))
   Label2 = CTA%
   'SI NO ELIGE MUESTRA TODAS

   CONDI2$ = "CUECONTAI > 0 "
   'SI SELECCIONO CTA FILTRA Y AGREGA ENCABEZADO A FRM
   If CTA% > 0 Then
       CONDI2$ = CONDI2$ + " AND CueContaI = " & CTA%
       DENOCUEX$ = TRIM$(ECOARCH$("ECUECON", CODCUE$(CTA%)))
       Label1.Caption = Label1.Caption + "  (" & DENOCUEX$ & ")"
   End If
   '
   CONDI2$ = CONDI2$ + " ORDER BY CUECONTAI"
   Call CARGALISTA(MSF, "GASCUECON", "CODEXCUE/A16+ECUECON/A16;Codicom_lar/A18;NUPROVE/I5+PROVED1/A32;FECHEMISI/D8;IMPORTE/F12.2", CONDI2$, "NOMESS")
   MSF.Cols = 7
   NUCOLU% = 1: TEPRUEBA = "ABCDEFGHIJ": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "C.Cont."
   NUCOLU% = 2: TEPRUEBA = "ABCDEFGHIJKLLOPQRSTUV": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLOPQRSTUV": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
   NUCOLU% = 4: TEPRUEBA = "12345678": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Cta."
   NUCOLU% = 5: TEPRUEBA = "ABCDEFGHIJKLLLMNÑOPQRSTUVABCDE": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "R.Social"
   NUCOLU% = 6: TEPRUEBA = "10/05/20123": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha"
   NUCOLU% = 7: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Importe"
   '
   'TOTALIZA LA COLUMNA IMPORTE
   SUMA# = 0
   For i& = 0 To MSF.Rows - 1
      SUMA# = SUMA# + XVALO(MSF.TextMatrix(i&, 6))
   Next i&
   Label18 = FORMATNUM$(SUMA#, "F13.2")
   Screen.MousePointer = 1
   '
End Sub

Private Sub MSF_DblClick()
   
   i& = MSF.MouseRow
   J& = MSF.MouseCol
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   '
   MSF.Row = i&
   MSF.COL = J&
   '
   NUDOCU1$ = TRIM$(MSF.TextMatrix(MSF.Row, 2))
   FE% = FECHANUM(MSF.TextMatrix(MSF.Row, 5))
   NUPROV% = XVALO(MSF.TextMatrix(MSF.Row, 3))
   If Left$(UCase$(NUDOCU1$), 3) <> "OP-" Then
      Call MUEFACO1(NUDOCU1$, FE%, NUPROV%)
      Exit Sub
   End If
   '
   TIDOCUM$ = "Orden de Pago"
   NPX& = XVALO(Mid$(NUDOCU1$, 11, 8))
   NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
   Screen.MousePointer = 11
   '
   DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
   If DOCUNU& < 1 Then
         Screen.MousePointer = 1
         Call COMUNI("Imposible Mostrar Documento")
         Exit Sub
   End If
   Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
   '
End Sub




Sub COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF As MSFlexGrid, COLOR As Variant, Optional REGISTRO)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.Row
    End If
    
    .Row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .COL = iCol%
          .CellBackColor = COLOR
    Next iCol%
   End With
End Sub


Private Sub MSF_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF.Row

          H& = 0: If MSF.COL < MSF.Cols Then H& = MSF.COL + 1
10        For U& = J& To MSF.Rows - 1
             HI& = H&: If U& > J& Then HI& = 0
             For i& = HI& To MSF.Cols - 1
              If InStr(UCase$(MSF.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF.Row = U&: MSF.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF.TopRow < MSF.Row - VISIROWS Then
                      MSF.TopRow = MSF.Row - VISIROWS
                    End If
                 End If
                MSF.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub
Sub MUEFACO1(TIDOCUM$, FE%, NUPROV%)
   '
   TIPODOCU$ = TRIM$(Left$(TIDOCUM$, 4)) 'FC-A
   NUDOCOR$ = TRIM$(TIDOCUM$) 'NUMERO DE DOC COMP
   NP1& = NUPROV%
   FECDOC1 = FETEXCOR1$(FE%) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   'jandy sql
   Dim SHOFACOTMP As ADODB.Recordset
   Set SHOFACOTMP = New ADODB.Recordset
   SHOFACOTMP.CursorLocation = adUseClient
   SQLZ$ = " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLZ$ = SQLZ$ + " AND NuProve =  " & NP1& & " "
   SQLX$ = " SELECT * FROM CAJABANC " + SQLZ$
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
25 On Error Resume Next
   SHOFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = NUDOCOR$ + " del " + FECHATEX$(FE%) ' TIPODOCU$ + " " + !Codicom_lar + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '\\\OT-08-988-29/10/08-FL
     SHOWFACO.Label19 = ""
     '\\\OT-08-988-FIN-FL
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CUECONTAI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DESCRIMOV
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.List1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !Idebecomp) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !Idebecomp)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !OBSERGEN
          '\\\OT-08-988-29/10/08-FL
          'SE CAPTURA UN POSIBLE ERROR PRODUCTO QUE EL CAMPO IDENALTA CONTENGA INFORMACION INVALIDA
          On Error Resume Next
          USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
          On Error GoTo 0
          SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
          '\\\OT-08-988-FIN-FL
        End If
        '
     .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ASIGLOAHP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASIGNACION AUTOMATICA DE LOTES "
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   19
      Top             =   3120
      Width           =   9720
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "R.Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   6360
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   480
      ScaleHeight     =   75
      ScaleWidth      =   11670
      TabIndex        =   17
      Top             =   2760
      Width           =   11730
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remito:                                                           Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13545
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   7720
         TabIndex        =   6
         Top             =   600
         Width           =   15
      End
      Begin VB.CommandButton Command1 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   12720
         Picture         =   "ASIGLOAHP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Aprueba Lotes Asignados"
         Top             =   1800
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Auto FIFO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   12720
         TabIndex        =   15
         ToolTipText     =   "Lotes Recomendados por Antiguedad"
         Top             =   240
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Frame Frame8 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   11880
         TabIndex        =   13
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   2330
         TabIndex        =   8
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6525
         TabIndex        =   7
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   10960
         TabIndex        =   4
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   9360
         TabIndex        =   5
         Top             =   600
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         ItemData        =   "ASIGLOAHP.frx":030A
         Left            =   200
         List            =   "ASIGLOAHP.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   12470
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   8080
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"ASIGLOAHP.frx":030E
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   12
         Top             =   600
         Width           =   12465
      End
      Begin VB.Label lblfactura 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label LBLCLIENTE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2520
      Left            =   10320
      ScaleHeight     =   2460
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command4 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         Picture         =   "ASIGLOAHP.frx":03BA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba Lotes Asignados"
         Top             =   240
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "ASIGLOAHP.frx":06C4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1320
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ASIGLOAHP.frx":25E6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ASIGLOAHP.frx":281A
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LBLCODIGO 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12600
      TabIndex        =   26
      Top             =   2680
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "ASIGLOAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CALCSAL#()
Public EJECUTAR%
Public RECEPCION%



 Sub Command6_Click()
   '

   '
99 Command6.Enabled = True
   '
End Sub

Private Sub Command1_Click()
     Command1.Enabled = False
     Screen.MousePointer = 11
     '
     If List2.ListCount < 1 Then GoTo 99
     DEPO% = XVALO(Control$("", "DEPOCEN"))
     '
'     For I& = 1 To List2.ListCount ' VALIDACIONES
'          CIXI% = CODINT%(TRIM$(Left$(List2.List(I& - 1), 20)))
'          If TRIM$(Mid$(List2.List(I& - 1), 74, 16)) = "" Then
'             Call COMUNI("Falta Lote Para Código: " & CODEXT$(CIXI%))
'             Label7 = "FALTAN_LOTES"
''             GoTo 99 'comentado para que pueda continuar sin ingresar todos los lotes
'          End If
'     Next I&
     
     '
     For i& = 1 To List2.ListCount ' GRABA
          Call TRACE(20, i&, List2.ListCount)
          CIXI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 20)))
          ORD% = XVALO(TRIM$(Mid$(List2.List(i& - 1), 120, 4)))
          CANTASI = XVALO(TRIM$(Mid$(List2.List(i& - 1), 62, 11)))
          LOTE$ = TRIM$(Mid$(List2.List(i& - 1), 74, 16))
        
          CONDI$ = "DOPRILAR = '" & TRIM$(lblfactura) & "' AND CodiMovi = 'RT' AND Cod_Inte = " & CIXI% & " AND NUORITEM = " & ORD% & " AND NUME_CLIE = " & XVALO(LBLCLIENTE)
          If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
             GoTo 20
             
          End If
          Dim MOVISE As ADODB.Recordset
          Set MOVISE = New ADODB.Recordset
          Set MOVISE = FLEXCONN.Execute(FILTSQL$("SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " + CONDI$))
          'GUARDA TODOS LOS DATOS DE LA FILA
          CACAM& = MOVISE.Fields.Count
          ReDim VARS(CACAM&)
          For U& = 1 To CACAM&
             VARS(U&) = MOVISE.Fields(U& - 1).Value
          Next U&
          '
          'PONE EN 0 CANTIDAD SALIDA Y CANTIDAD INGRESADA PARA REALIZAR EL ACSALOTE
          LOTEVIE$ = TRIM$(VALOBADA("MOVISTO", "IdenLote", CONDI$))
          Call ACTUREGISTRO("Movisto", "CantIngre", CONDI$, 0, REGAF&, "NOMESS")
          Call ACTUREGISTRO("Movisto", "CantSalid", CONDI$, 0, REGAF&, "NOMESS")
          If TRIM$(LOTEVIE$) <> "" Then
             Call ACSALOTE(CIXI%, LOTEVIE$, "NOMESS")
          End If
          'BORRA LOS REGISTROS
          Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&, "NOMESS")
          '
20         Dim Movisto As ADODB.Recordset
          Set Movisto = New ADODB.Recordset
          strSQL = "SELECT * FROM MOVISTO WHERE " + CONDI$
'25        On Error Resume Next
          Movisto.Open strSQL, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'             On Error GoTo 0
'             Call CapturaErrorOpen
'             GoTo 25
'          End If
'          On Error GoTo 0
'          '
          SUMASI = 0
          If CANTASI >= FRACMIN(CIXI%) Then
                With Movisto
                   .AddNew
                   'GENERA NUEVO REGISTRO CON LA INFORMACION GUARDADA ANTERIORMENTE
                   For U& = 1 To CACAM&
                      Movisto.Fields(U& - 1).Value = VARS(U&)
                   Next U&
                   !CANTSALID = CANTASI: SUMASI = SUMASI + CANTASI
                   !idenlote = LOTE$
                   .Update
                   Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
                End With
             End If
          '
          RESTO = XVALO(Label24) - SUMASI
          If RESTO >= FRACMIN(CIXI%) Then
            With Movisto
               .AddNew
               For U& = 1 To CACAM&
                  Movisto.Fields(U& - 1).Value = VARS(U&)
               Next U&
               !CANTSALID = RESTO
               !idenlote = ""
               .Update
            End With
          End If
          '
          Movisto.Close: Set Movisto = Nothing
          On Error Resume Next
          MOVISE.Close: Set MOVISE = Nothing
          On Error GoTo 0
          'Call Me.lblfactura_Change
    '
60    Next i&
      '
'    Call LIMPIAR
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
99  Screen.MousePointer = 1
    Command1.Enabled = True
    Unload Me

End Sub

Private Sub Form_Activate()
   If EJECUTAR% = 1 And List2.ListCount > 0 And lblfactura.Caption <> "" Then
      Call Command15_Click
   End If
End Sub

'Private Sub Command4_Click()
'    Command4.Enabled = False
'    DEPO% = DEPOEXPE%: If DEPO% = 0 Then DEPO% = 1
'
'    For j& = 1 To ASIGLOAHP.MSF.Rows - 1
'            DOCU$ = XVALO(MSF.TextMatrix(j&, 1))
'            CI% = CODINT(MSF.TextMatrix(j&, 2))
'            DESCRI$ = MSF.TextMatrix(j&, 3)
'            CANTI = XVALO(MSF.TextMatrix(j&, 4))
'            LOTE$ = XVALO(MSF.TextMatrix(j&, 5))
'            NUROIT% = XVALO(MSF.TextMatrix(j&, 6))
'            NPROV = XVALO(MSF.TextMatrix(j&, 7))
'      CANECE = CANTI ' CANTIDAD FACTURADA
'      If LOTE$ <> "" Then
'         CANLOTE = SaldoLote(CI%, LOTE$): If CANLOTE < 0 Then CANLOTE = 0  'VEO EL SALDO DEL LOTE
'         If CANLOTE > 0 Then             'SI TIENE SALDO
'           If CANTI > CANLOTE Then CANTI = CANLOTE 'SI LA CANTIDAD ORIGINAL ES MAYOR AL SALDO DEL LOTE ASIGNO LA CANTIDAD DEL LOTE
'           'GENERA MOVIMIENTO DE  SALIDA CON LOTE POR EL TOTAL DE LA CANTIDAD SI ALCANZA O POR EL SALDO DEL LOTE
'           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, , TEXLIB$, , , PRECONIVA#) '
'           If Abs(CANECE) - Abs(CANLOTE) > 0 Then GoTo 60
'      End If
'    Else
'
'
'      Call ARMALISLOTES(CI%, DEPO%, , , 1, NPROV)
'      FIN& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
'
'      CANECE = CANTIN ' CANTIDAD FACTURADA
'      For U& = 1 To FIN&
'        XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
'        CALOTE = XVALO(Mid$(XX$, 35, 10)) ' LEO LA POSICION CANTIDAD
'        LOTSE$ = TRIM$(Mid$(XX$, 1, 16))
'        If CALOTE < FRACMIN(CI%) Then GoTo 80
'        If CALOTE < CANECE Then
'              DIFE = CANECE - CALOTE
'              CANECE = CALOTE
'          End If
'         'ACA MOVISTOK(CANECE)
'         CANECE = CANECE - CALOTE
'         If CANECE <= 0 Then Exit For
'80   Next U&
'
'     If CANECE > 0 Then MOVISTOK (CANECE)
'
'
'
'
'    Next I&
'
'99  Command4.Enabled = True
'
'End Sub
'
Private Sub Form_Load()
  Call APLICACIONSKINS(Me)
  EJECUTAR% = 0
End Sub
 Sub Command15_Click()
    Command15.Enabled = False
    Screen.MousePointer = 11
    'Call LIMPIAR
    DEPO% = 1
    MAXIMOORD% = List2.ListCount - 1
    Do While i& <= List2.ListCount - 1
      i& = i& + 1
      TX$ = List2.List((i& - 1))
      CI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 20))) 'CODIGO INTERNO
      A = CODEXT$(CI%)
      CANECE = XVALO(Mid$(List2.List(i& - 1), 62, 11)) ' CANTIDAD
      CANTIORIG = XVALO(Mid$(List2.List(i& - 1), 62, 11)) ' CANTIDAD
      
      NPROV = XVALO(Mid$(List2.List(i& - 1), 104, 8)) ' PROVEEDOR
      NUMEORD% = XVALO(Mid$(List2.List(i& - 1), 120, 4))
      Call ARMALISLOTES(CI%, DEPO%, , , 1, NPROV)
      FIN& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
      '
      ACUMULACALOTE = 0
      '
      'RECORRPO EL LILOTSE
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
        CALOTE = XVALO(Mid$(XX$, 35, 10)) ' LEO LA POSICION CANTIDAD
        CANECE = XVALO(Mid$(List2.List(i& - 1), 62, 11)) '
        LOTSE$ = TRIM$(Mid$(XX$, 1, 16))
'        RESERVADO1 = CANTRESER(CI%, LOTSE$)
'        CALOTE = CALOTE - RESERVADO1
        
        If CALOTE < FRACMIN(CI%) Then GoTo 80 'SI NO TIENE CANTIDAD PASA AL SIGUIENTE
        'RECORRE EL LIST VIENDO SI USO CANTIDAD DE ESTE LOTE Y EN ESE CASO RECALCULA LA CANTIDAD DISPONIBLE DEL LOTE
        'ESTA SITUACION NO DEBERIA PASAR SALVO SI HAY ITEMS REPETIDOS EN EL LIST.
        If LOTSE$ <> "" Then
          For KII% = 1 To List2.ListCount
            If KII% - 1 <> i& - 1 Then
              If TRIM$(Mid$(List2.List(KII% - 1), 62, 16)) = LOTSE$ Then
                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 62, 11))
              End If
            End If
          Next KII%
        End If
        '
        'AQUI AGREGA EL LOTE CORRESPONDIENTE Y SI NO CUBRE LA NECESIDAD GENERA UN NUEVO REGISTRO POR EL RESTO SIN LOTE
        TY$ = ""
        
        If CALOTE > 0 Then
           
           If CALOTE < CANECE Then
              DIFE = CANECE - CALOTE
              CANECE = CALOTE
              TY$ = TX$
              Call REPLA(TY$, FORMATNUM(DIFE, "F10.1"), 62, 11)
              
           End If
           ACUMULACALOTE = ACUMULACALOTE + CANECE
           A = CODEXT$(CI%)
           Call REPLA(TX$, FORMATNUM$(CANECE, "F11.1"), 62, 11)
           Call REPLA(TX$, TRIM$(LOTSE$), 74, 16)
           Call REPLA(TX$, FORMATNUM$(NUMEORD%, "F4"), 120, 4)
'           MAXIMOORD% = MAXIMOORD% + 1
'           NUMEORD% = MAXIMOORD%
           List2.List(i& - 1) = TX$
           '
           'GENERA UN NUEVO REGISTRO POR EL RESTO SIN LOTE
           If TY$ <> "" Then
                    MAXIMOORD% = MAXIMOORD% + 1
                    NUMEORD% = MAXIMOORD%

                    Call REPLA(TY$, Space$(16), 74, 16)
                    Call REPLA(TY$, FORMATNUM$(NUMEORD%, "F4"), 120, 4)
                    'List2.AddItem Left$(TY$, 74), I& ' que agregue cantidad pendiente sin lote
                    List2.AddItem TY$
                    i& = i& + 1
           End If
        End If

       If CANTIORIG <= ACUMULACALOTE Then Exit For
          
80    Next U&
      '
    Loop
    Screen.MousePointer = 1
    Command15.Enabled = True
    Call Command1_Click
End Sub

'Private Sub Label8_Change()
'    REMITO$ = TRIM$(Label8)
'
'    If REMITO$ = "" Then Exit Sub
'    MSF.Rows = 0
'
'    CAMPOS$ = "Remito/A20;Código/A20;Descripción/A32;Cant./F9.O;Lote/A18;Proveedor/F7;NUORD/F0"
'    Call CARGA_ENCABEZADO(ASIGLOAHP.MSF, CAMPOS$, ASIGLOAHP, , , ANTOT)
'    '
'    j& = 0
'    SQLX$ = " Select DOPRILAR,DOSECLAR,COD_INTE, DESCRIPCION,CANTSALID,IDENLOTE from MOVISTO with(nolock)"
'    SQLX$ = SQLX$ + " where DOPRILAR ='" & REMITO$ & "'"
'    SQLX$ = SQLX$ + " AND CODIMOVI = 'RT'"
'    SQLX$ = SQLX$ + " AND CANTSALID > 0"
'    SQLX$ = SQLX$ + " AND (IDENLOTE = '' OR IDENLOTE IS NULL)"
'    SQLX$ = SQLX$ + " ORDER BY COD_INTE"
'
'    Set RS1 = READSET(SQLX$)
'    With RS1
'        Do While Not .EOF
'            j& = j& + 1
'            LOTE$ = TRIM$(SAFETEXT$(!IDENLOTE))
'            CIXI% = XVALO(!COD_INTE)
'            ORD% = XVALO(!NuorItem)
'            CANTIDADLEIDA = XVALO(!CANTSALID)
'            REMIX$ = SAFETEXT$(!DOPRILAR)
'            '
'            If CIXI% <> CIXI_ANT Then
'              Call ARMALISLOTES(CI%, DEPO%, , , 1, NPROV)
'              FIN& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
'              CANTREMTIDA = CANTIDADLEIDA
'            End If
'
'            CANECE = CANTASI ' CANTIDAD FACTURADA
'            For U& = 1 To FIN&
'                XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
'                CALOTE = XVALO(Mid$(XX$, 35, 10)) ' LEO LA POSICION CANTIDAD
'                LOTSE$ = TRIM$(Mid$(XX$, 1, 16))
'                If CALOTE < FRACMIN(CIXI%) Then GoTo 90
'
'                If CANTREMTIDA > CALOTE Then
'                   CANTREMTIDA = CALOTE
'                End If
''                If VUELTA% < 1 Then
''                   CONDI$ = "DOPRILAR = '" & TRIM$(REMIX$) & "' AND CodiMovi = 'RT' AND Cod_Inte = " & CIXI% & " AND NUORITEM = " & ORD%
''                   Call ACTUREGISTRO("Movisto", "CantSalid", CONDI$, CANECE, REGAF&, "NOMESS")
''                   Call ACTUREGISTRO("Movisto", "IDENLOTE", CONDI$, LOTSE$, REGAF&, "NOMESS")
''                   VUELTA% = 1
''                Else
'                   'LEO GENERO NUEVOS REGISTROS SIEMPRE
'                   'ACTUALIZO LA CANTIDAD DE REGIRSTROS CON CANTREMITIDA
''                End If
'                CANTIDADLEIDA = CANTIDADLEIDA - CANTREMTIDA
'                If CANTIDADLEIDA > 0 Then
'                  GoTo 80
'
'                If CANTASI <= 0 Then
'                   .MoveNext
'                   If CIXI_ANT% = XVALO(!COD_INTE) Then
'                     CIXI% = XVALO(!COD_INTE)
'                     ORD% = XVALO(!NuorItem)
'                     CANTASI = XVALO(!CANTSALID)
'                     'ACA DEBERIA IR AL NEXT U& PERO DEBERIA CONSIDERARSE SI DE LA ULTIMA ASIGNACION QUEDO ALGUN SALDO DE ESE LOTE QUE SE  LE PUEDA ASIGNAR.
'                     'TAMBIEN LO DEBERIA REALIZAR POR ACTUREGISTRO
'                     GoTo 80
'                   End If
'                End If
'
'                CIXI_ANT% = CIXI%
''80              .MoveNext
''                If CIXI_ANT% = XVALO(!COD_INTE) Then
''
''                If CANECE <= 0 Then Exit For
''                IF CANECE > 0 AND
'            Next U&
'90         .MoveNext
'        Loop
'    End With
'
'       If CANECE > 0 Then MOVISTOK (CANECE)
'
'
'
'            CONDI$ = "DOPRILAR = '" & TRIM$(REMIX$) & "' AND CodiMovi = 'RT' AND Cod_Inte = " & CIXI% & " AND NUORITEM = " & ORD%
'            Dim MOVISE As ADODB.Recordset
'            Set MOVISE = New ADODB.Recordset
'            Set MOVISE = FLEXCONN.Execute(FILTSQL$("SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " + CONDI$))
'            'GUARDA TODOS LOS DATOS DE LA FILA
'            CACAM& = MOVISE.Fields.Count
'            ReDim VARS(CACAM&)
'            For U& = 1 To CACAM&
'              VARS(U&) = MOVISE.Fields(U& - 1).Value
'            Next U&
'            '
'            'PONE EN 0 CANTIDAD SALIDA Y CANTIDAD INGRESADA PARA REALIZAR EL ACSALOTE
'          LOTEVIE$ = TRIM$(VALOBADA("MOVISTO", "IdenLote", CONDI$))
'          Call ACTUREGISTRO("Movisto", "CantIngre", CONDI$, 0, REGAF&, "NOMESS")
'          Call ACTUREGISTRO("Movisto", "CantSalid", CONDI$, 0, REGAF&, "NOMESS")
'          If TRIM$(LOTEVIE$) <> "" Then
'             Call ACSALOTE(CIXI%, LOTEVIE$, "NOMESS")
'          End If
'          'BORRA LOS REGISTROS
'          Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&, "NOMESS")
'          '
'          Dim Movisto As ADODB.Recordset
'          Set Movisto = New ADODB.Recordset
'          strsql = "SELECT * FROM MOVISTO WHERE " + CONDI$
'25        On Error Resume Next
'          Movisto.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'             On Error GoTo 0
'             Call CapturaErrorOpen
'             GoTo 25
'          End If
'          On Error GoTo 0
'          '
'          SUMASI = 0
'          If CANTASI >= FRACMIN(CIXI%) Then
'                With Movisto
'                   .AddNew
'                   'GENERA NUEVO REGISTRO CON LA INFORMACION GUARDADA ANTERIORMENTE
'                   For U& = 1 To CACAM&
'                      Movisto.Fields(U& - 1).Value = VARS(U&)
'                   Next U&
'                   !CANTSALID = CANTASI: SUMASI = SUMASI + CANTASI
'                   !IDENLOTE = LOTE$
'                   .Update
'                   Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
'                End With
'             End If
'          '
'          RESTO = XVALO(Label24) - SUMASI
'          If RESTO >= FRACMIN(CIXI%) Then
'            With Movisto
'               .AddNew
'               For U& = 1 To CACAM&
'                  Movisto.Fields(U& - 1).Value = VARS(U&)
'               Next U&
'               !CANTSALID = RESTO
'               !IDENLOTE = ""
'               .Update
'            End With
'          End If
'          '
'          Movisto.Close: Set Movisto = Nothing
'          MOVISE.Close: Set MOVISE = Nothing
'
'            .MoveNext
'        Loop
'   End With
'   RS1.Close
'   Set RS1 = Nothing
'
'End Sub
'


Sub LOTES_FACTURA(CI%, NC%, NFACTUR$)
        'SOLUCION DADA POR MEDITEA EL CASO QUE SEA NOTA DE CREDITO Y ANULA FACTURA
        'EN ESTA RUTINA MUESTRA LOS LOTES SOLO DE LA FACTURA EN CUESTION
        CONDI$ = "DOSECLAR  = '" & NFACTUR$ & "'"
        CONDI$ = CONDI$ + " AND NUME_CLIE = " & NC% & ""
        CONDI$ = CONDI$ + " AND COD_INTE =  " & CI%
        CONDI$ = CONDI$ + " AND CANTSALID > 0 "
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With FACTTMP
          Do While Not .EOF ' CARGA TODOS LOS ITEMS DEL MOVISTO RELACIONADO CON FACTURA
           CI% = XVALO(!cod_inte)
           LOTE$ = TRIM$(SAFETEXT$(!idenlote))
           CANTI = XVALO(!CANTSALID)
30         TTXX$ = ""
            Call REPLA(TTXX$, LOTE$, 1, 12)
            Call REPLA(TTXX$, FORMATNUM$(CANTI, "F10." & CNTDC$ & ""), 33, 10)
            If CANTI > 0 Then List2.AddItem (TTXX$)
           .MoveNext
          Loop
        End With
        FACTTMP.Close
        Set FACTTMP = Nothing

End Sub

Private Sub LBLCLIENTE_Change()
   NC = XVALO(LBLCLIENTE)
   If NC > 0 Then
     Label5 = rasocli$(NC)
   Else
     Label5 = ""
   End If
   
End Sub

Private Sub lblfactura_Change()
     'DEBERA PEGAR EL CODIGO EN LBLCODIGO
     'ESTO ES PARA CUANDO VENGA DE LA RECEPCION EN LA RECEPCION DEBERIA RECORRER , MEDIANTE UNA CONSULTA LOS MOVIMIENTOS RT DONDE
     'COD_INTE = CODIGO, AND  CODIMOVI = 'RT' AND (IDENLOTE IS NULL OR IDENLOTE = '')"
     'CADA CODIGO DEBERA LLAMAR A ESTE FRM PONIENDO RECEPCION% = 1 PEGAN
     '
     If RECEPCION% > 0 Then
        J& = 0
        CI1% = XVALO(LBLCODIGO)
        If CI1% < 1 Then Unload Me
        Me.List2.Clear
        REMITOX$ = lblfactura
        SQLX$ = "SELECT CANTSALID  ,CANTINGRE, DESCRIP,DOSECLAR, NUME_CLIE,COD_INTE,COD_EXTE, VALOUNITIVA, NUORITEM,Idenlote,NumeOrdPed,NUPEDCLI FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT'"
        SQLX$ = SQLX$ + " AND (IDENLOTE IS NULL OR IDENLOTE = '')"
        SQLX$ = SQLX$ + " AND DOPRILAR = '" & REMITOX$ & "'"
        SQLX$ = SQLX$ + " AND COD_INTE =" & CI1%
        SQLX$ = SQLX$ + " ORDER BY COD_INTE ASC"
        Set RS1 = READSET(SQLX$)
        With RS1
            Do While Not .EOF
                 CI1% = XVALO(!cod_inte)
                 If ESTRAZABLE%(CI1%) = 1 Then
                     CANTI = 0
                     If XVALO(!CANTINGRE) > 0 Then
                       CANTI = XVALO(!CANTINGRE)
                     Else
                       CANTI = XVALO(!CANTSALID)
                     End If
                     '
                     J& = J& + 1
    
                    TXTX$ = ""
                    
                    Call REPLA(TXTX$, SAFETEXT$(!Cod_Exte), 1, 20)
                    Call REPLA(TXTX$, TRIM$(SAFETEXT$(!Descrip)), 21, 40)
                    Call REPLA(TXTX$, FORMATNUM$(CANTI, "F10.1"), 62, 11)
                    Call REPLA(TXTX$, SAFETEXT$(!idenlote), 74, 16)
                    Call REPLA(TXTX$, FORMATNUM$(XVALO(!VALOUNITIVA), "F12.2"), 91, 12)
                    Call REPLA(TXTX$, FORMATNUM$(XVALO(PROVEEDOR_SEGUN_PEDIDO(XVALO(!NUPEDCLI), XVALO(!NUMEORDPED), CI1%)), "F7.0"), 104, 8)
                    Call REPLA(TXTX$, FORMATNUM$(XVALO(!NuOrItem), "F4"), 120, 4)
                    ASIGLOAHP.List2.AddItem (TXTX$)
    
                End If
                .MoveNext
            Loop
        End With
        RS1.Close
        Set RS1 = Nothing
        If J& > 0 Then
           Call Command15_Click
        Else
           Unload Me
        End If
     End If
End Sub


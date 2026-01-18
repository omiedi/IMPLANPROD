VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form APRONP07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Pedidos de Clientes."
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4320
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PEDIDOS DE CLIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   915
         ScaleHeight     =   600
         ScaleWidth      =   2520
         TabIndex        =   6
         Top             =   1380
         Width           =   2550
         Begin VB.CommandButton Command6 
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
            Height          =   630
            Left            =   0
            Picture         =   "APRONP07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Listado Pendientes de Aprobación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   735
            TabIndex        =   12
            Top             =   105
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   600
         ScaleHeight     =   600
         ScaleWidth      =   2520
         TabIndex        =   5
         Top             =   885
         Width           =   2550
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
            Height          =   630
            Left            =   0
            Picture         =   "APRONP07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Ver por Pantalla Pedidos de Clientes Pendientes de Aprobación."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver Pendientes de Aprobación."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   720
            TabIndex        =   11
            Top             =   50
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   285
         ScaleHeight     =   615
         ScaleWidth      =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   2550
         Begin VB.CommandButton Command43 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   1970
            Picture         =   "APRONP07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Revertir Pedidos Aprobados"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command14 
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
            Height          =   630
            Left            =   0
            Picture         =   "APRONP07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Aprobar Pedidos de Clientes."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprobar Pedidos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   660
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   3960
      ScaleHeight     =   4035
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   -105
      Width           =   1380
      Begin VB.CommandButton Command2 
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
         Height          =   640
         Left            =   105
         Picture         =   "APRONP07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   640
         Left            =   105
         Picture         =   "APRONP07.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "APRONP07.frx":0EBC
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "APRONP07.frx":10F0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Autorizar Pedidos con Motivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   3600
      Begin VB.PictureBox Picture5 
         Height          =   690
         Left            =   600
         ScaleHeight     =   630
         ScaleWidth      =   2490
         TabIndex        =   14
         Top             =   360
         Width           =   2550
         Begin VB.CommandButton Command8 
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
            Height          =   650
            Left            =   0
            Picture         =   "APRONP07.frx":116A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Autorización de Pedidos con Motivos"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label4 
            Caption         =   "Autorización de Prestamos - Giras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   720
            TabIndex        =   15
            Top             =   105
            Width           =   1830
         End
      End
   End
End
Attribute VB_Name = "APRONP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GENERAOF_ENVAPLAST_APRO(NPED&)

'    Screen.MousePointer = 11
'    Avance.PROCESOX$ = "GENERAOF_ENVAPLAST_APRO" & "/" & NPED&
'    On Error Resume Next
'    Avance.Show 1
'    On Error GoTo 0
'    Screen.MousePointer = 1
'    Exit Sub

     Dim obj As New RECORXET
     Set RS11 = obj.RS_DetaPedido(NPED&)

     CODI$ = ""
     With RS11
      Do While Not .EOF
          NC = XVALO(!NROCLIE)
          CIXI% = XVALO(!CODIINT)
          NUOR% = XVALO(!NuOrItem)
          CAN = XVALO(!CanPed)
          If CAN > 0.05 Then
            If CIXI% > 0 Then
               If TIMATE%(CIXI%) < 2 Then GoTo 50 ' PARA QUE NO GENERE OT POR MATERIAS PRIMAS.
               FENTRESOL% = CVINT(!FeSolEnt)
               If TRIM$(APROBO$(CIXI%)) = "" Then
                 If CODI$ = "" Then
                   CODI$ = CODI$ + CODEXT$(CIXI%)
                   Articulo$ = " El "
                   GoTo 50
                 Else
                   CODI$ = CODI$ & " , " & CODEXT$(CIXI%)
                   Articulo$ = " Los "
                   GoTo 50
                 End If
              End If
              Call FORPED07.GENERAR_OF_SIMPLE(CIXI%, NC, CAN, FENTRESOL%, NPED&, OrdenCompraPed$(NPED&), RETORNO_NORFA$, NUOR%)
              FILTX$ = TRIM$(Str$(CIXI%)) & ";" & TRIM$(Str$(NPED&)) & ";" & TRIM$(RETORNO_NORFA$) & ";" & TRIM$(MEDIDA_ENVAPLAST$(CIXI%, RETCORTE#))
              Call STDFORM("ORTR-ENV", FILTX$, "PRINT")
           End If
          End If
50       .MoveNext
      Loop
     End With

    If CODI$ <> "" Then
        TXX1$ = "" & Articulo$ & "Código" & Mid$(Articulo$, 3, 1) & " " & CODI$ & " Faltan ser Aprobados No se Generará la Orden de Trabajo Para Estos"
        Call COMUNI(TXX1$)
    End If
    '
    On Error Resume Next
    Set obj = Nothing
    RS11.Close
    Set RS11 = Nothing


End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
     '
   If DERACCE%("APRONP07", "Aprobación de Pedidos") < 2 Then
     Command14.Enabled = True
     Exit Sub
    Else
     If COMALTER%("¿Realmente Desea\APROBAR Pedidos?\\No, Abandonar\Si, Aprobar Pedidos") <> 2 Then
       Command14.Enabled = True
       Exit Sub
     End If
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "PROCESO DE APROBACIÓN DE PEDIDOS"
         
     Call ABREPEDCLI
     ' Selecciono todas las N/P'S pendientes de aprobación.
2    SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE STATUS = -1 "
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     Dim PEDENCTEMP As ADODB.Recordset
     Set PEDENCTEMP = New ADODB.Recordset
     PEDENCTEMP.CursorLocation = adUseClient
25   On Error Resume Next
      PEDENCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
      
     'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     jj& = 0
     Call COPYSTRU("INDIPEDI", "IDPEDPEN")
     With PEDENCTEMP
       On Error Resume Next
       .MoveFirst
       '\\\OT-05-0252-WAR-04/05/05///
       If Err Then
         Call COMUNI("No Existen Pedidos Pendientes de Aprobación.")
         Command14.Enabled = True
         Screen.MousePointer = 1
         On Error GoTo 0
         Exit Sub
       End If
       '\\\OT-05-0252-WAR-FIN///
       Do While Not .EOF
       'ANTIGUO CODIGO ANTES DE \\\OT-05-0682-WAR-18/10/05///
'         NPX& = !NroPed
'         REFEPED$ = !Referencia
'         NCLIE = !NroClie
'         FEX = CVINT(!FECHEMIS)
'         ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'         YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'         JJ& = JJ& + 1
'         Call GRAREG("IDPEDPEN", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE = !NROCLIE
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE)) + ANULX$
             jj& = jj& + 1
             Call GRAREG("IDPEDPEN", TTXX$, jj&)
             '\\\OT-05-0682-WAR-FIN///
       .MoveNext
       Loop
     End With
     PEDENCTEMP.Close
     Set PEDENCTEMP = Nothing
     '
     Call SETULTREG("IDPEDPEN", jj&)
     Call CIERRARCH("IDPEDPEN")
     Call FRESHECHO("IDPEDPEN")
     Screen.MousePointer = 1
     '
3    Call SELECHO("IDPEDPEN/Indice General de Pedidos Pendientes de Aprobación.")
     NPX& = XVALO(VALACT1$("IDPEDPEN"))
     If NPX& < 1 Then GoTo 99
     '
     Call ABREPEDCLI
     'jandy sql
     ' Envio las dos instrucciones al mismo tiempo se puede hacer en un join
     strSQL = "Update PedEnca set Status=0,FApro_Nope=getdate(),"
     strSQL = strSQL & " Usua_Apro= " & USNBR% Mod 100
     strSQL = strSQL & " where NroPed=" & NPX
     
     strSQL = strSQL & " Update PedDeta set Status=0"
     strSQL = strSQL & " where NroPed=" & NPX
      FLEXCONN.Execute (FILTSQL$(strSQL))
     On Error Resume Next
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Aprobado Por: " & USERNAME$, 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     TTXYZ$ = TTXYZ$ + Chr$(10) + Chr$(13) + String$(88, "=") + Chr$(10) + Chr$(13)


     Call AGREGA_ANOTAPED(NPX&, TTXYZ$)
     
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "ENVAPLAST") > 0 Then
        Call GENERAOF_ENVAPLAST_APRO(XVALO(NPX))
     End If
     
'     With PedEnca
'       .FindFirst "NroPed = " & NPX & ""
'       If .NoMatch = False Then
'         .Edit
'         !Status = 0
'         !Usua_Apro = USNBR% Mod 100
'         !FApro_Nope = Now
'         .Update
'       End If
'     End With
'     '
    ' PRX% = 0
    ' With PedDeta
     
'4      If PRX% = 0 Then
'           .FindFirst "NroPed = " & NPX & ""
'         Else
'           .FindNext "NroPed = " & NPX & ""
'       End If
'       If .NoMatch = False Then
'         If !Status < 8 Then  ' excluye los anulados
'           .Edit
'           !Status = 0
'           .Update
'         End If
'         PRX% = 1
'         GoTo 4
'       End If
'     End With
     '
'     PedEnca.Close
'     PedDeta.Close
     
     Call COMUNI("Se ha Aprobado con Éxito\el Pedido Nº: " & NPX & "")
'     PedEnca.FindFirst "STATUS = -1"
'     If PedEnca.NoMatch = False Then
'       GoTo 2
'     End If
     '
   End If
   '
99 Command14.Enabled = True
End Sub
Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command3_Click()
     Call MENSERR(24, "No Disponible por el Momento.")
End Sub

Sub Command4_Click()
     Command4.Enabled = False
     'Selecciono todas las N/P pendientes de aprobación.
         '
    Command4.Enabled = False
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM PEDENCA "
    SQLX$ = SQLX$ + " WHERE STATUS = -1 "
    SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
    Dim PEDENCTEMP As ADODB.Recordset
    Set PEDENCTEMP = New ADODB.Recordset
    PEDENCTEMP.CursorLocation = adUseClient
25   On Error Resume Next
      PEDENCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
    'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    jj& = 0
    Call COPYSTRU("INDIPEDI", "IDPEDPEN")
    With PEDENCTEMP
      On Error Resume Next
      .MoveFirst
      '\\\OT-05-0251-WAR-04/05/05///
      If Err Then
        On Error GoTo 0
        Call COMUNI("No existen Pedidos\Pendientes de Aprobación.")
        Screen.MousePointer = 1
        Command4.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '\\\OT-05-0251-WAR-FIN///
      Do While Not .EOF
         'CODIGO ANTERIOR A \\\OT-05-0682-WAR-18/10/05///
'        NPX& = !NroPed
'        REFEPED$ = !Referencia
'        NCLIE = !NroClie
'        FEX = CVINT(!FECHEMIS)
'        ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'        YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'        JJ& = JJ& + 1
'        Call GRAREG("IDPEDPEN", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NROPED
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NROOCOM
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE = !NROCLIE
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE)) + ANULX$
             jj& = jj& + 1
             Call GRAREG("IDPEDPEN", TTXX$, jj&)
             '\\\OT-05-0682-WAR-FIN///
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("IDPEDPEN", jj&)
    Call CIERRARCH("IDPEDPEN")
    Call FRESHECHO("IDPEDPEN")
    Screen.MousePointer = 1
    '
3   Call SELECHO("IDPEDPEN/Indice General de Pedidos Pendientes de Aprobación.")
    NPX& = XVALO(VALACT1$("IDPEDPEN"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    
    
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
4     DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
    
     If DOCUNU& > 0 Then
           Call MUESTRADOC(DOCUNU&)
           Screen.MousePointer = 1
           GoTo 3
    Else
           If TIDOCUM$ = "Pedido de Cliente" Then
             TIDOCUM$ = "Pedido Interno"
             GoTo 4
           End If
           
           Screen.MousePointer = 1
           Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.")
           GoTo 3
    End If
    
''''4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
''''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
''''      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Then
''''        PPXX% = InStr(XX$, "Nro.")
''''        If PPXX% > 0 Then
''''          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
''''        End If
''''        If InStr(XX$, NUDOCU$) > 4 Then
''''          DOCUNU& = CVS(Left$(XX$, 4))
''''          GoTo 5
''''        End If
''''      End If
''''    Next UI&
''''    If TIDOCUM$ = "Pedido de Cliente" Then
''''      TIDOCUM$ = "Pedido Interno"
''''      GoTo 4
''''    End If
''''    Screen.MousePointer = 1
''''    Call COMUNI("Documento no Encontrado")
''''    GoTo 99
''''    '
''''5   Call MUESTRADOC(DOCUNU&)
''''    GoTo 3
    '
99  Command4.Enabled = True
   End Sub

Private Sub Command43_Click()
   Command43.Enabled = False
   If DERACCE%("APRONP07", "Aprobación de Pedidos") > 1 Then
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     '"PEDENCA+PEDDETA/PEDENCA.NroPed = PEDDETA.NroPed"
     CONDI$ = "Status >= 0 And Status < 8 ORDER BY NROPED DESC"
     Call CARGALISEL("REVPEAPR", "PEDENCA+DEUDOR12/PEDENCA.NROCLIE = DEUDOR12.Nume_Cli ", "NROPED/F8;FECHEMIS/D8;NROCLIE/F8;Raso_Cli/A24;NROOCOM/A116", CONDI$, "NOMESS")
     
     Call SELECHO("REVPEAPR/Nro.Ped.     FECHA                        Cta.    Razón Social                                      Orden de Compra")
     NPX& = XVALO(VALACT1$("REVPEAPR"))
     If NPX& < 1 Then GoTo 99
     
     ' Envio las dos instrucciones al mismo tiempo se puede hacer en un join
     strSQL = "Update PedEnca set Status= -1,FApro_Nope=getdate(),"
     strSQL = strSQL & " Usua_Apro= " & USNBR% Mod 100
     strSQL = strSQL & " where NroPed=" & NPX
     
     strSQL = strSQL & " Update PedDeta set Status=-1"
     strSQL = strSQL & " where NroPed=" & NPX
     FLEXCONN.Execute (FILTSQL$(strSQL))
      
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Marca Pend. Aprobación Por: " & USERNAME$, 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     TTXYZ$ = TTXYZ$ + Chr$(10) + Chr$(13) + String$(88, "=") + Chr$(10) + Chr$(13)
     Call AGREGA_ANOTAPED(NPX&, TTXYZ$)
     '
     Call COMUNI("Pedido Seleccionado Grabado como Pendiente de Aprobación")
     '
   End If
99 Command43.Enabled = True

End Sub

Sub Command6_Click()
   Command6.Enabled = False
   Screen.MousePointer = 11
   '
   Call ABREPEDCLI       ' SE ABRE CONEXION EN CASO QUE SE ACCEDA DESDE EL MENU
           SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE Status < 0 "
           '
           Dim PEDTEMP As ADODB.Recordset
           Set PEDTEMP = New ADODB.Recordset
           PEDTEMP.CursorLocation = adUseClient
25         On Error Resume Next
           PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
           
           'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
           '
           With PEDTEMP
             On Error Resume Next
             .MoveFirst
             If Err Then
               Screen.MousePointer = 1
               Call COMUNI("No Existen Pedidos\Pendientes de Aprobación")
               On Error GoTo 0
               GoTo 99
             End If
             On Error GoTo 0
             '

             jj& = 0
             Call COPYSTRU("LIPEDPE", "LIPEDPE1")
             Y$ = REGBLAN("LIPEDPE1")
             '
             Do While Not .EOF
                    '

                    ' esto es para resolver inconsistencia de status entre el
                    ' encabezado y el detalle del pedido - EPB 12/05/05
                    NUPE& = !NROPED
                    'PedEnca.FindFirst "NroPed = " & NUPE&
                    'If PedEnca.NoMatch = False Then
                    '  PedEnca.Edit
                    '  PedEnca!Status = (-1)
                    '  PedEnca.Update
                    'End If
                    
                    strSQL = "Update PedEnca set Status="
                    strSQL = strSQL & " case when Status <> -1 then -1 else Status end "
                    strSQL = strSQL & " where NroPed=" & NUPE&
                     FLEXCONN.Execute (FILTSQL$(strSQL))

                    CAPEN = XVALO(!CanPed) - XVALO(!CANTENT)
                    '
                    If CAPEN < 0 Then CAPEN = 0
                      CANOASI = CAPEN - XVALO(!CanAsig)
                    '
                    If CANOASI < 0 Then CANOASI = 0
                      MONEMISPED% = XVALO(!MoneEmis)
                    PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(XVALO(!NROCLIE))) / 100 * TICAMONE(MONEMISPED%)
                    '
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    '
                    Call REPLA(Y$, MKI$(XVALO(!NROCLIE)), 1, 2)
                    Call REPLA(Y$, MKI$(XVALO(!CODIINT)), 3, 2)
                    Call REPLA(Y$, !TCOLOR, 5, 16)
                    Call REPLA(Y$, !TTALLE, 21, 8)
                    Call REPLA(Y$, MKS$(!TLargo), 29, 4)
                    Call REPLA(Y$, MKS$(!TAncho), 33, 4)
                    Call REPLA(Y$, MKS$(!TEspesor), 37, 4)
                    Call REPLA(Y$, MKS$(!NROPED), 47, 4)
                     FENTRESOL% = CVINT(!FeSolEnt)
                    Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                    Call REPLA(Y$, MKS$(!CanPed), 81, 4)
                    Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                    Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                    Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                    Call REPLA(Y$, MKS$(XVALO(!CANTENT)), 97, 4)
                    Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                    Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                    Call REPLA(Y$, Space$(34), 127, 34)
                    jj& = jj& + 1

                    Call GRAREG("LIPEDPE1", Y$, jj&)

                    '
                .MoveNext
             Loop
             On Error GoTo 0
           End With
           '

           Call SETULTREG("LIPEDPE1", jj&)
           Call CIERRARCH("LIPEDPE1")

           Call FINAL("*LIPENC1")

99  Screen.MousePointer = 1
    Command6.Enabled = True

90 End Sub

Private Sub Picture7_Click()
End Sub

Private Sub Command8_Click()
     '
     Command8.Enabled = False
     If DERACCE%("AUTOMOTI", "Autorización de Pedidos Con Motivos") < 2 Then GoTo 99
     '
     'PRESENTO LISTA DE SELECCION DE PEDIDOS CON MOTIVO PENDIENTES SIN AUTORIZACIONES
10   CONDI$ = "PEDENCA.Status < 9 AND (OV_AUTORIZA1 = '' OR   OV_AUTORIZA1 IS NULL OR OV_AUTORIZA2 = '' OR   OV_AUTORIZA2 IS NULL)"
     CONDI$ = CONDI$ + " AND OV_MOTIVO > 0 AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
     CONDI$ = CONDI$ + " AND PEDDETA.Status < 8 ORDER BY PEDENCA.NROPED"
     Call CARGALISEL("!INDIPEDI", "PEDENCA+PEDDETA/PEDENCA.NroPed = PEDDETA.NroPed", "PEDENCA.NroPed/F6;PEDENCA.REFERENCIA/A32", CONDI$, "NOMESS/DISTINCT")
     Call SELECHO("!INDIPEDI/Pedidos Pendientes de Autorización")
     NPED& = XVALO(VALACT1$("!INDIPEDI"))
     If NPED& < 1 Then GoTo 99
     '
     'PRESENTO VENTANA DE CARGA DE DATOS
     CONDI$ = "NROPED = " & NPED&
     NUMOT% = XVALO(VALOBADA("PEDENCA", "OV_Motivo", CONDI$, "NOMESS"))
     DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
     '
     FILTX$ = TRIM$(Str$(NPED&)) & ";" & TRIM$(Str$(NUMOT%)) & ";" & TRIM$(DESCMOTI$)
     Call STDFORM("DESMO", FILTX$)
     GoTo 10
     '
99   Command8.Enabled = True
End Sub

Private Sub Form_Load()

    ' SI ES MEDITEA MUESTRA BOTON DE AUTORIZACIÓN DE PEDIDOS CON MOTIVOS
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
       Me.Height = 4400
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)

End Sub


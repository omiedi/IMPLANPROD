VERSION 5.00
Begin VB.Form OPRETRA09 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos de Traslado - Opciones"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   4050
      ScaleHeight     =   4635
      ScaleWidth      =   1185
      TabIndex        =   8
      Top             =   -120
      Width           =   1245
      Begin VB.CommandButton Command7 
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
         Height          =   700
         Left            =   150
         Picture         =   "OPRETRA09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1290
         Width           =   800
      End
      Begin VB.CommandButton Command2 
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
         Height          =   700
         Left            =   150
         Picture         =   "OPRETRA09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   405
         Width           =   800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         Picture         =   "OPRETRA09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   6960
         Width           =   800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Trasla."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         Picture         =   "OPRETRA09.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   4920
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   6000
         Width           =   650
         Begin VB.Frame Frame10 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -225
         Picture         =   "OPRETRA09.frx":0DC8
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -250
         Picture         =   "OPRETRA09.frx":2CEA
         Top             =   6240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Setup - Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   1680
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
         Top             =   420
         Width           =   2640
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
            Left            =   0
            Picture         =   "OPRETRA09.frx":4C0C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso Directo a Configuración de Funcionaliento"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Configuracion de Funcionamiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   7
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1635
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   480
         ScaleHeight     =   585
         ScaleWidth      =   2715
         TabIndex        =   1
         Top             =   420
         Width           =   2775
         Begin VB.CommandButton Command15 
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
            Left            =   2160
            Picture         =   "OPRETRA09.frx":504E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Re-impresion de Remitos de Traslado"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command25 
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
            Left            =   0
            Picture         =   "OPRETRA09.frx":56B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación General de Remito / Orden de Despacho Emitidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anula Remito / Re-impresion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   3
            Top             =   105
            Width           =   1740
         End
      End
   End
End
Attribute VB_Name = "OPRETRA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    OPREMIT09.Show 1
    Command1.Enabled = True
    '
End Sub




Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Dim CISL%(2048), CANSL#(2048), DEPSL%(2048)
    '
    CMO$ = TRIM$(Control$("", "CODRETRA"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    DES1$ = CVDAT(DES%)
    HAS1$ = CVDAT(HAS%)
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, ReferCor FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    'SQLX$ = SQLX$ + "ORDER BY val(MID$(DoPriCor,4)) asc "
    Call APERBASDAT("STOCKS")
    'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      j& = 0: TXCONTROL$ = ""
      Do While (.EOF = False)
        If Left$(UCase$(!ReferCor), 4) <> "ANUL" Then   ' para evitar comprobantes arruinados
          CLI% = XVALO(!Nume_Clie)
          REMI$ = SAFETEXT$(!DOPRICOR)
          If TRIM$(REMI$) <> "" Then
            If InStr(TXCONTROL$, REMI$) < 1 Then
              FF% = CVINT(!FECHMOV)
              '
              Z$ = Space$(40) + String$(8, 0)
              Call REPLA(Z$, REMI$, 1, 10)
                RASOX$ = TRIM$(RASOCLI$(CLI%))
                If RASOX$ = "" Then RASOX$ = "CLIENTE Nro." + TRIM$(Str$(CLI%))
              Call REPLA(Z$, RASOX$, 11, 30)
              Call REPLA(Z$, MKS$(U&), 41, 4)
              Call REPLA(Z$, MKI$(FF%), 45, 2)
              Call REPLA(Z$, MKI$(CLI%), 47, 2)
              j& = j& + 1
              Call GRAREG("!SELREMI", Z$, j&)
              TXCONTROL$ = TXCONTROL$ + AJUSTI$(REMI$, 24)
            End If
          End If
        End If
      .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '
    Call SETULTREG("!SELREMI", j&)
    Call CIERRARCH("!SELREMI")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEXI% = CVI(Mid$(RERE$, 45, 2))
    FEMOV = FETEXCOR1$(FEXI%)
    CLISEL% = CVI(Mid$(RERE$, 47, 2))
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEXI) + "\   \\Abandonar\Confirma Anular"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    '\\\OT-5-0400-FG-14/06/05///
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND FechMov = '" & FEMOV & "'"
    'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim MoviTemp As ADODB.Recordset
    Set MoviTemp = New ADODB.Recordset
    MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MoviTemp
      On Error Resume Next
      .MoveFirst
      If Err > 1 Then
        On Error GoTo 0
        Call MENSERR(24, "No se Encontró Detalle del Remito")
        GoTo 10
      End If
      '
      On Error GoTo 0
      ' VERIFICA SI EL REMITO ESTA FACTURADO
      Do While Not .EOF
        TIDOCSEC$ = ""
        On Error Resume Next
        TIDOCSEC$ = UCase$(Mid$(TRIM$(SAFETEXT$(!DoSecCor)), 1, 2))
        On Error GoTo 0
        If TIDOCSEC$ = "FC" Then
          NUMERFAC$ = TRIM$(SAFETEXT$(!DoSecCor))
          Call APERBASDAT("CABAN")
          Screen.MousePointer = 11
          'CajaBanc.FindFirst "TipoMovim = 'FVEN' AND CodiCom_Cor = '" & NUMERFAC$ & "'"
          SQLX$ = " SELECT * FROM CajaBanc "
          SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
          SQLX$ = SQLX$ + " AND CodiCom_Cor = '" & NUMERFAC$ & "'"
          Dim CAJJJABANC As ADODB.Recordset
          Set CAJJJABANC = New ADODB.Recordset
          CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          With CAJJJABANC
            On Error Resume Next
            .MoveFirst
            'If CajaBanc.NoMatch = False Then
            If Err < 1 Then
              On Error GoTo 0
              If (XVALO(!IDEBECOMP) + XVALO(!IHABECOMP)) < 0.005 Then GoTo 16
            End If
            On Error GoTo 0
            Call COMUNI("Imposible Anular. Remito Ya Facturado\Mediante Comprobante " & NUMERFAC$ & "." _
                      & "\   \Debe Anular la Factura y Re-Intentar.")
            
          End With
          GoTo 10
        End If
      .MoveNext
      Loop
    End With
    '
16  Call APERBASDAT("STOCKS")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND FechMov = '" & FEMOV & "'"
    Dim MOVVVISTO As ADODB.Recordset
    Set MOVVVISTO = New ADODB.Recordset
    MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

'20  MoviTemp.FindFirst SQLX$
    
'21  If MoviTemp.NoMatch = False Then
      ' VERIFICA SI EL REMITO TIENE ASOCIADO UN PEDIDO
    With MOVVVISTO
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While (Not .EOF)
          NROPEDID& = XVALO(!NUPEDCLI)
          If NROPEDID& > 0 Then
          ' ACTUALIZA LOS PEDIDOS QUE FUERON ENTREGADOS CON EL REMITO
            CODINTE% = XVALO(!Cod_Inte)
            CANTSALD# = XVALO(!CANTSALID)
            Call APERBASDAT("PEDCLI")
            SQLY$ = "SELECT * FROM PEDDETA "
            SQLY$ = SQLY$ + " where NroPed = " & NROPEDID& & " "
            SQLY$ = SQLY$ + "AND CodiInt = " & CODINTE% & " "
            SQLY$ = SQLY$ + "AND NroClie = " & CLISEL% & " "
            SQLY$ = SQLY$ + "AND CantEnt >= " & FORMATNUM$((CANTSALD# - 0.05), "F15.2") & " "
            SQLY$ = SQLY$ & " AND Status < 9 "
            Dim PEDETTTA As ADODB.Recordset
            Set PEDETTTA = New ADODB.Recordset
            'MsgBox sqly$
            PEDETTTA.Open FILTSQL$(SQLY$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  
            'PedDeta.FindFirst SQLY$
            With PEDETTTA
              On Error Resume Next
              .MoveFirst
              'If PedDeta.NoMatch = False Then
              If Err < 1 Then
                On Error GoTo 0
                'PedDeta.Edit
                ENTREGA# = XVALO(!CantEnt) - CANTSALD#
                SALDO# = XVALO(!CanSaldo) + CANTSALD#
                !CantEnt = ENTREGA#
                !CanSaldo = SALDO#
                !NroRemi = ""
                .Update
              End If
              On Error GoTo 0
            End With
            Set PEDETTTA = Nothing
          End If
          .MoveNext
        Loop
      End If
      On Error GoTo 0
      '
    End With
    Set MOVVVISTO = Nothing
    '
    SQLX$ = "DELETE FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND FechMov = '" & FEMOV & "'"
    '
    On Error Resume Next
    FLEXCONN.Execute SQLX$
    If Err Then
         Call MENSERR(24, "No fue Posible Completar la Anulación\del Remito '" + REMISEL$ + "'")
       Else
         CONDI$ = "Dopricor = '" & REMISEL$ & "'"
         Call ACTUREGISTRO("DETACARTEM", "Cant_Entre", CONDI$, 0, REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "DopriLar", CONDI$, "", REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "DopriCor", CONDI$, "", REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "FechEmis", CONDI$, CVDAT(0), REGAF&)
         Call ACTUREGISTRO("DETACARTEM", "NuorItem", CONDI$, 0, REGAF&)
         Call COMUNI("Remito " + REMISEL$ + " Anulado.")
    End If
    GoTo 10
    '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    
    Dim CISL%(2048), CANSL#(2048), DEPSL%(2048)
    '
    CMO$ = TRIM$(Control$("", "CODRETRA"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    Call FRESHECHO("SELMESA")
    '
10  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    DES1$ = FETEXCOR1$(DES%)
    HAS1$ = FETEXCOR1$(HAS%)
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, ReferCor FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    
    Call APERBASDAT("STOCKS")
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 10
      End If
      On Error GoTo 0
      '
      j& = 0: TXCONTROL$ = ""
      Do While (.EOF = False)
        If Left$(UCase$(!ReferCor), 4) <> "ANUL" Then   ' para evitar comprobantes arruinados
          CLI% = XVALO(!Nume_Clie)
          REMI$ = SAFETEXT$(!DOPRICOR)
          If TRIM$(REMI$) <> "" Then
            If InStr(TXCONTROL$, REMI$) < 1 Then
              FF% = CVINT(!FECHMOV)
              '
              Z$ = Space$(40) + String$(8, 0)
              Call REPLA(Z$, REMI$, 1, 10)
                RASOX$ = TRIM$(RASOCLI$(CLI%))
                If RASOX$ = "" Then RASOX$ = "CLIENTE Nro." + TRIM$(Str$(CLI%))
              Call REPLA(Z$, RASOX$, 11, 30)
              Call REPLA(Z$, MKS$(U&), 41, 4)
              Call REPLA(Z$, MKI$(FF%), 45, 2)
              Call REPLA(Z$, MKI$(CLI%), 47, 2)
              j& = j& + 1
              Call GRAREG("!SELREMI", Z$, j&)
              TXCONTROL$ = TXCONTROL$ + AJUSTI$(REMI$, 24)
            End If
          End If
        End If
      .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    '
    Call SETULTREG("!SELREMI", j&)
    Call CIERRARCH("!SELREMI")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    Screen.MousePointer = 1
    Call FRESHECHO("!SELREMI")
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    '
    NUDOCU$ = TRIM$(Mid$(REMISEL$, 4))
    If NUDOCU$ = "" Then GoTo 10
    TIDOCUM$ = "Remito Traslado"
    '
    DOCUNU& = 0
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, NUDOCU$, "NOMESS")
    GoTo 10
    '
99  Command15.Enabled = True
    
End Sub

Private Sub Command7_Click()
    Call HELPONLINE("REMITRAS")
    
'    FIN& = ULTREG("RTRASLA")
'    For U& = 1 To ULTREG("RTRASLA")
'      Call TRACE(24, U&, FIN&)
'      XX$ = REGLEIDO$("RTRASLA", U&)
'        VENVI% = CVI(Mid$(XX$, 115, 2))     ' veces enviado
'        REMI& = XVALO(Mid$(XX$, 6, 11))
'        If REMI& = 3853 Then
'            Call COMUNI("Veces " & CStr(VENVI%))
'            Exit For
'        End If
'    Next U&
    
End Sub


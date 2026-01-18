VERSION 5.00
Begin VB.Form APRONP04 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Pedidos de Clientes."
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Pedidos de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   210
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox Picture4 
         Height          =   690
         Left            =   1155
         ScaleHeight     =   630
         ScaleWidth      =   2490
         TabIndex        =   8
         Top             =   1500
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
            Height          =   640
            Left            =   0
            Picture         =   "APRONP04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label3 
            Caption         =   "Listado Pendientes de Aprobación"
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
            Left            =   735
            TabIndex        =   14
            Top             =   105
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   690
         Left            =   840
         ScaleHeight     =   630
         ScaleWidth      =   2490
         TabIndex        =   7
         Top             =   1005
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
            Height          =   640
            Left            =   0
            Picture         =   "APRONP04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ver por Pantalla Pedidos de Clientes Pendientes de Aprobación."
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label2 
            Caption         =   "Ver Pendientes de Aprobación."
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
            TabIndex        =   13
            Top             =   50
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   690
         Left            =   525
         ScaleHeight     =   630
         ScaleWidth      =   2490
         TabIndex        =   6
         Top             =   510
         Width           =   2550
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
            Height          =   640
            Left            =   0
            Picture         =   "APRONP04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Aprobar Pedidos de Clientes."
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label1 
            Caption         =   "Aprobar Pedidos de Clientes"
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
            Left            =   840
            TabIndex        =   12
            Top             =   50
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3585
      Left            =   4620
      ScaleHeight     =   3525
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -105
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   1890
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
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
         Picture         =   "APRONP04.frx":091E
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
         Picture         =   "APRONP04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "APRONP04.frx":0D72
         Top             =   2205
         Width           =   1515
      End
   End
End
Attribute VB_Name = "APRONP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Sub Command14_Click()
    Command14.Enabled = False
     '
   If DERACCE%("APRONP04", "Aprobación de Pedidos") < 2 Then
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
         
     ' Selecciono todas las N/P'S pendientes de aprobación.
2    SQLX$ = "SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + " WHERE STATUS = -1 "
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     JJ& = 0
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
'         NCLIE% = !NroClie
'         FEX = CVINT(!FECHEMIS)
'         ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'         YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'         JJ& = JJ& + 1
'         Call GRAREG("IDPEDPEN", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$
             JJ& = JJ& + 1
             Call GRAREG("IDPEDPEN", TTXX$, JJ&)
             '\\\OT-05-0682-WAR-FIN///
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("IDPEDPEN", JJ&)
     Call CIERRARCH("IDPEDPEN")
     Call FRESHECHO("IDPEDPEN")
     Screen.MousePointer = 1
     '
3    Call SELECHO("IDPEDPEN/Indice General de Pedidos Pendientes de Aprobación.")
     NPX& = XVALO(VALACT1$("IDPEDPEN"))
     If NPX& < 1 Then GoTo 99
     '
     Call ABREPEDCLI
     With PedEnca
       .FindFirst "NroPed = " & NPX & ""
31     If .NoMatch = False Then
         .Edit
         !Status = 0
         !Usua_Apro = USNBR% Mod 100
         !FApro_Nope = Now
         .Update
         .FindNext "NroPed = " & NPX & ""
         GoTo 31
       End If
     End With
     '
     PRX% = 0
     With PedDeta
4      If PRX% = 0 Then
           .FindFirst "NroPed = " & NPX & ""
         Else
           .FindNext "NroPed = " & NPX & ""
       End If
       If .NoMatch = False Then
         If !Status < 8 Then  ' excluye los anulados
           .Edit
           !Status = 0
           .Update
         End If
         PRX% = 1
         GoTo 4
       End If
     End With
     '
     PedEnca.Close
     PedDeta.Close
     '
     If CONTROL$("PEDIDO", "IMPRIME") = "ALAPROBA" Then
        TIDOCUM$ = "Pedido de Cliente"
        TIDOCUX$ = "Pedido ABIERTO"
        TIDOCUY$ = "Pedido Interno"
        VUELTA% = 1
        LOMOSTRO% = 0
        '
15      NUDOCU$ = TRIM$(Str$(NPX&))
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
          XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
          If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
            PPXX% = InStr(XX$, "Nro.")
            If PPXX% > 0 Then
              XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
            End If
            If InStr(XX$, NUDOCU$) > 4 Then
              DOCUNU& = CVS(Left$(XX$, 4))
              GoTo 5
            End If
          End If
        Next UI&
        '
        Screen.MousePointer = 1
        If VUELTA% = 1 Then GoTo 10
        If LOMOSTRO% < 1 Then
           Call COMUNI("Imposible Imprimir este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
        End If
        GoTo 90
        '
5       Call MUESTRADOC(DOCUNU&, "PRINT")
        LOMOSTRO% = 1
        '
10      If VUELTA% = 1 Then
           VUELTA% = 2
           TIDOCUM$ = "Orden de Despacho"
           TIDOCUX$ = "Orden de Despacho"
           TIDOCUY$ = "Orden de Despacho"
           GoTo 15
        End If
        '
     End If
     '
90   Call COMUNI("Se ha Aprobado con Éxito\el Pedido Nº: " & NPX & "")
     PedEnca.FindFirst "STATUS = -1"
     If PedEnca.NoMatch = False Then
       GoTo 2
     End If
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
    Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
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
'        NCLIE% = !NroClie
'        FEX = CVINT(!FECHEMIS)
'        ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'        YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'        JJ& = JJ& + 1
'        Call GRAREG("IDPEDPEN", YY$, JJ&)
            '\\\OT-05-0682-WAR-18/10/05///
             NRODES& = !NroPed
             TTXX$ = MKS$(NRODES&)
               NOCO$ = !NroOcom
             If NOCO$ <> "" Then
               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
             End If
               FEX = CVINT(!FECHEMIS)
             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
               NCLIE% = !NroClie
               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + ANULX$
             JJ& = JJ& + 1
             Call GRAREG("IDPEDPEN", TTXX$, JJ&)
             '\\\OT-05-0682-WAR-FIN///
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("IDPEDPEN", JJ&)
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
4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    If TIDOCUM$ = "Pedido de Cliente" Then
      TIDOCUM$ = "Pedido Interno"
      GoTo 4
    End If
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    GoTo 3
    '
99  Command4.Enabled = True
   End Sub

Sub Command6_Click()
   Command6.Enabled = False
   Screen.MousePointer = 11
   '
   Call ABREPEDCLI
           SQLX$ = " SELECT * "
           SQLX$ = SQLX$ + " FROM PEDDETA "
           SQLX$ = SQLX$ + " WHERE Status < 0 "
           '
           Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
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
             JJ& = 0
             Call COPYSTRU("LIPEDPE", "LIPEDPE1")
             Y$ = REGBLAN("LIPEDPE1")
             '
             Do While Not .EOF
                    '
                    ' esto es para resolver inconsistencia de status entre el
                    ' encabezado y el detalle del pedido - EPB 12/05/05
                    NUPE& = !NroPed
                    PedEnca.FindFirst "NroPed = " & NUPE&
                    If PedEnca.NoMatch = False Then
                      PedEnca.Edit
                      PedEnca!Status = (-1)
                      PedEnca.Update
                    End If
                    '
                    CAPEN = !CANPED - !CantEnt
                    '
                    If CAPEN < 0 Then CAPEN = 0
                      CANOASI = CAPEN - !CanAsig
                    '
                    If CANOASI < 0 Then CANOASI = 0
                      MONEMISPED% = !MoneEmis
                    PREUNI = !PREUNID * (100 - PORDESCLI(!NroClie)) / 100 * TICAMONE(MONEMISPED%)
                    '
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    '
                    Call REPLA(Y$, MKI$(!NroClie), 1, 2)
                    Call REPLA(Y$, MKI$(!CodiInt), 3, 2)
                    Call REPLA(Y$, !TColor, 5, 16)
                    Call REPLA(Y$, !TTalle, 21, 8)
                    Call REPLA(Y$, MKS$(!TLargo), 29, 4)
                    Call REPLA(Y$, MKS$(!TAncho), 33, 4)
                    Call REPLA(Y$, MKS$(!TEspesor), 37, 4)
                    Call REPLA(Y$, MKS$(!NroPed), 47, 4)
                     FENTRESOL% = CVINT(!FeSolEnt)
                    Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
                    Call REPLA(Y$, MKS$(!CANPED), 81, 4)
                    Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                    Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                    Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                    Call REPLA(Y$, MKS$(!CantEnt), 97, 4)
                    Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                    Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                    Call REPLA(Y$, Space$(34), 127, 34)
                    JJ& = JJ& + 1
                    Call GRAREG("LIPEDPE1", Y$, JJ&)
                    '
                .MoveNext
             Loop
             On Error GoTo 0
           End With
           '
           Call SETULTREG("LIPEDPE1", JJ&)
           Call CIERRARCH("LIPEDPE1")
           Call FINAL("*LIPENC1")
99  Screen.MousePointer = 1
    Command6.Enabled = True

90 End Sub

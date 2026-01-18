VERSION 5.00
Begin VB.Form FAPOCAB04 
   BackColor       =   &H00C0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Ordenes de Compra Abiertas"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Opciones de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   5
      Top             =   105
      Width           =   4320
      Begin VB.PictureBox Picture4 
         Height          =   690
         Left            =   1155
         ScaleHeight     =   630
         ScaleWidth      =   2580
         TabIndex        =   8
         Top             =   1575
         Width           =   2640
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
            Picture         =   "FAPOCAB04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label3 
            Caption         =   "Listado de O/C's Pendientes de Aprobación"
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
            Top             =   0
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   690
         Left            =   840
         ScaleHeight     =   630
         ScaleWidth      =   2580
         TabIndex        =   7
         Top             =   990
         Width           =   2640
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
            Picture         =   "FAPOCAB04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ver por Pantalla Orden de Compra Pendiente de Aprobación."
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label2 
            Caption         =   "Ver O/C's Pendientes de  Aprobación"
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
            TabIndex        =   13
            Top             =   60
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   690
         Left            =   525
         ScaleHeight     =   630
         ScaleWidth      =   2580
         TabIndex        =   6
         Top             =   420
         Width           =   2640
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
            Picture         =   "FAPOCAB04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Aprueba Orden de Compra"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label1 
            Caption         =   "Aprobar Ordenes de Compra"
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
            Top             =   60
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   3585
      Left            =   4725
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
         Picture         =   "FAPOCAB04.frx":091E
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
         Picture         =   "FAPOCAB04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FAPOCAB04.frx":0D72
         Top             =   2310
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FAPOCAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Command14.Enabled = False
     '
     If DERACCE%("FOAPOCOM", "Aprobación de Ordenes de Compra") < 2 Then
        Exit Sub
     Else
         If COMALTER%("Realmente Desea\APROBAR Ordenes de Compra ?\\No, Abandonar\Si, Aprobar O/C's") <> 2 Then
           Command14.Enabled = True
           Exit Sub
         End If
         '
         HO% = HOY(HOS$)
         LU$ = LUDAT$
         '
          TITU$ = "PROCESO DE APROBACIÓN DE ORDENES DE COMPRA"
        
         ' Selecciono todas las O/C's pendientes de aprobación.
         '
         Call ABRECOMPRAS
         SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
         SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
         SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
         Set VEROCTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
         With VEROCTEMP
           On Error Resume Next
            .MoveFirst
            If Err Then
              Call MENSERR(24, "No Existen Ordenes de Compra Pendientes de Aprobación.")
              On Error GoTo 0
              GoTo 99
            End If
           On Error GoTo 0
            '
            JJ& = 0
            Call COPYSTRU("INDIOCOM", "INDOCPEN")
            Do While (.EOF = False)
                XXY$ = REGBLAN$("INDOCPEN")
                Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
                Call REPLA(XXY$, !Refe_OcoM, 5, 44)
                JJ& = JJ& + 1
                Call GRAREG("INDOCPEN", XXY$, JJ&)
                .MoveNext
            Loop
         End With
        '
        Set VEROCTEMP = Nothing
        Call SETULTREG("INDOCPEN", JJ&)
        Call CIERRARCH("INDOCPEN")
        Call FRESHECHO("INDOCPEN")
        '
        '
        Call FRESHECHO("INDOCPEN")
3       Call SELECHO("INDOCPEN/Indice General de Ordenes de Compra Pendientes de Aprobación.")
        NPX& = Val(VALACT1$("INDOCPEN"))
        If NPX& < 1 Then GoTo 99
        '////***NUEVO PARA Q NO SE PUEDAN APROBAR O/C SIN PRECIO****////
        Call ABRECOMPRAS
        SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
        SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
        SQLX$ = SQLX$ + " AND Total_Ocom = 0 "
        SQLX$ = SQLX$ + " AND Nume_Ocom = " & NPX& & " "
        Set VEROCTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With VEROCTEMP
          On Error Resume Next
          .MoveFirst
          If Err Then
             GoTo 10
          Else
             'Call MENSERR(24, "No se Puede Aprobar la Orden de Compra Nº " & NPX& & "\ Porque no Tiene Precio.")
             TTXX$ = "La Orden de Compra Nº " & NPX& & "\es SIN PRECIOS !!!\¿ Está Seguro de Aprobarla ?"
             If COMALTER%(TTXX$ + "\\No, Cancelar Aprobación\Si, Aprobarla") = 2 Then GoTo 10
             On Error GoTo 0
             GoTo 3
          End If
        End With
        '
        '/////****FIN******//////
10      On Error GoTo 0
        Call ABRECOMPRAS
        With EncoCabi
            .FindFirst "Nume_Ocom = " & NPX & ""
            If .NoMatch = False Then
              If !Stat_Ocom < 0 Then ' SOLO PENDIENTES DE APROBACION
               .Edit
               !Stat_Ocom = 0
               !Usua_Apro = USNBR% Mod 100
               !FApro_Ocom = Now
               .Update
              End If
            End If
        End With
        '
        PRX% = 0
        With DetoCabi
4           If PRX% = 0 Then
              .FindFirst "Nume_Ocom = " & NPX & ""
            Else
              .FindNext "Nume_Ocom = " & NPX & ""
            End If
            If .NoMatch = False Then
              If !Stat_Ocom < 0 Then    ' SOLO PENDIENTES DE APROBACION
                .Edit
                !Stat_Ocom = 0
                .Update
              End If
              PRX% = 1
              GoTo 4
            End If
        End With
        '
        'EncoCabi.Close
        'DetoCabi.Close
        Call COMUNI("Se ha Aprobado con Éxito\la Orden de Compra Nº " & NPX & "")
        '
    End If
99  Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command3_Click()
     Call MENSERR(24, "No Disponible por el Momento.")
End Sub

Private Sub Command4_Click()
     Command4.Enabled = False
     'Selecciono todas las O/C's pendientes de aprobación.
     Call ABRECOMPRAS
     SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
     SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
     SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
     Set VEROCTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With VEROCTEMP
       On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No Existen Ordenes de Compra Pendientes de Aprobación.")
          On Error GoTo 0
          GoTo 99
        End If
       On Error GoTo 0
        '
        JJ& = 0
        Call COPYSTRU("INDIOCOM", "INDOCPEN")
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDOCPEN")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("INDOCPEN", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    '
    Set VEROCTEMP = Nothing
    Call SETULTREG("INDOCPEN", JJ&)
    Call CIERRARCH("INDOCPEN")
    '
    '
    Call FRESHECHO("INDOCPEN")
3   Call SELECHO("INDOCPEN/Indice General de Ordenes de Compra Pendientes de Aprobación.")
    NPX& = Val(VALACT1$("INDOCPEN"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido - Abierta"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
      '
      '
99  Command4.Enabled = True
   End Sub


Private Sub Command6_Click()
Screen.MousePointer = 11
   'Listado de O/C'S PENDIENTES DE APROBACIÓN.
   Call ABRECOMPRAS
   SQLX$ = " Select * FROM DETOCABI "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
   SQLX$ = SQLX$ + " Order by Nume_Ocom ASC "
   Set OCOMPETEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   With OCOMPETEMP
     On Error Resume Next
     .MoveLast
     If Err Then
       Call MENSERR(24, "No Existen O/C's Pendientes de Aprobación")
       On Error GoTo 0
       GoTo 90
     End If
     On Error GoTo 0
     '
     FIN& = .RecordCount
     .MoveFirst
     JJ& = 0
     Call COPYSTRU("OPENDETA", "OPENDTEX")
     Do While (.EOF = False)
       XXX$ = REGBLAN$("OPENDTEX")
       XXY$ = XXX$ + String$(32, 0)
         MONEMI% = !Mone_Emis
       If MONEMI% <= 1 Then MONEX$ = "   "
       If MONEMI% = 2 Then MONEX$ = "U$S "
       If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
       TICAMBIO = 1
       If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
       '
       Call REPLA(XXY$, MONEX$, 121, 2)
       CAPEN = !Cant_Ocom - !Cant_Rec
       If CAPEN < 0.05 Then CAPEN = 0
       PREUNID# = !PrUn_Neto * TICAMBIO
       '
       Call REPLA(XXY$, MKS$(!Nume_Ocom), 1, 4)
       Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
       Call REPLA(XXY$, MKI$(!Npro_Ocom), 7, 2)
       Call REPLA(XXY$, MKI$(!Cod_inte), 9, 2)
       Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
       Call REPLA(XXY$, !Uni_Com, 59, 4)
       Call REPLA(XXY$, MKD$(!PrUn_Neto), 79, 8)
       Call REPLA(XXY$, MKD$(!Cant_Ocom * !PrUn_Neto), 87, 8)
       Call REPLA(XXY$, MKI$(CVINT%(!FVig_Ocom)), 95, 2)
       Call REPLA(XXY$, !CONSIMAT, 97, 2)
       Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
       Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
       Call REPLA(XXY$, MKI$(!Stat_Ocom), 124, 1)
       Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
       Call REPLA(XXY$, MKD$(PREUNID#), 133, 8)
       Call REPLA(XXY$, MKD$(PREUNID# * CAPEN), 141, 8)
       JJ& = JJ& + 1
       Call GRAREG("OPENDTEX", XXY$, JJ&)
     .MoveNext
     Loop
   End With
   '
   Call SETULTREG("OPENDTEX", JJ&)
   Call CIERRARCH("OPENDTEX")
   Call FINAL("*LOCAPEX")
   '
90 Screen.MousePointer = 1
   '
End Sub

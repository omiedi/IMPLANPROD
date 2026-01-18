VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FOAPROC07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "APROBACION DE ORDENES DE COMPRA"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   4560
      ScaleHeight     =   3555
      ScaleWidth      =   1350
      TabIndex        =   13
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   16
         Top             =   2160
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FOAPROC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FOAPROC07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3480
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPCIONES DE TRABAJO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4320
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   1155
         ScaleHeight     =   600
         ScaleWidth      =   2610
         TabIndex        =   3
         Top             =   1575
         Width           =   2640
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "FOAPROC07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Listado de O/C's Pen- dientes de Aprobación"
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
            TabIndex        =   9
            Top             =   90
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   840
         ScaleHeight     =   600
         ScaleWidth      =   2610
         TabIndex        =   2
         Top             =   990
         Width           =   2640
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "FOAPROC07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Ver por Pantalla Orden de Compra Pendiente de Aprobación."
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver O/C's Pendientes de  Aprobación"
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
            TabIndex        =   8
            Top             =   90
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   525
         ScaleHeight     =   600
         ScaleWidth      =   2730
         TabIndex        =   1
         Top             =   420
         Width           =   2760
         Begin VB.CommandButton Command21 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Detalle MRP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   2065
            Picture         =   "FOAPROC07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Aprobación Con Detalle de Neciesidades de MRP"
            Top             =   0
            Width           =   670
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "FOAPROC07.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Aprueba Orden de Compra"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprobar Orde- nes de Compra"
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
            Left            =   840
            TabIndex        =   7
            Top             =   90
            Width           =   1350
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FOAPROC07.frx":0EBC
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FOAPROC07.frx":10F0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FOAPROC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim VEROCTEMP As ADODB.Recordset
 Dim OCOMPETEMP As ADODB.Recordset
  
    
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
        Call APERBASDAT("COMPRAS")
        SQLX$ = " SELECT Nume_Ocom, Refe_Ocom, Tobru_Ocom FROM OCOMENCA"
        SQLX$ = SQLX$ + " WHERE Stat_Ocom = -1 " 'SE CAMBIA < 0 POR = -1 PARA QUE NO MUESTRE LAS QUE ESTAN EN -2 QUE SON PENDIENTES
        SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "    'POR QUE NO PASAN LA VALIDACION DE STOCKS O PEDIDOS SEGUN ARCON
        SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
        'Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
         
        Set VEROCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
            SELOCOPEN.LICOPEN.Clear
            jj& = 0
            'Call COPYSTRU("INDIOCOM", "INDOCPEN")
            Do While (.EOF = False)
                XXY$ = Space$(64)
                Call REPLA(XXY$, FORMATNUM$(!nume_ocom, "F6.0"), 1, 6)
                Call REPLA(XXY$, !Refe_OcoM, 9, 44)
                Call REPLA(XXY$, FORMATNUM(!ToBru_Ocom, "F12.2"), 53, 12)
                jj& = jj& + 1
                'Call GRAREG("INDOCPEN", XXY$, JJ&)
                SELOCOPEN.LICOPEN.AddItem XXY$
            .MoveNext
            Loop
        End With
        '
        VEROCTEMP.Close
        Set VEROCTEMP = Nothing
3       SELOCOPEN.Show 1
          '
        If XVALO(SELOCOPEN.Label5) < 1 Then GoTo 99 ' CANTIDAD DE REGISTROS APROBADOS
        If SELOCOPEN.Label2 = "OK" Then
          If COMALTER%("¿ Confirma la Aprobación\de las Ordenes de Compra Tildadas ?\\No, Abandonar\Si, Aprobarlas") = 2 Then
             Call ABRECOMPRAS
             Screen.MousePointer = 11
             FIN& = SELOCOPEN.LICOPEN.ListCount
             For ILIS& = 1 To FIN&
                Call TRACE(20, ILIS&, FIN&)
                If SELOCOPEN.LICOPEN.Selected(ILIS& - 1) = True Then
                   XXY$ = SELOCOPEN.LICOPEN.List(ILIS& - 1)
                   NPX& = XVALO(Left$(XXY$, 6))
                    If NPX& < 1 Then GoTo 19

                    '////***NUEVO PARA Q NO SE PUEDAN APROBAR O/C SIN PRECIO****////
                    Call APERBASDAT("COMPRAS")
                    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
                    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
                    SQLX$ = SQLX$ + " AND Total_Ocom = 0 "
                    SQLX$ = SQLX$ + " AND Nume_Ocom = " & NPX& & " "
                    'Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
                    '
                    Set VEROCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
                    VEROCTEMP.Close
                    Set VEROCTEMP = Nothing
                    '/////****FIN******//////
            
10                  On Error GoTo 0
                    Call APERBASDAT("COMPRAS")
             
                    STRSQL = "Update OCOMENCA"
                    STRSQL = STRSQL + " set Stat_ocom=0"
                    STRSQL = STRSQL + ", Usua_Apro= " & USNBR% Mod 100
                    STRSQL = STRSQL + ", FApro_Ocom= GETDATE()"
                    STRSQL = STRSQL + " where Stat_Ocom < 0 "
                    STRSQL = STRSQL + " and Nume_Ocom = " & NPX
                     FLEXCONN.Execute (FILTSQL$(STRSQL))
                    '
                    STRSQL = "Update OCOMDETA"
                    STRSQL = STRSQL + " set Stat_ocom=0"
                    STRSQL = STRSQL + " where Nume_Ocom=  " & NPX
                    STRSQL = STRSQL + " and Stat_Ocom < 0"
                     FLEXCONN.Execute (FILTSQL$(STRSQL))
                    '
                    '\\\OT-08-01047-FL-21/11/08///Esto es para encontrar el/los documentos seleccionados e imprimirlos
                    If CONTROL$("OCOMPRA", "IMPRIME") = "ALAPROBA" Then
                         TIDOCUM$ = "Nota de Pedido"
    
                         VUELTA% = 1
                         LOMOSTRO% = 0
                         '
15                       NUDOCU$ = TRIM$(Str$(NPX&))
                         While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
                         NUDOCU$ = NUDOCU$ + ""
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
                               DOCUNU& = CVS(Left$(XX$, 4))
                               GoTo 5
                             End If
                           End If
                         Next UI&
                         Screen.MousePointer = 1
5                       Call MUESTRADOC(DOCUNU&, "PRINT")
                    End If
                '\\\OT-08-01047-FL-FIN///
                End If
19            Next ILIS&
             Call COMUNI("Se han Aprobado con Éxito\las Ordenes de Compra Seleccionadas.")
          End If
        End If
        'Call COMUNI("Se ha Aprobado con Éxito\la Orden de Compra Nº " & NPX & "")
        '
    End If
99  Screen.MousePointer = 1
    Command14.Enabled = True

End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command21_Click()
        Command21.Enabled = False
        Call GENERBASE("STOCKS")
        If DERACCE%("FOAPOCOM", "Aprobación de Ordenes de Compra") < 2 Then GoTo 99
        Call APERBASDAT("COMPRAS")
40      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom, Tobru_Ocom FROM OCOMENCA with(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE Stat_Ocom = -1 " 'SE CAMBIA < 0 POR = -1 PARA QUE NO MUESTRE LAS QUE ESTAN EN -2 QUE SON PENDIENTES
        SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "    'POR QUE NO PASAN LA VALIDACION DE STOCKS O PEDIDOS SEGUN ARCON
        SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
        Set RS = READSET(SQLX$)
        With RS
            If .EOF = True Then
              Call MENSERR(24, "No Existen Ordenes de Compra Pendientes de Aprobación.")
              On Error GoTo 0
              GoTo 99
            End If
            j& = 0
            
            Campos$ = "O.Compra/F8;Referencia/A64;Imp.Neto/F12.2"
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 11500)
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , 11500)
            FRMZELECHO.Width = 12000
            FRMZELECHO.Caption = "Ordenes de Compra Pendientes de Aprobación"
            '
            FRMZELECHO.MSF1.Redraw = False: FRMZELECHO.msf2.Redraw = False
            
            Do While Not .EOF
               NUME& = XVALO(!nume_ocom)
               REFE$ = SAFETEXT$(!Refe_OcoM)
               IMPO$ = FORMATNUM(XVALO(!ToBru_Ocom), "F12.2")
                j& = j& + 1
               FRMZELECHO.msf2.Rows = j& + 1
               FRMZELECHO.msf2.TextMatrix(j&, 1) = NUME&
               FRMZELECHO.msf2.TextMatrix(j&, 2) = REFE$
               FRMZELECHO.msf2.TextMatrix(j&, 3) = IMPO$
               .MoveNext
            Loop
        End With
        FRMZELECHO.MSF1.Redraw = True: FRMZELECHO.msf2.Redraw = True

        FRMZELECHO.Show 1
        NUMOC& = XVALO(RESP_ZELE_VECT(1))
        If NUMOC& < 1 Then GoTo 99
        REFE$ = RESP_ZELE_VECT(2)
        IMPO$ = RESP_ZELE_VECT(3)
        Unload FRMZELECHO
        If NUMOC& > 0 Then
            SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
            SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUMOC&
            SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
            SQLX$ = SQLX$ + "ORDER BY NuOrd_Item ASC"
            '
            Set RS = READSET(SQLX$)
            With RS
               If .EOF Then
                Call MENSERR(24, "Imposible cargar Detalle de O/Compra\Orden de Compra Anulada")
                Screen.MousePointer = 1
                GoTo 99
              End If
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Campos$ = "CI/F0;Código/A16;Descripción/A82;A Comprar/F10.1;U.Stock/A6;A Comprar/F10.1;U.Proveed./A10;Nec.Sem.Ac./f11;Nec.Restante/F11;Recibido/F0;Pendiente/f11"
              Call CARGA_ENCABEZADO(SITOCMRP.GRID, Campos$, SITOCMRP, 1)
              Call CARGA_ENCABEZADO(SITOCMRP.MSF_2, Campos$, SITOCMRP)
              Campos$ = "CI/F0;Código/A16;Descripción/A0;Semana/A18;Referencia/A20;Cant.Disp./F11;Cant.Comprom./F15;Saldo/f11"
              Call CARGA_ENCABEZADO(SITOCMRP.MSFDET, Campos$, SITOCMRP)

              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              FRMZELECHO.MSF1.Redraw = False: FRMZELECHO.msf2.Redraw = False
              FEULCA = FECHAULTIMO_CALCULO
              j& = 0
              Do While Not .EOF
                UNIDADOCOM = FORMATNUM$(XVALO(!UNIDS_OCOM), "F10.1")
                UNICOMP$ = SAFETEXT$(!Uni_Com)
                COEF = XVALO(!Coef_Unids)
                
                CIXI% = XVALO(!Cod_Inte)
                CI1 = CIXI%
                CAPEN = XVALO(!Cant_Ocom)
                UNIMA$ = TRIM$(SAFETEXT$(!Uni_Mas))
                NECESEM1 = FORMATNUM$(NECECOMPRA_MRP(CI1, 1), "F10.1")  'NECESIDAD DE COMPRA DE LA SEMANA ACTIVA SEGUN MRP
                SALDONECESIDAD = NECECOMPRA_MRP(CIXI%) - NECECOMPRA_MRP(CI1, 1): If SALDONECESIDAD < 0 Then SALDONECESIDAD = 0
                NECESEM2 = FORMATNUM$(SALDONECESIDAD, "F10.1") 'NECESIDAD DEL MRP - LA NECESIDAD DE LA SEMANA ACTIVA
                RECIBIDO = RECE_POST_FECHA(CIXI%, FEULCA)
                COMPENDACU = CompraPendientesAcumuladasPos(CIXI%, NUMOC&, 0, FEULCA, 1)
                CANRECIPOST = CANRECEP(CIXI%, FEULCA)
                PENDIENTEX = COMPENDACU + CANRECIPOST
                
                With SITOCMRP.GRID
                    j& = j& + 1
                    SITOCMRP.GRID.Rows = j& + 1
                    .Rows = j& + 1
                    .TextMatrix(j&, 1) = CIXI%
                    .TextMatrix(j&, 2) = CODEXT$(CIXI%)
                    .TextMatrix(j&, 3) = DESCRIT0$(CIXI%)
                    .TextMatrix(j&, 4) = UNIDADOCOM
                    .TextMatrix(j&, 5) = UNIMA$
                    .TextMatrix(j&, 6) = FORMATNUM$(XVALO(CAPEN), "F10.1")
                    .TextMatrix(j&, 7) = UNICOMP$
                    .TextMatrix(j&, 8) = FORMATNUM$(NECESEM1, "F10.1")
                    .TextMatrix(j&, 9) = FORMATNUM$(NECESEM2, "F10.1")
                    .TextMatrix(j&, 10) = FORMATNUM$(RECIBIDO, "F10.1")
                    .TextMatrix(j&, 11) = FORMATNUM$(PENDIENTEX, "F10.1")
                End With
               .MoveNext
              Loop
              RS.Close
              Set RS = Nothing
              If j& > 0 Then
                SITOCMRP.Caption = "Detalle de Orden de Compra N°: " & NUMOC&
                SITOCMRP.Label5.Caption = NUMOC&
                SITOCMRP.Show 1
                GoTo 99
                
                 If VENTABUX.APROBADO% > 0 Then
                    If COMALTER%("Realmente Desea\APROBAR Orden de Compra N° " & NUMOC& & "?\\No, Abandonar\Si, Aprobar O/C's") = 2 Then
                        STRSQL = "Update OCOMENCA"
                        STRSQL = STRSQL + " set Stat_ocom=0"
                        STRSQL = STRSQL + ", Usua_Apro= " & USNBR% Mod 100
                        STRSQL = STRSQL + ", FApro_Ocom= GETDATE()"
                        STRSQL = STRSQL + " where Stat_Ocom < 0 "
                        STRSQL = STRSQL + " and Nume_Ocom = " & NUMOC&
                         FLEXCONN.Execute (FILTSQL$(STRSQL))
                        '
                        STRSQL = "Update OCOMDETA"
                        STRSQL = STRSQL + " set Stat_ocom=0"
                        STRSQL = STRSQL + " where Nume_Ocom=  " & NUMOC&
                        STRSQL = STRSQL + " and Stat_Ocom < 0"
                        FLEXCONN.Execute (FILTSQL$(STRSQL))
                        Call IMPRIME_OCOMPRA(NUMOC&)
                        GoTo 40
                     End If
                 Else
                     GoTo 40
                 End If
              End If
           End With
        End If
  
99    Command21.Enabled = True

End Sub

Private Sub Command21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "Presenta Ventana con Ordenes de Compra Pendientes de Aprobación" + vbCrLf
       TXTO$ = TXTO$ + "Dada una O.Compra Seleccionada se Visualiza el Detalle de la Misma " & vbCrLf
       TXTO$ = TXTO$ + "Visualizando el siguiente detalle:" & vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf

       TXTO$ = TXTO$ + "Cantidad = Cantidad pedida en la oreden de compra seleccionada" & vbCrLf
       TXTO$ = TXTO$ + "Nec.Sem.Ac = Acumulado de Necesidad de semana activa hacia atras segun ultimo calculo mrp" + vbCrLf
       TXTO$ = TXTO$ + "Nec.Restante. = Acumulado a partir de donde termina la Semana Activa  en adelante segun ultimo calculo mrp " + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "Recibido = Material Recibido del codigo posterior al ultimo calculo" + vbCrLf
       TXTO$ = TXTO$ + "Pendiente = Compras Pendientes de Recepcion + Recepciones pendientes de Aprobacion" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "Si se Aprueba dicha Ventana se Aprueba la orden de Compra." + vbCrLf
       TXTO$ = TXTO$ + "Caso Contrario no se Realiza Ninguna Transacción" + vbCrLf
       MsgBox TXTO$
     End If
   End If


End Sub


Private Sub Command3_Click()
     Call MENSERR(24, "No Disponible por el Momento.")
End Sub

Private Sub Command4_Click()
     Command4.Enabled = False
     'Selecciono todas las O/C's pendientes de aprobación.
     Call APERBASDAT("COMPRAS")
     SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
     SQLX$ = SQLX$ + " WHERE Stat_Ocom = -1 " 'SE CAMBIA < 0 POR = -1 PARA QUE NO MUESTRE LAS QUE ESTAN EN -2 QUE SON PENDIENTES
     SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "    'POR QUE NO PASAN LA VALIDACION DE STOCKS O PEDIDOS SEGUN ARCON
     SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
     'Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Set VEROCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
        jj& = 0
        Call COPYSTRU("INDIOCOM", "INDOCPEN")
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDOCPEN")
            Call REPLA(XXY$, MKS(!nume_ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            jj& = jj& + 1
            Call GRAREG("INDOCPEN", XXY$, jj&)
        .MoveNext
        Loop
    End With
    '
    VEROCTEMP.Close
    Set VEROCTEMP = Nothing
    Call SETULTREG("INDOCPEN", jj&)
    Call CIERRARCH("INDOCPEN")
      '
      '
3   Call SELECHO("INDOCPEN/Indice General de Ordenes de Compra Pendientes de Aprobación.")
    NPX& = XVALO(VALACT1$("INDOCPEN"))
    If NPX& < 1 Then GoTo 99
    '
    DOCUNU& = 0
    TIDOCUM$ = "Nota de Pedido"
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "ORDEN DE COMPRA"

    NUDOCU$ = TRIM$(Str$(NPX&))
'    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    
    
    NoPuedeImprimir% = 1
    MENSAJEDOC$ = "Imposible Imprimir Esta Orden de Compra, La Misma Está Pendiente de Aprobación"
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, , , MENSAJEDOC$, NoPuedeImprimir%)
    NoPuedeImprimir% = 0
    MENSAJEDOC$ = ""
    Screen.MousePointer = 1

    GoTo 3

'    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'    NUDOCU1$ = " " + NUDOCU$ + " "
'    NUDOCU2$ = " " + NUDOCU$ + "-"
'    Screen.MousePointer = 11
'    '
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        PPXX% = InStr(XX$, "Nro.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'          DOCUNU& = CVS(Left$(XX$, 4))
'          GoTo 5
'        End If
'      End If
'    Next UI&
'    Screen.MousePointer = 1
'    Call COMUNI("Documento no Encontrado")
'    GoTo 99
'    '
'5   Call MUESTRADOC(DOCUNU&)
'    '
'    GoTo 3
    '
99  Command4.Enabled = True

End Sub

Private Sub Command6_Click()
   '
   Screen.MousePointer = 11
   'Listado de O/C'S PENDIENTES DE APROBACIÓN.
   Call APERBASDAT("COMPRAS")
   SQLX$ = " Select * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
   '   Set OCOMPETEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OCOMPETEMP As ADODB.Recordset
   Set OCOMPETEMP = New ADODB.Recordset
   OCOMPETEMP.CursorLocation = adUseClient
   OCOMPETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   
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
     jj& = 0
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
       PreUnid# = !prun_neto * TICAMBIO
       '
       Call REPLA(XXY$, MKS$(!nume_ocom), 1, 4)
       Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
       Call REPLA(XXY$, MKI$(!Npro_Ocom), 7, 2)
       Call REPLA(XXY$, MKI$(!Cod_Inte), 9, 2)
       Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
       Call REPLA(XXY$, !Uni_Com, 59, 4)
       Call REPLA(XXY$, MKD$(!prun_neto), 79, 8)
       Call REPLA(XXY$, MKD$(!Cant_Ocom * !prun_neto), 87, 8)
       Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Sol)), 95, 2)
       Call REPLA(XXY$, !CONSIMAT, 97, 2)
       Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
       Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
       Call REPLA(XXY$, MKI$(!Stat_Ocom), 124, 1)
       Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
       Call REPLA(XXY$, MKD$(PreUnid#), 133, 8)
       Call REPLA(XXY$, MKD$(PreUnid# * CAPEN), 141, 8)
       jj& = jj& + 1
       Call GRAREG("OPENDTEX", XXY$, jj&)
     .MoveNext
     Loop
   End With
   OCOMPETEMP.Close
   Set OCOMPETEMP = Nothing
   '
   
   Call SETULTREG("OPENDTEX", jj&)
   Call CIERRARCH("OPENDTEX")
   Call FINAL("*LICOPEX")
   '
90 Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
End Sub


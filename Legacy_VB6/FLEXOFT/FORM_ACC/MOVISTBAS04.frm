VERSION 5.00
Begin VB.Form MOVISTBAS04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5850
   Icon            =   "MOVISTBAS04.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   3900
      Left            =   4935
      ScaleHeight     =   3840
      ScaleWidth      =   1320
      TabIndex        =   12
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "MOVISTBAS04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Height          =   690
         Left            =   105
         Picture         =   "MOVISTBAS04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   2500
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MOVISTBAS04.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3050
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Movimientos de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Movimientos Varios"
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
         Left            =   210
         TabIndex        =   3
         Top             =   1900
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   8
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command3 
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
               Picture         =   "MOVISTBAS04.frx":2680
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes de Stock"
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
               TabIndex        =   11
               Top             =   105
               Width           =   795
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   6
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command8 
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
               Picture         =   "MOVISTBAS04.frx":27CA
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Remito de Traslado"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Despacho Mercaderia"
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
               Left            =   630
               TabIndex        =   10
               Top             =   105
               Width           =   1320
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Entradas a Depósito"
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
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   17
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command2 
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
               Picture         =   "MOVISTBAS04.frx":2AD4
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Producción"
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
               Left            =   630
               TabIndex        =   19
               Top             =   105
               Width           =   1005
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   4
            Top             =   400
            Width           =   1695
            Begin VB.CommandButton Command28 
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
               Picture         =   "MOVISTBAS04.frx":2DDE
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción  Materiales"
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
               Left            =   630
               TabIndex        =   5
               Top             =   105
               Width           =   1320
            End
         End
      End
   End
End
Attribute VB_Name = "MOVISTBAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
    '05/04/07 (OT-07-0150)
    Call RECPROP
    '05/04/07 (FIN)
End Sub

Private Sub Command28_Click()
    Command28.Visible = False
    Call CONECRUN("BRECEP04", "Recepción de Materiales de Stock")
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   AJUSTOBAS04.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("REMITO04", "Remito de Expedición")
    Command8.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label5_Click()
    LU$ = LUDAT$
    If EXISTE%(LU$ + "GETRAMO.DAT") > 0 Then
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("GETRAMO")
            X$ = REGLEIDO$("GETRAMO", U&)
            COD$ = Left$(X$, 16)
            CI% = CODINT%(COD$)
            Y$ = Mid$(X$, 17, 128)
            Call REPLA(Y$, MKI$(CI%), 3, 2)
            Call REGAPP("MOVISTO", Y$)
        Next U&
        Call CIERRARCH("*.*")
        Call FILEKILL(LU$ + "GETRAMO.DAT")
        Screen.MousePointer = 1
    End If
End Sub

Sub RECECOMO() ' recepcion de materiales Comodoro
        '
   'Call BLIMESS("Seleccionando ...")
   FIN& = ULTREG&("ATRASLA")
   Call BLANARCH("!ATRASLA")
   CABOPE% = 0
   FF% = HOY(FECHAREC$)
   '
   For I& = 1 To FIN&
     X$ = REGLEIDO$("ATRASLA", I&)
     If CVD(Mid$(X$, 105, 8)) >= 0.05 Then
       If CVD(Mid$(X$, 117, 8)) < 0.05 Then
         CANOM# = CVD(Mid$(X$, 105, 8))
         Call REPLA(X$, MKD$(CANOM#), 117, 8)
         Call REPLA(X$, MKS$(I&), 125, 4) ' registro
         Call REGAPP("!ATRASLA", X$)
       End If
     End If
   Next I&
   '
   If ULTREG&("!ATRASLA") < 1 Then
     Call COMUNI("No hay Recepciones\Pendientes de Ingreso.")
     Exit Sub
   End If
   '
   FERECEP$ = REPLACAR$(FECHAREC$, "/", "-")
   VTB% = VENTABU%("RECECOM/NC/TITUPAN=Boleta de Recepción - " + FERECEP$)
   '
   If VTB% >= 0 Then
     If COMALTER%("Confirme Ingreso de Materiales Recibidos\\Cancelar\Confirmar Ingreso") = 2 Then
       AINDI% = 0
       For U& = 1 To ULTREG&("!ATRASLA")
         X$ = REGLEIDO$("!ATRASLA", U&)
         APRO# = CVD(Mid$(X$, 117, 8))
         If APRO# > 0.05 Then
           Call REPLA(X$, MKI$(FF%), 115, 2)
           REGORI& = CVS(Mid$(X$, 125, 4))
           If REGORI& > 0 Then
             Y$ = REGLEIDO$("ATRASLA", REGORI&)
           End If
           If REGORI& < 1 Or Left$(Y$, 112) <> Left$(X$, 112) Then
               Call REGAPP("ATRASLA", X$)
             Else
               Call GRAREG("ATRASLA", X$, REGORI&)
           End If
           '
           CI% = CVI(Mid$(X$, 39, 2))
           REFE$ = "Recepción " + FECHAREC$
           AINDI% = AINDI% + 1
           LOTE$ = TRIM$(Mid$(X$, 3, 14)) + "-" + TRIM$(Str$(AINDI%))
           DOPRI$ = Mid$(X$, 3, 14)
           DOSEC$ = "BR." + FECHAREC$
           DEPX% = 1
           CANTI = APRO#
           Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
         End If
       Next U&
     End If
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'
Sub RECPROP()
    '
    '05/04/07 (OT-07-0150)
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If CONTROL$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    '
100 Screen.MousePointer = 11
    For U& = ULTREG&("RECPROP") To 1 Step -1
        XX$ = REGLEIDO$("RECPROP", U&)
        If CVS(Mid$(XX$, 59, 4)) > 0 Then
            UPARTE& = CVS(Mid$(XX$, 59, 4))
            GoTo 105
        End If
    Next U&
    UPARTE& = 0
    '
105 Screen.MousePointer = 1
    Call BLANARCH("!CAREPRO")
    '
    VDEF$ = String$(88, 0)
    Call REPLA(VDEF$, MKI$(HOOI%), 1, 2)
    Call REPLA(VDEF$, MKS$(UPARTE& + 1), 85, 4)
    Call REPLA(VDEF$, Chr$(0), 37, 1)
    Call REPLA(VDEF$, Chr$(0), 38, 1)
    '
110 XX$ = OBJPLA$("RECPROP", VDEF$)
    If XX$ = "" Then Exit Sub
    '
160 VDEF$ = XX$
    FECHA% = CVI(Left$(XX$, 2))
    If FECHA% > HOOI% Or FECHA% < 1 Then
        Call MENSERR(24, "Fecha Incorrecta")
        GoTo 110
    End If
    '
'    CI% = CVI(Mid$(XX$, 3, 2))
'    If TRIM$(CODEXT$(CI%)) = "" Then
'        Call MENSERR(24, "Codigo Inexistente.")
'        GoTo 110
'    End If
'    CODCON$ = CODEXT$(CI%)
'    '
    DP1% = Asc(Mid$(XX$, 37, 1))
    DP2% = Asc(Mid$(XX$, 38, 1))
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP1%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 110
    End If
    '
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 110
    End If
    '
120 U3& = 0
    For U1& = 1 To ULTREG&("!CAREPRO")
      XY$ = REGLEIDO$("!CAREPRO", U1&)
      CANTI = CVD(Mid$(XY$, 23, 8))
      If Abs(CANTI) >= 0.05 Then
        U3& = U3& + 1
        Call GRAREG("!CAREPRO", XY$, U3&)
      End If
    Next U1&
    Call SETULTREG("!CAREPRO", U3&)
    '
    VTB% = VENTABU%("CAREPRO")
    If VTB% < 0 Then GoTo 110
    '
    CARE& = ULTREG&("!CAREPRO")
    If CARE& < 1 Then GoTo 120
    '
    NUORIT% = 0
    For U& = 1 To CARE&
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CANTI = CVD(Mid$(XY$, 23, 8))
       SCRAP = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
       '
       '\\\OT-06-0165-WAR-08/05/06///
       'If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
       If SCRAP < 0 Then
       '\\\OT-06-0165-WAR-FIN///
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + SCRAP)
       VACONTROL% = 1
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
       '
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
         For YX& = 1 To ULTREG&("!MATEROF")
            YY$ = REGLEIDO$("!MATEROF", YX&)
            CI% = CVI(Mid$(YY$, 1, 2))
            CAN# = CVS(Mid$(YY$, 5, 4))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
         Next YX&
       If PUEDEMOVI%(DP2%) < 0 Then GoTo 110 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       NRO& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = NRO&
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       If CANTI < (-0.005) Then
          REFE$ = "Desarmado de Conjunto"
       End If
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       MONEII% = 0
       CMOV$ = "PF"
       '
       NUORIT% = 0
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
       Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(NRO&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
       Call REGAPP("RECPROP", YY$)
    
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
       For U2& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(ZZ$, 2))
          CAN = -1 * CVS(Mid$(ZZ$, 5, 4))
          '
          If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
                   REFE$ = "CONSUMO S/DESP." + CODCON$
                   '\\\OT-06-0165-WAR-08/05/06///
                   If CAN > 0 Then
                    REFE$ = "DESARMADO CONJ." + CODCON$
                   End If
                   '\\\OT-06-0165-WAR-FIN///
                   DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                   '
                   If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                     If InStr(EPAC$, "YALTRE") > 0 Then
                       Call CARDEPOS
                       For KKII% = 1 To NUSU%
                         If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                           DEPX% = COSU%(KKII%)
                           GoTo 124
                         End If
                       Next KKII%
                     End If
                   End If
                   '
124                NUORIT% = NUORIT% + 1
                   Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
       Call CIERRARCH("!MATEROF")
       '
       XY$ = REGLEIDO$("!CAREPRO", U&)
       Call REPLA(XY$, MKD$(0), 23, 8)
       Call GRAREG("!CAREPRO", XY$, U&)
       '
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '05/04/07 (FIN)
    '
End Sub


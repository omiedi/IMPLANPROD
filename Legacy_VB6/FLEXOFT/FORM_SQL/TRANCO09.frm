VERSION 5.00
Begin VB.Form TRANCO09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Sucursal"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5955
   Icon            =   "TRANCO09.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6180
      Left            =   4920
      ScaleHeight     =   6120
      ScaleWidth      =   1245
      TabIndex        =   17
      Top             =   0
      Width           =   1305
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   645
         TabIndex        =   21
         Top             =   4080
         Width           =   705
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   660
         Left            =   150
         Picture         =   "TRANCO09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   700
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
         Height          =   660
         Left            =   150
         Picture         =   "TRANCO09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   700
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ac-Dir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   150
         Picture         =   "TRANCO09.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   700
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -225
         Picture         =   "TRANCO09.frx":0A68
         Top             =   4440
         Width           =   1515
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -255
         Picture         =   "TRANCO09.frx":298A
         Top             =   5115
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -435
         Picture         =   "TRANCO09.frx":48AC
         Top             =   7860
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Movimientos de Sucursales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4710
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Liquidacion de Entregas"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   4110
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   750
            ScaleHeight     =   585
            ScaleWidth      =   2580
            TabIndex        =   10
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command5 
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
               Left            =   2035
               Picture         =   "TRANCO09.frx":7E96
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "TRANCO09.frx":81A0
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Control Interactivo de Liquidacion"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Control y Liqui- dación Entregas"
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
               Left            =   600
               TabIndex        =   12
               Top             =   105
               Width           =   1440
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Consultas y Listados"
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
         Top             =   3240
         Width           =   4110
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   240
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   13
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
               Picture         =   "TRANCO09.frx":82EA
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Consulta por Liquidacion"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Control de Liquidacion"
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
               Left            =   600
               TabIndex        =   14
               Top             =   120
               Width           =   1320
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   6
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
               Picture         =   "TRANCO09.frx":8434
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado Entregas Sucursal"
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
               Left            =   735
               TabIndex        =   8
               Top             =   0
               Width           =   795
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Transferencia Electrónica"
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
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   4
            Top             =   420
            Width           =   2115
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
               Picture         =   "TRANCO09.frx":873E
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción Transferencia de Sucursal"
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
               Left            =   735
               TabIndex        =   5
               Top             =   0
               Width           =   1320
            End
         End
      End
   End
End
Attribute VB_Name = "TRANCO09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*CODELIQ")
    Command2.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    Call RECIBETRA
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   '
'   ' DEPURA ENTRECOM
'   Call COPYSTRU("ENTRECOM", "ENTRECOY")
'   Call SETULTREG("ENTRECOY", 0)
'   '
'   JJ& = 0: FIN& = ULTREG&("ENTRECOM")
'   FF1% = FECHANUM("01/06/12")
'   For U& = 1 To FIN&
'       Call TRACE(20, U&, FIN&)
'       x$ = REGLEIDO$("ENTRECOM", U&)
'       FF% = CVI(Left$(x$, 2))
'       REMITO$ = TRIM$(Mid$(x$, 3, 14))
'       NC% = CVI(Mid$(33, 2))
'       '
'       If REMITO$ <> "" Then
'         CONDI$ = "DOPRICOR='" & REMITO$ & "' AND FECHMOV='" & FETEXCOR1$(FF%) & "' AND (DEPOMOVI=9 OR DEPOMOVI=10)"
'         If CantRegistros("MOVISTO", CONDI$) > 0 Then
'             JJ& = JJ& + 1
'             Call GRAREG("ENTRECOY", x$, JJ&)
''             If FF% >= FF1% Then
''                Call MsgBox("Importado" & vbCrLf & "fecha: " & FECHATEX$(FF%) & " - Remito: " & REMITO$ & " - Cliente: " & NC%)
''             End If
'         Else
'             'Call MsgBox("fecha: " & FECHATEX$(FF%) & " - Remito: " & REMITO$ & " - Cliente: " & NC%)
'         End If
'       End If
'   Next U&
'   Call SETULTREG("ENTRECOY", JJ&)
   '
10 Des% = COMESACT%
   Has% = HOY(HOS$)
   '
   
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
   VDEF$ = MKI$(Des%) + MKI$(Has%) + MKI$(NCMIN%) + MKI$(NCMAX%)
   OBX$ = OBJPLA$("SELRAFECHCLI", VDEF$)
   If OBX$ <> "" Then
     Des% = CVI(Left$(OBX$, 2))
     Has% = CVI(Mid$(OBX$, 3, 2))
     NCMIN% = CVI(Mid$(OBX$, 5, 2))
     NCMAX% = CVI(Mid$(OBX$, 7, 2))
     If Des% > Has% Then GoTo 10
     '
     JJ& = 0
     Screen.MousePointer = 11
     Call COPYSTRU("ENTRECOM", "ENTRECOX")
     FIN& = ULTREG&("ENTRECOM")
     'FIN& = ULTREG&("ENTRECOY")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ENTRECOM", U&)
       'x$ = REGLEIDO$("ENTRECOY", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= Des% Then
         If FF% <= Has% Then
           NCLX% = CVI(Mid$(X$, 33, 2))
           If NCLX% >= NCMIN% Then
             If NCLX% <= NCMAX% Then
               IDDOC& = CVS(Mid(X$, 97, 4))
               CI1% = CVI(Mid(X$, 117, 2))
               CONDI$ = "IDOCUM = " & IDDOC& & " AND COD_INTE = " & CI1%
               IDDOCUM& = XVALO(VALOBADA("ENTRECOM", "IDOCUM", CONDI$, "NOMESS"))
               If IDDOCUM& > 0 Then GoTo 40
               JJ& = JJ& + 1
               Call GRAREG("ENTRECOX", X$, JJ&)
             End If
           End If
         End If
       End If
40   Next U&
     Call SETULTREG("ENTRECOX", JJ&)
     Call FILESORT("ENTRECOX", "", 1, 2)
     '
     If JJ& < 1 Then
        Call COMUNI("Sin Movimientos")
      Else
        Call FINAL("*ENTRECOM")
     End If
     Screen.MousePointer = 1
   End If
   '
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call ACCDIR("CONSTOCK")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
   '\\\OT-06-0384-MH-11/09/06///
   HOII% = HOY(HOS$)
   Des% = 0
   Has% = HOII%
   '
10 OBX$ = OBJPLA$("SELECLI", VDEF$)
   '
   If OBX$ <> "" Then
     NC% = CVI(Left$(OBX$, 2))
     If NC% < 1 Then GoTo 10
     If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then GoTo 10
     Call DESHASFECHA(Des%, Has%, PERIODO$)
     '
     Screen.MousePointer = 11
     RFF$ = RECFILT$("ENTRECOM", 4, MKI$(NC%))
     FIN& = Len(RFF$)
     KK& = 0
     For U& = 1 To Len(RFF$) Step 4
       Call TRACE(20, U&, FIN&)
       RELE& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("ENTRECOM", RELE&)
       If CVI(Mid$(X$, 225, 2)) <= 8224 Or CVS(Mid$(X$, 227, 4)) < 0.05 Then
         If CVI(Mid(X$, 1, 2)) >= Des% Then
           If CVI(Mid(X$, 1, 2)) <= Has% Then
             TTXX$ = REGBLAN$("LIQUIENT")
              '
              FEC% = CVI(Mid$(X$, 1, 2))
             Call REPLA(TTXX$, MKI$(FEC%), 1, 2)
              REMI$ = Mid$(X$, 3, 14)
             Call REPLA(TTXX$, REMI$, 3, 14)
              MODO1$ = Mid$(X$, 17, 16)
             Call REPLA(TTXX$, MODO1$, 17, 16)
              ALMAC$ = Mid$(X$, 65, 32)
             Call REPLA(TTXX$, ALMAC$, 33, 32)
              CODII% = CVI(Mid$(X$, 117, 2))
             Call REPLA(TTXX$, MKI$(CODII%), 65, 2)
              CANTU# = CVD(Mid$(X$, 183, 8))
             Call REPLA(TTXX$, MKD$(CANTU#), 67, 8)
             '
             KK& = KK& + 1
             Call GRAREG("LIQUIENT", TTXX$, KK&)
             '
           End If
         End If
       End If
44   Next U&
     '
     Call SETULTREG("LIQUIENT", KK&)
     Call CIERRARCH("LIQUIENT")
     NUMCLI$ = FORMATNUM(NC%, "F5.0")
     NOMCLI$ = ECOARCH$("DEUDOR1", MKI$(NC%))
     Call HEADERS("LIQUIENT", "Cliente: " + NUMCLI$ + " - " + NOMCLI$)
     Call HEADERS("LIQUIENT", "Período: " + PERIODO$)
     Call FINAL("*LIQUIENT")
     Screen.MousePointer = 1
     '
   End If
   '
99 Command5.Enabled = True
'\\\OT-06-0384-MH-FIN///
End Sub



Private Sub Command8_Click()
   Command8.Enabled = False
'   Call FINAL("*ENTRECOV")
   HOII% = HOY(HOS$)
10 VDEF$ = MKI$(0) + MKI$(HOII%) + Space$(30)
   OBX$ = OBJPLA$("SELIQENTCLI", VDEF$)
   If OBX$ <> "" Then
     NC% = CVI(Left$(OBX$, 2))
     FF% = CVI(Mid$(OBX$, 3, 2))
     DELIQ$ = TRIM$(Mid$(OBX$, 5, 30))
     If NC% < 1 Then GoTo 10
     If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then GoTo 10
     If Len(DELIQ$) < 3 Then GoTo 10
     '
     DELIQU$ = UCase$(TRIM$(DELIQ$))
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("INDILIQ")
       X$ = REGLEIDO$("INDILIQ", U&)
       If UCase$(TRIM$(Mid$(X$, 3, 30))) = DELIQU$ Then
         Screen.MousePointer = 1
         Call MENSERR(24, "Liquidacion Ya Registrada")
         GoTo 10
       End If
     Next U&
     '
     Screen.MousePointer = 11
     LIQUIEN09.List2.Clear
     RFF$ = RECFILT$("ENTRECOM", 4, MKI$(NC%))
     FIN& = Len(RFF$)
     For U& = 1 To Len(RFF$) Step 4
       Call TRACE(20, U&, FIN&)
       RELE& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("ENTRECOM", RELE&)
       If CVI(Mid$(X$, 225, 2)) <= 8224 Or CVS(Mid$(X$, 227, 4)) < 0.05 Then
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, FECHATEX$(CVI(Mid$(X$, 1, 2))), 1, 8) ' FECHA REMITO
         Call REPLA(TTXX$, Mid$(X$, 3, 14), 11, 14)
         Call REPLA(TTXX$, Mid$(X$, 17, 5), 25, 5)
         Call REPLA(TTXX$, Mid$(X$, 65, 32), 32, 23)
         Call REPLA(TTXX$, Mid$(X$, 119, 16), 56, 16)
           CAN = CVD(Mid$(X$, 183, 8))
         Call REPLA(TTXX$, FORMATNUM$(CAN, "F10.1"), 72, 10)
         Call REPLA(TTXX$, FORMATNUM(RELE&, "F8.0"), 121, 8)
         Call REPLA(TTXX$, Mid$(X$, 35, 30), 150, 30)
         LIQUIEN09.List2.AddItem TTXX$
       End If
     Next U&
     Screen.MousePointer = 1
     '
     If LIQUIEN09.List2.ListCount < 1 Then
       Call MENSERR(24, "Sin Entregas Pendientes")
       GoTo 10
     End If
     '
     LIQUIEN09.Label4 = TRIM$(Str$(NC%))
     LIQUIEN09.Label7 = rasocli$(NC%)
     LIQUIEN09.Label8 = DOMICLI$(NC%)
     LIQUIEN09.Label10 = CPOSCLI$(NC%)
     LIQUIEN09.Label12 = LOCACLI$(NC%)
     LIQUIEN09.Label14 = FECHATEX$(FF%)
     LIQUIEN09.Label13 = DELIQ$
     '
     LIQUIEN09.Show 1
     '
   End If
   '
99 Command8.Enabled = True
End Sub


Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
     Call MENSERR(24, "Debe Salir con el Botón 'Exit'")
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub RECECOMO() ' recepcion de materiales Comodoro
        '
   'Call BLIMESS("Seleccionando ...")
   FIN& = ULTREG&("ATRASLA")
   Call BLANARCH("!ATRASLA")
   CABOPE% = 0
   FF% = HOY(FECHAREC$)
   '
   For i& = 1 To FIN&
     X$ = REGLEIDO$("ATRASLA", i&)
     If CVD(Mid$(X$, 105, 8)) >= 0.05 Then
       If CVD(Mid$(X$, 117, 8)) < 0.05 Then
         CANOM# = CVD(Mid$(X$, 105, 8))
         Call REPLA(X$, MKD$(CANOM#), 117, 8)
         Call REPLA(X$, MKS$(i&), 125, 4) ' registro
         Call REGAPP("!ATRASLA", X$)
       End If
     End If
   Next i&
   '
   If ULTREG&("!ATRASLA") < 1 Then
     Call COMUNI("No hay Recepciones\Pendientes de Ingreso.")
     Exit Sub
   End If
   '
   FERECEP$ = REPLACAR$(FECHAREC$, "/", "-")
   VTB% = VENTABU%("RECECOM/NC/TITUPAN=Boleta de Recepción - " + FERECEP$)
   '
   JJ& = 0
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
           NROORD% = NROORD% + 1
           Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NROORD%, VUNI#, 1, NC1%, DEPX%, CANTI)
           'Call Movisto(FF%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
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

Sub RECIBETRA()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ENTRECOM", "TRAENTRE")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
5   OPX% = COMALTER%("Archivo de Transferencia\\Recibir Vía Web\Desde Unidad de Disco")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 1 Then
        NC1% = XVALO(Control("RETRASLA", "SUC1CLIE"))
        RASOC1$ = TRIM$(ECOARCH$("DEUDOR1", MKI$(NC1%)))
        RUTA1$ = TRIM$(Control("RETRASLA", "SUC1RUTA"))
        
        NC2% = XVALO(Control("RETRASLA", "SUC2CLIE"))
        RASOC2$ = TRIM$(ECOARCH$("DEUDOR1", MKI$(NC2%)))
        RUTA2$ = TRIM$(Control("RETRASLA", "SUC2RUTA"))
        LISTAOP$ = RASOC1$ & "\" & RASOC2$
        OPXX% = COMALTER%("Sucursales\\" & LISTAOP$)
        If OPXX% < 1 Or OPXX% > 2 Then GoTo 5

        Select Case OPXX%
            Case 1: ORIGARCH$ = RUTA1$
            Case 2: ORIGARCH$ = RUTA2$
        End Select
        GoTo 18
        
'        CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
'        If CARPETASINCRO$ <> "" Then
'            ORIGARCH$ = CARPETASINCRO$
'            GoTo 18
'        Else
'            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
'            GoTo 5
'        End If
    End If
    
    FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then Exit Sub
    '
    LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       Exit Sub
    End If
18  On Error GoTo 20
    MsgBox RUTA2$
    FileCopy ORIGARCH$ + "\ATRACEN.DAT", LU$ + "ATRACEN.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Unidad de Transferencia\Leída Exitosamente.\  \Pulse 'Aceptar Mensaje' para\Iniciar Importación de Datos.")
    GoTo 30
    '
20  Resume 21
21  On Error GoTo 0
    Call MENSERR(24, "Imposible Leer Unidad '" + ORIGARCH$ + "' de Transferencia.")
    Exit Sub
    '
30  Call SELECHO("TADEPOSI/Almacen de Despachos")
    DEPO% = Val(VALACT1$("TADEPOSI"))
    If DEPO% < 1 Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim FINUAR%(128), ARTRAS$(128), UREAR&(128)
    Dim BUFGRA As String * 128
    CANAR% = 0
    LU$ = LUDAT$
    FIN& = ULTREG&("ATRACEN")
    For U& = 1 To FIN&
      X$ = REGLEIDO$("ATRACEN", U&)
      ANOM$ = Left$(X$, 8)
      TTXX$ = Mid$(X$, 9, 128)
      '
      For KKI% = 1 To CANAR%
        If ARTRAS$(KKI%) = ANOM$ Then GoTo 35
      Next KKI%
      CANAR% = CANAR% + 1
      KKI% = CANAR%
      FINUAR%(KKI%) = FILEOPEN%(LU$ + TRIM$(ANOM$) + ".DAT", 0, 128)
      ARTRAS$(KKI%) = ANOM$
      UREAR&(KKI%) = 0
      '
35    UREAR&(KKI%) = UREAR&(KKI%) + 1
      BUFGRA$ = TTXX$
      Put #FINUAR%(KKI%), UREAR&(KKI%), BUFGRA$
    Next U&
    Screen.MousePointer = 1
    '
    For KKI% = 1 To CANAR%
      FIN& = LOF(FINUAR%(KKI%)) / 128
      If LOF(FINUAR%(KKI%)) > 128 * FIN& Then FIN& = FIN& + 1
      For U& = UREAR&(KKI%) + 1 To FIN&
        BUFGRA$ = String$(128, 0)
        Put #FINUAR%(KKI%), UREAR&(KKI%), BUFGRA$
      Next U&
      Close #FINUAR%(KKI%)
      FINUAR%(KKI%) = 0
    Next KKI%
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For U& = 1 To ULTREG("TRAENTRE")
      XX$ = REGLEIDO$("TRAENTRE", U&)
      IDENDO$ = Mid$(XX$, 97, 4)
      RFF$ = RECFILT$("ENTRECOM", 7, IDENDO$)
      For UJ& = 1 To Len(RFF$) Step 4
        RELE& = CVS(Mid$(RFF$, UJ&, 4))
        If RELE& > 0 Then
          If RELE& <= ULTREG&("ENTRECOM") Then
            YY$ = REGLEIDO$("ENTRECOM", RELE&)
            'SE MODIFICA PORQUE SE INCORPORA LA RAZON SOCIAL EN 35,30 Y COMO EL ENTRECOM DE AVELLANEDA
            'NO LO TIENE RESULTA QUE SON DISTINTOS Y LO AGREGA POR SEGUNDA VEZ
            MM1$ = Left$(YY$, 190): Call REPLA(MM1$, Space(30), 35, 30)
            MM2$ = Left$(XX$, 190): Call REPLA(MM2$, Space(30), 35, 30)
            'If Left$(YY$, 190) = Left$(XX$, 190) Then GoTo 45 ' YA RECIBIDO
            If MM1$ = MM2$ Then GoTo 45 ' YA RECIBIDO
          End If
        End If
      Next UJ&
      '
      YY$ = Left$(XX$, 192) + String$(256, 0)
      Call REGAPP("ENTRECOM", YY$)
      '
      FF% = CVI(Left$(XX$, 2))
      CI% = CVI(Mid$(XX$, 117, 2))
      If CI% < 1 Or CI% > NUMAS% Then CI% = CODINT%(Mid$(XX$, 119, 16))
      LTX$ = Mid$(XX$, 3, 14)
      CMOV$ = Mid$(XX$, 3, 2)
      DOPRI$ = Mid$(XX$, 3, 14)
      DOSEC$ = ""
      REFE$ = Mid$(XX$, 35, 30)
      NC% = CVI(Mid$(XX$, 33, 2))
      
      If TRIM$(UCase$(CMOV$)) <> "AJ" Then
         CANTI = (-1) * CVD(Mid$(XX$, 183, 8))
      Else
         CANTI = CVD(Mid$(XX$, 183, 8)) 'MODIFICADO PARA QUE NO INVIERTA EL MOVIMIENTO CUANDO ES UN AJUSTE.
      End If
      Call MOVISTOK(FF%, CI%, LTX$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NROORD%, VUNI#, 1, NC%, DEPO%, CANTI)
      'Call MOVISTO(FF%, CI%, LTX$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
      '
45  Next U&
    '
    Call CIERRARCH("ENTRECOM")
    Call CIERRARCH("TRAENTRE")
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    Screen.MousePointer = 1
    Call COMUNI("Importación Completa.")
    '
End Sub



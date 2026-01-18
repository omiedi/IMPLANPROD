VERSION 5.00
Begin VB.Form OPDESPIMP 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despachos de Importación"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4550
      ScaleHeight     =   75
      ScaleWidth      =   825
      TabIndex        =   13
      Top             =   1285
      Width           =   885
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actualización"
      Height          =   1905
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   9
         Top             =   735
         Width           =   2640
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
            Left            =   -20
            Picture         =   "OPDESPIMP.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Carga Base de Datos de Despachos por Teclado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso Despachos de Importación"
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
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas - Listados"
      Height          =   2745
      Left            =   210
      TabIndex        =   1
      Top             =   2310
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Secuencial por Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   12
         Top             =   525
         Value           =   -1  'True
         Width           =   4215
      End
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
         Height          =   525
         Left            =   4095
         Picture         =   "OPDESPIMP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Emite Reporte según Opciones Elegidas"
         Top             =   2060
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   4
         Top             =   1995
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2100
            TabIndex        =   5
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Comparativo Despachos Ingresados vs Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   3
         Top             =   1365
         Width           =   4215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Código de Pieza o Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         TabIndex        =   2
         Top             =   945
         Width           =   3270
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
      Height          =   855
      Left            =   4540
      Picture         =   "OPDESPIMP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Finaliza - Cierra la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "OPDESPIMP.frx":075E
      Top             =   1585
      Width           =   2010
   End
End
Attribute VB_Name = "OPDESPIMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   '
   If Option1.Value = True Then
         If Option5.Value = True Then
             Call CONECRUN("*CONDESP", "Consulta Despachos por Codigo de Pieza")
           ElseIf Option4.Value = True Then
             Call FIDESPP
         End If
         GoTo 99
       ElseIf Option2.Value = True Then
         Call CONDESPA
       ElseIf Option3.Value = True Then
         Call LIDESPAFECHA
   End If
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   'Call CONECRUN("*CARDESP", "Carga Despachos Importacion por Número")
   Call CARDESPAIMP
   Command3.Enabled = True
End Sub


Private Sub Command4_Click()
   'FLIM% = FECHANUM("29/01/01")
   'For U& = 1 To ULTREG&("DESPANEW")
   '   X$ = REGLEIDO$("DESPANEW", U&)
   '   If CVI(Mid$(X$, 3, 2)) > FLIM% Then
   '      Call REGAPP("DESP-IMP", X$)
   '   End If
   'Next U&
   'Call CIERRARCH("*.*")
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
'
Sub CONDESPA()
   '
   Screen.MousePointer = 11
   Dim DESPA(32767), SALDE(32767), PUNDESPA$(32767)
   NM% = NUMAS%
   '
   FIN& = ULTREG&("DESP-IMP")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      X$ = REGLEIDO$("DESP-IMP", IL&)
      CI% = CVI(Left$(X$, 2))
      CADES = CVD(Mid$(X$, 33, 8))
      CASAL = CVD(Mid$(X$, 41, 8))
      If CI% > 0 Then
         If CI% <= NM% Then
            DESPA(CI%) = DESPA(CI%) + CADES
            SALDE(CI%) = SALDE(CI%) + CASAL
            PUNDESPA$(CI%) = PUNDESPA$(CI%) + MKS$(IL&)
         End If
      End If
   Next IL&
   '
   JJ& = 0
   FIN& = NM%
   URG& = ULTREG&("DESP-IMP")
   For CI% = 1 To NM%
      IL& = CI%
      Call TRACE(20, IL&, FIN&)
      P$ = Left$(CODEXT$(CI%), 1)
      If P$ <> "T" And P$ <> "W" Then
         STO = STOCKACT(CI%)
         CAPUN% = (Len(PUNDESPA$(CI%))) / 4
         If STO > DESPA(CI%) + 0.05 Then
              X$ = MKI$(CI%) + MKS$(STO) + MKS$(DESPA(CI%)) + MKS$(STO - DESPA(CI%)) + MKI$(0)
              JJ& = JJ& + 1
              Call GRAREG("CONDESP", X$, JJ&)
              For K& = 1 To CAPUN%
                 IL& = CVS(Mid$(PUNDESPA$(CI%), 4 * K& - 3, 4))
                 If IL& > 0 Then
                   If IL& <= URG& Then
                     X$ = REGLEIDO$("DESP-IMP", IL&)
                     SALCAPEN# = CVD(Mid$(X$, 33, 8))
                     Call REPLA(X$, MKD$(SALCAPEN#), 41, 8)
                     Call GRAREG("DESP-IMP", X$, IL&)
                   End If
                 End If
              Next K&
            Else
              For K& = CAPUN% To 1 Step -1
                 IL& = CVS(Mid$(PUNDESPA$(CI%), 4 * K& - 3, 4))
                 If IL& > 0 Then
                   If IL& <= URG& Then
                     X$ = REGLEIDO$("DESP-IMP", IL&)
                     SALCAPEN# = CVD(Mid$(X$, 33, 8))
                     If SALCAPEN# > STO Then SALCAPEN# = CDbl(STO)
                     Call REPLA(X$, MKD$(SALCAPEN#), 41, 8)
                     Call GRAREG("DESP-IMP", X$, IL&)
                     STO = STO - SALCAPEN#
                   End If
                 End If
              Next K&
         End If
      End If
   Next CI%
   Call SETULTREG("CONDESP", JJ&)
   Call CIERRARCH("CONDESP")
   '
   Call CIERRARCH("DESP-IMP")
   Call CIERRARCH("*.*")
   '
   FIN& = ULTREG&("DESP-IMP")
   JJ& = ULTREG&("DESPAIMP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DESP-IMP", U&)
      Call GRAREG("DESPAIMP", X$, U&)
      JJ& = U&
   Next U&
   Call SETULTREG("DESPAIMP", JJ&)
   Call CIERRARCH("DESPAIMP")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   If ULTREG&("CONDESP") < 1 Then
      Call COMUNI("Despachos Cubren el Stock")
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Call CIERRARCH("*.*")
   Call FINAL("*SOCONDE")
   '
End Sub
'

Sub FIDESPP()
   '
10 XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   '
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
      GoTo 99
   End If
   CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   If cod1$ > cod2$ Then GoTo 10
   '
   OPXX% = COMALTER%("Puede Suprimir Despachos Sin Saldo\\Suprimir\Listar Todos")
   If OPXX% < 1 Or OPXX% > 2 Then GoTo 10
   '
   Call BLOQARCH("DESPAIML")
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("DESPAIMP")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("DESPAIMP", IL&)
     CI% = CVI(Left$(X$, 2))
     CODD$ = CODEXT$(CI%)
     If CODD$ >= cod1$ Then
       If CODD$ <= cod2$ Then
         If CVD(Mid$(X$, 41, 8)) >= 0.05 Or OPXX% = 2 Then
           JJ& = JJ& + 1
           Call GRAREG("DESPAIML", X$, JJ&)
         End If
       End If
     End If
   Next IL&
   '
   Call SETULTREG("DESPAIML", JJ&)
   Call CIERRARCH("DESPAIML")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call FILESORT("DESPAIML", "", 1, 4)
   Call CONECRUN("*LIDESPP", "Listado Despachos por Codigo de Pieza")
   '
99 Screen.MousePointer = 1
   '
End Sub
'

Sub LIDESPAFECHA()
   '
   If Option5.Value = True Then
      Call COMUNI("Listado no Disponible por Pantalla.\  \Emita por Impresora\y Exporte a Archivo de Texto.")
      Exit Sub
   End If
   '
10 Call COPYSTRU("DESPAIMP", "DESPAIML")
   FEX = HOY(HO$)
   HAS% = FEX
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 10
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Call BLOQARCH("DESPAIML")
   FIN& = ULTREG("DESPAIMP")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DESPAIMP", U&)
      FEDE% = CVI(Mid$(X$, 3, 2))
      If FEDE% >= DES% Then
         If FEDE% <= HAS% Then
            CIXI% = CVI(Left$(X$, 2))
            If CVI(Mid$(X$, 57, 2)) <> CIXI% Then
               Call REPLA(X$, MKI$(CIXI%), 57, 2)
               Call GRAREG("DESPAIMP", X$, U&)
            End If
            JJ& = JJ& + 1
            Call GRAREG("DESPAIML", X$, JJ&)
         End If
      End If
   Next U&
   '
   Call SETULTREG("DESPAIML", JJ&)
   Call CIERRARCH("DESPAIML")
   Call CIERRARCH("DESPAIMP")
   '
   Call CONECRUN("*LIDESPF", "Listado Despachos por Rango de Fechas")
   '
End Sub


Sub CARDESPAIMP()
   '
   OBX$ = TRIM$(OBJPLA$("ENTRADESPA", ""))
   If OBX$ = "" Then Exit Sub
   '
   ABUS$ = OBX$
   Screen.MousePointer = 11
   JJ& = 0
   Call COPYSTRU("DESPAIMP", "DESP-IMP")
   If ULTREG&("DESP-IMP") < 1 Then
      FIN& = ULTREG&("DESPAIMP")
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$("DESPAIMP", U&)
         Call GRAREG("DESP-IMP", X$, U&)
      Next U&
      Call CIERRARCH("DESP-IMP")
   End If
   '
   FIN& = ULTREG&("DESP-IMP")
   '
   ENCADES$ = "": IDENDES$ = ""
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DESP-IMP", U&)
      If TRIM$(Mid$(X$, 17, 16)) = ABUS$ Then
         Call REPLA(X$, TRIM$(Str$(U&)), 59, 6)
         JJ& = JJ& + 1
         Call GRAREG("!DESPAIMP", X$, JJ&)
         If ENCADES$ = "" Then
            FEX = CVI(Mid$(X$, 3, 2))
            FEDES$ = REPLACAR$(FECHATEX$(FEX), "/", ".")
            ENCADES$ = "Despacho: " + ABUS$ + " - Fecha: " + FEDES$ + " - Aduana: " + TRIM$(Mid$(X$, 5, 12))
            IDENDES$ = Mid$(X$, 3, 30)
         End If
      End If
   Next U&
   '
   Call SETULTREG("!DESPAIMP", JJ&)
   Call CIERRARCH("!DESPAIMP")
   Call CIERRARCH("DESP-IMP")
   Screen.MousePointer = 1
   '
   If ENCADES$ = "" Then
      FEXI% = HOY(HOS$)
      VDEF$ = MKI$(FEXI%) + Space$(12) + AJUSTI$(ABUS$, 16)
10    OBX$ = OBJPLA$("IDENDESPA", VDEF$)
      If OBX$ = "" Then Exit Sub
      VDEF$ = OBX$
      '
      FEX = CVI(Mid$(OBX$, 1, 2))
      ADUA$ = TRIM$(Mid$(OBX$, 3, 12))
      If ADUA$ = "" Then GoTo 10
      '
      FEDES$ = REPLACAR$(FECHATEX$(FEX), "/", ".")
      ENCADES$ = "Despacho: " + ABUS$ + " - Fecha: " + FEDES$ + " - Aduana: " + ADUA$
      IDENDES$ = OBX$
   End If
   '
   VTB% = VENTABU%("CARDESP/TITUPAN=" + ENCADES$)
   If VTB% < 0 Then Exit Sub
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("!DESPAIMP")
      X$ = REGLEIDO$("!DESPAIMP", U&)
      REDESPA& = Val(Mid$(X$, 59, 8))
      If REDESPA& > 0 Then
         If REDESPA& <= ULTREG&("DESP-IMP") Then
            Y$ = REGLEIDO$("DESP-IMP", REDESPA&)
            If Left$(Y$, 32) = Left$(X$, 32) Then
               DIFECAN# = CVD(Mid$(X$, 33, 8)) - CVD(Mid$(Y$, 33, 8))
               SALNUEVO# = CVD(Mid$(Y$, 41, 8)) + DIFECAN#
               If SALNUEVO# < 0 Then SALNUEVO# = 0
               Call REPLA(X$, MKD$(SALNUEVO#), 41, 8)
               Call REPLA(X$, String$(6, 0), 59, 6)
               Call GRAREG("DESP-IMP", X$, REDESPA&)
               GoTo 19
            End If
         End If
      End If
      '
      If CVI(Left$(X$, 2)) > 0 Then
         Call REPLA(X$, IDENDES$, 3, 30)
         SALNUEVO# = CVD(Mid$(X$, 33, 8))
         Call REPLA(X$, MKD$(SALNUEVO#), 41, 8)
         Call REPLA(X$, String$(6, 0), 59, 6)
         Call REGAPP("DESP-IMP", X$)
      End If
      '
19 Next U&
   '
   Call CIERRARCH("DESP-IMP")
   Call CIERRARCH("*.*")
   '
   FIN& = ULTREG&("DESP-IMP")
   JJ& = ULTREG&("DESPAIMP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DESP-IMP", U&)
      Call GRAREG("DESPAIMP", X$, U&)
      JJ& = U&
   Next U&
   Call SETULTREG("DESPAIMP", JJ&)
   Call CIERRARCH("DESPAIMP")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
End Sub


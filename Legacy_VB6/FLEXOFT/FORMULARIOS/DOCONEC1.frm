VERSION 5.00
Begin VB.Form DOCONEC1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculación Documentos Relacionados"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Archivo de Conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   9465
      Begin VB.CommandButton Command3 
         Height          =   960
         Left            =   8715
         Picture         =   "DOCONEC1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Iniciar Lectura de Archivo"
         Top             =   2460
         Width           =   645
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   3780
         TabIndex        =   9
         Top             =   420
         Width           =   5475
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8400
         TabIndex        =   13
         Top             =   3105
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   12
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2940
         TabIndex        =   10
         Top             =   2460
         Width           =   5685
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   8
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   315
         TabIndex        =   7
         Top             =   2895
         Width           =   8310
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -210
         TabIndex        =   6
         Top             =   2625
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Convert"
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
      Left            =   9135
      Picture         =   "DOCONEC1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Tabla Conversión de Códigos Comerciales - Técnicos"
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clientes"
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
      Left            =   9135
      Picture         =   "DOCONEC1.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Productos Importados (Material de Reventa)"
      Top             =   1260
      Width           =   855
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
      Left            =   9135
      Picture         =   "DOCONEC1.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8580
      Picture         =   "DOCONEC1.frx":09E0
      Top             =   3270
      Width           =   2010
   End
End
Attribute VB_Name = "DOCONEC1"
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
   Call CONECRUN("AMACLI02", "A-B-M-L Clientes Cuenta Corriente")
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    RUTARCH$ = Label11 + "\" + Label1
    FileCopy RUTARCH$, "C:\FLEXOFT\PRUEBA.DOC"
    'WORDDOCS.OLE1.CreateEmbed "F:\SOLO_LEER\ARCHIVOSCIME-SOI\Plantillas Word y Excel\hOJAS FLEXOFT.DOT"
    WORDDOCS.OLE1.SourceDoc = "C:\FLEXOFT\PRUEBA.DOC" ' Label11 + "\" + Label1
    
    'WORDDOCS.OLE1.CreateEmbed Label11 + "\" + Label1
    WORDDOCS.OLE1.Action = 1
    WORDDOCS.OLE1.Action = 6
    'AAA = WORDDOCS.OLE1.OLEType
    'ABC$ = WORDDOCS.OLE1.SourceItem
    WORDDOCS.Show 1
    Command3.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    '
   Command5.Enabled = False
   Call IMPORMOVI
   Screen.MousePointer = 1
   Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CARCONVERCO
   Command6.Enabled = True
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   'Call MENSERR(24, "Unidad no Disponible.")
   'Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub File1_Click()
  Label1 = File1.filename
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Show
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     'Call CIERRARCH("*.*")
     'Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Call CIERRARCH("*.*")
   'Call FINAL("")
End Sub

Sub LEEARTRA()
   '
   If TRIM$(Label1) = "" Then
     Call MENSERR(24, "Falta Nombre de Archivo")
     Exit Sub
   End If
   '
   ARIMPO$ = TRIM$(Label11)
   If Right$(ARIMPO$, 1) <> "\" Then
     ARIMPO$ = ARIMPO$ + "\"
   End If
   ARIM$ = TRIM$(Label1): If InStr(ARIM$, ".") < 1 Then ARIM$ = ARIM$ + "."
   ARIMPO$ = ARIMPO$ + ARIM$
   '
   If EXISTE%(ARIMPO$) < 1 Then
     Call MENSERR(24, "Archivo Inexistente")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Call INDEXARCH("IMPORMOV", 1)
   REBLA$ = REGBLAN$("IMPORMOV")
   n1% = FILEOPEN%(ARIMPO$, 1, 128)
   If EOF(n1%) Then GoTo 99
   Line Input #n1%, AA$     ' REGISTRO DE CABECERA
   '
   CARELE& = 0
   Label3 = ""
   Command3.Height = 540
   DoEvents
   While Not EOF(n1%)
     Line Input #n1%, AA$
     Label3 = TRIM$(Str$(Val(Label3) + 1))
     Label3.Refresh
     '
     TICOM$ = Mid$(AA$, 2, 3)
     NUCOM$ = Mid$(AA$, 7, 7)
     LICOM$ = Mid$(AA$, 15, 5)
     FECOM$ = Mid$(AA$, 21, 6)
     CLCOM$ = TRIM$(Mid$(AA$, 29, 10))
     COCOM$ = TRIM$(Mid$(AA$, 88, 13))
     CACOM$ = Mid$(AA$, 103, 12)
     PRCOM$ = Mid$(AA$, 116, 12)
     BOCOM$ = Mid$(AA$, 129, 6)
     '
     If COCOM$ <> "" Then
       If UCase$(COCOM$) <> "COMENT" Then
         If UCase$(COCOM$) <> "CRED" Then
           '
           FETEX$ = Left$(FECOM$, 2) + "/" + Mid$(FECOM$, 3, 2) + "/" + Mid$(FECOM$, 5, 2)
           FEMOI% = FECHANUM(FETEX$)
           '
           NCLIX& = Val(Mid$(CLCOM$, 2))
           While NCLIX& > 10000: NCLIX& = NCLIX& - 10000: Wend
           NCLIE% = Abs(NCLIX&)
           If UCase$(Left$(CLCOM$, 2)) = "P" Then
             NCLIE% = (-1) * NCLIE%
           End If
           '
           LINEA& = Val(LICOM$)
           While LINEA& > 10000: LINEA& = LINEA& - 10000: Wend
           ILIN% = LINEA&
           '
           REIMPO$ = REBLA$
           Call REPLA(REIMPO$, TICOM$, 1, 4)
           Call REPLA(REIMPO$, MKS$(Val(NUCOM$)), 5, 4)
           Call REPLA(REIMPO$, MKI$(ILIN%), 9, 2)
           Call REPLA(REIMPO$, MKI$(FEMOI%), 11, 2)
           Call REPLA(REIMPO$, MKI$(NCLIE%), 13, 2)
           Call REPLA(REIMPO$, AJUSTI$(COCOM$, 16), 15, 16)
           Call REPLA(REIMPO$, CACOM$, 31, 12)
           Call REPLA(REIMPO$, PRCOM$, 43, 12)
           Call REPLA(REIMPO$, BOCOM$, 55, 6)
           Call REPLA(REIMPO$, MKS$(0), 61, 4)
           '
'TTX1$ = "": For KKI% = 1 To 60: TTX1$ = TTX1$ + Str$(Asc(Mid$(REIMPO$, KKI%, 1))): Next KKI%
           RFF$ = RECFILT$("IMPORMOV", 4, MKI$(FEMOI%))
'AAA = Len(RFF$)
'BBB = FEMOI%: CCC$ = FECHATEX$(BBB)
           For KKL& = 1 To Len(RFF$) Step 4
             REIM& = CVS(Mid$(RFF$, KKL&, 4))
             If REIM& > 0 Then
               RELEIM$ = REGLEIDO$("IMPORMOV", REIM&)
'TTX2$ = "": For KKI% = 1 To 60: TTX2$ = TTX2$ + Str$(Asc(Mid$(RELEIM$, KKI%, 1))): Next KKI%
'MsgBox TTX1$ + "          " + TTX2$
               If Left$(RELEIM$, 60) = Left$(REIMPO$, 60) Then GoTo 19
             End If
           Next KKL&
           '
           Call REGAPP("IMPORMOV", REIMPO$)
           Call CIERRARCH("IMPORMOV")
           CARELE& = CARELE& + 1
           '
         End If
       End If
     End If
     '
19 Wend
   '
   Call CIERRARCH("IMPORMOV")
'ABC$ = LOADFILE$(LUDAT$ + "IMPORMOV.X04")
'TTX1$ = "": For KKI% = 1 To Len(ABC$) Step 2: TTX1$ = TTX1$ + Str$(CVI(Mid$(ABC$, KKI%, 2))): Next KKI%
'MsgBox TTX1$
   Close #n1%
   Call BAJALDISCO
   Command3.Height = 960
   DoEvents
   Screen.MousePointer = 1
   Call COMUNI("Proceso Completo\Leidos " + TRIM$(Str$(CARELE&)) + " Nuevos Movimientos.")
   '
99 End Sub

Sub IMPORMOVI()
   '
   Dim PUCLINOAL As String
   '
   For KKI% = 0 To 3
     Label6(KKI%) = ""
     Label6(KKI%).Refresh
   Next KKI%
   '
   RFF$ = RECFILT$("IMPORMOV", 10, MKS$(0))
   If Len(RFF$) < 4 Then
     Call MENSERR(24, "No hay Movimientos\Pendientes de Importación.")
     Exit Sub
   End If
   '
   FALTACODI% = 0
   REBLACO$ = REGBLAN$("CONVERCO")
   Label6(3) = TRIM$(Str$(Len(RFF$) / 4))
   Label6(3).Refresh
   '
   Screen.MousePointer = 11
   For UL& = 1 To Len(RFF$) Step 4
     Label6(0) = TRIM$(Str$((UL& + 3) / 4))
     Label6(0).Refresh
     REIM& = CVS(Mid$(RFF$, UL&, 4))
     '
10   XX$ = REGLEIDO$("IMPORMOV", REIM&)
     CODICOM$ = AJUSTI$(TRIM$(Mid$(XX$, 15, 16)), 24)
     REQUI& = REGBUS&("CONVERCO", 1, CODICOM$)
     If REQUI& < 1 Then
       ZZ$ = REBLACO$
       Call REPLA(ZZ$, CODICOM$, 1, 24)
       Call REPLA(ZZ$, String$(8, 0), 25, 8)
       Call REGAPP("CONVERCO", ZZ$)
       Call CIERRARCH("CONVERCO")
       Label6(2) = TRIM$(Str$(1 + Val(Label6(2))))
       Label6(2).Refresh
       FALTACODI% = 1
       GoTo 19
     End If
     '
     ZZ$ = REGLEIDO$("CONVERCO", REQUI&)
     CIEQUI% = CVI(Mid$(ZZ$, 25, 2))
     If TRIM$(CODEXT$(CIEQUI%)) = "" Then
       Label6(2) = TRIM$(Str$(1 + Val(Label6(2))))
       Label6(2).Refresh
       FALTACODI% = 1
       GoTo 19
     End If
     '
     NCLIE% = CVI(Mid$(XX$, 13, 2))
     If REGICLI&(NCLIE%) < 1 Then
       Label6(2) = TRIM$(Str$(1 + Val(Label6(2))))
       DoEvents
       '
       For KKI% = 1 To Len(PUCLINOAL$) Step 2
         If CVI(Mid$(PUCLINOAL$, KKI%, 2)) = NCLIE% Then GoTo 19
       Next KKI%
       '
       TTXX$ = "Cuenta " + TRIM$(Str$(NCLIE%)) + " no Registrada\\Alta Ahora\Continuar\Cancelar"
       Select Case COMALTER%(TTXX$)
         Case 1
           Call CIERRARCH("*.*")
           Call CONECRUN("AMACLI02", "A-B-M-L Clientes Cuenta Corriente")
           Exit Sub
         Case 3
           Call CIERRARCH("*.*")
           Exit Sub
         Case Else
           PUCLINOAL$ = PUCLINOAL$ + MKI$(NCLIE%)
           GoTo 19
       End Select
       '
     End If
     '
15   FECHA% = CVI(Mid$(XX$, 11, 2))
     LOTE$ = ""
     CMOV$ = "RT"
     DOPRI$ = "RT." + TRIM$(Str$(CVS(Mid$(XX$, 5, 4))))
     SIG% = (-1)
     If Left$(XX$, 2) = "NC" Then
       DOPRI$ = "DV." + TRIM$(Str$(CVS(Mid$(XX$, 5, 4))))
       SIG% = 1
     End If
     DOSEC$ = Left$(XX$, 2) + "." + TRIM$(Str$(CVS(Mid$(XX$, 5, 4))))
     REFE$ = RASOCLI$(NCLIE%)
     VUNI# = (Val(Mid$(XX$, 43, 12))) * (1 - (Val(Mid$(XX$, 55, 6))) / 100)
     MONE$ = "$"
     DEPO% = 4
     CANTI = SIG% * Abs(Val(Mid$(XX$, 31, 12)))
     '
     For UK% = 25 To 32 Step 2
       CIEQUI% = CVI(Mid$(ZZ$, UK%, 2))
       If CIEQUI% > 0 Then
         Call MOVISTO(FECHA%, CIEQUI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NCLIE%, DEPO%, CANTI)
       End If
     Next UK%
     '
     UREMO& = ULTREG&("MOVISTO")
     Call CIERRARCH("MOVISTO")
     '
     Call REPLA(XX$, MKS$(UREMO&), 61, 4)
     Call GRAREG("IMPORMOV", XX$, REIM&)
     '
     Label6(1) = TRIM$(Str$(1 + Val(Label6(1))))
     Label6(1).Refresh
     '
19 Next UL&
   Call CIERRARCH("CONVERCO")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   If FALTACODI% > 0 Then
      Call FILESORT("CONVERCO", "", 1, 1)
      Call CARCONVERCO
   End If
   '
End Sub

Sub CARCONVERCO()
   '
   JJ& = 0: CODANT$ = ""
   FIN& = ULTREG&("CONVERCO")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("CONVERCO", U&)
     If TRIM$(Left$(XX$, 24)) <> CODANT$ Then
       CODANT$ = TRIM$(Left$(XX$, 24))
       JJ& = JJ& + 1
       Call GRAREG("!CONVERCO", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CONVERCO", JJ&)
   Call CIERRARCH("!CONVERCO")
   '
   VTB% = VENTABU%("CONVERCO")
   If VTB% >= 0 Then
     Call TRACENTRAL("CONVERCO", "")
     Call FILESORT("CONVERCO", "", 1, 1)
   End If
   '
End Sub

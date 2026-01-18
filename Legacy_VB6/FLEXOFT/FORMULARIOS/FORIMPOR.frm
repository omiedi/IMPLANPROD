VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIMPOR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Maestro"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command15 
      Caption         =   "Pl.Cuentas"
      Height          =   225
      Left            =   6615
      TabIndex        =   22
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Deu-Pro"
      Height          =   225
      Left            =   2415
      TabIndex        =   21
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Provedor"
      Height          =   225
      Left            =   5565
      TabIndex        =   20
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Trans portes"
      Height          =   645
      Left            =   11025
      TabIndex        =   19
      Top             =   5460
      Width           =   750
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Codigos"
      Height          =   225
      Left            =   315
      TabIndex        =   18
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deu-Cli"
      Height          =   225
      Left            =   1365
      TabIndex        =   17
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Precios"
      Height          =   225
      Left            =   3465
      TabIndex        =   16
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clientes"
      Height          =   225
      Left            =   4515
      TabIndex        =   15
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   540
      Left            =   11025
      TabIndex        =   14
      Top             =   3045
      Width           =   750
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   540
      Left            =   11025
      TabIndex        =   13
      Top             =   1785
      Width           =   750
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   540
      Left            =   11760
      TabIndex        =   12
      Top             =   4620
      Width           =   750
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   315
      TabIndex        =   11
      Top             =   105
      Width           =   8520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   540
      Left            =   11025
      TabIndex        =   10
      Top             =   2415
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11760
      Picture         =   "FORIMPOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Inicia Importación"
      Top             =   3780
      Width           =   750
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\CLIENTES.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "'cli1#csv-vdr'$"
      Top             =   2415
      Width           =   1905
   End
   Begin VB.CommandButton Command4 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMPOR.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\art-cj-finales.xls-vdr.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Hoja1$"
      Top             =   105
      Width           =   1905
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
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMPOR.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIMPOR.frx":075E
      Height          =   1800
      Left            =   0
      OleObjectBlob   =   "FORIMPOR.frx":076E
      TabIndex        =   1
      Top             =   315
      Width           =   10830
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FORIMPOR.frx":1120
      Height          =   5475
      Left            =   0
      OleObjectBlob   =   "FORIMPOR.frx":1130
      TabIndex        =   8
      Top             =   2625
      Width           =   10830
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aro Plastico"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resorte"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   4
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   315
      Width           =   1380
   End
End
Attribute VB_Name = "FORIMPOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
   Close
   End
End Sub

Private Sub Command10_Click()
FORIMDEU.Show 1
End Sub

Private Sub Command11_Click()
FORIMCOD.Show 1
End Sub

Private Sub Command12_Click()
   URECORD& = Data2.Recordset.Recordcount
   DBGrid2.Height = (URECORD& + 50) * DBGrid2.RowHeight
   '
   LU$ = LUDAT$
   '
   '
   For U& = 0 To URECORD&
     On Error Resume Next
     DBGrid2.Row = U&
     If Err Then GoTo 19
     On Error GoTo 0
     DBGrid2.Col = 0: NC% = Val(DBGrid2.TEXT)
     DBGrid2.Col = 30: NTRA% = Val(DBGrid2.TEXT)
     DBGrid2.Col = 31: DENOTRA$ = TRIM$(DBGrid2.TEXT)
     DBGrid2.Col = 33: CONTACTO$ = DBGrid2.TEXT
     '
     If NC% > 0 Then
       If NTRA% > 0 Then
         For KU& = ULTREG&("TRANSPOR") + 1 To NTRA%
           TTTT$ = REGBLAN$("TRANSPOR")
           KUI% = KU&
           Call REPLA(TTTT$, MKI$(KUI%), 1, 2)
           Call REPLA(TTTT$, "...", 3, 32)
           Call GRAREG("TRANSPOR", TTTT$, KU&)
         Next KU&
         '
         KU& = NTRA%
         TTTT$ = REGLEIDO$("TRANSPOR", KU&)
         If DENOTRA$ <> "" Then
           If Mid$(TTTT$, 3, 3) = "..." Then
             Call REPLA(TTTT$, DENOTRA$, 3, 32)
             Call REPLA(TTTT$, DENOTRA$, 35, 48)
             Call REPLA(TTTT$, Chr$(1), 194, 1)
             Call REPLA(TTTT$, CONTACTO$, 339, 48)
             Call GRAREG("TRANSPOR", TTTT$, KU&)
           End If
         End If
         TTTT$ = REGLEIDO$("TRANSPOR", KU&)
         DENOTRA$ = TRIM$(Mid$(TTTT$, 3, 32))
         '
         RFF$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
         If Len(RFF$) >= 4 Then
             RELEN& = CVS(Left$(RFF$, 4))
           Else
             RELEN& = 1 + ULTREG&("LENTRECL")
         End If
         XX$ = REGLEIDO$("LENTRECL", RELEN&)
         Call REPLA(XX$, MKI$(NC%), 1, 2)
         If TRIM$(Mid$(XX$, 3, 32)) = "" Then
           Call REPLA(XX$, DOMICLI$(NC%), 3, 32)
         End If
         If TRIM$(Mid$(XX$, 35, 32)) = "" Then
           Call REPLA(XX$, LOCACLI$(NC%), 35, 32)
         End If
         Call REPLA(XX$, DENOTRA$, 67, 32)
         Call REPLA(XX$, MKI$(NTRA%), 99, 2)
         Call GRAREG("LENTRECL", XX$, RELEN&)
       End If
     End If
19 Next U&
   '
   Call CIERRARCH("TRANSPOR")
   Call CIERRARCH("LENTRECL")
   Call CIERRARCH("*.*")
   MsgBox "FIN"
End Sub

Private Sub Command13_Click()
   FORIMPRO.Show 1
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   FORIDEUP.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   'FORIPLAN.Show 1
   LU$ = LUDAT$
   '
   N1% = FILEOPEN%("C:\FLEXOFT\PLASANJ.TXT", 1, 128)
   KK& = 0
   JJ1& = ULTREG&("CUENTAS")
   While ULTREG&("DENOCUE") < JJ1&
     Call REGAPP("DENOCUE", ".")
     Call CIERRARCH("DENOCUE")
   Wend
   JJ2& = ULTREG&("INVECUE")
   '
   Line Input #N1%, A$
   '
   While Not EOF(N1%)
     Line Input #N1%, A$
     COCUE$ = TRIM$(Left$(A$, 12))
     DECUE$ = TRIM$(Mid$(A$, 13))
     NIVEI% = 4
     If Right$(COCUE$, 3) = "/00" Then NIVEI% = 3
     If Right$(COCUE$, 6) = "/00/00" Then NIVEI% = 2
     If Right$(COCUE$, 8) = ".0/00/00" Then NIVEI% = 1
     If Right$(COCUE$, 10) = ".0.0/00/00" Then NIVEI% = 0
     '
     For U& = 1 To ULTREG&("CUENTAS")
       XX$ = REGLEIDO$("CUENTAS", U&)
       If TRIM$(Left$(XX$, 12)) = COCUE$ Then
         JI% = U&
         YY$ = Space$(128)
         Call REPLA(YY$, DECUE$, 1, 100)
         Call REPLA(YY$, DECUE$, 101, 28)
         Call GRAREG("DENOCUE", YY$, U&)
         GoTo 20
       End If
     Next U&
     '
     XX$ = REGBLAN$("CUENTAS")
     Call REPLA(XX$, COCUE$, 1, 12)
     Call REPLA(XX$, "D", 13, 1)
     '
     YY$ = Space$(128)
     Call REPLA(YY$, DECUE$, 1, 100)
     Call REPLA(YY$, DECUE$, 101, 28)
     '
     JJ1& = JJ1& + 1: JI% = JJ1&
     JJ2& = JJ2& + 1
     '
     ZZ$ = REGBLAN$("INVECUE")
     Call REPLA(ZZ$, COCUE$, 1, 12)
     Call REPLA(ZZ$, MKI$(JI%), 13, 2)
     '
     Call GRAREG("CUENTAS", XX$, JJ1&)
     Call GRAREG("DENOCUE", YY$, JJ1&)
     Call GRAREG("INVECUE", ZZ$, JJ2&)
     '
20   QQ$ = REGBLAN$("PLAGRAL")
     Call REPLA(QQ$, MKI$(JI%), 1, 2)
     Call REPLA(QQ$, MKI$(NIVEI%), 3, 2)
     KK& = KK& + 1
     Call GRAREG("PLAGRAL", QQ$, KK&)
     '
19 Wend
   '
   Call SETULTREG("CUENTAS", JJ1&)
   Call SETULTREG("DENOCUE", JJ1&)
   Call SETULTREG("INVECUE", JJ2&)
   Call SETULTREG("PLAGRAL", KK&)
   '
   Call CIERRARCH("CUENTAS")
   Call CIERRARCH("DENOCUE")
   Call CIERRARCH("PLAGRAL")
   Call FILESORT("INVECUE", "", 1, 1)
   Call CIERRARCH("*.*")
   '
MsgBox "COMPLETO"

   Command15.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Dim RESUR$(20)
   'resur$(1) = "'cjtos# VGA'$"
   'resur$(1) = "'repar# VGA'$"
   RESUR$(1) = "'cjtos SR33'$"
   'resur$(4) = "'cjtos SR33-1'$"
   RESUR$(2) = "'repar SR33'$"
   RESUR$(3) = "'repar SR33-1'$"
   RESUR$(4) = "'repar SR33-2'$"
   RESUR$(5) = "'repar SR33-3'$"
   RESUR$(6) = "'cil. rueda-1'$"
   'resur$(10) = "'cil. rueda'$"
   RESUR$(7) = "servos$"
   '
   N1% = FILEOPEN%(LUDAT$ + "ESTRUCRE.TXT", 2, 128)
   For UCOMA% = 1 To 11
      Data2.RecordSource = RESUR$(UCOMA%)
      Data2.Refresh
      CACOL% = DBGrid2.Columns.Count
      CAFIL% = Data2.Recordset.Recordcount
      '
      DBGrid2.RowHeight = 25
      DBGrid2.Height = (CAFIL% + 50) * DBGrid2.RowHeight
      For U& = 4 To CACOL% Step 2
        If U& > 255 Then GoTo 29
         DBGrid2.Row = 0
         DBGrid2.Col = U&
         COCON$ = TRIM$(DBGrid2.Columns(U&).Caption)
         '
         If COCON$ <> "" Then
           If Left$(COCON$, 1) <> "F" Then
            For JJ& = 6 To CAFIL%
              'On Error Resume Next
              DBGrid2.Col = U& - 1
              DBGrid2.Row = JJ&
              'If Err Then GoTo 29
              On Error GoTo 0
              USOELE = Val(DBGrid2.TEXT)
              If USOELE > 0 Then
                DBGrid2.Col = 1
                DBGrid2.Row = JJ&
                On Error GoTo 0
                COELE$ = DBGrid2.TEXT
                If COELE$ = "" Then
                  DBGrid2.Col = 2
                  COELE$ = DBGrid2.TEXT
                End If
'If COELE$ = "2400752" Or COELE$ = "2404706" Or COELE$ = "2404707" Then
'AAA = BBB
'End If
                DBGrid2.Col = 0
                VALOUN$ = FORMATNUM$(Val(DBGrid2.TEXT), "F12.4")
                
                TTXX$ = Space$(100)
                Call REPLA(TTXX$, COCON$, 1, 20)
                Call REPLA(TTXX$, COELE$, 21, 20)
                Call REPLA(TTXX$, Str$(USOELE), 41, 10)
                Call REPLA(TTXX$, VALOUN$, 55, 12)
                Call REPLA(TTXX$, RESUR$(UCOMA%), 71, 20)
                Print #N1%, TTXX$
                '
                List1.AddItem TTXX$
                List1.ListIndex = List1.ListCount - 1
                List1.Refresh
              End If
19          Next JJ&
           End If
         End If
         '
29    Next U&
   Next UCOMA%
   Close #N1%
   MsgBox "FIN"
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
  PUCOX$ = ""
  N1% = FILEOPEN%(LUDAT$ + "ESTRUC.TXT", 1, 128)
  While Not EOF(N1%)
    Line Input #N1%, AA$
    COCON$ = TRIM$(Left$(AA$, 20))
    If COCON$ <> "" Then
      CICON% = CODINT%(COCON$)
      If CICON% < 1 Then
        If InStr(PUCOX$, AJUSTI$(COCON$, 16)) < 1 Then
          DEDEX$ = "Articulo " + COCON$
          RENUE$ = REGBLAN$("MASTER")
          NMS% = ULTREG&("MASTER")
          Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
          Call REPLA(RENUE$, COCON$, 77, 16)
          Call REPLA(RENUE$, DEDEX$, 5, 64)
          Call REPLA(RENUE$, Chr$(1), 93, 1)
          Call REPLA(RENUE$, "U.", 109, 2)
          Call REGAPP("MASTER", RENUE$)
          PUCOX$ = PUCOX$ + AJUSTI$(COCON$, 16)
          AGREGA% = 1
        End If
      End If
    End If
    '
    COCON$ = TRIM$(Mid$(AA$, 21, 20))
    If COCON$ <> "" Then
      CICON% = CODINT%(COCON$)
      If CICON% < 1 Then
        If InStr(PUCOX$, AJUSTI$(COCON$, 16)) < 1 Then
          DEDEX$ = "Articulo " + COCON$
          RENUE$ = REGBLAN$("MASTER")
          NMS% = ULTREG&("MASTER")
          Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
          Call REPLA(RENUE$, COCON$, 77, 16)
          Call REPLA(RENUE$, DEDEX$, 5, 64)
          Call REPLA(RENUE$, Chr$(1), 93, 1)
          Call REPLA(RENUE$, "U.", 109, 2)
          Call REGAPP("MASTER", RENUE$)
          PUCOX$ = PUCOX$ + AJUSTI$(COCON$, 16)
          AGREGA% = 1
        End If
      End If
    End If
    '
  Wend
  Close #N1%
  '
  If AGREGA% > 0 Then
    MsgBox "ELEMENTOS AGREGADOS - CORRER CONTROL"
    Exit Sub
  End If
  '
  Dim PESTRU$(32767)
  N1% = FILEOPEN%(LUDAT$ + "ESTRUC.TXT", 1, 128)
  While Not EOF(N1%)
    Line Input #N1%, AA$
    COCON$ = TRIM$(Left$(AA$, 20))
    If COCON$ <> "" Then
      CICON% = CODINT%(COCON$)
      COELE$ = TRIM$(Mid$(AA$, 21, 20))
      If COELE$ <> "" Then
        CIELE% = CODINT%(COELE$)
        USOX = Val(Mid$(AA$, 41, 10))
        TAGRE$ = MKI$(CIELE%) + MKS$(USOX) + MKI$(0)
        For KKU% = 1 To Len(PESTRU$(CICON%)) Step 8
          If Mid$(PESTRU$(CICON%), KKU%, 8) = TAGRE$ Then
            GoTo 19
          End If
        Next KKU%
        PESTRU$(CICON%) = PESTRU$(CICON%) + MKI$(CIELE%) + MKS$(USOX) + MKI$(0)
      End If
    End If
19 Wend
  '
  JJ& = 0
  For U& = 1 To 32767
    If PESTRU$(U&) <> "" Then
      CICON% = U&
      PSTRUX$ = PESTRU$(U&) + String$(376, 0)
      A1$ = MKI$(CICON%) + Left$(PSTRUX$, 120) + String$(6, 0)
      A2$ = Mid$(PSTRUX$, 121, 255) + Chr$(1)
      JJ& = JJ& + 1
      Call GRAREG("ESTRUNUE", A1$, JJ&)
      Call GRAREG("ESTRUDOS", A2$, JJ&)
    End If
  Next U&
  Call SETULTREG("ESTRUNUE", JJ&)
  Call SETULTREG("ESTRUDOS", JJ&)
  Call CIERRARCH("*.*")
  Close
  MsgBox "FIN"
      
End Sub

Private Sub Command4_Click()
   'URECORD& = Data1.Recordset.Recordcount
   'DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   '
   JJ& = 1: PUCOX$ = ""
   N1% = FILEOPEN%("F:\CLIENTES\VDR\TRAIDOS\LISCOMP.PRN", 1, 128)
   Line Input #N1%, A$
   While Not EOF(N1%)
     Line Input #N1%, A$
     CODEX$ = TRIM$(Left$(A$, 16))
     If InStr(PUCOX$, AJUSTI$(CODEX$, 16)) < 1 Then
         PUCOX$ = PUCOX$ + AJUSTI$(CODEX$, 16)
         DEDEX$ = TRIM$(Mid$(A$, 17))
         RENUE$ = REGBLAN$("MASTER")
         NMS% = JJ&
         Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
         Call REPLA(RENUE$, CODEX$, 77, 16)
         Call REPLA(RENUE$, DEDEX$, 5, 64)
         Call REPLA(RENUE$, Chr$(1), 93, 1)
         Call REPLA(RENUE$, "U.", 109, 2)
         If Len(CODEX$) >= 3 Then
           JJ& = JJ& + 1
           Call GRAREG("MASTER", RENUE$, JJ&)
           'Call ALTACODIGO(RENUE$, OKX%)
         Else
           AAA = BBB
         End If
     End If
   Wend
   Call SETULTREG("MASTER", JJ&)
   
Close
MsgBox "CORRA CONTROL"
End

   'JJ& = 1
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   For U& = 1 To URECORD&
        TTXX$ = Space$(254)
     DBGrid1.Row = U&
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     'If CODEX$ = "" Then
     '  DBGrid1.Col = 2
     '  CODEX$ = TRIM$(DBGrid1.TEXT)
     'End If
     'f CODEX$ = "" Then CODEX$ = TRIM$(Str$(DBGrid1.Value))
       'Call REPLA(TTXX$, CODEX$, 1, 16)
       '
     DBGrid1.Col = 3
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then DEDEX$ = "Artículo " + CODEX$
     '
     RENUE$ = REGBLAN$("MASTER")
     NMS% = JJ&
     Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
     Call REPLA(RENUE$, CODEX$, 77, 16)
     Call REPLA(RENUE$, DEDEX$, 5, 64)
     Call REPLA(RENUE$, Chr$(1), 93, 1)
     Call REPLA(RENUE$, "U.", 109, 2)
     If Len(CODEX$) >= 3 Then
       JJ& = JJ& + 1
       Call GRAREG("MASTER", RENUE$, JJ&)
       'Call ALTACODIGO(RENUE$, OKX%)
     Else
       AAA = BBB
     End If
19 Next U&
   '
MsgBox JJ&
Call CIERRARCH("MASTER")
   Close #N1%
   '
   Exit Sub
   '
   '
   Call CIERRARCH("*.*")
   Close: End
   '
End Sub



Private Sub Command5_Click()
   Dim RESUR$(20)
   RESUR$(1) = "'cjtos SR33'$"
   'resur$(4) = "'cjtos SR33-1'$"
   RESUR$(2) = "'repar SR33'$"
   RESUR$(3) = "'repar SR33-1'$"
   RESUR$(4) = "'repar SR33-2'$"
   RESUR$(5) = "'repar SR33-3'$"
   RESUR$(6) = "'cil. rueda-1'$"
   'resur$(10) = "'cil. rueda'$"
   RESUR$(7) = "servos$"
   '
   N1% = FILEOPEN%(LUDAT$ + "ESTRUCRE.TXT", 2, 128)
   For UCOMA% = 1 To 11
      Data2.RecordSource = RESUR$(UCOMA%)
      Data2.Refresh
      CACOL% = DBGrid2.Columns.Count
      CAFIL% = Data2.Recordset.Recordcount
      '
      DBGrid2.RowHeight = 25
      DBGrid2.Height = (CAFIL% + 50) * DBGrid2.RowHeight
      DBGrid2.Row = 8
      TTXX$ = ""
      For U& = 0 To CACOL% - 1 'Step 2
        If U& <= 255 Then
         DBGrid2.Col = U&
         TTXX$ = TTXX$ + DBGrid2.TEXT + ","
'         COCON$ = TRIM$(DBGrid2.Columns(U&).Caption)
        End If
      Next U&
   Next UCOMA%
End Sub

Private Sub Command6_Click()
Dim TTXX$(2048)
N2% = FILEOPEN%(LUDAT$ + "ESTRUC.TXT", 2, 128)
RUTA$ = "F:\CLIENTES\VDR\TRAIDOS\HOJA"
For UI% = 1 To 11
  ASEC$ = RUTA$ + TRIM$(Str$(UI%)) + ".PRN"
  N1% = FILEOPEN%(ASEC$, 1, 128)
  For KU% = 1 To 2048
    TTXX$(KU%) = ""
  Next KU%
  '
  CAREN% = 0: CAREMA% = 0
  While Not EOF(N1%)
    Line Input #N1%, A$
    A1$ = ""
    For KK1% = 1 To Len(A$)
      CARA% = Asc(Mid$(A$, KK1%, 1))
      If CARA% < 32 Or CARA% > 128 Then CARA% = 32
      A1$ = A1$ + Chr$(CARA%)
    Next KK1%
    If Len(A1$) < 1 Then A1$ = Space$(240)
    While Len(A1$) Mod 240 <> 0
      A1$ = A1$ + " "
    Wend
    If InStr(A1$, "USO") > 0 Then
      JJ& = 0
    End If
    If JJ& < 2048 Then
      JJ& = JJ& + 1
      If JJ& > CAREN% Then CAREN% = JJ&
      TTXX$(JJ&) = TTXX$(JJ&) + A1$
    End If
  Wend
  Close #N1%
  '
  ASEC$ = RUTA$ + "0" + TRIM$(Str$(UI%)) + ".PRN"
  N3% = FILEOPEN%(ASEC$, 2, 128)
  For JJ& = 1 To CAREN%
    Print #N3%, TTXX$(JJ&)
  Next JJ&
  Close #N3%
  'GoTo 39
  '
  '
  
  'For REN% = 1 To CAREN%
  For U& = 49 To Len(TTXX$(1)) Step 32
    COCON$ = TRIM$(Mid$(TTXX$(1), U&, 16))
    If COCON$ <> "" Then
      For REN% = 8 To CAREN%
        If U& <= Len(TTXX$(REN%)) Then
            USOELE = Val(Mid$(TTXX$(REN%), U& - 16, 16))
            If USOELE > 0 Then
              COELE$ = TRIM$(Mid$(TTXX$(REN%), 17, 16))
              VALOX$ = REPLACAR$(TRIM$(Mid$(TTXX$(REN%), 1, 16)), ",", ".")
              VALOUN$ = FORMATNUM$(Val(VALOX$), "F12.4")
              '
              'TTTT$ = Mid$(TTXX$(REN%), U& - 64)
              If USOELE > 20 Then
                USOELE1 = Val(Mid$(TTXX$(REN%), U& - 32, 16))
                If USOELE1 > 0 And USOELE1 < USOELE Then
                  USOELE = USOELE1
                  GoTo 15
                End If
                USOELE = 1
              End If
              '
15            TTYY$ = Space$(100)
              Call REPLA(TTYY$, COCON$, 1, 20)
              Call REPLA(TTYY$, COELE$, 21, 20)
              Call REPLA(TTYY$, Str$(USOELE), 41, 10)
              Call REPLA(TTYY$, VALOUN$, 55, 12)
              Print #N2%, TTYY$
              '
              List1.AddItem TTYY$
              List1.ListIndex = List1.ListCount - 1
              List1.Refresh
            End If
        End If
19    Next REN%
    End If
    '
29 Next U&
   Close #N1%
DoEvents
39 Next UI%
Close
MsgBox "FIN"
'
End Sub

Private Sub Command7_Click()
  N1% = FILEOPEN%(LUDAT$ + "ESTRUC.TXT", 1, 128)
  HOII% = HOY(HOS$)
  While Not EOF(N1%)
    Line Input #N1%, AA$
    COCON$ = TRIM$(Left$(AA$, 20))
    If COCON$ <> "" Then
      CICON% = CODINT%(COCON$)
      COELE$ = TRIM$(Mid$(AA$, 21, 20))
      If COELE$ <> "" Then
        CIELE% = CODINT%(COELE$)
        VUELE = Val(Mid$(AA$, 55, 12))
        If VUELE > 0 Then
          REMA& = CIELE% + 1
          XX$ = REGLEIDO$("MASTER", REMA&)
          Call REPLA(XX$, MKS$(VUELE), 111, 4)
          Call REPLA(XX$, MKI$(HOII%), 115, 2)
          Call REPLA(XX$, Chr$(1), 95, 1)
          Call GRAREG("MASTER", XX$, REMA&)
        End If
      End If
    End If
  Wend
  Call CIERRARCH("MASTER")
  
End Sub

Private Sub Command8_Click()
   FORIMCLI.Show 1
End Sub

Private Sub Command9_Click()
FORIMPRE.Show 1
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Exit Sub
   Data1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
5  OBX$ = VDEF$
   GoTo 15
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
15 VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 99
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 99
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 99
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 99
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   CODINUEB$ = AJUSTI$(CODINUE$, 15) + Chr$(96)
   For U& = 1 To NUMAS%
     XXZ$ = Mid$(REGLEIDO$("MASTER", U& + 1), 77, 16)
     If XXZ$ = CODINUEB$ Then
       Screen.MousePointer = 1
       Call MENSERR(24, "Código '" + CODINUE$ + "' figura de Baja Lógica.\Utilice Opción de Restauración Items de Baja.")
       GoTo 10
     End If
   Next U&
   '
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   NMSI% = NUMAS% + 1
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'
Sub ppp()

ALTANUE:
   DIREL$ = LCase$(TRIM$(DIREL$))
   PPXX% = InStr(DIREL$, " ")
   RESTO$ = ""
   If PPXX% > 0 Then
      DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
      RESTO$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
   End If
   PPXX% = InStr(DIREL$, "/")
   If PPXX% > 0 Then DIREL$ = TRIM$(Left$(DIREL$, PPXX%))
   If Left$(DIREL$, 1) = Chr$(39) Then
     DIREL$ = TRIM$(Mid$(DIREL$, 2))
   End If
   '
   If TRIM$(DIREL$) <> "" Then
     If InStr(DIREL$, "@") > 0 Then
       '
       POARRO% = InStr(DIREL$, "@")
       If InStr(POARRO% + 1, DIREL$, "@") > 0 Then
         GoTo 89
         'PARA EVITAR CARGAR DIRECCIONES MALAS
       End If
       '
       For KKU& = 1 To List1.ListCount
         If List1.List(KKU& - 1) = DIREL$ Then GoTo 89
       Next KKU&
       '
       If InStr(UCase$(APENO$), "MAILER-DAEMON") > 0 Then GoTo 89
       If InStr(UCase$(APENO$), "UNDISCLOSED-RECIPIEN") > 0 Then GoTo 89
       '
       List1.AddItem DIREL$
       RENUE$ = REGBLAN$("LIDIRMAI")
       '
       PPXX% = InStr(APENO$, "@")
       If PPXX% > 0 Then APENO$ = TRIM$(Left$(APENO$, PPXX% - 1))
       If APENO$ <> "" Then
         APENO$ = LCase$(APENO$)
         APENO$ = REPLACAR$(APENO$, "_", " ")
         '
         Call REPLA(APENO$, UCase$(Left$(APENO$, 1)), 1, 1)
         For KKI% = 2 To Len(APENO$)
           If Asc(Mid$(APENO$, KKI% - 1, 1)) < 48 Then
             Call REPLA(APENO$, UCase$(Mid$(APENO$, KKI%, 1)), KKI%, 1)
           End If
         Next KKI%
       End If
       '
'MsgBox APENO$ + Chr$(10) + Chr$(13) + EMPRA$ + Chr$(10) + Chr$(13) + DIREL$
       JJ& = 1 + ULTREG("LIDIRMAI")
       Call REPLA(RENUE$, MKS$(JJ&), 1, 4)
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       '
       Call REPLA(RENUE$, MUESTRA$, 5, 60)
       Call REPLA(RENUE$, APENO$, 65, 40)
       Call REPLA(RENUE$, DIREL$, 105, 60)
       Call REPLA(RENUE$, APELA$, 165, 8)
       Call REPLA(RENUE$, EMPRA$, 173, 64)
       Call REPLA(RENUE$, ORIGE$, 237, 16)
       Call REPLA(RENUE$, MKI$(HOII%), 253, 2)
       Call REGAPP("LIDIRMAI", RENUE$)
       '
     End If
   End If
   '
89 Return
'
End Sub

Private Sub Label1_DblClick(Index As Integer)
   'Data1.DatabaseName = Label1(Index).Caption
   'Data1.RecordSource = Label2(Index).Caption
   'Data1.Refresh
   'Label3(Index).Caption = TRIM$(Str$(Data1.Recordset.Recordcount))
End Sub

Sub ALTACODIGO1(VDEF$, OKX%)
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   NMSI% = NUMAS% + 1
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'


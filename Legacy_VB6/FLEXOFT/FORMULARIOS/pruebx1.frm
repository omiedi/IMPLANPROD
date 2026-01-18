VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PRUEBX1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Producción"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   435
      Left            =   3990
      TabIndex        =   5
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\ANOVE\Ajorna.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "BASE$"
      Top             =   1155
      Width           =   1905
   End
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
      Picture         =   "pruebx1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Genera Plan Maestro de Entregas"
      Top             =   0
      Width           =   560
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
      Left            =   3885
      Picture         =   "pruebx1.frx":1366
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "pruebx1.frx":14B0
      Height          =   5790
      Left            =   0
      OleObjectBlob   =   "pruebx1.frx":14C0
      TabIndex        =   2
      Top             =   2310
      Width           =   11565
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   330
      Left            =   6930
      TabIndex        =   6
      Top             =   1365
      Width           =   3900
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   330
      Left            =   6930
      TabIndex        =   4
      Top             =   840
      Width           =   3900
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   330
      Left            =   6930
      TabIndex        =   3
      Top             =   210
      Width           =   3900
   End
End
Attribute VB_Name = "PRUEBX1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
   DBGrid1.RowHeight = 1
   DBGrid1.Height = 1000 * DBGrid1.RowHeight
   JJ& = 5513
   For RXX = 0 To 200
      DBGrid1.Row = RXX
      DBGrid1.Col = 1
      Label1 = TRIM$(DBGrid1.Text)
      DBGrid1.Col = 2
      TTXX$ = Label1 + " " + DBGrid1.Text
      Label1 = TTXX$
      DBGrid1.Col = 7
      Label2 = DBGrid1.Text
      DBGrid1.Col = 8
      Label3 = DBGrid1.Text
      If TRIM$(Label1) <> "" Then
        'If TRIM$(Label2) <> "" Then
          If TRIM$(Label3) <> "" Then
             RENUE$ = REGBLAN$("LIDIRMAI")
             APENO$ = LCase$(Label1)
             Call REPLA(APENO$, UCase$(Left$(APENO$, 1)), 1, 1)
             For KKI% = 2 To Len(APENO$)
               If Asc(Mid$(APENO$, KKI% - 1, 1)) < 48 Then
                 Call REPLA(APENO$, UCase$(Mid$(APENO$, KKI%, 1)), KKI%, 1)
               End If
             Next KKI%
             '
             EMPRA$ = ""
             DIREL$ = Label3
             ORIGE$ = "AOTS: CONF-UCES"
'MsgBox APENO$ + Chr$(10) + Chr$(13) + EMPRA$ + Chr$(10) + Chr$(13) + DIREL$
             JJ& = JJ& + 1
             Call REPLA(RENUE$, MKS$(JJ&), 1, 4)
             MUESTRA$ = ""
             If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
             If TRIM$(EMPRA$) <> "" Then
              If MUESTRA$ <> "" Then
                 MUESTRA$ = MUESTRA$ + " -"
               End If
               MUESTRA$ = MUESTRA$ + " " + EMPRA$
             End If
             MUESTRA$ = MUESTRA$ + " (" + Label2 + ")"
             '
             Call REPLA(RENUE$, MUESTRA$, 5, 60)
             Call REPLA(RENUE$, APENO$, 65, 40)
             Call REPLA(RENUE$, DIREL$, 105, 60)
             Call REPLA(RENUE$, APELA$, 165, 8)
             Call REPLA(RENUE$, EMPRA$, 173, 64)
             Call REPLA(RENUE$, ORIGE$, 237, 16)
             Call REPLA(RENUE$, MKI$(HOII%), 253, 2)
             Call GRAREG("LIDIRMAI", RENUE$, JJ&)
             '
          End If
        'End If
      End If
      
   Next RXX
   Call CIERRARCH("*.*")
   Close: End
   '
End Sub

Private Sub DBGrid1_DblClick()
   For RXX = 0 To 250
   DBGrid1.Row = RXX
   DBGrid1.Col = 1
   Label1 = TRIM$(DBGrid1.Text)
   DBGrid1.Col = 2
   TTXX$ = DBGrid1.Text + " "
   DBGrid1.Col = 4
   Label2 = TTXX$
   If Len(DBGrid1.Text) > 2 Then Label2 = Label2 + DBGrid1.Text
   '
   If Label1 <> "" Then
     If Left$(Label1, 1) = "0" Then
     AAA = BBB
'VDEF$ = REGBLAN$("MASTER")
'Call REPLA(VDEF$, Label1, 77, 16)
'Call REPLA(VDEF$, Label2, 5, 64)
'Call REPLA(VDEF$, Chr$(1), 93, 1)
'Call REPLA(VDEF$, "U.", 109, 2)
'Call ALTACODIGO(VDEF$, OKX%)
End If
End If
   Next RXX
End Sub

Private Sub Form_Load()
   Data1.Refresh
   Data1.Recordset.MOVELAST
   
   AA1 = Data1.Recordset.Fields.Count
   Data1.Recordset.MOVEFIRST
   'Data1.Recordset.MOVENEXT
   'Data1.Recordset.MOVENEXT
   'Data1.Recordset.MOVENEXT
   'Data1.Recordset.MOVENEXT
   
   'TTXX$ = ""
   'For KKI% = 0 To AA1 - 1
   '  AA2 = Data1.Recordset(KKI%).Type
   '  AA3 = VarType(Data1.Recordset(KKI%).Value)
   '   TTXX$ = TTXX$ & "-" & AA3
     '.DefaultValue
   'Next KKI%

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


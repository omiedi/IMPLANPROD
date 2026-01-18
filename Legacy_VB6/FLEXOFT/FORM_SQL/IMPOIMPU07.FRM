VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPOIMPU07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Pedidos Desde Planillas Excel"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
      Height          =   2535
      Left            =   2100
      TabIndex        =   12
      Top             =   2205
      Visible         =   0   'False
      Width           =   8835
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   105
      TabIndex        =   8
      Top             =   1155
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPOIMPU07.frx":0000
         Height          =   3750
         Left            =   150
         OleObjectBlob   =   "IMPOIMPU07.frx":0010
         TabIndex        =   9
         Top             =   360
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11025
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   210
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   10815
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   6060
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   11
            Top             =   15
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   7470
         TabIndex        =   3
         Text            =   " "
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton Command22 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7035
         TabIndex        =   2
         Top             =   465
         Width           =   220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7110
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
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
      Height          =   645
      Left            =   11025
      Picture         =   "IMPOIMPU07.frx":09E3
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   420
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2200
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "IMPOIMPU07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
End Sub

Private Sub Combo1_Click()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Recordset.MoveFirst
   Data1.Refresh
   '
End Sub


Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.filename)
   Data1.DatabaseName = Label1
   Combo1.Clear
   Combo1.TEXT = ""
   On Error Resume Next
   Data1.Refresh
   '
   CATAB% = Data1.Database.TableDefs.Count
   For TTB% = 1 To CATAB%
     Combo1.AddItem Data1.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo1.TEXT = Data1.Database.TableDefs(TTB% - 1).Name
     End If
   Next TTB%
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   'Frame5.Visible = True
   List1.Clear
   '
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   Data1.Recordset.MoveFirst
   '
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   '
   If DBGrid1.Columns.Count < 6 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     '
     CY$ = ""
     For KKI% = 0 To 2
        DBGrid1.Col = 2 * KKI%
        CX$ = TRIM$(DBGrid1.TEXT)
        If CX$ <> "" Then
          While Len(CX$) < 2: CX$ = "0" + CX$: Wend
          DBGrid1.Col = 2 * KKI% + 1
          DX$ = TRIM$(DBGrid1.TEXT)
          CY$ = CY$ + CX$
          QX$ = ""
          For KKJ% = DBGrid1.Col + 1 To 5
             DBGrid1.Col = KKJ%
             If TRIM$(DBGrid1.TEXT) <> "" Then GoTo 15
          Next KKJ%
          '
          DBGrid1.Col = 6
          QX$ = TRIM$(DBGrid1.TEXT)
          '
15        List1.AddItem AJUSTI$(CY$, 16) + AJUSTI$(DX$, 32) + QX$
          RFF$ = RECFILT$("IMPUSOL", 1, CY$)
          TTXX$ = AJUSTI$(CY$, 12) + AJUSTI$(DX$, 64) + AJUSTI$(QX$, 12)
          If RFF$ <> "" Then
              Call FILTERPUTRECORD("IMPUSOL", 1, CY$, TTXX$)
            Else
              Call REGAPP("IMPUSOL", TTXX$)
          End If
        End If
          
     Next KKI%
     '
19 Next U&

20 Screen.MousePointer = 1
   '
   Call CIERRARCH("IMPUSOL")
   Call COMUNI("Se Ha Completado la Pre-Importación.")
   '
 DBGrid1.Height = DBHANT
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   Call CARGADET
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Sub CARGADET()
   Label1 = ""
   Combo1.Clear
   DBGrid1.ClearFields
End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub


Private Sub Text10_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text2_gotfocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text3_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text4_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text5_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Sub PREIMPORMAS()
   '
   If DBGrid1.Columns.Count < 10 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(254)
        CODEX$ = ""
        On Error Resume Next
        CODEX$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        On Error GoTo 0
        If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
        CDXI% = CODINT%(CODEX$)
        If CDXI% > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      Call REPLA(TTXX$, CODEX$, 1, 16)
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
        If Len(DEDEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DEDEX$, 25, 64)
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(2), 4)
        On Error GoTo 0
        If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 91, 4)
        '
        UME$ = "U."
        On Error Resume Next
        UME$ = TRIM$(.Recordset.Fields(3))
        On Error GoTo 0
        Call REPLA(TTXX$, UME$, 97, 4)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(4))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 101, 12)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(5))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 113, 12)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(6))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.4")
        Call REPLA(TTXX$, VALO$, 125, 12)
        '
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(7), 4)
        On Error GoTo 0
         If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 137, 4)
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = Format$(.Recordset.Fields(8), "dd/mm/yy")
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 143, 8)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(9))
        On Error GoTo 0
        If VALOR < 0 Or VALOR > 32766 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F8.0")
        Call REPLA(TTXX$, VALO$, 151, 8)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(10))
        On Error GoTo 0
        If VALOR < (-9999999) Or VALOR > 9999999 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F10.1")
        Call REPLA(TTXX$, VALO$, 161, 10)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORSTRU()
   '
   If DBGrid1.Columns.Count < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     '
     DBGrid1.Col = 1
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 19, 16)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 37, 32)
     '
     DBGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 69, 12)
     '
     DBGrid1.Col = 3                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 81, 4)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORUTS()
   '
   If DBGrid1.Columns.Count < 12 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     If CODEX$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = CODEX$
     '
     DBGrid1.Col = 1
     DEOPE$ = TRIM$(DBGrid1.TEXT)
     If DEOPE$ = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     NOPE% = NOPE% + 10
     VALO$ = FORMATNUM$(NOPE%, "I4")
     Call REPLA(TTXX$, VALO$, 17, 4)
     Call REPLA(TTXX$, DEOPE$, 23, 48)
     '
     DBGrid1.Col = 2                     ' MAQUINA
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 73, 6)
     '
     DBGrid1.Col = 3                     ' HORAS PREPARACION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 79, 10)
     '
     DBGrid1.Col = 4                     ' HORAS STANDARD FIJO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 89, 10)
     '
     DBGrid1.Col = 5                     ' MINUTOS STANDARD VARIABLE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.5")
     Call REPLA(TTXX$, VALO$, 99, 12)
     '
     DBGrid1.Col = 6                     ' LOTE STANDARD UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 1 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F10.1")
     Call REPLA(TTXX$, VALO$, 111, 10)
     '
     DBGrid1.Col = 7                     ' % DEM.DISTRIBUIDA
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     If VALOR < 0 Then VALOR = 0
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F8.1")
     Call REPLA(TTXX$, VALO$, 121, 8)
     '
     DBGrid1.Col = 8                     ' OPERARIOS
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0.01 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F8.2")
     Call REPLA(TTXX$, VALO$, 129, 8)
     '
     DBGrid1.Col = 9                     ' MAQUINA AUX 1
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 139, 6)
     '
     DBGrid1.Col = 10                     ' MAQUINA AUX 2
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 147, 6)
     '
     DBGrid1.Col = 11                     ' MAQUINA AUX 3
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 155, 6)
     '
     DBGrid1.Col = 12                     ' H.P.FABR.
     HOPROFA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, HOPROFA$, 163, 8)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 9, 30)
     '
     DBGrid1.Col = 2                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 999999.99 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 41, 12)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(DBGrid1.TEXT)
     If VTX$ <> "" And XVALO(VTX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(VTX$)
     On Error GoTo 0
     If VALOR > 999.999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 53, 10)
     '
     DBGrid1.Col = 4                    ' Maquina equivalente
     CODEQ$ = TRIM$(DBGrid1.TEXT)
     If CODEQ$ <> "" Then
       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CODEQ$, 66, 6)
     '
     DBGrid1.Col = 5                     ' Coeficiente de equivalencia
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
     If CODEQ$ = "" Then VALOR = 0
     If VALOR < VAMIN Or VALOR > 999.9 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     Call REPLA(TTXX$, VALO$, 72, 12)
     '
     DBGrid1.Col = 6                    ' Observaciones
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   
20 End Sub

Sub PREIMPORCAT()
   '
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   If URECORD& > 255 Then
     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
     GoTo 20
   End If
   '
   If ULTREG&("CATOPER") > 0 Then
     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
       GoTo 20
     End If
   End If
   '
   JJ& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     DECAT$ = TRIM$(DBGrid1.TEXT)
     If DECAT$ <> "" Then
       JJ& = JJ& + 1
       COCAT$ = AJUSTD$(Str$(JJ&), 4)
       Call REPLA(TTXX$, COCAT$, 1, 4)
       Call REPLA(TTXX$, DECAT$, 7, 30)
       '
       DBGrid1.Col = 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       If VALOR < 0 Or VALOR > 999999.99 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
         List1.AddItem MERROX$
         GoTo 19
       End If
       VALO$ = FORMATNUM$(VALOR, "F12.4")
       Call REPLA(TTXX$, VALO$, 37, 12)
       '
       DBGrid1.Col = 2
       OBSER$ = TRIM$(DBGrid1.TEXT)
       Call REPLA(TTXX$, OBSER$, 51, 32)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&
   
20 End Sub

Sub PREIMPORCLI()
   '
   If DBGrid1.Columns.Count < 11 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   HROW = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(284)
     DBGrid1.Row = U& - 1
     '
     DBGrid1.Col = 0
     NUCLIE = XVALO(TRIM$(DBGrid1.TEXT))
     If NUCLIE < 1 Or NUCLIE > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cliente no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     CODEX$ = AJUSTD$(Str$(Int(NUCLIE + 0.5)), 5)
12   Call REPLA(TTXX$, CODEX$, 1, 5)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social de Cliente"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 64)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 48)
     Call REPLA(TTXX$, DOMI$, 55, 48)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 12)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 32)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DBGrid1.Col = 9                     ' GRUPO DE CLIENTE
     GRUPOCLI$ = AJUSTI$(DBGrid1.TEXT, 4)
     Call REPLA(TTXX$, GRUPOCLI$, 236, 4)
     '
     DBGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODICUEN$, 241, 12)
     '
     DBGrid1.Col = 11                    ' VENDEDOR
     VENDE$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, VENDE$, 253, 4)
     '
     DBGrid1.Col = 12                    ' PROVINCIA
     PROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, PROVI$, 259, 2)
     '
     DBGrid1.Col = 13                    ' CATEGORIA IVA
     CIVA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, CIVA$, 261, 4)
     '
     DBGrid1.Col = 14                    ' CONDICION DE PAGO
     COPA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, COPA$, 265, 4)
     '
     DBGrid1.Col = 15                    ' LIMITE DE CREDITO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Then VALOR = 0
     VALO$ = FORMATNUM$(VALOR, "F12.2")
     Call REPLA(TTXX$, VALO$, 269, 12)
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   '
   DBGrid1.RowHeight = HROW   ' RESTAURA ALTURA ANTERIOR
   '
End Sub

Sub PREIMPORPRO()
   '
   If DBGrid1.Columns.Count < 15 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(290)
     DBGrid1.Row = U& - 1
     '
     DBGrid1.Col = 0
     NUPROV = XVALO(TRIM$(DBGrid1.TEXT))   ' NUMERO DE CUENTA
     If NUPROV < 1 Or NUPROV > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     CODEX$ = AJUSTD$(Str$(Int(NUPROV + 0.5)), 5)
12   Call REPLA(TTXX$, CODEX$, 1, 5)
     '
     DBGrid1.Col = 1                     ' RAZON SOCIAL
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social del Proveedor"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 48)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 64)
     Call REPLA(TTXX$, DOMI$, 55, 64)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 16)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 48)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DBGrid1.Col = 9                     ' TIPO DE CUENTA
     TIPOCTA$ = TRIM$(DBGrid1.TEXT)
     If TIPOCTA$ = "0" Or TIPOCTA$ = "1" Then
       Call REPLA(TTXX$, TIPOCTA$, 235, 4)
     Else
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Tipo de Cuenta"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     DBGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODICUEN$, 240, 12)
     '
     '\\\OT-05-0202-WAR-18/04/05///
     DBGrid1.Col = 11                    ' CODIGO DE PROVINCIA
     CODIPROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODIPROVI$, 253, 2)
     '
     DBGrid1.Col = 12                    ' CODIGO DE POSICION IVA
     CODIVA% = XVALO(DBGrid1.TEXT)
     Call REPLA(TTXX$, FORMATNUM$(CODIVA%, "I4"), 255, 4)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 13                    ' CUENTA MADRE
     CTAMADRE$ = TRIM$(DBGrid1.TEXT)
     If Len(CTAMADRE$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Longitud de Cuenta Madre"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CTAMADRE, 261, 12)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 14                    ' numbre de fantasia
     FANTASI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FANTASI$, 275, 16)
     '
    '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORPLA()
   '
   If DBGrid1.Columns.Count < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   HROW = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     '
     NIVEI% = XVALO(DBGrid1.TEXT)
     If NIVEI% < 0 Or NIVEI% > 5 Then NIVEI% = 0
     Call REPLA(TTXX$, AJUSTD$(Str$(NIVEI%), 4), 1, 4)
     '
     DBGrid1.Col = 1
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CODEX$, 7, 12)
     '
     For KKU% = 1 To List4.ListCount
       If TRIM$(Mid$(List4.List(KKU% - 1), 7, 12)) = CODEX$ Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Duplicado o Pre-Existente"
         List1.AddItem MERROX$
         GoTo 19
       End If
     Next KKU%
     '
     DBGrid1.Col = 2
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 21, 48)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ <> "D" And VTX$ <> "H" And VTX$ <> "I" Then VTX$ = "G"
     Call REPLA(TTXX$, VTX$, 71, 1)
     '
     DBGrid1.Col = 4
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ = "0" Or VTX$ = "NO" Then VTX$ = "NO"
     If VTX$ <> "NO" Then VTX$ = "SI"
     Call REPLA(TTXX$, VTX$, 75, 2)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   DBGrid1.RowHeight = HROW   ' RESTAURA ALTURA ANTERIOR
   '
20 End Sub
 '\\\OT-05-0207-WAR-21/04/05///
Sub PREIMPOPED()
   '
End Sub
Sub IMPORPED()
   Dim FORMORIG As Form
   Set FORMORIG = FORPED07
   If PROLLAMA = "ABIE" Then Set FORMORIG = PEDABIE07
    With FORMORIG
        Call .CARDACLI
  'CARGO EL DETALLE
  Screen.MousePointer = 11
  '
  REBLA$ = REGBLAN$("!CARDETPE")
  J& = 0
  For KI% = 1 To List4.ListCount
    TTXX$ = List4.List(KI% - 1)
    NCI2$ = (Mid$(TTXX$, 1, 16)) 'Nº ARTICULO
    NCI1% = CODINT(NCI2$)
    JJD$ = Mid$(TTXX$, 114, 12)
    CA1# = XVALO(Mid$(TTXX$, 114, 12))  'CANTIDAD
    FE% = FECHANUM(Mid$(TTXX$, 128, 8)) 'FECHA
    PU1# = XVALO(Mid$(TTXX$, 136, 14)) 'PRECIO UNITARIO
    Call REPLA(Z$, MKI(NCI1%), 1, 2)
    Call REPLA(Z$, MKS$(CA1#), 47, 4)
    Call REPLA(Z$, MKI$(FE%), 63, 2)
    Call REPLA(Z$, MKS$(PU1#), 39, 4)
    J& = J& + 1
    Call GRAREG("!CARDETPE", Z$, J&)
  Next KI%
  Call SETULTREG("!CARDETPE", J&)
  '
  On Error Resume Next
  Call .CARLISPE
  .MARCONOTA.Visible = False
  .MARSELPRESU.Visible = False
  Call .CALTOPED
  Call .CALPERIB
  End With
  On Error GoTo 0
  IMPPED07.Hide
  Screen.MousePointer = 1
  '
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FORACPLAN07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Action Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   105
      TabIndex        =   20
      Top             =   5145
      Width           =   11655
      Begin VB.CommandButton Command2 
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
         Left            =   7920
         TabIndex        =   24
         Top             =   315
         Width           =   220
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   23
         Text            =   " "
         Top             =   315
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   945
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   21
         Top             =   315
         Width           =   6975
         Begin VB.Label Label5 
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
            TabIndex        =   22
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.Label Label7 
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
         TabIndex        =   27
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Left            =   8160
         TabIndex        =   26
         Top             =   135
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   1785
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Formato de Importacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   4800
      TabIndex        =   7
      Top             =   105
      Width           =   6100
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   11
         Top             =   350
         Width           =   5655
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Col. Dato / Variable            Tipo     Rango"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   12
            Top             =   0
            Width           =   5655
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Uso por Vehículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   105
      TabIndex        =   2
      Top             =   1995
      Width           =   11655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   945
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   14
         Top             =   315
         Width           =   6975
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
            TabIndex        =   15
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   4
         Text            =   " "
         Top             =   315
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
         Left            =   7920
         TabIndex        =   3
         Top             =   315
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
         Left            =   8160
         TabIndex        =   6
         Top             =   135
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
         TabIndex        =   5
         Top             =   375
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
      Picture         =   "FORACPLAN07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   840
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   0
      Visible         =   0   'False
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
      Height          =   645
      Left            =   11025
      Picture         =   "FORACPLAN07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "(MRP - II)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1695
      Left            =   1680
      TabIndex        =   18
      Top             =   1365
      Width           =   4005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "de Action Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   1050
      TabIndex        =   17
      Top             =   735
      Width           =   4005
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Importador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   315
      TabIndex        =   16
      Top             =   105
      Width           =   4005
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FORACPLAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANCELIMP%
Sub PREIMPINVSTO()
   '
   '03/04/07 (OT-07-0141)
   If DBGrid1.Columns.Count < 5 Then
      Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
      GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   HROW = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   ATX% = 0
   HUBOAMAS% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(120)
        '
        codex$ = ""                             ' CODIGO DE ARTICULO
        On Error Resume Next
        codex$ = .Recordset.Fields(0)
        On Error GoTo 0
        '
5       CDXI% = CODINT%(codex$)
        If CDXI% < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código '" + TRIM$(codex$) + "' No Existente en el Maestro de Artículos"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
12      codex$ = CODEXT$(CDXI%)
        Call REPLA(TTXX$, codex$, 1, 16) 'CODIGO DEL ARTICULO
        '
        Call REPLA(TTXX$, DESCRIT0$(CDXI%), 19, 48) ' DESCRIPCION DEL ARTICULO
        '
        FECHMOVX$ = ""
        On Error Resume Next
        FECH% = CVINT(.Recordset.Fields(2))
        FECHMOVX$ = FECHATEX$(FECH%) 'FECHA
        On Error GoTo 0
        If FECHMOVX$ = "" Or FECHANUM(FECHMOVX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Fecha No Válida"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, FECHMOVX$, 67, 8)
        '
        DEPX% = 1 ' DEPóSITO
        On Error Resume Next
        DEPX% = CDbl(.Recordset.Fields(3))
        On Error GoTo 0
        If DEPX% < 1 Or DEPX% > 255 Then DEPX% = 1
        VALO$ = FORMATNUM$(DEPX%, "F4.0")
        Call REPLA(TTXX$, VALO$, 75, 4)
        '
        VALOR = 0 'STOCK TEORICO
        VALOR = STODEPFECH(CDXI%, DEPX%, FECH%) '04/04/07 (OT-07-0141)
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        Call REPLA(TTXX$, VALO$, 79, 12)
        '
        VALOR = 0 'STOCK AJUSTADO
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(4))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Ajustado No Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 91, 12)
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
20 DBGrid1.RowHeight = HROW
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.TEXT  ' "272W(407L)PART LIST"
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

Private Sub Combo2_Change()
   '
   Data2.RecordSource = Combo2.TEXT  ' "272W(407L)PART LIST"
   Data2.Refresh
   Data2.Recordset.MoveFirst
   '
End Sub

Private Sub Combo2_Click()
   '
   Data2.RecordSource = Combo2.TEXT
   Data2.Recordset.MoveFirst
   Data2.Refresh
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Action Plan"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label5 = UCase$(CommonDialog1.FileName)
   Data2.DatabaseName = Label5
   Combo2.Clear
   Combo2.TEXT = ""
   On Error Resume Next
   Data2.Refresh
   '
   CATAB% = Data2.Database.TableDefs.Count
   For TTB% = 1 To CATAB%
     Combo2.AddItem Data2.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo2.TEXT = Data2.Database.TableDefs(TTB% - 1).Name
     End If
   Next TTB%
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.FileName)
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

Private Sub Command3_Click()
   Command3.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Printer.Print "Registros no Procesables"
     Printer.Print " "
     For U& = 1 To List1.ListCount
       Printer.Print List1.List(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   '
   REBLA$ = REGBLAN$("PROGENT")
   DBGrid1.Height = 300 * DBGrid1.RowHeight
   DBGrid2.Height = 300 * DBGrid2.RowHeight
   For CL2% = 4 To 13 Step 2
     II& = CL2%
     Call TRACE(20, II&, 13)
     For RG2% = 17 To 41
       DBGrid2.Row = RG2%
       DBGrid2.Col = CL2%
       CANPRO = XVALO(DBGrid2.TEXT)
       DBGrid2.Col = CL2% + 1
       CANPRO = CANPRO + XVALO(DBGrid2.TEXT)
       FEPROGRA$ = TRIM$(FORMATNUM$(CL2% / 2, "I2"))
       While Len(FEPROGRA$) < 2: FEPROGRA$ = "0" + FEPROGRA$: Wend
       FEPROGRA$ = "10/" + FEPROGRA$ + "/09"
       DBGrid2.Col = 3
       MODELOX$ = DBGrid2.TEXT
       '
       MODELOA$ = ""
       For CL1% = 12 To 76
         DBGrid1.Row = 6
         DBGrid1.Col = CL1%
         If DBGrid1.TEXT = MODELOA$ Then GoTo 39
         MODELOA$ = DBGrid1.TEXT
         '
         If DBGrid1.TEXT = MODELOX$ Then
            For RG1% = 9 To 64
               DBGrid1.Row = RG1%
               DBGrid1.Col = CL1%
               USO = XVALO(DBGrid1.TEXT)
               If USO > 0 Then
                  DBGrid1.Col = 4
                  codex$ = DBGrid1.TEXT
                  CANPROGRA = USO * CANPRO
                  '
                  REGRA$ = REBLA$
                  CIXI% = CODINT%(codex$)
                  If CIXI% > 0 Then
                    Call REPLA(REGRA$, MKI$(CIXI%), 1, 2)
                    FEXI% = FECHANUM(FEPROGRA$)
                    Call REPLA(REGRA$, MKI$(FEXI%), 3, 2)
                    Call REPLA(REGRA$, MKS$(CANPROGRA), 7, 4)
 If codex$ = "53711-0K031" Then
   If CANPROGRA > 0 Then
     SUPROGRA = SUPROGRA + CANPROGRA
     AAA = BBB
   End If
 End If
                    For I& = 1 To JJ&
                       XYX$ = Mid$(REGRA$, 1, 4) ' VALOR NUEVO
                       YXY$ = Mid$(REGLEIDO$("PROGENT", I&), 1, 4) ' VALOR EXISTENTE EN ARCHIVO
                       CC = 0
                       If XYX$ = YXY$ Then
                         CANPROGRA = CANPROGRA + CVS(Mid$(REGLEIDO$("PROGENT", I&), 7, 4))
                         Call REPLA(REGRA$, MKS$(CANPROGRA), 7, 4)
                         Call GRAREG("PROGENT", REGRA$, I&)
                         CC = 1: Exit For
                       End If
                    Next I&
                    If CC = 0 Then
                       JJ& = JJ& + 1
                       Call GRAREG("PROGENT", REGRA$, JJ&)
                    End If
                  End If
                  '
               End If
29          Next RG1%
         End If
         '
39     Next CL1%
     Next RG2%
   Next CL2%
   '
   Call SETULTREG("PROGENT", JJ&)
   Call FILESORT("PROGENT", "", 1, 3, "ASC")
   Call COMUNI("Conversion Completada")
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub



Private Sub Form_Load()
   '03/04/07 (OT-07-0141)
   Call CENTRAFORM(Me)
   '
   List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
   List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
   List2.AddItem "  3  Fecha                      Date"
   List2.AddItem "  4  Depósito Número            Num      1 - 255"
   List2.AddItem "  5  Stock Ajustado             Num      +/- 999999"
   Label8 = "Código            Descripción                                     Fecha     Dep     S.Teo.     S.Ajus."
   '03/04/07 (FIN)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
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

Sub IMPINVSTO()
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE CodiMovi = 'AJ'"
   SQLX$ = SQLX$ + "ORDER BY VAL(SUBSTRING(DoPriCor,4,)) DESC"

   UNUM& = 0
   'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim MoviTemp As ADODB.Recordset
   Set MoviTemp = New ADODB.Recordset
   MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   On Error Resume Next
   MoviTemp.MoveFirst
   If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
   End If
   On Error GoTo 0
   NRO& = 1 + UNUM&
   MoviTemp.Close
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NOCLOSE% = 1
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = CODINT%(CODIT$)
     FEMO% = FECHANUM(Mid$(TTXX$, 67, 8))
     LOTE$ = ""
     CMOV$ = "AJ"
     DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
     DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
     DOPRILA$ = DOPRICO$
     DOSELA$ = DOSECO$
     REFE$ = "Importación Inventario al " + FECHATEX$(HOY(HO$))
     NUORIT% = 0
     VUNI# = 0
     MONEII% = 0
     NC% = 0
     DEPO% = XVALO(Mid$(TTXX$, 75, 4))
     CANTI = XVALO(Mid$(TTXX$, 91, 12)) - XVALO(Mid$(TTXX$, 79, 12))
     '
     If Abs(CANTI) >= 0.005 Then
        Call MOVISTOK(FEMO%, CIXI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
        NRO& = NRO& + 1
     End If
     '
   Next U&
   '
   Movisto.Close
   NOCLOSE% = 0
   Screen.MousePointer = 1
   '
End Sub


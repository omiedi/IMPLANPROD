VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPMOCON 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Importación de Movimientos Contables"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Progreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   210
      TabIndex        =   15
      Top             =   1155
      Width           =   2430
      Begin VB.Label Label7 
         Height          =   225
         Left            =   1260
         TabIndex        =   18
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "DE"
         Height          =   225
         Left            =   840
         TabIndex        =   17
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label5 
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "PLANILLA ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4770
      Left            =   105
      TabIndex        =   8
      Top             =   2250
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPMOCON.frx":0000
         Height          =   4380
         Left            =   105
         OleObjectBlob   =   "IMPMOCON.frx":0010
         TabIndex        =   9
         Top             =   315
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   2250
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H8000000E&
      Height          =   2130
      Left            =   2940
      TabIndex        =   2
      Top             =   105
      Width           =   7875
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   6090
         TabIndex        =   20
         Top             =   1680
         Width           =   1590
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   4815
         TabIndex        =   11
         Top             =   1110
         Width           =   4875
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "contab_sin_subst_al_31-10-2004.xls"
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
            TabIndex        =   12
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   945
         TabIndex        =   3
         Text            =   "0UOYKCUF"
         Top             =   1680
         Width           =   4890
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Diferencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5985
         TabIndex        =   19
         Top             =   1260
         Width           =   1680
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\FLGUIDI\TRAIDOS\MIGRACION NUCLEO\"
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
         Left            =   945
         TabIndex        =   14
         Top             =   420
         Width           =   6840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   105
         TabIndex        =   13
         Top             =   420
         Width           =   735
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   105
         TabIndex        =   5
         Top             =   1785
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1110
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
      Height          =   750
      Left            =   11025
      Picture         =   "IMPMOCON.frx":09C2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\FLGUIDI\TRAIDOS\migracion nucleo\contab_sin_subst_al_31-10-2004.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "'0UOYKCUF'$"
      Top             =   105
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
      Height          =   750
      Left            =   11025
      Picture         =   "IMPMOCON.frx":0CCC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   105
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPMOCON"
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

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'
'Private Sub Command2_Click()
'   Command2.Enabled = False
'   Screen.MousePointer = 11
'   Select Case Val(NARIM)
'     Case 0
'       Call IMPORMAS
'     Case 1
'       Call IMPORSTRU
'     Case 2
'       Call IMPORUTS
'     Case 3
'       Call IMPORMAQ
'     Case 4
'       Call IMPORCAT
'   End Select
'   Screen.MousePointer = 1
'   Call COMUNI("Se ha Completado la Importacion Definitiva")
'   Frame5.Visible = False
'   DoEvents
'   Command2.Enabled = True
'End Sub

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
   '
        Call IMPMOVCONT
   '
   MsgBox "Fin Importación"
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   '
   Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call FINAL("")
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub Option1_Click(Index As Integer)
   Label1 = ""
   Combo1.Clear
   DBGrid1.ClearFields
   NARIM = TRIM$(Str$(Index))
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
'
Sub IMPMOVCONT()
   '
   CUEDIFE$ = TRIM$(Text1)
   If CUEDIFE$ = "" Or CUEINT%(CUEDIFE$) < 1 Then
     Call MENSERR(24, "Falta Cuenta diferencia")
     Exit Sub
   End If
   '
   Set ENCASIEN1 = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   Set DETASIEN1 = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   On Error Resume Next
   ENCASIEN1.MoveLast
   While ENCASIEN1.RecordCount > 0
     ENCASIEN1.MoveLast
     ENCASIEN1.Delete
   Wend
   '
   DETASIEN1.MoveLast
   While DETASIEN1.RecordCount > 0
     DETASIEN1.MoveLast
     DETASIEN1.Delete
   Wend
   On Error GoTo 0
   '
   ULOI$ = Left$(LUDAT$, 2)
   '
   TEPRUE = ""
   Data1.DatabaseName = ULOI$ + Label9 + Label1
   Data1.RecordSource = "'0UOYKCUF'$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   Label7 = URECORD&
   Label7.Refresh
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   FLIMIT% = FECHANUM("31/12/00")
   '
6  ENCASIEN1.Close
   DETASIEN1.Close
   Set ENCASIEN1 = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   Set DETASIEN1 = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   '
   With Data1
    .Recordset.MoveFirst
    U& = 0
    QREN& = 1
    J = 1
    Do While Not .Recordset.EOF
        'LECTURA
        FEASIEN = .Recordset.Fields(0) 'FECHA DEL ASIENTO
          FEINAS% = CVINT(FEASIEN)
        NUMASIEN& = Val(.Recordset.Fields(1)) 'NUMERO ASIENTO
        If NUMASI$ = "" Then
          NUMASI$ = NUPROA$(FEINAS%)
        End If
        RFASIEN$ = TRIM$(.Recordset.Fields(2)) 'REFERENCIA ASIENTO
        CODIGCUE$ = TRIM$(.Recordset.Fields(3)) 'CODIGO DE CUENTA
          CUEIX% = CUECOINT%(CODIGCUE$)
        IMPDEB# = 0: IMPHAB# = 0
        On Error Resume Next
        IMPDEB# = CDbl(.Recordset.Fields(5)) 'IMPORTE DEBE
        IMPHAB# = CDbl(.Recordset.Fields(6)) 'IMPORTE HABER
        On Error GoTo 0
        RFPASE$ = TRIM$(.Recordset.Fields(7)) 'REFERENCIA PASE
        '
        'GUARDO DETALLE DE ASIENTO EN DetaAsien.
        On Error GoTo 0
        With DETASIEN1
          .AddNew
             !NUIDASIEN = NUMASI$ 'FALTA DEFINIR CRITERIO
               QRENGLON% = QRENGLON% + 1
             !OrNuPase = QRENGLON 'ORDEN NUMERO PASE
             !RefePase = RFPASE$ 'REFERENCIA PASE
             !FECHASIEN = FEASIEN
             !OrFeAsien = Val(Mid$(NUMASI$, 7)) 'FALTA DEFINIR CRITERIO
             !SucuAsien = 0
             !NuDfAsien = 0
             !StatPase = 1
             !CueCoExt = CODIGCUE$
             !CUECOINT = CUEINT(CODIGCUE$)
             !IDebePase = IMPDEB#
             !IHabePase = IMPHAB#
          .Update
        End With
        '
        If J = 1 Then
          RFASIEN1$ = RFASIEN$
          FEASIEN1 = FEASIEN
        End If
        J = 0
        '
        STD# = IMPDEB# + STD#
        STH# = IMPHAB# + STH#
        '
    .Recordset.MoveNext
    '
    On Error Resume Next
    A12 = Val(.Recordset.Fields(1))
    If Err > 0 Then
       On Error GoTo 0
       GoTo 150
    End If
    On Error GoTo 0
    If Val(.Recordset.Fields(1)) <> NUMASIEN& Then
        'GUARDO ENCABEZADO SI CAMBIA EL NUMERO DE ASIENTO
150     If Abs(STD# - STH#) > 0.005 Then
          IMPDEB# = STH# - STD#: IMPHAB# = 0
          If IMPDEB# < 0 Then
            IMPDEB# = 0
            IMPHAB# = STD# - STH#
          End If
        With DETASIEN1
          .AddNew
             !NUIDASIEN = NUMASI$ 'FALTA DEFINIR CRITERIO
               QRENGLON% = QRENGLON% + 1
             !OrNuPase = QRENGLON 'ORDEN NUMERO PASE
             !RefePase = "Diferencia Asiento " + TRIM$(Str$(NUMASIEN)) 'REFERENCIA PASE
             !FECHASIEN = FEASIEN
             !OrFeAsien = Val(Mid$(NUMASI$, 7)) 'FALTA DEFINIR CRITERIO
             !SucuAsien = 0
             !NuDfAsien = 0
             !StatPase = 1
             !CueCoExt = CUEDIFE$
             !CUECOINT = CUEINT(CUEDIFE$)
             !IDebePase = IMPDEB#
             !IHabePase = IMPHAB#
          .Update
        End With
        End If
        With ENCASIEN1
          .AddNew
          !NUIDASIEN = NUMASI$ 'FALTA DEFINIR CRITERIO
          !RefeAsien = RFASIEN1$
          !FECHASIEN = FEASIEN1
          !OrFeAsien = Val(Mid$(NUMASI$, 7))
          !SucuAsien = 0
          !NuDfAsien = 0 'FALTA DEFINIR CRITERIO
          !TipoAsien = 0
          !StatAsien = 1
          !EditAsien = 1
          !QPasAsien = QRENGLON%  'CANTIDAD DE PASES DEL ASIENTO.
          !DebeAsien = STD# 'IMPORTE DEL ASIENTO
          !Observa = ""
          .Update
        End With
        QRENGLON% = 0
        J = 1
        NUMASI$ = ""
        STD# = 0
        STH# = 0
    End If
    '
    U& = U& + 1
    QREN& = QREN& + 1
    Label5 = U&
    Label5.Refresh
    Loop
   End With
99 End Sub

Function NUPROA$(FEXII%)
   '
   Static FLIMIX%
   If FLIMIX% = 0 Then FLIMIX% = FECHANUM("31/12/00")
   '
   NUPROA$ = ""
   '
   FEASX = FEXII%
   FECHAS$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = CVDAT(FECHAU%)
   If FECHAU% <= FLIMIX% Then
       NUIDAS$ = TRIM$(Str$(FECHAU%))
       While Len(NUIDAS$) < 5: NUIDAS$ = "0" + NUIDAS$: Wend
     Else
       NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
       If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   End If
   ORDAS% = 0
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset("SELECT * FROM EncaBEAsien WHERE (STR$(FECHASIEN) = '" & FEASYYYY & "' and SucuAsien = " & SUCUI% & ") ORDER BY  ORFEASIEN DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then GoTo 20
   ORDAS% = SELECASIEN!OrFeAsien
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function
'
Sub IMPORPREFAC()
   '
   Dim FECHAMV As Date
   Screen.MousePointer = 11
   '
   ULOI$ = Left$(LUDAT$, 2)
   '
   TEPRUE = ""
   Data1.DatabaseName = ULOI$ + Label9 + Label1
   Data1.RecordSource = "'0UOYKCUF'$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   With Data1
    .Recordset.MoveFirst
    U& = 0
   Call ABRECABAN
   Do While Not .Recordset.EOF
     '
     'Label7 = U&
     'Label7.Refresh
     '
     If IsNull(.Recordset.Fields(0)) = True And IsNull(.Recordset.Fields(1)) = True And IsNull(.Recordset.Fields(2)) = True And IsNull(.Recordset.Fields(3)) = True And IsNull(.Recordset.Fields(6)) = True Then
        GoTo 36
     End If
     '
     FECHAMV = TRIM$(.Recordset.Fields(0))  'NUME MOVI CTA. CTE.
     
     TICOMP$ = TRIM$(.Recordset.Fields(1))  'TIPO DE COMPROBANTE
     
     NUMCOMP$ = TRIM$(.Recordset.Fields(3))  'NUMERO DE COMPROBANTE
     '
     PREFIJ1$ = ""
     If IsNull(.Recordset.Fields(2)) = False Then
       PREFIJ1$ = CStr(TRIM$(.Recordset.Fields(2))) 'PREFIJO DE FACTURA
     End If
     provenom$ = TRIM$(.Recordset.Fields(6)) 'CODI PROV STRING
     '
     NPROV% = NUMEPRO(provenom$)
     If NPROV% < 1 Then
       MsgBox "PROVEEDOR NO VALIDO"
       GoTo 36
     End If
     '
        TIPOCOMPRO$ = ""
        Select Case TICOMP$
            
            Case "NODEB"
              TIPOCOMPRO$ = "ND"
            Case "ND"
              TIPOCOMPRO$ = "ND"
            Case "NDC"
              TIPOCOMPRO$ = "ND"
            Case "NDEBP"
              TIPOCOMPRO$ = "ND"
            Case "NDIP"
              TIPOCOMPRO$ = "ND"
            Case "NDP"
              TIPOCOMPRO$ = "ND"
            Case "O/P"
              TIPOCOMPRO$ = "OP"
            Case "FPR"
              TIPOCOMPRO$ = "FC"
            Case "FP"
              TIPOCOMPRO$ = "FC"
            Case "APINP"
              TIPOCOMPRO$ = "AS"
            Case "CREIN"
              TIPOCOMPRO$ = "AS"
            Case "DEBAN"
              TIPOCOMPRO$ = "AS"
            Case "DEINT"
              TIPOCOMPRO$ = "AS"
            Case "NC"
              TIPOCOMPRO$ = "NC"
            Case "NCP"
              TIPOCOMPRO$ = "NC"
            Case "NCREG"
              TIPOCOMPRO$ = "NC"
            Case "NOCRE"
              TIPOCOMPRO$ = "NC"
            Case "RANSE"
              TIPOCOMPRO$ = "AS"
        End Select
        'ACA ARMO
        'CODICOMCOR
       If TIPOCOMPRO$ = "FC" Or TIPOCOMPRO$ = "ND" Or TIPOCOMPRO$ = "NC" Then
        NCOMPCOR$ = TIPOCOMPRO$ & "-" & NUMCOMP$ & "-" & "A"
       Else
        NCOMPCOR$ = TIPOCOMPRO$ & "-" & NUMCOMP$
       End If
       'CODICOMLAR
       NPREFIJ$ = 0
       'COMPLETO PREFIJO
       If Len(NPREFIJ$) < 4 Then
          NJJ% = Len(NPREFIJ$)
          For I = NJJ% To 3
            NPREFIJ$ = "0" + NPREFIJ$
          Next I
       End If
       'COMPLETO NUMERO
       NUMECOMP1 = NUMCOMP$
       If Len(NUMCOMP$) < 8 Then
          NJJ1% = Len(NUMCOMP$)
          For I = NJJ1% To 7
             NUMCOMP$ = "0" + NUMCOMP$
          Next I
       End If
       '
       PREFIJNUE$ = TRIM$(PREFIJ1$)
       While Len(PREFIJNUE$) < 4: PREFIJNUE$ = "0" + PREFIJNUE$: Wend
       PREFIJNUE$ = Right$(PREFIJNUE$, 4)
       '
       Select Case TIPOCOMPRO$
        Case "FC"
            NCOMPLAR$ = TIPOCOMPRO$ & "-A-" & NPREFIJ$ & "-" & NUMCOMP$
        Case "NC"
            NCOMPLAR$ = TIPOCOMPRO$ & "-A-" & NPREFIJ$ & "-" & NUMCOMP$
        Case "ND"
            NCOMPLAR$ = TIPOCOMPRO$ & "-A-" & NPREFIJ$ & "-" & NUMCOMP$
        Case "OP"
            NCOMPLAR$ = TIPOCOMPRO$ & "-X-" & NPREFIJ$ & "-" & NUMCOMP$
        Case Else
            NCOMPLAR$ = TIPOCOMPRO$ & "-X-" & NPREFIJ$ & "-" & NUMCOMP$
       End Select
       '
       Colarvie$ = NCOMPLAR$
       COLARNUE$ = Colarvie$
       Call REPLA(COLARNUE$, PREFIJNUE$, 6, 4)
       '
       FechEmi = CVDAT(CVINT(FECHAMV))
       '
       On Error GoTo 0
       'BUSCO EN CAJABANC
       
          With CajaBanc
            YYY$ = "CodiCom_Lar = '" & Colarvie$ & "'   "
            YYY$ = YYY$ + " AND NuProve = " & NPROV% & " "
            YYY$ = YYY$ + "AND cdbl(FechEmisi) = " & CDbl(FECHAMV) & " "
            .FindFirst YYY$
            If .NoMatch = False Then
                .Edit
                !CODICOM_LAR = COLARNUE$
                !DescriMov = AJUSTI$(COLARNUE$ + " - " + RASOCLI$((-1) * NPROV%), 48)
                .Update
                GoTo 101
            End If
          End With
101        'BUSCO EN SACREPEN
            With SacrePen
              YYY$ = "CodiCom_Lar = '" & Colarvie$ & "'   "
              YYY$ = YYY$ + " AND NuProve = " & NPROV% & " "
              YYY$ = YYY$ + "AND cdbl(FechEmisi) = " & CDbl(FECHAMV) & " "
             .FindFirst YYY$
              If .NoMatch = False Then
                .Edit
                !CODICOM_LAR = COLARNUE$
                .Update
                GoTo 102
              End If
            End With
102        'BUSCO EN APLIPAGO
            With ApliPago
              YYY$ = "CoDocOr_Lar = '" & Colarvie$ & "'   "
              YYY$ = YYY$ + " AND NuProve = " & NPROV% & " "
              .FindFirst YYY$
              If .NoMatch = False Then
                .Edit
                !COdocor_lar = COLARNUE$
                .Update
              End If
              '
              YYY$ = "CoDocAp_Lar = '" & Colarvie$ & "'   "
              YYY$ = YYY$ + " AND NuProve = " & NPROV% & " "
              .FindFirst YYY$
              If .NoMatch = False Then
                .Edit
                !COdocAp_lar = COLARNUE$
                .Update
              End If
            End With
        '
36    .Recordset.MoveNext
     U& = U& + 1
     If U& = 10000 Then MsgBox "10000" 'para trazar
    DoEvents
    Loop
   End With
20 Screen.MousePointer = 1
   '
99 End Sub

Sub IMPORMAS2()
   '
5  Label1 = Label9 + "MAESTRO PRODUCTOS\STO.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "STO$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   '
   Call ABREMASTER
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
10   CDXI% = CODINT%(CODEX$)
     If CDXI% = 0 Then
       Call MENSERR(24, "Código / Campo Clave '" + CODEX$ + "' Inexistente en Base de Datos")
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     DBGrid1.Col = 3
     STOMI = Val(DBGrid1.TEXT)
     '
     DBGrid1.Col = 4
     PUNPE = Val(DBGrid1.TEXT)
     If STOMI < 1 Then STOMI = PUNPE
     '
     DBGrid1.Col = 5
     STOMA = Val(DBGrid1.TEXT)
     If STOMA < STOMI Then STOMA = STOMI
     '
     LOREP = STOMA - STOMI
     If LOREP > 10000 Then LOREP = 10000
     If LOREP < STOMI Then LOREP = STOMI
     '
     With Master
       .Edit
       !STOMIN = STOMI
       !STOMAX = STOMA
       !lorepos = LOREP
       .Update
     End With
19 Next U&
   '
99 End Sub
'
Sub AGRECODBAJ(CODEX$)
    If TRIM$(CODEX$) <> "" Then
      With Master
       .AddNew
       !CODIGO = CODEX$
       !DESCRIPCION = "Codigo " + CODEX$
       !STATUS = (-1)
         NMS% = NUMAS% + 1
       !CODINT = NMS%
       .Update
     End With
   End If
End Sub

Sub IMPORSTRUC()
   '
   Call ABRESTRUC
   On Error Resume Next
   Formulas.MoveLast
   While Formulas.RecordCount > 0
     Formulas.MoveLast
     Formulas.Delete
   Wend
   '
6  Label1 = Label9 + "ESTRUCTURAS\EST.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 20
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "EST$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 20
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   Call ABREMASTER
   Call ABRESTRUC
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
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then GoTo 19
     '
10   If CODINT%(CODEX$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     DBGrid1.Col = 1
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Then GoTo 19
     If COELE$ = CODEX$ Then GoTo 19
     '
     If Len(COELE$) = 3 Then
       If Left$(COELE$, 2) = "MO" Then
         GoTo 19   ' MANO DE OBRA
       End If
     End If
     '
12   If CODINT%(COELE$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento Inexistente o No Valido"
       Call AGRECODBAJ(COELE$)
       GoTo 12
     End If
     '
     CICON% = CODINT%(CODEX$)
     CIELE% = CODINT%(COELE$)
     '
     DBGrid1.Col = 2
     USONE = CDbl(DBGrid1.TEXT)
     '
     DBGrid1.Col = 3
     UMEDEL$ = TRIM$(DBGrid1.TEXT)
     '
     If UMEDEL$ <> UNIMED$(CIELE%) Then
       With Master
         .Edit
         !UMEDIDA = UMEDEL$
         .Update
       End With
     End If
     '
     With Formulas
       .AddNew
       !CODICONJ = Abs(CICON%)
       !CODXCONJ = CODEX$
       !CODIELEM = Abs(CIELE%)
       !CODXELEM = COELE$
       !USOBRUTO = USONE
       !PLADISPO = 0
       .Update
     End With
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub IMPORRUTAS()
   '
   Dim UOP%(32767)
   '
   Call ABRESECOP
   On Error Resume Next
   Secoper.MoveLast
   While Secoper.RecordCount > 0
     Secoper.MoveLast
     Secoper.Delete
   Wend
   '
6  Label1 = Label9 + "ESTRUCTURAS\EST.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 20
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "EST$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 20
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   Call ABREMASTER
   Call ABRESECOP
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
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 1
     COELE$ = UCase$(TRIM$(DBGrid1.TEXT))
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Then GoTo 19
     If COELE$ = CODEX$ Then GoTo 19
     '
     If Len(COELE$) <> 3 Then GoTo 19
     '
     If UCase$(Left$(COELE$, 2)) <> "MO" Then
       GoTo 19   ' NO ES MANO DE OBRA
     End If
     '
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then GoTo 19
     '
10   If CODINT%(CODEX$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     CICON% = CODINT%(CODEX$)
     '
     DBGrid1.Col = 2
     USONE = CDbl(DBGrid1.TEXT)
     '
     DBGrid1.Col = 3
     UMEDEL$ = UCase$(TRIM$(DBGrid1.TEXT))
     If Left$(UMEDEL$, 3) <> "MIN" Then
        AAA = BBB
        MsgBox "Unidad de M.O. no Reconocida"
     End If
     '
     If USONE * LOTESTAN(CICON%) / 60 > 999 Then
       ' limita tamaño del lote a menos de 999 horas
       Master.Edit
       Master!lorepos = 1
       Master.Update
     End If
     '
     With Secoper
       .AddNew
       !CODICONJ = Abs(CICON%)
       !CODXCONJ = CODEX$
         UOP%(CICON%) = UOP%(CICON%) + 10
       !NUMEOPER = UOP%(CICON%)
         If COELE$ = "MOA" Then
             DESOP$ = "ARMADO"
           ElseIf COELE$ = "MOE" Then
             DESOP$ = "ESTAMPADO"
           ElseIf COELE$ = "MOP" Then
             DESOP$ = "PINTURA"
           Else
             AAA = BBB
         End If
       !DescOper = DESOP
       !codmaq0 = COELE$
       !STATOPER = 1
       !MinuStan = USONE
       !LOTESTAN = LOTESTAN(CICON%)
       If !LOTESTAN > 10000 Then
         !LOTESTAN = 10000
       End If
       !CantiOps = 1
       If USONE > 0.05 Then
         !Cadencia = 60 / USONE
       End If
       '
       ' PARA INICIALIZAR VALORES
       !HorsPrep = 0
       !HorsFijo = 0
       !HorsLimp = 0
       !HorsMHer = 0
       !PieCiclo = 0
       !PorDemor = 0
       !CoefSolp = 0
       !CategOps = 0
       .Update
     End With
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
     If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
       TEPRUE = TRIM$(TTXX$) + "*"
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
     If VTX$ <> "" And Val(VTX$) < 1 Then
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
     DBGrid1.Col = 4
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
     DBGrid1.Col = 5                     ' Cuota Horaria
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
     DBGrid1.Col = 6
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
       TEPRUE = TRIM$(TTXX$) + "*"
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
       If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
         TEPRUE = TRIM$(TTXX$) + "*"
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&
   
20 End Sub


Sub IMPORSTRU()
   '
   Dim PUNSTRU$(32767)
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   ATENTIA% = ATENTI%: ATENTI% = 12345 ' PA QUE NO REORGANICE IMPLOSI
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       CAIT% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = Space$(64)
         Call REPLA(TTYY$, Mid$(TTXX$, 19, 16), 1, 16)
         Call REPLA(TTYY$, Mid$(TTXX$, 69, 12), 49, 12)
         Call REPLA(TTYY$, Mid$(TTXX$, 81, 4), 61, 4)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("ESTRULAR")
   ATENTI% = ATENTIA%
   '
End Sub
'
Sub IMPORUTS()
   '
   Dim PUNSTRU$(32767)
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   REBLA$ = REGBLAN$("SECOPER")
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       RFF$ = RECFILT$("SECOPER", 1, MKI$(CDXI%))
       '
       For UKU& = 1 To Len(RFF$) Step 4
         RESE& = CVS(Mid$(RFF$, UKU&, 4))
         Call GRAREG("SECOPER", String$(128, 0), RESE&)
       Next UKU&
       '
       NOPEI% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = REBLA$
         Call REPLA(TTYY$, MKI$(CODINT%(Left$(TTXX$, 16))), 1, 2)
           NOPEI% = NOPEI% + 10
         Call REPLA(TTYY$, MKI$(NOPEI%), 3, 2)
         Call REPLA(TTYY$, Mid$(TTXX$, 23, 48), 5, 48)
         Call REPLA(TTYY$, Mid$(TTXX$, 73, 6), 53, 6)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 79, 10))), 59, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 89, 10))), 63, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 99, 12))), 67, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 111, 10))), 91, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 121, 8))), 75, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 129, 8))), 83, 4)
         Call REPLA(TTYY$, Mid$(TTXX$, 139, 6), 103, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 147, 6), 109, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 155, 6), 115, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 163), 121, 6)
         Call REGAPP("SECOPER", TTYY$)
       Next UKU&
     End If
   Next CDXI%
   '
   Call CIERRARCH("SECOPER")
   Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAQ()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADMAQ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COMAQ$ = TRIM$(Left$(TTXX$, 6))
     DEMAQ$ = TRIM$(Mid$(TTXX$, 9, 30))
     CUHOR = Val(Mid$(TTXX$, 41, 12))
     SECCI = Val(Mid$(TTXX$, 53, 10))
     MAQEQ$ = TRIM$(Mid$(TTXX$, 66, 6))
     COEFI% = Val(Mid$(TTXX$, 72, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 87, 30))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, COMAQ$, 1, 6)
     Call REPLA(REMAQ$, DEMAQ$, 7, 30)
     Call REPLA(REMAQ$, MKS$(SECCI), 37, 4)
     Call REPLA(REMAQ$, MAQEQ$, 43, 6)
     Call REPLA(REMAQ$, MKI$(COEFI), 49, 2)
     Call REPLA(REMAQ$, OBSER$, 51, 30)
     Call REPLA(REMAQ$, MKS$(CUHOR), 93, 4)
     Call FILTERSETRECORD("PADMAQ", 1, COMAQ$, REMAQ$) ' sobreescribe SI IGUAL CODIGO
     '
   Next U&
   Call CIERRARCH("PADMAQ")
   '
End Sub

Sub IMPORCAT()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("CATOPER")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     ICAT% = Val(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = Val(Mid$(TTXX$, 37, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 51, 32))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, Chr$(ICAT%), 1, 1)
     Call REPLA(REMAQ$, DECAT$, 2, 27)
     Call REPLA(REMAQ$, MKS$(CUHOR), 29, 4)
     Call REPLA(REMAQ$, OBSER$, 33, 32)
     JJ& = JJ& + 1
     Call GRAREG("CATOPER", REMAQ$, JJ&)
   Next U&
   '
   Call SETULTREG("CATOPER", JJ&)
   Call CIERRARCH("CATOPER")
   '
End Sub



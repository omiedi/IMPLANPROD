VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FIKARDEX 
   BackColor       =   &H00FEFEE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Movimientos de Material (Ficha Kardex)"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FIKARDEX1.frx":0000
      Height          =   4110
      Left            =   105
      OleObjectBlob   =   "FIKARDEX1.frx":0010
      TabIndex        =   25
      Top             =   1785
      Width           =   11460
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2205
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5145
      Width           =   3795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3950
      Index           =   8
      Left            =   11360
      TabIndex        =   24
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEFEE0&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   19
      Top             =   630
      Width           =   5220
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   23
         Top             =   360
         Width           =   4845
      End
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3045
      Picture         =   "FIKARDEX1.frx":09C2
      TabIndex        =   17
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   6615
      Picture         =   "FIKARDEX1.frx":0CCC
      TabIndex        =   16
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      Left            =   6090
      TabIndex        =   15
      Text            =   " "
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FEFEE0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   5460
      TabIndex        =   8
      Top             =   630
      Width           =   4500
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FEFEE0&
         Caption         =   "Incluir Antiguos"
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
         Left            =   2625
         TabIndex        =   22
         Top             =   280
         Width           =   1695
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4095
         Picture         =   "FIKARDEX1.frx":0FD6
         TabIndex        =   12
         Top             =   600
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1995
         Picture         =   "FIKARDEX1.frx":12E0
         TabIndex        =   11
         Top             =   600
         Width           =   175
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   10
         Text            =   " "
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3045
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   0
         TabIndex        =   14
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   2205
         TabIndex        =   13
         Top             =   630
         Width           =   750
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1050
      TabIndex        =   3
      Top             =   210
      Width           =   2010
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   1670
      Left            =   10100
      ScaleHeight     =   1605
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   880
         Picture         =   "FIKARDEX1.frx":15EA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   880
         Picture         =   "FIKARDEX1.frx":1734
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FIKARDEX1.frx":1D9E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1065
         Width           =   1515
      End
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   11600
      Y1              =   5860
      Y2              =   5860
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4515
      TabIndex        =   18
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6930
      TabIndex        =   7
      Top             =   210
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep:"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   260
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3570
      TabIndex        =   5
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   0
      TabIndex        =   4
      Top             =   250
      Width           =   960
   End
End
Attribute VB_Name = "FIKARDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTSEL_Click(Index As Integer)
    If Index < 2 Then
         Call SELECHO("SELFECHA")
         If VALACT1$("SELFECHA") <> "" Then
           If Index = 0 Then
                Text11.TEXT = VALACT1$("SELFECHA")
              Else
                Text12.TEXT = VALACT1$("SELFECHA")
           End If
         End If
      ElseIf Index = 2 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If
      ElseIf Index = 3 Then
        Call SELECHO("MASTER")
        If VALACT1$("MASTER") <> "" Then
          Text1 = VALACT1$("MASTER")
        End If
    End If
End Sub

Private Sub Check1_Click()
   '
   If Check1.Value = 1 Then
      Call MENSERR(24, "No Disponible en la Presente Versión.")
      Check1.Value = 0
      Exit Sub
   End If
   '
End Sub

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command8_Click()
   '
   If Data1.Recordset.RecordCount < 1 Then
      Exit Sub
   End If
   '
   Command8.Enabled = False
   ZBLA$ = REGBLAN$("COKARDEX")
   JJ& = 0
   FIN& = Data1.Recordset.RecordCount
   '
20 Data1.Recordset.MoveFirst
     '
   With Data1.Recordset
     Do While Not .EOF
       J2& = J1% + 1
       Call TRACE(24, J2&, FIN&)
       '
       FEC1% = CVINT(.Fields(0).Value)
       CM1$ = .Fields(1).Value
       DOPRI1$ = .Fields(2).Value
       DOSEC1$ = .Fields(3).Value
       REFE1$ = .Fields(4).Value
       DEP1% = .Fields(5).Value
       ENTRA1# = .Fields(6).Value
       SALE1# = .Fields(7).Value
       SALD1# = .Fields(8).Value
       '
       YY$ = ZBLA$
       Call REPLA(YY$, MKI$(FEC1%), 1, 2)
       CI1% = CODINT%(Text1.TEXT)
       Call REPLA(YY$, MKI$(CI1%), 3, 2)
       Call REPLA(YY$, CM1$, 21, 2)
       Call REPLA(YY$, DOPRI1$, 23, 10)
       Call REPLA(YY$, DOSEC1$, 33, 10)
       Call REPLA(YY$, REFE1$, 43, 30)
       Call REPLA(YY$, Chr$(DEP1%), 83, 1)
       Call REPLA(YY$, MKD$(ENTRA1#), 88, 8)
       Call REPLA(YY$, MKD$(SALE1#), 96, 8)
       Call REPLA(YY$, MKD$(SALD1#), 104, 8)
       JJ& = JJ& + 1
       Call GRAREG("COKARDEX", YY$, JJ&)
       '
     .MoveNext
     Loop
     '
   End With
   Call SETULTREG("COKARDEX", JJ&)
  
   '\\\OT-06-0390-MH-FIN///
   '
   Call HEADERS("COKARDEX", "")
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Private Sub Form_Load()
   Text11 = "01/01/80"
   HOII% = HOY(HOS$)
   Text12 = HOS$
   Call Text4_Change
End Sub

Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 75, 10))
   CANTSAL = XVALO(Mid$(TTXX$, 85, 10))
   '
   If CANTING > 0 Then
       IDLOT$ = TRIM$(Mid$(TTXX$, 129, 16))
       If IDLOT$ <> "" Then
         On Error Resume Next
         CONSUMLOT.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT.Label1 = IDLOT$
         CONSUMLOT.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM04.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM04.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM04.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM04.Show 1
   End If
   '
End Sub

Private Sub Text1_Change()
   '
   CI1% = CODINT%(Text1.TEXT)
   Label6.Caption = DESCRIT0$(CI1%)
   Label5.Caption = UNIMED$(CI1%)
   Call FRESHALL
   '
   Call GENKARDEX
End Sub

Private Sub Text1_DblClick()
   Call BOTSEL_Click(3)
End Sub

Private Sub Text11_Change()
   'List1.Clear
   If FECHANUM(Text11) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text11_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text12_Change()
   'List1.Clear
   If FECHANUM(Text12) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text12_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text4_Change()
   'List1.Clear
   If Val(Text4) < 256 Then
     If Val(Text4) < 1 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) = "" Then
         Label4 = "Todos los Depósitos"
       Else
         Label4 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
     End If
     Call GENKARDEX
   End If
End Sub

Private Sub Text4_DblClick()
   Call BOTSEL_Click(2)
End Sub

Sub GENKARDEX()
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   LU$ = LUDAT$
'   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
'     Call COPYSTRU("MOVISTO", "COKARDEX")
'   End If
'   REBLA$ = REGBLAN$("COKARDEX")
   '
   'List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   Set Data1.Recordset = Nothing
   DBGrid1.ClearFields
   DBGrid1.Refresh
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11)
   DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12)
   HASTA = CDbl(CVDAT(HASFE%))
   If DESDE > HASTA Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0
   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   JJ& = 0
   SARRAS# = 0
   INCLUPOS% = 0
   If CIDD% > 0 Then
     Call ABRESTOCKS
     '
     ' OBTIENE EL ACUMULADO ANTERIOR A LA FECHA DESDE
     ' FALTA CONSIDERAR LOS MOVIMIENTOS ANTIGUOS
     SQLX$ = "select SUM(CantIngre) as TINGRE, sum(CantSalid) as TEGRE from movisto "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     SQLX$ = SQLX$ + "AND Cdbl(FechMov) < " & DESDE & " "
     Set AA12 = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     AA12.MoveFirst
     ACUMU = AA12!TINGRE - AA12!TEGRE
     If IsNull(ACUMU) Then ACUMU = 0
     '
     ' LLENA EL RECORDSET CON LOS MOVIMIENTOS ACTUALES
     ' FALTA CONSIDERAR LOS MOVIMIENTOS ANTIGUOS
     SQLX$ = " SELECT fechmov, CodiMovi , DoPriCor, DoSecCor , ReferCor, DepoMovi , CantIngre, CantSalid , CantSaldo FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
     SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set TRANSIT = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     With TRANSIT
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
              ACUMU = ACUMU + !cantIngre - !CantSalid
              .Edit
              !CantSaldo = ACUMU
              .Update
           .MoveNext
           Loop
        End If
        On Error GoTo 0
     End With
     '
     Set Data1.Recordset = TRANSIT
     DBGrid1.Refresh
     '
     On Error Resume Next
     Set colui = DBGrid1.Columns
     colui(0).Caption = "Fecha"
     colui(1).Caption = "CM"
     colui(2).Caption = "Doc-Pri"
     colui(3).Caption = "Doc-Sec"
     colui(4).Caption = "Referencia"
     colui(5).Caption = "D"
     colui(6).Caption = "Entrada"
     colui(7).Caption = "Salida"
     colui(8).Caption = "Saldo"
     colui(6).NumberFormat = "        0.0"
     colui(7).NumberFormat = "########0.0"
     colui(8).NumberFormat = "########0.0"
     '
     colui(0).Width = 1000
     colui(1).Width = 300
     colui(2).Width = 1320
     colui(3).Width = 1320
     colui(4).Width = 2750
     colui(5).Width = 300
     colui(6).Width = 1300
     colui(7).Width = 1300
     colui(8).Width = 1300
     '
Screen.MousePointer = 1
Exit Sub
     
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           If CVINT(!FechMov) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!FechMov) < DESFE% Then
               SARRAS# = SARRAS# + !cantIngre - !CantSalid
             ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
               Call REPLA(YY$, MKI$(!Cod_Inte), 3, 2)
               Call REPLA(YY$, !IdenLote, 5, 16)
               Call REPLA(YY$, !CodiMovi, 21, 2)
               Call REPLA(YY$, !DoPriCor, 23, 10)
               Call REPLA(YY$, !DoSecCor, 33, 10)
               Call REPLA(YY$, !ReferCor, 43, 30)
               Call REPLA(YY$, MKD$(!ValoUnit), 73, 8)
                 MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(!DepoMovi), 83, 1)
               Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
               Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
               Call REPLA(YY$, MKD$(!cantIngre), 88, 8)
               Call REPLA(YY$, MKD$(!CantSalid), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(!DoPriCor, 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(!DoSecCor, 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
               JJ& = JJ& + 1
               'Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
       End If
     End With
     Set MoviTemp = Nothing
     'Movisto.Close
     '
     On Error GoTo 0
     'Call SETULTREG("!COKARDEX", JJ&)
     Call SETULTREG("#COKARDEX", JJ&)
     Screen.MousePointer = 11
     '
     If INCLUVIE% = 1 Then
       Call BLOQARCH("COKARDEX")
       'Call TRACENTRAL("COKARDEX", "")
       Call FILEDUMP("COKARDEX")
       Call FILESORT("COKARDEX", "", 1, 1, "ASC")
       'Call TRALOCAL("COKARDEX", "")
       For KKUL& = 1 To JJ&
         CKX$ = REGLEIDO$("COKARDEX", KKUL&)
         Call GRAREG("#COKARDEX", CKX$, KKUL&)
       Next KKUL&
       Call CIERRARCH("COKARDEX")
     End If
     '
     'FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
       TTXX$ = Space$(256)
         FEX = DIANTE(DESFE%)
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
       Call REPLA(TTXX$, REFEM$, 43, 36)
         CASALDO# = SARRAS#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
       List1.AddItem TTXX$
     End If
     '
     For KU& = 1 To FIN&
       'YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       TTXX$ = Space$(256)
         FEX = CVI(Left$(YY$, 2))
       Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
         CMOVI$ = Mid$(YY$, 21, 2)
       Call REPLA(TTXX$, CMOVI$, 11, 2)
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       Call REPLA(TTXX$, DOPRI$, 15, 12)
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       Call REPLA(TTXX$, DOSEC$, 29, 12)
         REFEM$ = Mid$(YY$, 43, 30)
       Call REPLA(TTXX$, REFEM$, 43, 26)
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       Call REPLA(TTXX$, DEPOM$, 70, 3)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       Call REPLA(TTXX$, FORMATNUM$(CAINGRE#, "F10.1"), 75, 10)
         CASALI# = CVD(Mid$(YY$, 96, 8))
       Call REPLA(TTXX$, FORMATNUM$(CASALI#, "F10.1"), 85, 10)
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
         IDLOT$ = Mid$(YY$, 5, 16)
       Call REPLA(TTXX$, IDLOT$, 129, 16)
       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     Screen.MousePointer = 1
     '
 Exit Sub
     Call SETULTREG("COKARDEX", JJ&)
     Call CIERRARCH("COKARDEX")
     Screen.MousePointer = 1
     Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
     Screen.MousePointer = 1
    Else
     Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   'If TRIM$(CODL$) = "" Then GoTo 5
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub


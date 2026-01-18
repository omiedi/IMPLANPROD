VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.ocx"
Begin VB.Form FIKARDEX09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Movimientos de Material (Ficha Kardex)"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID 
      Height          =   5655
      Left            =   120
      TabIndex        =   34
      Top             =   1800
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9975
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   8433840
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      GridColor       =   12632256
      GridColorFixed  =   0
      FocusRect       =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "FIKARDEX09.frx":0000
      Left            =   8160
      List            =   "FIKARDEX09.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "DESCRIPCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1180
      Left            =   105
      TabIndex        =   19
      Top             =   630
      Width           =   4755
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   700
         TabIndex        =   31
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   75
         TabIndex        =   21
         Top             =   240
         Width           =   4605
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
      Height          =   360
      Index           =   3
      Left            =   3630
      Picture         =   "FIKARDEX09.frx":0004
      TabIndex        =   17
      Top             =   215
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
      Height          =   300
      Index           =   2
      Left            =   6060
      Picture         =   "FIKARDEX09.frx":030E
      TabIndex        =   16
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5535
      TabIndex        =   15
      Text            =   " "
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   4980
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Incluir Antiguos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2880
         TabIndex        =   20
         Top             =   680
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
         Height          =   300
         Index           =   1
         Left            =   2385
         Picture         =   "FIKARDEX09.frx":0618
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
         Height          =   300
         Index           =   0
         Left            =   2385
         Picture         =   "FIKARDEX09.frx":0922
         TabIndex        =   11
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Text            =   " "
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1305
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   360
         TabIndex        =   14
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   13
         Top             =   675
         Width           =   750
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   210
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   10620
      ScaleHeight     =   1725
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   1575
         TabIndex        =   32
         Top             =   1080
         Width           =   1605
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   50
         Picture         =   "FIKARDEX09.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   120
         Width           =   600
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
         Left            =   1500
         Picture         =   "FIKARDEX09.frx":1F92
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   120
         Width           =   600
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
         Left            =   780
         Picture         =   "FIKARDEX09.frx":20DC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   120
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID1 
      Height          =   5625
      Left            =   105
      TabIndex        =   24
      Top             =   1800
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   9922
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      ForeColorFixed  =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "FIKARDEX09.frx":2746
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   -240
      OleObjectBlob   =   "FIKARDEX09.frx":297A
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   720
      TabIndex        =   25
      Top             =   16120
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   10155
      Picture         =   "FIKARDEX09.frx":29F4
      Stretch         =   -1  'True
      ToolTipText     =   "Detalle Movimientos por Lote Activo"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Partidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7020
      TabIndex        =   23
      Top             =   285
      Width           =   1065
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4395
      TabIndex        =   18
      Top             =   210
      Width           =   405
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5535
      TabIndex        =   7
      Top             =   555
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5040
      TabIndex        =   6
      Top             =   285
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3830
      TabIndex        =   5
      Top             =   290
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -30
      TabIndex        =   4
      Top             =   290
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
   End
   Begin VB.Menu filtros 
      Caption         =   "Filtros"
      Begin VB.Menu Filtroporproveedor 
         Caption         =   "Filtrar Código Activo por Proveedor"
      End
   End
End
Attribute VB_Name = "FIKARDEX09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$

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
   Call GENKARDEX
End Sub

Private Sub Combo1_Click()
  'SeleccionEnList
'  For I = 0 To List1.ListCount - 1
'     LOTMOV$ = Mid$(List1.List(I), 129, 16)
'     If LOTMOV$ = Combo1.TEXT Then
'          List1.Selected(I) = True
'        Else
'          List1.Selected(I) = False
'     End If
'   Next I
  For i = 1 To GRID.Rows - 1
     LOTMOV$ = TRIM$(GRID.TextMatrix(i, 4))
     GRID.Row = i
     For J& = 0 To GRID.Cols - 1
        GRID.COL = J
        If LOTMOV$ = TRIM$(Combo1.TEXT) And TRIM$(Combo1.TEXT) <> "" Then
             GRID.CellBackColor = RGB(0, 0, 255)
             GRID.CellForeColor = RGB(255, 255, 255)
           Else
             GRID.CellBackColor = RGB(255, 255, 255)
             GRID.CellForeColor = RGB(0, 0, 0)
        End If
     Next J&
   Next i
End Sub

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
   
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To GRID.Rows - 1
      For i& = 0 To GRID.Cols - 1
       If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        GRID.Row = U&: GRID.COL = i&
        If GRID.TopRow > GRID.Row Then GRID.TopRow = GRID.Row
        On Error Resume Next
        VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If GRID.TopRow < GRID.Row - VISIROWS Then
              GRID.TopRow = GRID.Row - VISIROWS
            End If
         End If
        GRID.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Private Sub Command3_Click()
'    '
'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        '
'        Call BORRAREGISTROS("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05", REGAF&)
'        '
'80  Next ii%

'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        sqlx$ = "select * from movisto where cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)"
'        '
'        Dim rst As ADODB.Recordset
'        Set rst = New ADODB.Recordset
'        rst.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        '
'        With rst
'            Do While Not .EOF
'                dopricor$ = TRIM$(SAFETEXT$(!dopricor))
'                DEPO% = XVALO(!depomovi)
'                CANTI# = XVALO(!cantsalid)
'                '
'                CONDI$ = "cod_inte=" & CIXI% & " and depomovi=" & DEPO% & " and dopricor='" & dopricor$ & "'"
'                careg& = CantRegistros("movisto", CONDI$)
'                '
'                If careg& > 1 Then
'                    !cantsalid = 0
'                    !cantingre = 0
'                    .Update
'                     MsgBox "entro " & careg& & vbCrLf & CONDI$
'                End If
'                .MoveNext
'            Loop
'        End With
'        On Error Resume Next
'        rst.Close
'        Set rst = Nothing
'        On Error GoTo 0
'80  Next ii%
    '
    Call COMUNI("proceso finalizado")
    '
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   ZBLA$ = REGBLAN$("COKARDEX")
   JJ& = 0
   FIN& = GRID.Rows - 1 ' List1.ListCount
   '\\\OT-06-0390-MH-13/09/06///
   TODO% = 0
'   For KU& = 1 To FIN&
'     If List1.Selected(KU& - 1) = True Then GoTo 10
'   Next KU&
   TODO% = 1
   GoTo 20
   '
10 OPXX% = COMALTER%("Seleccione Tipo de Listado\\Texto Seleccionado\Todo\Cancelar")
   TODO% = OPXX% - 1
   If TODO% > 1 Then GoTo 99
   '
20  OPXX% = COMALTER%("Seleccione Tipo de Listado\\Cancelar\Reporte Standard\Exportar a Excel")
    If OPXX% <= 1 Then GoTo 99
    If OPXX% = 3 Then
       Screen.MousePointer = 11
       Call ImprimirEnExcelH(Me.GRID)
       Screen.MousePointer = 1
       GoTo 99
    End If
    For J% = 1 To GRID.Rows - 1   'List1.ListCount
     '
     J2& = J% + 1
     Call TRACE(24, J2&, FIN&)
     '
'     If TODO% = 1 Or List1.Selected(J% - 1) Then
        '
        'X$ = List1.List(J% - 1)
        FEC1% = FECHANUM(GRID.TextMatrix(J%, 0))  ' Left$(X$, 8))
        CM1$ = GRID.TextMatrix(J%, 1)             ' Mid$(X$, 11, 2)
        DOPRI1$ = GRID.TextMatrix(J%, 2)          ' Mid$(X$, 15, 12)
        DOSEC1$ = GRID.TextMatrix(J%, 3)          ' Mid$(X$, 29, 12)
        REFE1$ = GRID.TextMatrix(J%, 5)           ' Mid$(X$, 43, 26)
        DEP1% = XVALO(GRID.TextMatrix(J%, 6))     ' Mid$(X$, 70, 3))
        ENTRA1# = XVALO(GRID.TextMatrix(J%, 7))   ' Mid$(X$, 75, 10))
        SALE1# = XVALO(GRID.TextMatrix(J%, 8))    ' Mid$(X$, 85, 10))
        SALD1# = XVALO(GRID.TextMatrix(J%, 9))    ' Mid$(X$, 95, 12))
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
'     End If
     '
   Next J%
   Call SETULTREG("COKARDEX", JJ&)
  
   '\\\OT-06-0390-MH-FIN///
   '
   Call HEADERS("COKARDEX", "")
   If Val(Text4) > 0 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) <> "" Then
      Call HEADERS("COKARDEX", "Deposito: " & TRIM$(Text4) & " - " & ECOARCH$("TADEPOSI", Chr$(Val(Text4))))
   End If
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Sub ImprimirEnExcelH(MSFH As MSHFlexGrid)
    'ESTE RECIBE MSHFlexGrid   HHHHH
    'SI SE IMPRIME Y NO SE CIERRA EL EXCEL DA ERROR
    
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim flexGrid As MSHFlexGrid
    Dim i As Integer
    Dim J As Integer
    
    ' Asignar el nombre de tu MSFlexGrid aquí
    Set flexGrid = MSFH
    FIN& = flexGrid.Rows
    ' Crear una instancia de Excel
    Set excelApp = CreateObject("Excel.Application")
    
    ' Agregar un nuevo libro de trabajo
    Set excelWorkbook = excelApp.Workbooks.Add
    Set excelWorksheet = excelWorkbook.Worksheets(1)
    
    ' Fijar la fila de los encabezados
    excelWorksheet.Range("A2").Select ' Selecciona la celda justo debajo de la primera fila
    excelApp.ActiveWindow.FreezePanes = True ' Congela la primera fila
'    excelWorksheet.Rows(1).Select
'    excelApp.ActiveWindow.FreezePanes = True
    
    ' Aplicar color de fondo a la fila de los encabezados
    excelWorksheet.Rows(1).Interior.COLOR = RGB(192, 192, 192) ' Puedes ajustar los valores RGB según el color que desees para el fondo de los encabezados
    
    ' Copiar los datos desde el MSFlexGrid a Excel
    For i = 0 To flexGrid.Rows - 1
        II& = i
        For J = 0 To flexGrid.Cols - 1
            Call TRACE(20, II&, FIN&)
            excelWorksheet.Cells(i + 1, J + 1).Value = flexGrid.TextMatrix(i, J)
            excelWorksheet.Cells(i + 1, J + 1).WrapText = True
            
            ' Agregar marco a la celda
            excelWorksheet.Cells(i + 1, J + 1).Borders.LineStyle = 1 ' 1 = xlContinuous, puedes ajustar el estilo de línea según tus necesidades
        Next J
    Next i
    
    ' Ajustar automáticamente la altura de las celdas
    excelWorksheet.Cells.EntireRow.AutoFit
    
    ' Aplicar formatos a las celdas de fecha y números
    For i = 0 To flexGrid.Rows - 1
        For J = 0 To flexGrid.Cols - 1
            If IsDate(flexGrid.TextMatrix(i, J)) And InStr(flexGrid.TextMatrix(i, J), "/") > 0 Then
                excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "dd/mm/yyyy"
                excelWorksheet.Cells(i + 1, J + 1).Value = CDate(flexGrid.TextMatrix(i, J))
            ElseIf IsNumeric(flexGrid.TextMatrix(i, J)) Then
                If InStr(flexGrid.TextMatrix(i, J), ".") > 0 Then
                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0.00"
                Else
                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0"
                End If
            End If
        Next J
    Next i
    
    ' Ajustar el ancho de las columnas
    For J = 0 To flexGrid.Cols - 1
        Dim headerText As String
        headerText = flexGrid.TextMatrix(0, J)
        excelWorksheet.Columns(J + 1).AutoFit
        If excelWorksheet.Columns(J + 1).ColumnWidth < Len(headerText) Then
            excelWorksheet.Columns(J + 1).ColumnWidth = Len(headerText)
        End If
    Next J
    
    ' Mostrar Excel después de que se haya cargado completamente
    excelApp.Visible = True
    
    ' Liberar recursos de Excel
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
End Sub

Private Sub Filtroporproveedor_Click()
    CI1% = CODINT%(Text1)
    If CI1% < 1 Then Exit Sub
    
    Campos$ = "Cta.Prov/F9;R.Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
'    If ANTOT < 5500 Then ANTOT = 5500
'    FRMZELECHO.Width = ANTOT + 1500

    SQLX$ = " EXECUTE ProveedorSegunCodigo  " & CI1%
    Dim rs As New ADODB.Recordset
    Set rs = FLEXCONN.Execute(SQLX$)
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    '
    With rs
      Do While Not .EOF
        NPROV = XVALO(!NumProv)
        RAZONPROVEEDOR$ = rasocli$(-1 * NPROV)
        '
        J& = J& + 1
        FRMZELECHO.MSF2.Rows = J& + 1
        FRMZELECHO.MSF2.TextMatrix(J&, 1) = NPROV
        FRMZELECHO.MSF2.TextMatrix(J&, 2) = RAZONPROVEEDOR$
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    If J& = 0 Then
       Call COMUNI("No se Detectaron Entregas Ingresos Con Proveedor Para el Código Activo")
       Exit Sub
    End If
    
    FRMZELECHO.Caption = "Proveedores de Código: " & CODEXT$(CI1%)
    FRMZELECHO.Show 1
    
    PROVED = XVALO(RESP_ZELE_VECT(1))
    If PROVED > o Then
       Label8 = PROVED
       Label9 = rasocli$(-1 * PROVED)
    Else
       Label8 = ""
    End If

End Sub

Private Sub Form_Load()
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture3)
   '
   Text11 = "01/01/80"
   FEINIKAR$ = TRIM$(Control("", "FEINIKAR"))
   FF% = FECHANUM(FEINIKAR$)
   If FF% > 0 Then Text11 = FECHATEX$(FF%)
   HOII% = HOY(HOS$)
   Text12 = HOS$
   Call Text4_Change
   '\\\OT-08-0990-OD-30/10/08///
   LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LOMA% > 15 Then Text1.MaxLength = LOMA%
   '\\\OT-08-0990-OD-FIN///
   '
   
   TEPRUEBA.Font = GRID.Font
   TEPRUEBA.FontSize = GRID.FontSize
   GRID.Cols = 11
   GRID.Row = 1
   GRID.FixedRows = 1
'   Campos$ = "Fecha/D8" '1
'   Campos$ = Campos$ + ";CM/A2" '2
'   Campos$ = Campos$ + ";Transaccion/A16"   '3
'   Campos$ = Campos$ + ";Doc.Numero/A16"       '4
'   Campos$ = Campos$ + ";Id-Bulto/A12" '5
'   Campos$ = Campos$ + ";Referencia/A32"        '6
'   Campos$ = Campos$ + ";DP/F3"     '7
'   Campos$ = Campos$ + ";Entrada/F9.1" '8
'   Campos$ = Campos$ + ";Salida/F9.1" '9
'   Campos$ = Campos$ + ";Saldo/F10.1" '10
'   Campos$ = Campos$ + ";ROWGUID/A0;" '11
'   Call CARGA_ENCABEZADOMSFH(GRID, Campos$, Me)
   NUCOLU% = 1: TEPRUEBA = "99/99/99M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Fecha"
   NUCOLU% = 2: TEPRUEBA = "OOM": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "CM"
   NUCOLU% = 3: TEPRUEBA = "MM-0000-00000000M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Transaccion"
   NUCOLU% = 4: TEPRUEBA = "MM-0000-00000000M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Doc.Numero"
   NUCOLU% = 6: TEPRUEBA = "ABCDEFGH1JKLMNOPQR-MMMMMMMM": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Referencia"
   NUCOLU% = 7: TEPRUEBA = "99M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Dp"
   NUCOLU% = 8: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Entrada"
   NUCOLU% = 9: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Salida"
   NUCOLU% = 10: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Saldo"
   NUCOLU% = 5: TEPRUEBA = "MM-999999-99": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "ID-Bulto"
   NUCOLU% = 11: TEPRUEBA = "": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "USUARIO"

   '
End Sub

Private Sub GRID_DblClick()
    '
    iRow& = GRID.Row
    CANTING = XVALO(GRID.TextMatrix(iRow&, 7))
    CANTSAL = XVALO(GRID.TextMatrix(iRow&, 8))
    COMO$ = TRIM$(GRID.TextMatrix(iRow&, 1))
    BRE$ = TRIM$(GRID.TextMatrix(iRow&, 2))
    CI1% = CODINT%(TRIM$(Text1))
    
    '
    If CANTING > 0 Then
       IDLOT$ = TRIM$(GRID.TextMatrix(iRow&, 4))
       If IDLOT$ <> "" Then
         OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
         If OPX% < 1 Then Exit Sub
         If OPX% = 1 Then
             On Error Resume Next
             CONSUMLOT07.Label4 = Text1
             If Err Then Exit Sub
             On Error GoTo 0
             CONSUMLOT07.Label1 = IDLOT$
             CONSUMLOT07.Show 1
             Exit Sub
         End If
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(GRID.TextMatrix(iRow&, 2))
       If Err Then Exit Sub
       On Error GoTo 0
       OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
       If OPX% < 1 Then Exit Sub
       If OPX% = 1 Then
        LOCONSUM09.Label5 = TRIM$(GRID.TextMatrix(iRow&, 0))
        LOCONSUM09.Label6 = TRIM$(GRID.TextMatrix(iRow&, 4))
        LOCONSUM09.Show 1
        Exit Sub
       End If
    End If
    '
    Select Case TRIM$(UCase(COMO$))
      Case "BR"
        Campos$ = "O.Compra/F9;Cta./F8;Proveedor/A32;Id.Lote/A16"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Caption = "Orden de Compra Según Recepcion: " & BRE$
        FRMZELECHO.Width = 8600
        '
        NBR& = XVALO(Mid$(BRE$, 4))
        SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK) WHERE NUMEBORE = " & NBR& & " AND COD_INTE = " & CI1%
        Set rs = READSET(SQLX$): k& = 0
        Do While Not rs.EOF
          k& = k& + 1
          FRMZELECHO.MSF2.Rows = k& + 1
          FRMZELECHO.MSF2.TextMatrix(k&, 1) = XVALO(rs!NUMOC)
          FRMZELECHO.MSF2.TextMatrix(k&, 2) = XVALO(rs!NumProv)
          FRMZELECHO.MSF2.TextMatrix(k&, 3) = SAFETEXT$(rs!Raso_Prov)
          FRMZELECHO.MSF2.TextMatrix(k&, 4) = SAFETEXT$(rs!idenlote)
          rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        FRMZELECHO.Show 1
        NUOC& = XVALO(RESP_ZELE_VECT(1))
        DOCUNU& = NUOC&
        TIDOCU$ = "NOTA DE PEDIDO"
        NUDOCU$ = TRIM$(Str$(DOCUNU&))
        Call MUESTRADOC(0, "", TIDOCU$, NUDOCU$, NUDOCU$)

    End Select
    
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   i& = GRID.MouseRow
   J& = GRID.MouseCol
   If i& = 0 Or i& > GRID.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       GRID.Row = i&
       GRID.COL = J&
       GRID.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "ASEMBLI") > 0 Then
            FEMOX1$ = GRID.TextMatrix(i&, 0)
            TIMOI1$ = GRID.TextMatrix(i&, 1)
            DOPRIX1$ = GRID.TextMatrix(i&, 2)
            LOTEA1$ = GRID.TextMatrix(i&, 4)
            CANTING1 = GRID.TextMatrix(i&, 7)
            CANTSAL1 = GRID.TextMatrix(i&, 8)
            Call AJUASEMBLI(CODINT%(Text1), CANTING1, CANTSAL1, TIMOI1$, DOPRIX1$, FEMOX1$, LOTEA1$)
            GRID.CellBackColor = RGB(255, 255, 255)
            Exit Sub
        End If
       '
       fech% = FECHANUM(GRID.TextMatrix(i&, 0))
       If fech% < 1 Then Exit Sub
       FECHAX = FETEXCOR1$(fech%)
       COMOVI$ = TRIM$(GRID.TextMatrix(i&, 1))
       DOPRI$ = TRIM$(GRID.TextMatrix(i&, 2))
       DOSEC$ = TRIM$(GRID.TextMatrix(i&, 3))
       IDLOT$ = TRIM$(GRID.TextMatrix(i&, 4))
       NUMEROUSUARIO% = XVALO(GRID.TextMatrix(i&, 10))
       CONDI$ = "FECHMOV = '" & FECHAX & "' AND CODIMOVI = '" & COMOVI$ & "' AND DOPRILAR = '" & DOPRI$ & "' AND DOSECLAR = '" & DOSEC$ & "' And IDENLOTE = '" & IDLOT$ & "'"
    
       FECHAYHORA = VALOBADA("MOVISTO", "FEREMOVI", CONDI$, "NOMESS")
       MsgBox "Fecha y Hora Real del Movimiento  " & vbCrLf & vbCrLf & FECHAYHORA & vbCrLf & "Usuario: " & USERNAM$(NUMEROUSUARIO%), , "FLEXOFT"
       GRID.CellBackColor = RGB(255, 255, 255)
   End If
   
End Sub

Sub AJUASEMBLI(CIXI%, CANTING, CANTSAL, TIMOI$, DOPRIX$, FEMOX$, LOTEA$)
'         TTXX$ = List1.TEXT
         CANTING = XVALO(CANTING)
         CANTSAL = XVALO(CANTSAL)
         '
         If CIXI% < 1 Then Exit Sub
         '
         If CANTSAL > 0 Or TIMOI$ = "AJ" Or TIMOI$ = "TR" Then
            Call ARMALILOT(CIXI%)
            Call FRESHECHO("!LILOTSE")
            Screen.MousePointer = 1
            If ULTREG("!LILOTSE") < 1 Then
              Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
              Exit Sub
            End If
            '
            Call SELECHO("!LILOTSE")
            VDEVU$ = VALACT1$("!LILOTSE")
            VDEV2$ = VALACT2$("!LILOTSE")
            If TRIM$(VDEVU$) = "" Then Exit Sub
            '
            CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
            If CANTSAL > CAMAXI Then
               Call MENSERR(24, "Cantidad del Movimiento Supera\Saldo Fisico del Lote")
               Exit Sub
            End If
            '
            FEMOI% = FECHANUM(FEMOX$)
            TTZZ$ = "Aprueba Aplicar el Lote '" + VDEVU$ + "'\al Comprobante '" + TIMOI$ + "' - '" + DOPRIX$ + "'"
            TTZZ$ = TTZZ$ + "\de Fecha '" + FEMOX$ + "' ?\\No, Cancelar\Si, Aplicar"
            If COMALTER%(TTZZ$) = 2 Then
               SQLX$ = "Cod_Inte = " & CIXI%
               SQLX$ = SQLX$ + " AND CodiMovi = '" & TIMOI$ & "'"
               SQLX$ = SQLX$ + " AND DoPrilar = '" & DOPRIX$ & "'"
               SQLX$ = SQLX$ + " AND FechMov= '" & FETEXCOR1$(FEMOI%) & "'"
               If CANTSAL >= 0.05 Then SQLX$ = SQLX$ + " AND CantSalid = " & TRIM$(FORMATNUM(CANTSAL, "F12.4"))
               SQLX$ = SQLX$ + " AND IdenLote = '" & LOTEA$ & "' "
               '
               Screen.MousePointer = 11
               Call ACTUREGISTRO("MOVISTO", "IdenLote", SQLX$, VDEVU$, REGAF&, "NOMESS")
               If REGAF& > 0 Then
                    Call ACSALOTE(CIXI%, VDEVU$, "NOMESS")
                    Call GENKARDEX
               Else
                    Call MENSERR(24, "Actualización Imposible")
               End If
               Screen.MousePointer = 1
            End If
         End If
End Sub
Private Sub Image1_Click()
   IDLOT$ = TRIM$(Combo1.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = Text1
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
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
         CONSUMLOT07.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT07.Label1 = IDLOT$
         CONSUMLOT07.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM09.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM09.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM09.Show 1
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      '
      Linea$ = List1.List(List1.ListIndex)
      FECHA1$ = Mid$(Linea$, 1, 8)
      FF% = FECHANUM(FECHA1$)
      FECHA2$ = FETEXCOR1$(FF%)
      LOTE$ = Mid(Linea$, 15, 12)
      '
      CONDI$ = "DOPRICOR ='" & TRIM$(LOTE$) & "' AND FECHMOV='" & FECHA2$ & "'"
      '
      CODUSU% = XVALO(VALOBADA("MOVISTO", "NUMUSUAR", CONDI$))
      USU$ = USERNAM$(CODUSU%)
      FechaReal$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 1, 10)
      HORA$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 12, 8)
      '
      Call REPLA(VDEF$, USU$, 1, 20)
      Call REPLA(VDEF$, FechaReal$, 21, 10)
      Call REPLA(VDEF$, HORA$, 31, 8)
      Datos$ = OBJPLA("DATOSMOV", VDEF$)
   End If
   '
End Sub


Private Sub Label8_Change()
   Call GENKARDEX
End Sub

Private Sub Label8_DblClick()
   Label8 = ""
   Label9 = ""
End Sub

Private Sub Text1_Change()
   '
   CI1% = CODINT%(Text1.TEXT)
   Label6.Caption = DESCRIT0$(CI1%)
   Label5.Caption = Unimed$(CI1%)
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
Call GENKARDEX2
Exit Sub
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   'List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11)
   Desde = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12)
   Hasta = CDbl(CVDAT(HASFE%))
   If Desde > Hasta Then
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
   GRID.Redraw = False
   If CIDD% > 0 Then
     Call ABRESTOCKS
'     Call ABRESTOCKVIE
     iRow& = 0: GRID.Rows = 23
     For i& = 1 To GRID.Rows - 1
        For J& = 0 To GRID.Cols - 1
           GRID.TextMatrix(i&, J&) = ""
        Next J&
     Next i&
     If INCLUVIE% = 1 Then
     '
        sqlx1$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
        sqlx1$ = sqlx1$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "

        sqlx1$ = sqlx1$ + " FROM MOVIVIE WITH(NOLOCK) "
        sqlx1$ = sqlx1$ + "WHERE Cod_Inte = " & CIDD% & " "
        If DEPXX% > 0 Then
          If DEPXX% = PRIDEP% Then
              sqlx1$ = sqlx1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
            Else
              sqlx1$ = sqlx1$ + "AND DepoMovi = " & DEPXX% & " "
          End If
        End If
        sqlx1$ = sqlx1$ + "AND CodiMovi <> 'KU' "
        sqlx1$ = sqlx1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
'        Dim MOVIVIETMP As ADODB.Recordset
'        Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
        Set MOVIVIETMP = READSET(sqlx1$)
        With MOVIVIETMP
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
                  SARRAS# = SARRAS# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
                ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
                  YY$ = REBLA$
                  Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
                  Call REPLA(YY$, MKI$(XVALO(!cod_inte)), 3, 2)
                  Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
                  Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
                  Call REPLA(YY$, SAFETEXT$(!dopricor), 23, 10)
                  Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
                  Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
                  Call REPLA(YY$, MKD$(XVALO(!VALOUNIT)), 73, 8)
                    MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
                  Call REPLA(YY$, MONEX$, 81, 2)
                  Call REPLA(YY$, Chr$(XVALO(!depomovi)), 83, 1)
                  Call REPLA(YY$, MKI$(XVALO(!NUME_CLIE)), 84, 2)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
                  Call REPLA(YY$, MKD$(XVALO(!CANTINGRE)), 88, 8)
                  Call REPLA(YY$, MKD$(XVALO(!CANTSALID)), 96, 8)
                  Call REPLA(YY$, MKD$(0), 104, 8)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!dopricor), 11, 4), 112, 4)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
                  Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
                  Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
                  JJ& = JJ& + 1
                  Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
        End With
        MOVIVIETMP.Close
        Set MOVIVIETMP = Nothing
     End If
     
     SQLX$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
     SQLX$ = SQLX$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " AND (VERIFICADO <> -1 OR VERIFICADO IS NULL) "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov , FeReMovi ASC "
     
     

     '
'     Dim MoviTemp As ADODB.Recordset
'     Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     Set MoviTemp = READSET(SQLX$)
     With MoviTemp
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
               SARRAS# = SARRAS# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
             ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
               Call REPLA(YY$, MKI$(XVALO(!cod_inte)), 3, 2)
               Call REPLA(YY$, SAFETEXT$(!idenlote), 5, 16)
               Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
               Call REPLA(YY$, SAFETEXT$(!dopricor), 23, 10)
               Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
               Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
               Call REPLA(YY$, MKD$(XVALO(!VALOUNIT)), 73, 8)
                 MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(XVALO(!depomovi)), 83, 1)
               Call REPLA(YY$, MKI$(XVALO(!NUME_CLIE)), 84, 2)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
               Call REPLA(YY$, MKD$(XVALO(!CANTINGRE)), 88, 8)
               Call REPLA(YY$, MKD$(XVALO(!CANTSALID)), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(SAFETEXT$(!dopricor), 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
               Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
               JJ& = JJ& + 1
               ' Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
     End With
     MoviTemp.Close
     Set MoviTemp = Nothing
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
     iRow& = 0: GRID.Rows = 23
     For i& = 1 To GRID.Rows - 1
        For J& = 0 To GRID.Cols - 1
           GRID.TextMatrix(i&, J&) = ""
        Next J&
     Next i&
     Combo1.TEXT = "": Call Combo1_Click
     ' FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
        ' TTXX$ = Space$(256)
        FEX = DIANTE(DESFE%)
        ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
        ' IROW& = IROW& + 1
        ' GRID.Row = IROW&
        ' GRID.Col = 0
        GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
        REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
        ' Call REPLA(TTXX$, REFEM$, 43, 36)
        ' GRID.Col = 5: GRID.TEXT = REFEM$
        GRID.TextMatrix(iRow, 5) = REFEM$
        '
        CASALDO# = SARRAS#
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""), 95, 12)
        ' GRID.Col = 9: GRID.TEXT = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        GRID.TextMatrix(iRow, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        ' List1.AddItem TTXX$
     End If
     '
     GRID.Rows = FIN& + 1
     For KU& = 1 To FIN&
       Call TRACE(20, KU&, FIN&)
       ' YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       ' TTXX$ = Space$(256)
       FEX = CVI(Left$(YY$, 2))
       ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
       iRow& = iRow& + 1
       If iRow > GRID.Rows - 1 Then GRID.Rows = iRow& + 1
       ' GRID.Row = IROW&
       ' GRID.Col = 0:
       GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
         CMOVI$ = Mid$(YY$, 21, 2)
       GRID.TextMatrix(iRow&, 1) = CMOVI$
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       GRID.TextMatrix(iRow&, 2) = DOPRI$
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       GRID.TextMatrix(iRow&, 3) = DOSEC$
         REFEM$ = Mid$(YY$, 43, 30)
       GRID.TextMatrix(iRow&, 5) = REFEM$
       ' Call REPLA(TTXX$, REFEM$, 43, 26)
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       GRID.TextMatrix(iRow&, 6) = TRIM$(DEPOM$)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       GRID.TextMatrix(iRow&, 7) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         CASALI# = CVD(Mid$(YY$, 96, 8))
       GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
         IDLOT$ = Mid$(YY$, 5, 16)
       GRID.TextMatrix(iRow&, 4) = TRIM$(IDLOT$)
       '
       If TRIM$(IDLOT$) <> "" Then
          If InStr(VECOMBO$, IDLOT$) < 1 Then
             If PINI& > Len(VECOMBO$) Then VECOMBO$ = VECOMBO$ + Space$(1024)
             Mid$(VECOMBO$, PINI&, 16) = IDLOT$
             PINI& = PINI& + 16
          End If
       End If
       '
       Call CompletaCombo(IDLOT$)
'       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     '
     GRID.Redraw = True
     
     For KU& = 1 To PINI& - 1 Step 16
        Combo1.AddItem Mid$(VECOMBO$, KU&, 16)
     Next KU&
     '
     ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
     'TPXI& = List1.ListCount - 19
     'If TPXI& < 0 Then TPXI& = 0
     'List1.TopIndex = TPXI&
     '
     TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     GRID.TopRow = TPXI&
     On Error GoTo 0
     '
     Screen.MousePointer = 1
     '
Exit Sub
'       Call SETULTREG("COKARDEX", JJ&)
'       Call CIERRARCH("COKARDEX")
'       Screen.MousePointer = 1
'       Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
'       Screen.MousePointer = 1
'    Else
'       Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub CompletaCombo(IDLOT$)
    '
    If TRIM$(IDLOT$) <> "" Then
       IDLOT1$ = TRIM$(IDLOT$)
       Dim EXISTE As Boolean
       EXISTE = False
       '
'       For i = 0 To Combo1.ListCount - 1
'          If Combo1.List(i) = IDLOT1$ Then
'             EXISTE = True
'             Exit Sub
'          End If
'       Next i
       '
       If EXISTE = False Then
          Combo1.AddItem IDLOT1$
       End If
    End If
    '
End Sub

Sub GENKARDEX2()
   '
   Call ABRECONEXION
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
      Call CARDEPOS
      PRIDEP% = COSU%(1)
   End If
   '
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
''   LU$ = LUDAT$
'''   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
'''     Call COPYSTRU("MOVISTO", "COKARDEX")
'''   End If
'''   REBLA$ = REGBLAN$("COKARDEX")

'''   List1.Clear
   'PARA LIMPIAR LA GRILLA POR MAS QUE EL CODIGO NO SEA VALIDO
   
   iRow& = 0: GRID.Rows = 23
   For i& = 1 To GRID.Rows - 1
     GRID.Row = i
      For J& = 0 To GRID.Cols - 1
          GRID.COL = J
         GRID.TextMatrix(i&, J&) = ""
         GRID.CellBackColor = RGB(255, 255, 255)
         GRID.CellForeColor = RGB(0, 0, 0)
      Next J&
   Next i&
   '
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = XVALO(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11): Desde = FETEXCOR1$(DESFE%)
   ' DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12): Hasta = FETEXCOR1$(HASFE%)
   ' HASTA = CDbl(CVDAT(HASFE%))
   If DESFE% > HASFE% Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0:   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   condi0$ = "Cod_Inte = " & CIDD% & " "
   If DEPXX% > 0 Then
      If DEPXX% = PRIDEP% Then
          condi0$ = condi0$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
        Else
          condi0$ = condi0$ + "AND DepoMovi = " & DEPXX% & " "
      End If
   End If
   '
   Screen.MousePointer = 11

   Call ABRESTOCKS
   iRow& = 0: GRID.Rows = 23
   For i& = 1 To GRID.Rows - 1
      For J& = 0 To GRID.Cols - 1
         GRID.TextMatrix(i&, J&) = ""
      Next J&
   Next i&
   ACUIN = 0: ACUSA = 0
   If INCLUVIE% = 1 Then
      condi0$ = condi0$ + "AND CodiMovi <> 'KU' "
      ACUIN = XVALO(VALOBADA("MOVIVIE", "SUM(CANTINGRE)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
      ACUSA = XVALO(VALOBADA("MOVIVIE", "SUM(CANTSALID)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   End If
   ACUIN = ACUIN + XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   ACUSA = ACUSA + XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   '
   SARRAS# = ACUIN - ACUSA
   If Abs(SARRAS#) >= 0.05 Then
       iRow& = iRow& + 1
       FEX = DIANTE(DESFE%)
       GRID.TextMatrix(iRow&, 0) = FECHATEX$(FEX)
       REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
       GRID.TextMatrix(iRow, 5) = REFEM$
       CASALDO# = SARRAS#
       GRID.TextMatrix(iRow, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
       iRow& = iRow& + 1  ' PARA QUE DEJE UNO LIBRE
   End If
   JJ& = 0
   ' SARRAS# = 0
   INCLUPOS% = 0
'   If CIDD% > 0 Then
'     Call ABRESTOCKS
'     Call ABRESTOCKVIE
'     If INCLUVIE% = 1 Then
'     '
'        SQLX1$ = " SELECT * FROM MOVIVIE WITH(NOLOCK) "
'        SQLX1$ = SQLX1$ + "WHERE Cod_Inte = " & CIDD% & " "
'        If DEPXX% > 0 Then
'          If DEPXX% = PRIDEP% Then
'              SQLX1$ = SQLX1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
'            Else
'              SQLX1$ = SQLX1$ + "AND DepoMovi = " & DEPXX% & " "
'          End If
'        End If
'        SQLX1$ = SQLX1$ + "AND CodiMovi <> 'KU' "
        sqlx1$ = sqlx1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
   BASEAC$ = "MOVISTO": If INCLUVIE% = 1 Then BASEAC$ = "MOVIVIE"
   '
   GRID.Redraw = False
20 sqlx1$ = "SELECT FECHMOV, CodiMovi, DOPRILAR,  DoSecLar, IdenLote, ReferCor,  DEPOMOVI,  CANTINGRE, CANTSALID,NumUsuar"
   sqlx1$ = sqlx1$ + " FROM " + BASEAC$ + " WITH(NOLOCK)"
   sqlx1$ = sqlx1$ + " WHERE " + condi0$
   sqlx1$ = sqlx1$ + " AND FECHMOV >= '" & Desde & "' "
   sqlx1$ = sqlx1$ + " AND FECHMOV <= '" & Hasta & "' "
   'SQLX1$ = SQLX1$ + " ORDER BY FechMov , FeReMovi ASC "
   If EXISTE_CAMPO%("MOVISTO", "MOVISTO_ID") > 0 Then 'POR QUE FEREMOVI NO GUARDA LOS SEGUNDOS
        sqlx1$ = sqlx1$ + " ORDER BY FeReMovi, MOVISTO_ID"
     Else
        sqlx1$ = sqlx1$ + " ORDER BY FechMov , FeReMovi ASC "
   End If


'   Dim MOVIVIETMP As ADODB.Recordset
'   Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
   Set MOVIVIETMP = READSET(sqlx1$)
   GRID.Rows = MOVIVIETMP.RecordCount + 1

   With MOVIVIETMP
      Do While Not .EOF
              If CVINT(!FechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
         IDLOT$ = SAFETEXT$(!idenlote)
         If XVALO(Label8) > 0 Then
            If PROVELOTE%(IDLOT$) <> XVALO(Label8) Then GoTo 65
         End If
             '
         iRow& = iRow& + 1
         If iRow& > GRID.Rows - 1 Then
            GRID.Rows = iRow& + 1
         End If
         GRID.TextMatrix(iRow&, 0) = FECHATEX$(CVINT(!FechMov))
         GRID.TextMatrix(iRow&, 1) = SAFETEXT$(!CODIMOVI)
         GRID.TextMatrix(iRow&, 2) = SAFETEXT$(!DOPRILAR)
         GRID.TextMatrix(iRow&, 3) = SAFETEXT$(!DOSECLAR)
            IDLOT$ = SAFETEXT$(!idenlote)
         GRID.TextMatrix(iRow&, 4) = IDLOT$
         GRID.TextMatrix(iRow&, 5) = SAFETEXT$(!ReferCor)
         GRID.TextMatrix(iRow&, 6) = XVALO(!depomovi)
            CAINGRE# = XVALO(!CANTINGRE)
            CASALI# = XVALO(!CANTSALID)
            CASALDO# = CASALDO# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
         GRID.TextMatrix(iRow&, 7) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
         GRID.TextMatrix(iRow&, 10) = XVALO(!NumUsuar)
'            Call CompletaCombo(IDLOT$)
            Combo1.AddItem IDLOT$
65    .MoveNext
      Loop
   End With
   MOVIVIETMP.Close
   Set MOVIVIETMP = Nothing
   If BASEAC$ = "MOVIVIE" Then
      BASEAC$ = "MOVISTO"
      GoTo 20
   End If
   '
   Call Duplicados1(Combo1)
   GRID.Redraw = True
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
   On Error Resume Next
   GRID.TopRow = TPXI&
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   '
End Sub



Private Sub GRID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = GRID.Row

          H& = 0: If GRID.COL < GRID.Cols Then H& = GRID.COL + 1
10        For U& = J& To GRID.Rows - 1
             HI& = H&: If U& > J& Then HI& = 0
             For i& = HI& To GRID.Cols - 1
              If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                GRID.Row = U&: GRID.COL = i&
                On Error Resume Next
                 VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If GRID.TopRow < GRID.Row - VISIROWS Then
                      GRID.TopRow = GRID.Row - VISIROWS
                    End If
                 End If
                GRID.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub


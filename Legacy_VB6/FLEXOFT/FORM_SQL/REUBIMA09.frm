VERSION 5.00
Begin VB.Form REUBIMA09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Entre Ubicaciones"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   75
      TabIndex        =   13
      Top             =   1035
      Width           =   10980
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3420
         Left            =   6400
         TabIndex        =   26
         Top             =   330
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3435
         Left            =   8150
         TabIndex        =   17
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   3375
         Left            =   9150
         TabIndex        =   16
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   3405
         Left            =   2070
         TabIndex        =   15
         Top             =   330
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   3390
         Left            =   4635
         TabIndex        =   14
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         ItemData        =   "REUBIMA09.frx":0000
         Left            =   165
         List            =   "REUBIMA09.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   540
         Width           =   10710
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Codigo           Descripcion             Id.Bulto         Lote             F.Venc.       Cantidad"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   19
         Top             =   315
         Width           =   10695
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6585
      Left            =   11160
      ScaleHeight     =   6525
      ScaleWidth      =   915
      TabIndex        =   8
      Top             =   -30
      Width           =   975
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Picture         =   "REUBIMA09.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ficha Kardex"
         Top             =   1050
         Width           =   645
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Picture         =   "REUBIMA09.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1815
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   120
         Picture         =   "REUBIMA09.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   2985
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -135
            TabIndex        =   11
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         Picture         =   "REUBIMA09.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save"
         Top             =   3825
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "REUBIMA09.frx":0A6C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3255
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6060
      TabIndex        =   7
      Top             =   540
      Width           =   175
   End
   Begin VB.TextBox Text8 
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
      IMEMode         =   3  'DISABLE
      Left            =   5055
      TabIndex        =   6
      Top             =   555
      Width           =   1050
   End
   Begin VB.CommandButton Command25 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6060
      TabIndex        =   5
      Top             =   165
      Width           =   175
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   5055
      TabIndex        =   4
      Top             =   180
      Width           =   1050
   End
   Begin VB.CommandButton Command30 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3450
      TabIndex        =   1
      Top             =   255
      Width           =   175
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
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
      IMEMode         =   3  'DISABLE
      Left            =   2865
      MaxLength       =   3
      TabIndex        =   0
      Top             =   270
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubic. Destino:"
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
      Left            =   3495
      TabIndex        =   23
      Top             =   645
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6360
      TabIndex        =   22
      Top             =   540
      Width           =   4665
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6345
      TabIndex        =   21
      Top             =   180
      Width           =   4665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubic. Origen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3855
      TabIndex        =   20
      Top             =   225
      Width           =   1140
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   225
      TabIndex        =   3
      Top             =   615
      Width           =   3405
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   330
      Width           =   1065
   End
End
Attribute VB_Name = "REUBIMA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub
Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub
Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   '
   If DERACCE%("SAVEREUB", "Registrar nueva ubicacion") < 2 Then GoTo 99
   '
   If TRIM$(Text8) = "" Then
      COMUNI ("Falta Ingresar Ubicacion Destino")
      GoTo 99
   End If
   '
   DEP% = XVALO(Text6)
   UBICX$ = TRIM$(Text8)
   '
   DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    If DEP% = DEPOTRANS% Then
        COMUNI ("Depósito Seleccionado No Válido\Configurado como Material en Transito ")
       GoTo 99
    End If
   CONDI$ = "CODIUBIC ='" & UBICX$ & "'"
   If CantRegistros("UBIPRE", CONDI$) < 1 Then
      Call COMUNI("Ubicacion Destino Inexistente")
      GoTo 99
   End If
   '
   CONDI$ = "DEPOUBIC = " & DEP% & " AND UBICFIS = '" & UBICX$ & "'"
   If CantRegistros("LOTINGRE", CONDI$) > 0 Then
      COMUNI ("La Ubicacion Destino Debe Estar Libre")
      GoTo 99
   End If
   '
   Dim CheckSel As Boolean
   CheckSel = False
   For U& = 0 To List1.ListCount - 1
     If List1.Selected(U&) = True Then
        CheckSel = True
        LOTE$ = ""
        LOTE$ = TRIM$(Mid$(List1.List(U&), 42, 16))
        CONDI$ = "IDENLOTE ='" & LOTE & "' "
        Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBICX$, REG&)
     End If
   Next U&
   '
   If CheckSel = False Then
      Call COMUNI("Debe Marcar Algun Item de la Lista")
      GoTo 99
   End If
   '
   Call GENTAUBI(DEP%)
   Text2 = ""
   Text8 = ""
   List1.Clear
   Call COMUNI("Cambio de Ubicacion Realizado Correctamente")
   
99
   Command13.Enabled = True
   '
End Sub

Private Sub COMMAND15_CLICK()
   Command7.Enabled = False
   FIKARDEX09.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   FIKARDEX09.Show 1
   Command20.Enabled = True
End Sub

Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Text6 = VALACT1$("TADEPOSI")
End Sub
Private Sub Command25_Click()
    '
    DEPO% = XVALO(Text6)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text2 = UBISELE$
'       Label2 = ECOARCH$("TAUBICA", Text2)
        '
        CONDI$ = "Codiubic = '" & Text2 & "'"
        Label12 = VALOBADA("UBIPRE", "Referencia", CONDI$, "NOMESS")
        '
      End If
End Sub
Private Sub Command2_Click()
    DEPO% = XVALO(Text6)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text8 = UBISELE$
'       Label3 = ECOARCH$("TAUBICA", Text8)
       CONDI$ = "Codiubic = '" & Text8 & "'"
       Label3 = VALOBADA("UBIPRE", "Referencia", CONDI$, "NOMESS")
    End If
End Sub

Private Sub List1_DblClick()
    LINEA$ = List1.List(List1.ListIndex)
    IDLOT$ = Mid$(LINEA$, 42, 16)
    CODX$ = Mid$(LINEA$, 1, 16)
    CONSUMLOT07.Label4 = CODX$
    CONSUMLOT07.Label1 = IDLOT$
    CONSUMLOT07.Show 1
End Sub

Private Sub Text2_Change()
   '
   DEP% = XVALO(Text6)
   If DEP% > 0 Then
'     UBICORIG$ = ECOARCH$("TAUBICA", TRIM$(Text2))
      ' Lectura desde el SQL
      CONDI$ = "Codiubic = '" & Text2 & "'"
      UBICORIG$ = VALOBADA("UBIPRE", "Referencia", CONDI$, "NOMESS")
      Label2 = UBICORIG$
      If UBICORIG$ <> "" Then Call CARGARLISTA
   End If
   '
End Sub

Private Sub Text6_Change()
   DEP% = XVALO(Text6)
   If DEP% > 0 Then
      Label38 = ECOARCH$("TADEPOSI", Chr$(DEP%))
      Text2 = ""
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command30_Click
End Sub

Sub CARGARLISTA()

       List1.Clear
       DEP% = XVALO(Text6)
       UBICORIG$ = TRIM$(Text2)
'       Dim LOTES As ADODB.Recordset
'       SQLX$ = "SELECT * FROM LOTINGRE WHERE UBICFIS = '" & UBICORIG$ & "' "
'       SQLX$ = SQLX$ + " AND DEPOUBIC = " & DEP%
'       SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0 "
'       Set LOTES = FLEXCONN.Execute(SQLX$)
'       With LOTES
'         Do While Not (.EOF)
'           LINEA$ = ""
'           CIXI% = XVALO(!COD_INTE)
'           CODX$ = CODEXT$(CIXI%)
'           DESCX$ = DESCRIT0(CIXI%)
'           FECHA$ = FECHATEX(CVINT(!FechFab))
'           LOTE$ = SAFETEXT(!LoteCol_Prov)
'           CANTX$ = FORMATNUM(XVALO(!CANSALDO), "F13." & CNTDC$)
'
'           Call REPLA(LINEA$, CODX$, 1, 16)
'           Call REPLA(LINEA$, DESCX$, 17, 24)
'           Call REPLA(LINEA$, SAFETEXT(!idenlote), 42, 16)
'           Call REPLA(LINEA$, LOTE$, 58, 16)
'           Call REPLA(LINEA$, FECHA$, 74, 8)
'           Call REPLA(LINEA$, CANTX$, 83, 12)
'           List1.AddItem LINEA$
'           .MoveNext
'         Loop
'       End With
       CONDI$ = "UBICFIS = '" & UBICORIG$ & "' "
       CONDI$ = CONDI$ + " AND DEPOUBIC = " & DEP%
       CONDI$ = CONDI$ + " AND (CanAlta - CanConsu) > " & FORMATNUM(FRACMIN(CIXI%), "F13.6")
       FORMATO$ = "F13." & CNTDC$
       Call CARGALISTA(List1, "LOTINGRE+MASTER/COD_INTE=CODINT", "COD_INTE/C16;DESCRIPCION/A23;IDENLOTE/A16;LOTECOL_PROV/A16;FECHFAB/D8;CANSALDO/" & FORMATO$, CONDI$)
End Sub

Private Sub Text8_Change()
   '
   DEP% = XVALO(Text6)
   If DEP% > 0 Then
      ' Label3 = ECOARCH$("TAUBICA", TRIM$(Text8))
      ' Lectura desde el SQL
      CONDI$ = "Codiubic = '" & Text8 & "'"
      Label3 = VALOBADA("UBIPRE", "Referencia", CONDI$, "NOMESS")
      '
   End If
   '
End Sub

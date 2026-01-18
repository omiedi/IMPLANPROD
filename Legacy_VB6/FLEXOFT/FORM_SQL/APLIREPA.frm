VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form APLIREPA 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Aplcación de Pagos a Técnicos Externos."
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Documentos a Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   7320
      Width           =   10700
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   9720
         Picture         =   "APLIREPA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto de Reparación"
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label77 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   7440
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   6720
         TabIndex        =   22
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo a Aplicar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Documentos a Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   4860
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2385
         TabIndex        =   8
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1005
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
         Height          =   300
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Limpiar Casillero de Filtro"
         Top             =   480
         Width           =   2130
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Selección de Técnico Externo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   10620
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor Asociado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   8280
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   7080
         TabIndex        =   5
         ToolTipText     =   "Doble Click para Limpiar Casillero de Filtro"
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Doble Click para Limpiar Casillero de Filtro"
         Top             =   480
         Width           =   1410
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   2685
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Boton Derecho Elimina Aplicación (Irreversible)"
      Top             =   4560
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4736
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID0 
      Height          =   2325
      Left            =   0
      TabIndex        =   21
      Top             =   2040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuelimiApli 
         Caption         =   "Eliminar Aplicación"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "APLIREPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Private Sub Command1_Click()
    Command1.Enabled = False
    NC = XVALO(Label1)
    If NC < 1 Then
       Call COMUNI("No se Ha Elegido Ningún Técnico")
       GoTo 99
    End If
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    FECHALU = FETEXCOR1$(FECHALIM%)
    
    CAMPOS$ = "NDOC/A0;Documento/A20;Imp.Orig./F12.2;Aplicado/F12.2;Pendiente/F12.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    '
    SQLX$ = "SELECT NUPROVE,Niv_Clasi,ValAbsComp,CODICOM_LAR,FECHEMISI "
    SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK)  "
    SQLX$ = SQLX$ + " WHERE (TipoMovim = 'OPAG' OR (TipoMovim = 'AJUP' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + " AND NuProve = " & NC & " "
    SQLX$ = SQLX$ + " AND FechEmisi >= '" & FECHALU & "'"
    SQLX$ = SQLX$ + " AND (LEFT(CODICOM_LAR, 2) = 'OP' OR LEFT(CODICOM_LAR,2) = 'CC' )"
    'SQLX$ = SQLX$ + " AND not  exists (SELECT 1 FROM APLIREPA where APLIREPA.DOC_APLI = CAJABANC.CODICOM_LAR  )" 'QUE NO EXISTEN EN LA TABLA APLIREPA CON MISMO NUMERO DE ORDEN DE PAGO
    
    Set RS = READSET(SQLX$)
    '
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    With RS
        Do While Not .EOF
            If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
            IMPO# = XVALO(!ValAbsComp)
            IMPOS$ = FORMATNUM$(IMPO#, "F12.2")
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            DOCU$ = SAFETEXT$(!CODICOM_LAR)
            YAAPLI# = XVALO(VALOBADA("APLIREPA", "SUM(IMP_APLI)", "DOC_APLI= '" & DOCU$ & "'", "NOMESS"))
            SALDOAAPLI# = ROUND(XVALO(IMPOS$) - YAAPLI#, 2)
            If SALDOAAPLI# > 0 Then
                JJ& = JJ& + 1
                FRMZELECHO.msf2.Rows = JJ& + 1
                FRMZELECHO.msf2.TextMatrix(JJ&, 1) = NPX&
                FRMZELECHO.msf2.TextMatrix(JJ&, 2) = DOCU$
                FRMZELECHO.msf2.TextMatrix(JJ&, 3) = IMPOS$
                FRMZELECHO.msf2.TextMatrix(JJ&, 4) = YAAPLI#
                FRMZELECHO.msf2.TextMatrix(JJ&, 5) = SALDOAAPLI#
            End If
70      .MoveNext
        Loop
    End With
    Screen.MousePointer = 1
    FRMZELECHO.Width = 9000
    FRMZELECHO.Caption = "Ordenes de Pago Pendientes de Aplicación"
    FRMZELECHO.Show 1
    '
    NORPAG$ = TRIM(RESP_ZELE_VECT(2))
    If NORPAG$ <> TRIM$(Label8.Caption) Then
        IMPOS$ = FORMATNUM$(XVALO(RESP_ZELE_VECT(5)), "F12.2")
        Label8 = NORPAG$
        Label5 = IMPOS$
        Label7 = IMPOS$
    End If
   
99 Command1.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   
   For i& = 1 To GRID0.Rows - 1
      If XVALO(GRID0.TextMatrix(i&, 3)) > 0 Then
        SQLX$ = "INSERT INTO APLIREPA "
        SQLX$ = SQLX$ + " (CodiEmpr,NuProve,TEC_EXT,NumRepa,IMP_MOBRA,FECHAPLI,IMP_APLI,DOC_APLI,SALDO_OP)"
        SQLX$ = SQLX$ + " VALUES(" & CodiEmp%
        SQLX$ = SQLX$ + " ," & XVALO(Label1)
        SQLX$ = SQLX$ + " ," & XVALO(Label26)
        SQLX$ = SQLX$ + " ," & XVALO(GRID0.TextMatrix(i&, 1)) 'NREP
        SQLX$ = SQLX$ + " ," & FORMATNUM$(XVALO(GRID0.TextMatrix(i&, 2)), "F16.4") 'IMPO ORIGINAL
        SQLX$ = SQLX$ + " ,'" & CVDAT(HOY(HO$)) & "'" '
        SQLX$ = SQLX$ + " ," & FORMATNUM$(XVALO(GRID0.TextMatrix(i&, 3)), "F16.4") 'IMPO APLICADO
        SQLX$ = SQLX$ + " ,'" & Label8.Caption & "'"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
      End If
   Next i&
   TECNICO = XVALO(Label26)
   Label26 = ""
   Label26 = TECNICO 'REFRESH
   
99 Command14.Enabled = True
End Sub


Private Sub Command8_Click()
  nt = MUESTRATECNICO
  If nt > 0 Then
    Label26.Caption = nt
  End If

End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, REG&, COL%)
  End If
End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&

        PopupMenu MNUARCHIVO
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnuelimiApli.Visible = Not (mnuelimiApli.Visible)
  '
End Sub

Private Sub Label8_Change()
    Label5 = ""
    Label7 = ""
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuelimiApli_Click()

    ID = TRIM$(GRID.TextMatrix(GLOBALFRM_FILA&, 6))
    If XVALO(ID) > 0 Then
       If DERACCE%("DESAREPA", "DESAPLICAR ORDEN DE PAGO A REPARACIÓN") >= 2 Then
        SQLX$ = "DELETE FROM APLIREPA WHERE APLIREPA_ID = " & ID
        FLEXCONN.Execute (SQLX$)
        Call Label26_Change
       End If
    Else
       Call COMUNI("Imposible Eliminar !!!\Fila Seleccionada No Contiene Aplicación")
    End If
End Sub

Private Sub GRID0_Click()
    Static MODO  As Boolean
    i& = GRID0.MouseRow
    J& = GRID0.MouseCol

    If (GRID0.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID0.COL = GRID0.MouseCol
            GRID0.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID0.COL = GRID0.MouseCol
            GRID0.Sort = 1
            MODO = True
        End If
    End If
    If J& = 0 Then
       If XVALO(Label7.Caption) <= 0 And GRID0.TextMatrix(GRID0.Row, 0) = "" Then Exit Sub
       Call COLOREAR_GRILLA_SOLAFILA(GRID0, vbYellow, GRID0.Row, 0, 0, 0)
       Call APLICAIMPO(GRID0.Row)
    End If

End Sub
Sub APLICAIMPO(REG&)
    If REG& < 1 Then Exit Sub
    
    If GRID0.TextMatrix(REG&, 0) = "*" Then
         If XVALO(Label7) <= 0 Then Exit Sub
         SALDO = GRID0.TextMatrix(REG&, 5)
         If XVALO(Label7) > XVALO(SALDO) Then GRID0.TextMatrix(REG&, 3) = SALDO
         If XVALO(Label7) <= XVALO(SALDO) Then GRID0.TextMatrix(REG&, 3) = FORMATNUM$(XVALO(Label7), "F12.2")
         GRID0.TextMatrix(REG&, 4) = Label10
         GRID0.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(SALDO) - XVALO(GRID0.TextMatrix(REG&, 3)), "F14.2")
         Label7 = FORMATNUM$(XVALO(Label7) - XVALO(GRID0.TextMatrix(REG&, 3)), "F12.2")
    Else
          GRID0.TextMatrix(REG&, 2) = FORMATNUM$(XVALO(GRID0.TextMatrix(REG&, 3)) + XVALO(GRID0.TextMatrix(REG&, 5)), "F12.2")
          GRID0.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(GRID0.TextMatrix(REG&, 3)) + XVALO(GRID0.TextMatrix(REG&, 5)), "F12.2")
          Label7 = FORMATNUM$(XVALO(Label7) + XVALO(GRID0.TextMatrix(REG&, 3)), "F12.2")
          GRID0.TextMatrix(REG&, 3) = ""
    End If
End Sub

Private Sub Label26_Change()
  nt& = XVALO(Label26.Caption)
  If nt& > 0 Then
    SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)  WHERE LISTA = 'TECNIEXT' AND CAMPO1 = '" & FORMATOCON0$(nt&, 9) & "'"
    Set RS = READSET(SQLX$)
    Label14.Caption = SAFETEXT$(RS!CAMPO2) 'NOMBRE DEL TECNICO
    Label1.Caption = SAFETEXT$(RS!CAMPO5) 'nro DEL proveedor
    Label2.Caption = SAFETEXT$(RS!CAMPO6)  'NOMBRE DEL proveedor
    Call COMPLETARGRID(XVALO(Label26.Caption))
    Label8 = ""
    Label5 = ""
  Else
    Label14.Caption = "" 'NOMBRE DEL TECNICO
    Label1.Caption = ""  'nro DEL proveedor
    Label2.Caption = ""  'NOMBRE DEL proveedor
    Label8 = ""
    Label5 = ""
    Label77 = ""
    GRID.Rows = 1
  End If

End Sub
Sub COMPLETARGRID(NP)
   CAMPOS$ = "N.REP/F7;Imp.M.Obra/F12.2;Imp.Aplic./F12.2;Doc_Aplic/A24;Saldo/F14.2;ID/A0"
   CAMPOS2$ = "N.REP/F7;Imp.M.Obra/F12.2;Imp.Aplic./F12.2;Doc_Aplic/A0;Saldo/F14.2;ID/A0"
   Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me)
   Call CARGA_ENCABEZADO(Me.GRID0, CAMPOS2$, Me)
   Label77 = ""
   J& = 0: i& = 0
   SQLX$ = "SELECT R.NUMREPA,R.MOBRA_EN_GTIA, A.IMP_APLI, A.DOC_APLI,A.APLIREPA_ID  "
   SQLX$ = SQLX$ + "  FROM REPARA R WITH(NOLOCK) "
   SQLX$ = SQLX$ + "  LEFT JOIN APLIREPA A  "
   SQLX$ = SQLX$ + "  ON R.NUMREPA  = A.NUMREPA  AND R.TEC_EXT  = A.TEC_EXT "
   SQLX$ = SQLX$ + "  WHERE  R.TEC_EXT = " & XVALO(Label26.Caption)
   
   SQLX$ = SQLX$ + "  ORDER BY R.NUMREPA"
   Set RS = READSET(SQLX$)
   With RS
      Do While Not .EOF
         J& = J& + 1
'==========================================================================================================================
         'CARGO LA GRILLA INFERIOR CON EL HISTORIAL
'==========================================================================================================================
         GRID.Rows = J& + 1
         REP& = XVALO(!NumRepa)
         GRID.TextMatrix(J&, 1) = XVALO(!NumRepa)
         GRID.TextMatrix(J&, 2) = FORMATNUM$(XVALO(!MOBRA_EN_GTIA), "F12.2")
         GRID.TextMatrix(J&, 3) = FORMATNUM$(XVALO(!IMP_APLI), "F12.2")
         GRID.TextMatrix(J&, 4) = SAFETEXT$(!DOC_APLI)
         If REP& = REP_ANT& Then
            ACUMUIMPAPLI = ACUMUIMPAPLI + XVALO(!IMP_APLI)
         Else
            ACUMUIMPAPLI = XVALO(!IMP_APLI)
         End If
         GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!MOBRA_EN_GTIA) - ACUMUIMPAPLI, "F14.2")
         GRID.TextMatrix(J&, 6) = SAFETEXT(!APLIREPA_ID)

'==========================================================================================================================
         
'==========================================================================================================================
         'CARGO LA GRILLA SUPERIOR CON SOLO LAS QUE TIENEN SALDO PENDIENTE Y SI SON VARIOS REGISTROS LOS ACUMULO EN UNA FILA.
'==========================================================================================================================
         If REP& = REP_ANT& Then
            'SI ES LA MISMA REPA ACUMULO
            ACUMUIMPAPLI0 = ACUMUIMPAPLI0 + XVALO(!IMP_APLI)
         Else
            'SI CAMBIA LA REPARACION AGREGO UN REGISTRO
            i& = i& + 1
            GRID0.Rows = i& + 1
            GRID0.TextMatrix(i&, 1) = XVALO(!NumRepa)
            GRID0.TextMatrix(i&, 2) = FORMATNUM$(XVALO(!MOBRA_EN_GTIA), "F12.2")
            ACUMUIMPAPLI0 = XVALO(!IMP_APLI)
         End If
         GRID0.TextMatrix(i&, 5) = FORMATNUM$(XVALO(!MOBRA_EN_GTIA) - ACUMUIMPAPLI0, "F14.2")
         If XVALO(GRID0.TextMatrix(i&, 5)) <= 0 Then i& = i& - 1
'==========================================================================================================================

         REP_ANT& = REP&
       .MoveNext
     Loop
   End With
   Label77.Caption = FORMATNUM$(SUMACOLUMNA#(GRID0, 5, 1), "F12.2")
End Sub


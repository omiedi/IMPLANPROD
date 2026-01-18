VERSION 5.00
Begin VB.Form TRAZINVER09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Trazabilidad Inversa"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "                                                                                                                  "
      Height          =   2160
      Index           =   1
      Left            =   8460
      TabIndex        =   31
      Top             =   1695
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "                "
      Height          =   2145
      Left            =   3615
      TabIndex        =   29
      Top             =   4755
      Width           =   15
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Destino / Ulilización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   150
      TabIndex        =   24
      Top             =   1230
      Width           =   10740
      Begin VB.Frame Frame9 
         Caption         =   "                                                                                                                  "
         Height          =   2160
         Index           =   2
         Left            =   9345
         TabIndex        =   32
         Top             =   435
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "                                                                                                                  "
         Height          =   2160
         Index           =   0
         Left            =   7155
         TabIndex        =   30
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   2160
         Left            =   3195
         TabIndex        =   26
         Top             =   465
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   2145
         Left            =   1650
         TabIndex        =   25
         Top             =   480
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
         Height          =   2040
         ItemData        =   "TRAZINVER09.frx":0000
         Left            =   195
         List            =   "TRAZINVER09.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   10380
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"TRAZINVER09.frx":0004
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   28
         Top             =   360
         Width           =   10380
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Consultar"
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
      Left            =   120
      TabIndex        =   23
      Top             =   780
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   8640
      TabIndex        =   18
      Top             =   90
      Width           =   2235
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
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
         Left            =   825
         TabIndex        =   22
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
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
         Left            =   150
         TabIndex        =   21
         Top             =   735
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   825
         TabIndex        =   20
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -315
         TabIndex        =   19
         Top             =   345
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   300
      Left            =   3135
      TabIndex        =   17
      Top             =   345
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Height          =   300
      Left            =   7380
      TabIndex        =   11
      Top             =   390
      Width           =   175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   165
      TabIndex        =   6
      Top             =   4260
      Width           =   10695
      Begin VB.Frame Frame5 
         Caption         =   "                "
         Height          =   2145
         Left            =   2025
         TabIndex        =   10
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "                                                                                                                  "
         Height          =   2160
         Left            =   6330
         TabIndex        =   9
         Top             =   435
         Width           =   15
      End
      Begin VB.ListBox List2 
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
         Height          =   2040
         ItemData        =   "TRAZINVER09.frx":0096
         Left            =   210
         List            =   "TRAZINVER09.frx":0098
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   10305
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"TRAZINVER09.frx":009A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   8
         Top             =   360
         Width           =   10305
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000040&
      Height          =   8400
      Left            =   11100
      ScaleHeight     =   8340
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1020
      Begin VB.CommandButton Command2 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRAZINVER09.frx":0136
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2535
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRAZINVER09.frx":0440
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRAZINVER09.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   2
         Top             =   6465
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "TRAZINVER09.frx":0894
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6690
         Width           =   1515
      End
   End
   Begin VB.Label Label36 
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
      Left            =   1425
      TabIndex        =   16
      Top             =   795
      Width           =   6135
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5175
      TabIndex        =   15
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label32 
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
      Left            =   1440
      TabIndex        =   14
      Top             =   345
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Componente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -465
      TabIndex        =   13
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label Label15 
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
      Left            =   5655
      TabIndex        =   12
      Top             =   405
      Width           =   1740
   End
End
Attribute VB_Name = "TRAZINVER09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command11_Click()
    Hide
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   TTXX$ = List2.List(List2.ListIndex)
   CODILLA$ = TRIM$(Mid$(TTXX$, 1, 16))
   If CODILLA$ = "" Then CODILLA$ = TRIM$(Label32)
   FIKARDEX07.Text1 = CODILLA$
   FIKARDEX07.Show 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AMASTO07", "Maestro de Items de Stock") > 1 Then
       AMASTOK07.Text1 = Label11
       AMASTOK07.Show 1
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  '
  If TRIM(Label32) <> "" Then
   Call SELECHO("!LILOTSE")
   LOTE$ = TRIM$(VALACT1$("!LILOTSE"))
   If LOTE$ <> "" Then
      Label15 = LOTE$
   End If
  End If
  '
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    '
    CONDI$ = "Trazable = 1"
    Call CARGALISEL("!ECOMAST", "Master", "Codigo/A24;Descripcion/A64", CONDI$)
    Call FRESHECHO("!ECOMAST")
    '
End Sub
Private Sub Command1_Click()
  
   
   Call SELECHO("!ECOMAST")
   COD$ = TRIM$(VALACT1$("!ECOMAST"))
   If COD$ <> "" Then
     Label32 = COD$
     Label15 = ""
     Label1 = ""
     Label33 = ""
   End If
End Sub
  


Private Sub Label27_Click()
End Sub

Private Sub Label28_Change()
   Call LIMPIAR
   NORFA$ = Label28
   Ci% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   If Ci% > 0 Then
     Label11 = CODEXT(Ci%)
     Label13 = DESCRIT$(Ci%)
     CONDI$ = " DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'"
     Call ENUMERA("MOVISTO", "OrdConsum", CONDI$)
     Call CARGALISTA(List1, "MOVISTO", "Cod_Exte/A16; Descrip/A43; IdenLote/A16; OrdConsum/F12", CONDI$)
     '
10   ' ELIMINA LOS ITEMS NO TRAZABLES
     Screen.MousePointer = 11
     i& = 1
     Do While i& <= List1.ListCount
        CIXI% = CODINT%(TRIM$(Left$(List1.List(i& - 1), 16)))
        If ESTRAZABLE%(CIXI%) < 1 Then
           List1.RemoveItem (i& - 1)
         Else
           i& = i& + 1
        End If
     Loop
     Screen.MousePointer = 1
     '
     'RECOREEMOS PARA QUE NO QUEDEN REPETIDOS EN EL LIST
     CIXIA% = 0: LOTA$ = "":  i& = 1
     Do While i& <= List1.ListCount
       LOT$ = TRIM$(Mid$(List1.List(i& - 1), 59, 16))
       CIXI% = CODINT%(TRIM$(Left$(List1.List(i& - 1), 16)))
       If CIXI% <> CIXIA% And (LOT$ = "" Or LOT$ <> LOTA$) Then
           i& = i& + 1
           CIXIA% = CIXI%
           LOTA$ = LOT$
         Else
           List1.RemoveItem (i& - 1)
       End If
     Loop
     '
   End If

End Sub

Sub LIMPIAR()

     Label11 = ""
     Label13 = ""
     Label5 = ""
     Label3 = ""
     Label6 = ""
     Label7 = ""
     List1.Clear
     List2.Clear
    
End Sub


Private Sub Command3_Click()
   '
   List2.Clear
   List1.Clear
   '
   Call APERBASDAT("STOCKS")
   LOTE$ = TRIM$(Label15)
   If LOTE$ = "" Then
      Call MENSERR(24, "Componente no Trazado")
      Exit Sub
   End If
   SQLX$ = " SELECT DISTINCT DOPRICOR FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & LOTE$ & "' "
   SQLX$ = SQLX$ + "AND CODIMOVI = 'CF'"
   '
   Dim LOTETEMP As ADODB.Recordset
   Set LOTETEMP = New ADODB.Recordset
   LOTETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
   '
   With LOTETEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
       NUORFA$ = TRIM$(SAFETEXT$(!DOPRICOR))
'      SQLX$ = "SELECT * FROM TANUMSE "
'       SQLX$ = SQLX$ + "WHERE IDSUBOR = '" & NUORFA$ & "'"
       '
'       Dim ORFATEMP As ADODB.Recordset
'       Set ORFATEMP = New ADODB.Recordset
'       ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
'       '
'       With ORFATEMP
'         On Error Resume Next
'         .MoveFirst
'         If Err < 1 Then
'           Do While Not .EOF
'              CIXI% = XVALO(!COD_INTE)
'              NUMSER$ = SAFETEXT(!NUMEROSERIE)
'              ' VERIFICA QUE EL NUMERO DE SERIE NO ESTE PREVIAMENTE CARGADO
'              For I = 1 To List2.ListCount
'                NSER$ = Mid$(List2.List(I - 1), 60, 20)
'                If TRIM$(NUMSER$) = TRIM$(NSER$) Then
'                  GoTo 30
'                End If
'              Next I
'              '
              CONDI2$ = "IDSUBOR='" & NUORFA$ & "'"
              CANPROY$ = FORMATNUM$(VALOBADA("ORDEFABR", "CANPROY", CONDI2$), "F7." & CNTDC$)
              CANAPRO$ = FORMATNUM$(VALOBADA("ORDEFABR", "CANAPRO", CONDI2$), "F7." & CNTDC$)
              'CANRECH$ = FORMATNUM$(VALOBADA("ORDEFABR", "CANRECH", CONDI2$), "F7." & CNTDC$)
              SALDO$ = XVALO(CANPROY$) - XVALO(CANAPRO$)
              
              CIXI% = VALOBADA("ORDEFABR", "COD_INTE", CONDI2$)
              TXTX$ = Space(128)
              Call REPLA(TXTX$, NUORFA$, 1, 13) 'Nro OF
              Call REPLA(TXTX$, CODEXT$(CIXI%), 15, 16)
              Call REPLA(TXTX$, DESCRIT$(CIXI%), 30, 39)
              Call REPLA(TXTX$, CANPROY$, 70, 10)
              Call REPLA(TXTX$, CANAPRO$, 80, 10)
              Call REPLA(TXTX$, FORMATNUM$(SALDO$, "F7." & CNTDC$), 89, 10)
              List1.AddItem (TXTX$)
'30            .MoveNext
'           Loop
'         End If
'       End With
       .MoveNext
       Loop
     End If
   End With
   ORFATEMP.Close
   Set ORFATEMP = Nothing
   LOTETEMP.Close
   Set LOTETEMP = Nothing
    
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label15_Change()
   If TRIM$(Label15) <> "" Then
      List1.Clear
      CONDI$ = "IDENLOTE='" & TRIM$(Label15) & "'"
      Label1 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechAlta", CONDI$)))
      Label33 = FORMATNUM$(VALOBADA("LOTINGRE", "CanAlta", CONDI$), "F7." & CNTDC$)
   End If
End Sub

Private Sub Label15_DblClick()
   Call Command5_Click
End Sub

Private Sub Label32_Change()
     Label36 = DESCRIT$(CODINT(Label32))
     List1.Clear
     List2.Clear
     CIXI% = CODINT%(Label32)
     If CIXI% > 0 Then
        Call CREALOTES(CIXI%)
        Call ARMALISLOTES(CIXI%)
     End If
End Sub

Private Sub Label32_DblClick()
   Call Command1_Click
End Sub

Private Sub List1_Click()
   Screen.MousePointer = 11
   List2.Clear
   NUORFA$ = Mid(List1.TEXT, 1, 13)
   
   SQLX$ = "SELECT * FROM TANUMSE "
   SQLX$ = SQLX$ + "WHERE IDSUBOR = '" & NUORFA$ & "'"
       
   Dim ORFATEMP As ADODB.Recordset
   Set ORFATEMP = New ADODB.Recordset
   ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
      
   '
   With ORFATEMP
     If .BOF = True And .EOF = True Then GoTo 99
     .MoveFirst
     Do While Not .EOF
        CIXI% = XVALO(!COD_INTE)
        NUMSER$ = SAFETEXT(!NUMEROSERIE)
        'FECHA$ = SAFETEXT(!FECHALTA)
        FECHA$ = FECHATEX(CVINT(!FECHALTA))
        REMITO$ = SAFETEXT(!NUMREMITO)
        NC = XVALO(!NUMECLI)
        ' VERIFICA QUE EL NUMERO DE SERIE NO ESTE PREVIAMENTE CARGADO
        For i = 1 To List2.ListCount
          NSER$ = Mid$(List2.List(i - 1), 60, 20)
          If TRIM$(NUMSER$) = TRIM$(NSER$) Then
             GoTo 30
          End If
        Next i
        '
        TXTX$ = Space(128)
        Call REPLA(TXTX$, NUMSER$, 1, 13) 'Nro OF
        Call REPLA(TXTX$, FECHA$, 20, 8) 'FECHA
        Call REPLA(TXTX$, REMITO$, 32, 24) 'REMITO
        Call REPLA(TXTX$, rasocli$(NC), 60, 24) 'Cliente
        List2.AddItem TXTX$
    
30      .MoveNext
      Loop
    End With
99  Screen.MousePointer = 1
End Sub

Private Sub Text1_Change()
'
Label28.Caption = VALOBADA("TANUMSE", "IDSUBOR", "NUMEROSERIE= '" & TRIM$(Text1.TEXT) & "'")
If Label28.Caption = "" Then
   CIXI% = XVALO(VALOBADA("TANUMSE", "cod_Inte", "NUMEROSERIE= '" & TRIM$(Text1.TEXT) & "'"))
   If CIXI > 0 Then
     Label11.Caption = CODEXT$(CIXI%)
     Label13.Caption = DESCRIT$(CIXI%)
     Call MENSERR(24, "Número de Serie Probablemente Cargado en Forma Manual")
   End If
End If
'NSERIE$ = TRIM$(Text1)
'OF$ = SAFETEXT(VALOBADA("TANUMSE", "IDSUBOR", "NUMEROSERIE='" & NSERIE$ & "'"))
'If TRIM(OF$) <> "" Then Label28 = TRIM$(OF$)
'
End Sub

Private Sub Text1_DblClick()
   '
   Call Command1_Click
   '
End Sub

Private Sub LimpiarCajas()
   
   Label1 = ""
   Label33 = ""
  
End Sub

Private Sub List1_DblClick()
   '
   CODIX$ = TRIM(Mid(List1.TEXT, 15, 15))
   LOTEX$ = TRIM(Mid(List1.TEXT, 1, 13))
   Label32 = CODIX$ ' Codigo
   Label15 = LOTEX$ ' Lote OF
   Call Command3_Click
   '
End Sub


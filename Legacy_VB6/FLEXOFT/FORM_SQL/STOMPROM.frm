VERSION 5.00
Begin VB.Form STOMPROM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Minimo y Consumo Promedio"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "PROCESO DE CALCULO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4005
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Picture         =   "STOMPROM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cálculo de Stocks Minimo por Autonomía Proyectada"
         Top             =   480
         Width           =   640
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Picture         =   "STOMPROM.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cálculo de Consumo Promedio"
         Top             =   1200
         Width           =   640
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3960
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Cálculo de Consumos Promedios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Última Actualización de Stocks Mínimos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Fecha de Ultimo Cálculo de Autonomía Stocks Mínimo"
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         ToolTipText     =   "Fecha de Ultimo Cálculo de Consumo Promedio"
         Top             =   1320
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   4305
      ScaleHeight     =   3555
      ScaleWidth      =   930
      TabIndex        =   1
      Top             =   0
      Width           =   960
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   9
         Top             =   1560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "STOMPROM.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
End
Attribute VB_Name = "STOMPROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub




Private Sub Command2_Click()

End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    
    If DERACCE%("CALCPROM", "Proceso De Cálculo de Consumos") < 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    Call ABREMASTER
    Call ABRESTOCKS
    FE% = HOY(HO$)
    DIXA% = Left$(FECHATEX$(FE%), 2)
    '
'    ' LEE A PARTIR DE UN EXCEL EL PROMEDIO DE LOS ARTICULOS ANTES DE DEPURAR LA BASE DE DATOS
'    Dim CONSUCOD(32767) As Single
'    Set objexcel = New Excel.Application
'    RUTA$ = UCase$(LUDAT$ & "CONSUPROM.XLS")
'    Set Libro = objexcel.Workbooks.Open(RUTA$)
'    Set Hoja = Libro.Sheets(1)
'    '
'    Dim IMPORTACION As Boolean
'    IMPORTACION = False
'    URECORD& = UltimoRegistro(Hoja, 1)
'    For FILA& = 1 To URECORD&
'        On Error Resume Next
'        CODIGO$ = TRIM$(Hoja.Cells(FILA&, 1))
'        CI% = CODINT%(CODIGO$)
'        CONSUMO = XVALO(Hoja.Cells(FILA&, 2))
'        '
'        CONSUCOD(CI%) = CONSUMO
'    Next FILA&
'    Libro.Close
'    objexcel.Quit
'    ' FIN LECTURA DEL ARCHIVO EXCEL *************************
    '
    Call CARDEPOS
    CONDI1$ = ""
    For UI% = 1 To NUSU%
        ' VERIFICA SI EL DEPOSITO PERTENECE A UN ARMADOR.
        ' SI ES ASI, LO CONSIDERA AUNQUE ESTE COMO NO DISPONIBLE
        ENCONTRO% = 0
        For II& = 1 To ULTREG&("ARMADORS")
            XX$ = REGLEIDO$("ARMADORS", II&)
            A% = Asc(Mid$(XX$, 71, 1))
            If A% = UI% Then ENCONTRO% = 1
        Next II&
        If ENCONTRO% = 0 Then
            If DISTO%(UI%) < 0 Then
                CONDI1$ = CONDI1$ + " AND DepoMovi <> " & COSU%(UI%)
            End If
        End If
    Next UI%
    '
    'RECORRO TODOS LOS ARTICULOS
    For i% = 1 To NUMAS%
       URE& = NUMAS%
       U& = i%
       Call TRACE(20, U&, URE&)
       If i% > 0 And i% < 32768 Then
         Ci% = i%
         TIEMPOPROM% = 12 ' PLAZO FIJO 12 MESES
'         ' ESTO ES PARA CONSIDERAR EL PROMEDIO EN LOS MESES ANTERIOR A OCTUBRE 2012
'         CONPROANT% = 0: FE1% = 10: FE2% = XVALO(Mid$(FECHATEX$(HOY(HO$)), 4, 2))
'         DIFE% = FE2% - FE1%
'         If DIFE% < 0 Then DIFE% = (12 - FE1%) + FE2%
'         If DIFE% <= 6 Then
'            TIEMPOPROM% = DIFE%
'            CONPROANT% = 6 - DIFE%
'         End If
         '
         FECHADESDE = FETEXCOR1$(FechMesAtras%(DIXA%, TIEMPOPROM%)) ' ASIGNO LA FECHA DESDE
         CONDI$ = "CODIMOVI <> 'AJ' AND CODIMOVI <> 'TR' AND FECHMOV >= '" & FECHADESDE & "' AND COD_INTE = " & Ci% '& CONDI1$
         
         PROMEDSEIS = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
'         If CONSUCOD(CI%) > 0.05 Then
'            PROMEDSEIS = PROMEDSEIS + (CONPROANT% * CONSUCOD(CI%))
'         End If
         '
         PROMEDUNO = ROUND#((PROMEDSEIS / 12), 1)
         
         'ACTUALIZO EL MASTER CON LA CANTIDAD DEVUELTA DEL MOVISOT
         Call ACTUREGISTRO("MASTER", "consuprom", "CODINT = " & Ci%, PROMEDUNO, REGAF&, "NOMESS")
         '
      End If
    Next i%
    '
    'FECHA DE ULTIMO CONTROL DE PERIDO DE STOCKS MINIMO
    Label7 = FECHATEX$(HOY(HO$))
    Call GRACONTROL("AMASTO", "PROMCONS", Label7)
    '
    Screen.MousePointer = 1
    Call COMUNI("Proceso Finalizado")
    '
99  Command5.Enabled = True

End Sub

Private Sub Command8_Click()
    '
    Command8.Enabled = False
    '
    If DERACCE%("STPERMIN", "Proceso De Cálculo de Stocks Minimo por Periodo") < 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    Call ABREMASTER
    Call ABRESTOCKS
    FE% = HOY(HO$)
    DIXA% = Left$(FECHATEX$(FE%), 2)
    '
    'RECORRO TODOS LOS ARTICULOS
    For i% = 1 To NUMAS%
       URE& = NUMAS%
       U& = i%
       Call TRACE(20, U&, URE&)
       If i% > 0 And i% < 32768 Then
         Ci% = i%
         PDSTOMIN = XVALO(AUTONOMIASM$(Ci%)) ' ASIGNO EL PLAZO ASIGNADO PARA ESTE ARTICULO
         If PDSTOMIN > 0 Then
            'FECHADESDE = FETEXCOR1$(FechMesAtras%(DIXA%, PDSTOMIN%)) ' ASIGNO LA FECHA DESDE
            'CONDI$ = "CODIMOVI <> 'AJ' AND CODIMOVI <> 'TR' AND FECHMOV >= '" & FECHADESDE & "' AND COD_INTE = " & CI
            'STOCMIN = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
            CONDI$ = "CODINT = " & Ci%
            STOCMIN = XVALO(VALOBADA("MASTER", "ConsuProm", CONDI$, "NOMESS")) * PDSTOMIN
            '
            'ACTUALIZO EL MASTER CON LA CANTIDAD DEVUELTA DEL MOVISOT
            Call ACTUREGISTRO("MASTER", "STOMIN", "CODINT = " & Ci%, STOCMIN, REGAF&, "NOMESS")
            '
         End If
      End If
    Next i%
    '
    'FECHA DE ULTIMO CONTROL DE PERIDO DE STOCKS MINIMO
    Label6 = FECHATEX$(HOY(HO$))
    Call GRACONTROL("AMASTO", "PERSTMIN", Label6)
    '
    Screen.MousePointer = 1
    Call COMUNI("Proceso Finalizado")

   
   '
99 Command8.Enabled = True
End Sub

   Function FechMesAtras%(DIXA%, CantMes%)
   '
   'DEVUELVE CON MISMO NUMERO DEL DIA PASADO EN DIXA%
   ', PERO CON LA CANTIDAD DE MESES HACIA ATRAS PASADOS EN CantMes%
   
   DESFE$ = FECHATEX$(HOY(HO$))
   NUMDIA$ = Right$("00" + TRIM$(Str$(DIXA%)), 2)
   '
   ANIO% = XVALO(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = XVALO(Mid$(DESFE$, 4, 2))
'   HASTAMES% = 12 * ANIO% + MESO%: HASMES$ = ABREMESAN$(HASTAMES%, 1)
   '
   CAMESES% = XVALO(CantMes%)
   '
   For UK% = 1 To CAMESES%
     MESO% = MESO% - 1
     If MESO% < 1 Then
       MESO% = MESO% + 12
       ANIO% = ANIO% - 1
     End If
   Next UK%
   DESDEMES% = 12 * ANIO% + MESO%: DESMES$ = ABREMESAN$(DESDEMES%, 1)
   MESOX$ = Right$("00" + TRIM$(Str$(MESO%)), 2)
   ANIOX$ = Right$("00" + TRIM$(Str$(ANIO% Mod 100)), 2)
   DESDEFE% = FECHANUM(NUMDIA$ + "/" + MESOX$ + "/" + ANIOX$)
   '
   FechMesAtras% = DESDEFE%
End Function

Private Sub Form_Load()
    Label6 = CONTROL("AMASTO", "PERSTMIN")
    Label7 = CONTROL("AMASTO", "PROMCONS")
    Call APLICACIONSKINS(Me, Picture1)
End Sub
Function UltimoRegistro(HOJA, COLUMNA) As Double
   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For FILA& = 2 To HOJA.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(HOJA.Cells(FILA&, COLUMNA))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = FILA& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next FILA&
10 UltimoRegistro = UltFila&
End Function


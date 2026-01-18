VERSION 5.00
Begin VB.Form ANAPEN09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Pedidos Pendientes de Entrega"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   6540
      Begin VB.CommandButton Command3 
         Caption         =   "Con Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Todos los pedidos, total pendiente, solo cantidades, total por artículo y cliente"
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   8520
      Left            =   6960
      ScaleHeight     =   8460
      ScaleWidth      =   1965
      TabIndex        =   24
      Top             =   -720
      Width           =   2025
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   200
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   4935
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   195
         Picture         =   "ANAPEN09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1785
         Width           =   640
      End
      Begin VB.CommandButton command1 
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
         Height          =   640
         Left            =   195
         Picture         =   "ANAPEN09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   1020
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   640
         Left            =   195
         Picture         =   "ANAPEN09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Genera y Emite Reporte Solicitado"
         Top             =   4185
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -180
         Picture         =   "ANAPEN09.frx":075E
         Top             =   5190
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Condensación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   2940
      TabIndex        =   19
      Top             =   2520
      Width           =   3900
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total General por Artículo"
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
         Index           =   2
         Left            =   525
         TabIndex        =   22
         Top             =   1050
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total por Artículo y Cliente"
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
         Index           =   1
         Left            =   525
         TabIndex        =   21
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Detalle por Pedido"
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
         Index           =   0
         Left            =   525
         TabIndex        =   20
         Top             =   420
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Valorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   210
      TabIndex        =   16
      Top             =   2940
      Width           =   2535
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Cantidades y Precios"
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
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   635
         Width           =   2115
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Solo Cantidades"
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
         Index           =   0
         Left            =   315
         TabIndex        =   17
         Top             =   370
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   12
      Top             =   1575
      Width           =   2535
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total Pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   15
         Top             =   900
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Meses Posteriores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   14
         Top             =   635
         Width           =   2010
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Mes en Curso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   13
         Top             =   370
         Width           =   2010
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   2940
      TabIndex        =   3
      Top             =   210
      Width           =   3900
      Begin VB.TextBox Text10 
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
         Left            =   2625
         TabIndex        =   7
         Top             =   465
         Width           =   855
      End
      Begin VB.CommandButton Command9 
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
         Height          =   285
         Left            =   3465
         TabIndex        =   6
         Top             =   465
         Width           =   175
      End
      Begin VB.TextBox Text8 
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
         Left            =   2625
         TabIndex        =   5
         Text            =   " "
         Top             =   1365
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Height          =   285
         Left            =   3465
         TabIndex        =   4
         Top             =   1365
         Width           =   175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   11
         Top             =   1725
         Width           =   3450
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   10
         Top             =   825
         Width           =   3450
      End
      Begin VB.Label Label6 
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
         Left            =   210
         TabIndex        =   9
         Top             =   510
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   210
         TabIndex        =   8
         Top             =   1410
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Todos los Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   23
         Top             =   900
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Compras a Terceros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   2
         Top             =   635
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Fabricación Propia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   1
         Top             =   370
         Width           =   2115
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu ArmadoDeFamilias 
         Caption         =   "Asignación de Familas"
      End
      Begin VB.Menu tablaFamilia 
         Caption         =   "Acceso a Tabla de Familias"
      End
   End
   Begin VB.Menu mnureporte 
      Caption         =   "Reportes"
      Begin VB.Menu ListadodeItemsinFamilia 
         Caption         =   "Listado de Items sin Familia"
      End
      Begin VB.Menu confamilia 
         Caption         =   "Listado de Items Con Familia"
      End
   End
End
Attribute VB_Name = "ANAPEN09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function TotalManoObra(CI1%)
    SQLX$ = "SELECT * FROM SECOPER WHITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE codiconj = " & CI1%
    Set RS = FLEXCONN.Execute(SQLX$)
    TotalManoObra = 0
    With RS
      Do While Not .EOF
         NUMEOPERACION% = XVALO(!NUMEOPER)
         TotalManoObra = TotalManoObra + TIESTAN(CI1%, NUMEOPERACION%, 1, 0)
         .MoveNext
      Loop
    End With
 

End Function

Private Sub ArmadoDeFamilias_Click()
    Screen.MousePointer = 11
20  Campos$ = "CODIGOINTERNO/F0;Código/A20;Descripción/A48;C.Fam./A6;Familia/A32"
    Call CARGA_ENCABEZADO(SELFAMILI.GRID, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(SELFAMILI.MSF_2, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    SELFAMILI.Command3.Caption = "Asignar Familia"
    FIN& = CantRegistros&("MASTER")
    'CARGO LOS ITEMS CON LAS FAMILIAS
30  SQLX$ = " SELECT CODINT,CODIGO, DESCRIPCION, FAMILIA FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY CODIGO ASC "
    Set rs1 = READSET(SQLX$)
    J& = 0
    With rs1
      Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)

        CI1% = XVALO(!CODINT)
        CODIGOX$ = SAFETEXT$(!Codigo)
        DESCRI$ = SAFETEXT$(!Descripcion)
        COFA$ = TRIM$(SAFETEXT$(!familia))
        DESCRIFAMI$ = ""
        If COFA$ <> "" Then DESCRIFAMI$ = ECOARCH$("TAFAMIL", COFA$)
        '
        SELFAMILI.GRID.Rows = J& + 1
        SELFAMILI.GRID.TextMatrix(J&, 1) = CI1%
        SELFAMILI.GRID.TextMatrix(J&, 2) = CODIGOX$
        SELFAMILI.GRID.TextMatrix(J&, 3) = DESCRI$
        SELFAMILI.GRID.TextMatrix(J&, 4) = COFA$
        SELFAMILI.GRID.TextMatrix(J&, 5) = DESCRIFAMI$
        .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    If J& > 0 Then
       Screen.MousePointer = 1
       SELFAMILI.GRID.Visible = True
       SELFAMILI.GRID.ZOrder 0
       Screen.MousePointer = 1
50     SELFAMILI.Show 1
       If SELFAMILI.HaySeleccion > 0 Then ' SI HAY ITEMS SELECCIONADOS
            Call SELECHO("TAFAMIL")
            FAMILX$ = VALACT1$("TAFAMIL")
            If FAMILX$ = "" Then Screen.MousePointer = 1: GoTo 99 ' SI NO ELIGE FAMILIA NO GRABA SALE
            For i& = 1 To SELFAMILI.MSFACTIVO.Rows - 1
              If TRIM$(SELFAMILI.MSFACTIVO.TextMatrix(i&, 0)) = "*" Then
               CONDI$ = "CODINT = " & XVALO(SELFAMILI.MSFACTIVO.TextMatrix(i&, 1))
               Call ACTUREGISTRO("MASTER", "FAMILIA", CONDI$, FAMILX$, REGAF&, "NOMESS")
              End If
            Next i&
            Call COMUNI("Se ha Asingado la Familia a los Items Seleccionados")
            SELFAMILI.HaySeleccion = 0
            Screen.MousePointer = 11
            SELFAMILI.MSF_2.Clear
            SELFAMILI.GRID.Clear
            SELFAMILI.TXTBUSCAR = ""
            SELFAMILI.CMBFILTRO.Clear
            GoTo 20
       End If
    Else
       Call COMUNI("No se Han Encontrado Items")
       Screen.MousePointer = 1
       GoTo 20
    End If
    
99  Screen.MousePointer = 1

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GENLISPE
   Command2.Enabled = True
End Sub
'

Private Sub Command3_Click()
   Command3.Enabled = False
   DEPOSELECCIONADOS$ = ""
   SELDEPOS.Show 1
      STRINGDEPOS$ = ""
   FLI1% = FINMESACT% + 1: FLI2% = 32767
   FLI1% = -1: FLI2% = 32767
   Desde = FETEXCOR1$(FLI1%)
   Hasta = FETEXCOR1$(FLI2%)

   For U& = 1 To SELDEPOS.LISNUM.ListCount
       If SELDEPOS.LISNUM.Selected(U& - 1) = True Then
        TXTX$ = SELDEPOS.LISNUM.List(U& - 1)
        NDEP$ = TRIM$(Str$(Mid$(TXTX$, 1, 6)))
        STRINGDEPOS$ = STRINGDEPOS$ + NDEP$ + ","
      End If
   Next U&
    If TRIM$(STRINGDEPOS$) <> "" Then
       Command3.Caption = "Con Stock Dep.: " & Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1)  'para quitar la coma
    Else
       command1.Caption = "Con Stock "
    End If
    DEPOSELECCIONADOS$ = Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1)

    Campos$ = "ci/A0;" '1
    Campos$ = Campos$ + "Código/A20;" '2
    Campos$ = Campos$ + "Descripción/A30;" '3

    Campos$ = Campos$ + "Cta/F7;" '4
    Campos$ = Campos$ + "R.Social/A40;" '5

    Campos$ = Campos$ + "Pedido/F7;" '6
    Campos$ = Campos$ + "Sol.Ent./D8;" '7

    Campos$ = Campos$ + "Cant.Ped./F7;" '8
    Campos$ = Campos$ + "P.Unit/F12.2;" '9

    Campos$ = Campos$ + "Importe/F12.2;" '10
    Campos$ = Campos$ + "Cant.Entr./F7;" '11

    Campos$ = Campos$ + "Saldo/F7;" '12
    Campos$ = Campos$ + "Observaciones/A40;" '13
    
    Campos$ = Campos$ + "Atrazo/F7;" '14
    Campos$ = Campos$ + "T.M.Obra/F10" '15

    Erase CADENATEX$
    If TRIM$(DEPOSELECCIONADOS$) <> "" Then
      Call TEXTOLOTES(DEPOSELECCIONADOS$, ",")
        For i& = 1 To UBound(CADENATEX$)
           Campos$ = Campos$ + ";Dep." & CADENATEX$(i&) & "/A12"
        Next i&
    End If
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    TX$ = "Analisis de Pedidos Pendientes"
    
    
    
    

   NCL1% = Val(Text10)
   NCL2% = Val(Text8)
   If NCL1% > NCL2% Then Exit Sub
   '
   '
   GRCF% = 0
   If Option1(2).Value = True Then GRCF% = 1
   If Option1(1).Value = True Then GRCF% = 2
   
   If Option4(0).Value = True Then
     Call SELECHO("TAFAMIL")
     FAMISEL$ = VALACT1$("TAFAMIL")
   End If
   '
   FEHOY% = HOY(HO$)
   
   If InStr(Caption, "(POR CLIENTE)") > 0 Then
     If Option4(2).Value = True Then
       Call MENSERR(24, "Combinación no Válida")
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0: LL& = 0

   Call ABRECONEXION
   
   CONDI$ = " STATUS < 8 AND STATUS <> 3 "
   CONDI$ = CONDI$ + " AND nroclie >= " & NCL1%
   CONDI$ = CONDI$ + " AND nroclie <= " & NCL2%
   CONDI$ = CONDI$ + " AND FeSolEnt >= '" & Desde & "'"
   CONDI$ = CONDI$ + " AND FeSolEnt <= '" & Hasta & "'"
   CONDI$ = CONDI$ + " AND CANPED > CANTENT "
   FIN1& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
   '
   If FIN1& < 1 Then
     Call COMUNI("No se Han Encontrado Pedidos Pendientes Para El Rango Seleccionado")
     GoTo 99
   End If
   
   SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " ORDER BY NROPED, CODIINT,FeSolEnt"
    
   Dim PEDPED As ADODB.Recordset
   Set PEDPED = New ADODB.Recordset
   Set PEDEPD = FLEXCONN.Execute(FILTSQL$(SQLX$))
   On Error Resume Next
   '
   indDepo% = 0
   With PEDEPD
        Do While Not .EOF
               i& = i& + 1
               Call TRACE(20, i&, FIN1&)
               CI7% = XVALO(!CODIINT)
               If GRCF% = 0 Or GRUCOFAB_V2%(CI7%) = GRCF% Then
                    '
                    If FAMISEL$ <> "" Then ' SI TIENE FILTRO DE FAMILIA INGRESA
                      FAMI$ = CODFAMIL$(CI7%) 'LO HICE ASI PARA QUE SI NO SE FILTRO NO ENTRE EN LA FUNCION CODFAMIL$ CADA VEZ QUE PASA
                      If FAMI$ <> FAMISEL$ Then GoTo 25
                    End If
                    '
                    NCLIE = XVALO(!NroClie)
                    CODI$ = CODEXT$(CI7%)

                    FENTRESOL% = CVINT(!FeSolEnt)
                    CAPEN = XVALO(!Canped) - XVALO(!Cantent): If CAPEN < 0.05 Then CAPEN = 0
                    PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(NCLIE)) / 100
                    If OPX% <> 2 Then PREUNI = 0
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    DiasDeAtrazo% = DIENTRE(FENTRESOL%, FEHOY%)
                    J& = J& + 1
                    GRILLAGRAL.GRID.Rows = J& + 1
                    GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI7%
                    GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Codiext)
                    GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
                    
                    GRILLAGRAL.GRID.TextMatrix(J&, 4) = NCLIE
                    GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!RASOCLIE)
                    GRILLAGRAL.GRID.TextMatrix(J&, 6) = XVALO(!NROPED)
                    
                    GRILLAGRAL.GRID.TextMatrix(J&, 7) = FECHATEX$(FENTRESOL%)
                    GRILLAGRAL.GRID.TextMatrix(J&, 8) = XVALO(!Canped)
                    GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(XVALO(!PREUNI), "F10.2")
                    
                    GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(ITOTA#, "F10.2")
                    GRILLAGRAL.GRID.TextMatrix(J&, 11) = XVALO(!Cantent)
                      saldo = XVALO(!Canped) - XVALO(!Cantent): If saldo < 0 Then saldo = 0
                    GRILLAGRAL.GRID.TextMatrix(J&, 12) = saldo
                    
                    GRILLAGRAL.GRID.TextMatrix(J&, 13) = Space$(18)
                    GRILLAGRAL.GRID.TextMatrix(J&, 14) = DiasDeAtrazo%
                        TotalMobra = TotalManoObra(CI7%) 'DEVUELVE EL TIEMPO POR UNA PIEZA
                        TotalMobra = TotalMobra * 60 * CAPEN ' RECALCULO EL TIEMPO POR LA CANTIDAD PENDIENTE (SALDO)
                        TSTT$ = AJUSTD$(HORMINSE_LAR$(TotalMobra), 16)
                    GRILLAGRAL.GRID.TextMatrix(J&, 15) = TSTT$
                    
                    For COL& = 16 To GRILLAGRAL.GRID.Cols - 1
                     indDepo% = indDepo% + 1
                     CantEnDeposito = STODEPOS(CI7%, Int(CADENATEX$(indDepo%))) ' Y quiere que lo presente
                    GRILLAGRAL.GRID.TextMatrix(J&, COL&) = FORMATNUM$(CantEnDeposito, "F12.1")
                   Next COL&
                   indDepo% = 0
               End If
25       .MoveNext
       Loop
   End With
   
   '*******************************************
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
    
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1

    
   GRILLAGRAL.Show 1
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOSCRP", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
   End If
99 Command3.Enabled = True
End Sub


Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
       Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub confamilia_Click()
  Call FINAL("?ITEMCONFAM")
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " (" + APLINVO$ + ")  -  " + EPAC$
    End If
    '
    Show
    Refresh
    If ULTREG&("DEUDOR1") > 0 Then
      Screen.MousePointer = 11
      NCMIN% = 32767: NCMAX% = 0
      FIN& = ULTREG&("DEUDOR1")
      For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
          If NCX% > NCMAX% Then NCMAX% = NCX%
          If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
      Next U&
      '
      Screen.MousePointer = 1
      Text10 = TRIM$(Str$(NCMIN%))
      Text8 = TRIM$(Str$(NCMAX%))
    End If
    '
    Call ABRECONEXION
    Call PoneEnCeroNull("MASTER", "FAMILIA", REGAF&, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub ListadodeItemsinFamilia_Click()
   Call FINAL("?ITEMSINFAM")

End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub





Private Sub tablaFamilia_Click()
   Call TRALOCAL("TAFAMIL", "")

10  VTB% = VENTABU("TAFAMIL")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAFAMIL")
        X$ = REGLEIDO$("!TAFAMIL", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        Call REPLA(X$, CD$, 1, 2)
        Call GRAREG("!TAFAMIL", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(CD$) < 2 Then
          Call MENSERR(24, "Código de Familia Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        DESC$ = TRIM$(Mid$(X$, 3, 27))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Codigo de Familia sin Descripción")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TAFAMIL", U2&)
          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
            Call MENSERR(24, "Código de Familia Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TAFAMIL", "")
    End If
End Sub



Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = rasocli$(NCMIN%)
   End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Text8_Change()
   Label8 = ""
   NCMAX% = Val(Text8.TEXT)
   If Val(NCMAX%) > 0 Then
     Label8 = rasocli$(NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub
Sub GENLISPE()
   '
   If Option4(0).Value = True Then
     Call SELECHO("TAFAMIL")
     FAMISEL$ = VALACT1$("TAFAMIL")
   End If
   '
   FEHOY% = HOY(HO$)
   
   If InStr(Caption, "(POR CLIENTE)") > 0 Then
     If Option4(2).Value = True Then
       Call MENSERR(24, "Combinación no Válida")
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0: LL& = 0
'   FIN& = ULTREG&("PEDDETA")
   '
   MODO% = 3
   If Option2(0).Value = True Then MODO% = 1
   If Option2(1).Value = True Then MODO% = 2
   FLI1% = -1: FLI2% = FINMESACT%
   If MODO% = 2 Then
       FLI1% = FINMESACT% + 1: FLI2% = 32767
     ElseIf MODO% = 3 Or MODO% = 4 Then
       FLI1% = -1: FLI2% = 32767
   End If
   '
   Desde = FETEXCOR1$(FLI1%)
   Hasta = FETEXCOR1$(FLI2%)
   '
   OPX% = 2
   If Option3(0) = True Then OPX% = 1
   '
   Call HEADERS("LIPEDPE", "")
   'Call HEADERS("ANALIPED", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   Call HEADERS("LIPEDPE", VARIANTE$)
   'Call HEADERS("ANALIPED", VARIANTE$)
   '
   NCL1% = Val(Text10)
   NCL2% = Val(Text8)
   If NCL1% > NCL2% Then Exit Sub
   '
   GRCF% = 0
   If Option1(2).Value = True Then GRCF% = 1
   If Option1(1).Value = True Then GRCF% = 2
   '
   VECPEDI$ = ""
   '
   Call ABRECONEXION
   
   CONDI$ = " STATUS < 8 AND STATUS <> 3 "
   CONDI$ = CONDI$ + " AND nroclie >= " & NCL1%
   CONDI$ = CONDI$ + " AND nroclie <= " & NCL2%
   CONDI$ = CONDI$ + " AND FeSolEnt >= '" & Desde & "'"
   CONDI$ = CONDI$ + " AND FeSolEnt <= '" & Hasta & "'"
   CONDI$ = CONDI$ + " AND CANPED > CANTENT "
   FIN1& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
   '
   If FIN1& < 1 Then
     Call COMUNI("No se Han Encontrado Pedidos Pendientes Para El Rango Seleccionado")
     GoTo 99
   End If
   '
   FORMATOHORAORIGINAL$ = TRIM$(Control$("PADIAMOD", "CONFHORA"))
   Call GRACONTROL("PADIAMOD", "CONFHORA", "MINUCENT")
   
   SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
    
   SQLX$ = SQLX$ + " "
   Dim PEDPED As ADODB.Recordset
   Set PEDPED = New ADODB.Recordset
   'PEDEPD.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
   Set PEDEPD = FLEXCONN.Execute(FILTSQL$(SQLX$))
   On Error Resume Next
   '
   With PEDEPD
        Do While Not .EOF
               i& = i& + 1
               Call TRACE(20, i&, FIN1&)
               CI7% = XVALO(!CODIINT)
               If GRCF% = 0 Or GRUCOFAB_V2%(CI7%) = GRCF% Then
                    '
                    If FAMISEL$ <> "" Then ' SI TIENE FILTRO DE FAMILIA INGRESA
                      FAMI$ = CODFAMIL$(CI7%) 'LO HICE ASI PARA QUE SI NO SE FILTRO NO ENTRE EN LA FUNCION CODFAMIL$ CADA VEZ QUE PASA
                      If FAMI$ <> FAMISEL$ Then GoTo 25
                    End If
                    '
                    NCLIE = XVALO(!NroClie)
                    CODI$ = CODEXT$(CI7%)

                    FENTRESOL% = CVINT(!FeSolEnt)
                    CAPEN = XVALO(!Canped) - XVALO(!Cantent): If CAPEN < 0.05 Then CAPEN = 0
                    PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(NCLIE)) / 100
                    If OPX% <> 2 Then PREUNI = 0
                    ITOTA# = CDbl(CAPEN * PREUNI)
                    INOAS# = CDbl(CANOASI * PREUNI)
                    Call REPLA(Y$, MKS$(NCLIE), 1, 4)
                    Call REPLA(Y$, MKI$(CI7%), 5, 2)
                    Call REPLA(Y$, !TCOLOR, 7, 16)
                    Call REPLA(Y$, !TTALLE, 23, 8)
                    Call REPLA(Y$, MKS$(!TLargo), 31, 4)
                    Call REPLA(Y$, MKS$(!TAncho), 35, 4)
                    Call REPLA(Y$, MKS$(!TEspesor), 39, 4)
                    Call REPLA(Y$, MKS$(NCLIE), 47, 4)
                    Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                    Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                    Call REPLA(Y$, MKS$(!Canped), 83, 4)
                    Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                    Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                    Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                    Call REPLA(Y$, MKS$(!Cantent), 99, 4)
                    Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                    Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                    Call REPLA(Y$, XVALO(!NROOCOM), 129, 16)
                    Call REPLA(Y$, Space$(18), 145, 18)
                    DiasDeAtrazo% = DIENTRE(FENTRESOL%, FEHOY%)
                    Call REPLA(Y$, MKS$(DiasDeAtrazo%), 163, 4)
                    '
                    TotalMobra = TotalManoObra(CI7%) 'DEVUELVE EL TIEMPO POR UNA PIEZA
                    TotalMobra = TotalMobra * 60 * CAPEN ' RECALCULO EL TIEMPO POR LA CANTIDAD PENDIENTE (SALDO)
                    'TSTT$ = AJUSTD$(HORMINSEDIEM$(TotalMobra), 16)
                    TSTT$ = AJUSTD$(HORMINSE_LAR$(TotalMobra), 16)
                    Call REPLA(Y$, TSTT$, 167, 16)
                    JJ& = JJ& + 1
                    Call GRAREG("LIPEDPE", Y$, JJ&)
                    'Call GRAREG("ANALIPED", Y$, JJ&)
                    VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
               End If
25       .MoveNext
       Loop
   End With
   
   Call GRACONTROL("PADIAMOD", "CONFHORA", FORMATOHORAORIGINAL$)
   
30 If MODO% = 4 Then
     KK& = JJ&: JJ& = 0
     For U& = 1 To KK&
       Y$ = REGLEIDO$("LIPEDPE", U&)
       If CVS(Mid$(Y$, 103, 4)) >= 0.05 Then
         JJ& = JJ& + 1
         Call GRAREG("LIPEDPE", Y$, JJ&)
       End If
     Next U&
   End If
   '
   Call SETULTREG("LIPEDPE", JJ&)
   Call CIERRARCH("LIPEDPE")
   '
   Call HEADERS("LIPEDPE", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   '
   If Option4(0).Value = True Then
     If InStr(Caption, "(POR CLIENTE)") > 0 Then
         Call FILESORT("LIPEDPE", "", 1, 3)
         sali$ = "*LIPENCL"
       Else
         Call FILESORT("LIPEDPE", "", 3, 3)
         sali$ = "*LIPENMA"
     End If
     Call CIERRARCH("LIPEDPE")
     Call FINAL(sali$)
     GoTo 99
   End If
   '
   If Option4(1).Value = True Or Option4(2).Value = True Then ' concentrar por articulo y cliente
     FIN& = ULTREG&("LIPEDPE")
     PUCIX$ = ""
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI = CVS(Mid$(XX$, 1, 4))
       CII% = CVI(Mid$(XX$, 5, 2))
       '
       If CII% <= 0 Then
            PUCIX$ = PUCIX$ + MKI$(0)
         Else
            CAPENPE = CVS(Mid$(XX$, 83, 4)) - CVS(Mid$(XX$, 99, 4))
            If CAPENPE < 0 Then CAPENPE = 0
            For KKU& = 1 To Len(PUCIX$) Step 2
              If CVI(Mid$(PUCIX$, KKU&, 2)) = CII% Then
                YY$ = REGLEIDO$("LIPEDPE", (KKU& + 1) / 2)
                If CVI(Mid$(YY$, 1, 2)) = NCI Or Option4(2).Value = True Then
                  TOCANTI = CVS(Mid$(YY$, 83, 4)) + CVS(Mid$(XX$, 83, 4))
                  TOENTRE = CVS(Mid$(YY$, 99, 4)) + CVS(Mid$(XX$, 99, 4))
                  TOPENPE = CVS(Mid$(YY$, 103, 4)) + CAPENPE
                  ITOTAL# = CVD(Mid$(YY$, 91, 8)) + CVD(Mid$(XX$, 91, 8))
                  Call REPLA(YY$, MKS$(TOCANTI), 83, 4)
                  Call REPLA(YY$, MKD$(ITOTAL#), 91, 8)
                  Call REPLA(YY$, MKS$(TOENTRE), 99, 4)
                  Call REPLA(YY$, MKS$(TOPENPE), 103, 4)
                  Call GRAREG("LIPEDPE", YY$, (KKU& + 1) / 2)
                  PUCIX$ = PUCIX$ + MKI$(0)
                  Call GRAREG("LIPEDPE", String$(128, 0), U&)
                  GoTo 39
                End If
              End If
            Next KKU&
            '
            PUCIX$ = PUCIX$ + MKI$(CII%)
            Call GRAREG("LIPEDPE", XX$, U&)
       End If
39   Next U&
     '
     JJ& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI = CVS(Mid$(XX$, 1, 4))
       CII% = CVI(Mid$(XX$, 5, 2))
       '
       If NCI > 0 Then
         If CII% > 0 Then
           JJ& = JJ& + 1
           Call GRAREG("LIPEDPE", XX$, JJ&)
         End If
       End If
     Next U&
     '
     Call SETULTREG("LIPEDPE", JJ&)
     '
     If Option4(1).Value = True Then
       If InStr(Caption, "(POR CLIENTE)") > 0 Then
           Call FILESORT("LIPEDPE", "", 1, 3)
           sali$ = "*LIPENCL"
         Else
           Call FILESORT("LIPEDPE", "", 3, 3)
           sali$ = "*LIPENMA"
       End If
       Call FINAL(sali$)
       GoTo 99
     End If
     '
     If Option4(2).Value = True Then
       FIN& = ULTREG&("LIPEDPE")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("LIPEDPE", U&)
         Call REPLA(XX$, String$(53, 0), 107, 53)
         CIXI% = CVI(Mid$(XX$, 5, 2))
         NECEPE = CVS(Mid$(XX$, 103, 4))
         If NECEPE >= 0.005 Then
           DISPONI = STOCKACT(CIXI%)
           If DISPONI < 0 Then DISPONI = 0
           Call REPLA(XX$, MKS$(DISPONI), 107, 4)
           NECEPE = NECEPE - DISPONI
         End If
         DISPONI = STOCOPEN(CIXI%) + STORFAB(CIXI%)
         Call REPLA(XX$, MKS$(DISPONI), 111, 4)
         NECEPE = NECEPE - DISPONI
         If NECEPE >= 0.05 Then
             Call REPLA(XX$, MKS$(NECEPE), 119, 4)
             DISPONI = MAXIFAB(CIXI%)
             Call REPLA(XX$, MKS$(DISPONI), 123, 4)
           Else
             Call REPLA(XX$, MKS$(Abs(NECEPE)), 115, 4)
         End If
         '
         Call GRAREG("LIPEPEN", XX$, U&)
         '
       Next U&
       '
       Call SETULTREG("LIPEPEN", FIN&)
       Call FILESORT("LIPEPEN", "", 3, 3)
       '
       Call HEADERS("LIPEPEN", "")
       If Option1(0).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: Total Pendiente")
       If Option1(1).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Fabricar")
       If Option1(2).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Comprar")
       '
       Call FINAL("*LIPEPEN")
       GoTo 99
     End If
   End If

   '
99 Screen.MousePointer = 1
End Sub

Sub GENLISPE_VIEJO_090615()
   
   If InStr(Caption, "(POR CLIENTE)") > 0 Then
     If Option4(2).Value = True Then
       Call MENSERR(24, "Combinación no Válida")
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0: LL& = 0
'   FIN& = ULTREG&("PEDDETA")
   '
   MODO% = 3
   If Option2(0).Value = True Then MODO% = 1
   If Option2(1).Value = True Then MODO% = 2
   FLI1% = -1: FLI2% = FINMESACT%
   If MODO% = 2 Then
       FLI1% = FINMESACT% + 1: FLI2% = 32767
     ElseIf MODO% = 3 Or MODO% = 4 Then
       FLI1% = -1: FLI2% = 32767
   End If
   '
   OPX% = 2
   If Option3(0) = True Then OPX% = 1
   '
   Call HEADERS("LIPEDPE", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   Call HEADERS("LIPEDPE", VARIANTE$)
   '
   NCL1% = Val(Text10)
   NCL2% = Val(Text8)
   If NCL1% > NCL2% Then Exit Sub
   '
   GRCF% = 0
   If Option1(2).Value = True Then GRCF% = 1
   If Option1(1).Value = True Then GRCF% = 2
   '
   VECPEDI$ = ""
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     X$ = REGLEIDO$("PEDDETA", U&)
'     STATUPE% = Asc(Mid$(X$, 27, 1))
'     If STATUPE% <> 9 Then ' no esta anulado
'       If STATUPE% <> 3 Then ' no esta entregado
'         CI7% = CVI(Mid$(X$, 33, 2))
'         If CI7 > 0 And CI7% <= NUMAS% Then
'           If GRCF% = 0 Or GRUCOFAB%(CI7%) = GRCF% Then
'             NC% = CVI(Mid$(X$, 7, 2))
'             If NC% >= NCL1% Then
'               If NC% <= NCL2% Then
'                 CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
'                 If CAPEN > 0.05 Then
'                   X1$ = X$
'                   FENTRESOL% = CVI(Mid$(X$, 25, 2))
'                   If FENTRESOL% >= FLI1% Then
'                     If FENTRESOL% <= FLI2% Then
'                       '
'                       Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
'                       Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
'                       Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
'                       Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
'                       '
'                       CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
'                       If CANOASI < 0 Then CANOASI = 0
'                       '
'                       PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
'                       If OPX% <> 2 Then PREUNI = 0
'                       ITOTA# = CDbl(CAPEN * PREUNI)
'                       INOAS# = CDbl(CANOASI * PREUNI)
'                       '
'                       Call REPLA(Y$, MKS$(PREUNI), 85, 4)
'                       Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
'                       Call REPLA(Y$, MKS$(CANOASI), 111, 4)
'                       Call REPLA(Y$, MKD$(INOAS#), 119, 8)
'                       Call REPLA(Y$, Space$(34), 127, 34)
'                       '
'                       JJ& = JJ& + 1
'                       Call GRAREG("LIPEDPE", Y$, JJ&)
'                       VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
'                       '
'                     End If
'                   End If
'                 End If
'               End If
'            End If
'           End If
'         End If
'       End If
'     End If
'   Next U&
   '
   Call ABRECONEXION
   
   SQLX$ = "SELECT * FROM PEDDETA "
   Dim PEDPED As ADODB.Recordset
   Set PEDPED = New ADODB.Recordset
   'PEDEPD.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
   Set PEDEPD = FLEXCONN.Execute(FILTSQL$(SQLX$))
   On Error Resume Next
   '
   With PEDEPD
        Do While Not .EOF
            STATUPE% = XVALO(!Status)
            If STATUPE% < 8 Then ' no esta anulado
               If STATUPE% <> 3 Then ' no esta entregado
               CI7% = XVALO(!CODIINT)
               If CI7% > 0 And CI7% <= NUMAS% Then
                  If GRCF% = 0 Or GRUCOFAB%(CI7%) = GRCF% Then
                    NC% = XVALO(!NroClie)
                    If NC% >= NCL1% Then
                       If NC% <= NCL2% Then
'                      CAPEN = XVALO(!CanPed) ' Cantped > 0 (Articulos ya entregados por completo forman parte del listado)
                       '
                       CAPEN = XVALO(!Canped) - XVALO(!Cantent) ' (Muestra solamente aquellos con saldo pendiente de entrega)
                       '
                       If CAPEN > 0.05 Then
                       FENTRESOL% = CVINT(!FeSolEnt)
                       If FENTRESOL% >= FLI1% Then
                          If FENTRESOL% <= FLI2% Then
'                              Call REPLA(X$, MKS$(!NROPED), 1, 4)
'                              Call REPLA(X$, MKI$(!FECHEMIS), 5, 2)
'                              Call REPLA(X$, MKI$(NC%), 7, 2)
'
'                              Call REPLA(X$, SAFETEXT$(!NROOCOM), 9, 16)
'                              Call REPLA(X$, MKI$(CI7%), 33, 2)
'                              Call REPLA(X$, SAFETEXT$(!TColor), 35, 16)
'                              Call REPLA(X$, SAFETEXT$(!TTalle), 51, 8)
'                              Call REPLA(X$, SAFETEXT$(!TLargo), 59, 4)
'                              Call REPLA(X$, SAFETEXT$(!TAncho), 63, 4)
'                              Call REPLA(X$, SAFETEXT$(!TEspesor), 67, 4)
'                              Call REPLA(X$, MKS$(XVALO(!PRELIST)), 71, 4)
'                              Call REPLA(X$, MKS$(XVALO(!CANASIG)), 75, 4)
'                              Call REPLA(X$, MKS$(XVALO(!CanPed)), 79, 4)
'                              Call REPLA(X$, MKS$(XVALO(!PreUnid)), 83, 4)
'                              Call REPLA(X$, MKD$(XVALO(!IMPOTOT)), 87, 8)
'                              Call REPLA(X$, MKS$(XVALO(!CantEnt)), 95, 4)
'                              Call REPLA(X$, MKS$(XVALO(!CanSaldo)), 99, 4)
'                              Call REPLA(X$, MKI$(CVINT(!FeUltEnt)), 103, 2)
'                              Call REPLA(X$, MKS$(XVALO(!NroRemi)), 105, 4)
'
'                              Y$ = String$(51, 0) + Left$(X$, 6) + Mid$(X$, 9, 16) + String$(10, 0) + Mid$(X$, 79)
'                              Call REPLA(Y$, MKS$(NC%), 1, 4)
'                              Call REPLA(Y$, MKS$(NC%), 47, 4)
'                              Call REPLA(Y$, Mid$(X$, 33, 42), 5, 42)
'                              '
                                                      
                        PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(NC%)) / 100
                        If OPX% <> 2 Then PREUNI = 0
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        INOAS# = CDbl(CANOASI * PREUNI)
                              
                        Call REPLA(Y$, MKS$(NC%), 1, 4)
                        Call REPLA(Y$, MKI$(CI7%), 5, 2)
                        Call REPLA(Y$, !TCOLOR, 7, 16)
                        Call REPLA(Y$, !TTALLE, 23, 8)
                        Call REPLA(Y$, MKS$(!TLargo), 31, 4)
                        Call REPLA(Y$, MKS$(!TAncho), 35, 4)
                        Call REPLA(Y$, MKS$(!TEspesor), 39, 4)
                        Call REPLA(Y$, MKS$(NC%), 47, 4)
                        Call REPLA(Y$, MKS$(!NROPED), 51, 4)
                        Call REPLA(Y$, MKI$(FENTRESOL%), 73, 2)
                        Call REPLA(Y$, MKS$(!Canped), 83, 4)
                        Call REPLA(Y$, MKS$(CAPEN), 103, 4)
                        Call REPLA(Y$, MKS$(PREUNI), 87, 4)
                        Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
                        Call REPLA(Y$, MKS$(!Cantent), 99, 4)
                        Call REPLA(Y$, MKS$(CANOASI), 113, 4)
                        Call REPLA(Y$, MKD$(INOAS#), 121, 8)
                        Call REPLA(Y$, !NROOCOM, 129, 16)
                        Call REPLA(Y$, Space$(18), 145, 18)
                        
'                              CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
'                              If CANOASI < 0 Then CANOASI = 0
'                              '
'                              PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
'                              If OPX% <> 2 Then PREUNI = 0
'                              ITOTA# = CDbl(CAPEN * PREUNI)
'                              INOAS# = CDbl(CANOASI * PREUNI)
'                              '
'                              Call REPLA(Y$, MKS$(PREUNI), 87, 4)
'                              Call REPLA(Y$, MKD$(ITOTA#), 91, 8)
'                              Call REPLA(Y$, MKS$(CANOASI), 113, 4)
'                              Call REPLA(Y$, MKD$(INOAS#), 121, 8)
'                              Call REPLA(Y$, Space$(34), 129, 34)
                              '
                              JJ& = JJ& + 1
                              Call GRAREG("LIPEDPE", Y$, JJ&)
                              VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                        End If
                    End If
                  End If
                End If
              End If
            End If
          End If
         End If
         End If
         .MoveNext
       Loop
   End With
   
30 If MODO% = 4 Then
     KK& = JJ&: JJ& = 0
     For U& = 1 To KK&
       Y$ = REGLEIDO$("LIPEDPE", U&)
       If CVS(Mid$(Y$, 103, 4)) >= 0.05 Then
         JJ& = JJ& + 1
         Call GRAREG("LIPEDPE", Y$, JJ&)
       End If
     Next U&
   End If
   '
   Call SETULTREG("LIPEDPE", JJ&)
   Call CIERRARCH("LIPEDPE")
   '
   Call HEADERS("LIPEDPE", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   '
   If Option4(0).Value = True Then
     If InStr(Caption, "(POR CLIENTE)") > 0 Then
         Call FILESORT("LIPEDPE", "", 1, 3)
         sali$ = "*LIPENCL"
       Else
         Call FILESORT("LIPEDPE", "", 3, 3)
         sali$ = "*LIPENMA"
     End If
     Call CIERRARCH("LIPEDPE")
     Call FINAL(sali$)
     GoTo 99
   End If
   '
   If Option4(1).Value = True Or Option4(2).Value = True Then ' concentrar por articulo y cliente
     FIN& = ULTREG&("LIPEDPE")
     PUCIX$ = ""
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI = CVS(Mid$(XX$, 1, 4))
       CII% = CVI(Mid$(XX$, 5, 2))
       '
       If CII% <= 0 Then
            PUCIX$ = PUCIX$ + MKI$(0)
         Else
            CAPENPE = CVS(Mid$(XX$, 83, 4)) - CVS(Mid$(XX$, 99, 4))
            If CAPENPE < 0 Then CAPENPE = 0
            For KKU& = 1 To Len(PUCIX$) Step 2
              If CVI(Mid$(PUCIX$, KKU&, 2)) = CII% Then
                YY$ = REGLEIDO$("LIPEDPE", (KKU& + 1) / 2)
                If CVI(Mid$(YY$, 1, 2)) = NCI Or Option4(2).Value = True Then
                  TOCANTI = CVS(Mid$(YY$, 83, 4)) + CVS(Mid$(XX$, 83, 4))
                  TOENTRE = CVS(Mid$(YY$, 99, 4)) + CVS(Mid$(XX$, 99, 4))
                  TOPENPE = CVS(Mid$(YY$, 103, 4)) + CAPENPE
                  ITOTAL# = CVD(Mid$(YY$, 91, 8)) + CVD(Mid$(XX$, 91, 8))
                  Call REPLA(YY$, MKS$(TOCANTI), 83, 4)
                  Call REPLA(YY$, MKD$(ITOTAL#), 91, 8)
                  Call REPLA(YY$, MKS$(TOENTRE), 99, 4)
                  Call REPLA(YY$, MKS$(TOPENPE), 103, 4)
                  Call GRAREG("LIPEDPE", YY$, (KKU& + 1) / 2)
                  PUCIX$ = PUCIX$ + MKI$(0)
                  Call GRAREG("LIPEDPE", String$(128, 0), U&)
                  GoTo 39
                End If
              End If
            Next KKU&
            '
            PUCIX$ = PUCIX$ + MKI$(CII%)
            Call GRAREG("LIPEDPE", XX$, U&)
       End If
39   Next U&
     '
     JJ& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI = CVS(Mid$(XX$, 1, 4))
       CII% = CVI(Mid$(XX$, 5, 2))
       '
       If NCI > 0 Then
         If CII% > 0 Then
           JJ& = JJ& + 1
           Call GRAREG("LIPEDPE", XX$, JJ&)
         End If
       End If
     Next U&
     '
     Call SETULTREG("LIPEDPE", JJ&)
     '
     If Option4(1).Value = True Then
       If InStr(Caption, "(POR CLIENTE)") > 0 Then
           Call FILESORT("LIPEDPE", "", 1, 3)
           sali$ = "*LIPENCL"
         Else
           Call FILESORT("LIPEDPE", "", 3, 3)
           sali$ = "*LIPENMA"
       End If
       Call FINAL(sali$)
       GoTo 99
     End If
     '
     If Option4(2).Value = True Then
       FIN& = ULTREG&("LIPEDPE")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("LIPEDPE", U&)
         Call REPLA(XX$, String$(53, 0), 107, 53)
         CIXI% = CVI(Mid$(XX$, 5, 2))
         NECEPE = CVS(Mid$(XX$, 103, 4))
         If NECEPE >= 0.005 Then
           DISPONI = STOCKACT(CIXI%)
           If DISPONI < 0 Then DISPONI = 0
           Call REPLA(XX$, MKS$(DISPONI), 107, 4)
           NECEPE = NECEPE - DISPONI
         End If
         DISPONI = STOCOPEN(CIXI%) + STORFAB(CIXI%)
         Call REPLA(XX$, MKS$(DISPONI), 111, 4)
         NECEPE = NECEPE - DISPONI
         If NECEPE >= 0.05 Then
             Call REPLA(XX$, MKS$(NECEPE), 119, 4)
             DISPONI = MAXIFAB(CIXI%)
             Call REPLA(XX$, MKS$(DISPONI), 123, 4)
           Else
             Call REPLA(XX$, MKS$(Abs(NECEPE)), 115, 4)
         End If
         '
         Call GRAREG("LIPEPEN", XX$, U&)
         '
       Next U&
       '
       Call SETULTREG("LIPEPEN", FIN&)
       Call FILESORT("LIPEPEN", "", 3, 3)
       '
       Call HEADERS("LIPEPEN", "")
       If Option1(0).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: Total Pendiente")
       If Option1(1).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Fabricar")
       If Option1(2).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Comprar")
       '
       Call FINAL("*LIPEPEN")
       GoTo 99
     End If
   End If

   '
99 Screen.MousePointer = 1

End Sub

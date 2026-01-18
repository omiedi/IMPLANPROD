VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORLICLI07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listados de Clientes - Cuentas Corrientes"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9225
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
   ScaleHeight     =   3150
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "AVANCE DE PROCESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   22
      Top             =   2160
      Width           =   4005
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3240
         Picture         =   "FORLICLI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   180
         Width           =   640
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   24
         Top             =   375
         Width           =   2850
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   420
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RANGO DE CUENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   4200
      TabIndex        =   12
      Top             =   105
      Width           =   4005
      Begin VB.TextBox Text10 
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
         Left            =   2625
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3465
         TabIndex        =   17
         Top             =   300
         Width           =   175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   " "
         Top             =   675
         Width           =   3330
      End
      Begin VB.TextBox Text8 
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
         Left            =   2625
         TabIndex        =   15
         Text            =   " "
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3465
         TabIndex        =   14
         Top             =   1140
         Width           =   175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1515
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   210
         TabIndex        =   20
         Top             =   440
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   210
         TabIndex        =   19
         Top             =   1280
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   4005
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1440
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3465
         TabIndex        =   9
         Top             =   1020
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2940
         TabIndex        =   8
         Text            =   " "
         Top             =   1020
         Width           =   540
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   675
         Width           =   3330
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3465
         TabIndex        =   5
         Top             =   300
         Width           =   175
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
         Height          =   300
         Left            =   2940
         TabIndex        =   4
         Text            =   " "
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   525
         TabIndex        =   11
         Top             =   1155
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente"
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
         Left            =   525
         TabIndex        =   7
         Top             =   435
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ORDENAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4005
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Código Postal"
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
         Height          =   360
         Index           =   3
         Left            =   2835
         TabIndex        =   23
         Top             =   380
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Alfabético"
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
         Height          =   240
         Index           =   2
         Left            =   1470
         TabIndex        =   2
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Numérico"
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
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   8340
      ScaleHeight     =   4605
      ScaleWidth      =   1035
      TabIndex        =   25
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   30
         Top             =   2720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "FORLICLI07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FORLICLI07.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORLICLI07.frx":0688
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FORLICLI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GELISCLI_NUEVO()

    Screen.MousePointer = 11
    '
    NCMIN = XVALO(Text10.TEXT)
    NCMAX = XVALO(Text8.TEXT)
    '
    GRUCO% = XVALO(Text1)
    VENDE% = XVALO(Text3)
    '
    If NCMIN < 1 Then NCMIN = PrimerClienteActivo&
    If NCMAX < 1 Then NCMAX = UltimoClienteActivo&
    '
10  CAMPOS$ = "Cuenta/I10;Cliente/A36;Domicilio/A36;C-Postal/A10;Localidad/A20;Prov/A4;Provincia/A20;Telefono/A24;Fax/A24;Mail/A48;Cod/I3;Pos.Iva/A20;Cuit/A13;L-P/I3;Lista Precios/A15;Dto/A16;C-P/I3;Cond. Pago/A20;Cta Madre/A16"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
    GRILLAGRAL.Caption = "LISTADO DE CLIENTES (Rango: " & CStr(NCMIN) & " - " & CStr(NCMAX) & " | Grupo: " & CStr(GRUCO%) & " | Vendedor: " & CStr(VENDE%) & ")"
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu3.Caption = "Invertir selección"
    
    MODORDE% = 1
    If FORLICLI07.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLICLI07.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN & " "
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX & " "
    If TRIM$(Text1) <> "" Then SQLX$ = SQLX$ + "AND GrCo_Cli =  " & GRUCO% & "  "
    If TRIM$(Text3) <> "" Then SQLX$ = SQLX$ + "AND VEND_Cli = " & VENDE% & " "
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
      Case 1
      SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
      Case 2
      SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
      Case 3
      SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set ClieTemp = READSET(SQLX$)
    With ClieTemp
      On Error Resume Next
      .MoveFirst
      'Valida que la tabla no este vacia
      If Err Then
        Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
        Screen.MousePointer = 1
        On Error GoTo 0
        GoTo 90
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      Do While (.EOF = False)
        '
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
            NCLIE = !nume_cli
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = NCLIE
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = !raso_cli
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = SAFETEXT$(!domi_cli)
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(!Cpos_Cli)
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = SAFETEXT$(!loca_Cli)
            PROV$ = SAFETEXT$(!Prov_Cli)
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = PROV$
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = ECOARCH$("PROVINC", PROV$)
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = SAFETEXT$(!Tele_Cli)
        GRILLAGRAL.GRID.TextMatrix(J&, 9) = SAFETEXT$(!Faxi_Cli)
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = SAFETEXT$(!Mail_Cli)
        
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = PIVACLI%(NCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 12) = POSIVAC$(NCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 13) = SAFETEXT$(!Cuit_Cli)
            LPRX% = LIPRCLI%(NCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 14) = LPRX%
        GRILLAGRAL.GRID.TextMatrix(J&, 15) = DELIPRE$(LPRX%)
            DCCL$ = TRIM$(SAFETEXT$(!Dsta_Cli))
            If DCCL$ <> "" Then
               If InStr(DCCL$, "%") < 1 Then
                  DCCL$ = DCCL$ + " %"
               End If
               If Len(DCCL$) > 28 Then DCCL$ = Left$(DCCL$, 28)
               DDLX$ = Left$(DDLX$, 29 - Len(DCCL$))
               DDLX$ = DDLX$ + " (" + DCCL$ + ")"
            End If
        GRILLAGRAL.GRID.TextMatrix(J&, 16) = DCCL$
            CPAG% = CPAGCLI%(NCLIE)
        GRILLAGRAL.GRID.TextMatrix(J&, 17) = CPAG%
        GRILLAGRAL.GRID.TextMatrix(J&, 18) = DECONPA$(CPAG%)
        GRILLAGRAL.GRID.TextMatrix(J&, 19) = CUEMADRE$(NCLIE)
        '
      .MoveNext
      Loop
    End With
    '
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    Screen.MousePointer = 1
    
    GRILLAGRAL.Show 1
    Dim MSF0 As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF0 = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF0, CAMPOS$, "LISCLIE", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    
    Screen.MousePointer = 1


90 End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    'nuevo
    Call GELISCLI
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN))
        Text9.TEXT = rasocli$(NCMIN)
    End If
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
   ' Call APERBASDAT("CLIEN")
   ' PRICLI& = 999999: ULTCLI& = 0
   ' Screen.MousePointer = 11
    
    
   Dim ClieTemp As ADODB.Recordset
   Set ClieTemp = New ADODB.Recordset
   ClieTemp.CursorType = adOpenKeyset
   ClieTemp.LockType = adLockReadOnly
   Call APERBASDAT("CLIEN")
   '
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM Deudor12 "
   SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
   SQLX$ = SQLX$ + "ORDER BY Nume_Cli ASC"
   ClieTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
   '
   With ClieTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then pricli& = XVALO(!nume_cli)
     .MoveLast
     If Err < 1 Then ultcli& = XVALO(!nume_cli)
     On Error GoTo 0
   End With
   ClieTemp.Close
   Set ClienTemp = Nothing
   '
    If pricli& < 999999 Then
       Text10 = TRIM$(Str$(pricli&))
       Text9 = rasocli$(pricli&)
    End If
    If ultcli& > 0 Then
       Text8 = TRIM$(Str$(ultcli&))
       Text7.TEXT = rasocli$(ultcli&)
    End If
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
'    If XVALO(Text10.TEXT) > 32767 Then
'       Text10.TEXT = Left$(TRIM$(Text10.TEXT), 4)
'       Exit Sub
'    End If
    If XVALO(Text10.TEXT) < 1 Then Text10.TEXT = "": Exit Sub
    NCMIN = XVALO(Text10.TEXT)
    Text9.TEXT = rasocli$(NCMIN)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If XVALO(Text1.TEXT) < 256 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If XVALO(Text1.TEXT) < 999 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text8_Change()
'    If XVALO(Text8.TEXT) > 32767 Then
'       Text8.TEXT = Left$(TRIM$(Text8.TEXT), 4)
'       Exit Sub
'    End If
    If XVALO(Text8.TEXT) < 1 Then Text8.TEXT = "": Exit Sub
    NCMAX = XVALO(Text8.TEXT)
    Text7.TEXT = rasocli$(NCMAX)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = rasocli$(NCMIN%)
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = rasocli$(NCMAX%)
    End If
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub GELISCLI_VIEJA()

    '
    Screen.MousePointer = 11
    '
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCO% = XVALO(Text1)
    VENDE% = XVALO(Text3)
    '
    MODORDE% = 1
    If FORLICLI07.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLICLI07.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM DEUDOR12 "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN% & " "
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX% & " "
    If TRIM$(Text1) <> "" Then SQLX$ = SQLX$ + "AND GrCo_Cli =  " & GRUCO% & "  "
    If TRIM$(Text3) <> "" Then SQLX$ = SQLX$ + "AND VEND_Cli = " & VENDE% & " "
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
      Case 1
      SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
      Case 2
      SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
      Case 3
      SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set ClieTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ClieTemp
      On Error Resume Next
      .MoveFirst
      'Valida que la tabla no este vacia
      If Err Then
        Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
        Screen.MousePointer = 1
        On Error GoTo 0
        GoTo 90
      End If
      On Error GoTo 0
      '
      '.Move First
      Do While (.EOF = False)
        '
        NC% = !nume_cli
        '
        Label3 = rasocli$(NC%)
        Label3.Refresh
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, rasocli$(NC%), 3, 32)
        Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
        Call REPLA(Y$, CuitCli$(NC%), 51, 16)
        Call REPLA(Y$, POSIVAC$(NC%), 67, 10)
        '\\\OT-06-0081-WAR-06/03/06///
        LPRX% = LIPRCLI%(NC%)
        Call REPLA(Y$, MKI$(LPRX%), 77, 2)
        DDLX$ = TRIM$(DELIPRE$(LPRX%))
        DCCL$ = TRIM$(DESCUCLI$(NC%))
        If DCCL$ <> "" Then
           If InStr(DCCL$, "%") < 1 Then
              DCCL$ = DCCL$ + " %"
           End If
           If Len(DCCL$) > 28 Then DCCL$ = Left$(DCCL$, 28)
           DDLX$ = Left$(DDLX$, 29 - Len(DCCL$))
           DDLX$ = DDLX$ + " (" + DCCL$ + ")"
        End If
        Call REPLA(Y$, DDLX$, 79, 32)
        '\\\OT-06-0081-WAR-FIN///
        Call REPLA(Y$, CUEMADRE$(NC%), 111, 16)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, DOMICLI$(NC%), 3, 32)
        Call REPLA(Y$, FAXCLI$(NC%), 35, 16)
        Call REPLA(Y$, NIBRCLI$(NC%), 51, 16)
        CTRX$ = TRIM$(CATIBRU$(NC%))
        CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                             If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                             If CTRX$ = "2" Then CATIBX$ = "C.Multil."
        Call REPLA(Y$, CATIBX$, 67, 10)
        CPAG% = CPAGCLI%(NC%)
        Call REPLA(Y$, MKI$(CPAG%), 77, 2)
        Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        DOMI$ = TRIM$(CPOSCLI$(NC%))
        If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
        DOMI$ = DOMI$ + LOCACLI$(NC%)
        Call REPLA(Y$, DOMI$, 3, 32)
        Call REPLA(Y$, EMAILCLI$(NC%), 35, 16)
          CUECOX% = VENDCLI%(NC%)
        Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
        Call REPLA(Y$, DEVENDE$(NC%), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("LICLIEN", ULI&)
    Call CIERRARCH("LICLIEN")
    Screen.MousePointer = 1
    Call FINAL("*LICLIEN")
    '
90 End Sub
'
Sub GELISCLI()
    '
    Call GELISCLI_NUEVO
    Exit Sub
    Screen.MousePointer = 11
    '
    NCMIN = XVALO(Text10.TEXT)
    NCMAX = XVALO(Text8.TEXT)
    '
    GRUCO% = XVALO(Text1)
    VENDE% = XVALO(Text3)
    '
    If NCMIN < 1 Then NCMIN = PrimerClienteActivo&
    If NCMAX < 1 Then NCMAX = UltimoClienteActivo&
    '
    If GRUCO% > 0 And VENDE% > 0 Then
      'CON FILTRO DE VENDEDOR Y GRUPO CONTABLE
     If EXISTESQR_ENSQL%("R-LICLIEN3") > 0 Then
      Call FINAL("?R-LICLIEN3(" & NCMIN & "," & NCMAX & "," & VENDE% & "," & GRUCO% & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If
    ElseIf GRUCO% > 0 Then
      'CON FILTRO POR GRUPO CONTABLE
     If EXISTESQR_ENSQL%("R-LICLIEN2") > 0 Then
      Call FINAL("?R-LICLIEN2(" & NCMIN & "," & NCMAX & "," & GRUCO% & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If

    ElseIf VENDE% > 0 Then
      'CON FILTRO POR VENDEDOR
     If EXISTESQR_ENSQL%("R-LICLIEN1") > 0 Then
      Call FINAL("?R-LICLIEN1(" & NCMIN & "," & NCMAX & "," & VENDE% & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If
    Else
      'CON FILTRO SOLO POR CLIENTES
     If EXISTESQR_ENSQL%("R-LICLIEN4") > 0 Then
      Call FINAL("?R-LICLIEN4(" & NCMIN & "," & NCMAX & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If
    End If
    Exit Sub
    
10  MODORDE% = 1
    If FORLICLI07.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLICLI07.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM DEUDOR12 "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN & " "
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX & " "
    If TRIM$(Text1) <> "" Then SQLX$ = SQLX$ + "AND GrCo_Cli =  " & GRUCO% & "  "
    If TRIM$(Text3) <> "" Then SQLX$ = SQLX$ + "AND VEND_Cli = " & VENDE% & " "
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
      Case 1
      SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
      Case 2
      SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
      Case 3
      SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set ClieTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ClieTemp
      On Error Resume Next
      .MoveFirst
      'Valida que la tabla no este vacia
      If Err Then
        Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
        Screen.MousePointer = 1
        On Error GoTo 0
        GoTo 90
      End If
      On Error GoTo 0
      '
      '.Move First
      Do While (.EOF = False)
        '
        NCLIE = !nume_cli
        '
'        Label3 = RASOCLI$(NC%)
'        Label3.Refresh
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKS$(NCLIE), 1, 4)
        Call REPLA(Y$, SAFETEXT$(!raso_cli), 5, 30)
        Call REPLA(Y$, SAFETEXT$(!Tele_Cli), 35, 16)
        Call REPLA(Y$, CuitCli$(!Cuit_Cli), 51, 16)
        Call REPLA(Y$, POSIVAC$(NCLIE), 67, 10)
        '\\\OT-06-0081-WAR-06/03/06///
        LPRX% = LIPRCLI%(NCLIE)
        Call REPLA(Y$, MKI$(LPRX%), 77, 2)
        DDLX$ = TRIM$(DELIPRE$(LPRX%))
        DCCL$ = TRIM$(SAFETEXT$(!Dsta_Cli))
        If DCCL$ <> "" Then
           If InStr(DCCL$, "%") < 1 Then
              DCCL$ = DCCL$ + " %"
           End If
           If Len(DCCL$) > 28 Then DCCL$ = Left$(DCCL$, 28)
           DDLX$ = Left$(DDLX$, 29 - Len(DCCL$))
           DDLX$ = DDLX$ + " (" + DCCL$ + ")"
        End If
        Call REPLA(Y$, DDLX$, 79, 32)
        '\\\OT-06-0081-WAR-FIN///
        Call REPLA(Y$, CUEMADRE$(NCLIE), 111, 16)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKS$(NCLIE), 1, 4)
        Call REPLA(Y$, SAFETEXT$(!raso_cli), 5, 30)
        Call REPLA(Y$, SAFETEXT$(!Faxi_Cli), 35, 16)
        Call REPLA(Y$, SAFETEXT$(!Nibr_Cli), 51, 16)
        CTRX$ = TRIM$(SAFETEXT$(!Caib_Cli))
        CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                             If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                             If CTRX$ = "2" Then CATIBX$ = "C.Multil."
        Call REPLA(Y$, CATIBX$, 67, 10)
        CPAG% = CPAGCLI%(NCLIE)
        Call REPLA(Y$, MKI$(CPAG%), 77, 2)
        Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKS$(NCLIE), 1, 4)
        DOMI$ = TRIM$(CPOSCLI$(NCLIE))
        If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
        DOMI$ = DOMI$ + LOCACLI$(NCLIE)
        Call REPLA(Y$, DOMI$, 5, 30)
        Call REPLA(Y$, EMAILCLI$(NCLIE), 35, 16)
          CUECOX% = VENDCLI%(NCLIE)
        Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
        Call REPLA(Y$, DEVENDE$(NCLIE), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("LICLIEN", ULI&)
    Call CIERRARCH("LICLIEN")
    Screen.MousePointer = 1
    Call FINAL("*LICLIEN")
    '
90 End Sub
'


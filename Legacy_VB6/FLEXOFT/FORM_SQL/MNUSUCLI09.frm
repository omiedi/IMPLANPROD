VERSION 5.00
Begin VB.Form MNUSUCLI09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materiales Suministrados por el Cliente"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4935
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00404080&
      Height          =   6795
      Left            =   4080
      ScaleHeight     =   6735
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   -120
      Width           =   1170
      Begin VB.CommandButton Command2 
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
         Left            =   100
         Picture         =   "MNUSUCLI09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ver e Imprimir Remitos"
         Top             =   940
         Width           =   650
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
         Height          =   640
         Left            =   100
         Picture         =   "MNUSUCLI09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Corregir Formularios / Conexion a Servidor"
         Top             =   2000
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   100
         Picture         =   "MNUSUCLI09.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
         Top             =   230
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   2880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MNUSUCLI09.frx":0896
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3120
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Materiales no Codificados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   140
      TabIndex        =   0
      Top             =   240
      Width           =   3795
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1440
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   13
         Top             =   2050
         Width           =   2100
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
            Height          =   600
            Left            =   0
            Picture         =   "MNUSUCLI09.frx":27B8
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Listados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   750
            TabIndex        =   14
            Top             =   90
            Width           =   1305
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1040
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   1
         Top             =   1500
         Width           =   2100
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MNUSUCLI09.frx":2AC2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Despachar con Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Despacho con Remito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   650
            TabIndex        =   2
            Top             =   50
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   590
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   3
         Top             =   950
         Width           =   2100
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MNUSUCLI09.frx":2DCC
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Registrar Novedades"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Registrar Novedades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   4
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   5
         Top             =   400
         Width           =   2100
         Begin VB.CommandButton Command26 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MNUSUCLI09.frx":2F16
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Nuevo Ingreso de Suministros"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso Material de  Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   600
            TabIndex        =   7
            Top             =   45
            Width           =   1400
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11025
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "MNUSUCLI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub TRACLIEN()
   '
   ' RECUPERA LOS DATOS DEL DAT A LA BASE DE DATOS DE CLIENTES SQL
   ' Deshabilitado 22/03/2011 porque traia quilombo en ARCON
Exit Sub
   '
   FIN& = ULTREG&("DEUDOR1")
   For I& = 1 To FIN&
     TX$ = REGLEIDO$("DEUDOR1", I&)
     TX$ = TX$ + REGLEIDO$("DEUDOR2", I&)
     NCLIE% = CVI(Mid$(TX$, 1, 2))
     RASOCIAL$ = Mid$(TX$, 3, 30)
     DOMICILIO$ = Mid$(TX$, 33, 25)
     NUMERACION% = CVI(Mid$(TX$, 58, 2))
     TOTDOMI$ = DOMICILIO$ & " " & NUMERACION%
     COPOSTAL% = CVI(Mid$(TX$, 60, 2))
     LOCALIDAD$ = Mid$(TX$, 62, 25)
     CUITTT$ = Mid$(TX$, 107, 15)
     PIVAX% = Asc(Mid$(TX$, 136, 1))
     DESCPOSIVA$ = ECOARCH$("CAPOSIV", Chr$(PIVAX%))

     EXISCLI& = CantRegistros&("DEUDOR12", "NUME_CLI =" & NCLIE%)
     '
     If EXISCLI& < 1 Then
        Call CreaRegistro("DEUDOR12", "NUME_CLI", NCLIE%, "NOMESS")
     End If
     '
     CONDI$ = "NUME_CLI = " & NCLIE%
     Call ACTUREGISTRO("DEUDOR12", "Raso_Cli", CONDI$, RASOCIAL$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "Domi_Cli", CONDI$, TOTDOMI$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "Cpos_Cli", CONDI$, COPOSTAL%, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "Loca_Cli", CONDI$, LOCALIDAD$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "Piva_Cli", CONDI$, PIVAX%, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "Cuit_Cli", CONDI$, CUITTT$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("DEUDOR12", "DEPOSIVA", CONDI$, DESCPOSIVA$, REGAF&, "NOMESS")
     '
   Next I&
   

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False

    Call CARGALISEL("INDRREM", "MASUMCLI", "NumeIng/F6.0;REFERENCIA/A56", "STATUS = 2", "DISTINCT")
    Call SELECHO("INDRREM")
    NroIngreso& = XVALO(TRIM$(VALACT1("INDRREM")))
    If NroIngreso& < 1 Then GoTo 99
    '
    REMI$ = VALOBADA("MASUMCLI", "NUMREMITO", "NumeIng = " & NroIngreso&, "NOMESS")
    '
    FREP = CVINT(VALOBADA("MASUMCLI", "FechRemito", "NumeIng = " & NroIngreso&, "NOMESS"))
    '
    NC% = XVALO(VALOBADA("MASUMCLI", "NUME_CLI", "NumeIng = " & NroIngreso& & ""))
    '
    FILTX$ = TRIM$(Str$(NroIngreso&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
    Call STDFORM("REMX-ARC", FILTX$, "PRINT")
    
99  Command2.Enabled = True
End Sub

Private Sub Command23_Click()
   '
    'FALTA VER COMO SE HACE ESTO DE REMITO
15  Call CARGALISEL("INDREMI", "MASUMCLI", "NumeIng/F6.0;REFERENCIA/A56", "STATUS = 1", "DISTINCT")
    Call SELECHO("INDREMI")
    NroIngreso& = XVALO(TRIM$(VALACT1("INDREMI")))
    If NroIngreso& < 1 Then GoTo 99
    '
    REMI$ = VALOBADA("MASUMCLI", "NUMREMITO", "NumeIng = " & NroIngreso&, "NOMESS")
    '
    If TRIM$(REMI$) = "" Then
       PRONREM& = 1 + XVALO(VALOBADA("MASUMCLI", "max( substring( NUMREMITO, 9,8))", " NumeIng > 0"))
       Ing$ = TRIM$(FORMATNUM$(PRONREM&, "F8.0"))
       While Len(Ing$) < 8: Ing$ = "0" + Ing$: Wend
       REMI$ = "RT-0001-" & Ing$ ' Ver si el Formato de Remito esta bien
    End If
    '
    NC% = XVALO(VALOBADA("MASUMCLI", "NUME_CLI", "NumeIng = " & NroIngreso& & ""))
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NroIngreso&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
    Call STDFORM("REMX-ARC", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 15
    End If
    
    Call ACTUREGISTRO("MASUMCLI", "STATUS", "NumeIng = " & NroIngreso&, "2", REGAF&, "NOMESS") 'VERIFICAR COMO SE VA A GRABAR EL STATUS

    '
99  Command23.Enabled = True

End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    '
10  Call CARGALISEL("INDISUMI", "MASUMCLI", "NumeIng/F6.0;REFERENCIA/A56", "NumeIng > 0  AND STATUS = 1  ", "DISTINCT")
    Call SELECHO("INDISUMI/Indice General de Materiales Suministrados por Cliente")
    NroIngreso& = XVALO(TRIM$(VALACT1("INDISUMI")))
    If NroIngreso& < 1 Then GoTo 99
    '
    NC% = XVALO(VALOBADA("MASUMCLI", "NUME_CLI", "NumeIng = " & NroIngreso& & ""))
    FREP = CVINT(VALOBADA("MASUMCLI", "FechRecep", "NumeIng = " & NroIngreso& & ""))
'    FILTX$ = TRIM$(Str$(NroIngreso&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
    FILTX$ = TRIM$(Str$(NroIngreso&)) & ";" & TRIM$(Str$(NC%)) & ";" & TRIM$(Str$(FREP)) & ";" & USERNAME$
    '
    Call STDFORM("MASC-ARC", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 10
    End If
        
99  Command25.Enabled = True

End Sub

Private Sub Command26_Click()
    Command26.Enabled = False
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then GoTo 99
    '
15  CONDI$ = "NumeIng > 0"
    NroIngreso& = 1 + XVALO(VALOBADA("MASUMCLI", "MAX(NumeIng)", CONDI$, "NOMESS"))
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NroIngreso&)) & ";" & TRIM$(Str$(NC%)) & ";" & TRIM$(Str$(FREP)) & ";" & USERNAME$
    
    Call STDFORM("MASC-ARC", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 10
    End If
    '
    
    DESKRIPCION$ = VALOFORM("Text-2")
    If TRIM$(DESKRIPCION$) = "" Then
       Call MENSERR(24, "Debe Ingresar la Descripcion ")
       GoTo 15
    End If
    REFER$ = TRIM$(Left$(DESKRIPCION$, 22)) & " - " & NC% & " - " & Left$(RASOCLI$(NC%), 22) ' TOTAL 55
    '
    Call ACTUREGISTRO("MASUMCLI", "NUME_CLI", "NumeIng = " & NroIngreso&, NC%, REGAF&, "NOMESS") ' LO AGREGUE  ACA POR QUE NO LO GRABABA EL DISEÑADOR
    Call ACTUREGISTRO("MASUMCLI", "REFERENCIA", "NumeIng = " & NroIngreso&, REFER$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MASUMCLI", "STATUS", "NumeIng = " & NroIngreso&, "1", REGAF&, "NOMESS")  'VERIFICAR COMO SE VA A GRABAR EL STATUS
    Call ACTUREGISTRO("MASUMCLI", "USUAPFOR_1", "NumeIng = " & NroIngreso&, USERNAME$, REGAF&, "NOMESS")
    
99  Command26.Enabled = True

End Sub

Private Sub Command29_Click()

End Sub

Private Sub Command5_Click()
     Command5.Enabled = False
    '
     RTA1% = ALTERNA%("Corregir Formularios \Conexión a Servidor")
     If RTA1% < 1 Or RTA1% > 2 Then GoTo 99
     '
     If RTA1% = 1 Then
        If EXISTE%("DSGFRM09.EXE") > 0 Then
           Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
        End If
     ElseIf RTA1% = 2 Then
        CONN_SQLS.Show 1
     End If
     
     '
99   Command5.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
10 OPX% = ALTERNA%("Ingresos Por Fecha \Ingresos por Cliente y Rango de Fechas\Ingresos Sin Remitir\Ingresos Remitidos Por Rango de Fechas\Pendientes Unificados")
   '
   If OPX% < 1 Or OPX% > 5 Then GoTo 99
   '
   If OPX% = 1 Then
      Call FINAL("?INGREFEC")
      GoTo 10
   ElseIf OPX% = 2 Then
      Call FINAL("?INGPORCL")
      GoTo 10
   ElseIf OPX% = 3 Then
      Call FINAL("?INGSINRE")
      GoTo 10
   ElseIf OPX% = 4 Then
      Call FINAL("?INGCLSIR")
      GoTo 10
   ElseIf OPX% = 5 Then
      Call REPARAUNIFI ' LISTADO DE REPARACIONES E INGRESO DE SUMINISTROS UNIFICADO
      GoTo 10
   End If
   
   
99 Command8.Enabled = True

End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    Call ABRESTOCKS
    Call ABRECLIEN
    '
    Screen.MousePointer = 11
    '
    MNUSUCLI09.Visible = True
    ''' Call TRACLIEN ' ACTUALIZA LA BASE DE DATOS DESDE EL DAT A SQL
                      ' Deshabilitado 22/03/2011 porque traia quilombo
    Screen.MousePointer = 1
    '
End Sub

Sub REPARAUNIFI()

  Call SETULTREG("PRESRUNI", 0)
  FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, UI&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
        If CVI(Mid$(XX$, 123, 2)) < 1 Then
           NUREPAR& = CVS(Mid$(XX$, 1, 4)) ' DATOS DE PRESREPA
           DESCR$ = Mid$(XX$, 5, 60)
           NC% = CVI(Mid$(XX$, 67, 2))
           FEING% = CVI(Mid$(XX$, 121, 2))
           REFER$ = "REP."
           '
           Call REPLA(TX$, MKS$(NUREPAR&), 1, 4) ' GRABO EN PRESRUNI
           Call REPLA(TX$, DESCR$, 5, 64)
           Call REPLA(TX$, MKI$(NC%), 69, 2)
           Call REPLA(TX$, MKI$(FEING%), 71, 2)
           Call REPLA(TX$, REFER$, 73, 12)
           JJ& = JJ& + 1
           Call GRAREG("PRESRUNI", TX$, JJ&)
        End If
   Next U&
   '
   SQLX$ = "SELECT * FROM MASUMCLI WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE STATUS = 1"
   SQLX$ = SQLX$ + " ORDER BY NUMEING"
   Dim MASUTMP As New ADODB.Recordset
   Set MASUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MASUTMP
     Do While Not .EOF
        JJ& = JJ& + 1
        XX$ = REGBLAN$("PRESRUNI")
        Call REPLA(XX$, MKS$(XVALO(!NumeIng)), 1, 4)
        Call REPLA(XX$, SAFETEXT$(!Descripcion), 5, 64)
        Call REPLA(XX$, MKI$(XVALO(!NUME_CLI)), 69, 2)
        Call REPLA(XX$, MKI$(CVINT(!FechRecep)), 71, 2)
        Call REPLA(XX$, "MSC.", 73, 12)
        Call GRAREG("PRESRUNI", XX$, JJ&)
       .MoveNext
     Loop
   End With
   '
   MASUTMP.Close
   Set MASUTMP = Nothing
   Call SETULTREG("PRESRUNI", JJ&)
   Call FILESORT("PRESRUNI", "", 5, 5, "ASC")
   Call CIERRARCH("PRESRUNI")
   '
   Screen.MousePointer = 1
   '
   Call FINAL("*LIREPUNI")
End Sub

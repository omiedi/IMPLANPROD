VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FAPEPAG07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obligaciones Pendientes de Pago"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9210
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
   ScaleHeight     =   3360
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "GRUPO CONTABLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   105
      TabIndex        =   26
      Top             =   2400
      Width           =   4005
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
         Left            =   1260
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   2535
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
         Left            =   960
         TabIndex        =   28
         Top             =   360
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
         Left            =   360
         MaxLength       =   2
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   8295
      ScaleHeight     =   3555
      ScaleWidth      =   945
      TabIndex        =   23
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   32
         Top             =   2880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
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
         Left            =   120
         Picture         =   "Fapepag07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   120
         Picture         =   "Fapepag07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C7D9CC&
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
      Height          =   2250
      Left            =   120
      TabIndex        =   10
      Top             =   90
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
         Left            =   2640
         TabIndex        =   16
         Top             =   420
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
         Height          =   300
         Left            =   3465
         TabIndex        =   15
         Top             =   420
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   780
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
         TabIndex        =   13
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
         Height          =   300
         Left            =   3465
         TabIndex        =   12
         Top             =   1365
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1720
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
         Left            =   240
         TabIndex        =   18
         Top             =   525
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
         TabIndex        =   17
         Top             =   1470
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C7D9CC&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4200
      TabIndex        =   7
      Top             =   2450
      Width           =   4005
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   3360
         Picture         =   "Fapepag07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   105
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Pantalla"
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
         Index           =   4
         Left            =   525
         TabIndex        =   9
         Top             =   370
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Impresión"
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
         Index           =   8
         Left            =   1890
         TabIndex        =   8
         Top             =   370
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
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
      Height          =   765
      Left            =   4200
      TabIndex        =   6
      Top             =   1620
      Width           =   4005
      Begin VB.TextBox Text6 
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
         Left            =   2160
         TabIndex        =   21
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   3240
         TabIndex        =   20
         Top             =   270
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Vencimiento"
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
         TabIndex        =   22
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
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
      Height          =   1470
      Left            =   4200
      TabIndex        =   1
      Top             =   90
      Width           =   4005
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Por Fecha de Vencimiento"
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
         Index           =   3
         Left            =   420
         TabIndex        =   5
         Top             =   1080
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
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
         Left            =   2205
         TabIndex        =   4
         Top             =   720
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
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
         Left            =   735
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Por Proveedor"
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
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   360
         Width           =   1905
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "Fapepag07.frx":075E
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Fapepag07.frx":07D8
      Top             =   1200
   End
   Begin VB.Menu otrasConsultas 
      Caption         =   "Otras consultas"
      Begin VB.Menu ComposicionDeuda 
         Caption         =   "Composición de Deuda a Fecha"
      End
      Begin VB.Menu ComposicionDeudaProveedor 
         Caption         =   "Composición de Deuda por Proveedor"
      End
      Begin VB.Menu ObligacionesPagarPorFecha 
         Caption         =   "Obligaciones a Pagar por Rango de Fecha"
      End
   End
End
Attribute VB_Name = "FAPEPAG07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TICUE%(32767)
Dim MODOMONEI%

Private Sub Command1_Click()
   Call SELECHO("GRUCOPRO")
   VDEVU$ = VALACT1$("GRUCOPRO")
   If TRIM$(VDEVU$) <> "" Then
      Text1.TEXT = VDEVU$
   End If
End Sub

Private Sub Command2_Click()
    Call FAPEPAG2
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COPAPE1")
            Exit Sub
          Else
            Call FINAL("*COPAPE2")
            Exit Sub
         End If
       Else                          ' impresora
           If MODOMONEI% = 2 Then
             If EXISTE%(LUDAT$ + "FAPEPAGD.RFS") > 0 Then
                If Option1(3).Value = False Then
                   Call FINAL("*LIPAPED1")
                   Exit Sub
                 Else
                   Call FINAL("*LIPAPED2")
                   Exit Sub
                End If
             Else
                Call COMUNI("Falta Controlador fapegagd, Informe a su Supervisor en Sistemas")
             End If
           Else
             If Option1(3).Value = False Then
               Call FINAL("*LIPAPE1")
               Exit Sub
             Else
               Call FINAL("*LIPAPE2")
               Exit Sub
             End If
           End If
    End If
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPEPAG1
        CARFAPEN% = 1
    End If
    '
End Sub


Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub ComposicionDeudaAFecha()

    Dim SALINI#(999999)
    Call SELECHO("SELFECHA")
    
    VDEVU$ = VALACT1$("SELFECHA")
    If VDEVU$ <> "" Then
      FECOR% = FECHANUM(VDEVU$)
      If FECOR% > 0 Then
        Screen.MousePointer = 11
        FECORTE$ = FETEXCOR1$(FECOR%)
        '
        TABORI$ = "CAJABANC"
        'PONE UN SETULTREG DEL ARCHIVO QUE SE USA PARA GRABAR LA DEUDA EN CERO
        Call COPYSTRU("FAPECOBA", "CODEUPRO")
        Call SETULTREG("CODEUPRO", 0)
        '
        ' DETERMINA SALDO A FECHA POR CUENTA
        Dim CABATEMP As ADODB.Recordset
        Set CABATEMP = New ADODB.Recordset
        CABATEMP.CursorLocation = adUseClient
        Call ABRECONEXION
10      SQLX$ = "SELECT NUPROVE, SUM(IHABECOMP-IDEBECOMP) AS SALDOIN FROM " + TABORI$ + " "
        SQLX$ = SQLX$ + "WHERE FechConta <= '" & FECORTE$ & "' "
        SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
        SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
        SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
        SQLX$ = SQLX$ + "GROUP BY NUPROVE "
        '
25      On Error Resume Next
        CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        On Error Resume Next
        If Err Then
           On Error GoTo 0
           Screen.MousePointer = 1
           Call MENSERR(24, "Imposible Determinar Saldos")
           GoTo 99
        End If
        On Error GoTo 0
        '
        JU& = 0: CLIMAX = 0
        With CABATEMP
           If Not .BOF And Not .EOF Then
             .MoveLast
             FIN& = .RecordCount
             .MoveFirst
             JU& = 0
             Do While Not .EOF
               JU& = JU& + 1
               Call TRACE(20, JU&, FIN&)
               NCLI = XVALO(!NUPROVE)
               If NCLI > 0 Then
                 If TICUEPRO((-1) * NCLI) = 0 Then ' solo cuentas corrientes
                   If NCLI > CLIMAX Then CLIMAX = NCLI
                   SALINI#(NCLI) = SALINI#(NCLI) + !SALDOIN
                   If NCLI > CLIMAX Then CLIMAX = NCLI
                 End If
               End If
             .MoveNext
             Loop
           End If
        End With
        CABATEMP.Close
        '
        ' DETERMINA COMPOSICION DE DEUDA DE CADA PROVEEDOR
        JJ& = 0
        For NCLI = 1 To CLIMAX
           FIN& = CLIMAX
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           If Abs(SALINI#(NCLI)) >= 0.005 Then
              Call RECADEUDA(NCLI, SALINI#(NCLI), FECORTE$)
           End If
        Next NCLI
        
        Call CIERRARCH("CODEUPRO")
        '
        Screen.MousePointer = 1
        Call HEADERS("CODEUPRO", "")
        Call HEADERS("CODEUPRO", "Fecha de Corte al: " + FECHATEX$(FECOR%))
        '
        Call FINAL("*CODEUPRO")
      End If
    End If

99  Exit Sub

End Sub

Sub RECADEUDA(NCLI, SALIN#, FECORTE$)

    Call ABRECONEXION
    
    fechaCorte% = CVINT(FECORTE$)
    
    Dim rst1 As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM SACREPEN WITH(NOLOCK) WHERE NUPROVE=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND FECHEMISI <='" & FECORTE$ & "'"
    
    Set rst1 = READSET(FILTSQL$(ConsultaSql$))
    With rst1
        Do While Not .EOF
            comprobanteNro$ = SAFETEXT$(!CODICOM_COR)
            saldoComprobante# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
            importeOriginal# = XVALO(!IBruComp)
            FechaComprobante% = CVINT(!FECHEMISI)
            
            GoSub GRADEURECU
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rst1.Close
    Set rst1 = Nothing
    On Error GoTo 0


    Dim rst2 As ADODB.Recordset
    ConsultaSql$ = "SELECT CODICOM_LAR FROM CAJABANC WITH(NOLOCK) WHERE NUPROVE=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND (TIPOMOVIM='OPAG' OR TIPOMOVIM='AJUP' OR (TIPOMOVIM='FCOM' AND OPCIMOVIM='NC'))" ' FALTA COMPOSICION ENTRE CUENTAS
    ConsultaSql$ = ConsultaSql$ & " AND FECHCONTA >'" & FECORTE$ & "'"
    ConsultaSql$ = ConsultaSql$ & " GROUP BY CODICOM_LAR"
    
    Set rst2 = READSET(FILTSQL$(ConsultaSql$))
    With rst2
        Do While Not .EOF
            opagoNro$ = SAFETEXT$(!CODICOM_LAR)
            Dim rst3 As ADODB.Recordset
            ConsultaSql$ = "SELECT * FROM APLIPAGO WITH(NOLOCK) WHERE NUPROVE=" & CStr(NCLI)
            ConsultaSql$ = ConsultaSql$ & " AND CODOCAP_LAR='" & opagoNro$ & "'"
            Set rst3 = READSET(FILTSQL$(ConsultaSql$))
            With rst3
                Do While Not .EOF
                    facturaNro$ = SAFETEXT$(!codocor_lar)
                    If InStr(1, EMPREAC$, "FERVI") > 0 And Mid$(opagoNro$, 1, 2) = "FF" And (NCLI = 264 Or NCLI = 452 Or NCLI = 529 Or NCLI = 960 Or NCLI = 976 Or NCLI = 1186 Or NCLI = 1213 Or NCLI = 1516 Or NCLI = 1661 Or NCLI = 1682 Or NCLI = 1726 Or NCLI = 1750 Or NCLI = 185 Or NCLI = 1941 Or NCLI = 1991 Or NCLI = 2002 Or NCLI = 2015 Or NCLI = 2034 Or NCLI = 2107 Or NCLI = 2131 Or NCLI = 2144 Or NCLI = 1388 Or NCLI = 1723 Or NCLI = 1751 Or NCLI = 1847 Or NCLI = 2034) Then
                        FechaFactura% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & facturaNro$ & "' AND NUPROVE=" & CStr(NCLI)))
                        If FechaFactura% = 0 Then
                            facturaNro$ = SAFETEXT$(!CODOCOR_COR)
                            FechaFactura% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_COR='" & facturaNro$ & "' AND NUPROVE=" & CStr(NCLI)))
                        End If
                    Else
                        FechaFactura% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & facturaNro$ & "' AND NUPROVE=" & CStr(NCLI)))
                    End If
                    
                    If FechaFactura% <= fechaCorte% And FechaFactura% > 0 Then
                        comprobanteNro$ = SAFETEXT$(!CODOCOR_COR)
                        saldoComprobante# = XVALO(!IMPOAPLI)
                        importeOriginal# = XVALO(!ITOTORIG)
                        FechaComprobante% = FechaFactura%
                        
                        GoSub GRADEURECU
                    End If
                    
                    .MoveNext
                Loop
            End With
            
            On Error Resume Next
            rst3.Close
            Set rst3 = Nothing
            On Error GoTo 0
            
            .MoveNext
        Loop
    End With

    On Error Resume Next
    rst2.Close
    Set rst2 = Nothing
    On Error GoTo 0
    
    Dim rst4 As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NUPROVE=" & CStr(NCLI)
    ConsultaSql$ = ConsultaSql$ & " AND ((TIPOMOVIM='FCOM' AND OPCIMOVIM='FC') OR (TIPOMOVIM='FCOM' AND OPCIMOVIM='ND') OR TIPOMOVIM='AJUP')" ' FALTA COMPOSICION ENTRE CUENTAS
    ConsultaSql$ = ConsultaSql$ & " AND FECHCONTA >'" & FECORTE$ & "'"
    
    Set rst4 = READSET(FILTSQL$(ConsultaSql$))
    With rst4
        Do While Not .EOF
            facturaNro$ = SAFETEXT$(!CODICOM_LAR)
            Dim rst5 As ADODB.Recordset
            ConsultaSql$ = "SELECT * FROM APLIPAGO WITH(NOLOCK) WHERE NUPROVE=" & CStr(NCLI)
            ConsultaSql$ = ConsultaSql$ & " AND CODOCOR_LAR='" & facturaNro$ & "'"
            Set rst5 = READSET(ConsultaSql$)
            With rst5
                Do While Not .EOF
                    opagoNro$ = SAFETEXT$(!codocap_lar)
                    fechaOPago% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & opagoNro$ & "' AND NUPROVE=" & CStr(NCLI)))
                    
                    If fechaOPago% <= fechaCorte% And fechaOPago% > 0 Then
                        comprobanteNro$ = SAFETEXT$(!CODOCAP_COR)
                        saldoComprobante# = (-1) * XVALO(!IMPOAPLI)
                        importeOriginal# = (-1) * XVALO(!ITOTORIG)
                        FechaComprobante% = fechaOPago%
                        
                        GoSub GRADEURECU
                    End If
                    
                    .MoveNext
                Loop
            End With
            
            On Error Resume Next
            rst5.Close
            Set rst5 = Nothing
            On Error GoTo 0
            
            .MoveNext
        Loop
    End With

    On Error Resume Next
    rst4.Close
    Set rst4 = Nothing
    On Error GoTo 0


    Exit Sub

GRADEURECU:
   SUMADEU# = 0

        y$ = REGBLAN$("CODEUPRO")
        '
        Call REPLA(y$, MKS$(NCLI), 1, 4)
        Call REPLA(y$, rasocli$((-1) * NCLI), 5, 28)
        Call REPLA(y$, TEL1CLI$((-1) * NCLI), 33, 16)
        Call REPLA(y$, comprobanteNro$, 49, 12)
        Call REPLA(y$, MKI$(FechaComprobante%), 61, 2)
        Call REPLA(y$, MKD$(importeOriginal#), 63, 8)
        Call REPLA(y$, MKD$(saldoComprobante#), 73, 8)
        '
        If Abs(saldoComprobante#) >= 0.005 Then
           Call REGAPP("CODEUPRO", y$)
           SUMADEU# = SUMADEU# + saldoComprobante#
        End If
    Return
    
End Sub
 Sub ObligacionesPagarPorFecha_Click()

    Call DESHASFECHA(Des%, Has%, PERIO$)
    If Des% < 1 Then GoTo 99
    '
    CORRESP$ = PERIO$
    Call HEADERS("VENXRAFE", "")
    Call HEADERS("VENXRAFE", "Corresponde a: " + CORRESP$)
    '
    ' JANDY SQL
    Call APERBASDAT("CABAN")
    '
    Dim VEXRAFETMP As ADODB.Recordset
    Set VEXRAFETMP = New ADODB.Recordset
    VEXRAFETMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND FeVencim>= '" & FETEXCOR1$(Des%) & "' "
    SQLX$ = SQLX$ + "AND FeVencim<= '" & FETEXCOR1$(Has%) & "' "
    SQLX$ = SQLX$ + "AND (SaldAcre)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
25      On Error Resume Next
    VEXRAFETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With VEXRAFETMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Vencimientos en el Rango de Fecha Seleccionado")
        GoTo 98
      End If
      On Error GoTo 0
      J& = 0
      y$ = REGBLAN("VENXRAFE")
      Do While .EOF = False
          '
          NC% = !NUPROVE
          NFS$ = !CODICOM_LAR
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldAcre - !SaldDebe
          '
          Call REPLA(y$, MKI$(NC%), 1, 2)
          Call REPLA(y$, CODICLI$((-1) * NC%), 3, 24)
          Call REPLA(y$, rasocli$((-1) * NC%), 27, 30)
          Call REPLA(y$, NFS$, 57, 18)
          Call REPLA(y$, MKI$(FVI%), 75, 2)
          Call REPLA(y$, MKD$(SALDEU#), 77, 8)
          '
          J& = J& + 1
          Call GRAREG("VENXRAFE", y$, J&)
          '
        .MoveNext
        '
      Loop
      '
      Call SETULTREG("VENXRAFE", J&)
      Call FILESORT("VENXRAFE", "VENXRAFE", 5, 5, "ASC")
      Call FILESORT("VENXRAFE", "VENXRAFE", 1, 1, "ASC")
      '
    End With
    '
    Call FINAL("*LIVERAFE")
    '
98      VEXRAFETMP.Close
    Set VEXRAFETMP = Nothing
            
    '
99  Exit Sub
    '
End Sub
Sub ComposicionDeudaPorProveedorFecha()

    Call SELECHO("SELFECHA")
    
    VDEVU$ = VALACT1$("SELFECHA")
    If TRIM$(VDEVU$) = "" Then Exit Sub
    
    NCLI = ListaSeleccionProveedores()
    If NCLI > 0 Then
        Call COPYSTRU("FAPECOBA", "CODEUPRO")
        Call BLANARCH("CODEUPRO")
        FECHA$ = FETEXCOR1$(FECHANUM(VDEVU$))
        Call RECADEUDA(NCLI, SALIN#, FECHA$)
        
        Call CIERRARCH("CODEUPRO")
        '
        Screen.MousePointer = 1
        Call HEADERS("CODEUPRO", "")
        Call HEADERS("CODEUPRO", "Fecha de Corte al: " + VDEVU$)
        '
        Call FINAL("*CODEUPRO")
    End If

End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub


Private Sub Command8_Click()
'    Call SELECHO("ACREDOR")
'    NCMAX% = XVALO(VALACT1$("ACREDOR"))
    NCMAX% = ListaSeleccionProveedores()
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = rasocli$((-1) * NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
'    Call SELECHO("ACREDOR")
'    NCMIN% = XVALO(VALACT1$("ACREDOR"))
    NCMIN% = ListaSeleccionProveedores()
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = rasocli$((-1) * NCMIN%)
    End If
End Sub

Private Sub ComposicionDeuda_Click()

    Call ComposicionDeudaAFecha
    
End Sub

Private Sub ComposicionDeudaProveedor_Click()

    Call ComposicionDeudaPorProveedorFecha
    
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    MODOMONEI% = 0
    
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text1_Change()
   ECO$ = ECOARCH$("GRUCOPRO", MKI$(XVALO(Text1)))
   '
   If ECO$ <> "" Then
        Text2 = ECO$
      Else
        Text2 = "No Definido"
   End If
End Sub

Private Sub Text1_DblClick()
   Call Command1_Click
End Sub

Private Sub Text10_Change()
    NCMIN% = XVALO(Text10.TEXT)
    Text9.TEXT = rasocli$((-1) * NCMIN%)
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text8_Change()
    NCMAX% = XVALO(Text8.TEXT)
    Text7.TEXT = rasocli$((-1) * NCMAX%)
    Text7.Refresh
End Sub

Private Sub Text10_DblClick()
    Call Command9_Click
'    Call SELECHO("ACREDOR")
'    NCMIN% = XVALO(VALACT1$("ACREDOR"))
'    If NCMIN% > 0 Then
'        Text10.TEXT = TRIM$(Str$(NCMIN%))
'        Text9.TEXT = rasocli$((-1) * NCMIN%)
'    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call Command8_Click
'    Call SELECHO("ACREDOR")
'    NCMAX% = XVALO(VALACT1$("ACREDOR"))
'    If NCMAX% > 0 Then
'        Text8.TEXT = TRIM$(Str$(NCMAX%))
'        Text7.TEXT = rasocli$((-1) * NCMAX%)
'    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    Dim NROPO$(8196), CLIPO%(8196)
    Dim TICUE%(32767), SARRACLI#(32767), SUDEUCLI#(32767)
    'Dim TIS$(10)
    'TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    'For I% = 0 To 7
    '    TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    'Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
'    FIN& = ULTREG&("PROVED1")
'    For U& = 1 To FIN&
'       Call TRACE(20, U&, FIN&)
'       PP1$ = REGLEIDO$("PROVED1", U&)
'       NC% = CVI(Left$(PP1$, 2))
'       If NC% > 0 Then
'          PP2$ = REGLEIDO$("PROVED2", U&)
'          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
'          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
'             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
'             TICUE%(NC%) = 0
'          End If
'       End If
'    Next U&
    '
    FIN& = ULTREG&("CUECORR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCLIE% = CVI(Mid$(X$, 5, 2))
      If NCLIE% > 0 Then
        If TICUEPRO%((-1) * NCLIE%) = 0 Then
          FEDOC% = CVI(Left$(X$, 2))
          If FEDOC% <= FLIM% Then
            IMPO# = CVD(Mid$(X$, 123, 8))
            IMPO# = IMPO# - CVD(Mid$(X$, 115, 8))
            SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
            SARRATO# = SARRATO# + IMPO#
            '
            SALDEX$ = Mid$(X$, 139, 60)
            For KKI% = 1 To 6
              FVI% = CVI(Mid$(SALDEX$, 10 * KKI% - 9, 2))
              SALDO# = CVD(Mid$(SALDEX$, 10 * KKI% - 7, 8))
              If FVI% > 0 Then
                If Abs(SALDO#) >= 0.005 Then
                  FF% = CVI(Left$(X$, 2))
                  NC% = CVI(Mid$(X$, 5, 2))
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
                  es$ = "    "
                  If SALDO# < 0 Then
                      es$ = " CR "
                    Else
                      If FVI% <= HOO% + 8 Then es$ = " ** "
                      If FVI% <= HOO% Then es$ = "!!!!"
                  End If
                  '
                  Z$ = REGBLAN$("FAPEPAG")
                  Call REPLA(Z$, MKI$(NC%), 1, 2)
                  Call REPLA(Z$, NFS$, 3, 18)
                  Call REPLA(Z$, MKI$(FF%), 21, 2)
                  Call REPLA(Z$, MKI$(FVI%), 23, 2)
                  Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                  Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
                  Call REPLA(Z$, es$, 41, 4)
                  '
                  UJ& = UJ& + 1
                  Call GRAREG("FAPEPAG", Z$, UJ&)
                  SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDO#
                  SUDEUTO# = SUDEUTO# + SALDO#
                  '
                End If
              End If
            Next KKI%
          End If
          If FEDOC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO$(CAMOPO%) = Mid$(X$, 7, 18)
            CLIPO%(CAMOPO%) = NCLIE%
          End If
        End If
      End If
    Next U&
    '
    '
    ' LEER APLIPAG - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLIPAG")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLIPAG", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            OPOSTE$ = Mid$(X$, 5, 18)
            DOCUORI$ = Mid$(X$, 55, 18)
            'If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   'AAA = BBB
            'End If
            For KKI% = 1 To CAMOPO%
              If NROPO$(KKI%) = DOCUORI$ Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            'GoTo 29
            '
24          Z$ = REGBLAN$("FAPEPAG")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
            NFS$ = DOCUORI$
            Call REPLA(Z$, NFS$, 3, 18)
            FF% = FLIM%
            Call REPLA(Z$, MKI$(FF%), 21, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 23, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPEPAG", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
   '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          RFF$ = RECFILT$("CUECORR", 4, MKI$(CLIAPLI%))
          For U& = Len(RFF$) To 1 Step -4
            REG& = CVS(Mid$(RFF$, U& - 3, 4))
            If REG& > 0 Then
              If REG& <= ULTREG&("CUECORR") Then
                X$ = REGLEIDO$("CUECORR", REG&)
                If CVI(Left$(X$, 2)) <= FLIM% Then
                  FF% = CVI(Left$(X$, 2))
                  FVI% = FLIM%
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
                  es$ = "    "
                  If Abs(IMORIG# - DIFESAL#) < 0.005 Then
                    SALDO# = IMORIG#
                    '
                    Z$ = REGBLAN$("FAPEPAG")
                    Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
                    Call REPLA(Z$, NFS$, 3, 18)
                    Call REPLA(Z$, MKI$(FF%), 21, 2)
                    Call REPLA(Z$, MKI$(FVI%), 23, 2)
                    Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                    Call REPLA(Z$, MKD$(Abs(IMORIG#)), 33, 8)
                    Call REPLA(Z$, es$, 41, 4)
                    '
                    UJ& = UJ& + 1
                    Call GRAREG("FAPEPAG", Z$, UJ&)
                    SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDO#
                    SUDEUTO# = SUDEUTO# + SALDO#
                    GoTo 39
                    '
                  End If
                End If
              End If
            End If
          Next U&
      End If
39  Next UI&
          
              
     For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
           
          Z$ = REGBLAN$("FAPEPAG")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 18)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 21, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 23, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
          IMORIG# = Abs(DIFESAL#)
          Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPEPAG", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          SUDEUTO# = SUDEUTO# + SALDEU#
          '
      End If
49  Next UI&
    '
50  Call SETULTREG("FAPEPAG", UJ&)
    Call CIERRARCH("FAPEPAG")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    Screen.MousePointer = 11
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPEPAG")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPEPAG", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 23, 2))
            '
            TLISTA$ = Space$(36)
            If MODORDE% = 1 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
              ElseIf MODORDE% = 2 Then
                 Call REPLA(TLISTA$, rasocli$((-1) * NC%), 1, 30)
              ElseIf MODORDE% = 3 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
            End If
            Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            List1.AddItem TLISTA$
            '
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = XVALO(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPEPAG", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 18)
        FF% = CVI(Mid$(X$, 21, 2))
        FVI% = CVI(Mid$(X$, 23, 2))
        SALDEU# = CVD(Mid$(X$, 25, 8))
        IMORIG# = CVD(Mid$(X$, 33, 8))
        es$ = Mid$(X$, 41, 4)
        '
        y$ = REGBLAN$("FAPEPAGX")
        Call REPLA(y$, MKI$(NC%), 1, 2)
        Call REPLA(y$, rasocli$((-1) * NC%), 3, 30)
        Call REPLA(y$, TEL1CLI$((-1) * NC%), 33, 16)
        Call REPLA(y$, NFS$, 49, 18)
        Call REPLA(y$, MKI$(FF%), 67, 2)
        Call REPLA(y$, MKD$(IMORIG#), 69, 8)
        Call REPLA(y$, MKI$(FVI%), 77, 2)
        Call REPLA(y$, MKD$(SALDEU#), 79, 8)
        Call REPLA(y$, es$, 95, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPEPAGX", y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    '
    FEX = FLIM%
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIPAPE1")
    '
    Screen.MousePointer = 1

End Sub


Sub FAPEPAG1()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
      If NCX% > 0 Then
        If NCX% > NCMAX% Then NCMAX% = NCX%
        If NCX% < NCMIN% Then NCMIN% = NCX%
        TICUE%(NCX%) = 0
      End If
    Next U&
    Text10 = TRIM$(Str$(NCMIN%))
    Text9 = rasocli$((-1) * NCMIN%)
    Text8 = TRIM$(Str$(NCMAX%))
    Text7 = rasocli$((-1) * NCMAX%)
    '
    Dim SACTEMP As ADODB.Recordset
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    'Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set SACTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    On Error Resume Next
    FLIM = HOY(HOS$)
    With SACTEMP
      .MoveFirst
      If Err < 1 Then
        FLIM = CVINT(!FEVENCIM)
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    
    On Error GoTo 0
    '
    Text6 = FECHATEX$(FLIM)
    Text6.Refresh
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub FAPEPAG2()
    '
    Screen.MousePointer = 11
    FEXI% = FECHANUM(Text6.TEXT)
    FLIM = FETEXCOR1$(FEXI%)
    
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCOSEL% = XVALO(Text1)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
'    Call RECONSTRUYE_DEUDA_DE_PROVEEDORES_CON_DIFERENCIA
    '
    'JANDY SQL
    Call APERBASDAT("CABAN")
    
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = New ADODB.Recordset
    SACTEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND NuProve >= " & NCMIN% & " "
    SQLX$ = SQLX$ + "AND NuProve <= " & NCMAX% & " "
    SQLX$ = SQLX$ + "AND FeVencim<= '" & FLIM & "' "
    SQLX$ = SQLX$ + "AND (SaldAcre + SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    'Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    SACTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    On Error Resume Next
    With SACTEMP
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        .MoveFirst
        UJ& = 0
        NIVCONXXX% = NIVCONFI%

        Do While Not .EOF
          Call TRACE(20, UJ&, FIN&)
          NC% = !NUPROVE
            If TICUEPRO%((-1) * NC%) = 0 Then ' SOLO CUENTAS CORRIENTES
              If GRUCOSEL% < 1 Or GRUCOCLI%((-1) * NC%) = GRUCOSEL% Then
                '\\\OT-08-1032-OD-11/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
                If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70 'VER QUE FACTURAS MOSTRAR EN LISTADO
                '\\\OT-08-1032-OD-FIN
                NFS$ = !CODICOM_LAR
                FF% = CVINT(!FECHEMISI)
                FVI% = CVINT(!FEVENCIM)
                SALDEU# = !SaldAcre - !SaldDebe
                SALDEUDOL# = 0: MEMI% = 0
                On Error Resume Next
                SALDEUDOL# = !SaldAcre_DOLA - !SaldDebe_DOLA
                MEMI% = !Mone_Emis: If MEMI% < 0 Then MEMI% = 0
                If MEMI% < 2 Then SALDEUDOL# = 0
                On Error GoTo 0
                ' ESTA ES PARA DESCARTAR LOS CASOS DE FACTURAS EN PESOS QUE QUEDARON
                ' CON ALGÚN SALDITO EN DOLARES - RECLAMO CLAUDIA MATSUSAKA - CT-100294-FL
                If Abs(SALDEU#) < 0.005 Then
                  If Abs(SALDEUDOL#) < 0.005 Then
                    GoTo 70
                  End If
                End If
                IMORIG# = !IBruComp
                es$ = "    "
                
                '
                y$ = REGBLAN$("FAPEPAGX")
                Call REPLA(y$, MKI$(NC%), 1, 2)
                Call REPLA(y$, rasocli$((-1) * NC%), 3, 30)
                Call REPLA(y$, TEL1CLI$((-1) * NC%), 33, 16)
                Call REPLA(y$, NFS$, 49, 18)
                Call REPLA(y$, MKI$(FF%), 67, 2)
                Call REPLA(y$, MKD$(IMORIG#), 69, 8)
                Call REPLA(y$, MKI$(FVI%), 77, 2)
                Call REPLA(y$, MKD$(SALDEU#), 79, 8)
                Call REPLA(y$, es$, 95, 4)
                '\\\OT-06-0104-WAR-23/03/06///
                LEYMAN1$ = ".........."
                Call REPLA(y$, LEYMAN1$, 99, 10)
                Call REPLA(y$, LEYMAN1$, 109, 10)
                Call REPLA(y$, LEYMAN1$, 119, 10)
                '\\\OT-06-0104-WAR-FIN///
                UJ& = UJ& + 1
                Call GRAREG("FAPEPAGX", y$, UJ&)
                
                If MODOMONEI% = 2 Then
                 Z$ = REGBLAN$("FAPEPAGD")
                 Call REPLA(Z$, Left$(y$, 94), 1, 94)
                 Call REPLA(Z$, Mid$(y$, 95, 34), 115, 34)
                 Call REPLA(Z$, SIMBOLPES$(MEMI%), 95, 4)
                 Call REPLA(Z$, MKD$(SALDEUDOL#), 99, 8)
                 If MEMI% > 1 Then
                    TIPOCAM# = TICAMISORP(NFS$, FF%, NC%)
                    Call REPLA(Z$, MKD$(TIPOCAM#), 149, 8)
                 End If
                 Call GRAREG("FAPEPAGD", Z$, UJ&)
               '//OT-07-0408
              End If

              End If
            End If
70      .MoveNext
        Loop
        '
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Rango: Cuentas " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)) + " ... Vencim.Hasta: " + Text6)
    '
    ' ORDENA POR numero de comprobante
    Call FILESORT("FAPEPAGX", "", 4, 4, "ASC")
    ' ORDENA POR FECHA DE VENCIMIENTO
    Call FILESORT("FAPEPAGX", "", 7, 7, "ASC")
    ' AHORA ORDENA POR MODO SELECCIONADO
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPEPAGX", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPEPAGX", "", 2, 2, "ASC")
    End Select
    '
    If MODOMONEI% = 2 Then
       Call SETULTREG("FAPEPAGD", UJ&)
       Call CIERRARCH("FAPEPAGD")
       '
       Call HEADERS("FAPEPAGD", "")
       Call HEADERS("FAPEPAGD", "Rango: Cuentas " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)) + " ... Vencim.Hasta: " + Text6)
       '
       ' ORDENA POR numero de comprobante
       Call FILESORT("FAPEPAGD", "", 4, 4, "ASC")
       ' ORDENA POR FECHA DE VENCIMIENTO
       Call FILESORT("FAPEPAGD", "", 7, 7, "ASC")
       ' AHORA ORDENA POR MODO SELECCIONADO
       Select Case MODORDE%
         Case 1
           Call FILESORT("FAPEPAGD", "", 1, 1, "ASC")
         Case 2
           Call FILESORT("FAPEPAGD", "", 2, 2, "ASC")
       End Select
    End If

    Screen.MousePointer = 1
    '
End Sub
'

Sub RECONSTRUYE_DEUDA_DE_PROVEEDORES_CON_DIFERENCIA()
    '
    HOXX = HOY(HO$)
    Dim DEUDA#(32767), SALDO#(32767)
    '

    Call APERBASDAT("CABAN")
    CONDI$ = " TipoMovim = 'FCOM' "
    CONDI$ = CONDI$ + " or TipoMovim = 'OPAG' "
    CONDI$ = CONDI$ + " or TipoMovim = 'AJUP' "
    FIN& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE " & CONDI$
    Set CABATEMP = READSET(SQLX$)
     With CABATEMP
        Do While Not .EOF    '
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        NC% = XVALO(!NUPROVE)
        If NC% > 0 Then
            IFAC# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
        End If
       .MoveNext
      Loop
    End With
    '
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    CONDI$ = "Status < 3 "
    FIN& = CantRegistros&("SACREPEN", CONDI$, "NOMESS")
    '
    SQLX$ = "SELECT * FROM SACREPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE " & CONDI$
    Set SACTEMP = READSET(SQLX$)
    With SACTEMP
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)

          NC% = XVALO(!NUPROVE)
          SALPAR# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
          End If
        .MoveNext
        Loop
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    For i% = 1 To 32767
      If TICUEPRO%((-1) * i%) = 0 Then  ' SOLO CUENTAS CORRIENTES
        If Abs(SALDO#(i%)) >= 0.005 Or Abs(DEUDA#(i%)) >= 0.005 Then
          If Abs(SALDO#(i%) - DEUDA#(i%)) > 0.005 Then
            '
            Call RAPLIPRO(i%)
            '
            CONDI2$ = " Status < 3  AND NuProve = " & i%
            SADE# = XVALO(VALOBADA("SACREPEN", "Sum(salddebe)", CONDI2$, "NOMESS"))
            SACRE# = XVALO(VALOBADA("SACREPEN", "Sum(saldacre)", CONDI2$, "NOMESS"))
            DEUDAX# = SACRE# - SADE#
            '
            If Abs(SALDO#(i%) - DEUDAX#) > 0.005 Then
               Call COMUNI("No se ha Podido Recuperar la Diferencia (Saldo Y Deuda) Del Proveedor N° : " & i% & "\" & rasocli$((-1) * i%) & "\ Consulte a su Soporte de Sistemas")
            End If
            '
          End If
          '
        End If
      End If
      If i% = 32767 Then Exit For
    Next i%
    
    

End Sub



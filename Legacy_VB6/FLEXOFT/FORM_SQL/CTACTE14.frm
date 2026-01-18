VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CTACTE14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente de Clientes"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Detalle de Cuenta Corriente"
      Top             =   2100
      Width           =   10935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Honorario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   10935
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8550
         Begin VB.CommandButton Command1 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            ToolTipText     =   "Seleccionar Rango de Fechas"
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TXTHASTA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6675
            Locked          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Doble Click Para Limpiar Casillero"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Hasta: "
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TXTDESDE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6675
            Locked          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Doble Click Para Limpiar Casillero"
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Desde: "
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Sel. Fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   14
            ToolTipText     =   "Seleccionar Rango de Fechas"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   7680
            Picture         =   "CTACTE14.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Imprime Muestra del Formulario Activo"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   720
         End
         Begin VB.TextBox TXTID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   24
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "ID"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TXTNOMBRE 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   24
            TabIndex        =   0
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Cta."
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "Nombre"
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   750
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   195
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "CTACTE14.frx":066A
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CTACTE14.frx":06C8
      Top             =   120
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5235
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9234
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16501405
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "CTACTE14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
 

Private Sub Command1_Click()
  Call CTACTE(0)
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
    OBX$ = ""
    If TRIM$(TXTDESDE) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTDESDE)), 1, 2)
    End If
    If TRIM$(TXTHASTA) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTHASTA)), 3, 2)
    End If
    VDEF$ = OBX$
10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        VDEF$ = OBX$
        GoTo 10
    End If
    '
    If DES% < 1 Or HAS% < 1 Then
       Call COMUNI("Fechas Inexistentes o No Válidas")
       GoTo 99
    End If
    
    
    DESDE = FECHATEX$(DES%)
    HASTA = FECHATEX$(HAS%)
    '
    TXTDESDE = "" ' PARA QUE PROVOQUE EL EVENTO CHANGE DE ACTUALIZAR
    TXTHASTA.TEXT = HASTA
    TXTDESDE.TEXT = DESDE
    '
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   NC% = XVALO(Text4)
   If MSF1.Rows <= 1 Then
     Call COMUNI("Falta Detalle de Listado")
     GoTo 99
   End If
   '
   Call GENERATABLAIMPRE
   CONDI$ = "CAMPO1 = '' OR CAMPO1 <> '' OR CAMPO1 IS NULL AND CodiEmpr = " & CodiEmp%
   Call BORRAREGISTROS("TABLA_IMPRE", CONDI$, REGAF&, "NOMESS")
   '
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
25 On Error Resume Next
   SQLX$ = "SELECT *  FROM TABLA_IMPRE  WHERE CAMPO1 =  ''  "
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   
'   Call CARGA_ENCABEZADO(CTACTE.MSF1, "Documento/A12;Fecha/D8;Concepto;Debe/F12.2;Haber/F12.2;Saldo/F12.2;", CTACTE14)
   For I& = 1 To MSF1.Rows - 1
      DOCUMENTO$ = MSF1.TextMatrix(I&, 1)
      FECHA$ = MSF1.TextMatrix(I&, 2)
      CONCEPTO$ = MSF1.TextMatrix(I&, 3)
      DEBEX$ = MSF1.TextMatrix(I&, 4)
      HABEX$ = MSF1.TextMatrix(I&, 5)
      SALDOX$ = MSF1.TextMatrix(I&, 6)
      With RS1
        .AddNew
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !CodiEmpr = CodiEmp%
        !CAMPO1 = TRIM$(DOCUMENTO$)
        !CAMPO2 = TRIM$(FECHA$)
        !CAMPO3 = Left$(TRIM$(CONCEPTO$), 48)
        !CAMPO4 = TRIM$(DEBEX$)
        !CAMPO5 = TRIM$(HABEX$)
        !CAMPO6 = TRIM$(SALDOX$)
        .Update
      End With
   Next I&
   '
   SALDOANTERIOR$ = "00.00"
   DES% = FECHANUM(TXTDESDE)
   HAS% = FECHANUM(TXTHASTA)
   If CVI(FECHANUM(TXTDESDE)) > 0 And CVI(FECHANUM(TXTHASTA)) > 0 Then
     SALDOANTERIOR$ = FORMATNUM$(SALDOAFECHA#(NC%, DES%, HAS%), "F12.2")
   End If
   SALDOACTUAL$ = FORMATNUM$(SALDOAFECHA#(NC%), "F12.2")

   FEDES% = FECHANUM(TRIM$(TXTDESDE))
   FEHAS% = FECHANUM(TRIM$(TXTHASTA))
   FILTX$ = TRIM$(Str$(NC%)) & ";" & TRIM$(Str$(FEDES%)) & ";" & TRIM$(Str$(FEHAS%)) & ";" & TRIM$(SALDOACTUAL$)
    
   Call STDFORM("CTCT-GSM", FILTX$)

99 Command4.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text4.TEXT = TRIM$(Str$(NCLIE%))
    End If
    Command6.Enabled = True

End Sub

Private Sub Form_Load()
'Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
   
    Call APLICACIONSKINS(Me)
    '
    Call CARGA_ENCABEZADO(CTACTE14.MSF1, "Documento/A12;Fecha/D8;Concepto/A24;Debe/F12.2;Haber/F12.2;Saldo/F12.2;", CTACTE14)
    '
    
End Sub

Private Sub Text4_Change()
    TXTID = Text4
End Sub

Private Sub TXTDESDE_Change()
   MSF1.Rows = 1
   If TRIM$(TXTDESDE) <> "" Then
     Call TXTID_Change
   ElseIf TRIM$(TXTDESDE) = "" And TRIM$(TXTHASTA) <> "" Then
      Call TXTID_Change
   End If
End Sub

Private Sub TXTDESDE_DblClick()
   TXTDESDE = ""
   
End Sub


Private Sub TXTHASTA_DblClick()
    TXTHASTA = ""
End Sub


Private Sub TXTID_Change()
   Call BLANFORMU
   MSF1.Rows = 1
   NC% = XVALO(TXTID)
   If NC% < 1 Then Exit Sub
   '
   Call CTACTE(NC%)
End Sub
Sub CTACTE(Optional NC%)
   TXTNOMBRE = RASOCLI$(NC%)
'   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
'   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
   DES% = FECHANUM(TXTDESDE)
   HAS% = FECHANUM(TXTHASTA)
   MSF1.Rows = 1
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE Status = 0 "
      If NC% > 0 Then SQLX$ = SQLX$ + " And NUCLIEN = " & NC%
      SQLX$ = SQLX$ + " and (tipomovim = 'FHON' OR tipomovim = 'RHON' OR tipomovim = 'AJUD')"
      
      'FILTRO DE FECHAS
      SALFE# = 0
      If DES% > 0 And HAS% > 0 Then
         SQLX$ = SQLX$ + " and FECHEMISI >= '" & FETEXCOR1(DES%) & "' AND FECHEMISI <= '" & FETEXCOR1$(HAS%) & "'"
         SALFE# = SALDOAFECHA#(NC%, DES%, HAS%)
         I& = MSF1.Rows
        
         MSF1.Rows = MSF1.Rows + 1
            MSF1.TextMatrix(I&, 1) = "Transporte"
            MSF1.TextMatrix(I&, 2) = TXTDESDE
            MSF1.TextMatrix(I&, 3) = "Saldo Acumulado"
            HABEX1$ = "0"
            DEBEX1$ = "0"
            If SALFE# > 0 Then
               DEBEX1$ = FORMATNUM$(SALFE#, "F12.2")
            Else
               HABEX1$ = FORMATNUM$(SALFE#, "F12.2")
            End If
            MSF1.TextMatrix(I&, 4) = DEBEX1$
            MSF1.TextMatrix(I&, 5) = HABEX1$
            ACUMUMUSALDO# = Abs(XVALO(DEBEX1$) - XVALO(HABEX1$))
            MSF1.TextMatrix(I&, 6) = FORMATNUM$(ACUMUMUSALDO#, "F12.2")
      End If
      
      SQLX$ = SQLX$ + " ORDER BY FeHoReal"
      Set RS = READSET(SQLX$)
      VUELTA% = 0
      With RS
        If Not .EOF Then
          Do While Not .EOF
            I& = MSF1.Rows
            MSF1.Rows = I& + 1

            NC% = XVALO(!NUCLIEN)
            NURECI$ = SAFETEXT$(!CODICOM_COR)
            CONCEPTO$ = SAFETEXT$(!TEXTOCAE)
            FE% = CVINT(!FECHEMISI)
            DEBEX1$ = FORMATNUM$(XVALO(!IDEBECOMP), "F12.2")
            HABEX1$ = FORMATNUM$(XVALO(!IHABECOMP), "F12.2")
            SALDO1# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
            ACUMUMUSALDO# = ACUMUMUSALDO# + SALDO1#
            '
            MSF1.TextMatrix(I&, 1) = NURECI$
            MSF1.TextMatrix(I&, 2) = FECHATEX$(FE%)
            MSF1.TextMatrix(I&, 3) = CONCEPTO$
            MSF1.TextMatrix(I&, 4) = DEBEX1$
            MSF1.TextMatrix(I&, 5) = HABEX1$
            MSF1.TextMatrix(I&, 6) = FORMATNUM$(ACUMUMUSALDO#, "F12.2")
            .MoveNext
          Loop
        End If
      End With
      RS.Close
      Set RS = Nothing
      '
   
   End Sub

Function SALDOAFECHA#(NC%, Optional DES%, Optional HAS%)
   'CONSULTA EL SALDO DEL CLIENTE
   'SI TIENE RANGO DE FECHAS FILTRA LA SUMA DEL HABER Y DEL DEBE SEGUN FECHA
   'SI NO TIENE FECHA NO CONSIDERA RANGO DE FECHAS
   
''''   CONDI$ = "NUCLIEN = " & NC% & " AND TIPOMOVIM = 'FHON' AND STATUS < 9 "
''''   If DES% > 0 And HAS% > 0 Then
''''      DESDE = FETEXCOR1$(DES%): HASTA = FETEXCOR1$(HAS%)
''''      CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDE & "' AND FECHEMISI <= '" & HASTA & "'"
''''   End If
''''   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
''''   CONDI$ = "NUCLIEN = " & NC% & " AND TIPOMOVIM = 'RHON' AND STATUS < 9 "
''''   If DES% > 0 And HAS% > 0 Then
''''      CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDE & "' AND FECHEMISI <= '" & HASTA & "'"
''''   End If
''''   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
''''
''''   SALDOX# = IDEBEX# - IHABEX#
''''   SALDOAFECHA# = SALDOX#
   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'FHON' OR TIPOMOVIM = 'AJUD' OR TIPOMOVIM = 'RHON') AND STATUS < 9"
   If DES% > 0 And HAS% > 0 Then
      DESDE = FETEXCOR1$(DES%): HASTA = FETEXCOR1$(HAS%)
      CONDI$ = CONDI$ + " AND FECHEMISI < '" & DESDE & "'" ' AND FECHEMISI <= '" & HASTA & "'"
   End If
   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
   SALDOX# = IDEBEX# - IHABEX#
   SALDOAFECHA# = SALDOX#
   
End Function


Sub BLANFORMU()
   CTACTE14.TXTNOMBRE = ""
End Sub


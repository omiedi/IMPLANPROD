VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELNUMSEAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Números de Serie "
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7755
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LISNUM 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde para seleccionar"
      Top             =   960
      Width           =   6465
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   6405
      TabIndex        =   7
      Top             =   600
      Width           =   6465
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Serie          Descripción"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   180
         Left            =   225
         TabIndex        =   8
         Top             =   0
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   5775
      Left            =   6840
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
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
         Left            =   120
         Picture         =   "SELNUMSEAR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
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
         Height          =   650
         Left            =   120
         Picture         =   "SELNUMSEAR.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4425
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   120
         Picture         =   "SELNUMSEAR.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   3555
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELNUMSEAR.frx":059E
         Top             =   4770
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   6475
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
         Height          =   285
         Left            =   2940
         TabIndex        =   15
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton Command3 
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
         Left            =   6060
         TabIndex        =   14
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Consumo: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   21
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep.Entrada: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   20
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5760
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "SELNUMSEAR.frx":24C0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELNUMSEAR.frx":26F4
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº.O.Fabricación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label34 
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
      Left            =   5190
      TabIndex        =   11
      Top             =   120
      Width           =   1395
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5500
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line2 
      X1              =   5505
      X2              =   5505
      Y1              =   495
      Y2              =   3245
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   9
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELNUMSEAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'


Private Sub BOTREC_Click()
     BOTREC.Enabled = False
     '
     ' VERIFICO QUE HAYA AL MENOS UNO SELECCIONADO
     CANTIEQUI# = 0
     For II2% = 1 To SELNUMSEAR.LISNUM.ListCount
        If SELNUMSEAR.LISNUM.Selected(II2% - 1) = True Then CANTIEQUI# = CANTIEQUI# + 1
     Next II2%
     If CANTIEQUI# < 0.5 Then
        Call MENSERR(24, " Debe Seleccionar los Números de Serie ")
        GoTo 99
    End If
    '
'    'VALIDA QUE NO CAMBIE EL DEPOSITO EN EL CASO QUE SEA UN ARMADOR EXTERNO.
'    NORFA$ = TRIM$(Label34)
'    If ValidaArmador%(NORFA$) > 0 Then
'       DEPARMA% = XVALO(CONTROL("", "DEPOARMA"))
'       If XVALO(Label3) <> DEPARMA% Then
'        Call MENSERR(24, "No Válido Cambiar de Depósito\Orden de Fabricación Asignado a Armador Externo\Depósito de Consumo: " & DEPARMA%)
'        GoTo 99
'       End If
'    End If
    '
    
    Screen.MousePointer = 11
    NORFA$ = Label34
    CI1% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS"))
    VACONTROL% = 1
    CODICONJ$ = CODEXT$(CI1%)
'    Call CARGA_ENCABEZADO(COMPONEN.MSF, "N.Serie/F10.0;Código/A18;Descripción/A32;Lote/A14;Nec./F10.0;Uso.U./F10.0", Me)
'50  Call CARCOMPO(CODICONJ$, NORFA$, CANTIEQUI#, VACONTROL%, 3)
    '
    'RECORRO EL LIST Y ARMO LA CONDICION
    For II2% = 1 To SELNUMSEAR.LISNUM.ListCount
        If SELNUMSEAR.LISNUM.Selected(II2% - 1) = True Then
           NROSERIE$ = TRIM$(Mid$(SELNUMSEAR.LISNUM.List(II2% - 1), 1, 16))
           CONDICION$ = CONDICION$ & " OR NUMEROSERIE = '" & NROSERIE$ & "'"
           CANTIEQUI# = CANTIEQUI# + 1
        End If
    Next II2%
    '
    'LE QUITO EL PRIMER OR Y LE AGREGO LOS PARENTESIS
    CONDICION$ = Mid$(CONDICION$, 5)
    CONDICION$ = "(" & CONDICION$
    CONDICION$ = CONDICION$ & ")"
    
    SQLX$ = "SELECT NUMEROSERIE,COD_EXTE,DESCRIP,CANTRES,USOUNIT, FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ = " WHERE IDSUBOR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ = " AND " & CONDICION$
    SQLX$ = SQLX$ = " ORDER BY NUMEROSERIE,COD_EXTE"
    Set RS = READSET(SQLX$)
    '
    With RS
        Do While Not .EOF
          i& = i& + 1
          MSF.Rows = i& + 1
          COMPONEN.MSF.TextMatrix(i&, 1) = SAFETEXT$(!NUMEROSERIE)
             CODI$ = SAFETEXT$(!Cod_Exte)
             LOTEX$ = TRIM$(SAFETEXT$(!idenlote))
             CANTIRESERVADA = XVALO(!CANTRES)
          COMPONEN.MSF.TextMatrix(i&, 2) = CODI$
          COMPONEN.MSF.TextMatrix(i&, 3) = SAFETEXT$(!Descrip)
          COMPONEN.MSF.TextMatrix(i&, 4) = SAFETEXT$(!idenlote)
          COMPONEN.MSF.TextMatrix(i&, 5) = FORMATNUM$(CANTIRESERVADA, "F10." & XNTDC)
          COMPONEN.MSF.TextMatrix(i&, 6) = FORMATNUM$(XVALO(!UsoUnit), "F10." & XNTDC)
             '
             'TRAE EL STOCK DEL CODIGO
             CI11% = CODINT%(Code$)
             If ESTRAZABLE%(CI11%) >= 1 And LOTEX$ <> "" Then
                 STOXX = SaldoLote(CI11%, LOTEX$)
             Else
                 STOXX = STOCKDIS(CI11%)
             End If
             '
             If STOXX < FRACMIN(CI11%) Then STOXX = 0
          COMPONEN.MSF.TextMatrix(i&, 7) = FORMATNUM$(XVALO(!STOXX), "F10." & XNTDC)
             FALTAX = CANTIRESERVADA - STOXX 'NECESITO 6 MENOS TENGO 4 =FALTAN 2
             If FALTAX >= FRACMIN(CI11%) Then
               HAYFALTANTES% = 1
             Else
               FALTAX = 0
             End If
          COMPONEN.MSF.TextMatrix(i&, 8) = FORMATNUM$(FALTAX, "F10." & XNTDC)
          .MoveNext
        Loop
    End With
    '
50  COMPONEN.Show 1
    If COMPONEN.OKK% <= 0 Then GoTo 99                  ' SI CANCELA SE VA
    If HAYFALTANTES% > 0 Then
       OPX% = COMALTER%("Atención: Esta Orden de Fabricación Contiene Componentes con Faltantes\Cancelar Cierre\Continuar")
       If OPX% <> 2 Then ' SI NO TRANSFIERE
          GoTo 50
       End If

       GoTo 50
    End If
    '
'    'VALIDO QUE TODOS LOS COMPONENTES TENGAN LOTE
'    For I& = 1 To COMPONEN.MSF.Rows - 1
'       LOTE$ = TRIM$(COMPONEN.MSF.TextMatrix(I&, 4))
'       CODI$ = TRIM$(COMPONEN.MSF.TextMatrix(I&, 2))
'       CI11% = CODINT%(CODI$)
'       If ESTRAZABLE%(CI11%) >= 1 Then 'VALIDO QUE LOS TRAZABLES TENGAN LOTE
'        If LOTE$ = "" Then
'           CODI$ = TRIM$(COMPONEN.MSF.TextMatrix(I&, 2))
'           NSER$ = TRIM$(COMPONEN.MSF.TextMatrix(I&, 1))
'           Call COMUNI("Imposible Continuar !!!!\Falta Lote Para Código: " & CODI$ & "\Número de Serie: " & NSER$)
'           GoTo 50
'        End If
'      End If
'    Next I&
    '
    
    HOII% = HOY(HOS$)
    FECHA% = HOII%
    CI2% = CI1%
    CMOV$ = "OF"
    DOPRI$ = NORFA$
    DOSEC$ = "IF-" + TRIM$(Mid$(NORFA$, 4))
    REFE$ = "Fabricación Propia"
    VUNI# = 0
    DEPX% = XVALO(SELNUMSEAR.Label1)   ' DEPOSITO DE DESTINO
    DP2% = XVALO(SELNUMSEAR.Label3)    ' DEPOSITO DE CONSUMO
    '
    
    'GRABA EN LA TABLA DE APROBACION - SOLO PRODUCTO TERMINADO.
    For II2% = 1 To SELNUMSEAR.LISNUM.ListCount
        If SELNUMSEAR.LISNUM.Selected(II2% - 1) = True Then
           NROSERIE$ = TRIM$(Mid$(SELNUMSEAR.LISNUM.List(II2% - 1), 1, 16))
           LOTE$ = NROSERIE$
           'GRABA EN TABLA DE APROBACION
           Call MOVISTOCK_APROF(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, "PESOS", NC%, DEPX%, CAN, DP2%)
           'ESCRIBE EN EL ANOTADOR DE OF
           
           TXYZ$ = Space$(93)
           Call REPLA(TTXYZ$, HOS$, 1, 8)
           Call REPLA(TTXYZ$, "Generación de Cierre N.S.: " & NROSERIE$, 11, 56)
           Call REPLA(TTXYZ$, USERTER$, 69, 24)
           TTXYZ$ = TTXYZ$ & vbCrLf
           Call REPLA(TTXYZ$, "Pendiente de Aprobación", 11, 56)
           TTXYZ$ = TTXYZ$ & vbCrLf
           Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
           Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
         End If
    Next II2%
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("ECORDEP")
    Call CIERRARCH("*.*")
    '
    Call COMUNI("Orden de Fabricacion: " & NORFA$ & " Cerrada Por " & CANTIEQUI# & " Equipo/s\(Pendiente de Aprobación)")
    Label34 = ""
    Hide
99  Screen.MousePointer = 1
    BOTREC.Enabled = True
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LISNUM.TopIndex
   For U& = 1 To LISNUM.ListCount
     If LISNUM.Selected(U& - 1) = True Then
         LISNUM.Selected(U& - 1) = False
       Else
         LISNUM.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LISNUM.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   LISNUM.Clear
   Unload Me
End Sub


Private Sub Command3_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label3 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
  

End Sub

Private Sub Command6_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label1 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Form_Load()
     'Actualizacion de Depositos de consumo donde el mismo este vacio con movimiento 'cf'
     Screen.MousePointer = 11
     SELNUMSEAR.Visible = True
     UTILIZA_SKIN% = 1
     Call APLICACIONSKINS(Me)

     DepoDes% = XVALO(Control("", "DEPODEST"))
     CONDI$ = "CODIMOVI = 'CF' AND CANTSALID > 0 AND (DEPOMOVI < 1 OR DEPOMOVI IS NULL)"
     Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", CONDI$, DepoDes%, REGAF&, "NOMESS")
     Screen.MousePointer = 1
     '
End Sub

Private Sub label1_Change()
   Label8 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label1)))
End Sub

Private Sub Label3_Change()
   Label2 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label3)))
End Sub

Private Sub Label34_Change()
     '
     If Label34 = "" Then
        Label1 = "": Label3 = ""
        LISNUM.Clear
        Exit Sub
     End If
     '
     NORFA$ = Label34
     CI% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
     CONDI$ = "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI% & " AND DEPONUMSE < 0 "
     Call CARGALISTA(SELNUMSEAR.LISNUM, "TANUMSE", "NUMEROSERIE/A16; Descrip/A33", CONDI$)
     '
     'CARGA DEPOSITOS SEGUN CONFIGURACION O SI TIENE ASIGNADO ARMADOR EXTERNO
     SELNUMSEAR.Label3 = Control("", "DEPODEST")
     '
     For II2% = 1 To SELNUMSE.LISNUM.ListCount
        SELNUMSEAR.LISNUM.Selected(II2% - 1) = True
     Next II2%
     '
End Sub


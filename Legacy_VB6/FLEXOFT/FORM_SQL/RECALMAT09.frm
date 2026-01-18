VERSION 5.00
Begin VB.Form RECALMAT09 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo Necesidades de Materiales"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "O.K."
      Height          =   1050
      Left            =   6840
      Picture         =   "RECALMAT09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Inicia Proceso de Cálculo"
      Top             =   2625
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   210
      TabIndex        =   17
      Top             =   2625
      Width           =   6450
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1365
      TabIndex        =   10
      Top             =   210
      Width           =   540
   End
   Begin VB.CommandButton Command9 
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
      Left            =   1905
      TabIndex        =   9
      Top             =   210
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6825
      ScaleHeight     =   2355
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   -50
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   19
         Top             =   1575
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   -135
            TabIndex        =   20
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   120
         Picture         =   "RECALMAT09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RECALMAT09.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1890
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Avance de Proceso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   21
      Top             =   840
      Width           =   1785
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6825
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   16
      Top             =   225
      Width           =   1020
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   510
      Left            =   4410
      TabIndex        =   15
      Top             =   255
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   300
      Left            =   3000
      TabIndex        =   14
      Top             =   1710
      Width           =   1785
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4875
      TabIndex        =   13
      Top             =   1635
      Width           =   1800
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      Height          =   270
      Left            =   300
      TabIndex        =   12
      Top             =   255
      Width           =   915
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2145
      TabIndex        =   11
      Top             =   225
      Width           =   2205
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   7
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   300
      Left            =   105
      TabIndex        =   6
      Top             =   2100
      Width           =   1125
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   5
      Top             =   2025
      Width           =   5310
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   4
      Top             =   1635
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   300
      Left            =   -315
      TabIndex        =   3
      Top             =   1710
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "S/Fabr. Nº:"
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   1365
      Width           =   1125
   End
End
Attribute VB_Name = "RECALMAT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then
      If COMALTER%("¿ Realmente Cancela la Asignación\Automática de Lotes Propuesta ?\\No, Conservarla\Si, Cancelarla") = 2 Then GoTo 10
      Exit Sub
   End If
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Label28_Change()
   CIXI% = CODINT%(Label28)
End Sub

Private Sub Text6_Change()
   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    List1.Clear
    Dim ORSEL As ADODB.Recordset
    '
    If TRIM$(Label38) = "" Then
       Call MENSERR(24, "Identifique el Depósito de Asignacion de Consumos")
       GoTo 99
    End If
    '
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    If XVALO(Text6) = DEPOTRANS% Then
        COMUNI ("Imposible Realizar Cálculo\Deposito de Material en Transito ")
       GoTo 99
    End If

    Screen.MousePointer = 11
    Call ABRESTOCKS
    List1.AddItem "Asignando Materiales Existentes en Depósito "
5   CONDI$ = "IDSUBOR <> '' "
    CONDI$ = CONDI$ + " AND STATUORF < 3 "
    CONDI$ = CONDI$ + " AND CanApro < CanProy "
    CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
    FIN& = CantRegistros("ORDEFABR", CONDI$)
    JJ& = 0
    DepOrig% = XVALO(Text6)
    If FIN& > 0 Then
       Set ORSEL = New ADODB.Recordset
       SQLX$ = "SELECT IDSUBOR FROM ORDEFABR WHERE " & CONDI$ & " ORDER BY IDSUBOR ASC"

       Set ORSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With ORSEL
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           NORFA$ = !IdSubOr
           If NORFA$ <> "" Then
              Label32 = NORFA$
              NULIPI& = 0
              Call REEMPLAZACOMPRAS(NORFA$, DepOrig%)
              Call ASIGNAREQUERI(NORFA$, DepOrig%, NULIPI&)
           End If
         .MoveNext
         Loop
       End With
    End If
    '
    '   CALCULO DE NECESIDADES DE COMPRAS
    List1.AddItem "Calculando Necesidades de Compras"
    Call CALNECOM
    '
    List1.AddItem "Calculando Notas de Pedido a Emitir"
25  CONDI$ = "IDSUBOR <> '' "
    CONDI$ = CONDI$ + " AND STATUORF < 3 "
    CONDI$ = CONDI$ + " AND CanApro < CanProy "
    CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
    FIN& = CantRegistros("ORDEFABR", CONDI$)
    JJ& = 0
    If FIN& > 0 Then
       Set ORSEL = New ADODB.Recordset
       SQLX$ = "SELECT IDSUBOR FROM ORDEFABR WHERE " & CONDI$ & " ORDER BY IDSUBOR ASC"
       Set ORSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With ORSEL
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           NORFA$ = !IdSubOr
           If NORFA$ <> "" Then
             CLIDEST& = XVALO(VALOBADA("ORDEFABR", "Nume_Cli", "IdSubOr = '" & NORFA$ & "'"))
             CONDI$ = "Destino = '" & NORFA$ & "'"
             SQLX$ = "SELECT IDENREQUE,Cod_Inte,Sum(Kilo_Reque) as Sureque,Sum(Kilo_Asig) as Suasig FROM BALAREQMO WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE " & CONDI$
             SQLX$ = SQLX$ + " GROUP BY IDENREQUE,COD_INTE"
             '
             Dim BALAREQ As ADODB.Recordset
             Set BALAREQ = FLEXCONN.Execute(FILTSQL$(SQLX$))
             With BALAREQ
               Do While Not .EOF
                 CIXI% = XVALO(!Cod_Inte)
                 If CRITREP%(CIXI%) = 1 Then
                   If XVALO(!SUREQUE) > XVALO(!SUASIG) Then
                     '
                     IDREQ$ = SAFETEXT$(!IdenReque)
                     KILAPE = FORMATNUM$(XVALO(!SUREQUE) - XVALO(!SUASIG), "F16.6")
                     CONDI$ = "IdenReque = '" & IDREQ$ & "' AND KILO_REQUE > 0"
                     Call ACTUREGISTRO("BALAREQMO", "Nume_Cli", CONDI$, CLIDEST&, RAFE&)
                     Call ACTUREGISTRO("BALAREQMO", "Kilo_Aped", CONDI$, KILAPE, RAFE&)
                     '
                   End If
                 End If
               .MoveNext
               Loop
             End With
             BALAREQ.Close
             Set BALAREQ = Nothing
           End If
         .MoveNext
         Loop
       End With
       ORSEL.Close
       Set ORSEL = Nothing
    End If
    '
    Call COMUNI("Calculo Completo")
    '
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
End Sub
'
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       'Frame9.Visible = True
    End If
    '
    Call FRESHALL
    '
    If Left$(APLINVO$, 7) = "OFABRIC" Then
       Label32 = TRIM$(Mid$(APLINVO$, 8))
       If Label32 <> "" Then
          If ECOARCH$("ECORDEP", Label32) <> "" Then
            HOII% = HOY(HOS$)
            Label24 = HOS$
          End If
       End If
    End If
    '
End Sub

Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
    HOII% = HOY(HOS$)
    Label24 = HOS$
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Label28 = ""
    Label30 = ""
    Label34 = ""
    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
         '
         Dim ORDEFABTMP As ADODB.Recordset
         Set ORDEFABTMP = New ADODB.Recordset
         Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With ORDEFABTMP
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Exit Sub
           End If
           On Error GoTo 0
           CIXI% = !Cod_Inte
           Label28 = CODEXT$(CIXI%)
           Label30 = DESCRIT0$(CIXI%)
           Label34 = FORMATNUM$(!CanProy, "F13." + CNTDC$(CIXI%))
           '
         End With
         '
    End If
99  '
End Sub

Private Sub Label32_DblClick()
    ' USAR PARA LIMPIAR TODO
    Label32 = ""
End Sub
'
Sub LIMPLAB()
    
    Label6.Caption = ""
    Text4.TEXT = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text12 = ""
  
End Sub

Sub CALNECOM()
   ' leer las ordenes de compra en curso por el saldo pendiente de recepcion
   Call APERBASDAT("COMPRAS")
   '
   Screen.MousePointer = 11
   CONDI$ = "Nume_Ocom > 0 "
   CONDI$ = CONDI$ + " AND Stat_Ocom < 8 "
   CONDI$ = CONDI$ + " AND Stat_Ocom <> -1 "
   CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   FIN& = CantRegistros("OCOMDETA", CONDI$)
   '
   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) where " + CONDI$
   Dim OCOPEN As ADODB.Recordset
   Set OCOPEN = New ADODB.Recordset
   Set OCOPEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JI& = 0
   With OCOPEN
      Do While Not .EOF
         JI& = JI& + 1
         Call TRACE(20, JI&, FIN&)
         IDOC$ = "OC." + TRIM$(Str$(XVALO(!Nume_Ocom)))
         CAPENOC = XVALO(!CANT_OCOM) - XVALO(!Cant_Rec)
         CIXI% = XVALO(!Cod_Inte)
         If CAPENOC < 0 Then CAPENOC = 0
         '
         ' descontar de las ordenes  de compra lo ya asignado en balareqmo
         CANYASI = XVALO(VALOBADA("BALAREQMO", "SUM(Kilo_Asig)", "IdenLote = '" & IDOC$ & "' AND Cod_Inte = " & CIXI%))
         CAPENOC = CAPENOC - CANYASI
         ' AQÍ PONER CAPTURA DE SOBRE-ASIGNACION
         If CAPENOC < 0 Then
            CAPENOC = 0
            Call MENSERR(24, "Solicitud de Suministro '" + IDOC$ + "' - Código '" + CODEXT$(CIXI%) + "' está Sobre-Asignada.")
         End If
         '
         If CAPENOC > 0 Then
            'Se agrega Registro en Balareqmo Informando el Pedido Generado
            Dim REQUERI As ADODB.Recordset
            Set REQUERI = New ADODB.Recordset
            SQLX$ = "SELECT * FROM BALAREQMO WITH(NOLOCK) "
            SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
            SQLX$ = SQLX$ + " AND KILO_REQUE > 0"
            Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
            '
            With REQUERI
               Do While Not .EOF
                 IDREQ$ = SAFETEXT$(!IdenReque)
                 '
                 ' BORRA REGISTROS DUPLICADOS
                 CONDI2$ = "IdenReque = '" & IDREQ$ & "' AND KILO_REQUE > 0"
                 If CantRegistros("BALAREQMO", CONDI2$) > 1 Then
                    If COMALTER%("Registros Duplicados para Requerimiento '" & IDREQ$ & "'\\Continuar\Depurar") = 2 Then
                       REQUERI.Close
                       Set REQUERI = Nothing
                       Call ENUMERA("BALAREQMO", "NorItReq", CONDI2$)
                       Call BORRAREGISTROS("BALAREQMO", CONDI2$ + " AND NorItReq > 1", RAFE&)
                       Call COMUNI("Depuración Completada.\Vuelva a Lanzar el Cálculo")
                       GoTo 99
                    End If
                 End If
                 '
                 TOASI = VALOBADA("BALAREQMO", "SUM(KILO_ASIG)", "IDENREQUE = '" & IDREQ$ & "'")
                 SALDOL = XVALO(!KILO_REQUE) - TOASI
                 If XVALO(SALDOL) <= 0 Then SALDOL = 0
                 Call ACTUREGISTRO("BALAREQMO", "KILO_SALDO", "IDENREQUE = '" & IDREQ$ & "' AND KILO_REQUE > 0", SALDOL, RAFE&)
                 '
                 If XVALO(SALDOL) > 0 Then
                   CANASI = SALDOL
                   If CANASI > CAPENOC Then CANASI = CAPENOC
                   '
                   SQLX$ = "SELECT * FROM BALAREQMO "
                   SQLX$ = SQLX$ + " WHERE IDENREQUE = '" & IDREQ$ & "'"
                   Dim BALTEMP As ADODB.Recordset
                   Set BALTEMP = New ADODB.Recordset
                   BALTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockOptimistic
                   '
                   ' asignar en balareqmo el saldo de cada O/Compra
                   NROPED$ = IDOC$
                   BALTEMP.AddNew   ' YA CONVERTIDO
                          On Error Resume Next
                   BALTEMP!CodiEmpr = CodiEmp%
                          On Error GoTo 0
                   BALTEMP!CODICOMPRO = SAFETEXT(!CODICOMPRO)
                   BALTEMP!NUMECOMPRO = XVALO(!NUMECOMPRO)
                   BALTEMP!CODSUCU = XVALO(!CODSUCU)
                   BALTEMP!NUORITEM = XVALO(!NUORITEM)
                   BALTEMP!IdenReque = SAFETEXT(!IdenReque)
                   BALTEMP!Cod_Inte = XVALO(!Cod_Inte)
                   BALTEMP!IdenLote = NROPED$
                   BALTEMP!Kilo_Asig = CANASI ' XVALO(!KILO_SALDO)
                   BALTEMP!FeReMovi = FECHATEX$(HOY(HO$))
                   BALTEMP!fechMov = FECHATEX$(HOY(HO))
                   BALTEMP!CREADOPOR = "FLEX"
                   BALTEMP!Destino = SAFETEXT(!Destino)
                   BALTEMP.Update
                   BALTEMP.Close
                   Set BALTEMP = Nothing
                   '
                   CAPENOC = CAPENOC - CANASI
                   If CAPENOC <= 0 Then GoTo 29
                 End If
               .MoveNext
               Loop
            End With
29          REQUERI.Close
            Set REQUERI = Nothing
         End If
39    .MoveNext
      Loop
   End With
   '
   ' BORRA LA TABLA OCOEMI
   Call BORRAREGISTROS("OCOEMI", "Cod_Inte >= 0", RAFE&, "NOMESS")
   '
   Dim BALARE As ADODB.Recordset
   Set BALARE = New ADODB.Recordset
   SQLX$ = "SELECT Cod_Inte, idenreque, "
   SQLX$ = SQLX$ + "SUM(Kilo_Reque) AS SxKilo_Reque, "
   SQLX$ = SQLX$ + "SUM(Kilo_Asig) AS SxKilo_Asig "
   SQLX$ = SQLX$ + "FROM BALAREQMO WITH(NOLOCK) " ' WHERE Kilo_Reque > Kilo_Asig "
   SQLX$ = SQLX$ + "GROUP BY Cod_Inte, IdenReque "
   '
   Set BALARE = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With BALARE
     Do While Not .EOF
       CIXI% = XVALO(!Cod_Inte)
       If CRITREP%(CIXI%) <> 1 Then  ' NO ES POR PEDIDO
         CODD$ = CODEXT$(CIXI%)
         NASI# = XVALO(!SxKilo_Asig)
         NEBRUTA# = XVALO(!SxKilo_Reque) - NASI#
         If NEBRUTA# > 0 Then
            CONDI$ = "Cod_Inte = " & CIXI%
            If CantRegistros("OCOEMI", CONDI$) < 1 Then
               Call CreaRegistro("OCOEMI", "Cod_Inte", CIXI%)
            End If
            NECOM = NEBRUTA# + XVALO(VALOBADA("OCOEMI", "Cantidad", CONDI$, "NOMESS"))
            Call ACTUREGISTRO("OCOEMI", "Cantidad", CONDI$, NECOM, RAFE)
              FENSOL = FETEXCOR1$(HOY(HO$))
            Call ACTUREGISTRO("OCOEMI", "Fentre_Sol", CONDI$, FENSOL, RAFE)
              NPRX% = PROVHABI%(CIXI%)
            Call ACTUREGISTRO("OCOEMI", "Npro_Ocom", CONDI$, NPRX%, RAFE)
         End If
       End If
     .MoveNext
     Loop
   End With
   BALARE.Close
   Set BALARE = Nothing
   '
   'AQUI APLICAR EL REDON_INDIVIS  '                          NECOM = REDON_INDIVIS(CIXI%, NEBRUTA#)
   SQLX$ = "SELECT * FROM OCOEMI "
   Dim ACTOCLOIND As ADODB.Recordset
   Set ACTOCLOIND = New ADODB.Recordset
   Set ACTOCLOIND = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With ACTOCLOIND
      Do While Not .EOF
        CIXI% = XVALO(!Cod_Inte)
        LOTEIN = REDON_INDIVIS(CIXI%, XVALO(!Cantidad))
        '
        'ACTUALIZA LAS CANTIDADES DE LAS ORDENES DE COMPRA A EMITIR SEGUN EL LOTE INDIVISIBLE
        CONDI$ = "COD_INTE = " & CIXI%
        Call ACTUREGISTRO("OCOEMI", "CANTIDAD", CONDI$, LOTEIN, REGG&, "NOMESS")
        '
      .MoveNext
      Loop
   End With
   '
   ACTOCLOIND.Close
   Set ACTOCLOIND = Nothing
   '
99 Screen.MousePointer = 1

End Sub

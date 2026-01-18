VERSION 5.00
Begin VB.Form MENISO9K 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub-Sistema de Calidad"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sub-Sistema de Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acciones Correctivas / Preventivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   16
         Top             =   3120
         Width           =   4395
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Index           =   8
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   17
            Top             =   400
            Width           =   2430
            Begin VB.CommandButton Command3 
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
               Left            =   0
               Picture         =   "MENISO9K.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   690
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "No Conf. y Acc. Preventivas"
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
               Left            =   735
               TabIndex        =   19
               Top             =   75
               Width           =   1650
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comprobante Interno de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   12
         Top             =   1720
         Width           =   4395
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Index           =   0
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   13
            Top             =   400
            Width           =   2430
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
               Height          =   690
               Left            =   0
               Picture         =   "MENISO9K.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   690
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Comprobante Interno de Entrega"
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
               Left            =   735
               TabIndex        =   15
               Top             =   105
               Width           =   1650
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Calidad O-Fabricación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   8
         Top             =   360
         Width           =   4395
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Index           =   1
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   9
            Top             =   400
            Width           =   2430
            Begin VB.CommandButton Command15 
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
               Left            =   0
               Picture         =   "MENISO9K.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   690
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "O-Fab / Despacho Internos"
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
               Left            =   735
               TabIndex        =   11
               Top             =   100
               Width           =   1650
            End
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6840
      Left            =   5040
      ScaleHeight     =   6780
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton AGenRep 
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
         Left            =   120
         Picture         =   "MENISO9K.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2040
         Width           =   645
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
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
         Left            =   120
         Picture         =   "MENISO9K.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   645
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
         Left            =   120
         Picture         =   "MENISO9K.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   645
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
         Left            =   120
         Picture         =   "MENISO9K.frx":0FF4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   95
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   2
         Top             =   3960
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
         Picture         =   "MENISO9K.frx":12FE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4250
         Width           =   1515
      End
   End
End
Attribute VB_Name = "MENISO9K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CIERRETECNICO()
  
    If DERACCE%("CIENCTEC", "Cierre Técnico de No Conformidad ") < 2 Then Exit Sub
    '
    'PRESENTA CARGALISEL CON NO CONFORMIDADES  EN PROCESO
    Call CARGALISEL("DETANOC", "ACCORREC", "NUMENC/F7;ResponNC/A24;FECHANC/D8", "NUMENC > 0 AND STATUS = 1 ORDER BY ANONC,NumeNC ASC")
    Call SELECHO("DETANOC/NO CONFORMIDADES ")
    NUMNC& = XVALO(VALACT1$("DETANOC"))
    If NUMNC& < 1 Then Exit Sub
    ANO% = Year(Right$(VALACT2$("DETANOC"), 8))

10  OPPX% = COMALTER%("Confirmación de Cierre Técnico\\No Cancelar\Sí, Realizar Cierre Técnico")
    If OPPX% = 2 Then
       FE% = HOY(HO$)
       FECH = FETEXCOR1$(FE%)
       USERNU% = USNBR% Mod 100
       If USERNU% < 1 Then USERNU% = 1
       CONDI$ = "NUMENC = " & NUMNC& & " AND ANONC = " & ANO%
       
       Call ACTUREGISTRO("ACCORREC", "CIERRETEC", CONDI$, FECH, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "USCIERRETEC", CONDI$, USERNU%, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "STATUS", CONDI$, 2, REGAF&, "NOMESS")
    End If
    
End Sub

Sub MODICOMPINT()
    
    'PRESENTA CARGALISEL
    Call CARGALISEL("DETADES", "ENCADESIN", "NRODESPA/F6.0;ORDFAB/A16;FECHDESPA/D8;SECTOREMI/A18;SECTORDESTI/A18", "NRODESPA > 0 ORDER BY NRODESPA ASC")
    Call SELECHO("DETADES/COMPROBANTES INTERNOS ")
    NRODES& = XVALO(VALACT1$("DETADES"))
    If NRODES& < 1 Then Exit Sub
    '
    'TOMO LA INFORMACION DE LA ORDEN
    CONDI$ = " NRODESPA = " & NRODES&
    SectEmi$ = VALOBADA("ENCADESIN", "SECTOREMI", CONDI$, "NOMESS")
    SecDest$ = VALOBADA("ENCADESIN", "SECTORDESTI", CONDI$, "NOMESS")
    NORFA$ = VALOBADA("ENCADESIN", "ORDFAB", CONDI$, "NOMESS")
    CANTITOT = VALOBADA("ENCADESIN", "CANTTOTAL", CONDI$, "NOMESS")
    '
    'PASO LA INFORMACION AL OBJETO
    Call REPLA(VDEF$, MKS$(NRODES&), 1, 4)
    Call REPLA(VDEF$, SectEmi$, 5, 30)
    Call REPLA(VDEF$, SecDest$, 35, 30)
    Call REPLA(VDEF$, NORFA$, 65, 12)
    Call REPLA(VDEF$, MKS$(CANTITOT), 77, 4)
    '
10  OBX$ = OBJPLA$("ALTACOMP", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    'LEO EL OBJETO
    SectEmi$ = TRIM$(Mid$(OBX$, 5, 30))
    SecDest$ = TRIM$(Mid$(OBX$, 35, 30))
    NORFA$ = TRIM$(Mid$(OBX$, 65, 12))
    CANTITOT = CVS(Mid$(OBX$, 77, 4))
    '
    'VALIDACIONES
    If SectEmi$ = "" Then
      Call COMUNI("Falta Ingresar Sector Emisor")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    If SecDest$ = "" Then
      Call COMUNI("Falta Ingresar Sector Destino")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    'CARGAR LOS ITEMS EN LA VENTANA
    SQLX$ = "SELECT * FROM DETADESIN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE  NRODESPA = " & NRODES&
    JL& = 0
    Dim DETATMLP As ADODB.Recordset
    Set DETATMLP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With DETATMLP
      Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         CANTIX# = XVALO(!CANTIDAD)
         LOTE$ = SAFETEXT$(!IDENLOTE)
         If CIXI% > 0 Then
            Call REPLA(XY$, MKI$(CIXI%), 1, 2)
            Call REPLA(XY$, MKD$(CANTIX#), 23, 8)
            Call REPLA(XY$, LOTE$, 3, 16)
            JL& = JL& + 1
            Call GRAREG("!CAREPRO", XY$, JL&)
         End If
         .MoveNext
      Loop
    End With
    Call SETULTREG("!CAREPRO", JL&)
    
    DETATMLP.Close
    Set DETATMLP = Nothing
    
40  VTB% = VENTABU%("CAREPRO1/TITUPAN=Ingreso de Códigos")
    If VTB% < 0 Then GoTo 10
    
    JJ& = 0
    '
    'VALIDO QUE LOS ITEM INGRESADOS TENGAN CANTIDAD
    For U& = 1 To ULTREG&("!CAREPRO")
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CIXI% = CVI(Left$(XY$, 2))
       CANTI = CVD(Mid$(XY$, 23, 8))
       LOTE$ = TRIM$(Mid$(XY$, 3, 16))
       If CANTI <= 0 Then
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODEXT$(CIXI%)))
          VDEF$ = OBX$
          GoTo 40
       End If
       If LOTE$ = "" Then
          Call MENSERR(24, "Falta Lote en Código " + TRIM$(CODEXT$(CIXI%)))
          GoTo 40
       Else
          CONDI$ = "COD_INTE=" & CIXI% & " AND IDSUBOR='" & LOTE$ & "'"
          If CantRegistros("ORDEFABR", CONDI$) < 1 Then
             Call MENSERR(24, "El Lote Elegido No Corresponde Al Código " + TRIM$(CODEXT$(CIXI%)))
             GoTo 40
          End If
       End If
    Next U&
    '
    'GRABO ENCABEZADO
    CONDI$ = "NRODESPA =" & NRODES&
    FEX% = HOY(HOS$)
    FEX1 = FETEXCOR1$(FEX%)
    Call ACTUREGISTRO("ENCADESIN", "FECHDESPA", CONDI$, FEX1, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "SECTOREMI", CONDI$, SectEmi$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "SECTORDESTI", CONDI$, SecDest$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "ORDFAB", CONDI$, NORFA$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "CANTTOTAL", CONDI$, CANTITOT, REGAF&, "NOMESS")
    '
    'GRABO EL DETALLE
    Call ACTUREGISTRO("DETADESIN", "MARBORRA", " NRODESPA = " & NRODES&, 1, REG&)
    
    For U& = 1 To ULTREG&("!CAREPRO")
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CIXI% = CVI(Left$(XY$, 2))
       CANTI = CVD(Mid$(XY$, 23, 8))
       LOTE$ = Mid$(XY$, 3, 16)
       '
100    ABUS& = RANDSTRING$(-7)
       CONDI$ = "NRODESPA = " & ABUS& & ""
       If CantRegistros("DETADESIN", CONDI$) > 0 Then GoTo 100
       Call CreaRegistro("DETADESIN", "NRODESPA", ABUS&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "COD_INTE", CONDI$, CIXI%, REGAF&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "IDENLOTE", CONDI$, LOTE$, REGAF&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "CANTIDAD", CONDI$, FORMATNUM$(CANTI, "F16.5"), REGAF&, "NOMESS")
       '
       Call ACTUREGISTRO("DETADESIN", "NRODESPA", CONDI$, NRODES&, REGAF&, "NOMESS")
    Next U&
    '
    Call BORRAREGISTROS("DETADESIN", "MARBORRA = 1", REG&, "NOMESS")
    '
End Sub

Sub NUECOMPINT()
    
    PROXNUMCOMINT& = 1 + XVALO(VALOBADA("ENCADESIN", "Max(NRODESPA)", "NRODESPA > 0", "NOMESS"))
    Call REPLA(VDEF$, MKS$(PROXNUMCOMINT&), 1, 4)

10  OBX$ = OBJPLA$("ALTACOMP", VDEF$)
    If OBX$ = "" Then Exit Sub
    
    SectEmi$ = TRIM$(Mid$(OBX$, 5, 30))
    SecDest$ = TRIM$(Mid$(OBX$, 35, 30))
    NORFA$ = TRIM$(Mid$(OBX$, 65, 12))
    CANTITOT = CVS(Mid$(OBX$, 77, 4))
    '
    'VALIDACIONES
    If SectEmi$ = "" Then
      Call COMUNI("Falta Ingresar Sector Emisor")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    If SecDest$ = "" Then
      Call COMUNI("Falta Ingresar Sector Destino")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    Call SETULTREG("!CAREPRO", 0)
40  VTB% = VENTABU%("CAREPRO1/TITUPAN=Ingreso de Códigos")
    If VTB% < 0 Then GoTo 10
    
    JJ& = 0
    '
    'VALIDO QUE LOS ITEM INGRESADOS TENGAN CANTIDAD
    For U& = 1 To ULTREG&("!CAREPRO")
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CIXI% = CVI(Left$(XY$, 2))
       CANTI = CVD(Mid$(XY$, 23, 8))
       LOTE$ = TRIM$(Mid$(XY$, 3, 16))
       If CANTI <= 0 Then
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODEXT$(CIXI%)))
          GoTo 40
       End If
       If LOTE$ = "" Then
          Call MENSERR(24, "Falta Lote en Código " + TRIM$(CODEXT$(CIXI%)))
          GoTo 40
       Else
          CONDI$ = "COD_INTE=" & CIXI% & " AND IDSUBOR='" & LOTE$ & "'"
          If CantRegistros("ORDEFABR", CONDI$) < 1 Then
             Call MENSERR(24, "El Lote Elegido No Corresponde Al Código " + TRIM$(CODEXT$(CIXI%)))
             GoTo 40
          End If
       End If
    Next U&
    '
    'GRABO ENCABEZADO
    'ACTUALIZO POR SI CAMBIO
    PROXNUMCOMINT& = 1 + XVALO(VALOBADA("ENCADESIN", "Max(NRODESPA)", "NRODESPA > 0", "NOMESS"))
    Call CreaRegistro("ENCADESIN", "NRODESPA", PROXNUMCOMINT&, "NOMESS")
    
    CONDI$ = "NRODESPA =" & PROXNUMCOMINT&
    FEX% = HOY(HOS$)
    FEX1 = FETEXCOR1$(FEX%)
    Call ACTUREGISTRO("ENCADESIN", "FECHDESPA", CONDI$, FEX1, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "SECTOREMI", CONDI$, SectEmi$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "SECTORDESTI", CONDI$, SecDest$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "ORDFAB", CONDI$, NORFA$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("ENCADESIN", "CANTTOTAL", CONDI$, CANTITOT, REGAF&, "NOMESS")
    '
    'GRABO EL DETALLE
    For U& = 1 To ULTREG&("!CAREPRO")
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CIXI% = CVI(Left$(XY$, 2))
       CANTI = CVD(Mid$(XY$, 23, 8))
       LOTE$ = Mid$(XY$, 3, 16)
       '
100    ABUS& = RANDSTRING$(-7)
       CONDI$ = "NRODESPA = " & ABUS& & ""
       If CantRegistros("DETADESIN", CONDI$) > 0 Then GoTo 100
       Call CreaRegistro("DETADESIN", "NRODESPA", ABUS&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "COD_INTE", CONDI$, CIXI%, REGAF&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "IDENLOTE", CONDI$, LOTE$, REGAF&, "NOMESS")
       Call ACTUREGISTRO("DETADESIN", "CANTIDAD", CONDI$, FORMATNUM$(CANTI, "F16.5"), REGAF&, "NOMESS")
       '
       Call ACTUREGISTRO("DETADESIN", "NRODESPA", CONDI$, PROXNUMCOMINT&, REGAF&, "NOMESS")
    Next U&
    '
End Sub

Sub NUEVANOCONFORM()
    
    ANO% = Year(Date)
    CONDI$ = "NUMENC > 0 AND ANONC = " & ANO%
    NUMNC& = 1 + XVALO(VALOBADA("ACCORREC", "MAX(NUMENC)", CONDI$, "NOMESS"))
    
    FE% = HOY(HO$)
    FILTX$ = TRIM$(Str$(NUMNC&)) & ";" & TRIM$(Str$(FE%))
    ARCHIMPSQR$ = CONTROL$("", "DEMANACC")
    '
    If ARCHIMPSQR$ = "" Then
        Call COMUNI("Falta Formulario de Impresión de Demando de Acción Correctiva")
        Exit Sub
    End If
    '
    NOCONF09.MODIFICACION% = 0
    NOCONF09.Label4 = TRIM$(Str$(NUMNC&))
    HOII% = HOY(HOS$)
    NOCONF09.Label5 = FECHATEX$(HOII%)
    NOCONF09.Show 1
    Exit Sub
    '
'    Call STDFORM(ARCHIMPSQR$, FILTX$)
'
'    If STDSCREEN.APRO = "OK" Then
'       CONDI$ = "NUMENC =  " & NUMNC&
'       Call ACTUREGISTRO("ACCORREC", "STATUS", CONDI$, 1, REGAF&, "NOMESS")
'       Call GRACONCATENAR
'    End If

End Sub

Sub PRINTCOMPINT()

    'PRESENTA CARGALISEL
    Call CARGALISEL("DETADES", "ENCADESIN", "NRODESPA/F6.0;ORDFAB/A16;FECHDESPA/D8;SECTOREMI/A18;SECTORDESTI/A18", "NRODESPA > 0 ORDER BY NRODESPA ASC")
10  Call SELECHO("DETADES/COMPROBANTES INTERNOS ")
    NRODES& = XVALO(VALACT1$("DETADES"))
    If NRODES& < 1 Then Exit Sub
    '
    'VENTANA DE IMPRESION
    FILTX$ = TRIM$(Str(NRODES&))
    ARCHIMPSQR$ = CONTROL$("", "COMPINTE")
    '
    If ARCHIMPSQR$ = "" Then
        Call COMUNI("Falta Formulario de Impresión de Despachos Internos")
        Exit Sub
    End If

    Call STDFORM(ARCHIMPSQR$, FILTX$)
    GoTo 10

End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub



Private Sub Command11_Click()
     '
    Call CIERRARCH("*.*")
    Call FINAL("")
     '
End Sub

Private Sub Command15_Click()
   '
   Command15.Enabled = False
10  OPPX% = COMALTER%("Elija Opción :\\Imprimir O-Fabricación\Imprimir Despachos Internos")

   If OPPX% < 1 Or OPPX% > 2 Then GoTo 99
   
30 Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
   NORFA$ = TRIM$(VALACT1$("ECORDEF"))
   If NORFA$ = "" Then GoTo 10
   FILTX$ = TRIM$(NORFA$)

   If OPPX% = 1 Then
        'LLAMAR A REPORTE DE IMPRESION ORFA-ALR
        ARCHIMPSQR$ = CONTROL$("", "FORMOFAB")
        If ARCHIMPSQR$ = "" Then
           Call COMUNI("Falta Formulario de Impresión de Orden De Fabricación")
           GoTo 99
        End If
        Call STDFORM(ARCHIMPSQR$, FILTX$)
        GoTo 30
        '
    ElseIf OPPX% = 2 Then
        'LLAMAR A REPORTE DE IMPRESION DE DESPACHOS INTERNOS
        ARCHIMPSQR$ = CONTROL$("", "DESPINTE")
        If ARCHIMPSQR$ = "" Then
           Call COMUNI("Falta Formulario de Impresión de Despachos Interno")
           GoTo 99
        End If
        Call STDFORM(ARCHIMPSQR$, FILTX$)
        GoTo 30
   End If
   '
99  Command15.Enabled = True
End Sub

Private Sub Command25_Click()

   Command25.Enabled = False
   '
   Call ABRECONEXION
   'CREAR TABLA DE ENCABEZADO
   Call CREACAMPO("", "ENCADESIN", "NroDespa", 4, 0, 1)
   Call CREACAMPO("", "ENCADESIN", "FechDespa", 8, 0)
   Call CREACAMPO("", "ENCADESIN", "SectorEmi", 10, 128)
   Call CREACAMPO("", "ENCADESIN", "SectorDesti", 10, 128)
   Call CREACAMPO("", "ENCADESIN", "OrdFab", 10, 128)
   Call CREACAMPO("", "ENCADESIN", "CantTotal", 6, 0)
   '
   'CREAR TABLA DE DETALLE
   Call CREACAMPO("", "DETADESIN", "NroDespa", 4, 0, 1)
   Call CREACAMPO("", "DETADESIN", "Cod_Inte", 3, 0)
   Call CREACAMPO("", "DETADESIN", "IdenLote", 10, 12)
   Call CREACAMPO("", "DETADESIN", "Cantidad", 6, 0)
   Call CREACAMPO("", "DETADESIN", "MARBORRA", 3, 0)
   '
10 OPPX% = COMALTER%("Elija Opción:\\Nuevo Compobante Interno\Modificar Comprobante\Imprimir Comprobante Interno")
   If OPPX% < 1 Or OPPX% > 3 Then GoTo 99
   If OPPX% = 1 Then
      Call NUECOMPINT
      GoTo 10
    ElseIf OPPX% = 2 Then
      Call MODICOMPINT
      GoTo 10
    ElseIf OPPX% = 3 Then
      Call PRINTCOMPINT  ' IMPRIME COMPROBANTE
      GoTo 10
   End If
   
99 Command25.Enabled = True
End Sub

Private Sub Command26_Click()
  Command26.Enabled = False
  CFGISO9K.Show 1
  Command26.Enabled = True
End Sub

Private Sub Command3_Click()
  
   Command3.Enabled = False
  
     Call ABRECONEXION
    'CREAR TABLA DE ACCORREC
    Call CREACAMPO("", "ACCORREC", "NumeNC", 4, 0, 1)
    Call CREACAMPO("", "ACCORREC", "FechaNC", 8, 0)
    Call CREACAMPO("", "ACCORREC", "AnoNC", 3, 0, 1)
    Call CREACAMPO("", "ACCORREC", "Cod_Inte", 3, 0)
    Call CREACAMPO("", "ACCORREC", "Descripcion", 10, 128)
    Call CREACAMPO("", "ACCORREC", "AntecNC", 10, 32) ' se cambia a texto por que poseen formtato alfanumerico anteriores
    Call CREACAMPO("", "ACCORREC", "FechaAntecNC", 8, 0)
    Call CREACAMPO("", "ACCORREC", "ResponNC", 10, 128)
    Call CREACAMPO("", "ACCORREC", "Catergoria", 3, 0)
    Call CREACAMPO("", "ACCORREC", "Canatidad", 6, 0)
    Call CREACAMPO("", "ACCORREC", "IDSUBOR", 10, 24)
    Call CREACAMPO("", "ACCORREC", "IDSUBOR1", 10, 24)
    Call CREACAMPO("", "ACCORREC", "IDSUBOR2", 10, 24)
    Call CREACAMPO("", "ACCORREC", "IDSUBOR3", 10, 24)
    Call CREACAMPO("", "ACCORREC", "ORDFAB", 10, 128)
    Call CREACAMPO("", "ACCORREC", "DescripcionNC", 12, 0)
    Call CREACAMPO("", "ACCORREC", "TipoCausa", 10, 1)
    Call CREACAMPO("", "ACCORREC", "TipoCausa1", 10, 1)
    Call CREACAMPO("", "ACCORREC", "TipoCausa2", 10, 1)
    Call CREACAMPO("", "ACCORREC", "TipoCausa3", 10, 1)
    Call CREACAMPO("", "ACCORREC", "CantCausa", 6, 0)
    Call CREACAMPO("", "ACCORREC", "CantCausa1", 6, 0)
    Call CREACAMPO("", "ACCORREC", "CantCausa2", 6, 0)
    Call CREACAMPO("", "ACCORREC", "CantCausa3", 6, 0)
    Call CREACAMPO("", "ACCORREC", "Causas", 10, 24)
    Call CREACAMPO("", "ACCORREC", "DescrpCausa", 12, 0)
    Call CREACAMPO("", "ACCORREC", "Destino", 3, 0)
    Call CREACAMPO("", "ACCORREC", "DescripDestino", 12, 0)
    Call CREACAMPO("", "ACCORREC", "DescripAccCorr", 12, 0)
    Call CREACAMPO("", "ACCORREC", "ResponsableNC", 10, 24)
    Call CREACAMPO("", "ACCORREC", "FechaNotificacion", 8, 0)
    Call CREACAMPO("", "ACCORREC", "FechaImplementacion", 8, 0)
    Call CREACAMPO("", "ACCORREC", "Status", 3, 0)
    Call CREACAMPO("", "ACCORREC", "CierreTec", 8, 0)
    Call CREACAMPO("", "ACCORREC", "UsCierreTec", 3, 0)
    Call CREACAMPO("", "ACCORREC", "CierreAdm", 8, 0)
    Call CREACAMPO("", "ACCORREC", "UsCierreAdm", 3, 0)
    Call CREACAMPO("", "ACCORREC", "SecDetectoNC", 3, 0)
    Call CREACAMPO("", "ACCORREC", "NotaCredito", 10, 24)
    Call CREACAMPO("", "ACCORREC", "TextoImplementacion", 10, 64)
    '
    Call PoneEnCeroNull("ACCORREC", "USCIERRETEC", REGAF&, "NOMESS")
    Call PoneEnCeroNull("ACCORREC", "USCIERREADM", REGAF&, "NOMESS")
    '
    Call CARGAAÑO
10  OPX% = ALTERNA%("Nueva No Conformidad / Accion Preventiva\Modificar No Conformidad / Accion Preventiva\Cierre Técnico de NC / AP\Cierre Administrativo de NC / AP \ Consulta de N-C Cerradas")
    '
    
    If OPX% = 1 Then
        Call NUEVANOCONFORM
        GoTo 10
      ElseIf OPX% = 2 Then
        Call MODINOCONFORM
        GoTo 10
      ElseIf OPX% = 3 Then
        Call CIERRETECNICO
        GoTo 10
     ElseIf OPX% = 4 Then
        Call CIERREADMIN
        GoTo 10
     ElseIf OPX% = 5 Then
        Call CONSULTANC
        GoTo 10
     End If
      
99   Command3.Enabled = True
   
End Sub
Sub CIERREADMIN()
    If DERACCE%("CIENCADM", "Cierre Administrativo de No Conformidad ") < 2 Then Exit Sub
    '
    'PRESENTA CARGALISEL CON NO CONFORMIDADES  CON CIERRE TECNICO
    Call CARGALISEL("DETANOC", "ACCORREC", "NUMENC/F7;ResponNC/A24;FECHANC/D8", "NUMENC > 0 AND STATUS = 2 AND USCIERRETEC > 0 AND (USCIERREADM < 1 or  USCIERREADM is null) ORDER BY ANONC, NumeNC ASC")
    Call SELECHO("DETANOC/NO CONFORMIDADES ")

    NUMNC& = XVALO(VALACT1$("DETANOC"))
    If NUMNC& < 1 Then Exit Sub
    ANO% = Year(Right$(VALACT2$("DETANOC"), 8))

10  OPPX% = COMALTER%("Confirmación de Cierre Administrativo\\No Cancelar\Sí, Realizar Cierre Administrativo")
    If OPPX% = 2 Then
       FE% = HOY(HO$)
       FECH = FETEXCOR1$(FE%)
       CONDI$ = "NUMENC = " & NUMNC& & "  AND ANONC = " & ANO%
       USERNU% = USNBR% Mod 100
       If USERNU% < 1 Then USERNU% = 1
       Call ACTUREGISTRO("ACCORREC", "CIERREADM", CONDI$, FECH, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "USCIERREADM", CONDI$, USERNU%, REGAF&, "NOMESS")
    End If
End Sub
Sub MODINOCONFORM()
    'PRESENTA CARGALISEL CON NO CONFORMIDADES ABIERTAS EN PROCESO
    Call CARGALISEL("DETANOC", "ACCORREC", "NUMENC/F7;ResponNC/A24;FECHANC/D8", "NUMENC > 0 AND (STATUS < 2 OR STATUS IS NULL) ORDER BY ANONC, NumeNC ASC")
    Call SELECHO("DETANOC/NO CONFORMIDADES ")
    NUMNC& = XVALO(VALACT1$("DETANOC"))
    If NUMNC& < 1 Then Exit Sub
    NOCONF09.ANOX% = Year(Right$(VALACT2$("DETANOC"), 8))
    FILTX$ = TRIM$(Str$(NUMNC&))
    ARCHIMPSQR$ = CONTROL$("", "DEMANACC")
    '
    If ARCHIMPSQR$ = "" Then
        Call COMUNI("Falta Formulario de Impresión de Demando de Acción Correctiva")
        Exit Sub
    End If
    NOCONF09.MODIFICACION% = 1

    HOII% = HOY(HOS$)
    NOCONF09.Label5 = FECHATEX$(HOII%)
    NOCONF09.Label4 = TRIM$(Str$(NUMNC&))
    NOCONF09.Show 1
    Exit Sub

'    Call STDFORM(ARCHIMPSQR$, FILTX$)
    '
'    If STDSCREEN.APRO = "OK" Then
'        'GUARDO LAS ORDENES DE FABRICACION CONCATENADAS EN EL CAMPO ORDFAB
'        '*****************************************************************
'        ORFA$ = TRIM$(VALOFORM("TEXT-19"))
'        ORFA1$ = TRIM$(VALOFORM("TEXT-21"))
'        ORFA2$ = TRIM$(VALOFORM("TEXT-22"))
'        ORFA3$ = TRIM$(VALOFORM("TEXT-24"))
'        If ORFA$ <> "" Then CONCAORFA$ = ORFA$ + ";"
'        If ORFA$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA1$ + ";"
'        If ORFA2$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA2$ + ";"
'        If ORFA3$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA3$ + ";"
'        If InStr(CONCAORFA$, ";") > 0 Then CONCAORFA$ = Left$(CONCAORFA$, Len(CONCAORFA$) - 1)
'        '
'        'GUARDO LAS CAUSAS CONCATENADAS EN EL CAMPO CAUSA
'        '************************************************
'        CAUSA$ = TRIM$(VALOFORM("TEXT-8"))
'        CAUSA1$ = TRIM$(VALOFORM("TEXT-25"))
'        CAUSA2$ = TRIM$(VALOFORM("TEXT-26"))
'        CAUSA3$ = TRIM$(VALOFORM("TEXT-27"))
'        If CAUSA$ <> "" Then CONCACAUSA$ = CAUSA$ + "-"
'        If CAUSA1$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA1$ + "-"
'        If CAUSA2$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA2$ + "-"
'        If CAUSA3$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA3$ + "-"
'        If InStr(CONCACAUSA$, "-") > 0 Then CONCACAUSA$ = Left$(CONCACAUSA$, Len(CONCACAUSA$) - 1)
'        '
'        NUMNC& = XVALO(VALOFORM("PARAM-1"))
'
'        CONDI$ = "NUMENC  = " & NUMNC&
'        Call ACTUREGISTRO("ACCORREC", "ORDFAB", CONDI$, CONCAORFA$, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("ACCORREC", "Causas", CONDI$, CONCACAUSA$, REGAF&, "NOMESS")
'     End If
End Sub
Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    '
    
End Sub

Sub CONSULTANC()
    'PRESENTA CARGALISEL CON NO CONFORMIDADES ABIERTAS EN PROCESO
    Call CARGALISEL("DETANOC", "ACCORREC", "NUMENC/F7;ResponNC/A24;FECHANC/D8", "NUMENC > 0 AND STATUS = 2 ORDER BY ANONC, NumeNC ASC")
    Call SELECHO("DETANOC/NO CONFORMIDADES CERRADAS")
    NOCONF09.MODIFICACION% = 1 'PARA QUE CUANDO EJECUTE EL LABEL4 NO ENTRE A VALIDAR FECHA
    NUMNC& = XVALO(VALACT1$("DETANOC"))
    If NUMNC& < 1 Then Exit Sub
    NOCONF09.ANOX% = Year(Right$(VALACT2$("DETANOC"), 8))
    '
    CONDI$ = "FECHA"
    FE% = CVINT(VALOBADA("ACCORREC", "FechaNC", CONDI$, "NOMESS"))
    If FE% > 0 Then NOCONF09.Label5 = FECHATEX$(FE%)
    
    NOCONF09.Label4 = TRIM$(Str$(NUMNC&))
    CONDI$ = "NUMENC =  " & NUMNC&
'    HOII% = HOY(HOS$)
'    NOCONF09.Label5 = FECHATEX$(HOII%)
     Dim CONTROLX As CONTROL ' declaro el objeto control
     For Each CONTROLX In NOCONF09.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
         If TypeOf CONTROLX Is TextBox Then CONTROLX.Locked = True
         If CONTROLX.Name <> "Command13" And CONTROLX.Name <> "Command11" Then
            If TypeOf CONTROLX Is CommandButton Then CONTROLX.Enabled = False
         End If
         
     Next
  


   
    NOCONF09.Show 1
    

'    FILTX$ = TRIM$(Str$(NUMNC&))
'    ARCHIMPSQR$ = CONTROL$("", "DEMANACC")
'    '
'    If ARCHIMPSQR$ = "" Then
'        Call COMUNI("Falta Formulario de Impresión de Demando de Acción Correctiva")
'        Exit Sub
'    End If
'    '
'    Call STDFORM(ARCHIMPSQR$, FILTX$, "READONLY")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call Command11_Click
End Sub
Sub CARGAAÑO()
   'CARGA EL ANO A LOS QUE NO POSEEN ESTE DATO, (CAMPO AGREGADO POSTERIORMENTE)
   SQLX$ = "Update ACCORREC SET ANONC = YEAR(FechaNC) WHERE ANONC IS NULL OR ANONC = ''"
   FLEXCONN.Execute (SQLX$)
End Sub

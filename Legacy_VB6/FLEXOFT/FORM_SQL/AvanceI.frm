VERSION 5.00
Begin VB.Form Avance 
   Caption         =   "Avance de Proceso"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Proceso en Curso - Aguarde a que se Cierre esta Ventana."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "Avance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PROCESOX$

Private Sub Form_Activate()
   Me.Visible = True
End Sub

Private Sub Form_Load()
    If InStr(PROCESOX$, "/") > 0 Then 'PARA QUE SE LE PUEDA PASAR PARAMETROS
      Dim strArray() As String
      On Error Resume Next
      strArray = Split(PROCESOX$, "/")
      PROCESOX$ = strArray(0)
    End If

    Select Case PROCESOX$
       Case "REFRESHECO"
         Call REFRESHECO
       Case "GENERAOF_ENVAPLAST"
         NPED& = XVALO(strArray(1))
        ' Call PRO_GENERAOFENVAP(NPED&)
       Case "GENERAOF_ENVAPLAST_APRO"
         NPED& = XVALO(strArray(1))
         'Call GENERAOF_ENVAPLAST_APRO(NPED&)
    End Select
    Unload Me
End Sub
'Sub PRO_GENERAOFENVAP(NPED&)
'    '
'    If NPED& > 0 Then
'        CODI$ = ""
'        FIN& = ULTREG&("!CARDETPE")
'        For U& = 1 To ULTREG&("!CARDETPE")
'
'         Label2 = U& & "  de  " & FIN&
'         Label2.Refresh
'
'         X$ = REGLEIDO$("!CARDETPE", U&)
'         CAN = CVS(Mid$(X$, 47, 4))
'         If CAN > 0.05 Then
'           '
'           CIXI% = CVI(Left$(X$, 2))
'           If CIXI% > 0 Then
'              If TIMATE%(CIXI%) < 2 Then GoTo 50
'              FENTRESOL% = CVI(Mid$(X$, 63, 2))
'              If TRIM$(APROBO$(CIXI%)) = "" Then
'                 If CODI$ = "" Then
'                   CODI$ = CODI$ + CODEXT$(CIXI%)
'                   Articulo$ = " El "
'                   GoTo 50
'                 Else
'                   CODI$ = CODI$ & " , " & CODEXT$(CIXI%)
'                   Articulo$ = " Los "
'                   GoTo 50
'                 End If
'              End If
'              Call GENERAR_OF_SIMPLE(CIXI%, XVALO(Text1), CAN, FENTRESOL%, NPED&, TRIM$(Text10), RETORNO_NORFA$)
'              FILTX$ = TRIM$(Str$(CIXI%)) & ";" & TRIM$(Str$(NPED&)) & ";" & TRIM$(RETORNO_NORFA$)
'              Call STDFORM("ORTR-ENV", FILTX$, "PRINT")
'           End If
'         End If
'50       Next U&
'         '
'         If CODI$ <> "" Then
'            TXX1$ = "" & Articulo$ & "Código" & Mid$(Articulo$, 3, 1) & " " & CODI$ & " Faltan ser Aprobados No se Generará la Orden de Trabajo Para Estos"
'            Call COMUNI(TXX1$)
'         End If
'         '
'     End If
'End Sub


Sub REFRESHECO()
    '
    Dim RST As ADODB.Recordset
    Set RST = READSET("SELECT IDSUBOR, CanProy, COD_INTE, StatuOrf FROM ORDEFABR WITH(NOLOCK) ORDER BY IDSUBOR")
    FIN& = CantRegistros("ORDEFABR", "IDSUBOR <> ''", "NOMESS")
    '
    REGF& = 0: REGP& = 0
    With RST
        Do While Not .EOF
            Call TRACE(24, REGF&, FIN&)
            '
            If REGF& Mod 10 = 0 Then
              Label2 = REGF& & "  de  " & FIN&
              Label2.Refresh
            End If
            NORFA$ = SAFETEXT$(!IDSUBOR)
            CANTI# = XVALO(!CANPROY)
            CIXI% = XVALO(!COD_INTE)
            STAT% = XVALO(!StatuOrf)
            DESCRI$ = TRIM$(CODEXT$(CIXI%)) + " - " + TRIM$(DESCRIT$(CIXI%))
            CANS$ = " - (" + TRIM$(Str$(CANTI#)) + ")"
            DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
            ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
            '
            XX$ = REGBLAN$("ECORDEF")
            Call REPLA(XX$, NORFA$, 1, 12)
            Call REPLA(XX$, DESCRI1$, 13, 52)
            '
            REGF& = REGF& + 1
            Call GRAREG("ECORDEF", XX$, REGF&)
            If STAT% = 1 Then
                REGP& = REGP& + 1
                Call GRAREG("ECORDEP", XX$, REGP&)
            End If
            
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    '
    Call SETULTREG("ECORDEF", REGF&)
    Call SETULTREG("ECORDEP", REGP&)
    '
    Call CIERRARCH("ECORDEF")
    Call CIERRARCH("ECORDEP")
    '
End Sub

Sub GENERAOF_ENVAPLAST_APRO(NPED&)

     Dim obj As New RECORXET
     Set RS11 = obj.RS_DETAPEDIDO(NPED&)
     FIN& = CantRegistros("PEDDETA", "NROPED = " & NPED&, "NOMESS")
     
     CODI$ = "": U& = 0
     With RS11
      Do While Not .EOF
          NC = XVALO(!NROCLIE)
          CIXI% = XVALO(!CODIINT)
          U& = U& + 1
          Label2 = U& & "  de  " & FIN&
          Label2.Refresh
          CAN = XVALO(!CanPed)
          If CAN > 0.05 Then
            If CIXI% > 0 Then
               If TIMATE%(CIXI%) < 2 Then GoTo 50 ' PARA QUE NO GENERE OT POR MATERIAS PRIMAS.
               FENTRESOL% = CVINT(!FeSolEnt)
               If TRIM$(APROBO$(CIXI%)) = "" Then
                 If CODI$ = "" Then
                   CODI$ = CODI$ + CODEXT$(CIXI%)
                   Articulo$ = " El "
                   GoTo 50
                 Else
                   CODI$ = CODI$ & " , " & CODEXT$(CIXI%)
                   Articulo$ = " Los "
                   GoTo 50
                 End If
              End If
              Call FORPED07.GENERAR_OF_SIMPLE(CIXI%, NC, CAN, FENTRESOL%, NPED&, OrdenCompraPed$(NPED&), RETORNO_NORFA$)
              FILTX$ = TRIM$(Str$(CIXI%)) & ";" & TRIM$(Str$(NPED&)) & ";" & TRIM$(RETORNO_NORFA$) & ";" & TRIM$(MEDIDA_ENVAPLAST$(CIXI%, RETCORTE#))
              Call STDFORM("ORTR-ENV", FILTX$, "PRINT")
           End If
          End If
50       .MoveNext
      Loop
     End With

    If CODI$ <> "" Then
        TXX1$ = "" & Articulo$ & "Código" & Mid$(Articulo$, 3, 1) & " " & CODI$ & " Faltan ser Aprobados No se Generará la Orden de Trabajo Para Estos"
        Call COMUNI(TXX1$)
    End If
               
    '
    On Error Resume Next
    Set obj = Nothing
    RS11.Close
    Set RS11 = Nothing


End Sub





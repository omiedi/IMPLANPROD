VERSION 5.00
Begin VB.Form FRMLECTARCH 
   Caption         =   "Transferencia de Archivos"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Lectura de Transferencia de Archivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "FRMLECTARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '\\\OT-08-1044-14/11/08///
    Call ABREMASTER
    Call ABRESTOCKS
    
    Command1.Enabled = False
    '
    'RECORRO LA TABLA TRANMOVI
    '
    SQLX$ = "SELECT * FROM TRANMOVI"
    SQLX$ = SQLX$ + " WHERE STATUS < 2"
    Dim TRANTMP As ADODB.Recordset
    Set TRANTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    With TRANTMP
      Do While Not .EOF
        CODD$ = SAFETEXT$(TRIM$(!COD_EXTE))
        MAXCART% = CONTROL$("MASTER", "LOMAXCOD")
        If Len(CODD$) > MAXCART% Then
          Call MENSERR(24, "Código " & CODD$ & "Supera el Límite de   " & MAXCART% & "  Caracteres")
          GoTo 50
        End If

        DESS$ = SAFETEXT$(!DESCRIPCION)
        DES22$ = SAFETEXT$(!Descridos)
        'BUSCO EL CODIGO DE TRANMOVI SI ESTA EN EL MASTER
        
        'NOTA :EN SQL LA CONSULTA NO CONTEMPLA NI LOS ESPACIOS
        'EXTENOS DEL STRING NI LAS MAYUSCULAS O MINUSCULAS
        '
        CODMASTER$ = VALOBADA("MASTER", "Codigo", "Codigo = '" & CODD$ & "'")
        If CODMASTER$ <> "" Then ' SI LO ENCUENTRA EDITO LA DESCRIPCION
           SQLX$ = "UPDATE MASTER SET DESCRIPCION = '" & DESS$ & "'"
           SQLX$ = SQLX$ + " WHERE Codigo ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
           '
           'PONGO A 2 EL STATUS
           '
           SQLZ$ = "UPDATE TRANMOVI SET STATUS = 2"
           SQLZ$ = SQLZ$ + " WHERE COD_EXTE ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLZ$))
           '
        Else
           '
           'CONSIGO EL VALOR MAXIMO DEL CODIGO INTERNO
           '
           SQLXY$ = "SELECT MAX(Codint)AS NUMAXX FROM MASTER"
           Dim NUMAXTMP As ADODB.Recordset
           Set NUMAXTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLXY$))
           VALMAX% = 1 + XVALO(NUMAXTMP!NUMAXX)
           '
           'AGREGO EL NUEVO REGISTRO CON EL SIGUIENTE DE CODINT
           '
           SQLR$ = "INSERT INTO MASTER (CODIGO,  Codint, "
           SQLR$ = SQLR$ + " Descripcion, Descridos)VALUES('" & CODD$ & "' ,"
           SQLR$ = SQLR$ + "" & VALMAX% & ","
           SQLR$ = SQLR$ + "'" & DESS$ & "',"
           SQLR$ = SQLR$ + " '" & DES22$ & "')"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLR$))
           '
           'PONGO EN 2 EL STATUS
           '
           SQLZ$ = "UPDATE TRANMOVI SET STATUS = 2"
           SQLZ$ = SQLZ$ + " WHERE Cod_Exte ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLZ$))

        End If
50      .MoveNext
      Loop
    TRANTMP.Close
    Set TRANTMP = Nothing
    
    End With
    '
    '////AHORA LO MISMO PERO CON LA TABLA TRANRECEP
    '
    SQLX$ = "SELECT * FROM TRANRECEP"
    SQLX$ = SQLX$ + " WHERE STATUS < 2"
    Dim TRANRETMP As ADODB.Recordset
    Set TRANRETMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    With TRANRETMP
    'RECORRO TRANRECEP
      Do While Not .EOF
        CODD$ = SAFETEXT$(TRIM$(!COD_EXTE))
        'MAXIMO DE LONGITUD CONFIGURADA DE CODIGO
        MAXCART% = CONTROL$("MASTER", "LOMAXCOD")
        If Len(CODD$) > MAXCART% Then
          Call MENSERR(24, "Código " & CODD$ & "Supera el Límite de   " & MAXCART% & "  Caracteres")
          GoTo 60
        End If
        
        DESS$ = SAFETEXT$(!DESCRIPCION)
        'BUSCO EL CODIGO DE TRANRECEP EN EL MASTER
        CODMASTER$ = VALOBADA("MASTER", "Codigo", "Codigo = '" & CODD$ & "'")
        'SI EXISTE HAGO UN UPDATE EN EL MASTER
        '
        If CODMASTER$ <> "" Then
           SQLX$ = "UPDATE MASTER SET DESCRIPCION = '" & DESS$ & "'"
           SQLX$ = SQLX$ + " WHERE Codigo ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
        'HAGO UN UPDATE EN TRNARECEP
        '
           SQLZ$ = "UPDATE TRANRECEP SET STATUS = 2"
           SQLZ$ = SQLZ$ + " WHERE COD_EXTE ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLZ$))
        Else
           'CONSIGO EL VALOR MAXIMO DEL CODIGO INTERNO
           '
           SQLXY$ = "SELECT MAX(Codint)AS NUMAXX FROM MASTER"
           Set NUMAXTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLXY$))
           VALMAX% = 1 + XVALO(NUMAXTMP!NUMAXX)
           '
           ' LO INGRESO EN EL MASTER CON EL VALOR MAXIMO + 1
           '
           SQLR$ = "INSERT INTO MASTER (CODIGO,  Codint, "
           SQLR$ = SQLR$ + " Descripcion )VALUES('" & CODD$ & "' ,"
           SQLR$ = SQLR$ + "" & VALMAX% & ","
           SQLR$ = SQLR$ + "'" & DESS$ & "')"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLR$))
           '
           ' PONGO EN 2 EL STATUS
           '
           SQLZ$ = "UPDATE TRANRECEP SET STATUS = 2"
           SQLZ$ = SQLZ$ + " WHERE Cod_Exte ='" & CODD$ & "'"
            FLEXCONN.EXECUTE (FILTSQL$ (SQLZ$))
           '
        End If
60      .MoveNext
      Loop
    End With
    TRANRETMP.Close
    Set TRANRETMP = Nothing
    Call MENSERR(24, "Proceso Terminado")
    Command1.Enabled = True
    '\\\OT-08-1044-14/11/08///
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

End Sub

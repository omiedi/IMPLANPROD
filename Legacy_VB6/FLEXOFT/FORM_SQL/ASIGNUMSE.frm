VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ASIGNUMSE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Números de Serie"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8400
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
   ScaleHeight     =   4500
   ScaleWidth      =   8400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3120
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle de artículos trazables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7245
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowBigSelection=   0   'False
         GridLines       =   2
         AllowUserResizing=   1
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
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4500
      Left            =   7440
      ScaleHeight     =   4440
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   0
      Width           =   1035
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
         Picture         =   "ASIGNUMSE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   2880
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "ASIGNUMSE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Height          =   650
         Left            =   120
         Picture         =   "ASIGNUMSE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "ASIGNUMSE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   2760
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -255
         Picture         =   "ASIGNUMSE.frx":0A68
         Top             =   3840
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ASIGNUMSE.frx":298A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ASIGNUMSE.frx":2BBE
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "N-Serie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "ASIGNUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FINALASIGNACION As Integer

Private Sub BOTREC_Click()
    BOTREC.Enabled = False
    '
    ALTA% = 0
    '
    ' VALIDA QUE TODOS LOS ARTICULOS QUE NO TIENEN ESTRUCTURA, O SEA QUE SON EL COMIENZO DE LA VINCULACION
    ' DE NUMEROS DE SERIE, POSEAN UNO
    VALIDAOK% = 1
    For Fila& = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = Fila&
        MSFlexGrid1.COL = 0
        Codigo$ = MSFlexGrid1.TEXT: CIXI1% = CODINT%(Codigo$)
        If CantRegistros("FORMULAS", "CODICONJ=" & CStr(CIXI1%)) = 0 Then
            MSFlexGrid1.COL = 2
            NSE& = XVALO(MSFlexGrid1.TEXT)
            If NSE& = 0 Then VALIDAOK% = 0
        End If
    Next Fila&
    '
    If VALIDAOK% = 0 Then
        Call COMUNI("Imposible Continuar.\Todos los Artículos Deben Tener N-Serie")
        GoTo 99
    End If
    '
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    Dim RST1 As ADODB.Recordset
    Set RST1 = New ADODB.Recordset
    '
    SQLX$ = "SELECT * FROM TRAZARECE WHERE COD_INTE=0"
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    For Fila& = 1 To MSFlexGrid1.Rows - 1
        With Rst
            MSFlexGrid1.Row = Fila&
            MSFlexGrid1.COL = 0: CIXI% = CODINT(MSFlexGrid1.TEXT)
            MSFlexGrid1.COL = 2: NSERIE$ = MSFlexGrid1.TEXT
            MSFlexGrid1.COL = 3: CFAMIL$ = MSFlexGrid1.TEXT
            MSFlexGrid1.COL = 4: NO$ = MSFlexGrid1.TEXT
            MSFlexGrid1.COL = 5: idenlote$ = MSFlexGrid1.TEXT
            MSFlexGrid1.COL = 6: NSERIEBK$ = MSFlexGrid1.TEXT
            MSFlexGrid1.COL = 7: CIXIREF% = XVALO(MSFlexGrid1.TEXT)
            MSFlexGrid1.COL = 8: Origen$ = TRIM$(MSFlexGrid1.TEXT)
            
            'If TRIM$(NSERIE$) <> TRIM$(NSERIEBK$) And CIXIREF% = 0 Then NO$ = ""
            
            ' GRABA EL LOTE TANTO SI TIENE NUMERO DE SERIE O NO, Y SIEMPRE Y CUANDO EL CODIGO DE REFERENCIA
            ' SEA PROVENGA DE LA APROBACION DE UNA BOLETA DE RECEPCION (CIXIREF% >= 0)
            If CIXIREF% >= 0 Then
                .AddNew
                !cod_inte = CIXI%
                !Cod_Exte = CODEXT$(CIXI%)
                !Descrip = DESCRIT0$(CIXI%)
                !idenlote = idenlote$
                !NUMEROSERIE = XVALO(NSERIE$)
                !CFAMILIA = CFAMIL$
                !Origen = Origen$
                !DESTINO = ""
                !DESTINOCONSUMO = ""
                !Status = 0
                .Update
                ALTA% = 1
            End If
        End With
        '
        ' LE ESCRIBE EL NUMERO DE SERIE AL LOTE DE TIPO ORDEN DE FABRICACION EN TABLA TANUMSE
        If CIXIREF% > 0 And XVALO(NSERIE$) > 0 Then
            Dim RST00 As ADODB.Recordset
            Set RST00 = New ADODB.Recordset
            RST00.Open FILTSQL$("SELECT * FROM TANUMSE WHERE NUMEROSERIE LIKE '" & TRIM$(idenlote$) & "%' ORDER BY NUMEORD ASC"), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            With RST00
                If Not (.EOF And .BOF) Then
                    !NUMEROSERIE = NSERIE$
                    .Update
                End If
            End With
            RST00.Close
            Set RST00 = Nothing
        End If
    Next Fila&
    '
    Rst.Close
    Set Rst = Nothing
    '
    'If ALTA% = 1 Then GoTo 25
    '
    For Fila& = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = Fila&
        MSFlexGrid1.COL = 0: CIXI% = CODINT(MSFlexGrid1.TEXT)
        MSFlexGrid1.COL = 2: NSERIE$ = TRIM$(MSFlexGrid1.TEXT)
        MSFlexGrid1.COL = 3: CFAMIL$ = MSFlexGrid1.TEXT
        MSFlexGrid1.COL = 4: NO$ = MSFlexGrid1.TEXT
        MSFlexGrid1.COL = 5: idenlote$ = MSFlexGrid1.TEXT
        MSFlexGrid1.COL = 6: NSERIEBK$ = TRIM$(MSFlexGrid1.TEXT)
        MSFlexGrid1.COL = 7: CIXIREF% = XVALO(MSFlexGrid1.TEXT)
        MSFlexGrid1.COL = 8: Origen$ = TRIM$(MSFlexGrid1.TEXT)
        
        ' SI LA LLAMADA PROVIENE DE LA ASIGNACION MANUAL DE N-SERIE INGRESA ACA PARA ACTUALIZAR TRAZARECE
        If CIXIREF% < 0 And XVALO(NSERIE$) > 0 Then
            SQLX$ = "SELECT * FROM TRAZARECE WHERE COD_INTE=" & CStr(Abs(CIXI%)) & " AND IDENLOTE='" & idenlote$ & "' AND (NUMEROSERIE='' OR NUMEROSERIE='0')"
            RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            
            RST1!NUMEROSERIE = XVALO(NSERIE$)
            RST1!CFAMILIA = CFAMIL$
            RST1!Origen = Origen$
            RST1!DESTINO = ""
            RST1!DESTINOCONSUMO = ""
            RST1!Status = 1
            RST1.Update
            RST1.Close
        End If
        
        ' EN EL CASO DE QUE SE TRATE DE UNA ORDEN DE FABRICACION, TAMBIEN ACTUALIZA LA TABLA DE NUMEROS DE SERIE TANUMSE
        If CIXIREF% < 0 And XVALO(NSERIE$) > 0 Then
            Dim RST01 As ADODB.Recordset
            Set RST01 = New ADODB.Recordset
            RST01.Open FILTSQL$("SELECT * FROM TANUMSE WHERE NUMEROSERIE LIKE '" & TRIM$(idenlote$) & "%' ORDER BY NUMEORD ASC"), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            With RST01
                If Not (.EOF And .BOF) Then
                    !NUMEROSERIE = NSERIE$
                    .Update
                End If
            End With
            RST01.Close
            Set RST01 = Nothing
        End If
    Next Fila&
    '
25  For Fila& = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = Fila&
        MSFlexGrid1.COL = 2: NSERIE$ = TRIM$(MSFlexGrid1.TEXT)
        MSFlexGrid1.COL = 7: CIXIREF% = Abs(XVALO(MSFlexGrid1.TEXT))
        MSFlexGrid1.COL = 8: Origen$ = MSFlexGrid1.TEXT
        MSFlexGrid1.COL = 5: idenlote$ = TRIM$(MSFlexGrid1.TEXT)
        If CIXIREF% > 0 And XVALO(NSERIE$) > 0 Then
            CONDI$ = "COD_INTE=" & CStr(CIXIREF%) & " AND NUMEROSERIE='" & NSERIE$ & "' AND IDENLOTE='" & Origen$ & "'"
            Call ACTUREGISTRO("TRAZARECE", "DESTINO", CONDI$, idenlote$, REGAF&)
        End If
    Next Fila&
    '
    If ALTA% = 0 Then Call COMUNI("Se Ha Completado el Proceso de Asignación")
    '
    FINALASIGNACION = 1
    Unload Me
    '
99  BOTREC.Enabled = True

End Sub

Private Sub Command1_Click()

    If FINALASIGNACION = 0 Then
        Call COMUNI("Imposible Cancelar la Operación.\Debe Aprobar la Asignación de los Números de Serie\Para Finalizar con Este Paso.")
    ElseIf FINALASIGNACION = 1 Then
        OPX% = COMALTER%("¿Está Segura que Desea Cancelar la Aprobación?\\No, Regresar\Si, Continuar")
        If OPX% = 2 Then
            AA& = XVALO(Label6)
            Label6 = CStr((-1) * AA&)
            Hide
        End If
    End If

End Sub

Private Sub Command2_Click()
    '
    With MSFlexGrid1
        NSERIE$ = TRIM$(Text1.TEXT)
        Codigo$ = Label1
        Fila& = .Row
        .COL = 7
        CIXIREF% = XVALO(.TEXT)
        '
        If NSERIE$ = "" Or NSERIE$ = "0" Then
            .COL = 7: .TEXT = "0"  ' CODIGO DE REFERENCIA
            .COL = 2: .TEXT = "0"  ' NUMERO DE SERIE
            .COL = 6: .TEXT = ""   ' NUMERO DE SERIE PARA COMPROBAR CAMBIO
            .COL = 8: .TEXT = ""   ' ORIGEN
            .Refresh
            Exit Sub
        End If
        '
        ' BUSCA EL NUMERO INGRESA EN LA MISMA LISTA
        For II& = 1 To .Rows - 1
            If II& <> Fila& Then
                .Row = II&: .COL = 2
                If TRIM$(.TEXT) = NSERIE$ Then
                    .COL = 0
                    If TRIM$(.TEXT) = Codigo$ Then
                        Call COMUNI("Imposible Asignar Número de Serie.\El Mismo Ya Se Encuentra Asignado en Esta Lista.")
                        Exit Sub
                    End If
                End If
            End If
        Next II&
        '
        .Row = Fila&
        If CIXIREF% > 0 Then        ' ARTICULO CON FORMULA DE PRODUCTO
            CONDI$ = "(NUMEROSERIE='" & NSERIE$ & "' AND COD_INTE=" & CStr(CIXIREF%) & " AND DESTINO='' AND DESTINOCONSUMO='' AND STATUS=1)" ' VERIFICA SI ESTA OCUPADO
            If CantRegistros("TRAZARECE", CONDI$) = 0 Then
                Call COMUNI("Imposible Asignar N-Serie. El Numero de Serie Ya Se Encuentra Asignado o No Existe\" & _
                            "Codigo= '" & CODEXT$(CIXIREF%) & "' - N-SERIE= '" & NSERIE$ & "'")
                Exit Sub
            Else
                Origen$ = SAFETEXT$(VALOBADA("TRAZARECE", "IDENLOTE", CONDI$))
                .COL = 8: .TEXT = Origen$
            End If
        Else
            CIXI% = CODINT%(Label1.Caption)
            CONDI$ = "CODICONJ=" & CStr(CIXI%)
            ' VERIFICA SI EXISTEN COMPONENTES DEL CUAL SE OBTIENE EL NUMERO DE SERIE
            If CantRegistros("FORMULAS", CONDI$) > 0 Then   ' VERIFICA SI TIENE ALGUN COMPONENTE CON N-SERIE
                Encontro% = 0
                SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ=" & CStr(CIXI%)
                Set RST10 = READSET(SQLX$)
                With RST10
                    Do While Not .EOF
                        CIXIREF0% = XVALO(!codielem)
                        CONDI$ = "NUMEROSERIE='" & NSERIE$ & "' AND COD_INTE=" & CStr(CIXIREF0%) & " AND STATUS=1"
                        If CantRegistros("TRAZARECE", CONDI$) > 0 Then
                            Encontro% = 1
                            DESTINO$ = TRIM$(SAFETEXT$(VALOBADA("TRAZARECE", "DESTINO", CONDI$)))
                            DESTINOCONSUMO$ = TRIM$(SAFETEXT$(VALOBADA("TRAZARECE", "DESTINOCONSUMO$", CONDI$)))
                            If DESTINO$ <> "" Or DESTINOCONSUMO$ <> "" Then Encontro% = -1
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                End With
                On Error Resume Next
                RST0.Close
                Set RST0 = Nothing
                On Error GoTo 0
                If Encontro% = 0 Then
                    Call COMUNI("Imposible Asignar N-Serie. El Mismo No Corresponde\a Nignuno de los Componentes del Artículo '" & TRIM$(Label1.Caption) & "'.")
                    Exit Sub
                ElseIf Encontro% = -1 Then
                    Call COMUNI("Imposible Asignar N-Serie. El Mismo Se Encuentra Ocupado")
                    Exit Sub
                End If
                CIXIREF% = (-1) * CIXIREF0%
            Else ' BUSCA SI EXISTE EN LA FAMILIA
                FAMI$ = CODFAMIL$(CIXI%)
                CONDI$ = "NUMEROSERIE='" & NSERIE$ & "' AND CFAMILIA='" & FAMI$ & "'"
                If CantRegistros("TRAZARECE", CONDI$) > 0 Then
                    Call COMUNI("Imposible Asignar Número de Serie.\El Mismo Se Encuentra Utilizado Dentro de la Familia '" & FAMI$ & "'")
                    Exit Sub
                End If
            End If
        End If
        .COL = 2: .TEXT = NSERIE$
        .COL = 6: .TEXT = NSERIE$
        .COL = 7: .TEXT = CStr(CIXIREF%)
    End With
    
End Sub

Private Sub Command3_Click()
    
    CANTELEGIDA& = 0
    CANTIAPRO& = XVALO(Label6)
    
    With MSFlexGrid1
        For Fila& = 1 To .Rows - 1
            .Row = Fila&: .COL = 0
            If TRIM$(.TEXT) = "Si" Then CANTELEGIDA& = CANTELEGIDA& + 1
        Next Fila&
        If CANTELEGIDA& <> XVALO(CANTIAPRO&) Then
            Call COMUNI("Imposible Continuar. La Cantidad de Números de Serie Elegidos (" & CStr(CANTELEGIDA&) & ")\No se Corresponde con la Cantidad Aprobada (" & CStr(CANTIAPRO&) & ")")
            Exit Sub
        End If
        
        For Fila& = 1 To .Rows - 1
            .Row = Fila&: .COL = 0
            If TRIM$(.TEXT) = "Si" Then
                .COL = 1: Codigo$ = .TEXT: CIXI% = CODINT%(Codigo$)
                .COL = 3: NSERIE$ = .TEXT
                .COL = 4: ILO$ = .TEXT
                CONDI$ = "COD_INTE=" & CStr(CIXI%) & " AND NUMEROSERIE='" & NSERIE$ & "' AND IDENLOTE='" & ILO$ & "'"
                Call ACTUREGISTRO("TRAZARECE", "STATUS", CONDI$, 1, REGAF&)
            End If
        Next Fila&
    End With
    
    FINALASIGNACION = 1
    Unload Me
    
End Sub

Private Sub Form_Load()
    FINALASIGNACION = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FINALASIGNACION = 0 Then Cancel = 1
End Sub

Private Sub MSFlexGrid1_Click()
    
    With MSFlexGrid1
        Fila& = .Row
        .Row = 0: .COL = 0
        If TRIM$(.TEXT) = "Aprueba" Then
            .Row = Fila&
            If Fila& > 0 Then
                .COL = 1: Codigo$ = .TEXT
                .COL = 2: DESCRI$ = .TEXT
                .COL = 3: NSERIE$ = .TEXT
            End If
            Label1.Caption = Codigo$
            Label2.Caption = DESCRI$
            Text1.TEXT = NSERIE$
        Else
            .Row = Fila&
            If Fila& > 0 Then
                .COL = 0: Codigo$ = .TEXT
                .COL = 1: DESCRI$ = .TEXT
                .COL = 2: NSERIE$ = .TEXT
            End If
            Label1.Caption = Codigo$
            Label2.Caption = DESCRI$
            Text1.TEXT = NSERIE$
        End If
    End With
    
End Sub

Private Sub MSFlexGrid1_DblClick()
        
    With MSFlexGrid1
        Fila& = .Row: Columna1% = .COL
        .Row = 0: .COL = 0
        If TRIM$(.TEXT) = "Aprueba" Then
            .Row = Fila&
            ESTADO$ = .TEXT
            If ESTADO$ = "Si" Then
                .TEXT = ""
            Else
                .TEXT = "Si"
            End If
        Else
            .Row = Fila&: .COL = Columna1%
        End If
    End With
    
End Sub

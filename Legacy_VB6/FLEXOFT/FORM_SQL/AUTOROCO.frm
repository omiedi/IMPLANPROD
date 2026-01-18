VERSION 5.00
Begin VB.Form AUTOROCO 
   Caption         =   "Generar Autorización de un Item y Para Un Pedido de Reposición"
   ClientHeight    =   1425
   ClientLeft      =   1935
   ClientTop       =   3600
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleMode       =   0  'User
   ScaleWidth      =   5683.819
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Autorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7440
      TabIndex        =   19
      Top             =   840
      Width           =   2670
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   6720
      Picture         =   "AUTOROCO.frx":0000
      TabIndex        =   0
      Top             =   125
      Width           =   175
   End
   Begin VB.TextBox TXTHASTA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Autorizar Hasta:"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TXTORDENCOMPRA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Ped. Reposición N°:"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TXTSTOCKMAXIMO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Stock Máximo"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TXTRASOCLI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   360
      Width           =   3522
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "R.Social"
      Top             =   120
      Width           =   3522
   End
   Begin VB.TextBox TXTCLIENTE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   12
      Top             =   360
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Cliente:"
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox TXTCODIGO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox TXTDESCRIPCION 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   3173
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Código"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Descripción"
      Top             =   120
      Width           =   3173
   End
   Begin VB.TextBox TXTSTOCKMINIMO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Stock Mínimo"
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnueditar 
      Caption         =   "Editar"
      Begin VB.Menu modiexcep 
         Caption         =   "Modificar Excepcion"
      End
   End
End
Attribute VB_Name = "AUTOROCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FORMULARIOX As Form
Public MSF As MSFlexGrid
Private Sub cmdcancelar_Click()
   Unload Me
End Sub


Private Sub Command1_Click()
  Command1.Enabled = False
  If XVALO(Me.TXTCLIENTE) < 1 Then
    Call COMUNI("Atención Falta N° De Cliente")
    Me.TXTCLIENTE.SetFocus
    GoTo 99
  End If
  '
  If CODINT%(Me.TXTCODIGO) < 1 Then
    Call COMUNI("Atención Código Inexistente o no Válido")
    Command2.SetFocus
    GoTo 99
  End If
  '
  If XVALO(Me.TXTORDENCOMPRA) < 1 Then
    Call COMUNI("Atención Falta N° De Pedido de Reposición")
    Call PINTATEXT(Me.TXTORDENCOMPRA)
    GoTo 99
  End If
  '
  If XVALO(Me.TXTHASTA) < 0.05 Then
    OPX% = COMALTER%("Cantidad a Autorizar <= 0 \Si Es Una Excepción Existente Quedará Anulada\Caso Contrario No se Grabará\\Cancelar\Intentar Grabar")
    If OPX% < 2 Then
      Call PINTATEXT(Me.TXTHASTA)
      GoTo 99
    End If
  End If

  Call GENERAR_ARCHIVO_TRANSFERENCIA_EXCEPCION(XVALO(Me.TXTCLIENTE), XVALO(Me.TXTORDENCOMPRA), CODINT%(Me.TXTCODIGO), XVALO(Me.TXTHASTA))
99 Command1.Enabled = True
End Sub
  Sub GENERAR_ARCHIVO_TRANSFERENCIA_EXCEPCION(NCLIE, OCOM&, CI1%, CANT_AUTORIZADA)
  
   Ruta$ = ""
   OK% = 0
   NC_CONFIG1 = XVALO(CONTROL("RETRASLA", "SUC1CLIE"))
   NC_CONFIG2 = XVALO(CONTROL("RETRASLA", "SUC2CLIE"))
   If NC_CONFIG1 = NCLIE Then
        Ruta$ = CONTROL("RETRASLA", "SUC1RUTA")
    ElseIf NC_CONFIG2 = NCLIE Then
        Ruta$ = CONTROL("RETRASLA", "SUC2RUTA")
   End If
   '
   If Ruta$ = "" Then
      Call COMUNI("Atención !!!!! Sucursal Seleccinada No es Válida\Verifique Configuración de Remito a Sucursales\No se Generó Archivo de Envio de Stock Sugeridos")
      Exit Sub
   End If
   FLEXCONN.BeginTrans
   On Error GoTo ERROR_GUARDAR
   
   SQLX$ = "Select * from Excepcion"
   SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & OCOM&
   SQLX$ = SQLX$ + " AND CODINT = " & CI1%
   SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
    
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   NUSER1% = USNBR% Mod 100
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
       If .EOF Then
         If CANT_AUTORIZADA < 0.05 Then GoTo 40
         .AddNew
       End If
       !Nume_Ocom = OCOM&
       !NUCLIEN = NCLIE
       !CODINT = CI1%
       !CODIGO = CODEXT$(CI1%)
       !FECHA = CVDAT(HOY(HO$))
       !USUARIO = NUSER1%
       !CANTIEXCEPCION = CANT_AUTORIZADA
      .Update
   End With
   
    'CARGA LOS TITULOS
    A$ = Date
    A$ = REPLACAR(A$, "/", "_")
    B$ = Time
    B$ = REPLACAR(B$, ":", "")
    B$ = REPLACAR(B$, ".", "")

    NARCHI2% = FILEOPEN%(Ruta$ + "\ARCHIVO_TRANSF_EXCEPCION " + A$ + " " + B$ + ".TXT", 2, 256)  ' ABRE EL ARCHIVO DE ESCRITURA
    TX$ = AJUSTI$("NCLIENTE", 8): TX$ = TX$ + ";" 'NCLIENTE
    TX$ = TX$ + AJUSTI$("CODIGO", 8): TX$ = TX$ + ";"  'CODIGO
    TX$ = TX$ + AJUSTI$("ULT.ENV.", 8): TX$ = TX$ + ";" 'FECHA ULTIMO ENVIO
    TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$ + ";" 'USUARIO
    TX$ = TX$ + AJUSTI$("ORD.COMP.", 8): TX$ = TX$ + ";" 'ORDEN DE COMPRA
    TX$ = TX$ + AJUSTI$("CANTIDAD", 10):  'CANTIDAD AUTORIZADA
    Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
    
    TX$ = AJUSTI$(Me.TXTCLIENTE, 8): TX$ = TX$ + ";" 'NCLIENTE
    TX$ = TX$ + AJUSTI$(SAFETEXT$(CI1%), 8): TX$ = TX$ + ";"    'CODIGO
    TX$ = TX$ + AJUSTI$(FECHATEX$(HOY(HO$)), 8): TX$ = TX$ + ";" 'FECHA ULTIMO ENVIO
    TX$ = TX$ + AJUSTI$(SAFETEXT$(NUSER1%), 8): TX$ = TX$ + ";" 'USUARIO
    TX$ = TX$ + AJUSTI$(Me.TXTORDENCOMPRA, 8): TX$ = TX$ + ";"  'O.C.
    TX$ = TX$ + AJUSTI$(Me.TXTHASTA, 8): TX$ = TX$ + ";"   'CANTIDAD
    Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
    Close #NARCHI2%
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      
      Exit Sub
    Else
      FLEXCONN.CommitTrans
40    RS.Close
      Set RS = Nothing
      Call COMUNI("Proceso Completo")
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     TXTCODIGO.TEXT = COD$
   End If

End Sub

Private Sub mnusalir_Click()
   Unload Me
End Sub


Private Sub modiexcep_Click()
10  CAMPOS$ = "Fecha/D8;O.C./F8;Código./A20;Descripción/A32;Cant.Apr./F10.1"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    '
    Screen.MousePointer = 11
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    FRMZELECHO.MSF1.Rows = 1
    j& = 0
    ACUMLA_SALDO = 0
    SQLX$ = "SELECT  Nume_Ocom, NUCLIEN,  CODINT, CODIGO,  FECHA,USUARIO,CantiExcepcion FROM EXCEPCION M WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE  NUCLIEN = " & XVALO(Me.TXTCLIENTE)
    
    SQLX$ = SQLX$ + " ORDER BY FECHA"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             CI1% = XVALO(!CODINT)
             '
             j& = j& + 1
             FRMZELECHO.msf2.Rows = j& + 1
             FRMZELECHO.msf2.TextMatrix(j&, 1) = FECHATEX$(CVINT(!FECHA))
             FRMZELECHO.msf2.TextMatrix(j&, 2) = XVALO(!Nume_Ocom)
             FRMZELECHO.msf2.TextMatrix(j&, 3) = SAFETEXT$(!CODIGO)
             
             FRMZELECHO.msf2.TextMatrix(j&, 4) = DESCRIT$(CI1%)
             FRMZELECHO.msf2.TextMatrix(j&, 5) = FORMATNUM$(XVALO(!CANTIEXCEPCION), "f10.1")
           
          .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1
    FRMZELECHO.Show 1
    Me.TXTCODIGO = RESP_ZELE_VECT(3)
    Me.TXTORDENCOMPRA = RESP_ZELE_VECT(2) 'O.CO
    Me.TXTHASTA = RESP_ZELE_VECT(5)
'    TXTORDENCOMPRA.Locked = True
'    TXTORDENCOMPRA.BackColor = &HE0E0E0
End Sub
   
   



Private Sub Text3_Change()
  Text9 = ECOARCH$("PROVED1", MKI$(XVALO(Text3)))
End Sub

Private Sub TXTCLIENTE_Change()
   NC = XVALO(TXTCLIENTE)
   If NC > 0 Then
     Me.TXTRASOCLI = RASOCLI$(NC)
   Else
     Me.TXTRASOCLI = ""
   End If

End Sub

Private Sub TXTCODIGO_Change()
   CI1% = CODINT%(TRIM$(TXTCODIGO))
   TXTDESCRIPCION = DESCRIT$(CI1%)
   NC = XVALO(Me.TXTCLIENTE)
   Me.TXTSTOCKMINIMO = StockMinimoEnSucursal(CI1%, NC)
   Me.TXTSTOCKMAXIMO = StockMaximoEnSucursal(CI1%, NC)
   On Error Resume Next
   Me.TXTORDENCOMPRA.SetFocus
   On Error GoTo 0
End Sub

Private Sub TXTHASTA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command1.SetFocus
   End If
End Sub

Private Sub TXTORDENCOMPRA_DblClick()
    TXTORDENCOMPRA.Locked = False
    TXTORDENCOMPRA.BackColor = Width

End Sub

Private Sub TXTORDENCOMPRA_GotFocus()
    TXTORDENCOMPRA.BackColor = vbYellow
    Call PINTATEXT(TXTORDENCOMPRA)
End Sub


Private Sub TXTORDENCOMPRA_LostFocus()
    TXTORDENCOMPRA.BackColor = vbWhite
End Sub

Private Sub TXTHASTA_GotFocus()
    TXTHASTA.BackColor = vbYellow
    Call PINTATEXT(TXTHASTA)
End Sub


Private Sub TXTHASTA_LostFocus()
    TXTHASTA.BackColor = vbWhite
    TXTHASTA = FORMATNUM$(XVALO(TXTHASTA), "F12.1")
End Sub


Private Sub TXTORDENCOMPRA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call PINTATEXT(Me.TXTHASTA)
   End If

End Sub
Function StockMinimoEnSucursal(CI1%, NC)
   CONDI$ = "CODINT = " & CI1%
   CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
   StockMinimoEnSucursal = XVALO(VALOBADA("StockSugXSuc", "StockMinimoActualEnSuc", CONDI$, "NOMESS"))
End Function
Function StockMaximoEnSucursal(CI1%, NC)
   CONDI$ = "CODINT = " & CI1%
   CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
   StockMaximoEnSucursal = XVALO(VALOBADA("StockSugXSuc", "StockMaximoActualEnSuc", CONDI$, "NOMESS"))
End Function


Private Sub TXTPROV_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call PINTATEXT(Me.TXTHASTA)
   End If
End Sub



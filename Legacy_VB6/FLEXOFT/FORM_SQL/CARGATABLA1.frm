VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CARGATABLA1 
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Dados de Baja"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text5 
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
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "DOBLE CLICK PARA LIMPIAR"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar  a la lista"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   6240
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4200
      OleObjectBlob   =   "CARGATABLA1.frx":0000
      Top             =   120
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "DOBLE CLICK PARA DAR DE ALTA ITEMS DE BAJA O VICEVERSA"
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   8689198
      BackColorBkg    =   16777215
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "CARGATABLA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TABLA1$ 'LA TABLA
Public CAMPO1$ 'NUMERO DE ORDEN O CODIGO
Public CAMPO2$ 'CAMPO DE REFERENCIA
Public CAMPO3$ 'STATUS


Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
     
   Dim LenText As Long, ret As Long
     
   'Si los caracteres presionados están entre el 0 y la Z
   If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
     
   ret = SendMessage(Combo1.hwnd, &H14C&, -1, ByVal Combo1.Text)
     
         If ret >= 0 Then
            LenText = Len(Combo1.Text)
            Combo1.ListIndex = ret
            Combo1.Text = Combo1.List(ret)
            Combo1.SelStart = LenText
            Combo1.SelLength = Len(Combo1.Text) - LenText
              
         End If
   End If
End Sub

Private Sub Command1_Click()

    'DESHABILITADO DESDE LAS PROPIEDADES NO FUNCIONA BIEN CUANDO SE QUIERE MODIFICAR.
    
    Command1.Enabled = False
    PROVIN$ = UCase$(TRIM$(TEXT2))
    If PROVIN$ = "" Then GoTo 99
    '
    Call ABRECONEXION
    '
    'CAMPO1$ 'NUMERO DE ORDEN O CODIGO
    'CAMPO2$ 'CAMPO DE REFERENCIA
    'CAMPO3$ 'STATUS
    '
    CONDI$ = CAMPO2$ & "  = '" & PROVIN$ & "'"
    '
    'SI ESTA EDITANDO NO ENTRA ACA A VALIDAR
    If XVALO(Text5) = 0 Then
        If CantRegistros(TABLA1$, CONDI$, "NOMESS") > 0 Then MsgBox "Dato Existente": GoTo 99
        '
        NUMPROV% = 1 + XVALO(MAXNUMCAMPO&(TABLA1$, CAMPO1$))
    Else
       Command1.Caption = "Agregar  a la lista" ' vuelve a poner el caption
       GoTo 99
       Exit Sub
    End If
    
    SQLX$ = "SELECT " & CAMPO1$ & " AS CAMPOAS1 , " & CAMPO2$ & " AS CAMPOAS2 "
    SQLX$ = SQLX$ + " ," + CAMPO3$ & " AS CAMPOAS3 "
    SQLX$ = SQLX$ + " From " & TABLA1$
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    On Error GoTo ERROR_GUARDAR

    Set AUX1 = New ADODB.Recordset
    AUX1.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    BASSSEX.BeginTrans
    With AUX1
          If (.BOF And .EOF) Then
            .AddNew
            !CAMPOAS1 = NUMPROV%
          End If
            !CAMPOAS2 = PROVIN$
            !CAMPOAS3 = 0
            .Update
    End With
ERROR_GUARDAR:
    If Err <> 0 Then
      BASSSEX.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
    Else
      BASSSEX.CommitTrans
    End If
    
50  AUX1.Close
    Set AUX1 = Nothing
    On Error GoTo 0
    '
    'ACTUALIZO EL LIST
    Call CARGAGRID_CARGATABA(MSF, TABLA1$, CAMPO1$, CAMPO2$, CAMPO3$)

99  Command1.Enabled = True
End Sub

''''Private Sub Command2_Click()
''''    Command2.Enabled = False
''''
''''    If XVALO(Text5) = 0 Then
''''     MsgBox "DEBE SELECCIONAR UN DATO PARA MODIFICAR"
''''     GoTo 99
''''    End If
''''
''''    PROVIN$ = TRIM$(Text2.Text)
''''    If PROVIN$ = "" Then GoTo 99
''''    '
'''''    If CANTREGISTROS&(TABLA1$, CAMPO1$ & " = '" & PROVIN$ & "'") < 1 Then
'''''      MsgBox "DATO INEXISTENTE DEBE AGREGARLO PRIMERO"
'''''      GoTo 99
'''''    End If
''''    '
''''    Call ABRECONEXION
''''    SQLX$ = "SELECT " & CAMPO1$ & " AS CAMPOAS1 FROM " & TABLA1$
''''    SQLX$ = SQLX$ + " WHERE " & CAMPO3$ & "  = " & XVALO(Text5)
''''
''''    On Error GoTo ERROR_GUARDAR
''''
''''    Set AUX1 = New ADODB.Recordset
''''    AUX1.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
''''    BASSSEX.BeginTrans
''''    With AUX1
''''          If (.BOF And .EOF) Then
''''            MsgBox "DATO INEXISTENTE O NO VALIDO"
''''          Else
''''            !CAMPOAS1 = PROVIN$
''''            .Update
''''          End If
''''    End With
''''ERROR_GUARDAR:
''''    If Err <> 0 Then
''''      BASSSEX.RollbackTrans
''''      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
''''    Else
''''      BASSSEX.CommitTrans
''''    End If
''''
''''50  AUX1.Close
''''    Set AUX1 = Nothing
''''    On Error GoTo 0
''''    '
''''    'ACTUALIZO EL LIST
''''    Call CARGAGRID_CARGATABA(MSF, TABLA1$, CAMPO1$, CAMPO2$, CAMPO3$)
''''    Text2 = ""
''''    Text5 = ""
''''
''''
''''99 Command2.Enabled = True
''''End Sub


Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")

    If CAMPO1$ = "" Then Exit Sub ' PARA QUE SI NO SE LE PASEN LOS PARAMETROS
    'NO EJECUTE EL EVENTO LOAD
    NUCOLU% = 1: TEPRUEBA = "12345678": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Codigo"
    NUCOLU% = 2: TEPRUEBA = "ABCDEFGHIJKLLLMNÑOPQRSTUVWXYZABCDEFGHIJK": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Referencia"
    '
    Call CARGAGRID_CARGATABA(MSF, TABLA1$, CAMPO1$, CAMPO2$, CAMPO3$)
    ''
    
End Sub

Private Sub List1_Click()
   If List1.Text <> "" Then
     Command2.Enabled = True
     Command3.Enabled = True
     Text1 = List1.Text
   End If
End Sub
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   i& = MSF.MouseRow
   j& = MSF.MouseCol
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       MSF.Row = i&
       MSF.COL = j&
       '
       NUCOD% = XVALO(MSF.TextMatrix(MSF.Row, 0))
       REFERENCIA$ = TRIM$(MSF.TextMatrix(MSF.Row, 1))
       If NUCOD% > 0 Then
         CONDI$ = CAMPO1$ & " = " & NUCOD%
         CONDI1$ = " AND " & CAMPO3$ & "  < 0"
         
         If CantRegistros&(TABLA1$, CONDI$ + CONDI1$, "NOMESS") > 0 Then ' ESTA DADO DE BAJA
            TXTX$ = "DESEA DAR DE ALTA EL CODIGO: " & NUCOD%
            TXTX$ = TXTX$ + " REFERENCIA: " & REFERENCIA$
            If MsgBox(TTXX$, vbYesNo, TXTX$) <> vbYes Then Exit Sub
            SQLX$ = "UPDATE " & TABLA1$ & " SET STATUS = 0 WHERE " & CONDI$
            BASSSEX.Execute (SQLX$)
            Call CARGAGRID_CARGATABA(MSF, TABLA1$, CAMPO1$, CAMPO2$, CAMPO3$)
         Else
            TXTX$ = "DESEA DAR DE BAJA EL CODIGO: " & NUCOD%
            TXTX$ = TXTX$ + " REFERENCIA: " & REFERENCIA$
            If MsgBox(TTXX$, vbYesNo, TXTX$) <> vbYes Then Exit Sub
            SQLX$ = "UPDATE " & TABLA1$ & " SET STATUS = -1 WHERE " & CONDI$
            BASSSEX.Execute (SQLX$)
          End If
            Call CARGAGRID_CARGATABA(MSF, TABLA1$, CAMPO1$, CAMPO2$, CAMPO3$)

       End If
       '
   End If
End Sub

Private Sub MSF_Click()

   Text5 = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   TEXT2 = TRIM$(MSF.TextMatrix(MSF.Row, 1))
   Command1.Caption = "Modificar Item Activo"
   '
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   
   i& = MSF.MouseRow
   j& = MSF.MouseCol
   '
   Call SELECCION_COLUMNA_0(MSF, i&, j&)
   
End Sub

Private Sub Text1_Change()
   Combo1.Text = Text1
   
End Sub

Private Sub Text5_DblClick()
   If XVALO(Text5) < 1 Then Exit Sub
   If XVALO(Text5) > 0 Then
      Command1.Caption = "Agregar  a la lista"
      Text5 = ""
      TEXT2 = ""
   Else
      
      Command1.Caption = "Modificar Item Activo"
   End If
End Sub

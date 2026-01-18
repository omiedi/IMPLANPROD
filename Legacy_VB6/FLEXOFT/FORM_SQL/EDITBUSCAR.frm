VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form EDITBUSCAR 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación Masiva de Descripciones / Tiempo Unitario / Sector"
   ClientHeight    =   5775
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   5775
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C00000&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   13740
      TabIndex        =   15
      Top             =   4920
      Width           =   13800
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar Activa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Actualiza los Códigos Seleccionados "
         Top             =   120
         Width           =   1965
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   12600
         ScaleHeight     =   75
         ScaleWidth      =   705
         TabIndex        =   17
         Top             =   120
         Width           =   770
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualiza Códigos Involucrados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7560
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Actualiza los Códigos Seleccionados "
         Top             =   120
         Width           =   4365
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   12200
         Picture         =   "EDITBUSCAR.frx":0000
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13730
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2950
         Left            =   9550
         TabIndex        =   14
         Top             =   1700
         Width           =   15
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Códigos Involucrados"
         Top             =   960
         Width           =   6075
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
         Left            =   7550
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Tiem. Unit.''"
         Top             =   270
         Width           =   1100
      End
      Begin VB.TextBox TXTSECCION 
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   540
         Width           =   1950
      End
      Begin VB.TextBox Text11 
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
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   7530
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   540
         Width           =   7500
      End
      Begin VB.CommandButton Command2 
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
         Height          =   360
         Left            =   10600
         TabIndex        =   5
         ToolTipText     =   "Click Para Seleccionar Sección"
         Top             =   550
         Width           =   175
      End
      Begin VB.TextBox TXTNOPER 
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox TXTTIEMPO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   7550
         TabIndex        =   3
         Text            =   "   "
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox Text9 
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
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Sector    /   N°"
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   2160
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2985
         ItemData        =   "EDITBUSCAR.frx":1F22
         Left            =   7560
         List            =   "EDITBUSCAR.frx":1F29
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1680
         Width           =   6075
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Código                   Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   13
         Top             =   1320
         Width           =   6075
      End
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "EDITBUSCAR.frx":1F75
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "EDITBUSCAR.frx":21A9
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "EDITBUSCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON
Public TEXTO_ORINGINAL$
Public TIEMPOUNITARIO_ORIGINAL
Public NUMESECTOR_ORIGINAL&



Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub

Private Sub cmdelinminar_Click()

End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    'ACTUALIZAR
    If TRIM$(TXTDESCRIPCION.TEXT) = "" Then
         Call COMUNI("Falta Ingresar Descripción")
         GoTo 99
    End If
    '
    If XVALO(TXTTIEMPO) < 0.005 Then
          Call COMUNI("Falta Ingresar Tiempo")
          GoTo 99
    End If
    '
    If XVALO(TXTNOPER) < 0.005 Then
          Call COMUNI("Falta Ingresar Sector")
          GoTo 99
    End If
    '
    If XVALO(TXTTIEMPO) > 32767 Or XVALO(TXTTIEMPO) < -32767 Then
         Call COMUNI("Valor Ingresado No Válido")
         GoTo 99
    End If
    '
    'VALIDACION QUE NO SE INGRESEN TEXTOS QUE INCLUYAN COMILLAS SIMPLES.
    If InStr(TXTDESCRIPCION.TEXT, "'") > 0 Then
           Call COMUNI("Se Han Detectado Caracteres Reservados (Comilla Simple)\ El Mismo Serà Reemplazado Por Un Espacio Vacío")
           TXTDESCRIPCION.TEXT = REPLACAR$(TXTDESCRIPCION.TEXT, "'", " ")
    End If

    '
    Screen.MousePointer = 11
    '
    HAYSELEC% = 0

    CONDI$ = "DESCRIPCION = '" & TEXTO_ORINGINAL$ & "'"
    CONDI$ = CONDI$ + " AND TIEMPOUNITARIO = " & TIEMPOUNITARIO_ORIGINAL
    CONDI$ = CONDI$ + " AND SecciPro = " & NUMESECTOR_ORIGINAL&
    '
    M1% = 0
    For I% = 0 To List1.ListCount - 1
      If List1.Selected(I%) = True Then
        If M1% = 0 Then
          CIXI% = CODINT%(TRIM$(Left(List1.List(I%), 16)))
          CONDI$ = CONDI$ + " AND ( CODICONJ =  " & CIXI% & ""
          M1% = 1
        Else
          CIXI% = CODINT%(TRIM$(Left(List1.List(I%), 16)))
          CONDI$ = CONDI$ + " OR CODICONJ =  " & CIXI% & ""
        End If
      End If
    Next I%
    If M1% > 0 Then CONDI$ = CONDI$ + " )"
    
    'SI NO SE SELECCIONO NINGUN CODIGO PRESENTA MENSAJE  Y SALE
    If M1% < 1 Then Call COMUNI("No Se Han Seleccionado Códigos"): GoTo 99
    '
    'GRABO LAS MODIFICACIONES EN TIEMOBRA
    SQLX$ = "UPDATE TIEMOBRA SET DESCRIPCION = '" & TXTDESCRIPCION & "' , TIEMPOUNITARIO = " & XVALO(TXTTIEMPO) & " , SecciPro = " & XVALO(TXTNOPER)
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    For I% = 0 To List1.ListCount - 1 'REFRESCO EL TIEMPO DE LA MANO DE OBRA EN EL MASTER PARA LOS CODIGOS SELECCIONADOS
      If List1.Selected(I%) = True Then
        CI1% = CODINT%(TRIM$(Left(List1.List(I%), 16)))
        TIEMPX = XVALO(VALOBADA("TIEMOBRA", "SUM(TIEMPOUNITARIO * CANTIOPS)", "CODICONJ = " & CI1%, "NOMESS"))
        TIEMPX = FORMATNUM$(TIEMPX / 60, "F12.2")
        Call ACTUREGISTRO("MASTER", "MANO_OBRA_REAL", "CODINT = " & CI1%, TIEMPX, REGAF&, "NOMESS")
      End If
    Next I%
    '
    'REFRESCO VENTANA DE TIEMPOS DE MANO DE OBRA.
    CI1% = CODINT%(TRIM$(TIEMMOB09.Text15))
    TIEMMOB09.Text15 = ""
    TIEMMOB09.Text15 = CODEXT$(CI1%)
    Unload Me
    Unload BUSCA_DES

    '
99  Screen.MousePointer = 1
    Command1.Enabled = True

End Sub
Sub ACTUALIZA_FLEX(MSF As MSFlexGrid, RENGLON%, FILA%, VALOR1$)
    '
    If MSF.Row < 1 Then Exit Sub
    '
    MSF.TextMatrix(RENGLON%, FILA%) = TRIM$(VALOR1$)
    
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("SECPROD")
    COD$ = VALACT1$("SECPROD")
    If TRIM$(COD$) <> "" Then
      TXTNOPER.TEXT = COD$
    End If
    Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If Me.List1.ListCount > 0 Then
       OPX% = COMALTER%("Estos Datos Tienen Códigos Involucrados\\Cancelar\Eliminar Datos")
       If OPX% <= 1 Then GoTo 99
    End If
    CONDI$ = "DESCRIPCION = '" & TEXTO_ORINGINAL$ & "'"
    CONDI$ = CONDI$ + " AND TIEMPOUNITARIO = " & TIEMPOUNITARIO_ORIGINAL
    CONDI$ = CONDI$ + " AND SecciPro = " & NUMESECTOR_ORIGINAL&
    Call BORRAREGISTROS("TIEMOBRA", CONDI$, REGAF&, "NOMESS")
    
    Call BUSCA_DES.BUSCAR
    
99  Command3.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub TXTDESCRIPCION_Change()
  TIEMMOB09.Label7 = TXTDESCRIPCION
End Sub

Private Sub TXTNOPER_Change()
      CS1 = XVALO(TXTNOPER.TEXT)
      If CS1 > 0 Then
         TXTSECCION = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
      Else
         TXTSECCION = ""
      End If

End Sub
Sub cargalist()
   List1.Clear
   TIEMPOUNITARIO_ORIGINAL = REPLACAR(SAFETEXT$(TIEMPOUNITARIO_ORIGINAL), ",", ".")
   TEXTO_ORINGINAL$ = TXTDESCRIPCION.TEXT
   If TRIM$(TEXTO_ORINGINAL$) = "" Then Exit Sub
   SQLX$ = "SELECT CODICONJ FROM TIEMOBRA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE DESCRIPCION LIKE '" & TEXTO_ORINGINAL$ & "'"
   SQLX$ = SQLX$ + " AND TIEMPOUNITARIO = " & TIEMPOUNITARIO_ORIGINAL
   SQLX$ = SQLX$ + " AND SecciPro = " & XVALO(NUMESECTOR_ORIGINAL&)
   '
   Dim RST As ADODB.Recordset
   Set RST = READSET(SQLX$)
   With RST
     Do While Not .EOF
       CIXI% = XVALO(!CODICONJ)
       Call REPLA(TX$, CODEXT$(CIXI%), 1, 16)
       Call REPLA(TX$, DESCRIT0$(CIXI%), 18, 35)
       List1.AddItem TX$
       .MoveNext
     Loop
   End With
   '
   Call seleccion_check 'selecciona todos los items
End Sub
Sub seleccion_check()
   'selecciona todos los items
   For U& = 1 To List1.ListCount
      List1.Selected(U& - 1) = True
   Next U&
End Sub


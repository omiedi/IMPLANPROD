VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form BUSCA_DES 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar o Quitar de la Lista"
   ClientHeight    =   7470
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30081.89
   ScaleMode       =   0  'User
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TXTBUSCAR 
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
      Height          =   495
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   6840
      Width           =   12900
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
      Height          =   2055
      Left            =   5640
      ScaleHeight     =   1995
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   7270
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3120
      OleObjectBlob   =   "BUSCA_DES.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "BUSCA_DES.frx":0234
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   6645
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   870
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "BUSCA_DES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON


Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub



Private Sub Command1_Click()
   Command1.Enabled = False
   'ACTUALIZAR
       If TRIM$(TXTDESCRIPCION.TEXT) = "" Then
           Call COMUNI("Falta Ingresar Descripción")
           Exit Sub
       End If
      '
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           Call COMUNI("Falta Ingresar Cantidad")
           Exit Sub
        End If
        '
        If XVALO(TXTTIEMPO) < 0.005 Then
            Call COMUNI("Falta Ingresar Tiempo")
            Exit Sub
        End If
        '
        If XVALO(TXTNOPER) < 0.005 Then
            Call COMUNI("Falta Ingresar Sector")
            Exit Sub
        End If
        '
        If XVALO(TXTTIEMPO) > 32767 Or XVALO(TXTTIEMPO) < -32767 Then
           Call COMUNI("Valor Ingresado No Válido")
           Exit Sub
        End If
    '
    'CARGA EN LA GRILLA
    With TIEMMOB09
    
'        DESCRI$ = TXTDESCRIPCION.TEXT
'        Call FRATEXTO(DESCRI$, 68)
'        For JJ& = 1 To CARETEX%
'          TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
'        Next JJ&
'        .Label7.Height = 210
'        .Label7 = TXTX$
'        .MSF1.RowHeight(.MSF1.Rows - 1) = .Label7.Height + 100

        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 3, FORMATNUM$(TXTTIEMPO, "F10.0"))
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 2, TXTDESCRIPCION.TEXT)
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 4, FORMATNUM$(TXTCANTIDAD, "F10.0"))
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 5, TXTNOPER)
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 6, TRIM$(TXTSECCION))
        TIME_TOTAL = XVALO(TXTTIEMPO) * XVALO(TXTCANTIDAD)
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 7, SAFETEXT$(TIME_TOTAL))
        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 8, SAFETEXT$(TXTMODIFICADO))
        Call TIEMMOB09.ACOMODARMSF
        Call .ACTUALIZAR_DATOS
    End With
    '
    Unload Me
    
99  Command1.Enabled = True

End Sub


Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 6
   NUCOLU% = 1: TEPRUEBA = "99": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "123456789*123456789*123456789*123456789*123456789*123456789*123456789*12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 3: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "T.Unit"
   NUCOLU% = 4: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "QT."
   NUCOLU% = 5: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 6: TEPRUEBA = "SECCIONSECC": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Secciòn"

   Call APLICACIONSKINS(Me)
   Call BUSCAR
End Sub




Private Sub TXTDESCRIPCION_Change()

End Sub

Private Sub MSF1_Click()
   If TXTBUSCAR <> "" Then
     Cargar_art_seleccionado
     Unload Me
   End If
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Cargar_art_seleccionado
       TXTDESCRIPCION.Enabled = False
    End If
End Sub

Private Sub TXTBUSCAR_Change()
   Call BUSCAR
End Sub
Sub BUSCAR()
  
   SQLX$ = "SELECT * FROM TIEMOBRA WITH(NOLOCK)"
   If TRIM$(TXTBUSCAR) <> "" Then SQLX$ = SQLX$ + " where descripcion like " & "'*" & TXTBUSCAR.TEXT & "*'"
   SQLX$ = SQLX$ + " ORDER BY DESCRIPCION"
   '
   MSF1.Rows = 1
   I& = 1
   Set RST = READSET(SQLX$)
   With RST
      Do While Not .EOF
        I& = MSF1.Rows
        MSF1.Rows = I& + 1
        MSF1.TextMatrix(I&, 1) = SAFETEXT$(!Descripcion)
        MSF1.TextMatrix(I&, 2) = TRIM$(XVALO(!TIEMPOUNITARIO))
        MSF1.TextMatrix(I&, 3) = TRIM$(XVALO(!CantiOps))
        MSF1.TextMatrix(I&, 4) = TRIM$(XVALO(!SecciPro))
        MSF1.TextMatrix(I&, 5) = TRIM$(ECOARCH$("SECPROD", MKS$(XVALO(!SecciPro))))
      .MoveNext
      Loop
    End With
    '
    RST.Close
    Set RST = Nothing
   
    

End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
    If TXTDESCRIPCION.TEXT <> "" Then
      If KeyAscii = 13 Then
        MSF1.SetFocus
      End If
    End If
End Sub

Private Sub Cargar_art_seleccionado()
     If MSF1.Rows < 1 Then Exit Sub
     With TIEMMOB09
         Label7.MSF1.TextMatrix(MSF.Row, 2)
         .MSF1.RowHeight(MSF1.Rows - 1) = Label7.Height + 200
        .TXTDESCRIPCION = Label7
        .TXTTIEMPO = TRIM$(XVALO(MSF1.TextMatrix(MSF.Row, 2)))
        .TXTNOPER = TRIM$(XVALO(MSF1.TextMatrix(MSF.Row, 3)))
        .TXTCANTIDAD = TRIM$(XVALO(MSF1.TextMatrix(MSF.Row, 4)))
    End With
End Sub


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form BUSCA_DES 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Descripciones de Operaciones"
   ClientHeight    =   7725
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   31108.78
   ScaleMode       =   0  'User
   ScaleWidth      =   12698.98
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5280
      Top             =   3600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   19
      OleObjectBlob   =   "BUSCA_DES.frx":0000
      TabIndex        =   9
      Top             =   7440
      Width           =   11745
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Lista de Descripciones de Operaciones"
      Top             =   0
      Width           =   11791
   End
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
      Left            =   28
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   6840
      Width           =   11726
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
      Height          =   3735
      Left            =   4080
      ScaleHeight     =   3675
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   960
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
      OleObjectBlob   =   "BUSCA_DES.frx":0056
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "BUSCA_DES.frx":028A
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   6045
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Doble Click Para Seleccionar /  Boton Derecho Modificaciòn Maisva"
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10663
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
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
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
Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
End Sub

Private Sub Form_Load()
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 5
   NUCOLU% = 1: TEPRUEBA = "99": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "123456789*123456789*123456789*123456789*123456789*123456789*123456789*12": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Descripción"
   NUCOLU% = 3: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "T.Un. ''"
'   NUCOLU% = 4: TEPRUEBA = "1234567": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "QT."
   NUCOLU% = 4: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 5: TEPRUEBA = "SECCIONSECC123456789": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 3: MSF1.TextMatrix(0, NUCOLU% - 1) = "Sector"

   Call APLICACIONSKINS(Me)
   Call BUSCAR
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   i& = MSF1.MouseRow
   J& = MSF1.MouseCol
   If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    If Button = 2 Then
        '
        MSF1.Row = i&
        If MSF1.Row >= 1 Then
            With EDITBUSCAR
               .TEXTO_ORINGINAL$ = MSF1.TextMatrix(MSF1.Row, 1)
               .TIEMPOUNITARIO_ORIGINAL = MSF1.TextMatrix(MSF1.Row, 2)
               .NUMESECTOR_ORIGINAL& = MSF1.TextMatrix(MSF1.Row, 3)

               .TXTTIEMPO.TEXT = MSF1.TextMatrix(MSF1.Row, 2)
               .TXTNOPER.TEXT = MSF1.TextMatrix(MSF1.Row, 3)
               .TXTSECCION.TEXT = MSF1.TextMatrix(MSF1.Row, 4)
               .TXTDESCRIPCION.TEXT = MSF1.TextMatrix(MSF1.Row, 1)

               Call .cargalist

               .Show 1
           End With
        End If
    
    End If
   '0800_555_1616
End Sub

Private Sub Timer1_Timer()
    Call PINTATEXT(BUSCA_DES.TXTBUSCAR)
    Timer1.Enabled = False

End Sub



Private Sub TXTDESCRIPCION_Change()

End Sub

Private Sub MSF1_DblClick()
  Cargar_art_seleccionado
  On Error Resume Next
  TXTBUSCAR.SetFocus
  On Error GoTo 0
  Hide
End Sub


Private Sub MSF1_GotFocus()
   SkinLabel1.Caption = "Enter o Doble Click en la Grilla para Seleccionar Descripción -  Flechas Para Desplazarse en Grilla -  Escape para Salir  - "
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Cargar_art_seleccionado
       On Error Resume Next
       TXTBUSCAR.SetFocus
       On Error GoTo 0
       Hide
    ElseIf KeyAscii = 27 Then
       Hide
    End If
End Sub

Private Sub MSF1_LostFocus()
   SkinLabel1.Caption = "-  Escape para Salir  - "

End Sub

Private Sub TXTBUSCAR_Change()
   Call BUSCAR
   
End Sub
Sub BUSCAR()
  
   SQLX$ = "SELECT DESCRIPCION, TIEMPOUNITARIO,  SecciPro FROM TIEMOBRA WITH(NOLOCK)"
   If TRIM$(TXTBUSCAR) <> "" Then SQLX$ = SQLX$ + " where descripcion like " & "'%" & TXTBUSCAR.TEXT & "%'"
   SQLX$ = SQLX$ + " GROUP BY DESCRIPCION, TIEMPOUNITARIO, SecciPro ORDER BY DESCRIPCION"
   '
   MSF1.Rows = 1
   i& = 1
   Set Rst = READSET(SQLX$)
   With Rst
      Do While Not .EOF
        i& = MSF1.Rows
        MSF1.Rows = i& + 1
        Label7 = SAFETEXT$(!Descripcion)
        MSF1.RowHeight(i&) = Label7.Height + 100
        MSF1.TextMatrix(i&, 1) = Label7
        MSF1.TextMatrix(i&, 2) = TRIM$(XVALO(!TIEMPOUNITARIO))
'        MSF1.TextMatrix(I&, 3) = TRIM$(XVALO(!CantiOps))
        MSF1.TextMatrix(i&, 3) = TRIM$(XVALO(!SecciPro))
        MSF1.TextMatrix(i&, 4) = TRIM$(ECOARCH$("SECPROD", MKS$(XVALO(!SecciPro))))
      .MoveNext
      Loop
    End With
    '
    Rst.Close
    Set Rst = Nothing
   
    

End Sub

Private Sub TXTBUSCAR_GotFocus()
    SkinLabel1.Caption = "Enter Para Correr el Foco , Hacia la Grilla  -  Escape para Salir  - "
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        MSF1.SetFocus
      ElseIf KeyAscii = 27 Then
        Hide
      End If
End Sub

Private Sub Cargar_art_seleccionado()
     If MSF1.Rows < 1 Then Exit Sub
     With TIEMMOB09
         Label7 = MSF1.TextMatrix(MSF1.Row, 1)
        .TXTDESCRIPCION = Label7
        .TXTTIEMPO = TRIM$(XVALO(MSF1.TextMatrix(MSF1.Row, 2)))
'        .TXTCANTIDAD = TRIM$(XVALO(MSF1.TextMatrix(MSF1.Row, 3)))
        .TXTNOPER = TRIM$(XVALO(MSF1.TextMatrix(MSF1.Row, 3)))
        
    End With
End Sub

Private Sub TXTBUSCAR_LostFocus()
   SkinLabel1.Caption = " -  Escape para Salir  - "
End Sub

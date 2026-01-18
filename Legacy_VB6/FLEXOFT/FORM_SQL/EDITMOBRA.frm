VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EDITMOBRA 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar o Quitar de la Lista"
   ClientHeight    =   2220
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   13230
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
   ScaleHeight     =   8940
   ScaleMode       =   0  'User
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
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
      Height          =   1605
      Left            =   21
      TabIndex        =   6
      Top             =   0
      Width           =   13185
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
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Tiem. Unit.''"
         Top             =   270
         Width           =   1100
      End
      Begin VB.TextBox Text6 
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   270
         Width           =   1195
      End
      Begin VB.TextBox TXTMODIFICADO 
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
         Height          =   380
         Left            =   11970
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Left            =   11970
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Modificado"
         Top             =   270
         Width           =   1095
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
         TabIndex        =   17
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
         Left            =   10
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   7550
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
         Height          =   975
         Left            =   10
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   540
         Width           =   7500
      End
      Begin VB.TextBox TXTCANTIDAD 
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
         Height          =   360
         Left            =   10800
         TabIndex        =   12
         Text            =   "   "
         Top             =   550
         Width           =   1180
      End
      Begin VB.TextBox TXTCANTOP 
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
         Height          =   285
         Left            =   12075
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1230
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
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
         Left            =   12075
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Cant.Op."
         Top             =   960
         Width           =   1000
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Sector    /   N°"
         ToolTipText     =   "Doble Click Para Seleccionar Sección"
         Top             =   270
         Width           =   2160
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
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdelinminar 
      Caption         =   "Quitar de la lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "EDITMOBRA.frx":0000
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "EDITMOBRA.frx":0234
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "EDITMOBRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON


Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub

Private Sub cmdelinminar_Click()
    'QUITAR DE LA LISTA
    TIEMMOB09.MODIFICATIEMMOBRA% = 1
    Call borrafila_CON_ALTURA(TIEMMOB09.MSF1, TIEMMOB09.MSF1.Row)
    Call TIEMMOB09.ACTUALIZAR_DATOS
99  Unload Me

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
        'VALIDACION QUE NO SE INGRESEN TEXTOS QUE INCLUYAN COMILLAS SIMPLES.
        If InStr(TXTDESCRIPCION.TEXT, "'") > 0 Then
           Call COMUNI("Se Han Detectado Caracteres Reservados (Comilla Simple)\ El Mismo Serà Reemplazado Por Un Espacio Vacío")
           TXTDESCRIPCION.TEXT = REPLACAR$(TXTDESCRIPCION.TEXT, "'", " ")
        End If

    '
    TIEMMOB09.MODIFICATIEMMOBRA% = 1
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
'         .TXTDESCRIPCION.TEXT = MSF1.TextMatrix(MSF1.Row, 2)
'         .TXTNOPER.TEXT = MSF1.TextMatrix(MSF1.Row, 3)
'         .TXTSECCION.TEXT = MSF1.TextMatrix(MSF1.Row, 4)
'         .TXTCANTIDAD.TEXT = MSF1.TextMatrix(MSF1.Row, 5)
'         .TXTTIEMPO.TEXT = MSF1.TextMatrix(MSF1.Row, 6)
'         .TXTMODIFICADO.TEXT = MSF1.TextMatrix(MSF1.Row, 9)
'         .TXTCANTOP.TEXT = TXTCANTOP
         
        .MSF1.TextMatrix(.MSF1.Row, 2) = Me.TXTDESCRIPCION
        .MSF1.TextMatrix(.MSF1.Row, 3) = Me.TXTNOPER
        .MSF1.TextMatrix(.MSF1.Row, 4) = Me.TXTSECCION
         .MSF1.TextMatrix(.MSF1.Row, 5) = Me.TXTCANTIDAD
         .MSF1.TextMatrix(.MSF1.Row, 6) = Me.TXTTIEMPO
           TIME_TOTAL = XVALO(TXTTIEMPO) * XVALO(TXTCANTIDAD)
         .MSF1.TextMatrix(.MSF1.Row, 7) = TIME_TOTAL
          TIEM1 = XVALO(Me.TXTTIEMPO)
          'If Abs(XVALO(TIEM1)) > 0 Then .MSF1.TextMatrix(.MSF1.Row, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIEM1), "F10.0"))
          'jorge pidio cambiar  3600 / TIEMPO TOTLA EN VEZ DE T.UNITARIO 29-06-21
          If Abs(XVALO(TIME_TOTAL)) > 0 Then .MSF1.TextMatrix(.MSF1.Row, 8) = TRIM$(FORMATNUM$(3600 / XVALO(TIME_TOTAL), "F10.0"))
         .MSF1.TextMatrix(.MSF1.Row, 9) = FECHATEX$(CVINT(Date))
         
        
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 3, FORMATNUM$(TXTTIEMPO, "F10.0"))
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 2, TXTDESCRIPCION.TEXT)
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 4, FORMATNUM$(TXTCANTIDAD, "F10.0"))
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 5, TXTNOPER)
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 6, TRIM$(TXTSECCION))
'        TIME_TOTAL = XVALO(TXTTIEMPO) * XVALO(TXTCANTIDAD)
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 7, SAFETEXT$(TIME_TOTAL))
'        Call ACTUALIZA_FLEX(.MSF1, .MSF1.Row, 8, FECHATEX$(CVINT(Date)))
        Call TIEMMOB09.ACOMODARMSF
        Call .ACTUALIZAR_DATOS
        
    End With
    '
    Unload Me
    
99  Command1.Enabled = True

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

Function borrafila_CON_ALTURA(MSF As MSFlexGrid, REG&)

        For i& = REG& To MSF.Rows - 2
          For j& = 1 To MSF.Cols - 1
            MSF.TextMatrix(i&, j) = MSF.TextMatrix(i& + 1, j)
            MSF.Row = i&
          Next
          'PARA QUE NO SE QUEDE CON LA ALTURA ANTERIOR QUE TENIA EL REGISTRO
          'POR QUE  SE VAN MOVIENDO Y LAS ALTURAS PUEDEN SER DIFERENTES
          TIEMMOB09.Label7 = MSF.TextMatrix(i&, 2)
          MSF.RowHeight(MSF.Row) = TIEMMOB09.Label7.Height + 100
          MSF.TextMatrix(MSF.Row, 2) = TIEMMOB09.Label7
        Next
        MSF.Rows = MSF.Rows - 1
        
End Function

Private Sub Command3_Click()

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


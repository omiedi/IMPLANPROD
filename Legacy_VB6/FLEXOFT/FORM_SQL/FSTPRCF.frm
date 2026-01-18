VERSION 5.00
Begin VB.Form FSTPRCF 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formulario Registro Control Final"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6375
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
   ScaleHeight     =   2250
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   2850
      Left            =   5460
      ScaleHeight     =   2790
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "FSTPRCF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Asignar Cantidad de Ok\Tabla de Valores Admisibles de Mediciones"
         Top             =   1155
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "FSTPRCF.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   630
         Width           =   645
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "FSTPRCF.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -345
         Picture         =   "FSTPRCF.frx":059E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1770
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   22
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Impresión de Registro de Control Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   5160
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4080
         Picture         =   "FSTPRCF.frx":24C0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   21
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen de Stocks Remotos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   5160
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   945
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   4
      Top             =   6720
      Width           =   5160
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   6
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Sede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   1
      Top             =   5565
      Width           =   5160
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Identificador de Sede:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Width           =   3480
      End
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
      Height          =   855
      Left            =   5600
      Picture         =   "FSTPRCF.frx":27CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "FSTPRCF.frx":2AD4
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "FSTPRCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "RECONFI", Text4.TEXT)
    '
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Command4_Click()
    Command4.Enabled = False
    OPX1% = COMALTER%("Opciones \\Asignar Cantidad de Ok\Valores Admisibles de Mediciones\Tabla de Instrumentos")
    If OPX1% < 1 Or OPX1% > 3 Then GoTo 99
    If OPX1% = 1 Then GoTo 5 'Asignar Cantidad de Ok
    If OPX1% = 2 Then Call MEDICIONES: GoTo 99 'Tabla de Valores Admisibles de Mediciones
    If OPX1% = 3 Then Call COFUTIE07.Command7_Click: GoTo 99
    '
5   Call ABREMASTER
10  Call SELECHO("MASTER") 'Mostrar Tabla Master
    CI% = XVALO(CODINT%(VALACT1$("MASTER"))) 'Asignar a var el cod interno
    If CI% < 1 Then GoTo 99 'si es < 1 sale
    '
    CONDI$ = "Codint =" & CI% ' Compara el cod interno con el ingresado
    CANPRO% = XVALO(VALOBADA("MASTER", "CantApro", CONDI$, "NOMESS"))
    'Asignar a var el valor del campo con la condicion dada
    
    Call REPLA(VDEF$, MKI$(CI%), 1, 2) ' Formar String con los parametros y posicion
    Call REPLA(VDEF$, MKI$(CANPRO%), 3, 2)
    
    OBX$ = OBJPLA$("CantApro", VDEF$) 'Guardar informacion del objeto en var
    If OBX$ = "" Then GoTo 10  'Si no se elige nada vuelva al Selecho

    CANPRO% = CVI(Mid$(OBX$, 3, 2))
    Call ACTUREGISTRO("MASTER", "CantApro", CONDI$, CANPRO%, REGAF&, "NOMESS")
    GoTo 10
    ' Actualizar registro modificado con la condicion dada

99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    Text4.TEXT = CONTROL$("", "RECONFI")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    '
End Sub



Private Sub Text4_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_Click()
   Text4.SetFocus
End Sub
Sub MEDICIONES()
     '
     If DERACCE%("MEDICION", "Carga Valores Admisibles de Medicion") < 2 Then Exit Sub
     '
10   SQLX$ = "SELECT * FROM CF_MEDICION WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE  MARBORRA = 0 "
     SQLX$ = SQLX$ + " ORDER BY CODEXTE"
     Dim MEDTMP As ADODB.Recordset
     Set MEDTMP = READSET(SQLX$)
     '
     Call SETULTREG("!MEDICION", 0)
     JJ& = 0
     With MEDTMP
        Do While (.EOF = False)
           X$ = REGBLAN$("MEDICION")
           Call REPLA(X$, MKI(!COD_INTE), 1, 2)
           Call REPLA(X$, MKD(!MINIMO_1), 3, 8)
           Call REPLA(X$, MKD(!MAXIMO_1), 11, 8)
           Call REPLA(X$, MKD(!MINIMO_2), 19, 8)
           Call REPLA(X$, MKD(!MAXIMO_2), 27, 8)
           Call REPLA(X$, MKD(!MINIMO_3), 35, 8)
           Call REPLA(X$, MKD(!MAXIMO_3), 43, 8)
           JJ& = JJ& + 1
           Call GRAREG("!MEDICION", X$, JJ&)
            .MoveNext
         Loop
      End With
      MEDTMP.Close
      Set MEDTMP = Nothing

20    VTB% = VENTABU%("MEDICION/NC/TITUPAN=Tabla de Valores Admisibles de Medicion")

      If VTB% < 0 Then Exit Sub
      '
      CONDI$ = "COD_INTE > 0 "
      Call ACTUREGISTRO("CF_MEDICION", "MARBORRA", CONDI$, 1, REG&)
         
      SQLX$ = "SELECT * FROM CF_MEDICION "
      SQLX$ = SQLX$ + " WHERE COD_INTE < 1 "
      Set MEDITMP = New ADODB.Recordset
      MEDITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With MEDITMP
             For i& = 1 To ULTREG&("!MEDICION")
               XX$ = REGLEIDO$("!MEDICION", i&)
               CI% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !COD_INTE = CI%
               !CODEXTE = CODEXT$(CI%)
               !MINIMO_1 = CVD(Mid$(XX$, 3, 8))
               !MAXIMO_1 = CVD(Mid$(XX$, 11, 8))
               !MINIMO_2 = CVD(Mid$(XX$, 19, 8))
               !MAXIMO_2 = CVD(Mid$(XX$, 27, 8))
               !MINIMO_3 = CVD(Mid$(XX$, 35, 8))
               !MAXIMO_3 = CVD(Mid$(XX$, 43, 8))
               !MARBORRA = 0
               .Update
             Next i&
           End With
         MEDITMP.Close
         Set MEDITMP = Nothing
         '
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("CF_MEDICION", CONDI$, REG&)
         '
         SQLX$ = "SELECT COUNT(*) , COD_INTE FROM CF_MEDICION WITH(NOLOCK) WHERE MARBORRA = 0 GROUP BY COD_INTE HAVING COUNT(*) > 1"
         Set MEDTMP = READSET(SQLX$)
         CI% = XVALO(MEDTMP!COD_INTE)
         If CI% > 0 Then
           Call COMUNI("Código " & CODEXT$(CI%) & " Repetido.")
           GoTo 10
         End If
         '
         Call COMUNI("Valores Admisibles Guardados")
         '
End Sub



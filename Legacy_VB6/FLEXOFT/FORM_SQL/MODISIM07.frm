VERSION 5.00
Begin VB.Form MODISIM07 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Modificacion de SIM"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ORDITEM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   5160
      Left            =   6600
      ScaleHeight     =   5100
      ScaleWidth      =   1005
      TabIndex        =   20
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "MODISIM07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aprueba e Imprime Pedido de Cotización"
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton Command13 
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
         Height          =   650
         Left            =   130
         Picture         =   "MODISIM07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   600
      End
      Begin VB.CommandButton Command22 
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
         Height          =   650
         Left            =   130
         Picture         =   "MODISIM07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   100
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "MODISIM07.frx":075E
         Top             =   3360
         Width           =   1515
      End
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   6375
   End
   Begin VB.TextBox Text7 
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
      Left            =   960
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
      Width           =   135
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
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
      Left            =   5280
      TabIndex        =   6
      Top             =   735
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   3600
      TabIndex        =   4
      Top             =   735
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
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
      Left            =   4320
      TabIndex        =   5
      Top             =   735
      Width           =   135
   End
   Begin VB.TextBox Text3 
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
      Left            =   840
      TabIndex        =   2
      Top             =   735
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
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
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   135
   End
   Begin VB.TextBox Text1 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
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
      Left            =   2645
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
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
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   18
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
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
      Left            =   -360
      TabIndex        =   17
      Top             =   2505
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant:"
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
      Left            =   3960
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   -315
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de SIM:"
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
      Left            =   -60
      TabIndex        =   13
      Top             =   345
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden:"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   345
      Width           =   1935
   End
End
Attribute VB_Name = "MODISIM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    command1.Enabled = False
10  If ULTREG&("!INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
        Call SELECHO("!INSOLINT/Indice General de Solicitud de Materiales")
        Text1.TEXT = VALACT1$("!INSOLINT")
    Else
        GoTo 99
    End If
    'Verifica estado de la sim
    RESS% = XVALO(VALOBADA("ENCASOLINT", "STATUS", "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND Status = 9"))
    If RESS% <> 0 Then
        Call MENSERR(24, "Imposible Modificar.\SIM Anulada.")
        GoTo 10
    End If
99  command1.Enabled = True
    '
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    NUSOLI& = XVALO(Text1)
    ORSOLI& = XVALO(ORDITEM)
    If NUSOLI& < 1 Or ORSOLI& < 1 Then
        Call MENSERR(24, "Número de SIM o de Orden no Validas")
        GoTo 99
    End If
    
    CONDI$ = "NUME_SOLINT = " & NUSOLI& & " AND NUORD_ITEM = " & ORSOLI&
    If CantRegistros("DETASOLINT", CONDI$) < 1 Then
        Call MENSERR(24, "Número de SIM o de Orden no Validas")
        GoTo 99
    End If
    '
    UMEDI$ = TRIM$(Text4) ' VALIDA UNIDAD DE MEDIDA INGRESADA
    If ECOARCH$("UMEDIDA", UMEDI$) = "" Then
        Call MENSERR(24, "Unidad de Medida Inválida")
        GoTo 99
    End If
    '
    If ECOARCH$("ECUECON", Text7) = "" Then
        Call MENSERR(24, "Destino Contable no Válido.")
        GoTo 99
    End If
    '
    CODI% = CODINT(Text3.TEXT)
    Call ACTUREGISTRO("DETASOLINT", "COD_INTE", CONDI$, XVALO(CODI%), REG&)
    Call ACTUREGISTRO("DETASOLINT", "DESCRITEM", CONDI$, Text6.TEXT, REG&)
    Call ACTUREGISTRO("DETASOLINT", "UNI_SOL", CONDI$, Text4.TEXT, REG&)
      CANTI$ = TRIM$(FORMATNUM$(XVALO(Text5.TEXT), "F12.1")) ' FORMATEA EL NUMERO PARA QUE EVITAR LA COMA COMO SEPARADOR DECIMAL
    Call ACTUREGISTRO("DETASOLINT", "CANT_SOLINT", CONDI$, CANTI$, REG&)
    Call ACTUREGISTRO("DETASOLINT", "ICONTA_SOLINT", CONDI$, Text7.TEXT, REG&)
    Call ACTUREGISTRO("ENCASOLINT", "OBSE_SOLINT", "NUME_SOLINT = " & XVALO(Text1.TEXT), Text8.TEXT, REG&)
    '
    Call MENSERR(24, "Modificación Registrada.")
    '
    TTXX$ = "Modificacion de SIM" & CRLF$
    TTXX$ = TTXX$ + CRLF$
    TTXX$ = TTXX$ & "Item Numero: " & XVALO(ORDITEM) & CRLF$
    If TRIM$(Text3) <> "" Then
        TTXX$ = TTXX$ + TRIM$(Text3) & " - " & TRIM$(Text6)
    Else
        TTXX$ = TTXX$ + TRIM$(Text6)
    End If
    TTXX$ = TTXX$ + CRLF$
    TTXX$ = TTXX$ + "Cantidad: " & TRIM$(Text5) & " " & TRIM$(Text4) & CRLF$
    TTXX$ = TTXX$ + CRLF$
    TTXX$ = TTXX$ + "Observaciones: " & Text8.TEXT
    Call ANOTASIM07.ANOTASIM(XVALO(Text1.TEXT), TTXX$)
    '
99  Command14.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("UMEDIDA")
    VALONUE$ = VALACT1$("UMEDIDA")
    If VALONUE$ <> "" Then
       Text4.TEXT = VALONUE$
    End If
   '
End Sub

Private Sub Command22_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
   Call SELECHO("!ECOMASIM")
   VALONUE$ = VALACT1$("!ECOMASIM")
   If VALONUE$ <> "" Then
      Text3.TEXT = VALONUE$
   End If
End Sub

Private Sub Command4_Click()
    '
   Call SELECHO("ECUECON") 'SELECCION CUENTAS CONTABLES
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
      Text7.TEXT = VALACT1$("ECUECON")
      Label6.Caption = VALACT2$("ECUECON")
   End If
   '
End Sub



Private Sub ORDITEM_Change()
    '
    Call buscasim 'Carga codigo del articulo, descripcion y cantidad segun solicitud
    '
End Sub

Private Sub ORDITEM_Click()
    '
    Call buscasim 'Carga codigo del articulo, descripcion y cantidad segun solicitud
    '
End Sub

Private Sub Text1_Change()
    '
    ORDITEM.Clear  'LIMPIA COMBO PARA LIMPIAR FORMULARIO
    Call buscasim  'Carga codigo del articulo, descripcion y cantidad segun solicitud
    '
    XREG& = XVALO(VALOBADA("DETASOLINT", "COUNT(NUORD_ITEM)", "NUME_SOLINT = '" & Text1.TEXT & "'"))
    For JJ& = 1 To XREG&     'CARGA LOS NUMEROS DE LOS ITEMS SEGUN XREG%(REGISTROS QUE CUMPLEN CONDICION)
        ORDITEM.AddItem (JJ&)
    Next JJ&
    '
End Sub

'Private Sub Text2_Change()
'    '
'    Call buscasim 'Carga codigo del articulo, descripcion y cantidad segun solicitud
'    '
'End Sub
 Sub buscasim()
    '
    Text8.TEXT = ""
    '
    If XVALO(Text1.TEXT) > 9999999 Then
      Call MENSERR(24, "Numero de Solicitud Muy Grande")
      Exit Sub
    End If
    If XVALO(ORDITEM.TEXT) > 20 Then
      Call MENSERR(24, "Numero de Orden Muy Grande")
      Exit Sub
    End If
    '
    CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT)
    CODD% = XVALO(VALOBADA("DETASOLINT", "COD_INTE", CONDI$))
    Text3.TEXT = CODEXT(CODD%)
    DESCR = VALOBADA("DETASOLINT", "DESCRITEM", CONDI$)
    Text6.TEXT = DESCR
    CANTT = XVALO(VALOBADA("DETASOLINT", "CANT_SOLINT", CONDI$))
    Text5.TEXT = FORMATNUM(CANTT, "F10.1")
    UNIME = VALOBADA("DETASOLINT", "UNI_SOL", CONDI$)
    Text4.TEXT = UNIME
    CUECO$ = VALOBADA("DETASOLINT", "ICONTA_SOLINT", CONDI$)
    Text7.TEXT = CUECO$
    CUII% = CUEINT%(TRIM$(CUECO$))
    Label6.Caption = DENOCUE$(CUII%)
    OBSER = VALOBADA("ENCASOLINT", "OBSE_SOLINT", " NUME_SOLINT = " & XVALO(Text1.TEXT))
    Text8.TEXT = OBSER
    '
 End Sub

Private Sub Text3_Change()
  CODD$ = TRIM$(Text3)
  CI% = CODINT%(CODD$) ' asigna codigo interno
  Text6.TEXT = DESCRIT$(CI%) 'busca descripcion si el codigo existe, si no existe asiga Material No Codificado
  Text4 = UNIMED(CI%)
End Sub

Private Sub Text3_DblClick()
    Call Command3_Click
End Sub

Private Sub Text4_DblClick()
    Call Command2_Click
End Sub

Private Sub Text5_LostFocus()
    Text5.TEXT = FORMATNUM$(Text5.TEXT, "F10.1")
End Sub

Private Sub Text7_Change()
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text7.TEXT))
    If DECUE$ <> "" Then
        Label6 = DECUE$
      Else
        Label6 = ""
    End If
End Sub

Private Sub Text7_DblClick()
    Call Command4_Click
End Sub

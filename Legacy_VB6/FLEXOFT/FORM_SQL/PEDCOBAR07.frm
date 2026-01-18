VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PEDCOBAR07 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   105
      TabIndex        =   10
      Top             =   105
      Width           =   10005
      Begin VB.Line Line2 
         X1              =   9970
         X2              =   9970
         Y1              =   105
         Y2              =   420
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   105
         Y2              =   425
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"PEDCOBAR07.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   10005
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton BOTNEXT 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   10220
      Picture         =   "PEDCOBAR07.frx":0097
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar Item"
      Top             =   4470
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4110
      TabIndex        =   0
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   4750
      Width           =   4110
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4545
      Left            =   10185
      ScaleHeight     =   4485
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   -105
      Width           =   1170
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PEDCOBAR07.frx":03A1
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quitar Item de la Lista"
         Top             =   1260
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   105
         Picture         =   "PEDCOBAR07.frx":06AB
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PEDCOBAR07.frx":07F5
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      ItemData        =   "PEDCOBAR07.frx":2717
      Left            =   105
      List            =   "PEDCOBAR07.frx":2719
      TabIndex        =   1
      Top             =   555
      Width           =   10005
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PEDCOBAR07.frx":271B
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PEDCOBAR07.frx":294F
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pedido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8400
      TabIndex        =   13
      Top             =   4515
      Width           =   1590
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8350
      TabIndex        =   12
      Top             =   4750
      Width           =   1750
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   1470
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras (Artículo) >>>>>>>>:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   960
      TabIndex        =   6
      Top             =   4800
      Width           =   3150
   End
End
Attribute VB_Name = "PEDCOBAR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTNEXT_Click()
        '
        BOTNEXT.Enabled = False
        If List1.ListCount < 1 Then
           Call COMUNI("No Se Han Ingresado Artículos")
           GoTo 99
        End If
        '
        Call CARGADETPE
'        Call FORPED07.CARDETPEDI(2)
        '
        Hide
99      BOTNEXT.Enabled = True
        '
        
End Sub

Sub CARGADETPE()
        Call SETULTREG&("!CARDETPE", 0)
        For U& = 1 To List1.ListCount
            '
            CIXI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            CANTI = XVALO(Mid$(List1.List(U& - 1), 57, 10))
            PRELIX = XVALO(Mid$(List1.List(U& - 1), 69, 14))
            FE% = FECHANUM(Mid$(List1.List(U& - 1), 90, 8))
            'COMENTADO POR QUE AHORA YA SE CARGA EL PRECIO DE LISTA AL ELEGIR EL CODIGO
''            ARCHIPRE$ = TRIM$(FORPED07.TEXT21(0).TEXT)
''            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
''            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
'            PRELIX = PRELISTA#(TRIM$(FORPED07.TEXT21(0).TEXT), CIXI%)
            '
            Y$ = REGBLAN$("!CARDETPE")
            Call REPLA(Y$, MKI$(CIXI%), 1, 2)
            Call REPLA(Y$, MKS$(CANTI), 47, 4)
            Call REPLA(Y$, MKS$(PRELIX), 39, 4)
            Call REPLA(Y$, MKI(FE%), 63, 2)
            '
            Call GRAREG("!CARDETPE", Y$, U&)
            '
        Next U&
        Call SETULTREG("!CARDETPE", U& - 1)
End Sub
Private Sub Command2_Click()
  '
  Command2.Enabled = False
  Label2 = ""
  List1.Clear
  FORPED07.PORCODBARRA = 0
  Hide 'Unload Me
  FORPED07.BLANFORMU
  FORPED07.Text1 = ""
  Command2.Enabled = True
  '
End Sub
Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Text1.SetFocus
    End If
    '
    'ACTUALIZO LA VENTANA PARA QUE ACTUALICE LOS IMPORTES TOTALES
    Call CARGADETPE
    Call FORPED07.CALTOPED
    'presenta el total del pedido que trae del frm de pedidos
    PEDCOBAR07.Label6 = FORPED07.Total(2)

End Sub

Private Sub Form_Load()
    '
    Label2 = "OK"
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_Change()
    If Len(TRIM$(Text1)) = 13 Then
       Call Text1_KeyPress(13)
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
        '
        'VALIDACIONES
        If KeyAscii <> 13 Then Exit Sub
        If TRIM$(Text1) = "" Then Exit Sub
        '
        'OBTENGO CODIGO INTERNO DE EXTERNO O UPC
        CODI$ = TRIM$(Text1)
        Ci% = CODINT%(CODI$) ' CODIGO INTERNO DE CODIGO EXTERNO
        If Ci% < 1 Then
          Ci% = CODINTBARRA%(CODI$) ' CODIGO INTERNO DE UPC
        End If
        '
        'VALIDO QUE EL CODIGO SEA VALIDO
        If Ci% < 1 Then
           Call MENSERR(24, "Código Inexistente o No Válido")
           Text1.SelStart = 0
           Text1.SelLength = Len(Text1)
           GoTo 60
        End If
        '
        'LEO LA CANTIDAD DE LOTE INDIVISIBLE, SI NO TIENE CARGADO > INPUTBOX >
        'GRABO LA CANTIDAD COMO LOTE INDIVISIBLE
        CANTI = LINDIVEN(Ci%)
        If CANTI < 1 Then
5             CANTI = XVALO(InputBox$("Cantidad de Unidades", "Ingrese Cantidad de Unidades"))
          If CANTI < 1 Then
             Call COMUNI("Debe Ingresar una Cantidad Válida")
             GoTo 5
          End If
          Call ACTUREGISTRO("MASTER", "LINDIVEN", "CODINT = " & Ci%, FORMATNUM$(CANTI, "F16." & CNTDC$(Ci%)), REGAF&, "NOMESS")
        End If
        '
        'ARMO LA FILA
        CODI2$ = TRIM$(CODEXT$(Ci%))
        Call REPLA(TX$, CODI2$, 1, 16)
        Call REPLA(TX$, Left(DESCRIT$(Ci%), 30), 18, 36)
        Call REPLA(TX$, FORMATNUM$(CANTI, "F10.1"), 57, 10)
        Call REPLA(TX$, FORMATNUM$(PRELISTA#(FORPED07.TEXT21(0), Ci%), "F14.4"), 69, 14)
        '
        'COMPARAR SI EXISTE EN LA LISTA, EN ESE CASO ACUMULO SOBRE EL MISMO
        For U& = 1 To List1.ListCount
          If TRIM$(Left$(List1.List(U& - 1), 16)) = CODI2$ Then
             CANTLIST = XVALO(Mid$(List1.List(U& - 1), 57, 10))
             FES$ = Mid$(List1.List(U& - 1), 90, 8)
             Call REPLA(TX$, FES$, 90, 8)
             Call REPLA(TX$, FORMATNUM$((CANTLIST + CANTI), "F10.1"), 57, 10)
             List1.RemoveItem U& - 1
             List1.AddItem TX$, U& - 1
             GoTo 50
          End If
        Next U&
        '
        List1.AddItem TX$
        '
50      Text1 = ""
60      Text1.SetFocus
        Call CARGADETPE
        Call FORPED07.CALTOPED
        'presenta el total del pedido que trae del frm de pedidos
        PEDCOBAR07.Label6 = FORPED07.Total(2)
        '
   End Sub

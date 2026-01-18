VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CIERORFA 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de O-Fabricación"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Modo2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Modo1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   21
      Top             =   4560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      ItemData        =   "CIERORFA.frx":0000
      Left            =   105
      List            =   "CIERORFA.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   3780
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton BOTNEXT 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5040
      Picture         =   "CIERORFA.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cargar Items"
      Top             =   5160
      Width           =   855
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
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   4440
      Left            =   5040
      ScaleHeight     =   4380
      ScaleWidth      =   870
      TabIndex        =   3
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "CIERORFA.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Quitar  Producto de la Lista"
         Top             =   1440
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "CIERORFA.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CIERORFA.frx":0762
         Top             =   3990
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
      Height          =   960
      ItemData        =   "CIERORFA.frx":2684
      Left            =   120
      List            =   "CIERORFA.frx":2686
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1995
      Width           =   4815
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CIERORFA.frx":2688
      Top             =   -120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   -360
      OleObjectBlob   =   "CIERORFA.frx":28BC
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      TabIndex        =   20
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas Leídas:"
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
      Left            =   3000
      TabIndex        =   19
      Top             =   1755
      Width           =   1125
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pendientes de Cierre"
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
      Left            =   105
      TabIndex        =   18
      Top             =   3540
      Width           =   2400
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas a Cerrar"
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
      Left            =   105
      TabIndex        =   17
      Top             =   1750
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Conjunto"
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
      TabIndex        =   15
      Top             =   550
      Width           =   1965
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "O-Fabricación"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   200
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   12
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta de OF >>>>>"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1305
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta de Caja >>>>>"
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
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   2085
   End
End
Attribute VB_Name = "CIERORFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTNEXT_Click()
        '
        BOTNEXT.Enabled = False
        '
        If List1.ListCount < 1 Then
           Call COMUNI("No Se Ha Leido Ninguna Etiqueta")
           GoTo 99
        End If
        '
        'PREGUNTO SI LA CANTIDAD DE LA ULTIMA CAJA ES MULTIPLO DE LOTE INDIVISIBLE
        NORFA$ = TRIM$(Label3)
        CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
        LOTIND# = LINDIVEN(CIXI%)
        '
10      ULTCAJA# = XVALO(InputBox("Cantidad de Ultima Caja", "Cantidad de la Ultima Caja", CStr(LOTIND#)))
        If ULTCAJA# - LOTIND# > 0 Then
            Call MENSERR(24, "La Cantidad Ingresada No Puede Ser Mayor al Lote Indivisible")
            GoTo 10
        ElseIf Abs(ULTCAJA# - LOTIND#) < 0.05 Then
            ULTCAJA# = 0
        End If
        If ULTCAJA# > 0.05 Then
            TXT$ = TRIM$(List1.List(List1.ListCount - 1))
            TXTX$ = Space$(128)
            Call REPLA(TXTX$, TXT$, 1, 64)
            Call REPLA(TXTX$, str$(ULTCAJA#), 64, 12)
            List1.List(List1.ListCount - 1) = TXTX$
        End If
        '
        Option1.Value = True
        Hide
99      BOTNEXT.Enabled = True
        '
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  List1.Clear
  Hide
  Command2.Enabled = True
  '
End Sub
Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List2.AddItem List1.List(List1.ListIndex)
      List1.RemoveItem (List1.ListIndex)
      QETI% = XVALO(Label11): QETI% = QETI% - 1
      Label11 = CStr(QETI%)
      Text1.SetFocus
    End If
    '
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_Change()
  LGCB% = 13: If Option2.Value Then LGCB% = 12
  If Len(Text1) > LGCB% Then
        '
        If TRIM$(Text1) = "" Then Exit Sub
        '
        'OBTENGO NUMERO DE ORDEN DE FABRICACION Y ORDEN DE ETIQUETA
        XX$ = Text1
        NORFA$ = Mid$(XX$, 1, 2) & "-" & Mid$(XX$, 3, 6) & "-" & Mid$(XX$, 9, 1)
        NORFA$ = UCase(NORFA$)
        NORFA0$ = TRIM$(Label3)
        ORDEN% = XVALO(Mid$(XX$, 11))
        '
        'VALIDA QUE LA ETIQUETA HAYA INGRESADO POR RECEPCION DE MATERIALES
        ARMA% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IDSUBOR='" & NORFA$ & "'"))
        LORECE$ = SAFETEXT$(VALOBADA("ORFAETI", "LOTERECE", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & CStr(ORDEN%)))
'        If TRIM$(LORECE$) = "" And ARMA% > 0 Then
'           Call COMUNI("Imposible Cerrar. La Caja No Ha Sido Ingresada Desde la Recepción de Materiales")
'           Text1.SelStart = 0
'           Text1.SelLength = Len(Text1)
'           GoTo 60
'        End If
        '
        'VALIDA QUE LA OF DE LA ETIQUETA SE CORRESPONDA CON EL FORM DE OF
        If NORFA$ <> NORFA0$ Then
           Call COMUNI("La Etiqueta No Corresponde a la O-Fabricacón Elegida")
           Text1.SelStart = 0
           Text1.SelLength = Len(Text1)
           GoTo 60
        End If
        '
        STAT% = XVALO(VALOBADA("ORFAETI", "STATUS", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & CStr(ORDEN%)))
        If STAT% > 0 Then
            Call COMUNI("La Etiqueta Leída Ya Fue Registrada")
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1)
            GoTo 60
        End If
        '
        For II& = 0 To List1.ListCount - 1
            TX$ = List1.List(II&)
            ORLI% = XVALO(Mid$(TX$, 19))
            If ORLI% = ORDEN% Then
                Call COMUNI("La Etiqueta Leída Ya Se Encuentra en la Lista")
                Text1.SelStart = 0
                Text1.SelLength = Len(Text1)
                GoTo 60
            End If
        Next II&
        '
        NO$ = CStr(ORDEN%)
        While Len(NO$) < 4: NO$ = "0" & NO$: Wend
        TTXX$ = "Lectura Caja Nro. " & NO$
        List1.AddItem TTXX$
        QETI% = XVALO(Label11): QETI% = QETI% + 1
        Label11 = CStr(QETI%)
        '
        For AX% = 0 To List2.ListCount - 1
            TX$ = List2.List(AX%)
            If TRIM$(TX$) = TRIM$(TTXX$) Then
                List2.RemoveItem (AX%)
                Exit For
            End If
        Next AX%
        '
50      Text1 = ""
60      Text1.SetFocus
        '
    End If
End Sub

Private Sub Text2_Change()
    If Len(Text2) > 8 Then
        '
        If List1.ListCount > 0 Then
            Text2 = "": Text1.SetFocus: GoTo 99
        Else
            List2.Clear
        End If
        '
        'OBTENGO NUMERO DE ORDEN DE FABRICACION Y ORDEN DE ETIQUETA
        XX$ = Text2
        NORFA$ = Mid$(XX$, 1, 2) & "-" & Mid$(XX$, 3, 6) & "-" & Mid$(XX$, 9, 1)
        NORFA$ = UCase(NORFA$)
        '
        'VERIFICA SI LA ORDEN DE FABRICACION EXISTE (CORRESPONDE A LA EMPRESA)
        If CantRegistros("ORDEFABR", "IDSUBOR='" & NORFA$ & "'", "NOMESS") < 1 Then
           Call COMUNI("Orden de Fabricación No Encontrada")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        End If
        '
        'VERIFICA STATUS DE LA ORDEN DE FABRICACION
        STAT% = XVALO(VALOBADA("ORDEFABR", "StatuOrf", "IDSUBOR='" & NORFA$ & "'"))
        If STAT% = 3 Then
           Call MENSERR(24, "Orden de Fabricación Cerrada")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        ElseIf STAT% > 3 Then
           Call MENSERR(24, "Orden de Fabricación Anulada")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        End If
        '
        'VERIFICA SI LA O-F SE HA CERRADO
        CONDI$ = "IDSUBOR='" & NORFA$ & "' AND STATUS <1"
        If CantRegistros("ORFAETI", CONDI$) < 1 Then
           Call MENSERR(24, "La O-F Leída Se Ha Cerrado Completamente")
           Text2.SelStart = 0
           Text2.SelLength = Len(Text2)
           GoTo 60
        End If
        '
        CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
        Label3.Caption = NORFA$
        Label4.Caption = CODEXT$(CIXI%)
        Label5.Caption = DESCRIT0$(CIXI%)
        '
        SQLX$ = "SELECT NUORDEN FROM ORFAETI WITH(NOLOCK) WHERE IDSUBOR='" & NORFA$ & "' AND STATUS < 3 ORDER BY NUORDEN ASC"
        Set Rst = READSET(SQLX$)
        With Rst
            Do While Not .EOF
                ORDEN% = XVALO(!NUORDEN)
                NO$ = CStr(ORDEN%)
                While Len(NO$) < 4: NO$ = "0" & NO$: Wend
                TTXX$ = "Lectura Caja Nro. " & NO$
                List2.AddItem TTXX$
                .MoveNext
            Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        Label11.Caption = "0"
        '
50      Text2 = ""
        Text1.SetFocus: GoTo 99
60      Text2.SetFocus
        '
    End If
99      Exit Sub
End Sub


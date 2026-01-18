VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CARTEMBAR07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta de Embarque Toyota"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   285
      Left            =   1365
      TabIndex        =   8
      Top             =   210
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8760
      TabIndex        =   7
      Top             =   7035
      Width           =   1380
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6210
      Left            =   210
      TabIndex        =   6
      Top             =   735
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   10954
      _Version        =   393216
      Rows            =   16
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CARTEMBAR07.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CARTEMBAR07.frx":0234
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de O/S:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -525
      TabIndex        =   5
      Top             =   260
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1365
      TabIndex        =   4
      Top             =   210
      Width           =   1635
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Entrega:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4095
      TabIndex        =   3
      Top             =   260
      Width           =   1725
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5835
      TabIndex        =   2
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6615
      TabIndex        =   1
      Top             =   255
      Width           =   2100
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8865
      TabIndex        =   0
      Top             =   210
      Width           =   1170
   End
End
Attribute VB_Name = "CARTEMBAR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   If TRIM$(Label1.Caption) = "" Then
      Call MENSERR(24, "Imposible Guardar - Falta Numero de O/S")
      Text1.Visible = False
      GoTo 99
   End If
   '
   ' valida los codigos
   For i& = 1 To GRID.Rows - 1
     GRID.Row = i&
     GRID.COL = 1
     Codigo$ = TRIM$(GRID.TEXT)
     If Codigo$ <> "" Then
       CIXI% = XVALO(VALOBADA("MASTER", "Codint", "Codigo LIKE '" + Codigo$ + "%' And Status > 0"))
       If CIXI% > 0 Then GoTo 19
     End If
     '
     Call MENSERR(24, "Atención: Código '" + Codigo$ + "'\no Registrado en Base de Datos FLEXOFT.\   \Imposible Guardar Carta de Embarque - Consulte.")
     Text1.Visible = False
     GoTo 99
19 Next i&
   Text1.Visible = False
   '
   condi0$ = "Num_Os = '" & Label1 & "' AND Num_Entre = '" & Label2 & "' AND FechEntre = '" & Label3 & "'"
   YAESTA% = CantRegistros("ENCACARTEM", condi0$)
   If YAESTA% > 0 Then
      TTXX$ = "Carta de Embarque ya Registrada"
      CONDI$ = condi0$ + " AND Dopricor <> ''"
      CANREPRO% = CantRegistros("DETACARTEM", CONDI$)
      If CANREPRO% > 0 Then
         NOREMI$ = SAFETEXT$(VALOBADA("DETACARTEM", "DOPRICOR", CONDI$))
         TTXX$ = TTXX$ + "\y Despachada por Remito '" + NOREMI$ + "'.\  \Imposible Refrescar esta Carta de Embarque.\   \Anule Remito y Repita Petición."
         Call MENSERR(24, TTXX$)
         GoTo 99
      End If
      TTXX$ = TTXX$ + ".\  \¿ Desea Sobre-Escribir esta Carta de Embarque ?\\No, Mantener Anterior\Si, Sobre-Escribir"
      If COMALTER%(TTXX$) <> 2 Then GoTo 99
      FLEXCONN.Execute ("DELETE ENCACARTEM WHERE " & condi0$)
      FLEXCONN.Execute ("DELETE DETACARTEM WHERE " & condi0$)
   End If
   '
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM ENCACARTEM WHERE " & condi0$
   Dim ENCACAR As ADODB.Recordset
   Set ENCACAR = New ADODB.Recordset
   ENCACAR.Open SQLX$, FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   With ENCACAR
     .AddNew
     !Num_Os = Label1
     !Num_Entre = Label2
     !FECHENTRE = Label3
     .Update
   End With
   ENCACAR.Close
   Set ENCACAR = Nothing
   '
   SQLX$ = "SELECT * FROM DETACARTEM WHERE " & condi0$
   Dim DETACAR As ADODB.Recordset
   Set DETACAR = New ADODB.Recordset
   DETACAR.Open SQLX$, FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   With DETACAR
     For i& = 1 To GRID.Rows - 1
       GRID.Row = i&
       GRID.COL = 1
       Codigo$ = GRID.TEXT
       CIXI% = XVALO(VALOBADA("MASTER", "Codint", "Codigo LIKE '" + Codigo$ + "%' And Status > 0"))
       '
       If CIXI% < 1 Then
          Call MENSERR(24, "Atención: Código '" + Codigo$ + "' no Registrado en Base de Datos FLEXOFT.\Resultados Imprevisibles - Consulte.")
       End If
       '
       GRID.COL = 6    ' cantidad
       Cantidad = XVALO(GRID.TEXT)
       If Cantidad < 0.5 Then
          Call COMUNI("Codigo '" + Codigo$ + "' - Cantidad en Cero ha Sido Suprimido.")
          GoTo 29
       End If
       '
       .AddNew
       !Num_Os = Label1
       !Num_Entre = Label2
       !FECHENTRE = Label3
       !NorItCa = i&
       !Cod_Parte = Codigo$
          '
          GRID.COL = 2
       !SubCod = GRID.TEXT
          GRID.COL = 4
       !Num_Inter = XVALO(GRID.TEXT)
          GRID.COL = 5
       !Pie_Medio = XVALO(GRID.TEXT)
          GRID.COL = 6
       !Pie_Pedid = XVALO(GRID.TEXT)
       !Pie_Recu = 0
       !Pie_Tota = XVALO(GRID.TEXT)
       !Pie_Enviar = XVALO(GRID.TEXT)
       '
       !Cod_Exte = CODEXT(CIXI%)
       !cod_inte = CIXI%
       !DoPriCor = ""
       !DOPRILAR = ""
       !NuOrItem = 0
       !FECHEMIS = CVDAT(0)
       .Update
29   Next i&
   End With
   DETACAR.Close
   Set DETACAR = Nothing
   Screen.MousePointer = 1
   '
   Call COMUNI("Carta de Embarque '" + Label1 + " - " + Label2 + " del " + Label3 + "'\Registrada en la Base de Datos FLEXOFT.")
   '
99 Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)

End Sub

Private Sub Text1_Change()
    If Text1.Visible = True Then
       Label1 = Text1
    End If
End Sub

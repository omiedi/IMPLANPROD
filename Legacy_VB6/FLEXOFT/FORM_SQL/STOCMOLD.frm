VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form STOCMOLD 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proección  de Entregas para itmes de Pedidos actual"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Ver Situación de Pedidos y Proyección"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   10440
      TabIndex        =   10
      Top             =   1680
      Width           =   3435
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Según Pedidos Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   13875
      Begin VB.Label TXTDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8880
         TabIndex        =   22
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   7800
         TabIndex        =   21
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label TXTPedidosPendientes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   20
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos Pendientes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label TXTStockActual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label TXTExcedenteof 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente de O.F.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label TXTFinExcedenteof 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8880
         TabIndex        =   13
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de  Finalización Para Completar los Pedidos Pendientes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3480
         TabIndex        =   12
         Top             =   600
         Width           =   5325
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00808080&
      Height          =   7275
      Left            =   14160
      ScaleHeight     =   7215
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   70
         Picture         =   "STOCMOLD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Aceptar"
         Top             =   6240
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "STOCMOLD.frx":030A
         Top             =   6840
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pedido Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      Begin VB.Label LBLUNIMED 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   13280
         TabIndex        =   33
         Top             =   240
         Width           =   525
      End
      Begin VB.Label TXTFECHASECO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11760
         TabIndex        =   32
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F. de Entrega Seco:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   10080
         TabIndex        =   31
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label TxtCantidadU 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8520
         TabIndex        =   29
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. en Unidades:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   11
         Left            =   7680
         TabIndex        =   28
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total A Fabricar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   10080
         TabIndex        =   27
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label TXTSaldoaFabricar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11760
         TabIndex        =   26
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Fabric.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   10080
         TabIndex        =   25
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label TXTFinSaldoaFabricar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11760
         TabIndex        =   24
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitada (U.Med. Original):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label TXTCantidadOriginal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label TXTDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   6300
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   645
      End
      Begin VB.Label TXTCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   2220
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7800
      OleObjectBlob   =   "STOCMOLD.frx":222C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "STOCMOLD.frx":2460
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Click Para Activar Código y Ver Detalle /  Doble Click Para Modificar Detalle"
      Top             =   3000
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Pedidio Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   2685
   End
End
Attribute VB_Name = "STOCMOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Command19_Click()
   Unload Me
End Sub

 Sub Codigo_Change()
   Codigo$ = TRIM$(TXTCODIGO)
   CI1% = CODINT(Codigo$)
   If CI1% < 1 Then Exit Sub
   '
     STACT = STOCKACT(CI1%)
   Me.TXTStockActual = TRIM$(FORMATNUM$(STOCKACT(CI1%), "F12.1"))
   eXCEDENfAB = ExcedentedeFabrica(CI1%)
   Me.TXTExcedenteof = TRIM$(FORMATNUM$(eXCEDENfAB, "F12.1"))
     PEDIPENDIEN = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
   Me.TXTPedidosPendientes = TRIM$(FORMATNUM$(PEDIPENDIEN, "F12.1"))
      DISPONIBLEX = STACT - PEDIPENDIEN: If DISPONIBLEX < 0 Then DISPONIBLEX = 0
   Me.TXTDisponible = DISPONIBLEX
'     AFABRICAR = CANTISOLICITADA - (STACT + PEDIPENDIEN): If AFABRICAR < 0 Then AFABRICAR = 0
'   Me.TXTSaldoaFabricar = TRIM$(FORMATNUM$(AFABRICAR, "F12.1"))
   'aca calcular fecha de finalizacion de excedente de fabricacion
   Demoradeexcedente% = DemoraenDiasExcedenteFabricaciondeItem%(CI1%) ' ya devuelve la cantidad real de dias considerando los sab y dom y feriados
   DIX% = DIASPOST(HOY(HO$), Demoradeexcedente%)
85 If DiaNoLaborable(DIX%) > 0 Then
      DIX% = DIASPOST(DIX%, 1)
      GoTo 85
   End If
   Me.TXTFinExcedenteof = FECHATEX$(DIX%)
'  demorareal% = CantDiasDemoraReal%(HOY(HO$), Demoradeexcedente%)
   '
   'ACA CALCULAR PARA QUE FECHA ESTARIA LA PARTE QUE  HAY QUE FABRICAR.
   '
   UNME$ = Unimed$(CI1%)
   Me.LBLUNIMED = UNME$
    '
   PEDPENDI = XVALO(TXTPedidosPendientes)
'   If UNME$ = "M2" Then
'      PEDPENDI = XVALO(TXTPedidosPendientes) * 6
'   End If
   '
   AFABRIC = PEDPENDI + XVALO(Me.TXTCantidadOriginal) - (XVALO(TXTStockActual) + XVALO(TXTExcedenteof))
   If AFABRIC < 0 Then AFABRIC = 0
   Me.TXTSaldoaFabricar = AFABRIC
   '
   Me.TXTFinSaldoaFabricar = ""
   If AFABRIC > 0 Then
     Me.TXTFinSaldoaFabricar = FECHATEX$(DIASPOST(HOY(HO$), DemoraenDiasFabricaciondeItem%(CI1%, XVALO(Me.TXTSaldoaFabricar))))
     Me.TXTFECHASECO = FECHATEX$(DIASPOST(CVINT(Me.TXTFinSaldoaFabricar), XVALO(Control("CONFABPE", "DIASSECA"))))
   End If
   '
   
End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
    DIAGPREM.Show 1
    
End Sub

Private Sub MSF1_Click()
 REG& = MSF1.MouseRow
 COL& = MSF1.MouseCol
 If REG& > 0 Then
    Me.TXTCantidadOriginal = MSF1.TextMatrix(REG&, 3)
    Me.TxtCantidadU = MSF1.TextMatrix(REG&, 4)
    Me.TXTCODIGO = MSF1.TextMatrix(REG&, 1)
 End If
End Sub

Sub TXTCantidad_Click()

End Sub

Private Sub MSF1_DblClick()
   Call FORPED07.List1_DblClick
End Sub

Sub TXTCODIGO_Change()
   Me.TXTDESCRIPCION = DESCRIT$(CODINT%(Me.TXTCODIGO))
   Call Codigo_Change
End Sub

Private Sub TXTDisponible_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "(STOCK ACTUAL DEL CODIGO) - (PEDIDOS PENDIENTES DEL CODIGO)  " & vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub TXTExcedenteof_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "(TOTAL DE FABRICACION PENDIENTE DEL CODIGO) - (PEDIDOS PENDIENTES DEL CODIGO)  " & vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub TXTFinExcedenteof_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "PRESENTA LA FECHA DE FINALIZACION DE FABRICACION DEL EXCEDENTE DE FABRICACION" & vbCrLf
       TXTO$ = TXTO$ + "CONSIDERANDO CUAL ES EL ULTIMO PEDIDO PARA ENTREGAR DE DICHO ARTICULO." & vbCrLf
       TXTO$ = TXTO$ + "LUEGO CALCULA LA CANTIDAD PENDIENTE  DE TODOS LOS ITEM QUE COMPARTEN EL MISMO MOLDE QUE EL ITEM ACTIVO. " & vbCrLf
       TXTO$ = TXTO$ + "EN BASE A ESTA CANTIDAD CALCULA PARA QUE FECHA TENDRA LA CANTIDAD EXCEDENTE DE FABRICACION"
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub


Private Sub TXTSaldoaFabricar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "(PedidosPendientes + Cantidad Solicitada) - (StockActual + Excedente de Fabricación)" & vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub



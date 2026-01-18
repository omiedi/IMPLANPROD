VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ASIGLOT07 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Lotes "
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   8400
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   6840
         TabIndex        =   20
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5950
         TabIndex        =   19
         Top             =   420
         Width           =   810
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   3225
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -1000
         TabIndex        =   15
         Top             =   420
         Width           =   1770
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "LOTE A ASIGNAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   8400
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Height          =   2150
         Left            =   5400
         TabIndex        =   23
         Top             =   150
         Width           =   1815
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Falta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label Label10 
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
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label6 
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
            Height          =   285
            Left            =   240
            TabIndex        =   26
            Top             =   1080
            Width           =   1320
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label4 
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
            Height          =   285
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   1320
         End
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   7380
         Picture         =   "ASIGLOT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Asigna Lotes Seleccionados"
         Top             =   360
         Width           =   810
      End
      Begin VB.Frame Frame7 
         Caption         =   "                "
         Height          =   1800
         Left            =   2760
         TabIndex        =   21
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   3960
         TabIndex        =   10
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "ASIGLOT07.frx":030A
         Left            =   195
         List            =   "ASIGLOT07.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   600
         Width           =   5025
      End
      Begin VB.Label VALIDA 
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Lote                       F.Alta       Ancho(mm)            Kilos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   195
         TabIndex        =   13
         Top             =   360
         Width           =   5025
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CODIGOS TRAZABLES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8400
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6960
         TabIndex        =   8
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   1800
         Left            =   2000
         TabIndex        =   6
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   5
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1644
         ItemData        =   "ASIGLOT07.frx":030E
         Left            =   190
         List            =   "ASIGLOT07.frx":0310
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                   Cant.Sol.    Cant. Asig."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   195
         TabIndex        =   4
         Top             =   360
         Width           =   8025
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8400
      Left            =   8640
      ScaleHeight     =   8370
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   -20
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   33
         Top             =   5760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
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
         Height          =   690
         Left            =   120
         Picture         =   "ASIGLOT07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Aprueba Lotes Asignados"
         Top             =   1610
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Left            =   120
         Picture         =   "ASIGLOT07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
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
         Left            =   120
         Picture         =   "ASIGLOT07.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   920
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ASIGLOT07.frx":0A70
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ASIGLOT07.frx":0CA4
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "ASIGLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CALCSAL#()
Public DEPOEXP_REMITO%


Private Sub Command11_Click()
    '
    List1.Clear
    List2.Clear
    VALIDA = "RECHA"
    Unload Me
    '
End Sub

Private Sub Command4_Click()
    '
    VALIDA = "APRO"
    For i = 0 To List1.ListCount - 1
        XX$ = List1.List(i)
        LOTE$ = TRIM$(Mid$(XX$, 85))
        If LOTE$ = "" Then
            Call MENSERR(24, "Todos los Articulos Deben Tener un Lote Asignado")
            Exit Sub
        End If
    Next i
    
    Label8 = ""
    Label1 = ""
    Label3 = ""
    DEPOEXP_REMITO% = 0
    Hide
    '
End Sub

 Sub Command6_Click()
   '
   Command6.Enabled = False

   NotaCredito% = 6
   NFACTUR$ = TRIM$(Left$(VALACT2$("SELFCND"), 18)) ' FACTURA SELECCIONADA
   
   XX$ = List1.TEXT
   For KKU& = 1 To List2.ListCount
      If List2.Selected(KKU& - 1) = True Then
         If NFACTUR$ <> "" Then
            If FORFASQL.TIFADOC% = NotaCredito% Then GoTo 60
         End If
         '
         For II& = 0 To List1.ListCount - 1
            YY$ = List1.List(II&)
            If InStr(85, YY$, TRIM$(Left$(List2.List(KKU& - 1), 12)) & ";") > 0 Then
                Call MENSERR(24, "Lote ya Asignado")
                Label6 = "" 'FORMATNUM$(KILOO, "F10.1")
                Label10 = "" 'FORMATNUM$(XVALO(Label4) - XVALO(Label10), "F10.1")
                GoTo 99
            End If
60       Next II&
         LOOTEE$ = TRIM$(LOOTEE$) & TRIM$(Left$(List2.List(KKU& - 1), 12)) & ";"
         KILOO = KILOO + XVALO(Mid$(List2.List(KKU& - 1), 35, 10))
         Label6 = "" 'FORMATNUM$(KILOO, "F10.1")
         Label10 = "" 'FORMATNUM$(XVALO(Label4) - XVALO(Label10), "F10.1")
         Label4 = ""
      End If
   Next KKU&
'   Call REPLA(XX$, LOOTEE$, 85, 120)
   Call REPLA(XX$, LOOTEE$, 85, 390) 'CAMBIO REALIZADO PARA CONTEPLAST Y CAMBACERES QUE POR CADA BOBINA ES UN LOTE
   'AGREGADO PARA QUE ENTREGAR LO QUE SE VA A DESPACHAR Y NO EL TOTAL DE LOS LOTES ELEGIDOS
   CANTASI = XVALO(Mid$(XX$, 54, 12))
   If KILOO > CANTASI Then KILOO = CANTASI
   Call REPLA(XX$, FORMATNUM$(KILOO, "F10.1"), 64, 10)
   List1.List(List1.ListIndex) = XX$
   '
99 Command6.Enabled = True
   '
End Sub

Private Sub Form_Load()
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  '
  DF = DEPOEXP_REMITO%
End Sub

Private Sub Label8_Change()
    '
    List2.Clear
    If TRIM$(Label8) = "" Then Exit Sub
    CCII% = CODINT%(Label8)
    CADECI = FORMATNUM$(FRACMIN(CCII%), "F10." & CNTDC$ & "")
    'SOLUCION DADA POR MEDITEA EL CASO QUE SEA NOTA DE CREDITO Y ANULA FACTURA
    'EN ESTA RUTINA MUESTRA LOS LOTES SOLO DE LA FACTURA EN CUESTION
    NFACTUR$ = TRIM$(Left$(VALACT2$("SELFCND"), 18)) ' FACTURA SELECCIONADA
    If NFACTUR$ <> "" Then
        If FORFASQL.TIFADOC% = 6 Then
           Call LOTES_FACTURA(CCII%, XVALO(FORFASQL.Text1), NFACTUR$)
           Exit Sub
        End If
    End If
    SQLX$ = "SELECT IdenLote, CanAlta, FechAlta , AnchoMili,CANSALDO FROM LOTINGRE WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE COD_INTE = " & CCII%
    'SQLX$ = SQLX$ + " AND CANSALDO > 0.05"
    SQLX$ = SQLX$ + " AND (CANALTA - CANCONSU) > " & CADECI
'    Dim PARDIATMP As ADODB.Recordset
'    Set PARDIATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set PARDIATMP = READSET(SQLX$)
    With PARDIATMP
        '
        If .EOF = True Then
            Call MENSERR(24, "No Existen Lotes para el Articulo Seleccionado")
        End If
        '
        'SI SE LE PASA DEPOSITO DE CONSUMO EN LA VARIABLE PUBLICA DE FRM 'DEPOEXP_REMITO%' TOMA ESTE
        If DEPOEXP_REMITO% > 0 Then
           DEPO% = DEPOEXP_REMITO%
        Else ' SI NO EL  CONFIGURADO EN REMITO
           DEPO% = XVALO(Control$("", "DEPOEXPE"))
        End If
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        
        'SE PASA EL 44 PARA MEDITEA POR QUE VA A REVISAR EL 44 Y NO EL DE FACTURACION P Q LUEGO TRANSFIERE DEL 44 AL 3
        DEPOPANFACT% = XVALO(Control$("", "DPREPPAÑ"))
        If InStr(EPAC$, "MEDITEA") > 0 And DEPOPANFACT% > 0 Then
           DEPO% = DEPOPANFACT%
        End If
        
        '
        Do While Not PARDIATMP.EOF
            '
30          TTXX$ = ""
            LOTE$ = SAFETEXT(!idenlote)
            Call REPLA(TTXX$, LOTE$, 1, 12)
            'MODIFICADO, SI TIENE CONFIGURADO DEPOSITO THEN
            'MsgBox "Dep Expedicion= " & CStr(DEPOEXP_REMITO%) & vbCrLf & "Dep.= " & CStr(DEPO%) & vbCrLf & "Dep. Factura= " & CStr(DEPOPANFACT%)
            If DEPO% > 0 Then
               CANTI = SaldoLote(CCII%, LOTE$, DEPO%)
            Else ' SI NO TRAE LA CANTIDAD DEL LOTINGRE
               'CANTI = XVALO(!CanAlta)
               CANTI = XVALO(!CanSaldo) ' CAMBIE A CANSALDO  POR QUE CANALTA NUNCA SE ACTUALIZA
            End If
            '
            Call REPLA(TTXX$, FECHATEX$(CVINT%(!FechAlta)), 15, 8)
            Call REPLA(TTXX$, FORMATNUM$(XVALO(!AnchoMili), "F10." & CNTDC$ & ""), 23, 10)
            Call REPLA(TTXX$, FORMATNUM$(CANTI, "F10." & CNTDC$ & ""), 33, 10)
            If CANTI > 0 Then List2.AddItem (TTXX$)
            '
50          PARDIATMP.MoveNext
        '
        Loop
    End With
    PARDIATMP.Close
    Set PARDIATMP = Nothing
    
End Sub

Private Sub List1_DblClick()
    '
    Z$ = List1.TEXT
    CCII% = CODINT%(Mid$(Z$, 1, 16))
    Label8 = ""
    Label8 = CODEXT(CCII%)
    Label1 = DESCRIT$(CCII%)
    Label3 = TRIM$(Mid$(Z$, 56, 10))
    Label4 = FORMATNUM$(XVALO(Mid$(Z$, 56, 10)), "F10.1")
    Label10 = Label4
    Label6 = ""
    List1.ToolTipText = Mid$(Z$, 85, 120)
    '
End Sub

Private Sub List2_Click()
 TXTX$ = List2.TEXT
 If Left$(TXTX$, 3) = "OF-" Then
    CONDI$ = "IDSUBOR = '" & Left$(TXTX$, 11) & "'"
    OCOMP$ = VALOBADA("ORDEFABR", "ID_ORFA", CONDI$, "NOMESS")
    List2.ToolTipText = OCOMP$
 End If
End Sub

Private Sub List2_ItemCheck(Item As Integer)
    Label10 = Label4
    Label6 = ""
    Contador% = 0
    ITEMSELECCIONADO% = Item
   For KKU& = 1 To List2.ListCount
      If List2.Selected(KKU& - 1) = True Then
         Contador% = Contador% + 1
         If Contador% >= 30 Then
            Call COMUNI("Atención !!!! El Máximo de Lotes a Seleccionar Por Fila es de 30\En Caso que No Pueda Completar La Cantidad\Cancele esta Pantalla y Corrija")
             List2.Selected(ITEMSELECCIONADO%) = False
             Exit For
         End If
         KILOO = KILOO + XVALO(Mid$(List2.List(KKU& - 1), 35, 10))
         Label6 = FORMATNUM$(KILOO, "F10.1")
         FALTTAA = FORMATNUM$(XVALO(Label4) - XVALO(Label6), "F10.1")
         If XVALO(FALTTAA) < 0.05 Then
            Label10 = ""
         Else
            Label10 = FALTTAA
         End If
      End If
   Next KKU&
   
End Sub
Sub LOTES_FACTURA(CI%, NCLI, NFACTUR$)
        'SOLUCION DADA POR MEDITEA EL CASO QUE SEA NOTA DE CREDITO Y ANULA FACTURA
        'EN ESTA RUTINA MUESTRA LOS LOTES SOLO DE LA FACTURA EN CUESTION
        NCLIE = XVALO(VALOBADA("CAJABANC", "CLI_REAL", "CODICOM_LAR='" & NFACTUR$ & "' AND NUCLIEN=" & CStr(NCLI)))
        CONDI$ = "DOSECLAR  = '" & NFACTUR$ & "'"
        CONDI$ = CONDI$ + " AND NUME_CLIE = " & CStr(NCLIE) & ""
        CONDI$ = CONDI$ + " AND COD_INTE =  " & CI%
        CONDI$ = CONDI$ + " AND CANTSALID > 0 "
        CONDI$ = CONDI$ + " AND CODIMOVI = 'RT'"
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        Set FACTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With FACTTMP
          Do While Not .EOF ' CARGA TODOS LOS ITEMS DEL MOVISTO RELACIONADO CON FACTURA
           CI% = XVALO(!cod_inte)
           LOTE$ = TRIM$(SAFETEXT$(!idenlote))
           CANTI = XVALO(!CANTSALID)
30         TTXX$ = ""
            Call REPLA(TTXX$, LOTE$, 1, 12)
            Call REPLA(TTXX$, FORMATNUM$(CANTI, "F10." & CNTDC$ & ""), 33, 10)
            If CANTI > 0 Then List2.AddItem (TTXX$)
           .MoveNext
          Loop
        End With
        FACTTMP.Close
        Set FACTTMP = Nothing

End Sub

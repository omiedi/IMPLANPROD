VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LOTESPED 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad de Componentes"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RESERVAS"
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
      Left            =   105
      TabIndex        =   11
      Top             =   600
      Width           =   8400
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6600
         TabIndex        =   12
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   14
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   1920
         TabIndex        =   13
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
         Height          =   1650
         ItemData        =   "LOTESPED.frx":0000
         Left            =   190
         List            =   "LOTESPED.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
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
         Height          =   330
         Left            =   6650
         TabIndex        =   33
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Height          =   330
         Left            =   5805
         TabIndex        =   32
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Height          =   330
         Left            =   1980
         TabIndex        =   31
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Height          =   330
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label14 
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
         Height          =   330
         Left            =   190
         TabIndex        =   29
         Top             =   280
         Width           =   8025
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ACTUALIZAR"
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
      Left            =   105
      TabIndex        =   4
      Top             =   3120
      Width           =   8400
      Begin VB.CommandButton Command5 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   17
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   7680
         Picture         =   "LOTESPED.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Ingresar Lote"
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Act.Lote Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6400
         TabIndex        =   25
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label11 
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
         Left            =   6360
         TabIndex        =   24
         Top             =   840
         Width           =   1275
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
         Height          =   300
         Left            =   1920
         TabIndex        =   19
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Left            =   4560
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
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
         Left            =   -1040
         TabIndex        =   16
         Top             =   920
         Width           =   1770
      End
      Begin VB.Label Label10 
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
         Left            =   825
         TabIndex        =   10
         Top             =   840
         Width           =   1020
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
         Left            =   -1040
         TabIndex        =   9
         Top             =   510
         Width           =   1770
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
         Height          =   300
         Left            =   825
         TabIndex        =   8
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label5 
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
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   915
         Width           =   450
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
         Left            =   2350
         TabIndex        =   5
         Top             =   920
         Width           =   690
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
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   8640
      ScaleHeight     =   4770
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   -120
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   34
         Top             =   3360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "LOTESPED.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Grabar Asignación de Lotes"
         Top             =   3720
         Width           =   660
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Auto FIFO"
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
         TabIndex        =   21
         ToolTipText     =   "Lotes Recomendados por Antiguedad"
         Top             =   1080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "LOTESPED.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   660
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "LOTESPED.frx":0762
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "LOTESPED.frx":0996
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label16 
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
      Left            =   7080
      TabIndex        =   27
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Pedido"
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
      Left            =   5450
      TabIndex        =   26
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "LOTESPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

     Command1.Enabled = False
     Screen.MousePointer = 11
     NORFA$ = TRIM$(Label7)
     '
     'DEPO% = XVALO(Control$("", "DEPOCEN"))
     '
     For i& = 1 To List2.ListCount ' VALIDACIONES
          CIXI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 16)))
          If TRIM$(Mid$(List2.List(i& - 1), 62, 16)) = "" Then
             Call COMUNI("Falta Lote Para Código: " & CODEXT$(CIXI%))
             Label7 = "FALTAN_LOTES"
'             GoTo 99 'comentado para que pueda continuar sin ingresar todos los lotes
          End If
     Next i&
     
    '
    Call LIMPIAR
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Hide
    '
99  Screen.MousePointer = 1
    Command1.Enabled = True
    '
End Sub

Private Sub Command11_Click()
    List2.Clear
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command15VIE_Click()
    Command15.Enabled = False
    Screen.MousePointer = 11
    Call LIMPIAR
    DEPO% = XVALO(Control$("", "DEPOCEN"))
    For i& = 1 To List2.ListCount
      CI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 16))) 'CODIGO INTERNO
      CANTI = XVALO(Mid$(List2.List(i& - 1), 50, 11)) ' CANTIDAD
      Call ARMALILOT(CI%, DEPO%)
      FIN& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
      '
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
        CAN = XVALO(Mid$(XX$, 35, 10)) ' LEO LA POSICION CANTIDAD
        If CANTI <= CAN Then     ' CUBRE LA NECESIDAD
          IDLOT$ = Mid$(XX$, 1, 16)
          TX$ = List2.List(i& - 1)
          Call REPLA(TX$, IDLOT$, 62, 16)
          List2.List(i& - 1) = TX$ ' LO PRESENTA EN EL LIST
          Exit For
        End If
      Next U&
      '
    Next i&
    Screen.MousePointer = 1
    Command15.Enabled = True
End Sub
Private Sub Command15_Click()
    Command15.Enabled = False
    Screen.MousePointer = 11
    Call LIMPIAR
'    DEPO% = XVALO(Control$("", "DEPOCEN"))
    Do While i& <= List2.ListCount - 1
      i& = i& + 1
      TX$ = List2.List((i& - 1))
      CI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 16))) 'CODIGO INTERNO
      A = CODEXT$(CI%)
      CANECE = XVALO(Mid$(List2.List(i& - 1), 50, 11)) ' CANTIDAD
      CANTIORIG = XVALO(Mid$(List2.List(i& - 1), 50, 11)) ' CANTIDAD
      Call ARMALISLOTES(CI%, DEPO%, , , 1)
      FIN& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
      '
      ACUMULACALOTE = 0
      '
      'RECORRPO EL LILOTSE
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
        CALOTE = XVALO(Mid$(XX$, 35, 10)) ' LEO LA POSICION CANTIDAD
        CANECE = XVALO(Mid$(List2.List(i& - 1), 50, 11)) '
        LOTSE$ = TRIM$(Mid$(XX$, 1, 16))
        RESERVADO1 = CANTRESER(CI%, LOTSE$)
        CALOTE = CALOTE - RESERVADO1
        
        If CALOTE < FRACMIN(CI%) Then GoTo 80 'SI NO TIENE CANTIDAD PASA AL SIGUIENTE
        'RECORRE EL LIST VIENDO SI USO CANTIDAD DE ESTE LOTE Y EN ESE CASO RECALCULA LA CANTIDAD DISPONIBLE DEL LOTE
        'ESTA SITUACION NO DEBERIA PASAR SALVO SI HAY ITEMS REPETIDOS EN EL LIST.
        If LOTSE$ <> "" Then
          For KII% = 1 To List2.ListCount
            If KII% - 1 <> i& - 1 Then
              If TRIM$(Mid$(List2.List(KII% - 1), 62, 16)) = LOTSE$ Then
                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 50, 11))
              End If
            End If
          Next KII%
        End If
        '
        'AQUI AGREGA EL LOTE CORRESPONDIENTE Y SI NO CUBRE LA NECESIDAD GENERA UN NUEVO REGISTRO POR EL RESTO SIN LOTE
        TY$ = ""
        If CALOTE > 0 Then
           
           If CALOTE < CANECE Then
              DIFE = CANECE - CALOTE
              CANECE = CALOTE
              TY$ = TX$
              Call REPLA(TY$, FORMATNUM(DIFE, "F10.1"), 50, 11)
           End If
           ACUMULACALOTE = ACUMULACALOTE + CANECE
           A = CODEXT$(CI%)
           Call REPLA(TX$, FORMATNUM$(CANECE, "F11.1"), 50, 11)
           Call REPLA(TX$, TRIM$(LOTSE$), 62, 16)
           List2.List(i& - 1) = TX$
           '
           'GENERA UN NUEVO REGISTRO POR EL RESTO SIN LOTE
           If TY$ <> "" Then
              List2.AddItem Left$(TY$, 61), i& ' que agregue cantidad pendiente sin lote
              i& = i& + 1
           End If
        End If

       If CANTIORIG <= ACUMULACALOTE Then Exit For
          
80    Next U&
      '
    Loop
    Screen.MousePointer = 1
    Command15.Enabled = True
End Sub


'Private Sub Command3_Click()
'    Command3.Enabled = False
'    If DERACCE("REGTRAZA", "Registro de Trazabilidad") > 1 Then
'
'        CI% = CODINT(Label8)
'        If CI% < 1 Then
'            Call MENSERR(24, "Falta seleccionar Item"): GoTo 99
'        End If
'        '
'        NUVAL = TRIM$(UCase(Label4))
'        If NUVAL = "" Then
'            Call MENSERR(24, "Falta Ingresar el Lote")
'            GoTo 99
'        End If
'        '
'        NORFA$ = Label28
'        '
'        TX$ = List2.List(List2.ListIndex)
'        canece = XVALO(Mid$(TX$, 50, 11))
'        CALOTE = XVALO(Label11)
'        LOTSE$ = TRIM$(Label4)
'        If LOTSE$ <> "" Then
'          For KII% = 1 To List2.ListCount
'            If KII% - 1 <> List2.ListIndex Then
'              If TRIM$(Mid$(List2.List(KII% - 1), 62, 16)) = LOTSE$ Then
'                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 50, 11))
'              End If
'            End If
'          Next KII%
'        End If
'        '
''        TX$ = List2.List(List2.ListIndex)
''        CANECE = XVALO(Mid$(TX$, 50, 11))
''        CALOTE = XVALO(Label5)
''        LOTSE$ = TRIM$(Label4)
''        If LOTSE$ <> "" Then
''          For KII% = 0 To List2.ListCount - 1
''            If KII% <> List2.ListIndex Then
''              If TRIM$(Mid$(List2.List(KII%), 62, 16)) = LOTSE$ Then
''                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 50, 11))
''              End If
''            End If
''          Next KII%
''        End If
'
'        TY$ = ""
'        If CALOTE > 0 Then
'           If CALOTE < canece Then
'           If CALOTE > canece Then CALOTE = canece
'              DIFE = canece - CALOTE
'              TY$ = TX$
'              Call REPLA(TY$, FORMATNUM(DIFE, "F10.1"), 50, 11)
'           End If
'           Call REPLA(TX$, FORMATNUM$(canece, "F11.1"), 50, 11)
'           Call REPLA(TX$, TRIM$(Label4), 62, 16)
'           List2.List(List2.ListIndex) = TX$
'           '
'           If TY$ <> "" Then
'              List2.AddItem TY$
'           End If
'        End If
'        '
'    End If
'    '
'99  Command3.Enabled = True
'
'End Sub
Private Sub Command3_Click()
    Command3.Enabled = False
    If DERACCE("REGTRAZA", "Registro de Trazabilidad") > 1 Then
        
        CI% = CODINT(Label8)
        If CI% < 1 Then
            Call MENSERR(24, "Falta seleccionar Item"): GoTo 99
        End If
        '
        NUVAL = TRIM$(UCase(Label4))
        If NUVAL = "" Then
            Call MENSERR(24, "Falta Ingresar el Lote")
            GoTo 99
        End If
        '
        NORFA$ = Label28
        '
        TX$ = List2.List(List2.ListIndex)
        CANECE = XVALO(Mid$(TX$, 50, 11))
        CALOTE = XVALO(Label11)
        LOTSE$ = TRIM$(Label4)
        If LOTSE$ <> "" Then
          For KII% = 1 To List2.ListCount
            If KII% - 1 <> List2.ListIndex Then
              If TRIM$(Mid$(List2.List(KII% - 1), 62, 16)) = LOTSE$ Then
                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 50, 11))
              End If
            End If
          Next KII%
        End If
        '
        TY$ = ""
        If CALOTE > 0 Then
           If CALOTE < CANECE Then
              DIFE = CANECE - CALOTE
              CANECE = CALOTE
              TY$ = TX$
              Call REPLA(TY$, FORMATNUM(DIFE, "F11.1"), 50, 11)
           End If
           '
           Call REPLA(TX$, FORMATNUM$(CANECE, "F11.1"), 50, 11)
           Call REPLA(TX$, TRIM$(Label4), 62, 16)
           List2.List(List2.ListIndex) = TX$
           '
           If TY$ <> "" Then
              List2.AddItem Left$(TY$, 61) ' que agregue cantidad pendiente sin lote
           End If
        End If
        '
    End If
    '
99  Command3.Enabled = True

End Sub


'Private Sub Command5_Click()
'    '
'    If TRIM$(Label8) <> "" Then
'        Call SELECHO("!LILOTSE")
'        If VALACT1$("!LILOTSE") <> "" Then
'           Label4 = VALACT1$("!LILOTSE")
'           CANAS = XVALO(Mid$(List2.TEXT, 50, 11))
'           SALOTT = XVALO(Mid$(VALACT2$("!LILOTSE"), 49, 10))
'           CIXI% = CODINT(Label8)
'           'IF AGREGADO POR QUE CUANDO TIENE DEPOSITO LA RUTINA ARMALISLOTES
'           'NO IMPRIME EN LA LISTA EL LA RESERVA Y EL SALDO LOGICO
'           'Y QUEDANDO ASI NO TOMABA LA CANTIDAD PARA ASIGNAR EN EL LABEL5
'           If XVALO(CONTROL$("", "DEPOCEN")) > 0 Then
'               'EN EL CASO QUE TENGA CONFIGURADO DEPOSITO DE CONSUMO CALCULO CANTIDAD
'               'RESERVADA PARA DESCONTALA DEL SALDO DEL LOTE Y ASIGNAR COMO MAXIMO LA DIFERENCIA
'               'CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & Label4 & "'"
'               'CANTRESER = XVALO(VALOBADA("RESERVA", "SUM (CANTRES)", CONDI$, "NOMESS"))
'               'CANTCONSUM = XVALO(VALOBADA("RESERVA", "SUM (CANTCONS)", CONDI$, "NOMESS"))
'               'RESERVADO1 = CANTRESER - CANTCONSUM
'               'If RESERVADO1 < 0 Then RESERVADO1 = 0
'               '
'               RESERVADO1 = CANTRESER(CIXI%, TRIM$(Label4))
'               SALOTT = XVALO(Mid$(VALACT2$("!LILOTSE"), 19, 10)) 'LEO EL STOCKS
'               SALOTT = SALOTT - RESERVADO1 ' LE DESCUENTO LA RESERVA
'               '
'           End If
'           If SALOTT < CANAS Then CANAS = SALOTT
'           Label5 = TRIM$(FORMATNUM$(CANAS, "F11.1"))
'           Label11 = TRIM$(FORMATNUM$(SALOTT, "F11.1"))
'        End If
'    End If
'    '
'End Sub
Private Sub Command5_Click()
    '
    If TRIM$(Label8) <> "" Then
        Call SELECHO("!LILOTSE")
        If VALACT1$("!LILOTSE") <> "" Then
           Label4 = VALACT1$("!LILOTSE")
           CANAS = XVALO(Mid$(List2.TEXT, 50, 11))
           SALOTT = XVALO(Mid$(VALACT2$("!LILOTSE"), 49, 10))
           CIXI% = CODINT(Label8)
           'IF AGREGADO POR QUE CUANDO TIENE DEPOSITO LA RUTINA ARMALISLOTES
           'NO IMPRIME EN LA LISTA EL LA RESERVA Y EL SALDO LOGICO
           'Y QUEDANDO ASI NO TOMABA LA CANTIDAD PARA ASIGNAR EN EL LABEL5
           If XVALO(Control$("", "DEPOCEN")) > 0 Then
               'EN EL CASO QUE TENGA CONFIGURADO DEPOSITO DE CONSUMO CALCULO CANTIDAD
               'RESERVADA PARA DESCONTALA DEL SALDO DEL LOTE Y ASIGNAR COMO MAXIMO LA DIFERENCIA
               'CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & Label4 & "'"
               'CANTRESER = XVALO(VALOBADA("RESERVA", "SUM (CANTRES)", CONDI$, "NOMESS"))
               'CANTCONSUM = XVALO(VALOBADA("RESERVA", "SUM (CANTCONS)", CONDI$, "NOMESS"))
               'RESERVADO1 = CANTRESER - CANTCONSUM
               'If RESERVADO1 < 0 Then RESERVADO1 = 0
               '
               RESERVADO1 = CANTRESER(CIXI%, TRIM$(Label4))
               RESERVADO1 = CANTRESER(CIXI%, TRIM$(Label4))
               SALOTT = XVALO(Mid$(VALACT2$("!LILOTSE"), 19, 10)) 'LEO EL STOCKS
               SALOTT = SALOTT - RESERVADO1 ' LE DESCUENTO LA RESERVA
               '
           End If
           
           If SALOTT < CANAS Then CANAS = SALOTT
           Label11 = FORMATNUM$(SALOTT, "F11.1")
           Label5 = TRIM$(FORMATNUM$(CANAS, "F11.1"))
        End If
    End If
    '
End Sub
Sub ARMALISLOTES_PED(CIXI%, Optional DEPO%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
25     If Not (.BOF And .EOF) Then
         Do While Not .EOF
           TTXX$ = REBLA$
           FEX = CVINT(!FechAlta)
           IDLO$ = SAFETEXT$(!idenlote)
           FEVE% = CVINT(!FechVenc)
           SLOTE = SaldoLote(CIXI%, IDLO$, DEPO%, FESTO%) - SALDO_LOTE_PEDIDO(CIXI%, IDLO$, XVALO(LABLE16))
           If SOLOCONSALDO% = 1 Then
              If SLOTE < 0.05 Then
                 GoTo 29
              End If
           End If
           '
           CANRES = XVALO(!CanReser)
           SALOGIC = SLOTE - CANRES
           If SALOGIC < 0 Then SALOGIC = 0
           Call REPLA(TTXX$, TRIM$(IDLO$), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F10.1"), 25, 10)
           Call REPLA(TTXX$, FORMATNUM$(SLOTE, "F10.1"), 35, 10)
           Call REPLA(TTXX$, FECHATEX$(FEVE%), 47, 8)
           If DEPO% < 1 Then
              Call REPLA(TTXX$, FORMATNUM$(CANRES, "F10.1"), 55, 10)
              Call REPLA(TTXX$, FORMATNUM$(SALOGIC, "F10.1"), 65, 10)
           End If
           '
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
29       .MoveNext
         Loop
       End If
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub


Function CANTRESER(CIXI%, LOTE$)
'
DEPO% = XVALO(Control$("", "DEPOCEN"))
SQLX$ = "SELECT IDSUBOR, CantRes FROM RESERVA WITH(NOLOCK) WHERE COD_INTE=" & CIXI% & " AND IDENLOTE='" & LOTE$ & "'"
Set rs = READSET(SQLX$)
CRTOTAL = 0
'
With rs
    Do While Not .EOF
        ORFA$ = SAFETEXT$(!idsubor)
        CR = XVALO(!CANTRES)
        CONDI$ = "COD_INTE=" & CIXI% & " AND DOSECCOR='" & ORFA$ & "' AND IDENLOTE='" & LOTE$ & "' AND CODIMOVI='TR' "
        If DEPO% > 0 Then CONDI$ = CONDI$ & " AND DEPOMOVI=" & DEPO%
        CANTR = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$))
        CRL = CR - CANTR: If CRL < 0 Then CRL = 0
        CRTOTAL = CRTOTAL + CRL
    .MoveNext
    Loop
End With
'
CANTRESER = CRTOTAL
'
End Function

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture14)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    List2.Clear
    Call CIERRARCH("*.*")

End Sub

Private Sub Label8_Change()
     CIXI% = CODINT%(Label8)
     
     If CIXI% > 0 Then
        Call CREALOTES(CIXI%)
        DEPO% = XVALO(Control$("", "DEPOCEN"))
        Call ARMALISLOTES(CIXI%, DEPO%)
     End If
End Sub


Private Sub List2_Click()
   Call LIMPIAR ' PARA QUE AL SEELECCIONAR OTRO ITEM LIMPIE LOS LABEL DE LOTES
   Call SETULTREG("LILOTSE", 0)
   Call SETULTREG("!LILOTSE", 0)
   
End Sub

Private Sub List2_DblClick()
     Call LIMPIAR
     TTXX$ = List2.List(List2.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label10 = "Reservas"
     Label1 = TRIM$(Mid$(TTXX$, 17, 32))
     Label5 = Mid$(TTXX$, 50, 11)
     Label4 = Mid$(TTXX$, 62, 16)
     
End Sub

Sub LIMPIAR()

     Label8 = ""
     Label10 = ""
     Label1 = ""
     Label5 = ""
     Label4 = ""
     Label11 = ""
End Sub


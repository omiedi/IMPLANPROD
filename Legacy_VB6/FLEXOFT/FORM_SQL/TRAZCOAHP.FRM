VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TRAZCOAHP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad de Items Facturados sin Lote-Proveedor"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   12720
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   765
      Left            =   13200
      Picture         =   "TRAZCOAHP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Grabar Asignación de Lotes"
      Top             =   3330
      Width           =   765
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Factura:                                                           Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   12840
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   10960
         TabIndex        =   31
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   9360
         TabIndex        =   30
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   7720
         TabIndex        =   14
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6525
         TabIndex        =   16
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   2330
         TabIndex        =   15
         Top             =   600
         Width           =   15
      End
      Begin VB.ListBox List2 
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
         Height          =   1680
         ItemData        =   "TRAZCOAHP.frx":030A
         Left            =   190
         List            =   "TRAZCOAHP.frx":0311
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   12465
      End
      Begin VB.Label LBLCLIENTE 
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
         Left            =   3480
         TabIndex        =   38
         Top             =   240
         Width           =   9240
      End
      Begin VB.Label lblfactura 
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
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"TRAZCOAHP.frx":03A3
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
         Left            =   195
         TabIndex        =   18
         Top             =   600
         Width           =   12465
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar"
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
      TabIndex        =   6
      Top             =   2760
      Width           =   12840
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtro Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8400
         TabIndex        =   33
         Top             =   120
         Width           =   4335
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
            Height          =   300
            Left            =   4080
            TabIndex        =   34
            Top             =   240
            Width           =   175
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   37
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label LBLRASOPRO 
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
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label LBLNUMEPROVE 
            Alignment       =   1  'Right Justify
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
            Left            =   2760
            TabIndex        =   35
            Top             =   240
            Width           =   1305
         End
      End
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
         TabIndex        =   20
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   970
         Left            =   7800
         Picture         =   "TRAZCOAHP.frx":0445
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Ingresar Lote"
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Act.Lote Lote"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6400
         TabIndex        =   29
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   6360
         TabIndex        =   28
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         TabIndex        =   22
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Left            =   4560
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -1065
         TabIndex        =   19
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label10 
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
         Left            =   825
         TabIndex        =   12
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -1065
         TabIndex        =   11
         Top             =   420
         Width           =   1770
      End
      Begin VB.Label Label8 
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
         Left            =   825
         TabIndex        =   10
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         TabIndex        =   9
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2505
         TabIndex        =   7
         Top             =   855
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Left            =   2760
         TabIndex        =   1
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00800000&
      Height          =   3360
      Left            =   13080
      ScaleHeight     =   3300
      ScaleWidth      =   1320
      TabIndex        =   2
      Top             =   -120
      Width           =   1380
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
         Left            =   130
         TabIndex        =   25
         ToolTipText     =   "Lotes Recomendados por Antiguedad"
         Top             =   1800
         Visible         =   0   'False
         Width           =   690
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
         Picture         =   "TRAZCOAHP.frx":074F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   4
         Top             =   2640
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -260
         Picture         =   "TRAZCOAHP.frx":0899
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2880
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "TRAZCOAHP.frx":27BB
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "TRAZCOAHP.frx":29EF
      TabIndex        =   27
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "TRAZCOAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

     Command1.Enabled = False
     Screen.MousePointer = 11
     '
     DEPO% = XVALO(Control$("", "DEPOCEN"))
     '
     For i& = 1 To List2.ListCount ' VALIDACIONES
          CIXI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 20)))
          If TRIM$(Mid$(List2.List(i& - 1), 74, 16)) = "" Then
             Call COMUNI("Falta Lote Para Código: " & CODEXT$(CIXI%))
             Label7 = "FALTAN_LOTES"
'             GoTo 99 'comentado para que pueda continuar sin ingresar todos los lotes
          End If
     Next i&
     
'''     'VALIDO LAS CANTIDAD DE LOS LOTES
'''     For I& = 1 To List2.ListCount ' VALIDACIONES
'''         CIXI% = CODINT%(TRIM$(Left$(List2.List(I& - 1), 16)))
'''         LOTE$ = TRIM$(Mid$(List2.List(I& - 1), 62, 16))
'''         If LOTE$ <> "" Then
'''           CANTI = XVALO(TRIM$(Mid$(List2.List(I& - 1), 50, 11)))
'''           If SaldoLote(CIXI%, LOTE$, DEPO%) + FRACMIN(CIXI%) < CANTI Then
'''              'ACTUALIZAR EL LOTE CON ACSALOTE
'''              Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
'''              Call COMUNI("Cantidad Insuficiente de Lote :" & LOTE$ & "\Para Código: " & CODEXT$(CIXI%))
'''              GoTo 99
'''              'SALE - EN CASO QUE LO QUIERE VER MAS ADELANTE LO DEBERA DEJAR SIN LOTE.
'''           End If
'''         End If
'''     Next I&
     '
     For i& = 1 To List2.ListCount ' GRABA

          CIXI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 20)))
          ORD% = XVALO(TRIM$(Mid$(List2.List(i& - 1), 120, 4)))
          CANTASI = XVALO(TRIM$(Mid$(List2.List(i& - 1), 62, 11)))
          LOTE$ = TRIM$(Mid$(List2.List(i& - 1), 74, 16))
          
          CONDI$ = "DOSECLAR = '" & TRIM$(lblfactura) & "' AND CodiMovi = 'RT' AND Cod_Inte = " & CIXI% & " AND NUORITEM = " & ORD%
          Dim MOVISE As ADODB.Recordset
          Set MOVISE = New ADODB.Recordset
          Set MOVISE = FLEXCONN.Execute(FILTSQL$("SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " + CONDI$))
          'GUARDA TODOS LOS DATOS DE LA FILA
          CACAM& = MOVISE.Fields.Count
          ReDim VARS(CACAM&)
          For U& = 1 To CACAM&
             VARS(U&) = MOVISE.Fields(U& - 1).Value
          Next U&
          '
          'PONE EN 0 CANTIDAD SALIDA Y CANTIDAD INGRESADA PARA REALIZAR EL ACSALOTE
          LOTEVIE$ = TRIM$(VALOBADA("MOVISTO", "IdenLote", CONDI$))
          Call ACTUREGISTRO("Movisto", "CantIngre", CONDI$, 0, REGAF&, "NOMESS")
          Call ACTUREGISTRO("Movisto", "CantSalid", CONDI$, 0, REGAF&, "NOMESS")
          If TRIM$(LOTEVIE$) <> "" Then
             Call ACSALOTE(CIXI%, LOTEVIE$, "NOMESS")
          End If
          'BORRA LOS REGISTROS
          Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&, "NOMESS")
          '
          Dim Movisto As ADODB.Recordset
          Set Movisto = New ADODB.Recordset
          strSQL = "SELECT * FROM MOVISTO WHERE " + CONDI$
25        On Error Resume Next
          Movisto.Open strSQL, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          '
          SUMASI = 0
          If CANTASI >= FRACMIN(CIXI%) Then
                With Movisto
                   .AddNew
                   'GENERA NUEVO REGISTRO CON LA INFORMACION GUARDADA ANTERIORMENTE
                   For U& = 1 To CACAM&
                      Movisto.Fields(U& - 1).Value = VARS(U&)
                   Next U&
                   !CANTSALID = CANTASI: SUMASI = SUMASI + CANTASI
                   !idenlote = LOTE$
                   .Update
                   Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
                End With
             End If
          '
          RESTO = XVALO(Label24) - SUMASI
          If RESTO >= FRACMIN(CIXI%) Then
            With Movisto
               .AddNew
               For U& = 1 To CACAM&
                  Movisto.Fields(U& - 1).Value = VARS(U&)
               Next U&
               !CANTSALID = RESTO
               !idenlote = ""
               .Update
            End With
          End If
          '
          Movisto.Close: Set Movisto = Nothing
          MOVISE.Close: Set MOVISE = Nothing
          'Call Me.lblfactura_Change
    '
60    Next i&

''     For I& = 1 To List2.ListCount ' GRABA
''          CIXI% = CODINT%(TRIM$(Left$(List2.LIST(I& - 1), 16)))
''          NUVALO$ = UCase(TRIM$(Mid$(List2.LIST(I& - 1), 62, 16)))
''          CONDI1$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CIXI% & ""
''
''          Call ACTUREGISTRO("Reserva", "IdenLote", CONDI1$, NUVALO$, REGAF&, "NOMESS")
''          Call ACSALOTE(CIXI%, NUVALO$, "NOMESS")
''     Next I&
      '
    Call LIMPIAR
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
99  Screen.MousePointer = 1
    Command1.Enabled = True
    Unload Me
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
    DEPO% = 1
    Do While i& <= List2.ListCount - 1
      i& = i& + 1
      TX$ = List2.List((i& - 1))
      CI% = CODINT%(TRIM$(Left$(List2.List(i& - 1), 20))) 'CODIGO INTERNO
      A = CODEXT$(CI%)
      CANECE = XVALO(Mid$(List2.List(i& - 1), 62, 11)) ' CANTIDAD
      CANTIORIG = XVALO(Mid$(List2.List(i& - 1), 62, 11)) ' CANTIDAD
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
              Call REPLA(TY$, FORMATNUM(DIFE, "F10.1"), 62, 11)
           End If
           ACUMULACALOTE = ACUMULACALOTE + CANECE
           A = CODEXT$(CI%)
           Call REPLA(TX$, FORMATNUM$(CANECE, "F11.1"), 62, 11)
           Call REPLA(TX$, TRIM$(LOTSE$), 74, 16)
           List2.List(i& - 1) = TX$
           '
           'GENERA UN NUEVO REGISTRO POR EL RESTO SIN LOTE
           If TY$ <> "" Then
              List2.AddItem Left$(TY$, 74), i& ' que agregue cantidad pendiente sin lote
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


Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("ACREDOR")
    
    LBLNUMEPROVE = XVALO(TRIM$(VALACT1$("ACREDOR")))
    
99  Command2.Enabled = True

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
        CANECE = XVALO(Mid$(TX$, 62, 11))
        CALOTE = XVALO(Label11)
        LOTSE$ = TRIM$(Label4)
        If LOTSE$ <> "" Then
          For KII% = 1 To List2.ListCount
            If KII% - 1 <> List2.ListIndex Then
              If TRIM$(Mid$(List2.List(KII% - 1), 74, 16)) = LOTSE$ Then
                CALOTE = CALOTE - XVALO(Mid$(List2.List(KII% - 1), 62, 11))
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
              Call REPLA(TY$, FORMATNUM(DIFE, "F10.1"), 62, 11)
           End If
           '
           Call REPLA(TX$, FORMATNUM$(CANECE, "F10.1"), 62, 11)
           Call REPLA(TX$, TRIM$(Label4), 74, 16)
           
           List2.List(List2.ListIndex) = TX$
           '
           If TY$ <> "" Then
              Call REPLA(TY$, Space$(16), 74, 16)
              
              List2.AddItem TY$ ' que agregue cantidad pendiente sin lote
'              List2.AddItem Mid$(TY$, 90)
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
        CIXI% = CODINT(Label8)
        Call ARMALISLOTES(CIXI%, 0, 0, 0, 1, XVALO(Me.LBLNUMEPROVE))
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
               SALOTT = XVALO(Mid$(VALACT2$("!LILOTSE"), 19, 10)) 'LEO EL STOCKS
               SALOTT = SALOTT - RESERVADO1 ' LE DESCUENTO LA RESERVA
               '
           End If
           If SALOTT < CANAS Then CANAS = SALOTT
           Label11 = FORMATNUM$(SALOTT, "F11.1")
           Label5 = TRIM$(FORMATNUM$(CANAS, "F11.1"))
        End If
        Call ARMALISLOTES(CIXI%, 0, 0, 0, 1, 0)
    End If
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
'        Call CREALOTES(CIXI%)
        DEPO% = XVALO(Control$("", "DEPOCEN"))
        Call ARMALISLOTES(CIXI%, DEPO%, , , 1)
'        Call ARMALISLOTES_SIN_SALDOLOTE(CIXI%, DEPO%, , , 1)
        
     End If
End Sub


Private Sub lblfactura_Change()
    J& = 0
    Me.List2.Clear
    FACTURAX$ = lblfactura
    SQLX$ = "SELECT CANTSALID  ,CANTINGRE, DESCRIP,DOSECLAR, NUME_CLIE,COD_INTE,COD_EXTE, VALOUNITIVA, NUORITEM,Idenlote FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT'"
    'SQLX$ = SQLX$ + " AND (IDENLOTE IS NULL OR IDENLOTE = '')"
    SQLX$ = SQLX$ + " AND DOSECLAR = '" & FACTURAX$ & "'"
    SQLX$ = SQLX$ + " ORDER BY COD_INTE ASC"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             CI1% = XVALO(!cod_inte)
             If ESTRAZABLE%(CI1%) = 1 Then
                 CANTI = 0
                 If XVALO(!CANTINGRE) > 0 Then
                   CANTI = XVALO(!CANTINGRE)
                 Else
                   CANTI = XVALO(!CANTSALID)
                 End If
                 
                 '
                 J& = J& + 1
                 
                 If CI1% > 0 And ESTRAZABLE%(CI1%) = 1 Then
                    TXTX$ = ""
                    Call REPLA(TXTX$, SAFETEXT$(!Cod_Exte), 1, 20)
                    Call REPLA(TXTX$, TRIM$(SAFETEXT$(!Descrip)), 21, 40)
                    Call REPLA(TXTX$, FORMATNUM$(CANTI, "F10.1"), 62, 11)
                    Call REPLA(TXTX$, SAFETEXT$(!idenlote), 74, 16)
                    Call REPLA(TXTX$, FORMATNUM$(XVALO(!VALOUNITIVA), "F12.2"), 91, 12)
                    Call REPLA(TXTX$, FORMATNUM$(XVALO(!NuOrItem), "F4"), 120, 4)
                    TRAZCOAHP.List2.AddItem (TXTX$)
                 End If

            End If
            .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing

End Sub

Private Sub LBLNUMEPROVE_Change()
    NPROV = XVALO(LBLNUMEPROVE)
    If NPROV > 0 Then
      LBLRASOPRO = ECOARCH$("PROVED1", MKI$(NPROV))
    Else
      LBLRASOPRO = ""
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
     Label8 = Mid$(TTXX$, 1, 20)
     Label10 = "STOCK"
     Label1 = TRIM$(Mid$(TTXX$, 21, 40))
     Label5 = Mid$(TTXX$, 62, 11)
     Label4 = Mid$(TTXX$, 74, 16)
     
End Sub

Sub LIMPIAR()

     Label8 = ""
     Label10 = ""
     Label1 = ""
     Label5 = ""
     Label6 = ""
     Label4 = ""
     Label11 = ""
     LBLNUMEPROVE.Caption = ""
     LBLRASOPRO.Caption = ""
End Sub


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DESCOL09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Materiales ingresados por Colectora"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12060
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
   ScaleHeight     =   6585
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9330
      TabIndex        =   5
      ToolTipText     =   "Control de Validez de Código,lote,Cantidad"
      Top             =   5520
      Width           =   1740
   End
   Begin VB.TextBox Text11 
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
      Left            =   7905
      TabIndex        =   2
      Text            =   " "
      ToolTipText     =   "Fecha del Movimiento"
      Top             =   5520
      Width           =   1065
   End
   Begin VB.CommandButton COMMAND12 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8955
      Picture         =   "DESCOL09.frx":0000
      TabIndex        =   3
      Top             =   5520
      Width           =   175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PROCESAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9330
      TabIndex        =   6
      ToolTipText     =   "Procesa La Salida de los Items Marcados como OK"
      Top             =   6000
      Width           =   1740
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
      Height          =   300
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   0
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   5520
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Referencia Mandataria"
      Top             =   6000
      Width           =   7815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   10965
      Begin VB.Frame Frame3 
         Caption         =   "                "
         Height          =   4500
         Left            =   7320
         TabIndex        =   22
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "                "
         Height          =   4500
         Left            =   6360
         TabIndex        =   21
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "                "
         Height          =   4500
         Left            =   5000
         TabIndex        =   20
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   4500
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         ItemData        =   "DESCOL09.frx":030A
         Left            =   120
         List            =   "DESCOL09.frx":030C
         TabIndex        =   14
         Top             =   480
         Width           =   10725
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                     Descripción                                Lote                    Cantidad  Control"
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   10725
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   6690
      Left            =   11180
      ScaleHeight     =   6630
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command21 
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
         Picture         =   "DESCOL09.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Seleccionar Archivo de Datos Colectados"
         Top             =   1800
         Width           =   650
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "DESCOL09.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Informe de Items que no superaron el Control "
         Top             =   960
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   5760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "DESCOL09.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DESCOL09.frx":08AC
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6120
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "DESCOL09.frx":27CE
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DESCOL09.frx":2A02
      TabIndex        =   25
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   1680
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Height          =   225
      Left            =   7080
      TabIndex        =   23
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -2520
      TabIndex        =   17
      ToolTipText     =   "Depósito de Salida"
      Top             =   5580
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2265
      TabIndex        =   16
      Top             =   5520
      Width           =   4740
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -2520
      TabIndex        =   15
      Top             =   6060
      Width           =   3765
   End
End
Attribute VB_Name = "DESCOL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text1.TEXT = VDEV$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command12_Click()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
       Text11.TEXT = VALACT1$("SELFECHA")
    End If

End Sub

Private Sub Command2_Click()
    
    Command2.Enabled = False
    If List2.ListCount < 1 Then
       Call COMUNI("Debe Leer el Arhcivo Generado por Colectora")
       GoTo 99
    End If
    '
    If TRIM$(Label1) = "" Then
     Call COMUNI("Falta Ingresar Depósito")
     GoTo 99
    End If
    '
    DEPO% = XVALO(Text1)
    For i& = 1 To List2.ListCount
        CONTR$ = ""
        TX$ = List2.List(i& - 1)
        CI% = CODINT%(Mid$(TX$, 1, 16))
        If CI% < 1 Then
          CONTR$ = "Codigo: " & TRIM$(Mid$(TX$, 1, 16)) & " Inexistente"
          Call REPLA(TX$, CONTR$, 70, 40)
          List2.List(i& - 1) = TX$
          GoTo 30
        End If
        '
        LOTEX$ = Mid$(TX$, 47, 12)
        CONDI$ = "Cod_Inte = " & CI% & "  AND IdenLote = '" & LOTEX$ & "'"
        CANRE& = CantRegistros&("LOTINGRE", CONDI$)
        If CANRE& < 1 Then
            CONTR$ = "Lote: " & LOTEX$ & " No Válido"
            Call REPLA(TX$, CONTR$, 70, 40)
            List2.List(i& - 1) = TX$
            GoTo 30
        End If
        '
        CANTI = XVALO(Mid$(TX$, 61, 8))
        If CANTI < FRACMIN(CI%) Then
           CONTR$ = "Cantidad No Válida"
           Call REPLA(TX$, CONTR$, 70, 40)
           List2.List(i& - 1) = TX$
           GoTo 30
        End If
        '
        If CANTI > SaldoLote(CI%, LOTEX$, DEPO%) + FRACMIN(CI%) Then
            CONTR$ = "Cantidad Excede Lote"
            Call REPLA(TX$, CONTR$, 70, 40)
            List2.List(i& - 1) = TX$
            GoTo 30
         End If
30      If TRIM$(CONTR$) <> "" Then
            U& = U& + 1
            Call GRAREG("INFODESP", TX$, U&)
        Else
            CONTR$ = "OK"
            Call REPLA(TX$, CONTR$, 70, 40)
            List2.List(i& - 1) = TX$
        End If
         
   Next i&
   Call SETULTREG("INFODESP", U&)
   Call CIERRARCH("*.*")
99 Command2.Enabled = True
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
  
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "CAPTURA DE INVENTARIOS"
    CommonDialog1.Filter = "Archivo de Texto(*.TXT)|*.TXT|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    ARCH$ = UCase$(CommonDialog1.FileName)
    Label2 = "INGRESO" 'bandera para saber si abrio el CommonDialog1 para que no genere error al solicitar listado de informes
    NARCHIVO% = FILEOPEN%(ARCH$, 1, 128) 'ABRE EL ARCHIVO DE LECTURA
        '
    List2.Clear
    U& = 0
    Do While Not EOF(NARCHIVO%) ' RECORRE
        Line Input #NARCHIVO%, TX1$ ' LEE CADA REGISTRO
        If TRIM$(TX1$) = "" Then GoTo 50 ' VALIDO QUE EL REGISTRO NO ESTE VACIO
        For i& = 1 To Len(TX1$) Step 35
          TXTO$ = Mid$(TX1$, i&, 35)
          '
          CODIEXT$ = TRIM$(Left$(TXTO$, 16))  '        CODIGO EXTERNO
          LOTEX$ = Mid$(TXTO$, 17, 12) '        LOTE
          CI% = CODINT(CODIEXT$) '              CODIGO INTERNO
          Descr$ = DESCRIT$(CI%)
          CANTI = XVALO(Mid$(TXTO$, 29, 8)) '
          '
          Call REPLA(TX$, CODIEXT$, 1, 16)
          Call REPLA(TX$, Descr$, 19, 27)
          Call REPLA(TX$, LOTEX$, 47, 12)
          Call REPLA(TX$, FORMATNUM$(CANTI, "F8.0"), 61, 8)
          '
          List2.AddItem TX$
          '
        Next i&
50  Loop
    Close NARCHIVO%
    
    Screen.MousePointer = 1
    '
    '
    Call COMUNI("Lectura Realizada ")
99  Command21.Enabled = True
  
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   
   Screen.MousePointer = 11
    '
   'bandera para saber si abrio el CommonDialog1 para que no genere error al solicitar listado de informes
   If TRIM$(Label2) <> "" Then
      Call COMUNI("Debe Primero Reiniciar El Modulo \y Luego Solicitar Nuevamente el Listado ")
      Call FINAL("")
   End If
   '
   If ULTREG&("INFODESP") > 0 Then ' LISTADO DE ERRORES
        Call FINAL("*INFODESP")
   Else
        Call COMUNI("No se han Encontrado Errores en la Lectura del Archivo")
   End If
    '
   Screen.MousePointer = 1

   Call CIERRARCH("*.*")
   Command4.Enabled = True
   
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    '
    HOII% = HOY(HOS$)
    If List2.ListCount < 1 Then
       Call COMUNI("Debe Leer el Arhcivo Generado por Colectora")
       GoTo 99
    End If
    '
    TX$ = List2.List(0)
    If Len(TRIM$(TX$)) < 70 Then
       Call COMUNI("Debe Realizar Primero el Control ")
       GoTo 99
    End If
    '
    If TRIM$(Label1) = "" Then
     Call COMUNI("Falta Ingresar Depósito")
     GoTo 99
    End If
    '
    If TRIM$(Text3) = "" Then
     Call COMUNI("Falta Ingresar Referencia")
     GoTo 99
    End If
    '
    If FECHANUM(Text11) < 1 Then
     Call COMUNI("Fecha No Válida")
     GoTo 99
    End If

    DEPO% = XVALO(Text1)
    REFE$ = Text3
    FE% = FECHANUM(Text11)
    U& = 0
    Call ABRESTOCKS

    For i& = 1 To List2.ListCount
        CONTR$ = ""
        TX$ = List2.List(i& - 1)
        CI% = CODINT%(Mid$(TX$, 1, 16))
        '
        LOTEX$ = Mid$(TX$, 47, 12)
        '
        CANTI = XVALO(Mid$(TX$, 61, 8))
        '
30      If Mid$(TX$, 70, 2) = "OK" Then
            GoSub GENEROMOVIMIENTO
        End If
   Next i&
   Call COMUNI("Se han Realizado los Movimientos de Stocks\Para Los Items Correctamente Leidos\Este Modulo se Reiniciará\Vuelva a Entrar Para Ver Listado de Informes\Con Items No Considerados en Este Proceso ")
   Call FINAL("") 'Nota: Se  realiza un call final del formulario luego de realizar
    'la lectura debido a que cuando se utiliza el CommonDialog cambia el path
    ' y al intentar ejecutar el listado de errores (Dat)devolvia un mensaje
    ' que no encontraba el archivo MSTDRV04.EXE (FALTA UTILITARIO MSTDRV04.EXE)
99 Command7.Enabled = True
'
Exit Sub
'
GENEROMOVIMIENTO:
    '
    CMOV$ = "RD"
    NVM$ = TRIM$(ProNroComprobante("RD"))
    While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    NRO& = XVALO(NVM$)
    '
    DOSECO$ = "RD." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
    DOPRICO$ = CMOV$ + "." + NVM$
    NUORIT% = 0
    VUNI# = 0
    MONEII% = 0
    NC% = 0
    REFE$ = TRIM$(Text3)
    '
    Call MOVISTOK(FE%, CI%, LOTEX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, -1 * CANTI)
    '
Return

End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)

    Label2 = "" 'bandera para saber si abrio el CommonDialog1 para que no genere error al solicitar listado de informes
    HOII% = HOY(HOS$)
    Text11 = FECHATEX$(HOII%)
    '
End Sub

Private Sub Text1_Change()
  '
    DEPOSAL% = XVALO(Text1)
    On Error Resume Next
    If DEPOSAL% > 0 And ECOARCH("TADEPOSI", Chr$(DEPOSAL%)) Then
       Label1 = ECOARCH("TADEPOSI", Chr$(DEPOSAL%))
    Else
      Label1 = ""
    End If
  '
End Sub





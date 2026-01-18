VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form BCODECTP 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   525
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
      Height          =   690
      Left            =   9360
      Picture         =   "BCODECTP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Continúa Carga de Remito."
      Top             =   4440
      Width           =   650
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   185
      Width           =   9150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   9150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4545
      Left            =   9240
      ScaleHeight     =   4485
      ScaleWidth      =   1005
      TabIndex        =   2
      Top             =   -210
      Width           =   1065
      Begin VB.Frame Frame17 
         BackColor       =   &H00004080&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   630
         Left            =   90
         TabIndex        =   6
         Top             =   3360
         Width           =   630
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   105
            TabIndex        =   7
            Top             =   270
            Width           =   435
         End
      End
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
         Picture         =   "BCODECTP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminar Producto de la Lista"
         Top             =   1140
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
         Picture         =   "BCODECTP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "BCODECTP.frx":075E
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   9150
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "BCODECTP.frx":2680
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "BCODECTP.frx":28B4
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "QUILMES ---SI SON BOBINAS CODIGO  CANTIDAD LOTE (CA TT 47 60 CR|1|BR-43-3)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Depòsito Salida"
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
      Left            =   120
      TabIndex        =   15
      Top             =   4600
      Width           =   1275
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
      Height          =   300
      Left            =   1275
      TabIndex        =   14
      Top             =   4560
      Width           =   420
   End
   Begin VB.Label Label5 
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
      Left            =   1755
      TabIndex        =   13
      Top             =   4560
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "OJO SE USA EN REMITOS PARA ENVAPLAST Y CAMBACRES"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label VALI_BCODE09 
      Caption         =   "VALI_BCODE09"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "BCODECTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTITFO&

Function CARGARASIGLOT07_CAMBACERES%()


'************** FUERA DE USO  ******************************

        CARGARASIGLOT07_CAMBACERES% = 1
        EPAC$ = TRIM$(UCase$(TRIM$(EMPREAC$)))

        If InStr(EPAC$, "CAMBACERES") < 1 And InStr(EPAC$, "CONTEPLAST") < 1 Then
            Exit Function
        End If

        
        CARGARASIGLOT07_CAMBACERES% = 1
        Dim List As ListBox
        Set List = BCODECTP.List1
        Dim VECTORCANTI(32766)
        Dim VECTORLOTE$(32766)
        For U& = 1 To List.ListCount ' VALIDACIONES
           Z$ = List.List(U& - 1)
           Codigo$ = TRIM$(Left$(Z$, 16))
           CI1% = CODINT%(Codigo$)
           DESCRIX$ = TRIM$(Mid$(Z$, 18, 30))
           CANTI = XVALO(Mid$(Z$, 49, 8))
           LOTE$ = TRIM$(UCase$(Mid$(Z$, 61, 12)))


          VECTORCANTI(CI1%) = VECTORCANTI(CI1%) + CANTI 'ACUMLA LA CANTIDAD DE CADA CODIGO
          'AGREGA UN STRING CON LOS LOTES EN EL VECTOR
          If InStr(VECTORLOTE$(CI1%), LOTE$) < 1 Then 'SI SE REPITE EL LOTE  NO LO AGREGA EN EL STRING DE LOTES
            If TRIM$(VECTORLOTE$(CI1%)) = "" Then
               VECTORLOTE$(CI1%) = LOTE$ & ";"
            Else
               VECTORLOTE$(CI1%) = VECTORLOTE$(CI1%) & LOTE$ & ";"
               strArray = Split(TEXTOTEXT1, ";")
               If UBound(strArray) > 26 Then 'si para un mismo codigo tiene mas de 26 lotes presenta mensaje
                  Call COMUNI("Atenciòn !!! Demasiados Lotes Para Còdigo: " & Codigo$ & "\Considere Cargar en Màs de Un Remito")
                  CARGARASIGLOT07_CAMBACERES% = -1
                  Exit Function
               End If
            End If
          End If
        Next U&
        '
        'ARMA EL CALL REPLA PARA LLEVAR AL LIST1 DE ASGLOT07
        ASIGLOT07.List1.Clear
        REBLA$ = REGBLAN$("BOLREDET"): JJ& = 0
        For CI1% = 1 To 32766
           If VECTORCANTI(CI1%) > 0 Then
                JJ& = JJ& + 1
                Call REPLA(XX$, CODEXT$(CI1%), 1, 16)
                Call REPLA(XX$, DESCRIT0$(CI1%), 18, 33)
                Call REPLA(XX$, FORMATNUM$(VECTORCANTI(CI1%), "F10.1"), 54, 12)
                Call REPLA(XX$, FORMATNUM$(VECTORCANTI(CI1%), "F10.1"), 64, 10)
                Call REPLA(XX$, TRIM$(FORMATNUM$(JJ&, "I5")), 80, 5)
                Call REPLA(XX$, VECTORLOTE$(CI1%), 85, 390)
                '
                'CARGO EL BOLREDET
                UM$ = Unimed$(CI1%)
                CANTI = VECTORCANTI(CI1%)
                
                Call CARGARBOLREDETDEASIGLOT07(CI1%, CANTI, JJ&)
                ASIGLOT07.List1.AddItem XX$
                Call LLENARLISTDEBOLREDET
                
           End If
        Next CI1%
        Call SETULTREG("!BOLREDET", JJ&)
        ASIGLOT07.VALIDA = "APRO"


End Function

Sub CARGARBOLREDETDEASIGLOT07(CI1%, CANTI, JJ&)
      REBLA$ = REGBLAN$("!BOLREDET")
    

      Q1# = XVALO(CANTI)
      y$ = REBLA$
      Call REPLA(y$, MKI$(CI1%), 83, 2)
      Call REPLA(y$, Left$(Unimed$(CI1%), 4), 85, 4)
      Call REPLA(y$, MKD$(Q1#), 89, 8)
      NPED& = XVALO(FOREMIT07.lblNroPed.Caption)
      Call REPLA(y$, MKS$(NPED&), 151, 4)
      Call GRAREG("!BOLREDET", y$, JJ&)
      
End Sub


Sub LLENARLISTDEBOLREDET()
   FOREMIT07.List1.Clear
   v& = 0
   For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            v& = v& + 1
            Call GRAREG("!BOLREDET", X$, v&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, Unimed$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            FOREMIT07.List1.AddItem RELI$
        End If
    Next U&
End Sub

Private Sub BOTNEXT_Click()
        '
        EPAC$ = UCase$(TRIM$(EMPREAC$))
        BOTNEXT.Enabled = False
        For U& = 1 To List1.ListCount
            CI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            CANTSA = XVALO(Mid$(List1.List(U& - 1), 49, 8))
            LOTE1$ = TRIM$(Mid$(List1.List(U& - 1), 61, 12))
            DEPX% = XVALO(Label4)
            If Left(LOTE1$, 3) <> "OF-" And InStr(EPAC$, "ENVAPLAST") < 1 Then   'SI EL LOTE ES DE UNA OF NO PRESENTA MENSAJE.
                For UU& = 1 To List1.ListCount                                 'SOLO SI NO ES ENVAPLAST
                   If U& <> UU& Then
                      LOTE2$ = TRIM$(Mid$(List1.List(UU& - 1), 61, 12))
                      If LOTE1$ = LOTE2$ Then
                        
                          OPX% = COMALTER%("Se Detectaron Lotes Repetidos\\Cancelar y Controlar\Continuar")
                          
                          If OPX% < 2 Then
                            GoTo 99
                          Else
                            GoTo 60
                          End If
                          
                      End If
                   End If
                Next UU&
            End If
            '
        Next U&
        
60      For U& = 1 To List1.ListCount
            CI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            CANTSA = XVALO(Mid$(List1.List(U& - 1), 49, 8))
            LOTE$ = TRIM$(Mid$(List1.List(U& - 1), 61, 12))
            DEPX% = XVALO(Label4)
            '
            CANTILOTE = SaldoLote(CI%, LOTE$, DEPX%)
            If CANTILOTE < CANTSA Then
               DEPOS$ = TRIM$(ECOARCH$("ECODEP", MKI$(DEPX%)))
               Call COMUNI("Lote Insuficiente Para Código: " & CODEXT$(CI%) & " - Lote: " & LOTE$ & "\En Depósito: " & DEPOS$ & " - Saldo: " & CANTILOTE)
               GoTo 99
            End If
            '
        Next U&
        '
        'VALIDAR SI EL CODIGO ES TRAZABLE QUE TENGA LOTE
        For U& = 1 To List1.ListCount
            CI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            If ESTRAZABLE%(CI%) = 1 Then
                CANTSA = XVALO(Mid$(List1.List(U& - 1), 49, 8))
                LOTE$ = TRIM$(Mid$(List1.List(U& - 1), 61, 12))
                DEPX% = XVALO(Label4)
                '
                If LOTE$ = "" Then
                   Call COMUNI("Falta Lote Para Código: " & CODEXT$(CI%))
                   GoTo 99
                End If
            End If
            '
        Next U&

     
        ASIGLOT07.List1.Clear 'LIMPIO EL LISTE DEL FORMULARIO ASIGLOT07 POR SI QUEDA BASURA.
        Label18 = ""
        
        VALI_BCODE09 = "APROB"
'        If CARGARASIGLOT07_CAMBACERES% < 0 Then VALI_BCODE09 = "" 'POR DEFECTO DEVUELVE 1 Y NO CAMBIA VALI_BCODE09
        
        
        Hide
99      BOTNEXT.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    If DERACCE%("STPCOBAR", "Configuración de Funcionamiento de Codigos de Barra") = 2 Then
        STPKBN04.Show 1
    End If
    Command13.Enabled = True
    '
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  List1.Clear
  VALI_BCODE09 = ""
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  Command2.Enabled = True
  Hide 'Unload Me
  '
End Sub

Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Text1.SetFocus
    End If
    '
    Label18 = List1.ListCount
End Sub

Private Sub Form_Load()
    '
    XXTITU$ = Space$(0)
    '
    VALI_BCODE09 = ""
    XXCA$ = "Codigo"
    XXTC$ = "   Cant."
    XXTD$ = "Lote"
    Descr$ = "Descripción"
    Call REPLA(XXTITU$, XXCA$, 1, 17)
    Call REPLA(XXTITU$, Descr$, 18, 30)
    Call REPLA(XXTITU$, XXTC$, 49, 8)
    Call REPLA(XXTITU$, XXTD$, 61, 4)
    '
    List2.AddItem XXTITU$
    '
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIT07.Show 1
        Call FINAL("")
    End If
    TOTITFO& = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub Label4_Change()
   Label5.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Label4.Caption))))
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'EJEMPLOS DE QR
'BOBINAS DE RECEPCION: CA TP 120 190 BL|BR-745-5
'BIG BAG: CODIGO,LOTE,CANTIDAD,CLIENTE,CI%,PEDIDO,MOVISTO_ID
'PROBAR CON CONDIGO CA TT 180 120 BL|BR-237-1 EN DEPO 1
'CA TT 180 120 BL|BR-237-1
'CA TT 180 120 BL|BR-238-1
'CA TT 180 120 BL|BR-237-2

'BOLSAS Y CORTES:  CODIGO,LOTE,CANTIDAD,CLIENTE,CI%,PEDIDO,MOVISTO_ID
'CORTES: COR TP 101 2 097|OF-000439-1|200|3584|1224|402|100884
'BOLSAS: bol 8176-1|OF-001062-1|3045|0|2155|0|184503

'________________________________________________
'ETIQUETA GENERADA EN LA OF DE CAMBACERES
'TEXTOQR$ = Codigo$ & "|" & NORFA$ & "|" & LoteNorfa$ & "|" & PESOENKG$ & "|" & METROSREALES$
'CA TP 120 190 BL|OF-000100-1|BR-745-5|125|400
'________________________________________________
'CA E 6.5 1250 BL|OF-000322-1|12000|0|173|0|70053
'______________________________________________________
'ESLINGAS QUILMES
'CA E 6.5 1250 AZ|OF-000405-1|OF-000405001|486.2|12000.0
'CA C 1.5|OF-000025-1|OF-000025008|222.0|51200.0
      
      If KeyAscii <> 13 Then Exit Sub
      EPAC$ = UCase$(TRIM$(EMPREAC$))
      If Text1 <> "" Then
       LecturaEtiqueta$ = TRIM$(Text1)
       LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
       LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
       
       LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "ç", "|")
       LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "Ç", "|")

       TEXTOTEXT1 = TRIM$(LecturaEtiqueta$)
       
       strArray = Split(TEXTOTEXT1, "|")
       On Error Resume Next
       XXARTIDES$ = Space$(128)

       Codigo$ = strArray(0)
       NORFA$ = strArray(1)
       CI1% = CODINT%(Codigo$)
       DESCRIX$ = DESCRIT$(CI1%)
       '
       ESSCRAP% = 0
       CF$ = TRIM$(CODFAMIL$(CI1%))
       If CF$ = "SC" Then
          ESSCRAP% = 1
       End If

       If CODINT%(Codigo$) <= 0 Then
          Call MENSERR(24, "Código no Válido")
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          GoTo 99
        End If

       If InStr(EPAC$, "CAMBACERES") > 0 Then
          NC = 0
          LOTEX$ = strArray(2)
          'SE AGREGO ESTE CAMBIO PARA QUILMES
          If InStr(EPAC$, "QUILMES") > 0 Then 'POR QUE SI TIENE ETIQ DE BR SOLO TIENE CODIGO Y LOTE
            If TRIM$(CF$) <> "BO" Then
              If CODINT%(Codigo$) > 0 And UCase$(Left$(NORFA$, 3)) = "BR-" Then LOTEX$ = NORFA$ '01/02/2024
            Else
              LOTEX$ = NORFA$ '01/02/2024
              Cantidad = XVALO(strArray(2))
            End If
          End If
          If InStr(EPAC$, "QUILMES") < 1 And InStr(EPAC$, "CAMBACERES") > 0 Then 'SI NO SON BOLSAS Y ES CAMBACERES
              If CODINT%(Codigo$) > 0 And UCase$(Left$(NORFA$, 3)) = "BR-" Then LOTEX$ = NORFA$ '09/02/2024
              If TRIM$(CF$) <> "BO" Then Cantidad = 1
          End If
          NROPED& = XVALO(FOREMIT07.lblNroPed) 'POR DEFECTO POR QUE NO VIENE DE PEDIDO
          If NROPED& > 0 Then NC = XVALO(FOREMIT07.Text1)
            
          Call REPLA(XXARTIDES$, SAFETEXT$(NROPED&), 90, 9) '
          If TRIM$(CF$) <> "BO" And TRIM$(CF$) <> "EE" And InStr(EPAC$, "QUILMES") > 0 Then
            Cantidad = 1
          End If
          If ESSCRAP% > 0 Then
             Cantidad = XVALO(strArray(3))
          End If
          
          If InStr(EPAC$, "QUILMES") > 0 Then 'AGREGADO PARA QUILMES 13-04-23
             If InStr(CF$, "R1") > 0 Or InStr(CF$, "R2") > 0 Or InStr(CF$, "TP") > 0 Or InStr(CF$, "TT") > 0 Or InStr(CF$, "PL") > 0 Or InStr(CF$, "TL") > 0 Or CF$ = "CT" Or CF$ = "TE" Or CF$ = "EE" Or CF$ = "CT" Then ' QUITE ESTO Or CF$ = "BO"  POR  QUE TRAI MAL EL LOTE CUANDO ERAN BOLSAS
            'If CF$ = "BO" Or CF$ = "CT" Or CF$ = "TE" Then 'SI ES QUILMES Y ES BOLSAS
              NORFA$ = strArray(1)
              'Cantidad = XVALO(strArray(2)) LO COMENTE PARA QUE QUEDE CANTIDAD 1 ACA EN QUILMES Y ESTAS FAMLIAS
              LOTEX$ = strArray(2)
              IDMOVISTO& = XVALO(strArray(6))
            End If
            If CF$ = "EE" Or CF$ = "BH" Or CF$ = "CT" Then 'AGREGADO QUE SI SON HILOS TAMBIEN LEA LA CANTIDAD 22/07/24 PROBADO EN REMITO OK
                'LECTURA DE  HILO ZA HILO CON|BR-133-1|ENVAPLAST SRL (QUILMES)|200.0|229.0|INT.hiloconducti|22/07/24
                'CA E 6.5 1250 AZ|OF-000405-1|OF-000405001|486.2|12000.0 'ESTO ES AHORA QUE ES TRAZABLE 17/04/24
                 Cantidad = XVALO(strArray(4))
                 'ESTO LO COMENTE POR QUE ESTABA VALIDANDO PARA ESTA CADENA 'CA E 6.5 1250 BL|OF-000322-1|12000|0|173|0|70053
'                LOTEX$ = strArray(1) 'ESLINGAS
'                Cantidad = XVALO(strArray(2))
'                IDMOVISTO& = XVALO(strArray(6))
'
            End If
            If CF$ = "M1" Or CF$ = "M2" Then
               'CODIGO,LOTE,PROVEEDOR,CANTIDAD DE PAQUETES,CANTIDAD TOTAL DE KG, NRO DE REMITO PROVEEDOR,FECHA RECEPCION
               'MP 11|BR-372-1|PETROQUIMICA CUYO|60.0|1500.0|1319|14/03/24
               LOTEX$ = strArray(1) 'MATERIAL PARA EXTRUSION
               Cantidad = XVALO(strArray(4))
               
            End If
          End If
          
          '
          
          If ESSCRAP% > 0 Then
           Cantidad = XVALO(strArray(3))
          End If


          For U& = 1 To List1.ListCount
            If Mid(List1.List(U& - 1), 61, 12) = LOTEX$ Then  ' COINCIDE EL ID
              Call COMUNI("Atención ! Lote ya Informado en esta lista")
              GoTo 90
            End If
          Next U&
          'ACA VALIDAR SI NO SE PASA DEL PEDIDO
          '**VER POR QUE ESTO LO TIENE QUE H ACER AL GRABAR TIENE UNA VALIDACION
          
          If NROPED& > 0 Then
            CONDI$ = "NROPED = " & NROPED& & "  AND CODIINT = " & CI1% & " And Status < 8 And NROCLIE = " & NC
            CANTRE& = CantRegistros("PEDDETA", CONDI$)
            If CANTRE& < 1 Then
              Call MENSERR(24, "El Artículo " & CODEXT$(CI1%) & " No Corresponde al Pedido: " & NUPED&)
              GoTo 90
            End If
          End If
          Call REPLA(XXARTIDES$, Codigo$, 1, 16)
          Call REPLA(XXARTIDES$, DESCRIX$, 18, 30)
          Call REPLA(XXARTIDES$, FORMATNUM$(XVALO(Cantidad), "F8.0"), 49, 8)
          Call REPLA(XXARTIDES$, LOTEX$, 61, 12) ' LOTE

       Else
            'VALIDO PARA BIG BAG, BOLSAS,TELAS
            NORFA$ = strArray(1)
            Cantidad = XVALO(strArray(2))
            If ESSCRAP% > 0 Then
               Cantidad = XVALO(strArray(3))
            End If
            
            
            If CF$ = "TE" Then 'si son TELAS tiene solo codigo y lotebr
               If Cantidad < 1 Then Cantidad = SaldoLote(CI1%, NORFA$, XVALO(Label4))
               If Cantidad <= 0.05 Then
                  Call COMUNI("El lote: " & NORFA$ & " No posee Stock")
                  GoTo 90
               End If
            End If
            
            NC = XVALO(strArray(3))
            CI1% = XVALO(strArray(4))
            NROPED& = XVALO(strArray(5))
            If (CF$ <> "BO" And CF$ <> "TE" And CF$ <> "SC") Then 'SE AGREGO ESTO PARA CONTEPLAST 21/03/22  (And CF$ <> "SC")
                IDMOVISTO& = XVALO(strArray(6))
                On Error GoTo 0
                If IDMOVISTO& < 1 Then
                   OPX% = COMALTER%("Atención La Lectura de La Etiqueta No Contiene Identificador de Ubicación\Puede Continuar y Liberar Ubicación Normalmente\O Re-Imprimir Etiqueta\\Cancelar\Cargar de Todos Modos")
                   If OPX% < 2 Then GoTo 90
                End If
                
                If NC <> XVALO(FOREMIT07.Text1) Then 'SI SON BOLSAS O CORTES O TELAS NO VALIDA
                      Call COMUNI("Atención No Coincide El Nro. de Cliente Leido en la Etiqueta Con el Cliente Activo")
                      GoTo 90
                End If
            Else
               'si son bolsas
               If InStr(EPAC$, "CONTEPLAST") > 0 Then
                If CF$ = "BO" Or CF$ = "CT" Or CF$ = "TE" Then
                   IDMOVISTO& = XVALO(strArray(6))
                End If
               End If
            End If
       
            'VALIDO QUE CORRESPONDA AL MISMO PEDIDO
            If (CF$ <> "BO" And CF$ <> "TE") Then
              NPED_ENPANTALLA& = XVALO(FOREMIT07.lblNroPed)
              If NPED_ENPANTALLA& > 0 Then ''SI SON BOLSAS O CORTES NO VALIDA
               If NROPED& <> NPED_ENPANTALLA& Then
                  Call COMUNI("Atención El Item Seleccionado No Corresponde al Pedido a Despachar")
                  Exit Sub
               End If
              End If
            End If
            '
            If XVALO(Cantidad) < FRACMIN(CI1%) Then
             Call MENSERR(24, "Cantidad no Válida")
             Text1.SelStart = 0
             Text1.SelLength = Len(Text1)
             GoTo 99
            End If
           
            For U& = 1 To List1.ListCount
                    VALORMOVISTO_ID = XVALO(Mid(List1.List(U& - 1), 100, 9))
                    If VALORMOVISTO_ID > 0 Then
                        If VALORMOVISTO_ID = IDMOVISTO& Then  ' COINCIDE EL ID
                          Call COMUNI("Atención Pallet Yá Informado Para Este Remito")
                          GoTo 90
                        End If
                    End If
            Next U&
            
            Call REPLA(XXARTIDES$, Codigo$, 1, 16)
            Call REPLA(XXARTIDES$, DESCRIX$, 18, 30)
            Call REPLA(XXARTIDES$, FORMATNUM$(XVALO(Cantidad), "F8.0"), 49, 8)
            Call REPLA(XXARTIDES$, NORFA$, 61, 12) ' LOTE
            Call REPLA(XXARTIDES$, SAFETEXT$(NC), 80, 9)   ' CLIENTE
            Call REPLA(XXARTIDES$, SAFETEXT$(NROPED&), 90, 9) ' NRO PEDIDO
            Call REPLA(XXARTIDES$, SAFETEXT$(IDMOVISTO&), 100, 9) ' IDMOVISTO DEL CIERRE DE OF
       End If
       List1.AddItem XXARTIDES$
       Label18 = List1.ListCount
       '
90     Text1 = ""
       Text1.SetFocus

       
'        XXARTIDES$ = Space$(128)
'        '
'        XXCODI$ = Codigo
'        XXCANT1$ = Cantidad
'        XXCANT$ = FORMATNUM$(XVALO(XXCANT1$), "F8.0")
'        XXTARD1$ = TRIM$(NORFA$)
        'VALIDACIONES
'        CI% = CI1%
'        DESCRIX$ = DESCRIT$(CI%)
'        If XVALO(XXCANT1$) < FRACMIN(CI%) Then
'          Call MENSERR(24, "Cantidad no Válida")
'          Text1.SelStart = 0
'          Text1.SelLength = Len(Text1)
'          GoTo 99
'        End If
'        If CODINT%(XXCODI$) <= 0 Then
'          Call MENSERR(24, "Código no Válido")
'          Text1.SelStart = 0
'          Text1.SelLength = Len(Text1)
'          GoTo 99
'        End If
        '
'Call REPLA(XXARTIDES$, "01-00006-15-B", 1, 16)
'Call REPLA(XXARTIDES$, "MIPRODUCTO", 18, 30)
'Call REPLA(XXARTIDES$, "1", 49, 8)
'Call REPLA(XXARTIDES$, "OF-004656-1", 61, 12) ' LOTE
'
'Call REPLA(XXARTIDES$, "887", 80, 9)   ' CLIENTE
'Call REPLA(XXARTIDES$, "3675", 90, 9) ' NRO PEDIDO
'Call REPLA(XXARTIDES$, "999", 100, 9) ' IDMOVISTO DEL CIERRE DE OF
        
        
        
'        Call REPLA(XXARTIDES$, XXCODI$, 1, 16)
'        Call REPLA(XXARTIDES$, DESCRIX$, 18, 30)
'        Call REPLA(XXARTIDES$, XXCANT$, 49, 8)
'        Call REPLA(XXARTIDES$, XXTARD1$, 61, 12) ' LOTE
'
'        Call REPLA(XXARTIDES$, SAFETEXT$(NC), 80, 9)   ' CLIENTE
'        Call REPLA(XXARTIDES$, SAFETEXT$(NROPED&), 90, 9) ' NRO PEDIDO
'        Call REPLA(XXARTIDES$, SAFETEXT$(IDMOVISTO&), 100, 9) ' IDMOVISTO DEL CIERRE DE OF
        '
'        List1.AddItem XXARTIDES$
'        Label18 = List1.ListCount
'        '
'90      Text1 = ""
'        Text1.SetFocus
'        '
      End If
    
    '
99 End Sub




VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CARCOBAR07 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2535
      ItemData        =   "CARCOBAR07.frx":0000
      Left            =   105
      List            =   "CARCOBAR07.frx":0002
      TabIndex        =   1
      Top             =   420
      Width           =   8205
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
      ItemData        =   "CARCOBAR07.frx":0004
      Left            =   105
      List            =   "CARCOBAR07.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   8205
   End
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
      Left            =   8400
      Picture         =   "CARCOBAR07.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cargar Items"
      Top             =   4470
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   4110
      TabIndex        =   0
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   4680
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4545
      Left            =   8400
      ScaleHeight     =   4485
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   -105
      Width           =   1170
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
         Picture         =   "CARCOBAR07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quitar  Producto de la Lista"
         Top             =   1260
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
         Picture         =   "CARCOBAR07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CARCOBAR07.frx":0766
         Top             =   4080
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CARCOBAR07.frx":2688
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CARCOBAR07.frx":28BC
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiquetas No Leídas (Armadores):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   2955
   End
   Begin VB.Label Label3 
      Height          =   135
      Left            =   5280
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   1470
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras (Artículo) >>>>>>>>:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   4725
      Width           =   4005
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Código                         Descripción                                                                   Cantidad"
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
      TabIndex        =   6
      Top             =   210
      Width           =   8200
   End
End
Attribute VB_Name = "CARCOBAR07"
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
        NUORFA$ = Label3
        NORCOMP& = XVALO(VALOBADA("ORDEFABR", "NUME_OCOM", "IDSUBOR='" & NUORFA$ & "'"))
        NUORITEM% = 1
        NPROV% = XVALO(Label2)
        Call SETULTREG&("!BOLREDET", 0)
        For U& = 1 To List1.ListCount
            '
            ETIQUETA$ = TRIM$(Mid$(List1.List(U& - 1), 100, 14))
            CI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            CANTI = XVALO(Mid$(List1.List(U& - 1), 57, 10))
            If NUORFA$ = "" Then
                CONDI$ = "NPRO_OCOM = " & NPROV% & " AND   Stat_Ocom < 8 AND Stat_Ocom >=0 "
                CONDI$ = CONDI$ + " AND Cod_Inte = " & CI% & " AND Cant_Rec < (Cant_Ocom - 0.001)  ORDER BY FEMI_OCOM, Nume_Ocom ASC"
                '
                NORCOMP& = XVALO(VALOBADA("OCOMDETA", "Nume_Ocom", CONDI$, "NOMESS"))
                NUORITEM% = XVALO(VALOBADA("OCOMDETA", "NUORD_ITEM", CONDI$, "NOMESS"))
                FEENTRE% = CVINT(VALOBADA("OCOMDETA", "Fentre_Sol", CONDI$, "NOMESS"))
                CANTOCOMPRA = XVALO(VALOBADA("OCOMDETA", "Cant_Ocom", CONDI$, "NOMESS"))
                CANTENTRE = XVALO(VALOBADA("OCOMDETA", "Cant_Rec", CONDI$, "NOMESS"))
                SALDO = CANTOCOMPRA - CANTENTRE
            End If
            '
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, Unimed$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CDbl(CANTI)), 89, 8)
            Call REPLA(Y$, Chr$(XVALO(NUORITEM%)), 98, 1)
            Call REPLA(Y$, ETIQUETA$, 241, 16)
            Call REPLA(Y$, MKS$(NORCOMP&), 151, 4)
            Call REPLA(Y$, MKD$(CDbl(CANTI)), 163, 8)
            Call REPLA(Y$, MKI$(FEENTRE%), 187, 2)
            '
            Call GRAREG("!BOLREDET", Y$, U&)
            
        Next U&
        '
        Call Command2_Click
99      Label3 = ""
        BOTNEXT.Enabled = True
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  Label2 = ""
  Label3 = ""
  Hide 'Unload Me
  Command2.Enabled = True
  '
End Sub
Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      ETIQUETA$ = Mid$(List1.List(List1.ListIndex), 100, 14)
      NO$ = Mid$(ETIQUETA$, 11, 4)
      Call REPLA(TTXX$, "Lectura Caja Nro. " & NO$, 1, 22)
      
      List1.RemoveItem (List1.ListIndex)
      List2.AddItem TTXX$
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
    If Len(TRIM$(Text1)) = 14 Then
       Call Text1_KeyPress(13)
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
        '
        MODOGABAL% = 0
        If InStr(1, EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
        'VALIDACIONES
        If KeyAscii <> 13 Then Exit Sub
        If TRIM$(Text1) = "" Then Exit Sub
        '
        'VERIFICO SI ES O-FABRICACION
        If UCase(Mid$(Text1, 1, 2)) = "OF" And MODOGABAL% = 1 Then
            ETIQUETA$ = UCase(TRIM$(Text1))
            'VALIDA SI LA ETIQUETA YA ESTA EN LA LISTA
            NROET% = XVALO(Mid$(ETIQUETA$, 11))
            If NROET% < 1 Then
                Call COMUNI("Número de Etiqueta Inválido")
                GoTo 50
            End If
            
            For II = 0 To List1.ListCount - 1
                TTXX$ = Mid$(List1.List(II), 100, 14)
                If ETIQUETA$ = TTXX$ Then
                    Call COMUNI("La Etiqueta Ya Se Encuentra en la Lista")
                    GoTo 50
                End If
            Next II
            NUORFA$ = TRIM$(Label3)
            NO$ = CStr(XVALO(Mid$(ETIQUETA$, 11)))
            If NUORFA$ = "" Then
                NUORFA$ = UCase(Mid$(Text1, 1, 2)) & "-" & Mid$(Text1, 3, 6) & "-" & Mid$(Text1, 9, 1)
                'VALIDA QUE COINCIDA EL PROVEEDOR DEL ARMADOR
                NPRO1% = XVALO(Label2)
                NPRO2% = XVALO(VALOBADA("OCOMDETA", "Npro_Ocom", "Doc_Orig='" & NUORFA$ & "'"))
                If NPRO1% <> NPRO2% Then
                    Call COMUNI("Imposible Continuar.\Proveedor Elegido No Se Corresponde con '" & NUORFA$ & "'")
                    GoTo 50
                End If
            End If
            'VALIDA SI LA ETIQUETA LEIDA CORRESPONDE A LA MISMA OF
            NUORFA0$ = UCase(Mid$(ETIQUETA$, 1, 2)) & "-" & Mid$(ETIQUETA$, 3, 6) & "-" & Mid$(ETIQUETA$, 9, 1)
            If NUORFA0$ <> NUORFA$ Then
                Call COMUNI("Imposible Continuar.\La Etiqueta Leida Corresponde a Otra OF")
                GoTo 50
            End If
            'CARGA LISTA DE ETIQUETAS PENDIENTES DE RECEPCION
            If List2.ListCount < 1 Then
                SQLX$ = "SELECT * FROM ORFAETI WITH(NOLOCK) WHERE IDSUBOR='" & NUORFA$ & "' AND STATUS<3 ORDER BY NUORDEN ASC"
                Set Rst = READSET(SQLX$)
                With Rst
                    Do While Not .EOF
                        ORDEN% = XVALO(!NUORDEN)
                        NOO$ = CStr(ORDEN%)
                        While Len(NOO$) < 4: NOO$ = "0" & NOO$: Wend
                        Call REPLA(TTXX$, "Lectura Caja Nro. " & NOO$, 1, 22)
                        List2.AddItem TTXX$
                        .MoveNext
                    Loop
                End With
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
            End If
            
            'VALIDA QUE LA ETIQUETA EXISTA
            If CantRegistros("ORFAETI", "IDSUBOR='" & NUORFA$ & "' AND NUORDEN=" & NO$) < 1 Then
                Call COMUNI("Imposible Continuar.\La Etiqueta Leída No Existe")
                GoTo 50
            End If

            'VALIDA QUE LA ETIQUETA NO ESTE INGRESADA
            If CantRegistros("ORFAETI", "IDSUBOR='" & NUORFA$ & "' AND NUORDEN=" & NO$ & " AND LOTERECE<>''") > 0 Then
                BRECE$ = SAFETEXT$(VALOBADA("ORFAETI", "LOTERECE", "IDSUBOR='" & NUORFA$ & "' AND NUORDEN=" & NO$))
                Call COMUNI("Imposible Continuar.\Etiqueta Ya Recepcionada con '" & BRECE$ & "'")
                GoTo 50
            End If
            'VALIDA QUE LA ETIQUETA NO ESTE INFORMADA O ANULADA
            If CantRegistros("ORFAETI", "IDSUBOR='" & NUORFA$ & "' AND NUORDEN=" & NO$ & " AND STATUS > 0") > 0 Then
                Call COMUNI("Imposible Continuar.\Etiqueta Ya Leída o Anulada")
                GoTo 50
            End If
            'VERIFICA QUE LA OF NO ESTE ANULADA
'            If CantRegistros("ORDEFABR", "IDSUBOR='" & NUORFA$ & "' AND STATUORF > 3") > 0 Then
'                Call COMUNI("Imposible Continuar.\Orden de Fabricación Cerrada o Anulada")
'                GoTo 50
'            End If
            CI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NUORFA$ & "'"))
            CANTI = LINDIVEN(CI%) 'TOMA LA CANTIDAD DE UNA CAJA O BULTO
            If TRIM$(Label3) = "" Then Label3 = NUORFA$
            GoTo 40
        End If
        '
        'VERIFICA QUE NO HAYA RECIBIDO MATERIALES DE ARMADORES. UNA MODALIDAD U OTRA.
        If TRIM$(Label3) <> "" Then
            Call COMUNI("Se Ha Recibido Material de Armadores.\Para Recibir Otro Tipo de Material Utilice Otra B/Recepcion")
            GoTo 50
        End If
        '
        'OBTENGO CODIGO INTERNO DE EXTERNO O UPC
        CODI$ = TRIM$(Text1)
        CI% = CODINT%(CODI$) ' CODIGO INTERNO DE CODIGO EXTERNO
        If CI% < 1 Then
          CI% = CODINTBARRA%(CODI$) ' CODIGO INTERNO DE UPC
        End If
        '
        'VALIDO QUE EL CODIGO SEA VALIDO
        If CI% < 1 Then
           Call MENSERR(24, "Código Inexistente o No Válido")
           Text1.SelStart = 0
           Text1.SelLength = Len(Text1)
           GoTo 60
        End If
        '
        'LEO LA CANTIDAD DE LOTE INDIVISIBLE, SI NO TIENE CARGADO > INPUTBOX >
        'GRABO LA CANTIDAD COMO LOTE INDIVISIBLE
        CANTI = LINDIVIS(CI%)
40      If CANTI < 1 Then
5             CANTI = XVALO(InputBox$("Cantidad de Unidades", "Ingrese Cantidad de Unidades"))
          If CANTI < 1 Then
             Call COMUNI("Debe Ingresar una Cantidad Válida")
             GoTo 5
          End If
          Call ACTUREGISTRO("MASTER", "LINDIVIS", "CODINT = " & CI%, FORMATNUM$(CANTI, "F16." & CNTDC$(CI%)), REGAF&, "NOMESS")
        End If
        '
        'ARMO LA FILA
        CODI2$ = TRIM$(CODEXT$(CI%))
        Call REPLA(TX$, CODI2$, 1, 16)
        Call REPLA(TX$, Left(DESCRIT$(CI%), 30), 18, 37)
        Call REPLA(TX$, FORMATNUM$(CANTI, "F10.1"), 57, 10)
        
        'SI SE TRATA DE OF NO ACUMULA Y GUARDA ETIQUETA PARA LUEGO DARLA POR CUMPLIDA
        If TRIM$(Label3) <> "" And MODOGABAL% = 1 Then
            NO$ = Mid$(ETIQUETA$, 11, 4)
            DESCRI$ = Mid$(DESCRIT$(CI%), 1, 25) & " (Caja:" & NO$ & ")"
            Call REPLA(TX$, DESCRI$, 18, 37)
            Call REPLA(TX$, ETIQUETA$, 100, 14)
            GoTo 45
        End If
        
        'COMPARAR SI EXISTE EN LA LISTA, EN ESE CASO ACUMULO SOBRE EL MISMO
        For U& = 1 To List1.ListCount
          If TRIM$(Left$(List1.List(U& - 1), 16)) = CODI2$ Then
             CANTLIST = XVALO(Mid$(List1.List(U& - 1), 57, 10))
             Call REPLA(TX$, FORMATNUM$((CANTLIST + CANTI), "F10.1"), 57, 10)
             List1.RemoveItem U& - 1
             List1.AddItem TX$, U& - 1
             GoTo 50
          End If
        Next U&
        '
45      List1.AddItem TX$
        'ELIMINA LA ETIQUETA DE LA LISTA DE PENDIENTES
        For II = 0 To List2.ListCount - 1
            t1$ = List2.List(II)
            If Mid$(t1$, 19, 4) = Mid$(ETIQUETA$, 11, 4) Then
                List2.RemoveItem (II)
                Exit For
            End If
        Next II
        '
50      Text1 = ""
60      Text1.SetFocus
    '
   End Sub

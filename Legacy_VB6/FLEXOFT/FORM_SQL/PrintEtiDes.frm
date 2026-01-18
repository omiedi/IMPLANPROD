VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PrintEtiDes 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrucciones de Despacho"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12650
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar Instrucciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   21
         ToolTipText     =   "Modificar Instrucciones del Pedido Activo"
         Top             =   600
         Width           =   2220
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar - Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9600
         TabIndex        =   20
         Top             =   5160
         Width           =   2940
      End
      Begin VB.TextBox TXTCANTBULTOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2760
         TabIndex        =   0
         Top             =   480
         Width           =   885
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         TabIndex        =   16
         Top             =   120
         Width           =   4095
         Begin VB.OptionButton PagoOrigen 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pago en Origen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   18
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton PagoDestino 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Pago en Destino"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   17
            Top             =   240
            Width           =   1160
         End
      End
      Begin VB.TextBox NroCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   11400
         TabIndex        =   10
         Top             =   600
         Width           =   1160
      End
      Begin VB.TextBox NroPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   11400
         TabIndex        =   9
         Top             =   240
         Width           =   1160
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4080
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   12420
         Begin VB.TextBox TXTLABEL 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TXTKILOSXBULTO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Index           =   0
            Left            =   1200
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   8040
            TabIndex        =   6
            Top             =   5000
            Width           =   2100
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Peso de C/Bulto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese la Cantidad de Bultos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N°C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   10920
         TabIndex        =   12
         Top             =   705
         Width           =   435
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N°P."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   10920
         TabIndex        =   11
         Top             =   345
         Width           =   435
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   3480
         Y2              =   9120
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   -120
         Visible         =   0   'False
         Width           =   9495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   7320
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "PrintEtiDes.frx":0000
      Top             =   4800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4800
      OleObjectBlob   =   "PrintEtiDes.frx":0234
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   135
      Left            =   6480
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "PrintEtiDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim TextBoxValues() As String

 Sub CrearMatrizTextBox(Cantidad As Integer)

   For F = 1 To TXTKILOSXBULTO.UBound
      Unload TXTKILOSXBULTO(F)
      Unload TXTLABEL(F)
   Next F
 
 

    
    ' Establecer la posición inicial de la matriz
    FilaInicial = TXTKILOSXBULTO(0).Top + TXTKILOSXBULTO(0).Height + 100 ' Ajusta la posición vertical según tus necesidades
    columnaInicial = TXTKILOSXBULTO(0).Left
    columnaInicial2 = TXTLABEL(0).Left
    
    ' Establecer la separación vertical entre los TextBox
    separacion = 330 ' Ajusta la separación vertical según tus necesidades
    H = 0
    For i = 1 To Cantidad
        ' Crear un nuevo TextBox
        J = TXTKILOSXBULTO.Count
        Load TXTKILOSXBULTO(J)
        TXTKILOSXBULTO(J).TEXT = ""
        TXTKILOSXBULTO(J).Top = FilaInicial + (J - H - 1) * separacion
        TXTKILOSXBULTO(J).Left = columnaInicial
        TXTKILOSXBULTO(J).Width = TXTKILOSXBULTO(J).Width
        TXTKILOSXBULTO(J).Visible = True
        Load TXTLABEL(J)
        TXTLABEL(J).TEXT = "Bto. " & J
        TXTLABEL(J).Top = FilaInicial + (J - H - 1) * separacion
        TXTLABEL(J).Left = columnaInicial2
        TXTLABEL(J).Width = TXTLABEL(J).Width
        TXTLABEL(J).Visible = True

        If J Mod 10 = 0 Then
          columnaInicial = columnaInicial + 2000: H = J
          columnaInicial2 = columnaInicial2 + 2000
        End If
    Next i
        

        
        ' Agregar el nuevo TextBox al formulario
'        Me.Controls.Add txtInicial(i)
End Sub

Private Sub Command1_Click()
    'esta llamada esta copiada de la modificacion de pedidos
    Entrega.NroCliente = XVALO(Me.NroCliente)
    Entrega.NroPedido = XVALO(Me.NroPedido)
    Entrega.command1.Visible = True 'para que pueda cambiar el nro de sucursal y trae los datos maestro
    Entrega.Command25.Visible = False
    'NO FILTRO POR QUE SI SE GUARDO CUANDO SE GENERO DEBE TENER LA POSIBILIDAD DE MOSTRARLO, PARA MODIFICAR.
    CONDI$ = "Nume_Cli =" & XVALO(Me.NroCliente)  'SI NO TIENE SUCURSALES NO ENTRA
    If CantRegistros&("SUCENTRE", CONDI$, "NOMESS") > 0 Then
        Entrega.Show 1
        If Entrega.Label1.Caption <> "OK" And Entrega.Label1.Caption <> "OK-SIN DATOS" Then Unload Entrega: GoTo 99
'                LENTREGA(2).TEXT = Entrega.DenoTransporte
'                Text20.TEXT = XVALO(Entrega.SucuEntrega)
'                Text19.TEXT = XVALO(Entrega.NumeTransporte)
'                LENTREGA(0).TEXT = Entrega.DomiTransporte.TEXT
'                LENTREGA(1).TEXT = Entrega.LocTransporte.TEXT
    End If
'   '
'   Call GRAPEDIDO(OKX%, MODOAPROB$)
'   '
'    If RetiraDomicilioAHP% > 0 Then
'       Call Entrega.BLANFORMU
'       Entrega.SucuEntrega = ""
'       Call Entrega.grabarDatosEntregaAhp(NUPED&, XVALO(Text1))
'       CONDI$ = "NROPED = " & NUPED&
'       CONDI$ = CONDI$ + " AND NROCLIE = " & XVALO(Text1)
'       Call ACTUREGISTRO("DATOENTREGA", "EntEnDomi", CONDI$, 1, REGAF&)
'       Unload Entrega
'    End If
     
    If Entrega.Label1.Caption = "OK" Then
         Call Entrega.grabarDatosEntregaAhp(XVALO(Me.NroPedido), XVALO(Me.NroCliente))
         Unload Entrega
    End If
99
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   NP& = XVALO(NroPedido)
   If NP& < 1 Then
      Call COMUNI("Falta Nro. de Pedido")
      GoTo 99
   End If
   '
   Call GRABAR(NP&)
   CANTIDADCOPIAS& = XVALO(TXTCANTBULTOS)
   'If CANTIDADCOPIAS& < 1 Then CANTIDADCOPIAS& = 1
   Call PrintEtiqDespachoAhp(NP&)
   Unload Me
99 Command4.Enabled = True
End Sub
Sub GRABAR(NROPED&)
  CONDI$ = "NROPED = " & NROPED&
  PORIG% = 0: PDEST% = 0
  If PagoOrigen.Value = True Then PORIG% = 1
  If PagoDestino.Value = True Then PDEST% = 1
  CNTBTO% = XVALO(TXTCANTBULTOS): If CNTBTO% < 1 Then CNTBTO% = 1
  Call ACTUREGISTRO("DatoEntrega", "PagoOrigen", CONDI$, PORIG%, REGAF&, "NOMESS")
  Call ACTUREGISTRO("DatoEntrega", "PagoDestino", CONDI$, PDEST%, REGAF&, "NOMESS")
  Call ACTUREGISTRO("DatoEntrega", "CantBultos", CONDI$, CNTBTO%, REGAF&, "NOMESS")
  '
  CONDI$ = "NROPED= " & NROPED& & " AND MARBORRA = 0"
  Call ACTUREGISTRO("DETABULTOS", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        SQLX$ = "SELECT * FROM DETABULTOS WHERE NUMEROBULTO < 1"
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With rs
        For i% = 1 To Me.TXTKILOSXBULTO.UBound
            .AddNew
            !PESOKG = XVALO(TXTKILOSXBULTO(i%))
            !NROPED = NROPED&
            !NROCLIE = XVALO(NroCliente)
            !NumeroBulto = i%
            !MarBorra = 0
            
            .Update
        Next i%
       End With
       rs.Close
       Set rs = Nothing
       CONDI$ = "NROPED= " & NROPED& & " AND MARBORRA = 1"
       Call BORRAREGISTROS("DETABULTOS", CONDI$, REGAF&, "NOMESS")

End Sub
Sub PrintEtiqDespachoAhp(NROPED&)
        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""
        PIMPRE$ = Control$(IDENTER$, "PUERETDE")
        If PIMPRE$ = "" Then PIMPRE$ = Control$("*", "PUERETDE") ' en SETUP DE PEDIDO

        Call SETPRINPORT(PIMPRE$)
        PagoDeOrigen$ = "_"
        PAGODEENDESTINO$ = "_"
        CONDI$ = "NROPED = " & NROPED&
        LETRAPROVINCIA$ = TRIM$(VALOBADA("DatoEntrega", "ProvEntre_cli", CONDI$, "NOMESS"))
        PROVINCIAX$ = TRIM$(ECOARCH$("PROVINC", LETRAPROVINCIA$))
        If PROVINCIAX$ = "" Then PROVINCIAX$ = "-"
        PO% = XVALO(VALOBADA("DatoEntrega", "PagoOrigen", CONDI$, "NOMESS")): If PO% > 0 Then PagoDeOrigen$ = "X"
        PD% = XVALO(VALOBADA("DatoEntrega", "PagoDestino", CONDI$, "NOMESS")): If PD% > 0 Then PAGODEENDESTINO$ = "X"
        FORIPRE$ = "ETDP-AHP"
        SQLX$ = "SELECT * FROM DetaBultos WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NROPED = " & NROPED&
        SQLX$ = SQLX$ + " ORDER BY NUMEROBULTO"
        Set rs = READSET(SQLX$)
        With rs
        Do While Not .EOF
            'FORMATO DE TEXTOQR$ 8 POSICIONES PARA EL PEDIDO Y 4 PARA EL NRO DE BULTO
            PESOKG$ = XVALO(!PESOKG)
            ORDENBULTO$ = TRIM$(FORMATOCON0$(XVALO(!NumeroBulto), 4))
            NroPedido1$ = FORMATOCON0$(XVALO(!NROPED&), 8)
            nroPedidoYOrdenBulto$ = TRIM$(NroPedido1$) & ORDENBULTO$
            '
            imag$ = "IMDSPAHP_QR.JPG"
            IDENTIFICADOR$ = "QR_DESPA_AHP"
            ANCHOX% = 0
            ALTOX% = 0
           On Error Resume Next
           Kill "C:\FLEXOFT\" & imag$
           TEXTOQR$ = nroPedidoYOrdenBulto$
           On Error GoTo 0
           'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
'           QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, imag$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
           QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", imag$, TEXTOQR$, "QR_CODE", 0, 0))
           Call T_Espera(3)
'           Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
           Call GRABAR_QR("C:\FLEXOFT\" & imag$, IDENTIFICADOR$)
            '
            FILTX$ = TRIM$(FORMATNUM$(NROPED&, "F10.0")) & ";" & ORDENBULTO$ & ";" & PESOKG$ & ";" & PagoDeOrigen$ & ";" & PAGODEENDESTINO$ & ";" & PROVINCIAX$
'           Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0, CANCOPIAS&)
            Call STDFORM(FORIPRE$, FILTX$, "", "", 0)
            .MoveNext
        
        Loop
        End With
        rs.Close
        Set rs = Nothing
        
        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""

End Sub

Private Sub NroPedido_Change()
   NROPED& = XVALO(NroPedido)
   CONDI$ = "NROPED= " & NROPED&
   NroCliente = XVALO(VALOBADA("DatoEntrega", "NROCLIE", CONDI$, "NOMESS"))
   PORIG% = XVALO(VALOBADA("DatoEntrega", "PagoOrigen", CONDI$, "NOMESS"))
   PDEST% = XVALO(VALOBADA("DatoEntrega", "PagoDestino", CONDI$, "NOMESS"))
   
   CANTBULTO% = XVALO(VALOBADA("DatoEntrega", "CantBultos", CONDI$, "NOMESS"))
   If CANTBULTO% < 1 Then TXTCANTBULTOS = "" Else TXTCANTBULTOS = CANTBULTO%
   If PORIG% > 0 Then PagoOrigen.Value = True
   If PDEST% > 0 Then PagoDestino.Value = True
   If PDEST% + PORIG% < 1 Then PagoDestino.Value = True 'SI NO TRAE DATOS POR DEFECTO PAGO DESTINO
 
   TXTCANTBULTOS.Locked = True 'PARA QUE MIENTRAS SE CARGUE NO HAYA CAMBIOS.
   For i% = 1 To XVALO(TXTCANTBULTOS)
     CONDI$ = "NROPED= " & NROPED&
     CONDI$ = CONDI$ + " and NumeroBulto =" & i%
     TXTKILOSXBULTO(i%) = XVALO(VALOBADA("DetaBultos", "pesokg", CONDI$, "NOMESS"))
   Next i%
   TXTCANTBULTOS.Locked = False
End Sub

Private Sub TXTCANTBULTOS_Change()
'   Call BorrarMatrizTextBox(TXTKILOSXBULTO()
   If XVALO(Me.TXTCANTBULTOS) > 60 Then
      Call COMUNI("Valor Ingresado Supera el Máximo de 60 ")
      Exit Sub
   End If
   Call CrearMatrizTextBox(XVALO(TXTCANTBULTOS))
End Sub

Private Sub TXTKILOSXBULTO_LostFocus(Index As Integer)
      TXTKILOSXBULTO(Index).TEXT = FORMATNUM$(XVALO(TXTKILOSXBULTO(Index)), "F9.1")
End Sub

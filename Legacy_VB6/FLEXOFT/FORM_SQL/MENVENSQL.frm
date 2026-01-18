VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MENVENSQL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Ctas. Ctes."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5235
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   20
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Control Cuenta Corriente"
      Height          =   225
      Left            =   2310
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   330
      Left            =   0
      TabIndex        =   17
      Top             =   3675
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   4320
      ScaleHeight     =   4545
      ScaleWidth      =   1425
      TabIndex        =   12
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   22
         Top             =   3480
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   23
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recu"
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
         Left            =   120
         Picture         =   "MENVENSQL.frx":0000
         TabIndex        =   19
         Top             =   2830
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "MENVENSQL.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Acceso a Pedidos y Presupuestos"
         Top             =   2170
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "MENVENSQL.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   870
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "MENVENSQL.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   230
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "MENVENSQL.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Acceso a Listas de Precios de Venta"
         Top             =   1520
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "GESTION DE VENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   120
      TabIndex        =   2
      Top             =   100
      Width           =   4065
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONTROL DE GESTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   3585
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   9
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "MENVENSQL.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Acceso al Control General de Gestión Comercial"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión Comercial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   11
               Top             =   105
               Width           =   1620
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "OPERACIONES COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   4
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "MENVENSQL.frx":0EBC
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir o Registrar un Recibo de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Recibo Cobranza en Cta. Corriente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   5
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   6
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "MENVENSQL.frx":1006
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar una Factura en Cuenta Corriente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación C.Cte"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   7
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.Label LabelFuenteAnchoFijo 
            Caption         =   "Para aplicar fuente ancho fijo"
            Height          =   255
            Left            =   720
            TabIndex        =   24
            Top             =   1440
            Visible         =   0   'False
            Width           =   2295
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MENVENSQL.frx":1448
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "MENVENSQL.frx":167C
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op_comerc 
      Caption         =   "Transacciones"
      Begin VB.Menu fac_ven 
         Caption         =   "Facturación de Ventas"
      End
      Begin VB.Menu recibo 
         Caption         =   "Recibo de Cobranza"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ctacte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu ctrlgeneral 
         Caption         =   "Control General de Gestión Comercial"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuListaDePrecios 
         Caption         =   "Lista de Precios"
      End
      Begin VB.Menu mnuPedidosDeClientes 
         Caption         =   "Pedidos de Clientes"
      End
      Begin VB.Menu mnuPresupuestos 
         Caption         =   "Presupuestos a Clientes "
      End
   End
End
Attribute VB_Name = "MENVENSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GENERAMAILDEUDA()
    '
    'VALIDO QUE EN EL DIA NO SE HAYA GENERADO
    If XVALO(Control("FAPECOB", "VERIENVI")) = HOY(HO$) Then Exit Sub
    '
    Screen.MousePointer = 11
    'LLAMO A LA RUTINA PARA QUE CARGUE EL ARCHIVO FAPECOBA CON LA COMPOSICION DE DEUDA
    Call FAPECOB07.FAPECOB2
   
    'ARMO ENCABEZADO
    TXTX1$ = "": TX$ = "": TOTALTEX$ = "": TX1$ = "": TOTALTEX$ = ""
    Call REPLA(TX1$, AJUSTD("Cta.", 6), 1, 6)
    Call REPLA(TX1$, AJUSTI("Cliente", 24), 8, 24)
    Call REPLA(TX1$, AJUSTI("Teléfono", 12), 34, 12)
    Call REPLA(TX1$, AJUSTI("Documento", 12), 48, 12)
    Call REPLA(TX1$, AJUSTD("Emisi.", 8), 62, 8)
    Call REPLA(TX1$, AJUSTD("Importe", 14), 72, 14)
    Call REPLA(TX1$, AJUSTD("Venci.", 8), 88, 8)
    Call REPLA(TX1$, AJUSTD("Saldo", 14), 98, 14)
    Call REPLA(TX1$, AJUSTD("Total", 14), 114, 14)
   
    SUBSALDEU# = 0
    'ARMO EL CUERPO DEL MAIL
    FIN& = ULTREG&("FAPECOBA")
    For i& = 1 To FIN&
       TX$ = Space(128)
       Y$ = REGLEIDO$("FAPECOBA", i&)
       NC% = CVI(Mid$(Y$, 1, 2))
      
       RACLI$ = rasocli$(NC%)
       TEL$ = TEL1CLI$(NC%)
       NFS$ = Mid$(Y$, 49, 12)
       FECHEMIX$ = FECHATEX$(CVI(Mid$(Y$, 61, 2)))
       IMORIG$ = FORMATNUM$(CVD(Mid$(Y$, 63, 8)), "F14.2")
       FVI$ = FECHATEX$(CVI(Mid$(Y$, 71, 2)))
       SALDEU# = CVD(Mid$(Y$, 73, 8))
       SUBSALDEU# = SUBSALDEU# + (SALDEU#)
       '
       Call REPLA(TX$, Str$(NC%), 1, 6)
       Call REPLA(TX$, RACLI$, 8, 24)
       Call REPLA(TX$, TEL$, 34, 12)
       Call REPLA(TX$, NFS$, 48, 12)
       Call REPLA(TX$, FECHEMIX$, 62, 8)
       Call REPLA(TX$, IMORIG$, 72, 14)
       Call REPLA(TX$, FVI$, 88, 8)
       Call REPLA(TX$, FORMATNUM$(SALDEU#, "F14.2"), 98, 14)
       Call REPLA(TX$, Space(14), 114, 14)
       '
       'VALIDO SI CAMBIA DE CLIENTE PARA QUE TOTALICE Y AGREGUE UNA RAYA
       If i& < FIN& Then
          NC_SIG% = CVI(Mid$(REGLEIDO$("FAPECOBA", i& + 1), 1, 2)) 'LEO EL SIGUIENTE
          If NC% <> NC_SIG% Then 'SI ES DIFERENTE AGREGO EL TOTAL DE ESTE CLIENTE
             Call REPLA(TX$, FORMATNUM(SUBSALDEU#, "F14.2"), 114, 14)
             TOTALDEU# = TOTALDEU# + SUBSALDEU# 'ACUMULO
             TXTX1$ = TXTX1$ + TX$ + vbCrLf
             TXTX1$ = TXTX1$ + String$(128, "-") + vbCrLf 'AGREGO LINEA DE SEPARACION
             SUBSALDEU# = 0
             GoTo 50
          End If
       Else
          'PARA EL ULTIMO CASO
          Call REPLA(TX$, FORMATNUM(SUBSALDEU#, "F14.2"), 114, 14)
          TOTALDEU# = TOTALDEU# + SUBSALDEU#
       End If
       '
       'AGREGO LA LINEA
       TXTX1$ = TXTX1$ + TX$ + vbCrLf
       '
50  Next i&
    '
    'ARMO EL TOTAL DEL TEXTO CON EL ENCABEZADO
    TOTALTEX$ = "LISTADO DE DEUDA DE CLIENTES" + vbCrLf + String$(28, "*") + vbCrLf + vbCrLf + TX1$ + vbCrLf + String$(128, "=") + vbCrLf ' ENCABEZADO
    TOTALTEX$ = TOTALTEX$ + TXTX1$ 'DETALLE
    
    'AGREGO TOTALES
    TOTALTEX$ = TOTALTEX$ + String$(128, "-") + vbCrLf
    TOTALTEX$ = TOTALTEX$ + String$(128, "-") + vbCrLf
    Call REPLA(TX2$, AJUSTI("TOTAL GENERAL", 24), 8, 24)
    Call REPLA(TX2$, FORMATNUM$(TOTALDEU#, "F14.2"), 98, 14)
    Call REPLA(TX2$, FORMATNUM$(TOTALDEU#, "F14.2"), 114, 14)
    TOTALTEX$ = TOTALTEX$ + TX2$ + vbCrLf
    TOTALTEX$ = TOTALTEX$ + String$(128, "-") + vbCrLf
    
    'GRABO EN ARCHIVO TXT PARA ADJUNTO
    Call SAVEFILE(LUCOM$ + "DEUCLITS.TXT", TOTALTEX$)
    
    'ARMO EL MAIL
    MAI_DESTI$ = TRIM$(EMPREAC$)
    MAI_DIREL$ = Control("FAPECOB", "MAIDEUDA")
    MAI_TEXTO$ = "Listado de Deuda de Clientes Enviado Como Adjunto (" + "DEUCLITS.TXT" + ")"
    MAI_ASUNT$ = "Lista de Deudas de Clientes"
    MAI_ADJUN$(0) = LUCOM$ + "DEUCLITS.TXT"
    '
    Call GENERAMAIL
    '
    'GRABO LA FECHA QUE SE GENERO EL PROCESO PARA QUE SE REALICE SOLO UNA VEZ POR DIA
    FE$ = TRIM$(Str$(HOY(HO$)))
    Call GRACONTROL("FAPECOB", "VERIENVI", FE$)
    
    Call COMUNI("Envio de Mail de Deuda de Clientes Realizado")
    Screen.MousePointer = 1

End Sub



'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    OPP% = COMALTER%("Gestión de Ventas y Deudores\\Minimizar\Cerrar\Cancelar")
    If OPP% = 1 Then
        WindowState = 1
      ElseIf OPP% = 2 Then
        Call FINAL("")
        Exit Sub
    End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("RECIBO", "Recibo de Cobranza") > 1 Then
        TC = TICAMONE(-1) ' fuerza refresco de tabla
        RECIBO08.Show 1
   End If
   Command14.Enabled = True
End Sub


Sub Command18_Click()
   Command18.Enabled = False
   Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
   Command18.Enabled = True
End Sub


Sub Command2_Click()
   Command2.Enabled = False
   OPX% = ALTERNA%("Presupuestos a Clientes\Pedidos de Clientes")
   If OPX% = 1 Then
        Call CONECRUN("PRESUP07", "Presupuestos a Clientes")
      ElseIf OPX% = 2 Then
        Call CONECRUN("PEDIDO07", "Pedidos de Clientes")
   End If
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    '
    MODU$ = "FACTURACION DE VENTAS"
    HOII% = HOY(HOS$)
    If EXISTE%(LUDAT$ + "VENCICAI.RFS") = 1 Then
      For UU& = 1 To ULTREG&("VENCICAI")
        XYZ$ = REGLEIDO$("VENCICAI", UU&)
        FEVENCAI% = CVI(Mid$(XYZ$, 63, 2))
        If FEVENCAI% > 0 Then
          If FEVENCAI% < HOII% Then
            DEFORMU$ = TRIM$(Left$(XYZ$, 32))
            Call MENSERR(24, "Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
            'Call MENSERR(24, "Imposible Acceder a Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
            'GoTo 99
          ElseIf FEVENCAI% < HOII% + 32 Then
                DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                Call MENSERR(24, "ATENCION: \CAI de " + DEFORMU$ + "\Próximo a Vencer !!!")
          End If
        End If
      Next UU&
    End If
    '
    TC = TICAMONE(-1) ' fuerza refresco de tabla
    On Error Resume Next
    FORFASQL.Show 1 ' PARA QUE PRESENTE EL EXISTENTE
    FORFASQL09.Show 1
    On Error GoTo 0
99  Command22.Enabled = True
End Sub

Private Sub Command27_Click()
      Command27.Enabled = False
      
      If DERACCE%("RECUCOMP", "Compensación de Recuperación de Comprobantes") < 2 Then
         GoTo 99
      End If
     '
     Screen.MousePointer = 11
     '
     Call PONEFECONTABLE
     
     Screen.MousePointer = 1

     REDOCEL.Show 1
99   Command27.Enabled = True
End Sub
Sub PONEFECONTABLE()

      Call ABRECONEXION
     'GENERAR RECORDSET TABLA FARECAFIP DE DOCUMENTOS RECUPERADOS
     'RECORRER Y ACTURALIZAR FECHA CONTABLE EN EL CAJABANC A 31/12/10
     '***************************************************************

     'CREAR UN CAMPO DE CONTROL DE REGISTROS A LOS CUALES SE LE MODIFICO LA FECHA
     Call CREACAMPO("", "FARECAFIP", "STATUS", 3, 2)
     Call PoneEnCeroNull("FARECAFIP", "STATUS", REGAF&, "NOMESS")
     CONDI$ = "STATUS < 1 "
     '
     SQLX$ = "SELECT * FROM FARECAFIP WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE " & CONDI$
     Set RSCABAN = READSET(SQLX$)
     FIN& = CantRegistros&("FARECAFIP")
     FECON = FETEXCOR1$(FECHANUM("31/12/10"))
     i& = 0
     With RSCABAN
       Do While Not .EOF
         i& = i& + 1
         Call TRACE(24, i&, FIN&)
         '
         NUFACTU$ = SAFETEXT$(!NFACTURA)
         NC% = XVALO(!NUCLIEN)
         '
         'ACTUALIZO FECHA CONTABLE EN CAJABANC
         CONDI$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NC%
         Call ACTUREGISTRO("CAJABANC", "FECHCONTA", CONDI$, FECON, REGAF&, "NOMESS")
         '
         'MARCO CON 1 PARA INDICAR QUE YA SE HA MODIFICADO LA FECHA DE ESTE REGISTRO
         Condi1$ = "NFACTURA = '" & NUFACTU$ & "' AND NUCLIEN = " & NC%
         Call ACTUREGISTRO("FARECAFIP", "STATUS", Condi1$, 1, REGAF&, "NOMESS")
         .MoveNext
       Loop
     End With
     RSCABAN.Close
     Set RSCABAN = Nothing

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   COGEVE07.Show 1
   'REMEVE07.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   'CONTROL PARA VERIFICAR SALDOS DE CUENTA CORRIENTE VENTAS(POR IMPORTACION DE DATOS (DOSIVAC))
   Exit Sub
   Command5.Enabled = False
   CLAX$ = "": CLACO% = 0
30 PPX% = Int(Rnd * 63)
   If PPX% < 48 Or PPX% > 57 Then GoTo 30
   CLAX$ = CLAX$ + Chr$(PPX%)
   If Len(CLAX$) < 6 Then GoTo 30
   '
   RESPUES$ = CONTRACLA$(CLAX$)
   CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
   If CONTRASE$ <> RESPUES$ Then GoTo 99
   '
   FEC% = FECHANUM("22/08/08")
   FECH1 = FETEXCOR1$(FEC%)
   FEC0% = FECHANUM("01/01/02")
   FECH0 = FETEXCOR1$(FEC0%)
   '
   Screen.MousePointer = 11
    
   SQLX$ = "SELECT COUNT (*) AS CANTI FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN'"
   SQLX$ = SQLX$ + " AND FECHEMISI >= '" & FECH0 & " '"
   SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FECH1 & " '"
   '
   Dim CONTRCOUNT As ADODB.Recordset
   Set CONTRCOUNT = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CONTRCOUNT
     CantiReg& = !CANTI
   End With
   CONTRCOUNT.Close
   Set CONTRCOUNT = Nothing
    
   SQLX$ = "SELECT * FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN'"
   SQLX$ = SQLX$ + " AND FECHEMISI >= '" & FECH0 & " '"
   SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FECH1 & " '"
   'SQLX$ = SQLX$ + " ORDER BY FECHEMISI"
   '
   Dim CONTROLTMP As ADODB.Recordset
   Set CONTROLTMP = New ADODB.Recordset
25 On Error Resume Next
   CONTROLTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With CONTROLTMP
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, CantiReg&)
        NC% = !NUCLIEN
        FecEmision% = CVINT%(!FECHEMISI)
        ImporteTotal# = XVALO(!idebecomp) - XVALO(!ihabecomp)
        nufact& = XVALO(Mid$(!codicom_lar, 9, 8))
        '
        RFF$ = RECFILT$("FACTUR", 1, MKI(NC%))
        '
        For U& = 1 To Len(RFF$) Step 4
           REG& = CVS(Mid$(RFF$, U&, 4))
           X$ = REGLEIDO$("FACTUR", REG&)
           If CVS(Mid$(X$, 3, 4)) = nufact& And CVI(Mid$(X$, 7, 2)) = CVINT(!FECHEMISI) Then
               TI% = Asc(Mid$(X$, 13, 1))
               If TI% < 0 Or TI% > 7 Then TI% = 2
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IMPO# = CVD(Mid$(X$, 15, 8))
               IMPO# = IMPO# + (CVD(Mid$(X$, 23, 8)))
               IMPO# = IMPO# + (CVD(Mid$(X$, 31, 8)))
               IMPO# = IMPO# + (CVD(Mid$(X$, 49, 8)))
               IMPO# = IMPO# + (CVD(Mid$(X$, 57, 8)))
               IMPO# = SG% * ROUND(IMPO#)
               Diferencia# = IMPO# - ImporteTotal#
               If Abs(Diferencia#) >= 0.005 Then
                  '
                  MsgBox "CLIENTE: " & NC% & " - COMPROBANTE: " & !codicom_lar
                  '
                  CODILAR$ = !codicom_lar
                  NCLI% = !NUCLIEN
                  FEMIX$ = FETEXCOR1$(CVINT(!FECHEMISI))
                  If !idebecomp > 0 Then
                      !idebecomp = !idebecomp + Diferencia#
                    Else
                      !ihabecomp = !ihabecomp - Diferencia#
                  End If
                  !ValAbsComp = Abs(!idebecomp + !ihabecomp)
                  .Update
                  '
                  SQLX$ = "SELECT * FROM CAJABANC"
                  SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'VENTA'"
                  SQLX$ = SQLX$ + " AND CodiCom_Lar = '" & CODILAR$ & "'"
                  SQLX$ = SQLX$ + " AND FECHEMISI = '" & FEMIX$ & "'"
                  SQLX$ = SQLX$ + " AND NuClien = " & NCLI% & " "
                  '
                  Dim CONTROL1 As ADODB.Recordset
                  Set CONTROL1 = New ADODB.Recordset
26                On Error Resume Next
                  CONTROL1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                  If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                  End If
                  On Error GoTo 0
                  '
                  If CONTROL1!idebecomp > 0 Then
                      CONTROL1!idebecomp = CONTROL1!idebecomp - Diferencia#
                    Else
                      CONTROL1!ihabecomp = CONTROL1!ihabecomp + Diferencia#
                  End If
                  CONTROL1!ValAbsComp = Abs(CONTROL1!idebecomp + CONTROL1!ihabecomp)
                  CONTROL1.Update
                  'On Error Resume Next
                  CONTROL1.Close
                  'On Error GoTo 0
                  Set CONTROL1 = Nothing
                  '
               End If
               GoTo 19
             End If
        Next U&
19    .MoveNext
      Loop

    End With
    CONTROLTMP.Close
    Set CONTROLTMP = Nothing

    Screen.MousePointer = 1
    MsgBox "FIN DE CONTROL"
    '\\\OT-08-0827-OD-FIN
    
99 Command5.Enabled = True
End Sub






Private Sub ctacte_Click()
    Call COGEVE07.Command7_Click
End Sub

Private Sub ctrlgeneral_Click()
    Call Command3_Click
End Sub

Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_pen_Click()
   Call COGEVE07.Command3_Click
End Sub

Private Sub fac_ven_Click()
   Call Command22_Click
End Sub

Private Sub Form_Activate()
   Call refresca_botones(Me)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   HOII% = HOY(HOS$)
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
   Call InstalaFuenteAnchoFijo
   'If ControlMenu%(Me.Name, Me) < 1 Then Call FINAL("")
    
   Call APERBASDAT("CABAN")
   '
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'      Caption = Caption + "  -  " + EPAC$
'   End If
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
   End If
   '
'VALIDACION POR SI NO PAGA
'If InStr(EPAC$, "DAVTECNICA") > 0 Then
'     FF% = FECHANUM("30/06/20")
'     If FF% <= HOY(HO$) Then
'       Call Form_Load
'     End If
'End If
   '
   Call GRATITFOR(Caption)
   '
   'SE ASIGNA A FORFINFAC PARA EL CASO DE LA LECTURA DEL FRM EN FORHASAR
   Set FORFINFAC = FINFACTUR07
   '
   Call ABRECONEXION
   Call CREATABLA_BCONTROL
   Call CREATABLA_STOCKS
   Call CREAR_STORED_PROCEDURE
   Call CREATABLA_ULTRECIBO
   EsPorPedido& = 0
   If Left$(APLINVO$, 6) = "PEDIDO" Then ' VIENE DE PEDIDO
        If XVALO(Mid$(APLINVO$, 7)) > 0 Then  'TIENE NUMERO DE PEDIDO
          EsPorPedido& = XVALO(Mid$(APLINVO$, 7))
          Call Command22_Click
        End If
   End If
   '
   If Left$(APLINVO$, 6) = "RECIBO" Then
       RECIBO08.Show 1
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "CONLISRE" Then  'OPCIONES DE REECIBO DE COBRANZA
       RECIBO08.Command8_Click
       Call FINAL("")
       
       
     ElseIf Left$(APLINVO$, 8) = "OPCIRECI" Then  'OPCIONES DE REECIBO DE COBRANZA
       RECIBO08.Command8_Click
       Call FINAL("")

     ElseIf Left$(APLINVO$, 8) = "ANURECOB" Then  'ANULACION DE RECIBO DE COBRANZA
       RECICO07.Command25_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "VERIMPREC" Then  'VER IMPRESION DE RECIBO DE COBRANZA
       RECICO07.Command23_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 10) = "IMPDESCRAN" Then  'IMPUTACION Y DESAPLICACION DE CREDITOS Y ANTICIPOS
       RECICO07.Command3_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "VISUIMPU" Then  'VISUALIZAR IMPUTACION DE CREDITOS Y ANTICIPOS
       RECICO07.Command11_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "AJUSALDE" Then  'AJUSTE DE SALDO DEUDORES
       RECICO07.Command4_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 9) = "COMCLIPRO" Then  'COMPENSACION CLIENTE PROVEEDOR
       RECICO07.Command10_Click
       Call FINAL("")
       
       
     ElseIf Left$(APLINVO$, 8) = "LIRECGE" Then  'LISTADO DE RECIBOS REGISTRADOS
       RECICO07.Command8_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "CONGENRE" Then  'CONFIGURACION GENERAL DE RECIBOS DE COBRANZA
       RECIBO08.Command6_Click
       Call FINAL("")
  
     ElseIf Left$(APLINVO$, 9) = "DETPORCUE" Then  'DETALLE POR CUENTA
       Call Command3_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 9) = "CONESTADO" Then  'CONTROL POR ESTADO
       Call Command3_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 7) = "FACTURA" Then
       FORFASQL.Show 1
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "CONLIFAC" Then
       Call Command22_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "OPCIFACT" Then  'OPCIONES VARIAS
       FORFASQL.Command20_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "ANUFACUEN" Then  'ANULACION DE COMPROBANTES
       OPVARFA07.Command45_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "CORFECEMI" Then  'CORREGIR FECHA DE EMISION
       OPVARFA07.Command3_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "LISCONDC" Then  'LISTADO DE CONTROL D.C.
       OPVARFA07.Command2_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "REMPENFA" Then  'REMITOS PENDIENTES DE FACTURACION.
       FORFASQL.Command8_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "CONGENFA" Then  'CONFIGURACION GENERAL - FACTURACION EN CUENTA CORRIENTE
       FORFASQL.Command5_Click
       Call FINAL("")
     
     ElseIf Left$(APLINVO$, 8) = "LPRECIOS" Then
       Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "CGVENTAS" Then
       COGEVE07.Show 1
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "RECUCOCLI" Then  'CONSULTA RESUMEN CUENTA CORRIENTE
       COGEVE07.Command7_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "LILICREDE" Then  'LIMITE DECREDITOS Y DEUDAS
       COGEVE07.Command20_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "CUECOBRAR" Then  'CUENTAS POR COBRAR
       COGEVE07.Command3_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 8) = "FLUCAPRO" Then  'FLUJO DE CAJA PROYECTADO
       COGEVE07.Command4_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "SALDEUCLI" Then  'SALDOS Y DEUDAS DE CLIENTES
       COGEVE07.Option2.Value = True
       COGEVE07.Command5_Click
       Call FINAL("")
       
      ElseIf Left$(APLINVO$, 8) = "AVAFACOB" Then  'AVANCE DE FACTURACION Y COBRANZA
       COGEVE07.Command9_Click
       Call FINAL("")
       
     ElseIf Left$(APLINVO$, 9) = "SUBESTVEN" Then  'RESUMEN MENSUALES SUBDIARIO Y ESTADISTICAS
       REMEVE07.Show 1
       Call FINAL("")

   End If
   '
   DoEvents
   CUEM$ = Control$("SVD451S", "CUEIVA")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   CUEM$ = Control$("SVD451S", "CUEADU")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   CUEM$ = Control$("SVD451S", "CUERET")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   For IXX% = 0 To 11
     CUEM$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(IXX%)))
     If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   Next IXX%
   '
   'If InStr(EMPREAC$, "MEDITEA") Then Command27.Visible = True
   '
   MAI_INT$ = Control("FAPECOB", "MAIDEUDA")
   If InStr(MAI_INT$, "@") > 0 Then
      Call GENERAMAILDEUDA
   End If
   Screen.MousePointer = 1
   GoTo 99
   '

50 Call MENSERR(24, "Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal.")
   'Call CONECRUN("FSETUP04/OPCLAVEN", "Configuración de Funcionamiento")
   OPCLAV07.Show 1
   Call FINAL("")
   '
99 End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub print1_Click()
   RECUEC07.Text1 = ""
   Call RECUEC07.Command3_Click
End Sub

Private Sub print2_Click()
   COGEVE07.Option2.Value = True
   Call COGEVE07.Command20_Click
End Sub

Private Sub print3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC07.MESAN = MESASEL$
      AVAFAC07.Option3.Value = True
      AVAFAC07.Option2.Value = True
      Call COGEVE07.Command9_Click
    End If
End Sub

Private Sub print4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC07.MESAN = MESASEL$
      AVAFAC07.Option4.Value = True
      AVAFAC07.Option2.Value = True
      Call COGEVE07.Command9_Click
    End If
End Sub

Private Sub print5_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option2.Value = True
      Call REMEVE07.Command3_Click
    End If
End Sub

Private Sub print6_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option2.Value = True
      Call REMEVE07.Command4_Click
    End If
End Sub

Private Sub print7_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option2.Value = True
      Call REMEVE07.Command5_Click
    End If
End Sub

Private Sub print8_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option2.Value = True
      Call REMEVE07.Command6_Click
    End If
End Sub

Private Sub print9_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option2.Value = True
      Call REMEVE07.Command2_Click
    End If
End Sub

Private Sub mnuListaDePrecios_Click()
    Call Command18_Click
End Sub

Private Sub mnuPedidosDeClientes_Click()
    Call CONECRUN("PEDIDO07", "Pedidos de Clientes")
End Sub

Private Sub mnuPresupuestos_Click()
    Call CONECRUN("PRESUP07", "Presupuestos a Clientes")
End Sub

Private Sub recibo_Click()
   Call Command14_Click
End Sub


Private Sub scr1_Click()
  Call COGEVE07.Command7_Click
End Sub

Private Sub scr2_Click()
  COGEVE07.Option1.Value = True
  Call COGEVE07.Command20_Click
End Sub

Private Sub scr3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC07.MESAN = MESASEL$
      AVAFAC07.Option3.Value = True
      AVAFAC07.Option1.Value = True
      Call COGEVE07.Command9_Click
    End If
End Sub

Private Sub scr4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC07.MESAN = MESASEL$
      AVAFAC07.Option4.Value = True
      AVAFAC07.Option1.Value = True
      Call COGEVE07.Command9_Click
    End If
End Sub

Private Sub scr5_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option1.Value = True
      Call REMEVE07.Command3_Click
    End If
End Sub

Private Sub scr6_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option1.Value = True
      Call REMEVE07.Command4_Click
    End If
End Sub

Private Sub scr7_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option1.Value = True
      Call REMEVE07.Command5_Click
    End If
End Sub

Private Sub scr8_Click()
   REMEVE07.Option1.Value = True
   Call REMEVE07.Command6_Click
End Sub

Private Sub scr9_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      REMEVE07.Option1.Value = True
      Call REMEVE07.Command2_Click
    End If
End Sub

Private Sub sub_iva_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      Call REMEVE07.Command21_Click
    End If
End Sub

Private Sub sub_ven_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE07.MESAN = MESASEL$
      Call REMEVE07.Command6_Click
    End If
End Sub





Private Sub Command4_Click()
20 Screen.MousePointer = 11
   Dim ACLIX%(32767)
   For U& = 1 To ULTREG&("CONSICLI")
      X$ = REGLEIDO$("CONSICLI", U&)
      NCLIE% = CVI(Left$(X$, 2))
      If NCLIE% > 0 Then
        If NCLIE% <= 32767 Then
          ACLIX%(NCLIE%) = 1
        End If
      End If
   Next U&
   '
   KK& = 0
   Call COPYSTRU("RECONCLI", "REPEFAC")
   For NCLIE% = 1 To 32760
     If ACLIX%(NCLIE%) > 0 Then
        Call RECONCLI(NCLIE%)
        For U& = 1 To ULTREG&("RECONCLI")
          X$ = REGLEIDO$("RECONCLI", U&)
          KK& = KK& + 1
          Call GRAREG("REPEFAC", X$, KK&)
        Next U&
     End If
   Next NCLIE%
   Call SETULTREG("REPEFAC", KK&)
   Call CIERRARCH("REPEFAC")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Call FINAL("*REPEFAC")
   '
End Sub
Sub RECONCLI(NCX7%)
   '
   Dim SALCONSI(32767)
   Call SETULTREG("RECONCLI", 0)
   Call SETULTREG("RECONMAT", 0)
   '
   If NCX7% > 0 Then
     If rasocli$(NCX7%) <> "" Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("CONSICLI")
       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: Next KLK&
       RFF$ = RECFILT$("CONSICLI", 1, MKI$(NCX7%))
       For U2& = 1 To Len(RFF$) Step 4   ' FIN&
         If U2& > Len(RFF$) Then Exit For
         U& = CVS(Mid$(RFF$, U2&, 4))
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("CONSICLI", U&)
         NCX8% = CVI(Left$(XX$, 2))
         If NCX8% = NCX7% Then
           CIXI% = CVI(Mid$(XX$, 35, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
             End If
           End If
         End If
       Next U2&
       '
       JJ& = 0
       FIN& = NUMAS%  ' NUINV%
       For U& = 1 To NUMAS%   ' NUINV%
         Call TRACE(20, U&, FIN&)
         'II$ = REGLEIDO$("INVERT", U& + 1)
         CIXI% = U&    ' CVI(Mid$(II$, 17, 2))
         'If CIXI% > 0 Then
         '  If CIXI% <= NUMAS% Then
             If SALCONSI(CIXI%) > 0 Then
               XX$ = REGBLAN$("RECONCLI")
               YY$ = REGBLAN$("RECONMAT")
               Call REPLA(XX$, MKI$(NCX7%), 1, 2)
               Call REPLA(XX$, MKI$(CIXI%), 3, 2)
               Call REPLA(XX$, MKS$(SALCONSI(CIXI%)), 5, 4)
               Call REPLA(XX$, String$(16, "."), 9, 16)
               Call REPLA(YY$, MKI$(CIXI%), 1, 2)
               Call REPLA(YY$, MKS$(SALCONSI(CIXI%)), 3, 4)
               JJ& = JJ& + 1
               Call GRAREG("RECONCLI", XX$, JJ&)
               Call GRAREG("RECONMAT", YY$, JJ&)
             End If
         '  End If
         'End If
       Next U&
       Call SETULTREG("RECONCLI", JJ&)
       Call CIERRARCH("RECONCLI")
       Call SETULTREG("RECONMAT", JJ&)
       Call CIERRARCH("RECONMAT")
       Screen.MousePointer = 1
       '
     End If
   End If

End Sub




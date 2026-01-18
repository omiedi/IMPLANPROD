VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CONTPEDI2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTROL TOTAL DE PEDIDOS PARA VALE DE MATERIALES"
   ClientHeight    =   8850
   ClientLeft      =   -3210
   ClientTop       =   465
   ClientWidth     =   15135
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8805
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15135
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Situación del Cliente y Pediddo Activo"
         Top             =   600
         Width           =   14880
      End
      Begin VB.TextBox TXTNROPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   10245
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1070
      End
      Begin VB.TextBox TXTNUMCLI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1500
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1070
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11325
         TabIndex        =   15
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   8730
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Nro. Pedido:"
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox TXTREFEPEDI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   11520
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   3525
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2565
         TabIndex        =   12
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   29
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Cuenta Cliente:"
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox TXTRASOCLI 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   4245
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   9
         Top             =   8280
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   8
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Situación General de los Pedidos del Cliente Activo"
         Top             =   4440
         Width           =   14880
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14865
         _ExtentX        =   26220
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         Appearance      =   0
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   3435
         Left            =   120
         TabIndex        =   6
         Top             =   4800
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   6059
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "ARCHIVO"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   " Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnudesap 
         Caption         =   "Des-Aplicar Informe de Inicio de Proceso"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNUARCHIVO2 
      Caption         =   "ARCHIVO2"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar2 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "CONTPEDI2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Public NORFAEXTERNA$

Sub Carga_Sit_General()
     MSF2.Rows = 1
     HOX% = HOY(HO$)

     SQLX$ = "SELECT PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NROPED, PEDDETA.FeSolEnt"
     SQLX$ = SQLX$ + " ,PEDDETA.NuOrItem,PEDDETA_ID,PEDDETA.CODIINT "
     SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE  NROCLIE = " & XVALO(TXTNUMCLI)
     SQLX$ = SQLX$ + " AND CanPed > CantEnt "
     SQLX$ = SQLX$ + "  AND STATUS < 8 "
     SQLX$ = SQLX$ + "  ORDER BY PEDDETA.NROPED "
     Set RS1 = READSET(SQLX$)
     With RS1
       Do While Not .EOF
         CI1% = XVALO(!CODIINT)
         CODI$ = SAFETEXT$(!CODIEXT)
         DESCRIX$ = SAFETEXT$(!Descrip)
         CANP = XVALO(!CanPed)
         CANEN = XVALO(!CANTENT)
         NOR% = XVALO(!NuOrItem)
         FESOL% = CVINT(!FeSolEnt)
         NP& = XVALO(!NROPED)
         j& = j& + 1
         MSF2.Rows = j& + 1
         MSF2.TextMatrix(j&, 1) = NP&
         MSF2.TextMatrix(j&, 2) = CODI$
         MSF2.TextMatrix(j&, 3) = DESCRIX$
         MSF2.TextMatrix(j&, 4) = CANP
         MSF2.TextMatrix(j&, 5) = CANEN
         MSF2.TextMatrix(j&, 6) = FECHATEX$(FESOL%)
         ATRAZO = 0
         If HOX% > FESOL% Then ATRAZO = DIENTRE%(FESOL%, HOX%)
         MSF2.TextMatrix(j&, 7) = ATRAZO
         MSF2.TextMatrix(j&, 8) = VmatPendientedeRemitir(CI1%, NP&)
         MSF2.TextMatrix(j&, 9) = 0
         MSF2.TextMatrix(j&, 10) = XVALO(!PEDDETA_ID)
         .MoveNext
       Loop
     End With
   RS1.Close
   Set RS1 = Nothing
   'CantidadPendienteSegunPedidoCodigo
   Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atrazo/F5"
   Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"

End Sub



Private Sub Command1_Click()

  Command1.Enabled = False
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/A10;Razon Social/a48", FRMZELECHO)
  FRMZELECHO.Caption = "Clientes Activos"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_DEUDOR12()
  Call Carga_MSF_Recordset(RS1, "Cuenta/A10;Razon Social/a48", FRMZELECHO.MSF2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  
  NCLI = XVALO(RESP_ZELE_VECT(1))
  If NCLI > 0 Then Me.TXTNUMCLI = NCLI
99 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO)
    FRMZELECHO.Caption = "Pedidos Pendientes de Entrega del Cliente:"
    '****************************************************************
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_PedidoPendxCte2(XVALO(Me.TXTNUMCLI))
    Call Carga_MSF_Recordset(RS1, "Nro.Pedido/f9;Referencia/a48", FRMZELECHO.MSF2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    NRO = XVALO(RESP_ZELE_VECT(1))
    If NRO > 0 Then TXTNROPED.TEXT = NRO
99  Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Call GUARDARVM(MSF1)
   Unload Me
End Sub

Sub GUARDARVM(MSF As MSFlexGrid)
   Call SETULTREG("!BOLREDET", 0)
   For i& = 1 To MSF.Rows - 1
      NPED# = XVALO(MSF.TextMatrix(i&, 1))
      Codigo$ = MSF.TextMatrix(i&, 2)
      CI1% = CODINT%(Codigo$)
      
      If CI1% > 0 Then
        CANTI# = MSF.TextMatrix(i&, 9)
        If CANTI# > 0 Then
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI1%), 83, 2)
            Call REPLA(Y$, Unimed$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CANTI#), 89, 8)
            Call REPLA(Y$, MKD$(NPED#), 163, 8)
            
            Call REGAPP("!BOLREDET", Y$)
        End If
      End If
   Next i&
   Call SETULTREG("!BOLREDET", i& - 1)
End Sub

Private Sub Command4_Click()
   Call GUARDARVM(MSF2)
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   Me.TXTNUMCLI.SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 1
   Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atraso/F5"
   Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"

   Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
      Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atraso/F5"
   Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"

   Call CARGA_ENCABEZADO(MSF1, Campos$, Me)

   
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub MSF1_DblClick()
    Unload VENTABNEW
    REG& = MSF1.MouseRow
    COL% = MSF1.MouseCol

    Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atrazo/F5"
    Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"
    
    VENTANA.Inicializar = 0
    VENTANA.UtilizaFormula = 1
'    VENTANA.Formula(8) = "6;*;7"
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Vale de Consumo"
    VENTANA.CampEditables = "9"
    VENTANA.Campox = Campos$
    VENTANA.Inicializar = 1
25  Call EDITARFILA(MSF1, VENTABNEW.MSF, REG&)

    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call PEGARFILAEDITADA(VENTABNEW.MSF, MSF1, REG&)
       If XVALO(VENTABNEW.MSF.TextMatrix(1, 9)) <= 0.05 Then
            Call COMUNI("Cantidad Ingresada No Vàlida (en fila: " & i& & ")")
            GoTo 30
         End If
    End If
    '
    Set VENTANA = Nothing
End Sub
Function VALIDAR_CAMBIOS(MSFBASE As MSFlexGrid, MSFMODI As MSFlexGrid)
'
'   ' "Código/a16;Descripción/A32;Q.Pedida/f9;Q.Entreg./F9;O.Fabric./a12;Q.Proy./f9;Q.Aprob./f9;Orden/F0"
'   'SI LA CANTIDAD PEDIDA ES < A LA ENTREGADA NO SE PUEDE
'   With MSFMODI
'     For i& = 1 To .Rows - 1
'       CANTPED = XVALO(.TextMatrix(i&, 3))
'       CANTENT = XVALO(.TextMatrix(i&, 4))
'       If CANTPED < CANTENT Then
'          Call COMUNI("En Fila: " & i& & " Imposible Cambiar la Cantidad Pedida a un Valor Menor que la entregada")
'          VALIDAR_CAMBIOS = -1
'          Exit Function
'       End If
'     Next i&
'   End With
'   '
'    With MSFMODI
'     For i& = 1 To .Rows - 1
'       CANTPED = XVALO(.TextMatrix(i&, 3))
'       CANPROY = XVALO(.TextMatrix(i&, 6))
'       If CANTPED > CANPROY Then
'          OPX% = COMALTER("En Fila: " & i& & " Se Han Detectado Aumentos en Cantidad Pedida\\Mantener Cantidad a Fabricar\Aumentar en Mismo Modo La Cantidad a Fabricar")
'          If OPX% = 2 Then
'           ' VER QUE HACER POR QUE AL ACTUALIZAR PUEDE QUE YA TENGA PARCIALES APROBADOS, Y QUE PASARIA CON LAS RESERVAS.
'          End If
'          Exit Function
'       End If
'     Next i&
'   End With
'
'   'SI LA CANTIDAD PEDIDA ES > A LA ENTREGADA OK
'      ' CONSULTAR SI ACTUALIZA  LA CANTIDAD A FABRICAR
'
'   'SI LA CANTIDAD PEDIDA ES < A LA CANTIDAD FABRICADA AVISAR QUE LA MISMA ES PARA STOCK Y DAR POR CUMPLIDA LA ORDEN
End Function
Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF1.MouseRow
     COL% = MSF1.MouseCol
     Call MENU_CELDA(MSF1, REG&, COL%)
  End If
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnudesap.Visible = Not (mnudesap.Visible)
'  InformaCantidadEnProceso.Visible = Not InformaCantidadEnProceso.Visible
  '
End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub

        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCHIVO
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

''****************************************************************
Private Sub mnucopiar2_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub


Private Sub MSF2_DblClick()
    Unload VENTABNEW
    REG& = MSF2.MouseRow
    COL% = MSF2.MouseCol

    Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atrazo/F5"
    Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"
    
    VENTANA.Inicializar = 0
    VENTANA.UtilizaFormula = 1
'    VENTANA.Formula(8) = "6;*;7"
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Vale de Consumo"
    VENTANA.CampEditables = "9"
    VENTANA.Campox = Campos$
    VENTANA.Inicializar = 1
25  Call EDITARFILA(MSF2, VENTABNEW.MSF, REG&)

    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call PEGARFILAEDITADA(VENTABNEW.MSF, MSF2, REG&)
       If XVALO(VENTABNEW.MSF.TextMatrix(1, 9)) <= 0.05 Then
            Call COMUNI("Cantidad Ingresada No Vàlida (en fila: " & i& & ")")
            GoTo 30
         End If
    End If
    '
  
    Set VENTANA = Nothing
End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF2.MouseRow
     COL% = MSF2.MouseCol
     Call MENU_CELDA2(MSF2, REG&, COL%)
  End If
End Sub

Sub INVICTRL2()
  '
  mnucopiar2.Visible = Not (mnucopiar2.Visible)
  '
End Sub
Sub MENU_CELDA2(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL2
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCHIVO2
        Call INVICTRL2
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub


Private Sub TXTNROPED_Change()
     Me.TXTREFEPEDI = VALOBADA("PEDENCA", "REFERENCIA", "NROPED = " & XVALO(Me.TXTNROPED), "NOMESS")
     Call CARGAGRILLADETALLE(XVALO(TXTNROPED))
End Sub
Sub CARGAGRILLADETALLE(NROPED&)

     Dim obj As New RECORXET
     MSF1.Rows = 1
     If NROPED& < 1 Then Exit Sub
     HOX% = HOY(HO$)

     SQLX$ = "SELECT PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NROPED, PEDDETA.FeSolEnt"
     SQLX$ = SQLX$ + " ,PEDDETA.NuOrItem,PEDDETA_ID,PEDDETA.CODIINT "
     SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) "
     'SQLX$ = SQLX$ + " WHERE  NROCLIE = " & XVALO(TXTNUMCLI)
     SQLX$ = SQLX$ + " WHERE  NROPED =  " & NROPED&
     SQLX$ = SQLX$ + " AND CanPed > CantEnt "
     SQLX$ = SQLX$ + "  AND STATUS < 8 "
     SQLX$ = SQLX$ + "  ORDER BY PEDDETA.NROPED "
     Set RS1 = READSET(SQLX$)
     With RS1
       Do While Not .EOF
         CI1% = XVALO(!CODIINT)
         CODI$ = SAFETEXT$(!CODIEXT)
         DESCRIX$ = SAFETEXT$(!Descrip)
         CANP = XVALO(!CanPed)
         CANEN = XVALO(!CANTENT)
         NOR% = XVALO(!NuOrItem)
         FESOL% = CVINT(!FeSolEnt)
         NP& = XVALO(!NROPED)
         j& = j& + 1
         MSF1.Rows = j& + 1
         MSF1.TextMatrix(j&, 1) = NP&
         MSF1.TextMatrix(j&, 2) = CODI$
         MSF1.TextMatrix(j&, 3) = DESCRIX$
         MSF1.TextMatrix(j&, 4) = CANP
         MSF1.TextMatrix(j&, 5) = CANEN
         MSF1.TextMatrix(j&, 6) = FECHATEX$(FESOL%)
         ATRAZO = 0
         If HOX% > FESOL% Then ATRAZO = DIENTRE%(FESOL%, HOX%)
         MSF1.TextMatrix(j&, 7) = ATRAZO
         MSF1.TextMatrix(j&, 8) = VmatPendientedeRemitir(CI1%, NP&)
         MSF1.TextMatrix(j&, 9) = 0
         MSF1.TextMatrix(j&, 10) = XVALO(!PEDDETA_ID)
         .MoveNext
       Loop
     End With
   RS1.Close
   Set RS1 = Nothing
   'CantidadPendienteSegunPedidoCodigo
   Campos$ = "N.Ped./F8;Código/a16;Descripción/A24;Q.Ped/F9;Q.Entr./F9;F.Sol./D8;Atrazo/F5"
   Campos$ = Campos$ + ";V.M.Pendiente/F9.0;V.M.a Informar/F12.0;PEDDETA_ID/A0"
   
End Sub

Private Sub TXTNUMCLI_Change()
   TXTRASOCLI = rasocli$(XVALO(Me.TXTNUMCLI))
   Call Carga_Sit_General
End Sub
Private Sub TXTNUMCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.TXTNROPED.SetFocus
   End If
End Sub



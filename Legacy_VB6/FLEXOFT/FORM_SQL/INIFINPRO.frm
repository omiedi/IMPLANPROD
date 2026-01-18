VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INIFINPRO 
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Infomes de Inicio y Fin de Proceso"
   ClientHeight    =   9060
   ClientLeft      =   -3216
   ClientTop       =   768
   ClientWidth     =   15372
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   15372
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0D0D0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8805
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Ordenes de Fabricación con Inicio de Proceso"
         Top             =   375
         Width           =   14865
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Seleccionar Orden de Fabricación para Inicio de Proceso"
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
         Left            =   9800
         TabIndex        =   8
         Top             =   4080
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar Informes con Fin de Proceso"
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
         Left            =   10280
         TabIndex        =   7
         Top             =   8280
         Width           =   4695
      End
      Begin VB.TextBox TXTMOVIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
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
         Height          =   330
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Ordenes de Fabricación con Fin de Proceso"
         Top             =   4575
         Width           =   14865
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   14865
         _ExtentX        =   26226
         _ExtentY        =   5630
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   3195
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Doble Click Quita de Lista de Fin de Procesos (Solo Si No Tiene Cierres de O.F.)"
         Top             =   4920
         Width           =   14865
         _ExtentX        =   26226
         _ExtentY        =   5630
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
         BackColor       =   &H00C0D0D0&
         Caption         =   "TEPRUEBA"
         Height          =   192
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   876
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0D0D0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "ARCHIVO"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu InformaCantidadEnProceso 
         Caption         =   "Informar Cantidad en Proceso"
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
      Begin VB.Menu modfinproc 
         Caption         =   "Modificar Informe de Fin de Proceso"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnureporte 
      Caption         =   "Reportes"
      Begin VB.Menu sininformar 
         Caption         =   "Ordenes de Fabricación Sin Informar "
      End
      Begin VB.Menu inicioinfo 
         Caption         =   "Ordenes de Fabricación Informado  su inicio"
      End
      Begin VB.Menu cantacerrar 
         Caption         =   "Ordenes de Fabricación Informada Cantidad lista para Cerrar"
      End
   End
   Begin VB.Menu pp 
      Caption         =   "Planilla de Producción"
      Visible         =   0   'False
      Begin VB.Menu geplapro 
         Caption         =   "Generar Planilla de Producción"
      End
      Begin VB.Menu viplage 
         Caption         =   "Visualizar Planillas Generadas"
      End
   End
End
Attribute VB_Name = "INIFINPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Public NORFAEXTERNA$


Sub cargarCantfinProceso()
    'Generar un proceso (cargarCantfinProceso) que recorra las ordenes de fabricacion pendientes en los casos que la cantidad
    'aprobada sea mayor a 0 que ponga  complete el campo canti_fin_proceso = cantidad aprobada y el campo Status_procesos = 2
    Screen.MousePointer = 11
    Call PoneEnCeroNull("ORDEFABR", "canti_fin_proceso", REGAF&, "NOMESS")
   
    SQLX$ = "Select CanApro,canti_fin_proceso,StatuOrf ,Status_Proceso from ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf < 3 AND canti_fin_proceso < 1"
    SQLX$ = SQLX$ + "  AND CanApro > 0 "
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With RS
      Do While Not .EOF
        !canti_fin_proceso = XVALO(!CANAPRO)
        !Status_Proceso = 2
        .Update
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    Screen.MousePointer = 1
End Sub

Sub CARGAROFABPROCESOSINFORMADOS()
   MSF2.Rows = 1
   SQLX$ = "Select * from ORDEFABR WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Status_Proceso = 2 or Canti_fin_proceso > 0 order by IDSUBOR"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       MSF2.Rows = MSF2.Rows + 1
       MSF2.TextMatrix(MSF2.Rows - 1, 1) = SAFETEXT$(!IDSUBOR)
       MSF2.TextMatrix(MSF2.Rows - 1, 2) = SAFETEXT$(!COD_EXTE)
       MSF2.TextMatrix(MSF2.Rows - 1, 3) = SAFETEXT$(!Descrip)
       MSF2.TextMatrix(MSF2.Rows - 1, 4) = TRIM$(FORMATNUM$(XVALO(!CANPROY), "F10.1"))
       MSF2.TextMatrix(MSF2.Rows - 1, 6) = TRIM$(FORMATNUM$(XVALO(!CANAPRO), "F10.1"))
       MSF2.TextMatrix(MSF2.Rows - 1, 5) = TRIM$(FORMATNUM$(XVALO(!canti_fin_proceso), "F10.1"))
       MSF2.TextMatrix(MSF2.Rows - 1, 7) = Format(!FinDeProceso, "dd/mm/yy hh:mm:ss")
       
       MSF2.TextMatrix(MSF2.Rows - 1, 8) = TRIM$(FORMATNUM$(XVALO(MSF2.TextMatrix(MSF2.Rows - 1, 5)) - XVALO(MSF2.TextMatrix(MSF2.Rows - 1, 6)), "F10.1"))
       MSF2.TextMatrix(MSF2.Rows - 1, 9) = TRIM$(FORMATNUM$(XVALO(MSF2.TextMatrix(MSF2.Rows - 1, 4)) - XVALO(MSF2.TextMatrix(MSF2.Rows - 1, 5)), "F10.1"))
       
       .MoveNext
     Loop
   End With

End Sub

Private Sub mnucopiar2_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub modfinproc_Click()
    CantidadOriginal$ = TRIM$(MSF2.TextMatrix(GLOBALFRM_FILA&, 4))
    CantidadInformadaActual$ = TRIM$(MSF2.TextMatrix(GLOBALFRM_FILA&, 5))
    CantidadInformada$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad Total Informada", CantidadInformadaActual$))

    CantOrig# = XVALO(CantidadOriginal$)
    CantInfo# = XVALO(CantidadInformada$)
    If CantInfo# > CantOrig# Or CantInfo# = 0 Then
        Call MENSERR(24, "Falta Cantidad Informada o la Misma Supera la Original")
        Exit Sub
    End If

    NROOF$ = TRIM$(MSF2.TextMatrix(GLOBALFRM_FILA&, 1))
    CONDI$ = "IDSUBOR = '" & NROOF$ & "'"
    FechaFinProcesoDefault% = HOY(HO$)
    VDEF$ = MKI$(FechaFinProcesoDefault%) & Format(Now, "hh:mm")
    OBX$ = OBJPLA$("FECHAHORAFINPROC", VDEF$)
    
    If OBX$ = "" Then Exit Sub

    FechaFinProceso% = CVI(Mid$(OBX$, 1, 2))
    If FechaFinProceso% = 0 Or FechaFinProceso% < FechaFinProcesoDefault% Then
        Call MENSERR(24, "Fecha de Fin de Proceso Inválida o Menor a la de Inicio")
        Exit Sub
    End If
    HoraFinProceso = IsDate(Mid$(OBX$, 3, 5) & ":00")
    If HoraFinProceso = False Then
        Call MENSERR(24, "Hora de Fin de Proceso Inválida")
        Exit Sub
    End If
    
    If XVALO(CantidadInformada$) >= XVALO(MSF2.TextMatrix(GLOBALFRM_FILA&, 6)) Then
       MSF2.TextMatrix(GLOBALFRM_FILA&, 5) = FORMATNUM$(XVALO(CantidadInformada$), "F10.1")
       MSF2.TextMatrix(GLOBALFRM_FILA&, 7) = FECHATEX(FechaFinProceso%) & Space(1) & HORATEX$(HORANUM(Mid$(OBX$, 3, 5)))
       MSF2.TextMatrix(GLOBALFRM_FILA&, 8) = TRIM$(FORMATNUM$(XVALO(MSF2.TextMatrix(GLOBALFRM_FILA&, 5)) - XVALO(MSF2.TextMatrix(GLOBALFRM_FILA&, 6)), "F10.1"))
       MSF2.TextMatrix(GLOBALFRM_FILA&, 9) = TRIM$(FORMATNUM$(XVALO(MSF2.TextMatrix(GLOBALFRM_FILA&, 4)) - XVALO(MSF2.TextMatrix(GLOBALFRM_FILA&, 5)), "F10.1"))
    Else
       Call COMUNI("Imposible Modificar !!!\Cantidad Informada Como Fin de Proceso es Menor a la Cantidad Ya Aprobada")
    End If
End Sub

Private Sub MSF1_Click()
  REG& = MSF1.MouseRow
  COL% = MSF1.MouseCol
  If COL% = 0 Then
     On Error Resume Next
     Call COLOREAR_GRILLA_SOLAFILA(INFINI09.MSF1, vbYellow, REG&)
     On Error GoTo 0
  End If
End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSF2.MouseRow
     COL% = MSF2.MouseCol
     Call MENU_CELDA2(MSF2, REG&, COL%)
  End If
End Sub
Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSF1.MouseRow
     COL% = MSF1.MouseCol
     Call MENU_CELDA(MSF1, REG&, COL%)
  End If
End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu mnuarchivo
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
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

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnudesap.Visible = Not (mnudesap.Visible)
  InformaCantidadEnProceso.Visible = Not InformaCantidadEnProceso.Visible

  '
End Sub
Sub INVICTRL2()
  '
  mnucopiar2.Visible = Not (mnucopiar2.Visible)
  '
End Sub

Private Sub mnudesap_Click()
  NORFA$ = TRIM$(MSF1.TextMatrix(GLOBALFRM_FILA&, 1))
  CONDI$ = "IDSUBOR= '" & NORFA$ & "'"
  INFOFIN = XVALO(VALOBADA("ORDEFABR", "Canti_fin_proceso", CONDI$, "NOMESS"))
  If INFOFIN > 0 Then
      Call COMUNI("Imposible Revertir Orden de Fabricación Tiene Informados Fin de Proceso\Modifique los Mismos y Vuelva a Intentar")
  Else
     'DESAPLICACION DE INFORME DE INICIO DE PROCESO
     FECHX$ = FETEXCOR1$(FECHANUM("01/01/1980"))
     Call ACTUREGISTRO("ORDEFABR", "Status_Proceso", CONDI$, 0, REGAF&, "NOMESS")
     Call ACTUREGISTRO("ORDEFABR", "inicio_informado", CONDI$, FECHX$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("ORDEFABR", "Canti_Iniciada", CONDI$, 0, REGAF&, "NOMESS")
     
     If MSF1.Rows > 2 Then
        Call MSF1.RemoveItem(GLOBALFRM_FILA&)
     Else
        MSF1.Rows = 1
     End If
     TXTX$ = "Se Borra Inicio de Proceso"
     Call AGREGA_ANOTAORFAB_NEW(NORFA$, TXTX$)

  End If
End Sub
Private Sub InformaCantidadEnProceso_Click()
    CANTIGRILLA$ = TRIM$(MSF1.TextMatrix(GLOBALFRM_FILA&, 10))
    CANTIENPROCESO$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad Total Informada", CANTIGRILLA))
    If XVALO(CANTIENPROCESO$) > 0 Then
       NROOF$ = MSF1.TextMatrix(GLOBALFRM_FILA&, 1)
       MSF1.TextMatrix(GLOBALFRM_FILA&, 10) = FORMATNUM$(XVALO(CANTIENPROCESO$), "F10.1")
       CONDI$ = "IDSUBOR = '" & NROOF$ & "'"
       Call ACTUREGISTRO("ORDEFABR", "Canti_Iniciada", CONDI$, CANTIENPROCESO$, REGAF&, "NOMESS")
    Else
       Call COMUNI("Imposible Modificar !!!\Cantidad Informada Debe ser Mayor a 0")
    End If
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Sub CargarOFabEnProceso(Optional ORFA$)
     
   'SI SE LE LLAMA CON PARAMETRO DE ORDEN DE FABRIACION AGREGA EL REGISTRO DE DICHA ORDEN
   'SI ORFA$ ESTA VACIO LLENA TODA LA GRILLA
   DESDE = FETEXCOR1$(FECHANUM("01/01/1980"))
   SQLX$ = "Select * from ORDEFABR WITH(NOLOCK)"
   If ORFA$ = "" Then
       MSF1.Rows = 1
       SQLX$ = SQLX$ + " WHERE Status_Proceso = 1"
   Else
       SQLX$ = SQLX$ + " WHERE IDSUBOR  = '" & ORFA$ & "'"
   End If
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
      
       MSF1.Rows = MSF1.Rows + 1
       MSF1.TextMatrix(MSF1.Rows - 1, 1) = SAFETEXT$(!IDSUBOR)
       MSF1.TextMatrix(MSF1.Rows - 1, 2) = SAFETEXT$(!COD_EXTE)
       MSF1.TextMatrix(MSF1.Rows - 1, 3) = SAFETEXT$(!Descrip)
       MSF1.TextMatrix(MSF1.Rows - 1, 4) = FORMATNUM$(XVALO(!CANPROY), "F10.1")
       If ORFA$ = "" Then
         MSF1.TextMatrix(MSF1.Rows - 1, 5) = FORMATNUM$(XVALO(!canti_fin_proceso), "F10.1")
         MSF1.TextMatrix(MSF1.Rows - 1, 6) = FORMATNUM$(XVALO(!CANAPRO), "F10.1")
       Else
         'SI SE LE LLAMA CON PARAMETRO DE ORDEN DE FABRIACION  ESTOS CAMPOS DEBEN QUEDAR VACIOS
         MSF1.TextMatrix(MSF1.Rows - 1, 5) = ""
         MSF1.TextMatrix(MSF1.Rows - 1, 6) = ""
       End If
       
       MSF1.TextMatrix(MSF1.Rows - 1, 7) = Format(!inicio_informado, "dd/mm/yy hh:mm:ss")
       MSF1.TextMatrix(MSF1.Rows - 1, 10) = TRIM$(FORMATNUM$(XVALO(!Canti_Iniciada), "F10.1"))

       .MoveNext
     Loop
   End With
   
   'si no esta cargado el frm lo carga
   If Not IsFormLoaded(Me) Then
     Load Me
     Me.Visible = True
   End If

End Sub
Private Sub Command1_Click()

     VDEF$ = ""
     If NORFAEXTERNA$ <> "" Then  'VARIABLE A NIVEL DE FORMULARIO
       Call REPLA(VDEF$, NORFAEXTERNA$, 1, 12)
     End If
     '
     NORFAEXTERNA$ = "" 'LIMPIO LA VARIABLE PARA NO QUEDARME CON INFO
     NROOF$ = TRIM$(OBJPLA$("SELORFAP", VDEF$))
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     For i& = 1 To MSF1.Rows - 1
        OFGRILLA$ = TRIM$(MSF1.TextMatrix(i&, 1))
        If OFGRILLA$ = NROOF$ Then
           Call COMUNI("La Orden de Fabricación Seleccionada Cuenta con Informe de Inicio de Proceso")
           MSF1.TopRow = i&
           MSF1.Row = i&
           MSF1.ColSel = 1
           On Error Resume Next
           MSF1.SetFocus
           On Error GoTo 0
           Exit Sub
        End If
     Next i&
     '
     For i& = 1 To MSF2.Rows - 1
        OFGRILLA$ = TRIM$(MSF2.TextMatrix(i&, 1))
        If OFGRILLA$ = NROOF$ Then
            Call COMUNI("La Orden de Fabricación Seleccionada Cuenta con Fin de Proceso\ \Debe Quitarla de la Lista Inferior\para Modificar su Inicio de Proceso.")
            Exit Sub
'           OPX% = COMALTER%("La Orden de Fabriación Seleccionada Yá Tiene Informados Fin de Proceso\\Cancelar\Informar Comienzo de Proceso\Corregir el Mismo")
'           If OPX% = 3 Then  'SE POSICIONA EN LA FILA
'             MSF2.TopRow = I&
'             MSF2.Row = I&
'             MSF2.ColSel = 1
'             On Error Resume Next
'             MSF2.SetFocus
'             On Error GoTo 0
'             Exit Sub
'           ElseIf OPX% = 1 Then 'CANCELA SALE
'             Exit Sub
'           End If  ' SI NO ENTRA EN NINGUNA DE LAS 2 SIGUE Y AGREGA.
        End If
     Next i&
     
     'SI PASA LA VALIDACION
     FechaInicioProceso% = HOY(HO$)
     VDEF$ = MKI$(FechaInicioProceso%) & Format(Now, "hh:mm")
     OBX$ = OBJPLA$("FECHAHORAFINPROC", VDEF$)
     
     CONDI$ = "IDSUBOR = '" & NROOF$ & "'"
     FechaOrdenFabricacion% = CVINT(VALOBADA("ORDEFABR", "FechGen", CONDI$))
     
     FechaInicioProceso% = CVI(Mid$(OBX$, 1, 2))
     If FechaInicioProceso% = 0 Or FechaInicioProceso% < FechaOrdenFabricacion% Then
         Call MENSERR(24, "Fecha de Inicio de Proceso Inválida o Menor a la de Generación")
         Exit Sub
     End If
     HoraInicioProceso = IsDate(Mid$(OBX$, 3, 5) & ":00")
     If HoraInicioProceso = False Then
         Call MENSERR(24, "Hora de Inicio de Proceso Inválida")
         Exit Sub
     End If
    
     Call ACTUREGISTRO("ORDEFABR", "Status_Proceso", CONDI$, 1, REGAF&, "NOMESS")
     Call ACTUREGISTRO("ORDEFABR", "inicio_informado", CONDI$, FETEXCOR1$(FechaInicioProceso%) & " " & Mid$(OBX$, 3, 5) & ":00", REGAF&, "NOMESS")
     TXTX$ = "Se Informa Inicio de Proceso"
     Call AGREGA_ANOTAORFAB_NEW(NROOF$, TXTX$)
     '
     Call CargarOFabEnProceso(NROOF$)
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   For i& = 1 To MSF2.Rows - 1
      NORFA$ = MSF2.TextMatrix(i&, 1)
      CANTINFORMADAFINPROCESO = XVALO(MSF2.TextMatrix(i&, 5))
      If CANTINFORMADAFINPROCESO > 0 Then
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        'LEO LA CANTIDAD QUE TIENE EN LA BASE DE DATOS YA GUARDADA
        CANTIINFFINBDATOS = XVALO(VALOBADA("ORDEFABR", "Canti_fin_proceso", CONDI$, "NOMESS"))
        'ACTUALIZO LA CANTIDAD A FIN DE PROCESO
        FechaFinDeProceso = MSF2.TextMatrix(i&, 7)
        
        Call ACTUREGISTRO("ORDEFABR", "Canti_fin_proceso", CONDI$, CANTINFORMADAFINPROCESO, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "Status_Proceso", CONDI$, 2, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "FinDeProceso", CONDI$, FechaFinDeProceso, REGAF&, "NOMESS")
        'SI SE CAMBIO EL VALOR GRABA
        If CANTIINFFINBDATOS <> CANTINFORMADAFINPROCESO Then
          TXTX$ = "Se Informa Fin de Proceso por " & FORMATNUM$(CANTINFORMADAFINPROCESO, "F10.1")
          Call AGREGA_ANOTAORFAB_NEW(NORFA$, TXTX$)
          'VERIFICO SI  LA O.F. NO ESTA EN LA GRILLA DE ARRIBA LE BORRO LA CANTIDAD INFORMADA COMO INICIO DE PROCESO,
          Call PONE_EN_CERO_CANTI_INI_PROCESO(NORFA$)
        End If
        '
        'VERIFICO SI
      End If
   Next i&
99 Command2.Enabled = True
   Screen.MousePointer = 1

End Sub
Sub PONE_EN_CERO_CANTI_INI_PROCESO(NROOF$)
     'RECORRO LA GRILLA SUPERIOR SI ESTA SE VA.
     For i& = 1 To MSF1.Rows - 1
        OFGRILLA$ = TRIM$(MSF1.TextMatrix(i&, 1))
        If OFGRILLA$ = NROOF$ Then
           Exit Sub
        End If
     Next i&
     '
     'SI NO ESTA VA A LLEGAR AQUI Y LA BORRA.
     CONDI$ = "IDSUBOR = '" & NROOF$ & "'"
     Call ACTUREGISTRO("ORDEFABR", "Canti_Iniciada", CONDI$, 0, REGAF&, "NOMESS")
     '
End Sub

Private Sub Form_Load()
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    CAMPOS$ = "O.Fabricación/A12;Código/A16;Descripción/A26;Cantidad/F10.1;Fin Proceso/F11.1;Aprobada/F11.1;Generada/A14;Inf. a Cerrar/F13.1;Pend. Informar/F13.1;En Proceso/F13.1"
    CAMPOS2$ = "O.Fabricación/A12;Código/A20;Descripción/A32;Cantidad/F10.1;Fin Proceso/F10.1;Aprobada/F11.1;Cerrada/A14;Inf. a Cerrar/F13.1;Pend. Informar/F13.1"
    Call CARGA_ENCABEZADO(MSF1, CAMPOS$, INFINI09)
    Call CARGA_ENCABEZADO(MSF2, CAMPOS2$, INFINI09)
    'PROPIEDADES DE SELECCION DE FILA
    MSF1.SelectionMode = flexSelectionByRow
    MSF1.FocusRect = flexFocusNone
    '
    MSF2.SelectionMode = flexSelectionByRow
    MSF2.FocusRect = flexFocusNone
    '
    Screen.MousePointer = 11
    Call cargarCantfinProceso
    Call CargarOFabEnProceso
    Call CARGAROFABPROCESOSINFORMADOS
    Call PoneEnCeroNull("ORDEFABR", "Status_proceso", REGAF&, "NOMESS")
    If INFINI09.NORFAEXTERNA <> "" Then
       Call Command1_Click
    End If
    '
    If InStr(EPAC$, "PREMOTEC") > 0 Then
      pp.Visible = True
    End If

    Screen.MousePointer = 1
End Sub

Private Sub MSF1_DblClick()
   If (MSF1.MouseRow = 0) Then Exit Sub
   REG& = MSF1.MouseRow
   If REG& < 1 Or MSF1.Rows <= 1 Then Exit Sub
   NROOF$ = TRIM$(MSF1.TextMatrix(REG&, 1))
   If NROOF$ = "" Then
      Call COMUNI("Orden de Fabricación Inexistente o No Válida")
      Exit Sub
   End If
   '
   'posiciono sobre la lista Inferior
   For i& = 1 To MSF2.Rows - 1
        OFGRILLA$ = TRIM$(MSF2.TextMatrix(i&, 1))
        If OFGRILLA$ = NROOF$ Then
             MSF2.TopRow = i&
             MSF2.Row = i&
             MSF2.ColSel = 1
             MSF2.SetFocus
             If XVALO(MSF1.TextMatrix(i&, 10)) > 0 Then
             'SI TIENE CANTIDAD INFORMADA AVISA QUE LA VA A SUMAR A LA CANTIDAD INFORMADA COMO FIN DE PROCESO.
             TOTALSUMA$ = TRIM$(FORMATNUM$(XVALO(MSF1.TextMatrix(i&, 10)) + XVALO(MSF1.TextMatrix(i&, 5)), "F9.1"))
                TXX1$ = "Atención se Sumaran a la Cantidad Informada Como Fin de Proceso (" & TRIM$(MSF1.TextMatrix(REG&, 5)) & ") "
                TXX1$ = TXX1$ + "\La Cantidad En Proceso (" & TRIM$(MSF1.TextMatrix(i&, 10)) & ")"
                TXX1$ = TXX1$ + "\Total Fin Proc.: " & TOTALSUMA$
                Call COMUNI(TXX1$)
                MSF2.TextMatrix(i&, 5) = TOTALSUMA$
             End If
             GoTo 80
        End If
   Next i&
   '
   'SI NO EXISTE EN GRILLA INFERIOR LO AGREGO
   
   MSF2.Rows = MSF2.Rows + 1
   MSF2.Cols = MSF1.Cols
   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF1.TextMatrix(REG&, 1)
   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF1.TextMatrix(REG&, 2)
   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF1.TextMatrix(REG&, 3)
   MSF2.TextMatrix(MSF2.Rows - 1, 4) = MSF1.TextMatrix(REG&, 4)
   'MSF2.TextMatrix(MSF2.Rows - 1, 5) = XVALO(MSF1.TextMatrix(REG&, 5)) + XVALO(MSF1.TextMatrix(MSF1.MouseRow, 10))
   MSF2.TextMatrix(MSF2.Rows - 1, 6) = MSF1.TextMatrix(REG&, 6)
   'MSF2.TextMatrix(MSF2.Rows - 1, 7) = MSF1.TextMatrix(REG&, 7)
   MSF2.TextMatrix(MSF2.Rows - 1, 8) = MSF1.TextMatrix(REG&, 8)
   MSF2.TextMatrix(MSF2.Rows - 1, 9) = MSF1.TextMatrix(REG&, 9)
   
   MSF2.TopRow = MSF2.Rows - 1
   MSF2.Row = MSF2.Rows - 1
   MSF2.ColSel = 1
   MSF2.SetFocus
   '
80   Call borrafila(MSF1.Row, Me, Me.MSF1)
'   MSF2.Sort = 1
End Sub

Private Sub msf2_Click()

'          COLUM = MSF2.MouseCol
'          If COLUM <> 5 Then
'            TXTMOVIL.Visible = False
'            Exit Sub
'          End If
'          MSF2.COL = 5
'          With TXTMOVIL
'            .ZOrder 0
'            .Width = MSF2.CellWidth
'            .Left = MSF2.CellLeft + MSF2.Left
'            .Top = MSF2.CellTop + MSF2.Top
'            .TEXT = MSF2.TEXT
'            .Visible = True
'            .ToolTipText = MSF2.Row
'            .SetFocus
'          End With

End Sub

Private Sub MSF2_DblClick()
   If (MSF2.MouseRow = 0) Then Exit Sub
   '
   Dim FechaInicioProceso As Date
   '
   REG& = MSF2.MouseRow
   NORFA$ = MSF2.TextMatrix(REG&, 1)
   CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
   CANTIAPRO = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$, "NOMESS"))
   If CANTIAPRO < 1 Then
        REG& = MSF2.Row
        MSF1.Rows = MSF1.Rows + 1
        MSF1.Cols = MSF2.Cols
        MSF1.TextMatrix(MSF1.Rows - 1, 1) = MSF2.TextMatrix(REG&, 1)
        MSF1.TextMatrix(MSF1.Rows - 1, 2) = MSF2.TextMatrix(REG&, 2)
        MSF1.TextMatrix(MSF1.Rows - 1, 3) = MSF2.TextMatrix(REG&, 3)
        MSF1.TextMatrix(MSF1.Rows - 1, 4) = MSF2.TextMatrix(REG&, 4)
        MSF1.TextMatrix(MSF1.Rows - 1, 5) = " "
        MSF1.TextMatrix(MSF1.Rows - 1, 6) = MSF2.TextMatrix(REG&, 6)
        FechaInicioProceso = VALOBADA("ORDEFABR", "inicio_informado", CONDI$, "NOMESS")
        MSF1.TextMatrix(MSF1.Rows - 1, 7) = Format(FechaInicioProceso, "dd/mm/yy hh:mm:00")
        MSF1.TextMatrix(MSF1.Rows - 1, 8) = MSF2.TextMatrix(REG&, 8)
        MSF1.TextMatrix(MSF1.Rows - 1, 9) = MSF2.TextMatrix(REG&, 9)
        'MSF1.TextMatrix(MSF1.Rows - 1, 10) = MSF2.TextMatrix(REG&, 5)
        
        Call borrafila(Me.MSF2.Row, Me, Me.MSF2)
        'ACTUALIZO LA CANTIDAD A FIN DE PROCESO
        Call ACTUREGISTRO("ORDEFABR", "Canti_fin_proceso", CONDI$, 0, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "Status_Proceso", CONDI$, 1, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "FinDeProceso", CONDI$, "'" & FETEXCOR1$(FECHANUM("01/01/80")) & "'", REGAF&, "NOMESS")
   Else
        Call COMUNI("Imposible Subir Esta Orden de Fabricación a Grilla de Solo Informadas\La Misma Posee Cierres Informados")
        Exit Sub
   End If
End Sub

Private Sub sininformar_Click()
    Call FINAL("?SININFOR")
'   FILTX$ = ""
'   Call STDFORM("SINF-ARC", FILTX$)
   Command3.Enabled = True

End Sub

Private Sub TXTMOVIL_Change()
       REG& = XVALO(TXTMOVIL.ToolTipText)
       INFORMEFIN = XVALO(TXTMOVIL)
       CANTAPROB = XVALO(MSF2.TextMatrix(REG&, 6))
       If CANTAPROB > INFORMEFIN Then
          Call COMUNI("Cantidad Informada Como Finalizada en Proceso es Menor a la Cantidad ya Cerrada")
          MSF2.TextMatrix(REG&, 5) = FORMATNUM$(CANTAPROB, "F10.1")
       End If

End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub


Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      REG& = XVALO(TXTMOVIL.ToolTipText)
      ORFA$ = MSF2.TextMatrix(XVALO(TXTMOVIL.ToolTipText), 1)
'      TXTMOVIL.ToolTipText = REG&
      MSF2.TEXT = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
      TXTMOVIL.Visible = False
   End If
End Sub

Private Sub TXTMOVIL_LostFocus()
     REG& = XVALO(TXTMOVIL.ToolTipText)
     If TXTMOVIL.Visible = True Then
       MSF2.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
       TXTMOVIL.Visible = False
'       ORFA$ = MSF2.TextMatrix(REG&, 1)
'       INFORMEFIN = XVALO(MSF2.TextMatrix(REG&, 5))
'       CANTAPROB = XVALO(MSF2.TextMatrix(REG&, 6))
'       If CANTAPROB > INFORMEFIN Then
'          Call COMUNI("Cantidad Informada Como Finalizada en Proceso es Menor a la Cantidad ya Cerrada")
'          MSF2.TextMatrix(REG&, 5) = FORMATNUM$(CANTAPROB, "F12.2")
'       End If

     End If
     
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    MSF2.COL = 5
    Select Case KeyCode
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF2.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F10.1")
            'MSF2.SetFocus
            If MSF2.Row > 1 Then MSF2.Row = MSF2.Row - 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF2.CellWidth
            .Left = MSF2.CellLeft + MSF2.Left
            .Top = MSF2.CellTop + MSF2.Top
            
             .TEXT = FORMATNUM$(MSF2.TEXT, "F10.1")
            .Visible = True
            .ToolTipText = MSF2.Row
            .SetFocus
            Call PINTATEXT(TXTMOVIL)
          End With
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            'MSF2.SetFocus
            MSF2.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F10.1")
            If MSF2.Row < MSF2.Rows Then MSF2.Row = MSF2.Row + 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF2.CellWidth
            .Left = MSF2.CellLeft + MSF2.Left
            .Top = MSF2.CellTop + MSF2.Top
             .TEXT = FORMATNUM$(MSF2.TEXT, "F10.1")
            .Visible = True
            .ToolTipText = MSF2.Row
            .SetFocus
            Call PINTATEXT(TXTMOVIL)
          End With
            On Error GoTo 0
    End Select
    '
End Sub

Private Sub geplapro_Click()

   'VALIDAR QUE HAYA MARCADO FILAS
   HAYTILDADOS% = 0
   For i& = 1 To MSF1.Rows - 1
      If TRIM$(MSF1.TextMatrix(i&, 0)) = "*" Then
         HAYTILDADOS% = 1
         'VALIDAR QUE TENGAN CANTIDAD DE FIN DE PROCESO SOLO LOS QUE ESTAN MARCADOS.
         If XVALO(MSF1.TextMatrix(i&, 10)) <= 0 Then
           Call COMUNI("Atención !!! Fila Seleccionada N°: " & i& & " No Tiene Cantidad En Proceso")
           Exit Sub
         End If
      End If
   Next i&
   If HAYTILDADOS% = 0 Then
     Call COMUNI("Atención !!!  No Se Seleccionó Ningúna Fila")
     Exit Sub
   End If
   '
   FORIPRE$ = TRIM$(Control("CONFABPE", "PLANPROD"))
   If FORIPRE$ = "" Then Exit Sub
   '
   If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Planilla de Producción:\  \'" + RMCC$ + "'.")
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11

   'RECORRER NUEVAMENTE Y GENERAR LA IMPRESION.
   FEX = HOY(HOS$)
   FECHDOC$ = FECHATEX$(FEX)
   NUMEDOC$ = ""
   TIPODOC$ = "Planilla de Producción"
   '
   CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHDOC$
      DIASECADO% = XVALO(Control("CONFABPE", "DIASSECA"))
   CODPARAM$(2) = "FECH-LAN": DOCPARAM$(2) = FECHATEX$(DIASPOST(FEX, DIASECADO%))
   CODPARAM$(3) = "TOTA-CAN"
   CODPARAM$(4) = "IMP-TOTA"
   '
   CAPARDOC% = 4
   '
   CODTABU1$(1) = "ORD-ITEM"
   CODTABU1$(2) = "COD-CONJ"
   CODTABU1$(3) = "DES-MAT"
   CODTABU1$(4) = "DES-OPER"
   CODTABU1$(5) = "CAN-CONJ"
   CODTABU1$(6) = "CANTIDAD"
   
   CACOLTAB1% = 6
   CAITAB1% = 0
   Call LIMPIATETAB
   '
   ACUCANTIORIG = 0
   ACUCANTITOTM2 = 0
   ORD% = 0
   For i& = 1 To MSF1.Rows - 1
      If TRIM$(MSF1.TextMatrix(i&, 0)) = "*" Then
         ORD% = ORD% + 1
         NORFA$ = TRIM$(MSF1.TextMatrix(i&, 1))
         NPED& = XVALO(VALOBADA("ORDEFABR", "NroPed", "IdSubOr = '" & NORFA$ & "'", "NOMESS"))
         NL% = 0
         If NPED& > 0 Then
            NC = XVALO(VALOBADA("PEDENCA", "NroClie", "NROPED = " & NPED&, "NOMESS"))
            If NC > 0 Then NL% = LIPRCLI%(NC)
         End If
         CODI$ = TRIM$(MSF1.TextMatrix(i&, 2))
         DESXRIP$ = TRIM$(MSF1.TextMatrix(i&, 3))
         CANTIORIG = XVALO(TRIM$(MSF1.TextMatrix(i&, 10)))
         CI1% = CODINT%(CODI$)
         UNME$ = Unimed$(CI1%)
         '
         CANTITOTM2 = CANTIORIG
         CANM2 = 1
         If UNME$ = "M2" Then
           CANM2 = CANTIXM2(CI1%, NL%)
           If CANM2 < 0 Then CANM2 = 6
           CANTITOTM2 = CANTIORIG * CANM2
         Else
           CANTIORIG = 0
         End If
         
         '
         CAITAB1% = CAITAB1% + 1
         
         TETABU1$(1, CAITAB1%) = SAFETEXT$(ORD%)
         TETABU1$(2, CAITAB1%) = CODI$
         TETABU1$(3, CAITAB1%) = DESXRIP$ + " - " + UNME$
         TETABU1$(4, CAITAB1%) = NORFA$
         TETABU1$(5, CAITAB1%) = FORMATNUM$(CANTITOTM2, "F8.0")
         TETABU1$(6, CAITAB1%) = FORMATNUM$(CANTIORIG, "F8.0")
         ACUCANTIORIG = ACUCANTIORIG + CANTIORIG
         ACUCANTITOTM2 = ACUCANTITOTM2 + CANTITOTM2
         'LINEA DE SEPARACION
         Call CARLINSEP

      End If
   Next i&
   '
   DOCPARAM$(3) = ACUCANTIORIG
   DOCPARAM$(4) = ACUCANTITOTM2

   Call PRINTDOC("@" + FORIPRE$)   ' PLANILLA DE PRODUCCION
   '
   Call LIMPIATETAB

   Call CIERRARCH("*.*")
   Screen.MousePointer = 1

End Sub
Private Sub viplage_Click()
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Fecha/D8;ID/F8;Referencia/A36", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Fecha/D8;ID/F8;Referencia/A36", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Lista de Planilla de Producción"
    '
    FRMZELECHO.MSF1.Rows = 1
    SQLX$ = "SELECT FECHAYHORA, NUMEDOC, REFERENCIA, TIPODOCU "
    SQLX$ = SQLX$ + " FROM PDOC_TBL WITH(NOLOCK) WHERE TIPODOCU = 'Planilla de Producción' ORDER BY NUMEDOC"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        FE% = CVINT(!FECHAYHORA)
        NUMEDOX = XVALO(!NUMEDOC)
        REFE$ = SAFETEXT$(!REFERENCIA)
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = FECHATEX$(FE%)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(NUMEDOX)
        FRMZELECHO.MSF2.TextMatrix(i&, 3) = REFE$
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    FRMZELECHO.Show 1
    NUDOC& = RESP_ZELE_VECT(2)
    Call MUESTRADOC(NUDOC&)
    
End Sub

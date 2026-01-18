VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DIAGPREM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proyección de Pedidos "
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14940
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   9360
      Width           =   7605
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0C0&
         Height          =   285
         Left            =   50
         ScaleHeight     =   225
         ScaleWidth      =   675
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   50
         ScaleHeight     =   225
         ScaleWidth      =   675
         TabIndex        =   10
         Top             =   425
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   50
         ScaleHeight     =   225
         ScaleWidth      =   675
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Cantidad Acumulada de Pedidos Pendientes en la Fecha Activa"
         Top             =   120
         Width           =   6630
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Stock del item / Saldo (El Saldo Considera Proyectado de entrega Anterior)"
         Top             =   435
         Width           =   6630
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Cantidad a Fabricar (Cantidad Pedida - Stock)"
         Top             =   750
         Width           =   6630
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7800
      OleObjectBlob   =   "DIAGPREM.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "DIAGPREM.frx":0234
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1215
      Left            =   10920
      TabIndex        =   2
      ToolTipText     =   "Doble Click Editar"
      Top             =   8880
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   9255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Doble Click Editar"
      Top             =   120
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   16325
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Pedidio Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2685
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuRefrescar 
         Caption         =   "Refrescar"
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
   End
End
Attribute VB_Name = "DIAGPREM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$










Private Sub Command5_Click()
    
End Sub


Private Sub Form_Load()
    Call CARGA_ENCA_DIAG(DIAGPREM.MSF1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MSF1.Height = Me.Height - 1600
    
    MSF1.Width = Me.Width - 235
    On Error GoTo 0
    Frame1.Visible = True
    Frame1.Top = Me.Height - (Frame1.Height + 400)
End Sub

Private Sub mnuBuscar_Click()
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To MSF1.Rows - 1
      For i& = 0 To MSF1.Cols - 1
       If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        MSF1.Row = U&: MSF1.COL = i&
        If MSF1.TopRow > MSF1.Row Then MSF1.TopRow = MSF1.Row
        On Error Resume Next
        VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 20
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If MSF1.TopRow < MSF1.Row - VISIROWS Then
              MSF1.TopRow = MSF1.Row - VISIROWS
            End If
         End If
        MSF1.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub
Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF1.Row

          H& = 0: If MSF1.COL < MSF1.Cols Then H& = MSF1.COL + 1
10        For U& = J& To MSF1.Rows - 1
             HI& = H&: If U& > J& Then HI& = 0
             For i& = HI& To MSF1.Cols - 1
              If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF1.Row = U&: MSF1.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF1.TopRow < MSF1.Row - VISIROWS Then
                      MSF1.TopRow = MSF1.Row - VISIROWS
                    End If
                 End If
                MSF1.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

Private Sub mnuRefrescar_Click()
   Call MostrarDiagrama(DIAGPREM.MSF1, DIAGPREM.MSF2)
End Sub

Private Sub MSF1_Click()
 REG& = MSF1.MouseRow
 COL& = MSF1.MouseCol
 
' If REG& > 0 Then
'    Me.TXTCodigo = MSF1.TextMatrix(REG&, 1)
'    Me.TXTCantidad = MSF1.TextMatrix(REG&, 3)
'
' End If
End Sub

Sub TXTCantidad_Click()

End Sub


Private Sub MSF1_DblClick()
   REG& = MSF1.MouseRow
   COL& = MSF1.MouseCol
   CI1% = CODINT%(MSF1.TextMatrix(REG&, COL&))
   If CI1% < 1 Then Exit Sub ' SI NO HAY CANITIDAD EXIT
   SITUSTO.TXTCODIGO = MSF1.TextMatrix(REG&, COL&)
   SITUSTO.Show 1
   
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
   REG& = MSF1.MouseRow
   COL& = MSF1.MouseCol
   If XVALO(MSF1.TextMatrix(REG&, COL&)) < 1 Then Exit Sub ' SI NO HAY CANITIDAD EXIT
   If REG& < 1 Then Exit Sub
   CI1% = CODINT%(MSF1.TextMatrix(REG&, 0))
   If CI1% > 0 Then
      Call MOSTRAR_DETA_COD_PEDI(CI1%, REG&, COL&)
      Exit Sub
   End If
   CI1% = CODINT%(MSF1.TextMatrix(REG& - 1, 0))
   If CI1% > 0 Then
      Call MOSTRAR_KARDEX(CI1%)
      Exit Sub
   End If
   CI1% = CODINT%(MSF1.TextMatrix(REG& - 2, 0))
   If CI1% > 0 Then
      Call MOSTRAR_OFABR(CI1%)
      Exit Sub
   End If
  End If
End Sub
Sub MOSTRAR_OFABR(CI1%)
   Call ABREORFAB
   '
   Call PoneEnCeroNull("ORDEFABR", "Status_Proceso", REGAF&, "NOMESS")
   Call PoneEnCeroNull("ORDEFABR", "canti_fin_proceso", REGAF&, "NOMESS")
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "O.Fabricación/A12;Cant.Orig/F11.1;Cant.Cerr./F11.1;En Proceso/F11.1", FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "O.Fabricación/A12;Cant.Orig/F11.1;Cant.Cerr./F11.1;En Proceso/F11.1", FRMZELECHO, , , ANTOT)
   If ANTOT < 5500 Then ANTOT = 5500
   FRMZELECHO.Width = ANTOT + 1100
   FRMZELECHO.Caption = "Lista de Ordenes de Fabricación Generadas Para Código: " & CODEXT$(CI1%)

   SQLX$ = " SELECT IDSUBOR, CanProy,  CanApro,  NUMEORFA, canti_fin_proceso,  Status_Proceso FROM ORDEFABR WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI1%
   'SQLX$ = SQLX$ + "  AND Status_Proceso < 2 "
   SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
   SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
   SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
   '
   Set INFOFABTMP = READSET(SQLX$)
   FRMZELECHO.MSF1.Rows = 1
   With INFOFABTMP
     Do While Not .EOF
       NUMORFA$ = SAFETEXT$(!idsubor)
       CACOM = XVALO(!CANPROY)
       CAPRO = XVALO(!CanApro)
       CAFINPROCESO = XVALO(!canti_fin_proceso)
       i& = i& + 1

       FRMZELECHO.MSF2.Rows = i& + 1
       FRMZELECHO.MSF2.TextMatrix(i&, 1) = NUMORFA$
       FRMZELECHO.MSF2.TextMatrix(i&, 2) = FORMATNUM$(CACOM, "F9.1")
       FRMZELECHO.MSF2.TextMatrix(i&, 3) = FORMATNUM$(CAPRO, "F9.1")
       FRMZELECHO.MSF2.TextMatrix(i&, 4) = FORMATNUM$(CAFINPROCESO, "F9.1")
       .MoveNext
     Loop
   End With
   FRMZELECHO.Show 1
   '
   NORFA$ = TRIM$(RESP_ZELE_VECT(1))
   If NORFA$ <> "" Then
     Call CargarOFabEnProcesoSimple(NORFA$)
   End If
   '
End Sub
Sub MOSTRAR_DETA_COD_PEDI(CI1%, REG&, COL&)
   TEXTOBUS$ = SAFETEXT$(REG&) & ";" & SAFETEXT$(COL&)
   For i& = 1 To MSF2.Rows - 1
      If MSF2.TextMatrix(i&, 1) = TEXTOBUS$ Then
         PEDIDOS$ = MSF2.TextMatrix(i&, 2)
         Call TEXTOLOTES$(PEDIDOS$, ";")
         For J& = 1 To UBound(CADENATEX$)
           NROPED& = XVALO(CADENATEX$(J&))
           If J& = 1 Then
              CONDI$ = " (NROPED = " & NROPED&
           Else
              CONDI$ = CONDI$ + " OR NROPED = " & NROPED&
           End If
         Next J&
         Exit For
      End If
   Next i&
   CONDI$ = CONDI$ + ")"
   i& = 0
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro.Ped/F9;Cta./F6;R.Social/A24;Cant.Ped./F9.1;Cant.Ent./F9.1;Sol.Ent/D8", FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Nro.Ped/F9;Cta./F6;R.Social/A24;Cant.Ped./F9.1;Cant.Ent./F9.1;Sol.Ent/D8", FRMZELECHO, , , ANTOT)
   If ANTOT < 5500 Then ANTOT = 5500
   FRMZELECHO.Width = ANTOT + 1100
   FRMZELECHO.Caption = "Lista de Pedidos Pendientes Para Código: " & CODEXT$(CI1%)
   '
   FRMZELECHO.MSF1.Rows = 1
   SQLX$ = "SELECT NroPed,NroClie, RaSoClie,CanPed, CantEnt, FeSolEnt FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " AND CODIINT = " & CI1%
   SQLX$ = SQLX$ + " ORDER BY FeSolEnt, NroPed"
   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        NP% = XVALO(!NROPED)
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = NP%
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = XVALO(!NROCLIE)
        FRMZELECHO.MSF2.TextMatrix(i&, 3) = SAFETEXT$(!RaSoClie)
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = FORMATNUM$(XVALO(!CanPed), "F9.1")
        FRMZELECHO.MSF2.TextMatrix(i&, 5) = FORMATNUM$(XVALO(!CANTENT), "F9.1")
        FRMZELECHO.MSF2.TextMatrix(i&, 6) = FECHATEX$(CVINT(!FeSolEnt))
        .MoveNext
      Loop
   End With
   rs.Close
   Set rs = Nothing
   FRMZELECHO.Show 1
   
    NROPED& = XVALO(RESP_ZELE_VECT(1))
    NCLIE = XVALO(RESP_ZELE_VECT(2))
    CANTISALDO = XVALO(RESP_ZELE_VECT(4)) - XVALO(RESP_ZELE_VECT(5))
    
    If NROPED& > o Then
       OPX% = COMALTER%("Opciones de trabajo\\Cancelar\Modificar Datos del Pedido\Generar Orden de Fabricación")
       If OPX% < 2 Then Exit Sub
       If OPX% = 2 Then Call OPPED07.CAMBIAPED(NROPED&)
       If OPX% = 3 Then Call GENERA_OF(CI1%, NROPED&, CANTISALDO, NCLIE) 'GENERA UNA OF NUEVA
    End If

   
End Sub

Sub GENERA_OF(CI1%, NROPED&, CANTI, Optional NCLIE)
   'GENERA UNA OF NUEVA
'   UNME$ = UNIMED$(CI1%)
'   '
'   If UNME$ = "M2" Then
'      CAN = CANTI * 6
'   End If

   CONDI$ = " NROPED = " & NROPED&
   CONDI$ = CONDI$ + " And STATUORF < 3 "
   NORFA$ = TRIM$(VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESSS"))
   If NORFA$ = "" Then
     OPX% = COMALTER%("El Pedido Seleccionado No Generó Orden de Fabricación\\Cancelar\Generar una Nueva")
     If OPX% < 2 Then Exit Sub
     If OPX% = 2 Then
       CAN = CANTI
       CODE$ = CODEXT$(CI1%)
       NCLIE1 = NCLIE
       NUMPEDI% = NROPED&
       ORCOCLI$ = ""
   
       ACONEC1$ = Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE1)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$
       ACONEC1$ = REPLACAR$(ACONEC1$, "/", "ß") 'PARA EVITAR CODIGOS CON BARRA POR QUE EN  EL PASAJE LLEGA TRUNCADO POR ESTE CARACTER (JYM) (ALT-225)
       ACONEC$ = TRIM$("ORFABR07/ORFNUE" + ACONEC1$)
       'ORFABR07/ORFNUE|000004700|2|999|3697|
       Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
      End If
   Else
     Call CargarOFabEnProcesoSimple(NORFA$)
   End If
  
End Sub
Sub CargarOFabEnProcesoSimple(NORFA$)
   INFINI09.NORFAEXTERNA = NORFA$
   
   INFINI09.Show 1
   
End Sub

VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FMAIL09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Mail"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Height          =   330
      Left            =   5760
      Picture         =   "FMAIL09.frx":0000
      TabIndex        =   23
      ToolTipText     =   "Lista de Selección de Mail de Contactos "
      Top             =   2160
      Width           =   175
   End
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
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1480
      Width           =   4950
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   195
      Width           =   780
   End
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
      Height          =   330
      Left            =   3180
      Picture         =   "FMAIL09.frx":030A
      TabIndex        =   1
      Top             =   195
      Width           =   175
   End
   Begin VB.CommandButton Command9 
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
      Height          =   330
      Left            =   5640
      Picture         =   "FMAIL09.frx":0614
      TabIndex        =   3
      Top             =   195
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
      Height          =   330
      Left            =   4860
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   195
      Width           =   780
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2640
      Width           =   4950
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   6
      Top             =   2190
      Width           =   4950
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
      Height          =   330
      Left            =   840
      TabIndex        =   4
      Top             =   1035
      Width           =   4950
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00800000&
      Height          =   8520
      Left            =   6090
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   -105
      Width           =   1065
      Begin VB.PictureBox Picture16 
         Height          =   5200
         Left            =   600
         ScaleHeight     =   5145
         ScaleWidth      =   1005
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   5250
            Left            =   480
            Picture         =   "FMAIL09.frx":091E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.CommandButton command1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FMAIL09.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FMAIL09.frx":2E00
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Aprueba y Envía el Mail "
         Top             =   3360
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   4320
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FMAIL09.frx":310A
         Top             =   4680
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FMAIL09.frx":502C
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FMAIL09.frx":5260
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Proveedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   19
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   18
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3550
      TabIndex        =   17
      Top             =   240
      Width           =   1305
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6100
      Y1              =   980
      Y2              =   980
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   130
      TabIndex        =   14
      Top             =   2650
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   13
      Top             =   1500
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dir-El.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   130
      TabIndex        =   12
      Top             =   2200
      Width           =   645
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   1040
      Width           =   645
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu enviar 
         Caption         =   "Enviar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FMAIL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command14_Click()
   '
    Command14.Enabled = False
    '
    ' *** DETECTA IMPRESORA VIRTUAL INSTALADA
    IMPRESO$ = ""
    For Each PRX1 In Printers   ' BUSCA BIOPDF O BULLZIP
       If InStr(UCase$(PRX1.DeviceName), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then
        IMPRESO$ = PRX1.DeviceName
        GoTo 10
       End If
    Next
    For Each PRX1 In Printers   ' BUSCA PDFCREATOR
       If InStr(UCase$(PRX1.DeviceName), "PDFCREATOR") > 0 Then
        IMPRESO$ = PRX1.DeviceName
        GoTo 10
       End If
    Next
    Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDF.")
    Exit Sub

    ' *** VALIDACIONES
    '
10  EMAI$ = TRIM$(Text2)
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       Call MENSERR(24, "Mail No Válido")
       Text2.SetFocus
       GoTo 99
    End If
    '
    If Text1 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Falta Destino.")
      Text1.SetFocus
      GoTo 99
    End If
    '
    If Text6 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Asunto.")
      Text6.SetFocus
      GoTo 99
    End If
   '
    If Text5 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Texto.")
      Text5.SetFocus
      GoTo 99
    End If
    '
    NOMARCH$ = "DOC_ADJ"
                    
    ' *** MODO BULLZIP Y BIOPDF - GENERACION DE PDF
    '
    IMPRESONAME$ = ""
    If InStr(UCase$(IMPRESO$), "BULLZIP") > 0 Then IMPRESONAME$ = "BULLZIP"
    If InStr(UCase$(IMPRESO$), "BIOPDF") > 0 Then IMPRESONAME$ = "BIOPDF"
    If IMPRESONAME$ = "" Then GoTo 45
    SETTINGS_PROGID = IMPRESONAME$ & ".PDFSettings"
    UTIL_PROGID = IMPRESONAME$ & ".PDFUtil"
    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.SetValue "Output", LUCOM$ + NOMARCH$ & ".pdf"
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "no"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    '
    ' *** ELIMINAMOS EL ARCHIVO, EN CASO DE EXISTIR
    If EXISTE%(LUCOM$ & "" + NOMARCH$ + ".PDF") > 0 Then
       On Error GoTo FalloAlEliminarDocumentoAnterior
       Kill LUCOM$ & "" + NOMARCH$ + ".PDF" ' ELIMINAMOS EL ARCHIVO
       On Error GoTo 0
    End If
    
    ' *** IMPRIME DOCUMENTO
    EDITFORM.COMBO1.TEXT = IMPRESO$
    EDITFORM.Label1 = NOMARCH$
    Call EDITFORM.Command4_Click
    '
    ' *** VERIFICA SI EL ARCHIVO YA FUE CREADO
    Dim runonce As String
41  runonce = settings.GetSettingsFilePath(True)
    VUELTA% = 0
    While Dir(runonce, vbNormal) = ""
        Call T_Espera(0.1)
        VUELTA% = VUELTA% + 1
        If VUELTA% > 15 Then
            TX1$ = "Generando Archivo PDF "
            OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
            If OPX% = 2 Then
                Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
                GoTo 90
            Else
                GoTo 41
            End If
        End If
    Wend
    GoTo 60
    '
    ' *** MODO PDFCREATOR - GENERACION DE PDF
    '
45  If InStr(UCase$(IMPRESO$), "PDFCREATOR") < 1 Then GoTo 90
    Dim PDFCreator1 As Object
    Set PDFCreator1 = CreateObject("PDFCREATOR.clsPDFCreator")
    Dim noStart As Boolean
    Dim opt As Object
    
    With PDFCreator1
      .cVisible = True
      If .cStart("/NoProcessingAtStartup") = False Then
        If .cStart("/NoProcessingAtStartup", True) = False Then
         Exit Sub
        End If
        .cVisible = True
      End If
      Set opt = .cOptions
      .cClearCache
      noStart = False
    End With
    
    With opt
        opt.AutosaveDirectory = LUCOM$
        opt.AutosaveFilename = NOMARCH$ & ".pdf"
        opt.UseAutosave = 1
        opt.UseAutosaveDirectory = 1
        opt.AutosaveFormat = 0 ' PDF
    End With
    Set PDFCreator1.cOptions = opt
    Set Printer = Printers(PrinterIndex("PDFCreator"))
    '
    ' *** ELIMINAMOS EL ARCHIVO, EN CASO DE EXISTIR
    If EXISTE%(LUCOM$ & "" + NOMARCH$ + ".PDF") > 0 Then
       On Error Resume Next
       Kill LUCOM$ & "" + NOMARCH$ + ".PDF"
       On Error GoTo 0
    End If
    
    ' *** IMPRIME DOCUMENTO
    EDITFORM.COMBO1.TEXT = IMPRESO$
    EDITFORM.Label1 = NOMARCH$
    Call EDITFORM.Command4_Click
    
    PDFCreator1.cPrinterStop = False
    Sleep 1000
    If noStart = False Then
        PDFCreator1.cClose
        While PDFCreator1.cProgramIsRunning
            DoEvents
            Sleep 100
        Wend
    End If
    Set PDFCreator1 = Nothing
    Set opt = Nothing
    
    ' *** VERIFICA SI EL ARCHIVO YA FUE CREADO
46  While Dir(LUCOM$ & "" + NOMARCH$ + ".PDF", vbNormal) = ""
        Call T_Espera(0.1)
        VUELTA% = VUELTA% + 1
        If VUELTA% > 15 Then
            TX1$ = "Generando Archivo PDF "
            OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
            If OPX% = 2 Then
                Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
                GoTo 90
            Else
                VUELTA% = 0
                GoTo 46
            End If
        End If
    Wend
    '
    ' *** CARGA DATOS DEL DESTINATARIO Y ARCHIVO ADJUNTO
60  MAI_DESTI$ = TRIM$(Text1)
    MAI_DIREL$ = TRIM$(Text2)
    MAI_ASUNT$ = TRIM$(Text6)
    MAI_TEXTO$ = Text5
    DoEvents
    Call T_Espera(2)
    If EXISTE%(LUCOM$ + "" + NOMARCH$ + ".PDF") = 1 Then
        MAI_ADJUN$(0) = LUCOM$ + "" + NOMARCH$ + ".PDF"
    End If

    Call T_Espera(5)
    '
    Call GENERAMAIL
    Call BLANFORMU
    '
90  Call CIERRARCH("*.*")
    Call Command1_Click
    '
99  Command14.Enabled = True
    
    Exit Sub
    
FalloAlEliminarDocumentoAnterior:
    Call COMUNI("Problema al Eliminar Documento Anterior\ \" & Err.Description)
    On Error GoTo 0
    GoTo 99
    
End Sub

Private Sub Command22_Click()
   Hide '18/04/07 (OT-07-0158)
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   Call SELECHO("PROVED1")
   NPR% = XVALO(TRIM$(VALACT1$("PROVED1"))) ' DESPLIEGA LISTA DE CLIENTES
   '
   If NPR% < 1 Then GoTo 99  'SI  NO SELECCIONA
   '
   Text4 = NPR%
   '
99 Command2.Enabled = True
   

End Sub

Private Sub Command3_Click()
     Command3.Enabled = False
     If XVALO(Text3) + XVALO(Text4) <= 0 Then GoTo 99
     If XVALO(Text3) > 0 Then
        NCX1 = XVALO(Text3)
     Else
        NCX1 = XVALO(Text4)
        NCX1 = -1 * NCX1
     End If
     
     Call CONTACTO_MAIL(CODICLI$(NCX1), NCX1)   'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
     If SENDBYMAIL$ <> "" Then
        Text2.TEXT = SENDBYMAIL$
     Else
        Text2.TEXT = ""
     End If
99   Command3.Enabled = True
     
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   '
   Call SELECHO("DEUDOR1") ' DESPLIEGA LISTA DE CLIENTES
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   '
   If NCLIE% < 1 Then GoTo 99  'SI  NO SELECCIONA
   '
   Text3 = NCLIE%
   '
99 Command9.Enabled = True
   

End Sub

Private Sub Label5_Click()

End Sub

Private Sub enviar_Click()
    Call Command14_Click
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 4800
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

   Call APLICACIONSKINS(Me, Picture5)
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text3_Change()
   NC% = XVALO(Text3)
   If NC% > 0 Then
     Label8 = rasocli$(NC%)
     Text2 = TRIM$(EMAILCLI$(NC%))
     CONT$ = CONTACTCLI$(NC%)
     If TRIM$(CONT$) = "" Then
       Text1 = Label8
     Else
       Text1 = CONT$
     End If
     Text6 = EDITFORM.Label98
     '
     Text5 = "Enviamos Documento Adjunto: " & EDITFORM.Label98 & " En Formato PDF"
     '
     'AGREGADO PARA AHP
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "AHP") > 0 And UCase$(Left$(EDITFORM.Label98, 11)) = "PRESUPUESTO" Then
       DEXCRIDOX$ = "Enviamos Documento Adjunto: " & EDITFORM.Label98
       POS1% = InStr(DEXCRIDOX$, "-")
       DEXCRIDOX$ = Left$(DEXCRIDOX$, POS1%)
       TXX1$ = DEXCRIDOX$ & " " & SAFETEXT$(FRMPRESUAHP.TXTREFERENCIA)
       TXX1$ = TXX1$ + "" + CRLF$
       Text5 = TXX1$
     End If

     Text4 = "" ' limpio el text de proveedor
    End If
   

End Sub

Private Sub Text4_Change()
   NPRO% = XVALO(Text4)
   If XVALO(Text4) > 0 Then
     Label8 = rasocli$((-1) * NPRO%)
     Text2 = TRIM$(EMAILCLI$((-1) * NPRO%))
     CONT$ = CONTACTCLI$((-1) * NPRO%)
     If TRIM$(CONT$) = "" Then
        Text1 = Label8
     Else
        Text1 = CONT$
     End If
     Text6 = EDITFORM.Label98

     Text5 = "Enviamos Documento Adjunto: " & EDITFORM.Label98 & " En Formato PDF"
     'SE VALIDA QUE SI ES NOTA DE PEDIDO (ORDEN DE COMPRA) SE LE AGREGA EL TEXTO DE MAIL ORDEN DE COMPRA QUE  SE CONFIGURA EN EL SETUP
     If UCase$(Left$(EDITFORM.Label98, 14)) = "NOTA DE PEDIDO" Then Text5 = Text5 & vbCrLf & vbCrLf & TEXTO_MAIL_COMPRA$
     '
     
     Text3 = "" ' limpio el text de cliente
   End If
   
End Sub

Sub BLANFORMU()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Label8 = ""
End Sub



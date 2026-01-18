VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STDMAIL 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENVIO DE MAIL"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   3360
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text6 
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
      Height          =   570
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1480
      Width           =   4950
   End
   Begin VB.TextBox Text4 
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
      Picture         =   "STDMAIL.frx":0000
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
      Picture         =   "STDMAIL.frx":030A
      TabIndex        =   3
      Top             =   195
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Left            =   4860
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   195
      Width           =   780
   End
   Begin VB.TextBox Text5 
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
      Height          =   2385
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2640
      Width           =   4950
   End
   Begin VB.TextBox Text2 
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
      Left            =   840
      TabIndex        =   6
      Top             =   2190
      Width           =   4950
   End
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   4
      Top             =   1035
      Width           =   4950
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H009BBBC6&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   6090
      ScaleHeight     =   8490
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "STDMAIL.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   640
         Left            =   120
         Picture         =   "STDMAIL.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Aprueba y Envía el Mail "
         Top             =   4320
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "STDMAIL.frx":0A68
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "STDMAIL.frx":0C9C
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   17
      Top             =   300
      Width           =   1620
   End
   Begin VB.Label Label8 
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
      Height          =   330
      Left            =   840
      TabIndex        =   16
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   15
      Top             =   300
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
         Weight          =   400
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
      Caption         =   "Asunto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   13
      Top             =   1500
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   12
      Top             =   2320
      Width           =   645
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Para"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   1140
      Width           =   645
   End
End
Attribute VB_Name = "STDMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public NombreDeArchivo As String

Private Sub Command1_Click()
   Unload Me
End Sub

Sub Command14_Click()
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
10  EMAI$ = TRIM$(Text2)
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       Call MENSERR(24, "Mail No Válido")
       GoTo 90
    End If
    '
    If Text6 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Asunto.")
      Text6.SetFocus
      GoTo 90
    End If
    '
    If Text5 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Texto.")
      Text5.SetFocus
      GoTo 90
    End If
    '
    ' NOMBRE DE ARCHIVO POR DEFAULT O EL PASADO SOBRE LA VARIABLE PUBLICA NombreDeArchivo
    NOMARCH$ = "DOC_ADJ"
    If TRIM$(NombreDeArchivo) <> "" Then
        NOMARCH$ = TRIM$(NombreDeArchivo)
    End If
    '
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
    settings.SetValue "Output", LUCOM$ + NOMARCH$ & ".PDF"
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
       On Error Resume Next
       Kill LUCOM$ & "" + NOMARCH$ + ".PDF"
       On Error GoTo 0
    End If
    '
    ' *** CONFIGURACION DEL STDSCREEN
    STDSCREEN.Combo1.TEXT = IMPRESO$
    STDSCREEN.Label6 = NOMARCH$
    If ModoBullzip% = 0 Then STDSCREEN.Timer1.Enabled = True
    '
    ' *** IMPRIME Y GENERA PDF
    Call STDSCREEN.Command4_Click
    Call SETPRINPORT("RESTORE")
    '
    Dim runonce As String
41  runonce = settings.GetSettingsFilePath(True)
    VUELTA% = 0
    While Dir(runonce, vbNormal) <> ""
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
    EDITFORM.Label1.Caption = ARANEX$
    Call T_Espera(5)
    GoTo 60

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
        opt.AutosaveFilename = NOMARCH$ & ".PDF"
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
    
    ' *** CONFIGURACION DEL STDSCREEN
    STDSCREEN.Combo1.TEXT = IMPRESO$
    STDSCREEN.Label6 = NOMARCH$
    If ModoBullzip% = 0 Then STDSCREEN.Timer1.Enabled = True
    '
    ' *** IMPRIME Y GENERA PDF
    Call STDSCREEN.Command4_Click
    Call SETPRINPORT("RESTORE")
    
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
60  MAI_DESTI$ = TRIM$(Text1)
    MAI_DIREL$ = TRIM$(Text2)
    MAI_ASUNT$ = TRIM$(Text6)
    MAI_TEXTO$ = Text5
    MAI_ADJUN$(0) = LUCOM$ + "" + NOMARCH$ + ".PDF"
    '
    Call GENERAMAIL
90  Call BLANFORMU
    '
    Call CIERRARCH("*.*")
    Call Command1_Click
   '
99 Command14.Enabled = True
   
End Sub

Private Sub Command22_Click()
   Hide '18/04/07 (OT-07-0158)
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   Call SELECHO("ACREDOR")
   NPR% = XVALO(TRIM$(VALACT1$("ACREDOR"))) ' DESPLIEGA LISTA DE CLIENTES
   '
   If NPR% < 1 Then GoTo 99  'SI  NO SELECCIONA
   '
   Text4 = NPR%
   '
99 Command2.Enabled = True
   

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

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture5)
End Sub

Private Sub TEXT3_Change()
   NC = XVALO(Text3)
   If NC > 0 Then
     Label8 = rasocli$(NC)
     Text2 = TRIM$(EMAILCLI$(NC))
     CONT$ = CONTACTCLI$(NC)
     If TRIM$(CONT$) = "" Then
       Text1 = Label8
     Else
       Text1 = CONT$
     End If
'     Text6 = EDITFORM.Label98
'     Text5 = "Enviamos Documento Adjunto: " & EDITFORM.Label98 & " En Formato PDF"

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
'     Text6 = EDITFORM.Label98
'     Text5 = "Enviamos Documento Adjunto: " & EDITFORM.Label98 & " En Formato PDF"

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



Private Sub Timer1_Timer()
    Static CANTIVUELTAS%
    
    CANTIVUELTAS% = CANTIVUELTAS% + 1
    If EXISTE%(LUCOM$ + "DOC_ADJ.PDF") > 0 Then ' SI ENCONTRO EL ARCHIVO, GENERA EL MAIL
        Timer1.Enabled = False: CANTIVUELTAS% = 0
        Call GENERAMAIL
        Call BLANFORMU
        '
        Call CIERRARCH("*.*")
        Call Command1_Click
    ElseIf CANTIVUELTAS% > 3 Then ' INTENTA 3 VECES QUE EL PDF SE HAYA GENERADO
        Timer1.Enabled = False: CANTIVUELTAS% = 0 ' DEJA AL USUARIO QUE ELIJA SI SE REINTENTA
        OPX% = COMALTER%("Imposible Generar Archivo PDF Adjunto.\¿Desea Volver a Interar?\\Si, Reintentar\No, Cancelar")
        If OPX% = 1 Then Timer1.Enabled = True
    End If
End Sub



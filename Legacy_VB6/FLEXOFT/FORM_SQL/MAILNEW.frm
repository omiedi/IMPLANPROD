VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MAILNEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Mail"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTPARA 
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
      Height          =   330
      Left            =   6600
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar Adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   12120
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Elimina el Adjunto Seleccionado"
      Top             =   2520
      Width           =   1905
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   6960
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   645
      Left            =   14040
      Picture         =   "MAILNEW.frx":0000
      TabIndex        =   21
      Top             =   1680
      Width           =   300
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   7320
      TabIndex        =   19
      Top             =   1680
      Width           =   6735
   End
   Begin VB.TextBox TXTENCABEZADO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7305
      TabIndex        =   16
      Top             =   75
      Width           =   6750
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
      Height          =   405
      Left            =   13800
      Picture         =   "MAILNEW.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Aprueba y Envía el Mail "
      Top             =   8520
      Width           =   650
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   -120
      ScaleHeight     =   480
      ScaleWidth      =   390
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox TXTASUNTO 
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
      Height          =   570
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1245
      Width           =   5175
   End
   Begin VB.TextBox Text4 
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
      Height          =   330
      Left            =   1440
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
      Left            =   2220
      Picture         =   "MAILNEW.frx":0614
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
      Picture         =   "MAILNEW.frx":091E
      TabIndex        =   3
      Top             =   195
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Height          =   330
      Left            =   4860
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   195
      Width           =   780
   End
   Begin VB.TextBox TXTSUBENCABEZADO 
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
      Height          =   945
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   6990
   End
   Begin VB.TextBox TXTDIRMAIL 
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
      Height          =   330
      Left            =   960
      TabIndex        =   5
      Top             =   1950
      Width           =   4935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MAILNEW.frx":0C28
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2160
      OleObjectBlob   =   "MAILNEW.frx":0E5C
      TabIndex        =   14
      Top             =   -120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5505
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   9710
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle del Enío"
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
      Left            =   240
      TabIndex        =   26
      Top             =   2640
      Width           =   3585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto:"
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
      Left            =   5880
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjuntos"
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
      Left            =   6195
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Encabezado"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro.: Proveedor:"
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
      Left            =   -20
      TabIndex        =   12
      Top             =   240
      Width           =   1400
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   960
      TabIndex        =   11
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
      Left            =   3435
      TabIndex        =   10
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Encabezado"
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
      Left            =   5955
      TabIndex        =   9
      Top             =   600
      Width           =   1350
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
      TabIndex        =   8
      Top             =   1260
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
      Left            =   135
      TabIndex        =   7
      Top             =   1965
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
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu mnuConfigmail 
         Caption         =   "Configurar Datos de Mail"
      End
   End
End
Attribute VB_Name = "MAILNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Screen.MousePointer = 1
   RUTAX$ = SELECCION_Archivo$("*", CMD1)
   If RUTAX$ <> "" Then List1.AddItem RUTAX$
   Screen.MousePointer = 1
End Sub

Private Sub Command14_Click()
   '
    Command14.Enabled = False
    '
    'VALIDACIONES
    If TRIM$(Me.TXTASUNTO) = "" Then
       Call COMUNI("Falta Completar Casillero de Asunto")
       Call PINTATEXT(TXTASUNTO)
       GoTo 99
    End If
    If TRIM$(Me.TXTDIRMAIL) = "" Then
       Call COMUNI("Falta Completar Casillero Dirección de Correo")
       Call PINTATEXT(TXTDIRMAIL)
       GoTo 99
    End If
    If TRIM$(Me.TXTENCABEZADO) = "" Then
       OPX% = COMALTER%("No se Ha Completado el Casillero de Encabezado\\Modificar\Continuar")
       If OPX% < 2 Then
         Call PINTATEXT(TXTENCABEZADO)
         GoTo 99
       End If
    End If
    If TRIM$(Me.TXTSUBENCABEZADO) = "" Then
       OPX% = COMALTER%("No se Ha Completado el Casillero de Encabezado Complementario\\Modificar\Continuar")
       If OPX% < 2 Then
         Call PINTATEXT(TXTSUBENCABEZADO)
         GoTo 99
       End If
    End If
    '
    PLANTILLAHTML$ = Control("MAILNEW", "PLANMAIL")
    
    pos1% = InStrRev(PLANTILLAHTML$, "\")
    If pos1% > 0 Then
      If TRIM$(UCase$(Mid$(PLANTILLAHTML$, pos1% + 1))) = "PLANTILLA_STANDARD_ORIG.TXT" Then
       CARPETOK$ = LUDAT$
       If EXISTE%(PLANTILLAHTML$) = 0 Then
          TXT$ = REPLACAR$(LUDAT$, "\", "/")
          Call MENSERR(24, "Falta Archivo 'Instalar Archivo PLANTILLA_STANDARD_ORIG.txt' \en Carpeta " & TXT$)
          GoTo 99
       End If
      End If
    '
       '*********************************************************************************************
       'CREO EL ARCHIVO CON LOS DATOS DE ENVIO DE MAIL
       '*********************************************************************************************
       For I1% = 1 To 32767
         ARCHDAMAIL$ = "C:\Flexoft\" & "DATOSMAIL" & TRIM$(Str$(I1%)) & ".TXT"
         If EXISTE%(ARCHDAMAIL$) < 1 Then
           NAME_ARCH_DATMAIL$ = "DATOSMAIL" & TRIM$(Str$(I1%)) & ".TXT"
           NARCHI0% = FILEOPEN%(ARCHDAMAIL$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
           Exit For
         End If
       Next I1%
       Usertex$ = SAFETEXT$(USNBR% Mod 100)
       CUENTAPROPIA1 = Control(Usertex$, "CTADMAIL")
       CONTRASEÑA1 = Control(Usertex$, "CONTRASE")
       PUERTO1 = Control(Usertex$, "PUERTOCO")
       SERVIDORSMTP1 = Control(Usertex$, "SERVERSM")
       
       Print #NARCHI0%, "|MAILRECEPTOR|" & Me.TXTDIRMAIL
       Print #NARCHI0%, "|ASUNTO|      " & Me.TXTASUNTO
       Print #NARCHI0%, "|CUENTAPROPIA|" & CUENTAPROPIA1
       Print #NARCHI0%, "|CONTRASEÑA|  " & CONTRASEÑA1
'       PUERTO$ = "587"
       Print #NARCHI0%, "|PUERTO|      " & PUERTO1
'       servid$ = "boxer.avnam.net"
       Print #NARCHI0%, "|SERVIDORSMTP|" & SERVIDORSMTP1
       
       'IMPRIMO LOS ADJUNTOS
       For i& = 0 To List1.ListCount - 1
          t1$ = List1.List(i&)
          Print #NARCHI0%, "|ADJUNTOS|    " & t1$
       Next i&


       Close #NARCHI0%
       '*********************************************************************************************
       'FIN EL ARCHIVO CON LOS DATOS DE ENVIO DE MAIL
       '*********************************************************************************************
       '
       '*********************************************************************************************
       'BUSCO QUE EL ARCHIVO NO EXISTA
       '*********************************************************************************************
       ARCH$ = "C:\Flexoft\" & "PLANTILLA_STANDARD.TXT"
       If EXISTE%(ARCH$) > 0 Then
        For I1% = 1 To 32767
          ARCH$ = "C:\Flexoft\" & "PLANTILLA_STANDARD" & TRIM$(Str$(I1%)) & ".TXT"
          If EXISTE%(ARCH$) < 1 Then
            NAME_ARCH_HTML$ = "PLANTILLA_STANDARD" & TRIM$(Str$(I1%)) & ".TXT"
            NARCHI2% = FILEOPEN%(ARCH$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
            Exit For
          End If
        Next I1%
       Else
         NAME_ARCH_HTML$ = "PLANTILLA_STANDARD.TXT"
         NARCHI2% = FILEOPEN%(ARCH$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
       End If
       '
       'COPIO DEL ORIGINAL AL ARCHIVO QUE SE ENVIA COMO PARAMETRO AL .NET
       NARCHI1% = FILEOPEN%(LUDAT$ + "PLANTILLA_STANDARD_ORIG.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    

            Do While Not EOF(NARCHI1%) ' RECORRE
               Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
               Select Case XX$
               Case "<H1><CENTER>ENCABEZADO1</CENTER></H1>"
                  If TRIM$(TXTENCABEZADO) <> "" Then
                     TX$ = "<H1><CENTER>" & TXTENCABEZADO & "</CENTER></H1>"
                  Else
                     TX$ = ""
                  End If
               '
               Case "<H2><CENTER>ENCABEZADO2</CENTER></H2>"
                  If TRIM$(Me.TXTSUBENCABEZADO) <> "" Then
                     TX$ = "<H2><CENTER>" & TXTSUBENCABEZADO & "</CENTER></H2>"
                  Else
                     TX$ = ""
                  End If
               Case "|CARGAR_LA_TABLA|"
                      'CARGA EL ENCABEZADO DE LA TABLA
                        TXT1$ = TXT1$ + "<table border=" & Chr(34) & "1" & Chr(34) & "style=" & Chr(34) & "margin: 0 auto;" & Chr(34) & ">"
                        TXT1$ = TXT1$ + "<tr>"
                        For i& = 1 To MSF.Cols - 1 'CARGO EL ENCABEZADO
                           If MSF.ColWidth(i&) > 10 Then ' SI NO ESTA VISIBLE NO LO ESCRIBO
                              TXT1$ = TXT1$ + "<td>" & MSF.TextMatrix(0, i&) & "</td>"
                           End If
                        Next i&
                        'CARGA EL DETALLE EN LA TABLA
                        TXT1$ = TXT1$ + "<tr>"
                        For i& = 1 To MSF.Rows - 1
                           For c& = 1 To MSF.Cols - 1
                            TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & MSF.TextMatrix(i&, c&) & "</td>"
                           Next c&
                           TXT1$ = TXT1$ + "</tr>"
                          
                        Next i&
                        TXT1$ = TXT1$ + "</table>"
                        TX$ = TXT1$
               Case "|CARGAR_LA_IMAGEN_FIRMA|"
                        'ESTO PODRIA ESTAR POR FUERA, LO DEJE AQUI DE TODOS MODOS NO ESTA MAL _
                        POR QUE ES LA BANDERA QUE UTILIZA EL .NET PARA INCRUSTAR LA IMAGEN EN EL HTML.
                        TX$ = XX$
                        ruta_archi_imag$ = Control(Usertex$, "FIRMAMAI") 'si esta configurado en el control para este usuario
                        If TRIM$(ruta_archi_imag$) <> "" Then
                          If EXISTE%(TRIM$(ruta_archi_imag$)) > 0 Then
                             RUTA_ARCH_FIRMA$ = "C:\Flexoft\" & "FIRMA" & ".TXT"
                             NARCHI3% = FILEOPEN%(RUTA_ARCH_FIRMA$, 2, 128) 'ABRE EL ARCHIVO DE LECTURA
                             Print #NARCHI3%, ruta_archi_imag$
                             Close #NARCHI3%
                           End If
                         End If
               Case Else
                 TX$ = XX$
               End Select
               If TX$ <> "" Then Print #NARCHI2%, TX$
             Loop
            '
            Close #NARCHI1%
            Close #NARCHI2%
            XXX = Shell(App.Path + "\" & "EnvioMail2.exe " & NAME_ARCH_DATMAIL$ & " " & NAME_ARCH_HTML$ & " " & "FIRMA.TXT", 4)
            '
    End If

99
   '
   
   Command14.Enabled = True
   
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
    If List1.ListIndex >= 0 Then
       List1.RemoveItem (List1.ListIndex)
    End If
99  Command3.Enabled = True
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

Private Sub enviar_Click()
    'esta oculto
    Call Command14_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then Unload Me
End Sub

Private Sub List1_DblClick()
   Linea$ = List1.List(List1.ListIndex)
   Archi$ = TRIM$(Mid$(Linea$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        'Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If

End Sub

Private Sub mnuConfigFirma_Click()
End Sub

Private Sub mnuconfigmail_Click()
   CONFIMAIL.Show 1
End Sub

Private Sub mnuConfigPlantilla_Click()
End Sub

Private Sub salir_Click()
    Unload Me
End Sub

Private Sub Text3_Change()
   NC% = XVALO(Text3)
   If NC% > 0 Then
     Label8 = rasocli$(NC%)
     Me.TXTDIRMAIL = TRIM$(EMAILCLI$(NC%))
     cont$ = CONTACTCLI$(NC%)
     If TRIM$(cont$) = "" Then
       TXTPARA = Label8
     Else
       TXTPARA = cont$
     End If
     '
     Text4 = "" ' limpio el text de proveedor
    End If
   

End Sub

Private Sub Text4_Change()
   NPRO% = XVALO(Text4)
   If XVALO(Text4) > 0 Then
     Label8 = rasocli$((-1) * NPRO%)
     Me.TXTDIRMAIL = TRIM$(EMAILCLI$((-1) * NPRO%))
     cont$ = CONTACTCLI$((-1) * NPRO%)
     If TRIM$(cont$) = "" Then
        TXTPARA = Label8
     Else
        TXTPARA = cont$
     End If

     
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



VERSION 5.00
Begin VB.Form FORTOKEN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera y Obtiene TOKEN"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Height          =   2115
      Left            =   7500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   5280
      Width           =   3500
   End
   Begin VB.TextBox Text8 
      Height          =   2115
      Left            =   7500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2520
      Width           =   3500
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9720
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6670
      TabIndex        =   14
      Top             =   1160
      Width           =   630
   End
   Begin VB.TextBox Text7 
      Height          =   2115
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   5265
      Width           =   3500
   End
   Begin VB.TextBox Text6 
      Height          =   2115
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   5280
      Width           =   3500
   End
   Begin VB.TextBox Text5 
      Height          =   870
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1200
      Width           =   6300
   End
   Begin VB.TextBox Text4 
      Height          =   2115
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2520
      Width           =   3500
   End
   Begin VB.TextBox Text3 
      Height          =   2115
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2520
      Width           =   3500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   6670
      Picture         =   "FORTOKEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aprueba y Graba"
      Top             =   1530
      Width           =   650
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   390
      Width           =   4430
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   225
      TabIndex        =   0
      Top             =   390
      Width           =   2400
   End
   Begin VB.Label Label9 
      Caption         =   "Log File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7500
      TabIndex        =   19
      Top             =   5040
      Width           =   3075
   End
   Begin VB.Label Label8 
      Caption         =   "TOKEN Factura Exportación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7560
      TabIndex        =   18
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Log File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   15
      Top             =   5040
      Width           =   3075
   End
   Begin VB.Label Label6 
      Caption         =   "Log File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   13
      Top             =   5040
      Width           =   3075
   End
   Begin VB.Label Label5 
      Caption         =   "TOKEN Bienes Capital:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   9
      Top             =   2280
      Width           =   2190
   End
   Begin VB.Label Label4 
      Caption         =   "TOKEN Factura Electronica:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   8
      Top             =   2280
      Width           =   3075
   End
   Begin VB.Label Label3 
      Caption         =   "Atributos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   975
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "U.R.L."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   180
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "C.U.I.T.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   1230
   End
End
Attribute VB_Name = "FORTOKEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GENEROELPRIMERO%
Dim GENEROELSEGUNDO%
Dim GENEROELTERCERO%
Dim ESCRIBIOLOGFILE%
Dim CUDIRA$
Dim VEMOFREN%
Dim VECES%, VECES2%, VECES3%
'
Private Sub Command1_Click()
    ' MsgBox "Al cerrar el directorio actual es " & CurDir$
    ChDrive Left$(CUDIRA$, 2)
    ChDir CUDIRA$
    ' MsgBox "Lo volví a cambiar a " & CurDir$
    Close
    Cancel = 0
    End
    Unload Me
End Sub

Private Sub Command2_Click()
       Command2.Enabled = False
       VECES% = 0: VECES2% = 0: VECES3% = 0
       ruta$ = "C:\FLEXOFT\FACELEC\"
       
       GENEROELPRIMERO% = 0
       GENEROELSEGUNDO% = 0
       GENEROELTERCERO% = 0
       ESCRIBIOLOGFILE% = 0
       '
       On Error Resume Next
       Kill ruta$ + "TOKEN.TXT"
       Kill ruta$ + "TOKEN_WSBFE.TXT"
       Kill ruta$ + "WSAA_LOG.TXT"
       Kill ruta$ + "CUTIT_COMPARAR.TXT"
       On Error GoTo 0
       Text3 = "": Text4 = ""
       Text6 = "": Text7 = ""
       Text8 = "": Text9 = ""
       '
       'Valida CUIT
       If Trim(Text1.Text) = "" Then
          MsgBox ("Falta Ingresar CUIT")
          Text1.SetFocus
          GoTo 99
       End If
      
       'Valida Certificado
       ruta$ = App.Path
       On Error Resume Next
       GetAttr (ruta$ & "\micertificado.pfx")
       If Err Then
          MsgBox ("Falta Archivo: '" & ruta$ & "\micertificado.pfx'")
          GoTo 99
       End If
       '
       'Valida URL
       DINTER$ = Trim(Text2.Text)
       If DINTER$ = "" Then GoTo 99
       If Left$(DINTER$, 8) <> "https://" Then
          Call MsgBox("Direccion URL no Válida.")
          Text2.SetFocus
          GoTo 99
       End If
       '
       'Valida si Existe ClienteLoginCms_VB
       On Error Resume Next
       GetAttr (ruta$ & "\ClienteLoginCms_VB.EXE")
       If Err Then
          MsgBox ("Falta Archivo: '" & ruta$ & "\ClienteLoginCms_VB.EXE'")
          GoTo 99
       End If
       '
       'Guarda el Link
       N1% = FreeFile
       On Error Resume Next
       Open ruta$ + "\LINK.TXT" For Output As #N1%
       Print #N1%, Text2.Text
       Close #N1%
       '
       'Guarda el Cuit
       N1% = FreeFile
       On Error Resume Next
       Open ruta$ + "\CUITEMP.TXT" For Output As #N1%
       Print #N1%, Text1.Text
       Close #N1%
       '
       'valida si existe WSFE_Cliente.EXE
       On Error Resume Next
       GetAttr (ruta$ & "\WSFE_Cliente.EXE")
       If Err Then
          ' MsgBox ("Falta Archivo: '" & RUTA$ & "\WSFE_Cliente.EXE'")
          GoTo 40
       End If
       '
       'PEDIR TOKEN
       Timer1.Enabled = True
       ATRIBU$ = Text5.Text
       XXX = Shell(ruta$ & "\ClienteLoginCms_VB.EXE" + ATRIBU$, 4)
       If Err Then
          On Error GoTo 0
          Call MsgBox("No se ha Generado el archivo: '" & ruta$ & "\TOKEN.TXT'")
          GoTo 48
       End If
       GENEROELPRIMERO% = 1
       '
       'GENERO COPIA DE ARCHIVO DE CUIT PARA COMPARAR SI COINCIDE CON EL QUE ESTA EN LA EMPRESA POSICIONADA
       TX$ = (Text1.Text)
       Open "C:\FLEXOFT\FACELEC\CUTIT_COMPARAR.txt " For Output As #1
       Write #1, TX$
       Close #1
       
       'valida si existe WSBFE_Cliente.EXE
40     On Error Resume Next
       GetAttr (ruta$ & "\WSBFE_Cliente.EXE")
       If Err Then
          GoTo 48
       End If
       '
       If GENEROELPRIMERO% > 0 Then
45      Timer1.Enabled = True
        MsgBox "ESPERE a que se Cierre la Ventana Negra y Pulse Aceptar para Continuar"
        If ESCRIBIOLOGFILE% < 1 Then GoTo 45
       End If
       '
       'PEDIR TOKEN_WSBFE.TXT
       ATRIBU$ = Text5
       PPXX% = InStr(UCase$(ATRIBU$), "WSFE")
       If PPXX% < 1 Then GoTo 99
       ATRIBU$ = Left$(ATRIBU$, PPXX% - 1) + "wsbfe" + Mid$(ATRIBU$, PPXX% + 4)
       '
       XXX = Shell(ruta$ & "\ClienteLoginCms_VB.EXE" + ATRIBU$, 4)
       If Err Then
          On Error GoTo 0
          Call MsgBox("No se ha Generado el Archivo: '" & ruta$ & "\TOKEN_WSBFE.TXT'")
          GoTo 99
       End If
       GENEROELSEGUNDO% = 1
       
       'valida si existe WSFEX_Cliente.EXE
48     Timer1.Enabled = True
       On Error Resume Next
       GetAttr (ruta$ & "\WSFEX_Cliente.EXE")
       If Err Then
          GoTo 99
       End If
       '
       If GENEROELSEGUNDO% > 0 Or GENEROELPRIMERO% > 0 Then
50      MsgBox "ESPERE a que se Cierre la Ventana Negra y Pulse Aceptar para Continuar"
        If ESCRIBIOLOGFILE% < 1 Then GoTo 50
       End If
       'PEDIR TOKEN_WSFEX.TXT
       ATRIBU$ = Text5
       PPXX% = InStr(UCase$(ATRIBU$), "WSFE")
       If PPXX% < 1 Then GoTo 99
       ATRIBU$ = Left$(ATRIBU$, PPXX% - 1) + "wsfex" + Mid$(ATRIBU$, PPXX% + 4)
       '
       XXX = Shell(ruta$ & "\ClienteLoginCms_VB.EXE" + ATRIBU$, 4)
       If Err Then
          On Error GoTo 0
          Call MsgBox("No se ha Generado el Archivo: '" & ruta$ & "\TOKEN_WSFEX.TXT'")
          GoTo 99
       End If
       GENEROELTERCERO% = 1
       '
       '
99     Command2.Enabled = True
       '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AppActivate Caption
  Me.ZOrder 0
  DoEvents
End Sub

Private Sub Form_Load()
  'valida si existe WSFEX_Cliente.EXE
  ruta$ = App.Path
  On Error Resume Next
  GetAttr (ruta$ & "\WSFEX_Cliente.EXE")
  If Err Then
     FORTOKEN.Width = 7560
  End If

  ' Leer Cuit
  CUDIRA$ = CurDir
  ruta$ = App.Path
  ' MsgBox "eL aPP.pATH ES " & App.Path
  ChDrive Left$(App.Path, 2)
  ChDir ruta$
  N1% = FreeFile
  TT$ = ""
  On Error Resume Next
  Open ruta$ + "\CUITEMP.TXT" For Input As #N1%
  If Err Then GoTo 10
  Do While Not EOF(N1%)
    Line Input #N1%, XX$
    TT$ = TT$ + XX$
  Loop
  Close #N1%
  On Error GoTo 0
  Text1.Text = TT$
  
10  'LEER LINK
  N1% = FreeFile
  TT$ = ""
  On Error Resume Next
  Open ruta$ + "\LINK.TXT" For Input As #N1%
  If Err Then GoTo 20
  Do While Not EOF(N1%)
    Line Input #1, XX$
    TT$ = TT$ + XX$
  Loop
  Close #N1%
  On Error GoTo 0
20
  DINTER$ = TT$
  '
  If DINTER$ = "" Then DINTER$ = "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL"
  Text2.Text = DINTER$
  '
  'Lista de Atributos
  ATRIBU$ = " -s " + Chr$(34) + "wsfe" + Chr$(34)
  ATRIBU$ = ATRIBU$ + " -c " + Chr$(34) + ruta$ + "\MICERTIFICADO.PFX" + Chr$(34)
  ATRIBU$ = ATRIBU$ + " -w " + Chr$(34) + DINTER$ + Chr$(34)
  ATRIBU$ = ATRIBU$ + " -v ON"
  Text5.Text = ATRIBU$
  '
  Timer2.Enabled = True
  '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = 1
End Sub

Private Sub Timer1_Timer()
   ruta$ = App.Path
   If Text3 = "" Then
      'Leer Token y mostrarlo en Text3
      If GENEROELPRIMERO% < 1 Then GoTo 30
      '
      VUELTA% = 0
11    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      Open ruta$ + "\TOKEN.TXT" For Input As #N1%
      If Err Then
          On Error GoTo 0
          GoTo 20
      End If
      On Error GoTo 0
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 11
      Text3.Text = TT$
      Text3.Refresh
   End If
   '
20 ' LEE EL LOG FILE
   If Text3 = "" Then Exit Sub
   VECES% = VECES% + 1
   ' MsgBox "PASO 1"
   If VECES% < 5 Then Exit Sub
   '
   If Text7 = "" Then
      VUELTA% = 0
21    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      ZX = FileDateTime(ruta$ + "\WSAA_LOG.TXT")
      If Err Then
          On Error GoTo 0
          GoTo 30
      End If
      ' MsgBox "PASO"
      On Error GoTo 0
      If CDbl(Now) - CDbl(ZX) < 3 / (24# * 60# * 60#) Then GoTo 30  ' DEMORA 3 SEGUNDOS
      '
      MSGERROR% = 0
      Open ruta$ + "\WSAA_LOG.TXT" For Input As #N1%
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
          If InStr(1, UCase(XX$), "ERROR") > 0 Then MSGERROR% = 1
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 21
      Text7.Text = TT$
      Text7.Refresh
      ESCRIBIOLOGFILE% = 1
   End If
   If MSGERROR% = 1 Then MsgBox "Error de la AFIP al Obtener TOKEN para Factura Electronica Comun" & vbCrLf & "Espere Un Momento y Vuelva a Intentarlo"
   '
30 ' LEE EL OTRO TOKEN
   If GENEROELSEGUNDO% < 1 Then GoTo 60
   If Text4 = "" Then
      'Leer Token y mostrarlo en Text3
      VUELTA% = 0
31    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      Open ruta$ + "\TOKEN_WSBFE.TXT" For Input As #N1%
      If Err Then
          On Error GoTo 0
          GoTo 40
      End If
      On Error GoTo 0
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 31
      Text4.Text = TT$
      Text4.Refresh
   End If
   '
40 ' LEE EL LOG FILE
   If Text4 = "" Then Exit Sub
   VECES2% = VECES2% + 1
   If VECES2% < 5 Then Exit Sub
   '
   If Text6 = "" Then
      VUELTA% = 0
41    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      ZX = FileDateTime(ruta$ + "\WSAA_LOG.TXT")
      If Err Then
          On Error GoTo 0
          GoTo 60
      End If
      On Error GoTo 0
      If CDbl(Now) - CDbl(ZX) < 3 / (24# * 60# * 60#) Then GoTo 30  ' DEMORA 3 SEGUNDOS
      '
      MSGERROR% = 0
      Open ruta$ + "\WSAA_LOG.TXT" For Input As #N1%
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
          If InStr(1, UCase(XX$), "ERROR") > 0 Then MSGERROR% = 1
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 41
      Text6.Text = TT$
      Text6.Refresh
      ESCRIBIOLOGFILE% = 1
   End If
   If MSGERROR% = 1 Then MsgBox "Error de la AFIP al Obtener TOKEN para Bienes de Capitales." & vbCrLf & "Espere Un Momento y Vuelva a Intentarlo"
   '
60  ' LEE EL OTRO TOKEN FACTURA DE EXPORTACION
   If GENEROELTERCERO% < 1 Then Exit Sub
   If Text8 = "" Then
      'Leer Token y mostrarlo en Text3
      VUELTA% = 0
61    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      Open ruta$ + "\TOKEN_WSFEX.TXT" For Input As #N1%
      If Err Then
          On Error GoTo 0
          GoTo 70
      End If
      On Error GoTo 0
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 61
      Text8.Text = TT$
      Text8.Refresh
   End If
   '
70 ' LEE EL LOG FILE
   If Text8 = "" Then Exit Sub
   VECES3% = VECES3% + 1
   If VECES3% < 5 Then Exit Sub
   '
   If Text9 = "" Then
      VUELTA% = 0
71    N1% = FreeFile
      TT$ = ""
      On Error Resume Next
      ZX = FileDateTime(ruta$ + "\WSAA_LOG.TXT")
      If Err Then
          On Error GoTo 0
          GoTo 90
      End If
      On Error GoTo 0
      If CDbl(Now) - CDbl(ZX) < 3 / (24# * 60# * 60#) Then GoTo 60  ' DEMORA 3 SEGUNDOS
      '
      MSGERROR% = 0
      Open ruta$ + "\WSAA_LOG.TXT" For Input As #N1%
      Do While Not EOF(N1%)
          Line Input #N1%, XX$
          TT$ = TT$ + vbCrLf + XX$
          If InStr(1, UCase(XX$), "ERROR") > 0 Then MSGERROR% = 1
      Loop
      Close #N1%
      If VUELTA% < 1 Then VUELTA% = 1: GoTo 71
      Text9.Text = TT$
      Text9.Refresh
      ESCRIBIOLOGFILE% = 1
   End If
   If MSGERROR% = 1 Then MsgBox "Error de la AFIP al Obtener TOKEN para Factura Electronica Exportacion" & vbCrLf & "Espere Un Momento y Vuelva a Intentarlo"
   '
   '
90 End Sub

Private Sub Timer2_Timer()
   If VEMOFREN% < 10 Then          ' que lo haga 10 veces
      VEMOFREN% = VEMOFREN% + 1
      On Error Resume Next
      Me.ZOrder 1
      AppActivate Caption
      SendKeys Chr$(0)
      On Error GoTo 0
   End If
End Sub

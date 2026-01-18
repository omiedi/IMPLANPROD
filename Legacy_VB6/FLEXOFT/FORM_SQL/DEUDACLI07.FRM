VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DEUDACLI07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composición de Deuda del Cliente"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   4260
      Left            =   9840
      ScaleHeight     =   4230
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   0
      Width           =   1905
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
         Height          =   690
         Left            =   105
         Picture         =   "DEUDACLI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Deuda / Envio por Mail"
         Top             =   2180
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
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
         Left            =   100
         Picture         =   "DEUDACLI07.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Búsqueda por Texto del Detalle"
         Top             =   1490
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   100
         Picture         =   "DEUDACLI07.frx":19D0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Selección / deselección Automática"
         Top             =   800
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   100
         Picture         =   "DEUDACLI07.frx":1B1A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   650
      End
   End
   Begin VB.ListBox LICHESEL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      ItemData        =   "DEUDACLI07.frx":1E24
      Left            =   80
      List            =   "DEUDACLI07.frx":1E26
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   9675
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "DEUDACLI07.frx":1E28
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "DEUDACLI07.frx":205C
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label82 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8960
      TabIndex        =   18
      Top             =   115
      Width           =   840
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6500
      TabIndex        =   17
      Top             =   115
      Width           =   1080
   End
   Begin VB.Label Label84 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4895
      TabIndex        =   16
      Top             =   115
      Width           =   1440
   End
   Begin VB.Label Label85 
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2600
      TabIndex        =   15
      Top             =   115
      Width           =   840
   End
   Begin VB.Label Label86 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   14
      Top             =   115
      Width           =   1455
   End
   Begin VB.Label Label87 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3545
      TabIndex        =   13
      Top             =   115
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DEUDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5970
      TabIndex        =   12
      Top             =   3735
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECCIONADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      TabIndex        =   11
      Top             =   3735
      Width           =   1305
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Height          =   300
      Left            =   7440
      TabIndex        =   10
      Top             =   3615
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Height          =   300
      Left            =   4185
      TabIndex        =   9
      Top             =   3615
      Width           =   2010
   End
   Begin VB.Label Label88 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   80
      TabIndex        =   19
      Top             =   70
      Width           =   9675
   End
End
Attribute VB_Name = "DEUDACLI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub pruebahtml()
    'TITULOS DE DETALLE
    ENCA$ = ""
    TXT1$ = "<!DOCTYPE html>"
    TXT1$ = TXT1$ + "<html>"
    TXT1$ = TXT1$ + "<body>"
    TXT1$ = TXT1$ + "<table border=" & Chr(34) & "1" & Chr(34) & ">"
    TXT1$ = TXT1$ + "<tr>"
    TXT1$ = TXT1$ + "<td>China</td><td>1300 millones</td>"
    TXT1$ = TXT1$ + "</tr>"
    TXT1$ = TXT1$ + "<tr>"
    TXT1$ = TXT1$ + "<td>India</td><td>1080 millones</td>"
    TXT1$ = TXT1$ + "</tr>"
    TXT1$ = TXT1$ + "<tr>"
    TXT1$ = TXT1$ + "<td>Estados Unidos</td><td>295 millones</td>"
    TXT1$ = TXT1$ + "</tr>"
    TXT1$ = TXT1$ + "</table>"
    TXT1$ = TXT1$ + "</body>"
    TXT1$ = TXT1$ + "</html>"
    Call REPLA(ENCA$, "Documento", 1, 18)
    Call REPLA(ENCA$, "   Nro.", 19, 6)
    Call REPLA(ENCA$, "Emisión", 28, 8)
    Call REPLA(ENCA$, "Importe", 44, 8)
    Call REPLA(ENCA$, "Vto.", 57, 8)
    Call REPLA(ENCA$, "Saldo", 72, 8)
    ENCA$ = ENCA$ + vbCrLf + String$(77, "=") + vbCrLf
    '
    'CONDICION DE PAGO DEL CLIENTE
    CONPAG% = CPAGCLI%(NC%)
    CONDPAGO$ = ECOARCH$("CONPAG", MKI$(CONPAG%))
   '
   'CARGA DE DETALLE
    TX$ = ""
    For i& = 1 To LICHESEL.ListCount
      TX$ = LICHESEL.LIST(i& - 1)
      'SI ESTA SELECCIONADA SE LE AGREGA UN (*)
      If LICHESEL.Selected(i& - 1) = True Then TX$ = TX$ + " (*)"
      TXTX$ = TXTX$ + TX$ + vbCrLf
    Next i&
    '
    'CARGA DE ENCABEZADO
    TXTX$ = TXTX$ + vbCrLf + String$(77, "=")
    ENCA2$ = "Sres." + vbCrLf + rasocli(NC%) + vbCrLf + vbCrLf
    ENCA2$ = ENCA2$ + "Condición de Pago: " + CONDPAGO$ + vbCrLf + vbCrLf
    ENCA2$ = ENCA2$ + "Nos dirigimos a Uds a fin de recordarles " + vbCrLf
    ENCA2$ = ENCA2$ + "el proximo vencimiento de la(s) siguiente(s) facturas(s)" + vbCrLf
    textox$ = ENCA2$ + vbCrLf + ENCA$ + vbCrLf + TXTX$ + vbCrLf
    '
    'AGREGO AL FINAL LOS VALORES SEECCIONADOS Y TOTAL DEUDA
    Call REPLA(VALORES$, "Seleccionados: " & FORMATNUM$((XVALO(label2)), "F12.2"), 10, 27)
    Call REPLA(VALORES$, "   Total Deuda: " & FORMATNUM$((XVALO(Label15)), "F12.2"), 39, 30)
    '
    textox$ = textox$ + VALORES$ + vbCrLf + String$(77, "=")
    textox$ = textox$ & vbCrLf & vbCrLf & "<html><b>" & NOMBREFANTASIA$ & "</b></html>"
    '
    'si existe la impresora pdf genera un pdf y adjunta

End Sub

Private Sub BOTEXIT_Click()
    LICHESEL.Clear
    label2.Caption = ""
    Label15.Caption = ""
    Hide
End Sub
'
Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICHESEL.TopIndex
   For U& = 1 To LICHESEL.ListCount
     If LICHESEL.Selected(U& - 1) = True Then
         LICHESEL.Selected(U& - 1) = False
       Else
         LICHESEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICHESEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub
'
Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.LIST(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    OPX% = COMALTER%("Opciones de Salida\\Listar\Envio x Mail")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    NC% = XVALO(ACLI_GES07.Text1)
    
    'ENVIO POR MAIL
    If OPX% = 1 Then
      Call LISTADE(NC%)
    ElseIf OPX% = 2 Then
      Call MAILDEU(NC%)
    End If
    
99  Command3.Enabled = True
End Sub
Sub MAILDEU(NC%)
    '
    
    If TRIM$(EMAILCLI(NC%)) = "" Or InStr(EMAILCLI(NC%), "@") < 1 Then 'TIENE DIRECCION VALIDA
        Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
        Exit Sub
    End If
    '
    Dim XXXX$(8)
    Call CAREMPRE(LUDAT$, XXXX$)
    NOMBREFANTASIA$ = XXXX$(1) 'NOMBRE FANTASIA QUE SE GUARDA EN SETUP BOTON EMPRESA
    
    'TITULOS DE DETALLE
    ENCA$ = ""
    Call REPLA(ENCA$, "Documento", 1, 18)
    Call REPLA(ENCA$, "   Nro.", 19, 6)
    Call REPLA(ENCA$, "Emisión", 28, 8)
    Call REPLA(ENCA$, "Importe", 44, 8)
    Call REPLA(ENCA$, "Vto.", 57, 8)
    Call REPLA(ENCA$, "Saldo", 72, 8)
    ENCA$ = ENCA$ + vbCrLf + String$(77, "=") + vbCrLf
    '
    'CONDICION DE PAGO DEL CLIENTE
    CONPAG% = CPAGCLI%(NC%)
    CONDPAGO$ = ECOARCH$("CONPAG", MKI$(CONPAG%))
   '
   'CARGA DE DETALLE
    TX$ = ""
    For i& = 1 To LICHESEL.ListCount
      TX$ = LICHESEL.LIST(i& - 1)
      'SI ESTA SELECCIONADA SE LE AGREGA UN (*)
      If LICHESEL.Selected(i& - 1) = True Then TX$ = TX$ + " (*)"
      TXTX$ = TXTX$ + TX$ + vbCrLf
    Next i&
    '
    'CARGA DE ENCABEZADO
    TXTX$ = TXTX$ + vbCrLf + String$(77, "=")
    ENCA2$ = "Sres." + vbCrLf + rasocli(NC%) + vbCrLf + vbCrLf
    ENCA2$ = ENCA2$ + "Condición de Pago: " + CONDPAGO$ + vbCrLf + vbCrLf
    ENCA2$ = ENCA2$ + "Nos dirigimos a Uds a fin de recordarles " + vbCrLf
    ENCA2$ = ENCA2$ + "el proximo vencimiento de la(s) siguiente(s) facturas(s)" + vbCrLf
    textox$ = ENCA2$ + vbCrLf + ENCA$ + vbCrLf + TXTX$ + vbCrLf
    '
    'AGREGO AL FINAL LOS VALORES SEECCIONADOS Y TOTAL DEUDA
    Call REPLA(VALORES$, "Seleccionados: " & FORMATNUM$((XVALO(label2)), "F12.2"), 10, 27)
    Call REPLA(VALORES$, "   Total Deuda: " & FORMATNUM$((XVALO(Label15)), "F12.2"), 39, 30)
    '
    textox$ = textox$ + VALORES$ + vbCrLf + String$(77, "=")
    textox$ = textox$ + vbCrLf + String$(77, "=")
    lonx% = Len(NOMBREFANTASIA$)
    If lonx% >= 77 Then NOMBREFANTASIA$ = Left$(NOMBREFANTASIA$, 77): laterales% = 0: GoTo 80
    laterales% = Int((77 - lonx%) / 2)
80  textox$ = textox$ + vbCrLf + String$(laterales%, " ") + NOMBREFANTASIA$ + String$(laterales%, " ")
    textox$ = textox$ + vbCrLf + String$(77, "=")
    
    '
    'si existe la impresora pdf genera un pdf y adjunta
'    If EXISTEIMPRESORAPDF% > 0 Then
'       Call GeneraArchivoPDFdeTXT("C:\FLEXOFT\DEUDACLISS.TXT", TEXTOX$, "DEUDACLISS")
'       MAI_TEXTO$ = "Adjuntamos Saldos de Cuenta Corriente"
'       MAI_ADJUN$(0) = LUCOM$ + "" + TRIM$("DEUDACLISS") + ".PDF"
'    Else
     MAI_TEXTO$ = textox$
'    End If
    'ENVIO DE MAIL
    MAI_DESTI$ = rasocli$(NC%)
    MAI_DIREL$ = EMAILCLI(NC%)
    MAI_ASUNT$ = "Informe de Saldo"

    Call GENERAMAIL
''''''    ESTILO HTML FUNCIONA PERFECTO PERO A AHP QUE TIENE OUTLOOK LE MUESTRA EL CODIGO HTML , NO LO INTERPRETA

''''''    If TRIM$(EMAILCLI(NC%)) = "" Or InStr(EMAILCLI(NC%), "@") < 1 Then 'TIENE DIRECCION VALIDA
''''''        Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
''''''        Exit Sub
''''''    End If
''''''    '
''''''    Dim XXXX$(8)
''''''    Call CAREMPRE(LUDAT$, XXXX$)
''''''    NOMBREFANTASIA$ = XXXX$(1) 'NOMBRE FANTASIA QUE SE GUARDA EN SETUP BOTON EMPRESA
''''''    '
''''''    'CONDICION DE PAGO DEL CLIENTE
''''''    CONPAG% = CPAGCLI%(NC%)
''''''    CONDPAGO$ = ECOARCH$("CONPAG", MKI$(CONPAG%))
''''''    '
''''''    'CARGA DE ENCABEZADO
''''''    TXTX$ = TXTX$ + vbCrLf + String$(77, "=")
''''''    ENCA2$ = "Sres." + vbCrLf + rasocli(NC%) + vbCrLf + vbCrLf
''''''    ENCA2$ = ENCA2$ + "Condición de Pago: " + CONDPAGO$ + vbCrLf + vbCrLf
''''''    ENCA2$ = ENCA2$ + "Nos dirigimos a Uds a fin de recordarles " + vbCrLf
''''''    ENCA2$ = ENCA2$ + "el proximo vencimiento de la(s) siguiente(s) facturas(s)" + vbCrLf
''''''    TEXTOX$ = ENCA2$ + vbCrLf + ENCA$ + vbCrLf + TXTX$ + vbCrLf
''''''    '
''''''    'AGREGO AL FINAL LOS VALORES SEECCIONADOS Y TOTAL DEUDA
''''''    Call REPLA(VALORES$, "Seleccionados: " & FORMATNUM$((XVALO(Label2)), "F12.2"), 10, 27)
''''''    Call REPLA(VALORES$, "   Total Deuda: " & FORMATNUM$((XVALO(Label15)), "F12.2"), 39, 30)
''''''    '
''''''    TEXTOX$ = TEXTOX$ + VALORES$ + vbCrLf + String$(77, "=")
''''''    TEXTOX$ = TEXTOX$ & vbCrLf & vbCrLf & "<html><b>" & NOMBREFANTASIA$ & "</b></html>"
''''''    '
''''''    TXT1$ = "<!DOCTYPE html>"
''''''    TXT1$ = TXT1$ + "<html>"
''''''    TXT1$ = TXT1$ + "<body>"
''''''    TXT1$ = TXT1$ + "<br><br>"
''''''    'CARGA UN TEXTO DE ENCABEZADO
''''''    ENCA0$ = "Señores: " & rasocli(NC%)
''''''    ENCA1$ = "Nos dirigimos a Uds a fin de recordarles el proximo vencimiento de la(s) siguiente(s) facturas(s)"
''''''    ENCA2$ = "Condición de Pago: " & CONDPAGO$
''''''    TXT1$ = TXT1$ + "<h2><center>" & ENCA0$ & "</center></h2>"
''''''    TXT1$ = TXT1$ + "<br><br>"
''''''    TXT1$ = TXT1$ + "<h2><center>" & ENCA1$ & "</center></h2>"
''''''    TXT1$ = TXT1$ + "<br><br>"
''''''    TXT1$ = TXT1$ + "<h3><center>" & ENCA2$ & "</center></h3>"
''''''    'CARGA EL ENCABEZADO DE LA TABLA
''''''    TXT1$ = TXT1$ + "<table border=" & Chr(34) & "1" & Chr(34) & "style=" & Chr(34) & "margin: 0 auto;" & Chr(34) & ">"
''''''    TXT1$ = TXT1$ + "<tr>"
''''''    TXT1$ = TXT1$ + "<td>Documento</td><td>Numero</td><td>Emisión</td><td>Importe</td><td>Vencimiento</td><td>Saldo</td>"
''''''    TXT1$ = TXT1$ + "</tr>"
''''''    TXT1$ = TXT1$ + "<tr>"
''''''    'CARGA EL DETALLE EN LA TABLA
''''''    For i& = 1 To LICHESEL.ListCount
''''''      TX$ = LICHESEL.List(i& - 1)
''''''      If LICHESEL.Selected(i& - 1) = True Then
''''''        TXT1$ = TXT1$ + "<tr>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 1, 18) & "</td>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 19, 6) & "</td>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 27, 8) & "</td>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 35, 16) & "</td>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 51, 10) & "</td>"
''''''        TXT1$ = TXT1$ + "<td style=" & Chr(34) & "text-align:center;" & Chr(34) & ">" & Mid$(TX$, 61, 17) & "</td>"
''''''        TXT1$ = TXT1$ + "</tr>"
''''''      End If
''''''    Next i&
''''''    TXT1$ = TXT1$ + "<td></td><td></td><td></td><td></td>"
''''''    TXT1$ = TXT1$ + "<td>" & "   Total Deuda: " & "</td>"
''''''    TXT1$ = TXT1$ + "<td>" & FORMATNUM$((XVALO(Label15)), "F12.2") & "</td>"
''''''    TXT1$ = TXT1$ + "</tr>"
''''''    TXT1$ = TXT1$ + "</table>"
''''''    TXT1$ = TXT1$ + "<br><br>"
''''''    TXT1$ = TXT1$ + "<h2><center>" & NOMBREFANTASIA$ & "</center></h2>"
''''''    'EL CODIGO DE LA IMAGEN NO ME FUNCIONO
'''''''    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
'''''''    ISOLOGO$ = LUDAT$ + ISOLOGO$
'''''''    CODIGOHTML = "<p align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & ISOLOGO$ & Chr(34) & "width=" & Chr(34) & "122" & Chr(34) & "height=" & Chr(34) & "85> </p>"
'''''''    TXT1$ = TXT1$ + CODIGOHTML
''''''    TXT1$ = TXT1$ + "</body>"
''''''    TXT1$ = TXT1$ + "</html>"
''''''    MAI_TEXTO$ = TXT1$
''''''    'ENVIO DE MAIL
''''''    MAI_DESTI$ = rasocli$(NC%)
''''''    MAI_DIREL$ = EMAILCLI(NC%)
''''''    MAI_ASUNT$ = "Informe de Saldo"
''''''
''''''    Call GENERAMAIL
'
End Sub

Sub GeneraArchivoPDFdeTXT(RUTANOMBRE$, TEXTOPDF$, ARANEX$)
    
        If UCase$(Right(RUTANOMBRE$, 3)) <> "TXT" Then MsgBox "El archivo no es *.txt", vbCritical + vbOKOnly, "ERROR": Exit Sub
        '
        If EXISTEIMPRESORAPDF% <= 0 Then
           Call COMUNI("Imposible Generar Archivo PDF No se Detectó la Impresora")
           Exit Sub
        End If
        '
        'GENERA UN TXT CON EL TEXTO PASADO
        N1% = FILEOPEN%(RUTANOMBRE$, 2, 128)
        Print #N1%, TEXTOPDF$
        Close #N1%
    
        'Se abre el archivo
        Open RUTANOMBRE$ For Binary As #N1%
        'N1% = FILEOPEN%(TextFileName, 1, 128)
        TextToPrint$ = Space(LOF(1)) 'Se crea el buffer
        TextToPrint$ = TEXTOPDF$
        Get #N1%, , TextToPrint$ 'Se almacenan los datos en el buffer
        Close #N1% 'Se cierra el archivo
    
        Printer.Print TextToPrint$ 'Se imprimen los datos del buffer
        '
        'BUSCA LA IMPRESORA Y GENERA EL ARCHIVO
        For Each PRX1 In Printers
           ImpresoraNombre = PRX1.DeviceName
           If InStr(UCase$(ImpresoraNombre), "PDFWRITER") > 0 Or InStr(UCase$(ImpresoraNombre), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then
                
                If InStr(UCase$(ImpresoraNombre), "BULLZIP") > 0 Then ImpresoraNombre = "BULLZIP"
                If InStr(UCase$(ImpresoraNombre), "BIOPDF") > 0 Then ImpresoraNombre = "BIOPDF"
                SETTINGS_PROGID = ImpresoraNombre & ".PDFSettings"
                UTIL_PROGID = ImpresoraNombre & ".PDFUtil"
                Dim prtidx As Integer
                Dim sPrinterName As String
                Dim settings As Object
                Dim util As Object
                
                Set util = CreateObject(UTIL_PROGID)
                sPrinterName = util.defaultprintername
                
                Rem -- Configure the PDF print job
                Set settings = CreateObject(SETTINGS_PROGID)
                settings.printerName = sPrinterName
                settings.SetValue "Output", LUCOM$ + ARANEX$ & ".pdf"
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
                ModoBullzip% = 1
                Exit For
             End If
     Next

End Sub

'Cómo usarlo:

'La implementación de éste código no necesita explicaciones.
'Se llama a un archivo de texto que será imprimido:

'PrintTextFile "C:\texto.txt"

Sub LISTADE(NC%)
    '
    'ENCABEZADO
    ENCA$ = ""
    Call REPLA(ENCA$, "Documento", 1, 18)
    Call REPLA(ENCA$, "   Nro.", 19, 6)
    Call REPLA(ENCA$, "Emisión", 28, 8)
    Call REPLA(ENCA$, "Importe", 44, 8)
    Call REPLA(ENCA$, "Vto.", 57, 8)
    Call REPLA(ENCA$, "Saldo", 72, 8)
    '
    'GRABA EL ENCABEZADO
    Call GRAREG("DEUDACLI", ENCA$, 1)
    Call GRAREG("DEUDACLI", Space$(80), 2)
    'CARGA DE DETALLE
    TX$ = ""
    For i& = 1 To LICHESEL.ListCount
      TX$ = Mid$(LICHESEL.LIST(i& - 1), 1, 80)
      If TX$ <> "" Then
        'SI ESTA SELECCIONADA SE LE AGREGA UN (*)
        If LICHESEL.Selected(i& - 1) = True Then TX$ = TX$ + " (*)"
        Call GRAREG("DEUDACLI", TX$, i& + 2)
      End If
    Next i&
    '
    Call GRAREG("DEUDACLI", String$(77, "="), i& + 2)
    'AGREGO AL FINAL LOS VALORES SEECCIONADOS Y TOTAL DEUDA
    Call REPLA(VALORES$, "Seleccionados: " & Str(XVALO(label2)), 10, 27)
    Call REPLA(VALORES$, "   Total Deuda: " & Str(XVALO(Label15)), 33, 25)
    '
    Call GRAREG("DEUDACLI", VALORES$, i& + 3)
    '
    Call SETULTREG("DEUDACLI", i& + 3)
    Call CIERRARCH("DEUDACLI")
    Call CIERRARCH("*.*")

    '
    'CONDICION DE PAGO DEL CLIENTE
    CONPAG% = CPAGCLI%(NC%)
    CONDPAGO$ = ECOARCH$("CONPAG", MKI$(CONPAG%))
    DATCLIENTE$ = "Cliente: (" & TRIM$(Str(NC%)) & ") - " & rasocli$(NC%)
    '
    Call HEADERS("DEUDACLI", "")
    Call HEADERS("DEUDACLI", DATCLIENTE$)
    Call HEADERS("DEUDACLI", "Cond.Pago: " + CONDPAGO$)
    '
    Call FINAL("*DEUDACLI")
    '
End Sub

Private Sub Form_Activate()
    Call SUCHESEL
    If LICHESEL.ListCount > 0 Then LICHESEL.ListIndex = 0
End Sub

Sub SUCHESEL()
    SUCHESE# = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.LIST(i& - 1), 63, 14))
       End If
    Next i&
    label2.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    label2.Refresh
End Sub


Private Sub Form_Load()
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
End Sub

Private Sub LICHESEL_Click()
   Call SUCHESEL
End Sub





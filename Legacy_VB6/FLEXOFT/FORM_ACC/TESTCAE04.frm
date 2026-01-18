VERSION 5.00
Begin VB.Form TESTCAE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de CAE Manual"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6180
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4200
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4680
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3480
      Width           =   5010
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1245
      TabIndex        =   9
      Top             =   2760
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1245
      TabIndex        =   8
      Top             =   2280
      Width           =   2520
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   6255
      Left            =   5280
      ScaleHeight     =   6195
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox Picture16 
         Height          =   6200
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   1020
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   6200
            Left            =   0
            Picture         =   "TESTCAE04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sol. CAE con ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         TabIndex        =   25
         Top             =   2640
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "TESTCAE04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Re-Impresión de Comprobantes"
         Top             =   960
         Width           =   600
      End
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
         Height          =   600
         Left            =   120
         Picture         =   "TESTCAE04.frx":2A02
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Proceso Automático de Solicitud de Cae Pendientes"
         Top             =   4680
         Visible         =   0   'False
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "TESTCAE04.frx":2B4C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5400
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "TESTCAE04.frx":2C96
         Top             =   5640
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   140
      TabIndex        =   11
      Top             =   120
      Width           =   5040
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Este boton pide el CAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitud de CAE Manual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1575
         TabIndex        =   21
         ToolTipText     =   "Solicitud de Cae Manual"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   20
         Top             =   960
         Width           =   2805
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   17
         Top             =   600
         Width           =   165
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solicitar CAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   570
      End
      Begin VB.Label MODOAUTO 
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id Transaccion:"
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
         Left            =   -120
         TabIndex        =   19
         Top             =   720
         Width           =   1890
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Provisorio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Punto de Venta:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1410
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vto.:"
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
      Left            =   3750
      TabIndex        =   28
      Top             =   2355
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4230
      TabIndex        =   27
      Top             =   2280
      Width           =   940
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cae Devuelto:"
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
      Left            =   -570
      TabIndex        =   6
      Top             =   2355
      Width           =   1770
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Factura Asignado:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -330
      TabIndex        =   5
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Recibido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu reimpresiondefacturaselectronicas 
         Caption         =   "Re-Impresión de Facturas Electrónicas"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "TESTCAE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ESTADO%
Dim FACTU$()
Dim CONTAD&
Dim INDICE&
Dim MODO%
Dim CANTCOPIAS%
Dim FILTX$
Dim ARCSQR$
Dim SOLOCAE%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
End Sub

Sub Command4_Click()
    ' SOLOCAE% = 0
    Text3 = ""
    Command4.Enabled = False
    Screen.MousePointer = 11
    '
    Timer1.Enabled = False
    Call SOLICITACAE
    '
End Sub

Sub SOLICITACAE()
      EJECTUTA$ = "C:\FLEXOFT\FACELEC\"

      PVTA% = XVALO(Text4)
      IDENCAE$ = TRIM$(Text5)
      If PVTA% < 1 Or IDENCAE$ = "" Then
           If MODOAUTO = "NO" Then Call COMUNI("Imposible Solicitar CAE.\Datos Incompletos o Invalidos.")
           Exit Sub
         Else
           On Error Resume Next
           Kill EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT" ' ELIMINAMOS EL ARCHIVO
           On Error GoTo 0
           If EXISTE%(EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT") > 0 Then
              If MODOAUTO = "NO" Then Call MENSERR(24, "Imposible Solicitar CAE.\No se Pudo Eliminar Archivo '" + REPLACAR$(EJECTUTA$ + "CAE_" + IDENCAE$ + ".TXT'", "\", "/"))
              Exit Sub
           End If
           Call PIDECAE ' (PVTA%, IDENCAE$)
      End If

''      INDICE& = 1
''      ReDim FACTU$(INDICE&)
''      FACTU$(INDICE&) = IDENCAE$
      
''      ESTADO% = 0
''      CONTAD& = 0
''      Timer2.Enabled = True
      
End Sub

Sub PIDECAE() ' (PVTA%, ID$)
     '
     PVTA% = XVALO(Text4)
     IDENCAE$ = TRIM$(Text5)
     EJECTUTA$ = "C:\FLEXOFT\FACELEC\"
     '
     COMPROB$ = Mid$(Text6, 1, 2)
     VARIANTE$ = Mid$(Text6, 18, 1)
     TIPOCOM% = 1                 ' valor por omision = FACTURA
     Select Case COMPROB$
           Case "FC": If VARIANTE$ = "B" Then TIPOCOM% = 6
           Case "ND": TIPOCOM% = 2: If VARIANTE$ = "B" Then TIPOCOM% = 7
           Case "FC": TIPOCOM% = 3: If VARIANTE$ = "B" Then TIPOCOM% = 8
     End Select
     '
     PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
     PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
     PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
     '
     Select Case PVTA%
        Case XVALO(PVTANORMAL$)
           If EXISTE%(EJECTUTA$ + "wsfe_cliente.exe") = 1 Then
              XXX = Shell(EJECTUTA$ & "wsfe_cliente.exe AUTO " & IDENCAE$, 4)
           End If
        Case XVALO(PVTABCPTAL$)
          If EXISTE%(EJECTUTA$ + "wsbfe_cliente.exe") = 1 Then
            XXX = Shell(EJECTUTA$ & "wsbfe_cliente.exe AUTO " & IDENCAE$ & " " & TIPOCOM% & " " & PVTA%, 4)
          End If
        Case XVALO(PVTAFAEXP$)
          If EXISTE%(EJECTUTA$ + "wsfex_cliente.exe") = 1 Then
           XXX = Shell(EJECTUTA$ & "wsfex_cliente.exe AUTO " & IDENCAE$, 4)
          End If
        Case Else
           If MODOAUTO = "NO" Then Call COMUNI("No se Ha Encontrado el Punto de Venta " & PVTA%)
           Exit Sub
     End Select
     '
     Timer1.Enabled = True    ' DISPARA LAZO DE ESPERA DE LA RESPUESTA DE AFIP
     '
End Sub

Function FORELFACTU$(LETR$, TIPOCOM$, TIPOFACT$, TIFADOC%) ' DEVUELVE EL NOMBRE DEL FORMULARIO ELECTRONICO DE FACTURACIÓN
    '
    TIPOFACT$ = UCase$(TIPOFACT$)
    LTFX1$ = UCase$(LETR$)
    If LTFX1$ = "C" Then LTFX1$ = "A"
    FEFAC$ = TRIM$(CONTROL$("", "FORFAC-" + LTFX1$))
    If TIPOCOM$ = "(T.Libre)" Then
      If LTFX1$ <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FEFAC$ = TRIM$(CONTROL$("", "FACTEXLI"))
        If LTFX1$ = "B" Then
          FEFAC$ = TRIM$(CONTROL$("", "FACTELIB"))
        End If
      End If
    End If
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIPOFACT$ = "NC" Or TIPOFACT$ = "ND" Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If LTFX1$ = "A" Or LTFX1$ = "C" Then
          FORIALT$ = CONTROL$("", "FORNDE-A") 'ND-A
          ElseIf LTFX1$ = "B" Then
          FORIALT$ = CONTROL$("", "FORNDE-B") 'ND-B
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If LTFX1$ = "A" Then
          FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
        ElseIf LTFX1$ = "B" Then
          FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FEFAC$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    '
    ' VALIDACIONES
    If FEFAC$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       FEFAC$ = "NULL"
    End If
    '
    If EXISTE%(LUDAT$ + FEFAC$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FEFAC$ + ".FRM'.")
       FEFAC$ = "NULL"
    End If
    '
    FORELFACTU$ = FEFAC$
    '
End Function

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5800
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

End Sub

Private Sub reimpresiondefacturaselectronicas_Click()
   Call Command16_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text3_Change()
    
    If Command4.Enabled = False And Screen.MousePointer = 11 Then
       Command4.Enabled = True:  Screen.MousePointer = 1
    End If

End Sub

'
Private Sub Timer1_Timer()
   '
   Static CONTA%
   
   CONTA% = CONTA% + 1 ' SI SUPERA LOS 40 SEGUNDOS DE DEMORA EN TRAER EL CAE PRESENTA MENSAJE
   If CONTA% > 40 Then
'      If COMALTER%("No Se Ha Podido Obtener el CAE \\Seguir Intentando\Cancelar Factura") <> 1 Then
         'Call FINAL("")
      Timer1.Enabled = False
''         If MODO% = 1 And ESTADO% = 1 Then Timer3.Enabled = True
''         If MODO% = 2 And ESTADO% = 1 Then ESTADO% = 0
'      End If
      CONTA% = 0
   End If
   '
'   IDENTCAE$ = TRIM$(Text5)
   KAE$ = LEECAE$    '  (IDENTCAE$)
   '
End Sub

Function LEECAE$() ' (IDENTCAE$)
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LEECAE$ = ""
    CAE$ = ""
    TX1$ = ""
    IDENCAE$ = TRIM$(Text5)
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cae>")) + 5
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             CAE$ = Mid$(TXT$, 1, POS2% - 1)
             'caso de punto de venta 4
          ElseIf InStr(UCase$(XXXXX$), UCase$("<cbt_desde>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cbt_desde>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             NUFACT$ = Mid$(TXT$, 1, POS2% - 1)
             'FECHA DE VENCIMINTO DE CAE SEGUN PUNTO DE VENTA 4
          ElseIf InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             VtoCae$ = Mid$(TXT$, 1, POS2% - 1)

             'caso de punto de venta 5
          ElseIf InStr(UCase$(XXXXX$), UCase$("<cbte_nro>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cbte_nro>")) + 10
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             NUFACT$ = Mid$(TXT$, 1, POS2% - 1)
          'FECHA DE VENCIMINTO DE CAE SEGUN PUNTO DE VENTA 5
          ElseIf InStr(UCase$(XXXXX$), UCase$("<fch_venc_cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<fch_venc_cae>")) + 14
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             VtoCae$ = Mid$(TXT$, 1, POS2% - 1)
          End If
        Loop
        Close NARCHIVO%
    End If
    '
    Text3 = TX1$
    '
    If XVALO(CAE$) > 0 Then
        Timer1.Enabled = False
        Text1 = CAE$
        Label6 = VtoCae$
        Text2 = NUFACT$
        ' If ESTADO% = 1 And MODO% > 0 Then ESTADO% = 0
    End If
    '
    If InStr(TRIM$(TX1$), "CANCELAR") > 0 Then
        Timer1.Enabled = False
        Call PROCESA_NO_OBTUVO_CAE
        'If MODO% = 1 And ESTADO% = 1 Then Timer3.Enabled = True
        'If MODO% = 2 And ESTADO% = 1 Then ESTADO% = 0
    End If
    '
    LEECAE$ = CAE$
    '
End Function

Sub PROCESA_NO_OBTUVO_CAE()
     '
     TX$ = UCase$(Text3)
     '
     If InStr(TX$, "ID DE TRANSACCION NO REGISTRADO") > 0 Then
           Call PIDENUEVOCAE
        ElseIf InStr(TX$, "CORRELATIVO") > 0 Then
             Call OBTIENEIDEDEREEMPLAZO
             Call PIDENUEVOCAE
        Else
           If MODOAUTO = "NO" Then Call COMUNI("No se Pudo Obtener el Cae ")
     End If
     '
''     Select Case XXX ... PROCESANDO LO QUE ESTA EN EL TEXT3
''        CASE ID NO REGISTRADO
''           Call PIDENUEVOCAE
''       CASE NO CORRELATIVO
''           Call OBTIENEIDEDEREEMPLAZO
''           CALL REEMPLAZA IDE EN CAJABANC
''
''       Case Else
''         mENSAJE Y EXIT SUB
''
''      End Select
End Sub
'
Sub PIDENUEVOCAE()
      '
      MODO% = 2
      PVTA% = XVALO(Text4)
      IDENCAE$ = TRIM$(Text5)
      '
      PVTANORMAL$ = CONTROL("FACTURA", "PUVENELE") 'P VTA. NORMAL
      PVTABCPTAL$ = CONTROL("FACTURA", "PUVENCAP") 'P.VTA. B.CAPITAL
      PVTAFAEXP$ = CONTROL("FACTURA", "PUVEFAEX")  'P.VTA. FACT.EXP.
      '
      If PVTA% < 1 And IDENCAE$ = "" Then
          If MODOAUTO = "NO" Then Call COMUNI("Faltan Datos ")
         Exit Sub
      End If
      '
      Text1 = ""
      Text2 = ""
      Text3 = ""
      '
      Select Case PVTA%
         '
         Case XVALO(PVTANORMAL$)
              Call SOLICITA_CAE_ID(IDENCAE$)
         Case XVALO(PVTABCPTAL$)
              Call SOLICITA_CAE_BCAP_ID(IDENCAE$)
         Case XVALO(PVTAFAEXP$)
              Call SOLICITA_CAE_EXPO_ID(IDENCAE$)
      End Select

      Timer1.Enabled = True
      '
End Sub
Sub OBTIENEIDEDEREEMPLAZO()
   '
   IDENCAE$ = TRIM$(Text5)
   '
   'BUSCO EL VALOR MAXIMO DEL IDENCAE
   SQLX$ = "SELECT MAX(Val(Mid(IDCAE,1 ))) AS CANTIDAD FROM CAJABANC WHERE IDCAE IS NOT NULL AND TIPOMOVIM = 'FVEN'"
   Set MAXIDCAE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   IdentiCaeMax$ = TRIM$(Str$(1 + XVALO(MAXIDCAE!CANTIDAD)))
   '
   'GUARDO EL CAE NUEVO REEMPLAZANDO EL ANTERIOR
''''   CONDI$ = "IDCAE = '" & IDENCAE$ & "'"
''''   Call ACTUREGISTRO("CAJABANC", "IDCAE", CONDI$, IdentiCaeMax$, REAF&, "NOMESS")
   Call APERBASDAT("CABAN")
   SQLX$ = "UPDATE CAJABANC SET IDCAE = '" & IdentiCaeMax$ & "'"
   SQLX$ = SQLX$ + " WHERE IDCAE = '" & IDENCAE$ & "'"
   CueCorri.Execute (SQLX$)
   '
   'ACTUALIZO LA CAJA DE TEXTO DEL IDENTIFCADOR DE CAE
   Text5 = IdentiCaeMax$
   '
End Sub

Private Sub Command16_Click()
    '
    ' ESTE BOTÓN ES PARA RE-IMPRIMIR COMPROBANTES
    '
    Command16.Enabled = False
    Screen.MousePointer = 11
    Call COPYSTRU("FACTUSEL", "LISCOMP")
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT DISTINCT CODICOM_LAR, FechEmisi,NuClien FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN' AND NUMECOMP > 0 ORDER BY FECHEMISI, CODICOM_LAR "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    Call SETULTREG("!LISCOMP", 0)
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          TX$ = ""
          DOCUM$ = SAFETEXT$(!CODICOM_LAR)
          NC% = XVALO(!NuClien)
          RACLI$ = Left$(RASOCLI$(NC%), 30)
          FE% = CVINT(!FechEmisi)
          FECH$ = TRIM$(FECHATEX$(FE%))
          
          JJ& = JJ& + 1
          Call REPLA(TX$, DOCUM$, 1, 18)
          Call REPLA(TX$, AJUSTI$(Str$(NC%), 6), 19, 6)
          Call REPLA(TX$, AJUSTI$(RACLI$, 30), 27, 30)
          Call REPLA(TX$, AJUSTD$(FECH$, 8), 56, 8)
          Call GRAREG("!LISCOMP", TX$, JJ&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!LISCOMP", JJ&)
    Screen.MousePointer = 1

    
10  Call SELECHO("!LISCOMP/Lista de Comprobantes de Venta")
    NUFACTU$ = TRIM$(VALACT1("!LISCOMP"))
    If NUFACTU$ = "" Then GoTo 99
    
    TXT$ = VALACT2("!LISCOMP")
    NN = InStr(1, TXT$, " ")
    NC% = Val(Mid(TXT$, 1, NN - 1))
    LETR$ = Mid$(NUFACTU$, 18, 1)
    
    Call NUEVOPRIFACTU(NUFACTU$, NC%, LETR$)
    
    GoTo 10
99  Command16.Enabled = True
End Sub

Sub Command22_Click()
   
   SQLX$ = "SELECT COUNT(*) AS CANTIDAD FROM CAJABANC  WHERE NUMECOMP < 0  AND TIPOMOVIM = 'FVEN'"
   Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   INDICE& = XVALO(CATMP!CANTIDAD)
   If INDICE& > 0 Then ReDim FACTU$(INDICE&)
   '
   i& = 0
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT IDCAE FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NUMECOMP < 0  AND TIPOMOVIM = 'FVEN' "
   SQLX$ = SQLX$ + " ORDER BY FEHOREAL "
   Set CAJATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CAJATMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           DOCU$ = SAFETEXT$(!IDCAE)
           If TRIM$(DOCU$) <> "" Then
            i& = i& + 1
            FACTU$(i&) = DOCU$
           End If
           .MoveNext
         Loop
      End If
      On Error GoTo 0
   End With
   CAJATMP.Close
   '
   Set CAJATMP = Nothing
   Timer2.Enabled = True
   ESTADO% = 0
   CONTAD& = 0
   
End Sub


Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    'ARMA LISTA DE SELECCION DE FACTURAS PENDIENTES DE CAE
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT DISTINCT IDCAE, FechEmisi,NuClien FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + " AND NUMECOMP < 0 AND IDCAE <> '' AND IDCAE IS NOT NULL"
    SQLX$ = SQLX$ + " ORDER BY FECHEMISI "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    Call SETULTREG("!FACTUSEL", 0)
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          TX$ = ""
          IDCAEX$ = SAFETEXT$(!IDCAE)
          NC% = XVALO(!NuClien)
          RACLI$ = Left$(RASOCLI$(NC%), 30)
          FE% = CVINT(!FechEmisi)
          FECH$ = TRIM$(FECHATEX$(FE%))
          '
          JJ& = JJ& + 1
          Call REPLA(TX$, IDCAEX$, 1, 18)
          Call REPLA(TX$, AJUSTI$(Str$(NC%), 6), 19, 6)
          Call REPLA(TX$, AJUSTI$(RACLI$, 30), 27, 30)
          Call REPLA(TX$, AJUSTD$(FECH$, 8), 56, 8)
          Call GRAREG("!FACTUSEL", TX$, JJ&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    CABATEMP.Close
    Set CABATEMP = Nothing
    Call SETULTREG("!FACTUSEL", JJ&)
    Screen.MousePointer = 1
    '
    Call FRESHECHO("!FACTUSEL")
    Call SELECHO("!FACTUSEL/Facturas Pendientes de Solicitud de Cae")
    IDCAEX$ = TRIM$(VALACT1("!FACTUSEL"))
    If IDCAEX$ <> "" Then
     Text5 = IDCAEX$
    End If
    '
    Command3.Enabled = True
End Sub

''Sub Command5_Click()
''     '
''      Command5.Enabled = False
''      '
''      MODO% = 1
''      PVTA% = XVALO(Text4)
''      IDENCAE$ = TRIM$(Text5)
''      If PVTA% > 0 And IDENCAE$ <> "" Then
''        Call PIDECAE(PVTA%, IDENCAE$)
''      Else
''        If MODOAUTO = "NO" Then Call COMUNI("Faltan Datos")
''      End If
''      '
''      Command5.Enabled = True
''      '
''End Sub

Function TIPO_DOC$(NC%)
    'TIPO DE DOCUMENTO
    'NC% = XVALO(FORFACPRINT.Text1)
    POSDEIVA% = PIVACLI%(NC%)
    Select Case POSDEIVA%
        Case 0: VALOR$ = 96
        Case 1: VALOR$ = 80
        Case 2: VALOR$ = 80
        Case 3: VALOR$ = 80
        Case 4: VALOR$ = 80
        Case 5: VALOR$ = 80
        Case 6: VALOR$ = 80
        Case 7: VALOR$ = 80
    End Select
    TIPO_DOC$ = VALOR$
End Function


Function NRO_DOC$(NC%)
    NRO_DOC$ = ""
    POXIVA% = PIVACLI%(NC%)
    '
    NUMERO_CUIT_DOC1$ = CUITCLI$(NC%)
    NUMERO_CUIT_DOC$ = REPLACAR(NUMERO_CUIT_DOC1$, "-", "")
    '
    If POXIVA% < 1 And TRIM$(NUMERO_CUIT_DOC$) = "" Then NUMERO_CUIT_DOC$ = "0"
    '
    If Not IsNumeric(NUMERO_CUIT_DOC$) Then
      If MODOAUTO = "NO" Then Call MENSERR(24, "Documento Nro.: " & NUMERO_CUIT_DOC$ & " No Válido\El Dato debe ser Numérico\Modifique y Vuelva a Cargar la Factura ")
      Call FINAL("")
    End If
    '
    NRO_DOC$ = NUMERO_CUIT_DOC$
    '
End Function


Function IMP_TOT_CONC$(PERCEPIVA#, PERCEPIB#, IMPNONETO#)

   IMP_TOT_CONC$ = "" ' SUMA DE IVA NO GRAVADO + PERCEPCION DE IVA + PERCEPCION DE INGRESOS BRUTOS
   IMPTOTACONC$ = FORMATNUM$(XVALO(IMPNONETO#) + XVALO(PERCEPIVA#) + XVALO(PERCEPIB#), "F12.2")
   If TRIM$(IMPTOTACONC$) = "" Then IMPTOTACONC$ = "0"
   IMP_TOT_CONC$ = TRIM$(IMPTOTACONC$)
   
End Function

Function IMPTO_LIQ$()
  IMPTO_LIQ$ = ""
  IMPTOLIQUID$ = FORMATNUM$(XVALO(FORFACPRINT.Total(3)) + XVALO(FORFACPRINT.Total(4)) + XVALO(FORFACPRINT.Total(5)), "F12.2")
  IMPTO_LIQ$ = TRIM$(IMPTOLIQUID$)
End Function

Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPO_CBTE$(PREFI$, LETRA$)
    
    TIPOCOMP$ = UCase$(PREFI$ & LETRA$)
    Select Case TIPOCOMP$
        '
        Case "FCA": VALOR$ = "1" 'FACTURA A
        Case "NDA": VALOR$ = "2" 'N DEBITO A
        Case "NCA": VALOR$ = "3" 'N CREDITO A
        Case "FCB": VALOR$ = "6" 'FACTURA B
        Case "NDB": VALOR$ = "7" 'N DEBITO B
        Case "NCB": VALOR$ = "8" 'N CREDITO B
    End Select
    '
    TIPO_CBTE$ = VALOR$
    '
End Function

'Private Sub Command6_Click()
'
'      SOLOCAE% = 1
'      '
'      Call Command5_Click
'
'End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub

Private Sub Text2_Change()

   If XVALO(Text1) > 0 Then
      PREFI$ = "0000" & TRIM$(Text4): PREFI$ = Right(PREFI$, 4)
      NUMEROFACTURA$ = "00000000" & TRIM$(Text2): NUMEROFACTURA$ = Right(NUMEROFACTURA$, 8)
      NUMFACTUNUE$ = Mid$(Text6, 1, 3) & PREFI$ + "-" + NUMEROFACTURA & Mid$(Text6, 17, 2)
      NUFACTUVIE$ = TRIM$(Text6)
      
      Call GRABANUEFACTU(NUFACTUVIE$, NUMFACTUNUE$)
   End If
End Sub


Private Sub Text2_gotfocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub



Private Sub Text5_Change()
    Call ABRECABAN
    
    DOCUM$ = TRIM$(Text5)
    'IDENDIFICADOR DE CAE
    SQLX$ = "SELECT CAE,CODICOM_LAR FROM CAJABANC  WHERE IDCAE = '" & DOCUM$ & "'"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    CAE$ = TRIM$(SAFETEXT$(CATMP!CAE))
    If CAE$ <> "" Then Text6 = "": Exit Sub
    Text6 = TRIM$(SAFETEXT$(CATMP!CODICOM_LAR))
    '
    'PUNTO DE VENTA
    TIP$ = REPLACAR$(Mid$(Text6, 4, 4), "X", "0")
    Text4 = XVALO(TIP$)
    
    'limpio el resto de los casilleros
    Text1 = ""
    Text2 = ""
    Text3 = ""

End Sub

Private Sub Text6_GotFocus()
On Error Resume Next
Text5.SetFocus
On Error GoTo 0
End Sub

Sub GRABANUEFACTU(NUFACTUVIE$, NUFACTUNUE$)
   Call ABRECABAN
   '
   TICOMX$ = Left(NUFACTUVIE$, 2)
   NUMECOMPROB$ = "000000" & XVALO(Text2): NUMECOMPROB$ = Right(NUMECOMPROB$, 6)
   LETRACOMP$ = Right(TRIM$(NUFACTUVIE$), 1)
   FACTUCOR$ = TICOMX$ + "-" + Right(NUMECOMPROB$, 6) + "-" + LETRACOMP$
   CAE$ = TRIM$(Text1)

   '
''   CONDI$ = "CODICOM_LAR = '" & NUFACTUVIE$ & "'"
   '
   SQLX$ = "SELECT NUCLIEN FROM CAJABANC  WHERE CODICOM_LAR = '" & NUFACTUVIE$ & "'"
   Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   NC% = XVALO(CATMP!NuClien)
   FF% = HOY(HO$)
   FECH = Int(CDbl(CVDAT(FF%)))
   FECHVENCICAE$ = TRIM$(Label6)
   COBA$ = ARMACODIGOBARRA(NUFACTUNUE$, CAE$, FECHVENCICAE$)
   
   'GRABO EN CAJABANC
   Call ABRECABAN
   NUMECOMPRO% = XVALO(NUMECOMPROB$)
   SQLX$ = "UPDATE CAJABANC SET CODICOM_COR = '" & FACTUCOR$ & "'"
   SQLX$ = SQLX$ + " ,NUMECOMP =" & NUMECOMPRO%
   SQLX$ = SQLX$ + " ,CAE ='" & CAE$ & "'"
   SQLX$ = SQLX$ + " ,FECHCONTA = " & FECH
   SQLX$ = SQLX$ + " ,FECHEMISI =" & FECH
   SQLX$ = SQLX$ + " ,CODICOM_LAR ='" & NUFACTUNUE$ & "'"
   SQLX$ = SQLX$ + " ,FechVtoCae ='" & FECHVENCICAE$ & "'"
   SQLX$ = SQLX$ + " ,COD_BARRA = '" & COBA$ & "'"
   SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & NUFACTUVIE$ & "'"
   CueCorri.Execute (SQLX$)
   '
   'GRABO EN SADEUPEN
   SQLX$ = "UPDATE SADEUPEN SET CODICOM_COR = '" & FACTUCOR$ & "'"
   SQLX$ = SQLX$ + " ,CODICOM_LAR ='" & NUFACTUNUE$ & "'"
   SQLX$ = SQLX$ + " ,FECHEMISI = " & FECH
   SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & NUFACTUVIE$ & "'"
   CueCorri.Execute (SQLX$)
   '
   'GRABO EN MOVISTO
   Call ABRESTOCKS
   SQLX$ = "UPDATE MOVISTO SET DOSECCOR = '" & FACTUCOR$ & "'"
   SQLX$ = SQLX$ + " ,DOSECLAR ='" & NUFACTUNUE$ & "'"
   SQLX$ = SQLX$ + " WHERE DOSECLAR = '" & NUFACTUVIE$ & "'"
   Stocks.Execute (SQLX$)
   '
   'GRABA EN ARCHIVO CARBIMON
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    FIN& = Len(RFF$)
    For KKU& = 1 To FIN& Step 4
       U& = CVS(Mid$(RFF$, KKU&, 4))
       XX$ = REGLEIDO$("CARBIMON", U&)
       If TRIM$(Mid$(XX$, 13, 18)) = NUFACTUVIE$ Then
         Call REPLA(XX$, NUFACTUNUE$, 13, 18)
         Call REPLA(XX$, NUFACTUNUE$, 45, 18)
         Call GRAREG("CARBIMON", XX$, U&)
      End If
    Next KKU&
   '
   'TOMA CANTIDAD DE COPIAS CONFIGURADAS EN SETUP
   '********************************************
   CANTCOPIAS% = XVALO(CONTROL$("", "COPIAFAC"))
   If CANTCOPIAS% < 1 Then CANTCOPIAS% = 1
   Call NUEVOPRIFACTU(NUFACTUNUE$, NC%, LETRACOMP$)
   '
End Sub
Function ARMACODIGOBARRA(NUFACTU$, CAE$, VENCIMIENTO$)
    
    ' Obtiene CUIT
    N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    Get #N1%, 6, X
    CUITT$ = TRIM$(REPLACAR$(TRIM$(DESENCRI$(X)), "-", ""))
    Close #N1%
    ' Tipo de comrpobante
    COMPROBANTE$ = TRIM$(TIPOX_CBTEX$(NUFACTU$))
    ' Punto de venta
    PuntoVenta$ = TRIM$(Mid$(NUFACTU$, 4, 4))
    ' CAE
    CODIGOAUTORIZACION$ = TRIM$(CAE$)
    ' Fecha de vencimiento
    FECHAVENCIMIENTO$ = TRIM$(VENCIMIENTO$)
    ' Arma cadena
    XX$ = CUITT$ + COMPROBANTE$ + PuntoVenta$ + CODIGOAUTORIZACION$ + FECHAVENCIMIENTO$
    ' Obtiene digito verificador
    DIGITOVERIFICADOR$ = OBTIENEDIGITO$(XX$)
    ' Arma cadena final
    XX$ = XX$ + DIGITOVERIFICADOR$
    
    YY$ = CODIBARRA$(XX$, 1) ' SE REEMPLAZA POR ESTA FUNCION
    '
    ARMACODIGOBARRA = YY$

End Function
Function TIPOX_CBTEX$(NUFACTU$)

    TIPOCOMP$ = Mid$(NUFACTU$, 1, 2) & ":" & Mid$(NUFACTU$, 18, 1)
    
    Select Case TIPOCOMP$
        Case "FC:A": VALOR$ = "01"
        Case "ND:A": VALOR$ = "02"
        Case "NC:A": VALOR$ = "03"
        Case "FC:B": VALOR$ = "06"
        Case "ND:B": VALOR$ = "07"
        Case "NC:B": VALOR$ = "08"
        Case "FC:E": VALOR$ = "19"
        Case "ND:E": VALOR$ = "20"
        Case "NC:E": VALOR$ = "21"
    End Select
    '
    TIPOX_CBTEX$ = VALOR$
    '
End Function

Private Sub Timer2_Timer()
   
   If CONTAD& >= INDICE& Then
       Timer2.Enabled = False
       If MODOAUTO = "NO" Then
            Call COMUNI("Proceso de Solicitud de CAE Finalizado")
       Else
            MODOAUTO = "NO" ' AVISA AL MODOAUTO="SI" QUE FINALIZO EL PROCESO
            Unload Me
       End If
   End If
   '
   If ESTADO% = 0 Then
      CONTAD& = CONTAD& + 1
      NUMEFAC$ = FACTU$(CONTAD&)
      Text5 = NUMEFAC$
      ESTADO% = 1
   End If
   '
End Sub

Sub SOLICITA_CAE_BCAP_ID(IDENCAE$)
    '
    Call ABRECABAN
    '
    SQLX$ = "SELECT TEXTOCAE, FECHEMISI FROM CAJABANC  WHERE IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    DATOX$ = SAFETEXT$(CATMP!TEXTOCAE)
    FE% = CVINT(CATMP!FechEmisi)
    If DATOX$ = "" Then Exit Sub

    'FECHA DE HOY
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%) ' FORMATO AAAAMMDD
    FECH2$ = ARMAFECHCAE(FE%) ' FORMATO AAAAMMDD
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$) 'REEMPLAZO LA FECHA QUE TENIA POR LA DE HOY
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    SQLX$ = "SELECT NUMETALO FROM CAJABANC  WHERE IDCAE = '" & IDENCAE$ & "'"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    PVTA$ = Str$(XVALO(CATMP!NUMETALO))

    SQLX$ = "SELECT OPCIMOVIM , VARIMOVIM FROM CAJABANC  WHERE IDCAE = '" & IDENCAE$ & "' AND TIPOMOVIM = 'FVEN'"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    PREFI$ = SAFETEXT$(CATMP!OPCIMOVIM)
    LETRA$ = SAFETEXT$(CATMP!VARIMOVIM)
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
      If EXISTE%(CARPETOK$ + "wsbfe_cliente.exe") = 1 Then
         TC = TIPO_CBTE(PREFI$, LETRA$)
         XXX = Shell(CARPETOK$ + "wsbfe_cliente.exe AUTO " & IDENCAE$ & " " & TC & " " & PVTA$, 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsbfe_cliente.exe")
         Exit Sub
      End If
    End If
    '
End Sub

Sub SOLICITA_CAE_EXPO_ID(IDENCAE$)
    '
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%)
    
    Call ABRECABAN
    SQLX$ = "SELECT TEXTOCAE, FECHEMISI FROM CAJABANC  WHERE IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    DATOX$ = SAFETEXT$(CATMP!TEXTOCAE)
    FE% = CVINT(CATMP!FechEmisi)
    
    FECH2$ = ARMAFECHCAE(FE%)
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$)
    '
    If DATOX$ = "" Then Exit Sub
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
      If EXISTE%(CARPETOK$ + "wsfex_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfex_cliente.exe AUTO " & IDENCAE$, 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfex_cliente.exe")
         Exit Sub
      End If
    End If
    '
End Sub

Sub SOLICITA_CAE_ID(IDENCAE$)
    '
    Call ABRECABAN
    '
    SQLX$ = "SELECT TEXTOCAE, FECHEMISI FROM CAJABANC  WHERE IDCAE = '" & IDENCAE$ & "' And TEXTOCAE Is Not Null"
    Set CATMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    DATOX$ = SAFETEXT$(CATMP!TEXTOCAE)
    FE% = CVINT(CATMP!FechEmisi)
    If DATOX$ = "" Then Exit Sub
    '
    'FECHA DE HOY
    HOII% = HOY(HOS$)
    FECH1$ = ARMAFECHCAE(HOII%) ' FORMATO AAAAMMDD

    FECH2$ = ARMAFECHCAE(FE%) ' FORMATO AAAAMMDD
    If FECH1$ <> FECH2$ Then DATOX$ = REPLACAR(DATOX$, FECH2$, FECH1$)  'REEMPLAZO LA FECHA QUE TENIA POR LA DE HOY

    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    Archiv$ = CARPETOK$ + "CONSULTA_CAE_" + IDENCAE$ + ".TXT"
    NARCHI2% = FILEOPEN%(Archiv$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, DATOX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT"  ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    If EXISTE%(CARPETOK$ + "CAE_" + IDENCAE$ + ".TXT") = 0 Then
      If EXISTE%(CARPETOK$ + "wsfe_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfe_cliente.exe AUTO " & IDENCAE$, 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
         Exit Sub
      End If
    End If
    '
End Sub
Function ARMAFECHCAE(FE%)

    If FE% < 1 Then Exit Function
    
    FECH = FECHATEXLAR$(FE%)
    FECHAFORM = Right(FECH, 4) + Mid$(FECH, 4, 2) + Mid$(FECH, 1, 2)
    ARMAFECHCAE = FECHAFORM
    '
End Function

Private Sub Timer3_Timer()
'''
''    Timer3.Enabled = False
''    Call Command2_Click
'''
End Sub


'
Sub NUEVOPRIFACTUNO(NUFACTU$, NCBNC%, LETR$)
    Static AGREDESP% '
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("CABAN")
    Call APERBASDAT("PEDCLI")
    '
    'SE OBTIENE EL NUMERO DE CLIENTE DEL MOVISTO POR QUE EN LOS CASOS DE CLIENTES CON
    'GRUPO DIVISIONAL SE GUARDA EN EL CAJABANC EL CLIENTE CONTENEDOR, EN EL MOVISTO ESTA EL ORIGINAL
    NCMOV% = XVALO(VALOBADAMDB("MOVISTO", "Nume_Clie", "DOSECLAR = '" & NUFACTU$ & "'", "NOMESS"))
    If NCMOV% < 1 Then NCMOV% = NCBNC%

    'AVERIGUO SI ES FACTURA POR TEXTO LIBRE
    CANTREG& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    HAYMATNOCODIFIC& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "'  AND COD_INTE = 0 AND Nume_Clie = " & NCMOV%, "NOMESS"))
    If CANTREG& <= HAYMATNOCODIFIC& Then TIPODECOMPROB$ = "(T.Libre)"
    '
    TIPOFACT$ = Left$(NUFACTU$, 2)
    CAHOJAS% = 0
    '
    'DEFINE EL TIPO DE FACTURA
    TIPODOC$ = "Factura": TIFADOC% = 2: CODOCUM$ = "FAC"
    If UCase$(Left$(NUFACTU$, 2)) = "ND" Then TIFADOC% = 3: TIPODOC$ = "Nota de Debito": CODOCUM$ = "DEB"
    If UCase$(Left$(NUFACTU$, 2)) = "NC" Then TIFADOC% = 6: TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    'DESTIDOC% = NCMOV%
    DESTIDOCNUE = NCMOV%
    '
    FORIPRE$ = FORELFACTU$(LETR$, TIPODECOMPROB$, TIPOFACT$, TIFADOC%)        ' BUSCA POR LA FUNCION ASOCIADA AL TIPO DE COMPROBANTE
    If FORIPRE$ = "NULL" Then Exit Sub
    CODFOR$ = FORIPRE$
    '
    Screen.MousePointer = 11
    '
    'SELECT QUE SE UTILIZA MAS ABAJO PARA TOMAR NUMEROS DE SERIES
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCMOV% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
    Set CONSITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)

    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CANCELPRINT% = 0
    'FFF% = HOY(HO$)
    FFF% = CVINT(VALOBADAMDB("CAJABANC", "FechConta", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    DESCU$ = SAFETEXT$(VALOBADAMDB("CAJABANC", "Descu1", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    PORCENDESCU = COEFDESCU(DESCU$)

    FECHDOC$ = FECHATEX$(FFF%)
    'FECHDOC$ = HO$
    NUMEDOC$ = Mid$(NUFACTU$, 4, 13) 'FORMATO'0001-00000123'
    '
    'RECUPERO LOS DATOS DE LA BASE
    MONEMI% = XVALO(VALOBADAMDB("CAJABANC", "MONE_EMIS", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If MONEMI% < 1 Then MONEMI% = 1
    MONE$ = DEMONECO$(MONEMI%)
    TIPOCAM# = XVALO(VALOBADAMDB("CAJABANC", "Tipo_Cam", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))


    IMPORTETOTALNETO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    IMPORTETOTALNETO# = IMPORTETOTALNETO# + XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))

    IVA1050# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA1050# = IVA1050# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# / TIPOCAM#
    IVA1050# = IVA1050# / TIPOCAM#
    COEFIVAINCLU# = 1
    If IMPORTETOTALNETO# > 0.005 Then
      COEFIVAINCLU# = (IMPORTETOTALNETO# + IVA1050# + IVA2100#) / IMPORTETOTALNETO#
    End If
    '
    'TOTALIVAINCLUIDO SE UTILIZA PARA CALCULAR EL DESCUENTO
    If TIFADOC% <> 6 Then
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
      'NOTA DE CREDITO
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# / TIPOCAM#
    If LETR$ = "B" Then
       TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# * COEFIVAINCLU#
       IMPORTETOTALNETO# = IMPORTETOTALNETO# * COEFIVAINCLU#
    End If
    '
    BASEIMPO21# = XVALO(VALOBADAMDB(" CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO21# = BASEIMPO21# / TIPOCAM#
    BASEIMPO105# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO105# = BASEIMPO105# / TIPOCAM#
    INOGRAVADO# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# + XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# / TIPOCAM#
    On Error Resume Next
    PORCENTAJEDESCUENTO = DESCU$
    On Error GoTo 0
    IMPORTEDESC# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
'''''    TOTALBRUTOSINIVACONDESCUENTO# = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN= " & NCBNC%, "NOMESS")) / TIPOCAM#
'''''    On Error Resume Next
'''''    PORCENTAJEDESCUENTO = ROUND((1 - (TOTALDOCUMENTO# / TOTALIVAINCLUIDO#)) * 100, 2)
'''''    On Error GoTo 0
'''''    TOTALBRUTOSINIVACONDESCUENTO# = TOTALBRUTOSINIVA# / (1 - (PORCENTAJEDESCUENTO / 100))
    If TIFADOC <> 6 Then
        PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
        PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
       PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
       PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    '
    DESCRILAR$ = VALOBADAMDB("MOVISTO", "DESCRI_LAR", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS")
    CAE$ = VALOBADAMDB("CAJABANC", "CAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    IDCAE$ = VALOBADAMDB("CAJABANC", "IDCAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    'SE AGREGA FECHA DE VENCIMIENTO DE CAE
    FECHVENCICAE$ = VALOBADAMDB("CAJABANC", "FechVtoCae", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")

    COBA$ = TRIM$(VALOBADAMDB("CAJABANC", "COD_BARRA", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If COBA$ = "" Then COBA$ = ARMACODIGOBARRA(NUFACTU$, CAE$, FECHVENCICAE$)

    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Observa$ = "CAE: " + CAE$ + " ... Vto Cae: " + FECHVENCICAE$ + " ... Id.Tran: " + IDCAE$ + vbCrLf + VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")
    If InStr(EPAC$, "MECANIZA") > 0 Or InStr(EPAC$, "EDR") > 0 Then Observa$ = VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")

    '
    'NUMERO DE PEDIDO PARA TOMAR DATOS DEL MISMO
    NROP& = XVALO(VALOBADAMDB("MOVISTO", "NuPedCli", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS"))
    NROPEDIDO$ = Str$(NROP&)
    If NROP& > 0 Then
       NUOCOMPRA$ = VALOBADAMDB("PEDENCA", "NroOcom", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
       VENDE% = XVALO(VALOBADAMDB("PEDENCA", "Vended", "NROPED = " & NROP&))
    Else
       'SI NO TIENE NUMERO DE PEDIDO PONE EL VENDEDOR DEL MAESTRO
       VENDE% = VENDCLI%(NCMOV%)
    End If
    NUVENDEX$ = Str$(VENDE%) + " " + ECOARCH$("VENDEDOR", Chr$(VENDE%))
    'CONDICION DE PAGO SI NO TIENE TOMA LA DEL MAESTRO
    CPAG% = XVALO(VALOBADAMDB("CAJABANC", "CondPag", "CodiCom_Lar = '" & NUFACTU$ & "' AND NuClien = " & NCBNC%, "NOMESS"))
    If CPAG% < 1 Then CPAG% = CPAGCLI%(NCMOV%)
    CONDPAG$ = ECOARCH$("CONPAG", Chr$(CPAG%))
    '
    NroRemi$ = VALOBADAMDB("MOVISTO", "DOPRILAR", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS")
    PORCENTAJEIVA1$ = "10.50 %"
    PORCENTAJEIVA2$ = "21.00 %"
    IEMBALA = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "CODICOM_LAR = '" & NUFACTU$ & "'AND TIPOMOVIM = 'EMBAL' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    '
    DOMICI$ = VALOBADAMDB("PEDENCA", "DomiEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    'SI NO TIENE DOMICILIO EL TRANSPORTE LE PONE EL TRANSPORTE DEL MAESTRO
    If TRIM$(DOMICI$) <> "" Then
      TRANSPO$ = VALOBADAMDB("PEDENCA", "TranspEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
      LOCALENT$ = VALOBADAMDB("PEDENCA", "LocaEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    Else
      DOMICI$ = DTRANCLI$(NCMOV%)
      TRANSPO$ = RTRANCLI$(NCMOV%)
      NT% = NUTRANCLI(NCMOV%)
      TLOCALENT$ = LOCATRANS$(NT%)
    End If
    '

    'PARAMETROS DE IMPRESION
    '***********************
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST" 'ESTOS YA ESTABAN COMENTADOS
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    CODTABU1$(27) = "N-AUTENT"
    CODTABU1$(28) = "R-PLANO"
    CODTABU1$(29) = "F-PLANO"
    '
    CACOLTAB1% = 29
    '
    CAITAB1% = 0
    CAITCRU% = 0

    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETR$
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(NroRemi$)
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = TRIM$(NUOCOMPRA$)
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = TRIM$(NROPEDIDO$)
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = TRIM$(NUVENDEX$)
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = TRIM$(CONDPAG$)
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORCENTAJEIVA2$
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORCENTAJEIVA1$ 'PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = PORCENTAJEDESCUENTO 'TRIM$(PORDESCU.TEXT) ' VER CALCULO DE PORCENTAJE DE DESCUENTO
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""    ' Valor en PESOS equivalente a dolares
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = TRIM$(DOMICI$) 'LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = TRIM$(LOCALENT$) 'LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = TRIM$(TRANSPO$) 'LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONE$
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
       Incoterms$ = VALOBADAMDB("CAJABANC", "Incoterms", "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCBNC%, "NOMESS")
    CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Incoterms$)          ' VALOR FOB ' ver como conseguir el val
       TIPOCAMM# = TIPOCAM#: If InStr(1, EMPREAC$, "GRAMY") > 0 And MONEMI% = 1 Then TIPOCAMM# = TICAMONE(2)
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = FORMATNUM$(TIPOCAMM#, "F12.4")              ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOCNUE)   '  CONTACTO
    CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(Str$(IEMBALA))                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NCMOV%)))
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2") ' VER ESTE TOTAL COMO SE CALCULA
    CODPARAM$(54) = "LAR-CUTE": DOCPARAM$(54) = CAE$
    CODPARAM$(55) = "LAR-MATE": DOCPARAM$(55) = FECHVENCICAE$
    CODPARAM$(56) = "LAR-CUCO": DOCPARAM$(56) = IDCAE$
    CODPARAM$(57) = "LAR-MACO": DOCPARAM$(57) = COBA$
    '
    'OBTENGO EL VARIDOC PARA TOMAR LA TASA CON ALIPERBRU
    '***************************************************
    VARIDOC% = 1: PIVADOC% = PIVACLI%(NCLIE%)
    If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
    ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
    End If
    TASAX = ALIPERBRU#(NCMOV%, MINIPER#, VARIDOC%)
    CODPARAM$(52) = "IMPO-CHE": DOCPARAM$(52) = FORMATNUM(TASAX, "F5.1")
    CODPARAM$(53) = "NUOPERS": DOCPARAM$(53) = GRUDIVCLI%(NCMOV%)
    '
    'IDCAE + CAE + OBSERVACIONES
    '****************************
    Call SETULTREG("!OBSERVOF", 0)
    If Observa$ <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(Observa$, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    CAPARDOC% = 57
    '
    REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      If (U& Mod ITPORHOJA%) = 1 Then
        ' PONE ENCABEZADO
        CAITAB1% = 0
        CAHOJAS% = CAHOJAS% + 1
      End If
      '
      U& = 0
      'LO LEO ACA POR QUE SE UTILIZA EL VALOR CUANDO RECORRE EL MOVISTO PARA IMPRIIMIR EL DETALLE
      CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC"))
      '
      'SE AGREGA PARA OBTENER EL VALOR MAXIMO DE ITEMS DE LA FACTURA QUE SE VALIDA MAS ABAJO
      CONDI$ = "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%
      FIN& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", CONDI$, "NOMESS"))

      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY NuorItem "
      Set MOVTMP = Stocks.OpenRecordset(SQLX$, 4)
      '
      With MOVTMP
       Do While .EOF = False
         U& = U& + 1
         If (U& Mod ITPORHOJA%) = 1 Then
            ' PONE ENCABEZADO
            CAITAB1% = 0
            CAHOJAS% = CAHOJAS% + 1
         End If

         CIPRE% = XVALO(!Cod_Inte)
         CANTID = Abs(XVALO(!CANTSALID) - XVALO(!CANTINGRE))
         CAPEX$ = Str$(CANTID)
         NPEDIDD& = XVALO(!NUPEDCLI)
         SUBCO$ = SUBCODI$(NPEDIDD&, CIPRE%)
         ORDENCOMPRA$ = SAFETEXT$(NumeOrdPed)
         NroRemi$ = SAFETEXT$(!DOPRILAR)
         VALORUNITARIO# = XVALO(!VALOUNITIVA)
         'AQUI LE CARGA EL PRECIO ORIGINAL SIN EL DESCUENTO APLICADO
         'EJ: SI EL MOVISTO DEVUELVE VALORUNITARIO 80 Y TENIA UN DESCUENTO APLICADO DE 20% LA FUNCION VA A DEVOLVER VALORUNITARIO=100
         VALORUNITARIO# = PrecioOriginalPreDescuento(VALORUNITARIO#, PORCENDESCU)

           If LETR$ = "B" Then
              VALORUNITARIO# = VALORUNITARIO# / (1 + PORCEIVA(CIPRE%) / 100)
           End If
         CAITAB1% = CAITAB1% + 1
         CODDX$ = CODEXT$(CIPRE%) + SUBCO$ ' en coddx$ trae el codigo externo a partir de cipre(int)
         If REPLAORI% = 1 Then
           CODDX$ = CODORIG$(CIPRE%, NCMOV%) ' + SUBCO$
           'se pasa el codigo interno, mas el nº cliente(NCMOV%) y trae el codigo externo solicitado
        End If
        TETABU1$(1, CAITAB1%) = CODDX$
        DESCRILAR$ = SAFETEXT$(!DESCRI_LAR)
        If CIPRE% > 0 Then
         TEXTO$ = DESCRIT0$(CIPRE%)
        Else
         TEXTO$ = DESCRILAR$
        End If
        TETABU1$(2, CAITAB1%) = TEXTO$
        TETABU1$(3, CAITAB1%) = CAPEX$
        'TETABU1$(4, CAITAB1%) = PRELI$
        'TETABU1$(5, CAITAB1%) = DESCX$
        'TETABU1$(6, CAITAB1%) = PREUX$
        'TETABU1$(7, CAITAB1%) = ITOTX$
        'TETABU1$(8, CAITAB1%) = FENTX$
        TETABU1$(9, CAITAB1%) = Str$(U&)
        TETABU1$(10, CAITAB1%) = UNIMED$(CIPRE%)    ' unidad de medida
       TETABU1$(14, CAITAB1%) = NUPLANO$(CIPRE%) ' numero de plano
       TETABU1$(16, CAITAB1%) = ORDENCOMPRA$  ' numero de o.compra
       TETABU1$(17, CAITAB1%) = NroRemi$  ' numero de remito
      '
      FACIVITEM = 1
      If LETR$ = "B" Then FACIVITEM = 1 + PORCEIVA(CIPRE%) / 100
      '
      VALORUNITARIO# = VALORUNITARIO# * FACIVITEM
      TETABU1$(22, CAITAB1%) = FORMATNUM$(VALORUNITARIO#, "F15.4")  ' precio unitario
      TETABU1$(23, CAITAB1%) = TRIM$(FORMATNUM$(CANTID * VALORUNITARIO#, "F15.4"))   ' total neto
         If TETABU1$(23, CAITAB1%) = "" Then TETABU1$(23, CAITAB1%) = "S/Cargo"
      TETABU1$(24, CAITAB1%) = FORMATNUM$((PORCEIVA(CIPRE%)), "F5.2") ' alicuota iva
      TETABU1$(25, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#) - (VALORUNITARIO#), "F15.4") ' importe iva del item
      TETABU1$(26, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#), "F15.4") ' total bruto
         If TETABU1$(26, CAITAB1%) = "" Then TETABU1$(26, CAITAB1%) = "S/Cargo"
      TETABU1$(27, CAITAB1%) = Mid$(XXX$, 371, 16) ' numero de pedido cliente
      ' AQUI BUSCA EL NUMERO DE SERIE
        '
        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        '
        With CONSITEMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
            SERNU$ = !ReferCor
            If Left$(SERNU$, 5) = "S.No." Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = SERNU$
            End If
          End If
        End With

      '
      ' AQUI AGREGA DESPACHO DE IMPORTACION 'VER
      If AGREDESP% = 0 Then
        AGREDESP% = (-1)
        If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      '
'SE COMENTO POR QUE ESTA TRAYENDO LA ESPECIFICACION SIN CONFIGURACION Y EN TODO CASO SE HARA POR CONFIGURACION
'      ESPECIFIX$ = ESPECIFIT$(CIPRE%) ' TOMA LA ESPECIFICACION DEL ARTICUL
'      PARAMEXT$ = CODTABU1$(2) ' NOMBRE DEL PARAMETRO
'      Call PRINTEXT(ESPECIFIX$, 2, PARAMEXT$)
      '
      'INTERVENCION DEL CIELO
      NROPEDI& = XVALO(!NUPEDCLI)
      CONDI$ = "NROPED = " & NROPEDI&
      REFERPEDI$ = TRIM$(VALOBADAMDB("PEDENCA", "Referencia", CONDI$, "NOMESS"))
      'If REFERPEDI$ <> "" And InStr(EPAC$, "DEL CIELO") > 0 Then
      If REFERPEDI$ <> "" And CONTROL("FACTURA", "SISTMANT") = "SI" Then
         'TEXTO EJ REFERENCIA: LIMPIEZA DE TANQUE AGOSTO 2015 (1/6)
        POS1% = InStr(REFERPEDI$, "|") ' ME QUEDO SOLAMENTE CON EL MES Y EL NRO DE CUOTA AGOSTO 2015 (1/6)
        If POS1% > 0 Then REFERPEDI$ = TRIM$(Mid$(REFERPEDI$, POS1% + 1))
        Call PRINTEXT(REFERPEDI$, 2, PARAMEXT$)
      End If
      'FIN INTERVENCION DEL CIELO
      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo

      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            Call PRINTDOC("@" + FORIPRE$)
            Call CIERRARCH("*.*")
            Call BAJALDISCO
          End If
        End If
      End If

       .MoveNext
       Loop
     End With
'''''''''''    '
    TOTALIVAINCLUIDO# = PrecioOriginalPreDescuento(TOTALIVAINCLUIDO#, PORCENDESCU)
    IMPORTETOTALNETO# = PrecioOriginalPreDescuento(VALORUNITARIO#, PORCENDESCU)
    IMPORTEDESC# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(TOTALIVAINCLUIDO# + IEMBALA, "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = IVA2100#
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = IVA1050#
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = FORMATNUM$(Abs(IMPORTEDESC#), "F15.4") 'Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = IMPORTETOTALNETO# - Abs(IMPORTEDESC#)
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = TOTALDOCUMENTO#
     VADOLA# = TOTALDOCUMENTO# * TIPOCAM#
     DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = PERCEPBRUTO#
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = PERCEPIVA#                         'PARA IMPRIMIR PERC-IVA
    ITOGRAVA# = TOTALBRUTOSINIVA#
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
    INOGRAVA# = INOGRAVADO#
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
    DOCPARAM$(47) = BASEIMPO21#
    DOCPARAM$(48) = BASEIMPO105#

'''''    If LETR$ = "B" Then
'''''      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * TOTALBRUTOSINIVA# + IEMBALA, "F13.2")
'''''      DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * PORCENTAJEDESCUENTO, "F13.2")
'''''      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2")
'''''      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
'''''      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2"))
'''''      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLU# * INOGRAVA#, "F13.2"))
'''''      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA1050#, "F13.2"))
'''''      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA2100#, "F13.2"))
'''''    End If
 '''''''''''    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TOTFAC$ = NUMTEX$(TOTALDOCUMENTO#)
    If MONE$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS") ' MODIFICADO POR QUE SIEMPRE IMPRIMIA PESOS
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONE$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
'''''''''''
    PRINTERPORT$ = PORTUFAC$
'    CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC")) LO LEO ARRIBA
    If CACOPDOC% > 0 Then
       Call PRINTDOC("@" + FORIPRE$)
    End If
'''''''''''    '
'    NUCON& = XVALO(CONTROL("ALATALO", FORIPRE$))
'    NUPANT& = XVALO(Text11)
'    If NUCON& > 0 Then
'       If NUCON& < NUPANT& Then
'          Call COMUNI("ATENCION: Alarma de Talonario Proximo a Terminarse.")
'       End If
'    End If
'''''''''''    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub
'
Sub NUEVOPRIFACTU(NUFACTU$, NCBNC%, LETR$)
    Static AGREDESP% '
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("CABAN")
    Call APERBASDAT("PEDCLI")
    '
    'SE OBTIENE EL NUMERO DE CLIENTE DEL MOVISTO POR QUE EN LOS CASOS DE CLIENTES CON
    'GRUPO DIVISIONAL SE GUARDA EN EL CAJABANC EL CLIENTE CONTENEDOR, EN EL MOVISTO ESTA EL ORIGINAL
    NCMOV% = XVALO(VALOBADAMDB("MOVISTO", "Nume_Clie", "DOSECLAR = '" & NUFACTU$ & "'", "NOMESS"))
    If NCMOV% < 1 Then NCMOV% = NCBNC%

    'AVERIGUO SI ES FACTURA POR TEXTO LIBRE
    CANTREG& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    HAYMATNOCODIFIC& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "'  AND COD_INTE = 0 AND Nume_Clie = " & NCMOV%, "NOMESS"))
    If CANTREG& <= HAYMATNOCODIFIC& Then TIPODECOMPROB$ = "(T.Libre)"
    '
    TIPOFACT$ = Left$(NUFACTU$, 2)
    CAHOJAS% = 0
    '
    'DEFINE EL TIPO DE FACTURA
    TIPODOC$ = "Factura": TIFADOC% = 2: CODOCUM$ = "FAC"
    If UCase$(Left$(NUFACTU$, 2)) = "ND" Then TIFADOC% = 3: TIPODOC$ = "Nota de Debito": CODOCUM$ = "DEB"
    If UCase$(Left$(NUFACTU$, 2)) = "NC" Then TIFADOC% = 6: TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    'DESTIDOC% = NCMOV%
    DESTIDOCNUE = NCMOV%
    '
    FORIPRE$ = FORELFACTU$(LETR$, TIPODECOMPROB$, TIPOFACT$, TIFADOC%)        ' BUSCA POR LA FUNCION ASOCIADA AL TIPO DE COMPROBANTE
    If FORIPRE$ = "NULL" Then Exit Sub
    CODFOR$ = FORIPRE$
    '
    Screen.MousePointer = 11
    '
    'SELECT QUE SE UTILIZA MAS ABAJO PARA TOMAR NUMEROS DE SERIES
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCMOV% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
    Set CONSITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
 
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CANCELPRINT% = 0
    'FFF% = HOY(HO$)
    FFF% = CVINT(VALOBADAMDB("CAJABANC", "FechConta", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    DESCU$ = SAFETEXT$(VALOBADAMDB("CAJABANC", "Descu1", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    PORCENDESCU = COEFDESCU(DESCU$)
    
    FECHDOC$ = FECHATEX$(FFF%)
    'FECHDOC$ = HO$
    NUMEDOC$ = Mid$(NUFACTU$, 4, 13) 'FORMATO'0001-00000123'
    '
    'RECUPERO LOS DATOS DE LA BASE
    MONEMI% = XVALO(VALOBADAMDB("CAJABANC", "MONE_EMIS", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If MONEMI% < 1 Then MONEMI% = 1
    MONE$ = DEMONECO$(MONEMI%)
    TIPOCAM# = XVALO(VALOBADAMDB("CAJABANC", "Tipo_Cam", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
     
     
    IMPORTETOTALNETO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    IMPORTETOTALNETO# = IMPORTETOTALNETO# + XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    
    IVA1050# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA1050# = IVA1050# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# / TIPOCAM#
    IVA1050# = IVA1050# / TIPOCAM#
    COEFIVAINCLU# = 1
    If IMPORTETOTALNETO# > 0.005 Then
      COEFIVAINCLU# = (IMPORTETOTALNETO# + IVA1050# + IVA2100#) / IMPORTETOTALNETO#
    End If
    '
    'TOTALIVAINCLUIDO SE UTILIZA PARA CALCULAR EL DESCUENTO
    If TIFADOC% <> 6 Then
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
      'NOTA DE CREDITO
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# / TIPOCAM#
    If LETR$ = "B" Then
       TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# * COEFIVAINCLU#
       IMPORTETOTALNETO# = IMPORTETOTALNETO# * COEFIVAINCLU#
    End If
    '
    BASEIMPO21# = XVALO(VALOBADAMDB(" CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO21# = BASEIMPO21# / TIPOCAM#
    BASEIMPO105# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO105# = BASEIMPO105# / TIPOCAM#
    INOGRAVADO# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# + XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# / TIPOCAM#
    On Error Resume Next
    PORCENTAJEDESCUENTO = 100 * (1 - TOTALBRUTOSINIVA# / TOTALIVAINCLUIDO#)
    On Error GoTo 0
    IMPORTEDESC# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
'''''    TOTALBRUTOSINIVACONDESCUENTO# = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN= " & NCBNC%, "NOMESS")) / TIPOCAM#
'''''    On Error Resume Next
'''''    PORCENTAJEDESCUENTO = ROUND((1 - (TOTALDOCUMENTO# / TOTALIVAINCLUIDO#)) * 100, 2)
'''''    On Error GoTo 0
'''''    TOTALBRUTOSINIVACONDESCUENTO# = TOTALBRUTOSINIVA# / (1 - (PORCENTAJEDESCUENTO / 100))
    If TIFADOC <> 6 Then
        PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
        PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
       PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
       PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    '
    DESCRILAR$ = VALOBADAMDB("MOVISTO", "DESCRI_LAR", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS")
    CAE$ = VALOBADAMDB("CAJABANC", "CAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    IDCAE$ = VALOBADAMDB("CAJABANC", "IDCAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    'SE AGREGA FECHA DE VENCIMIENTO DE CAE
    FECHVENCICAE$ = VALOBADAMDB("CAJABANC", "FechVtoCae", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    
    COBA$ = TRIM$(VALOBADAMDB("CAJABANC", "COD_BARRA", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If COBA$ = "" Then COBA$ = ARMACODIGOBARRA(NUFACTU$, CAE$, FECHVENCICAE$)
    
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Observa$ = "CAE: " + CAE$ + " ... Vto Cae: " + FECHVENCICAE$ + " ... Id.Tran: " + IDCAE$ + vbCrLf + VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")
    If InStr(EPAC$, "MECANIZA") > 0 Or InStr(EPAC$, "EDR") > 0 Then Observa$ = VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")
       
    '
    'NUMERO DE PEDIDO PARA TOMAR DATOS DEL MISMO
    NROP& = XVALO(VALOBADAMDB("MOVISTO", "NuPedCli", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS"))
    NROPEDIDO$ = Str$(NROP&)
    If NROP& > 0 Then
       NUOCOMPRA$ = VALOBADAMDB("PEDENCA", "NroOcom", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
       VENDE% = XVALO(VALOBADAMDB("PEDENCA", "Vended", "NROPED = " & NROP&))
    Else
       'SI NO TIENE NUMERO DE PEDIDO PONE EL VENDEDOR DEL MAESTRO
       VENDE% = VENDCLI%(NCMOV%)
    End If
    NUVENDEX$ = Str$(VENDE%) + " " + ECOARCH$("VENDEDOR", Chr$(VENDE%))
    'CONDICION DE PAGO SI NO TIENE TOMA LA DEL MAESTRO
    CPAG% = XVALO(VALOBADAMDB("CAJABANC", "CondPag", "CodiCom_Lar = '" & NUFACTU$ & "' AND NuClien = " & NCBNC%, "NOMESS"))
    If CPAG% < 1 Then CPAG% = CPAGCLI%(NCMOV%)
    CONDPAG$ = ECOARCH$("CONPAG", Chr$(CPAG%))
    '
    NroRemi$ = VALOBADAMDB("MOVISTO", "DOPRILAR", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS")
    PORCENTAJEIVA1$ = "10.50 %"
    PORCENTAJEIVA2$ = "21.00 %"
    IEMBALA = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "CODICOM_LAR = '" & NUFACTU$ & "'AND TIPOMOVIM = 'EMBAL' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    '
    DOMICI$ = VALOBADAMDB("PEDENCA", "DomiEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    'SI NO TIENE DOMICILIO EL TRANSPORTE LE PONE EL TRANSPORTE DEL MAESTRO
    If TRIM$(DOMICI$) <> "" Then
      TRANSPO$ = VALOBADAMDB("PEDENCA", "TranspEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
      LOCALENT$ = VALOBADAMDB("PEDENCA", "LocaEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    Else
      DOMICI$ = DTRANCLI$(NCMOV%)
      TRANSPO$ = RTRANCLI$(NCMOV%)
      NT% = NUTRANCLI(NCMOV%)
      TLOCALENT$ = LOCATRANS$(NT%)
    End If
    '
    
    'PARAMETROS DE IMPRESION
    '***********************
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST" 'ESTOS YA ESTABAN COMENTADOS
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    CODTABU1$(27) = "N-AUTENT"
    CODTABU1$(28) = "R-PLANO"
    CODTABU1$(29) = "F-PLANO"
    '
    CACOLTAB1% = 29
    '
    CAITAB1% = 0
    CAITCRU% = 0

    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETR$
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(NroRemi$)
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = TRIM$(NUOCOMPRA$)
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = TRIM$(NROPEDIDO$)
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = TRIM$(NUVENDEX$)
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = TRIM$(CONDPAG$)
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORCENTAJEIVA2$
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORCENTAJEIVA1$ 'PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = DESCU$ 'FORMATNUM$(PORCENTAJEDESCUENTO, "F13.2") 'TRIM$(PORDESCU.TEXT) ' VER CALCULO DE PORCENTAJE DE DESCUENTO
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""    ' Valor en PESOS equivalente a dolares
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = TRIM$(DOMICI$) 'LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = TRIM$(LOCALENT$) 'LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = TRIM$(TRANSPO$) 'LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONE$
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
       Incoterms$ = VALOBADAMDB("CAJABANC", "Incoterms", "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCBNC%, "NOMESS")
    CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Incoterms$)          ' VALOR FOB ' ver como conseguir el val
       TIPOCAMM# = TIPOCAM#: If InStr(1, EMPREAC$, "GRAMY") > 0 And MONEMI% = 1 Then TIPOCAMM# = TICAMONE(2)
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = FORMATNUM$(TIPOCAMM#, "F12.4")              ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOCNUE)   '  CONTACTO
    CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(Str$(IEMBALA))                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NCMOV%)))
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2") ' VER ESTE TOTAL COMO SE CALCULA
    CODPARAM$(54) = "LAR-CUTE": DOCPARAM$(54) = CAE$
    CODPARAM$(55) = "LAR-MATE": DOCPARAM$(55) = FECHVENCICAE$
    CODPARAM$(56) = "LAR-CUCO": DOCPARAM$(56) = IDCAE$
    CODPARAM$(57) = "LAR-MACO": DOCPARAM$(57) = COBA$
    '
    'OBTENGO EL VARIDOC PARA TOMAR LA TASA CON ALIPERBRU
    '***************************************************
    VARIDOC% = 1: PIVADOC% = PIVACLI%(NCLIE%)
    If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
    ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
    End If
    TASAX = ALIPERBRU#(NCMOV%, MINIPER#, VARIDOC%)
    CODPARAM$(52) = "IMPO-CHE": DOCPARAM$(52) = FORMATNUM(TASAX, "F5.1")
    CODPARAM$(53) = "NUOPERS": DOCPARAM$(53) = GRUDIVCLI%(NCMOV%)
    '
    'IDCAE + CAE + OBSERVACIONES
    '****************************
    Call SETULTREG("!OBSERVOF", 0)
    If Observa$ <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(Observa$, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    CAPARDOC% = 57
    '
    REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      If (U& Mod ITPORHOJA%) = 1 Then
        ' PONE ENCABEZADO
        CAITAB1% = 0
        CAHOJAS% = CAHOJAS% + 1
      End If
      '
      U& = 0
      'LO LEO ACA POR QUE SE UTILIZA EL VALOR CUANDO RECORRE EL MOVISTO PARA IMPRIIMIR EL DETALLE
      CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC"))
      '
      'SE AGREGA PARA OBTENER EL VALOR MAXIMO DE ITEMS DE LA FACTURA QUE SE VALIDA MAS ABAJO
      CONDI$ = "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%
      FIN& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", CONDI$, "NOMESS"))
      
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY NuorItem "
      Set MOVTMP = Stocks.OpenRecordset(SQLX$, 4)
      '
      With MOVTMP
       Do While .EOF = False
         U& = U& + 1
         If (U& Mod ITPORHOJA%) = 1 Then
            ' PONE ENCABEZADO
            CAITAB1% = 0
            CAHOJAS% = CAHOJAS% + 1
         End If
 
         CIPRE% = XVALO(!Cod_Inte)
         CANTID = Abs(XVALO(!CANTSALID) - XVALO(!CANTINGRE))
         CAPEX$ = Str$(CANTID)
         NPEDIDD& = XVALO(!NUPEDCLI)
         SUBCO$ = SUBCODI$(NPEDIDD&, CIPRE%)
         ORDENCOMPRA$ = SAFETEXT$(NumeOrdPed)
         NroRemi$ = SAFETEXT$(!DOPRILAR)
         VALORUNITARIO# = XVALO(!VALOUNITIVA)
           If LETR$ = "B" Then
              VALORUNITARIO# = VALORUNITARIO# / (1 + PORCEIVA(CIPRE%) / 100)
           End If
         CAITAB1% = CAITAB1% + 1
         CODDX$ = CODEXT$(CIPRE%) + SUBCO$ ' en coddx$ trae el codigo externo a partir de cipre(int)
         If REPLAORI% = 1 Then
           CODDX$ = CODORIG$(CIPRE%, NCMOV%) ' + SUBCO$
           'se pasa el codigo interno, mas el nº cliente(NCMOV%) y trae el codigo externo solicitado
        End If
        TETABU1$(1, CAITAB1%) = CODDX$
        DESCRILAR$ = SAFETEXT$(!DESCRI_LAR)
        If CIPRE% > 0 Then
         TEXTO$ = DESCRIT0$(CIPRE%)
        Else
         TEXTO$ = DESCRILAR$
        End If
        TETABU1$(2, CAITAB1%) = TEXTO$
        TETABU1$(3, CAITAB1%) = CAPEX$
        'TETABU1$(4, CAITAB1%) = PRELI$
        'TETABU1$(5, CAITAB1%) = DESCX$
        'TETABU1$(6, CAITAB1%) = PREUX$
        'TETABU1$(7, CAITAB1%) = ITOTX$
        'TETABU1$(8, CAITAB1%) = FENTX$
        TETABU1$(9, CAITAB1%) = Str$(U&)
        TETABU1$(10, CAITAB1%) = UNIMED$(CIPRE%)    ' unidad de medida
       TETABU1$(14, CAITAB1%) = NUPLANO$(CIPRE%) ' numero de plano
       TETABU1$(16, CAITAB1%) = ORDENCOMPRA$  ' numero de o.compra
       TETABU1$(17, CAITAB1%) = NroRemi$  ' numero de remito
      '
      FACIVITEM = 1
      If LETR$ = "B" Then FACIVITEM = 1 + PORCEIVA(CIPRE%) / 100
      '
      DIFIERE# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
      If Abs(DIFIERE#) < 0.5 Then 'SE CONSIDERA PARA QUE CALCULE EL PRECIO ANTES DEL DESCUENTO CUANDO HAY DIFRENCIA
                                  'ENTRE EL CAJABANC Y EL MOVISTO.
        VALORUNITARIO# = PrecioOriginalPreDescuento(VALORUNITARIO#, PORCENDESCU)
      End If
      VALORUNITARIO# = VALORUNITARIO# * FACIVITEM
      TETABU1$(22, CAITAB1%) = FORMATNUM$(VALORUNITARIO#, "F15.4")  ' precio unitario
      TETABU1$(23, CAITAB1%) = TRIM$(FORMATNUM$(CANTID * VALORUNITARIO#, "F15.4"))   ' total neto
         If TETABU1$(23, CAITAB1%) = "" Then TETABU1$(23, CAITAB1%) = "S/Cargo"
      TETABU1$(24, CAITAB1%) = FORMATNUM$((PORCEIVA(CIPRE%)), "F5.2") ' alicuota iva
      TETABU1$(25, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#) - (VALORUNITARIO#), "F15.4") ' importe iva del item
      TETABU1$(26, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#), "F15.4") ' total bruto
         If TETABU1$(26, CAITAB1%) = "" Then TETABU1$(26, CAITAB1%) = "S/Cargo"
      TETABU1$(27, CAITAB1%) = Mid$(XXX$, 371, 16) ' numero de pedido cliente
      ' AQUI BUSCA EL NUMERO DE SERIE
        '
        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        '
        With CONSITEMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
            SERNU$ = !ReferCor
            If Left$(SERNU$, 5) = "S.No." Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = SERNU$
            End If
          End If
        End With
      
      '
      ' AQUI AGREGA DESPACHO DE IMPORTACION 'VER
      If AGREDESP% = 0 Then
        AGREDESP% = (-1)
        If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      '
'SE COMENTO POR QUE ESTA TRAYENDO LA ESPECIFICACION SIN CONFIGURACION Y EN TODO CASO SE HARA POR CONFIGURACION
'      ESPECIFIX$ = ESPECIFIT$(CIPRE%) ' TOMA LA ESPECIFICACION DEL ARTICUL
'      PARAMEXT$ = CODTABU1$(2) ' NOMBRE DEL PARAMETRO
'      Call PRINTEXT(ESPECIFIX$, 2, PARAMEXT$)
      '
      'INTERVENCION DEL CIELO
      NROPEDI& = XVALO(!NUPEDCLI)
      CONDI$ = "NROPED = " & NROPEDI&
      REFERPEDI$ = TRIM$(VALOBADAMDB("PEDENCA", "Referencia", CONDI$, "NOMESS"))
      If REFERPEDI$ <> "" And InStr(EPAC$, "DEL CIELO") > 0 Then
         'TEXTO EJ REFERENCIA: LIMPIEZA DE TANQUE AGOSTO 2015 (1/6)
        POS1% = InStr(REFERPEDI$, "|") ' ME QUEDO SOLAMENTE CON EL MES Y EL NRO DE CUOTA AGOSTO 2015 (1/6)
        If POS1% > 0 Then REFERPEDI$ = TRIM$(Mid$(REFERPEDI$, POS1% + 1))
        Call PRINTEXT(REFERPEDI$, 2, PARAMEXT$)
      End If
      'FIN INTERVENCION DEL CIELO
      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
 
      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            Call PRINTDOC("@" + FORIPRE$)
            Call CIERRARCH("*.*")
            Call BAJALDISCO
          End If
        End If
      End If
 
       .MoveNext
       Loop
     End With
'''''''''''    '
    If Abs(DIFIERE#) < 0.5 Then 'SE CONSIDERA PARA QUE CALCULE EL PRECIO ANTES DEL DESCUENTO CUANDO HAY DIFRENCIA
                                'ENTRE EL CAJABANC Y EL MOVISTO.
       TOTALIVAINCLUIDO# = PrecioOriginalPreDescuento(TOTALIVAINCLUIDO#, PORCENDESCU)
    End If
    IMPORTEDESC# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
    If Abs(IMPORTEDESC#) > 0.05 And TRIM$(DESCU$) = "" Then
       CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = ""
       DescuPorc = Abs(PORCENT_ENTRE_2_VALORES(IMPORTETOTALNETO#, TOTALIVAINCLUIDO#))
       CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = FORMATNUM$(DescuPorc, "F8.2")
    End If
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(TOTALIVAINCLUIDO# + IEMBALA, "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = IVA2100#
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = IVA1050#
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = FORMATNUM$(IMPORTEDESC#, "F15.4") 'Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = IMPORTETOTALNETO#
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = TOTALDOCUMENTO#
     VADOLA# = TOTALDOCUMENTO# * TIPOCAM#
     DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = PERCEPBRUTO#
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = PERCEPIVA#                         'PARA IMPRIMIR PERC-IVA
    ITOGRAVA# = TOTALBRUTOSINIVA#
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
    INOGRAVA# = INOGRAVADO#
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
    DOCPARAM$(47) = BASEIMPO21#
    DOCPARAM$(48) = BASEIMPO105#
 
'''''    If LETR$ = "B" Then
'''''      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * TOTALBRUTOSINIVA# + IEMBALA, "F13.2")
'''''      DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * PORCENTAJEDESCUENTO, "F13.2")
'''''      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2")
'''''      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
'''''      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2"))
'''''      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLU# * INOGRAVA#, "F13.2"))
'''''      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA1050#, "F13.2"))
'''''      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA2100#, "F13.2"))
'''''    End If
 '''''''''''    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TOTFAC$ = NUMTEX$(TOTALDOCUMENTO#)
    If MONE$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS") ' MODIFICADO POR QUE SIEMPRE IMPRIMIA PESOS
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONE$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
'''''''''''
    PRINTERPORT$ = PORTUFAC$
'    CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC")) LO LEO ARRIBA
    If CACOPDOC% > 0 Then
       Call PRINTDOC("@" + FORIPRE$)
    End If
'''''''''''    '
'    NUCON& = XVALO(CONTROL("ALATALO", FORIPRE$))
'    NUPANT& = XVALO(Text11)
'    If NUCON& > 0 Then
'       If NUCON& < NUPANT& Then
'          Call COMUNI("ATENCION: Alarma de Talonario Proximo a Terminarse.")
'       End If
'    End If
'''''''''''    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub NUEVOPRIFACTU_RESPALDO(NUFACTU$, NCBNC%, LETR$)
    Static AGREDESP% '
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC")):   If Len(CMO$) <> 2 Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("CABAN")
    Call APERBASDAT("PEDCLI")
    '
    'SE OBTIENE EL NUMERO DE CLIENTE DEL MOVISTO POR QUE EN LOS CASOS DE CLIENTES CON
    'GRUPO DIVISIONAL SE GUARDA EN EL CAJABANC EL CLIENTE CONTENEDOR, EN EL MOVISTO ESTA EL ORIGINAL
    NCMOV% = XVALO(VALOBADAMDB("MOVISTO", "Nume_Clie", "DOSECLAR = '" & NUFACTU$ & "'", "NOMESS"))
    If NCMOV% < 1 Then NCMOV% = NCBNC%

    'AVERIGUO SI ES FACTURA POR TEXTO LIBRE
    CANTREG& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    HAYMATNOCODIFIC& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", "DOSECLAR = '" & NUFACTU$ & "'  AND COD_INTE = 0 AND Nume_Clie = " & NCMOV%, "NOMESS"))
    If CANTREG& <= HAYMATNOCODIFIC& Then TIPODECOMPROB$ = "(T.Libre)"
    '
    TIPOFACT$ = Left$(NUFACTU$, 2)
    CAHOJAS% = 0
    '
    'DEFINE EL TIPO DE FACTURA
    TIPODOC$ = "Factura": TIFADOC% = 2: CODOCUM$ = "FAC"
    If UCase$(Left$(NUFACTU$, 2)) = "ND" Then TIFADOC% = 3: TIPODOC$ = "Nota de Debito": CODOCUM$ = "DEB"
    If UCase$(Left$(NUFACTU$, 2)) = "NC" Then TIFADOC% = 6: TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    'DESTIDOC% = NCMOV%
    DESTIDOCNUE = NCMOV%
    '
    FORIPRE$ = FORELFACTU$(LETR$, TIPODECOMPROB$, TIPOFACT$, TIFADOC%)        ' BUSCA POR LA FUNCION ASOCIADA AL TIPO DE COMPROBANTE
    If FORIPRE$ = "NULL" Then Exit Sub
    CODFOR$ = FORIPRE$
    '
    Screen.MousePointer = 11
    '
    'SELECT QUE SE UTILIZA MAS ABAJO PARA TOMAR NUMEROS DE SERIES
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCMOV% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DESELEC# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HASELEC# & " "
    Set CONSITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
 
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CANCELPRINT% = 0
    'FFF% = HOY(HO$)
    FFF% = CVINT(VALOBADAMDB("CAJABANC", "FechConta", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    DESCU$ = SAFETEXT$(VALOBADAMDB("CAJABANC", "Descu1", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    PORCENDESCU = COEFDESCU(DESCU$)
    
    FECHDOC$ = FECHATEX$(FFF%)
    'FECHDOC$ = HO$
    NUMEDOC$ = Mid$(NUFACTU$, 4, 13) 'FORMATO'0001-00000123'
    '
    'RECUPERO LOS DATOS DE LA BASE
    MONEMI% = XVALO(VALOBADAMDB("CAJABANC", "MONE_EMIS", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If MONEMI% < 1 Then MONEMI% = 1
    MONE$ = DEMONECO$(MONEMI%)
    TIPOCAM# = XVALO(VALOBADAMDB("CAJABANC", "Tipo_Cam", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
     
     
    IMPORTETOTALNETO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    IMPORTETOTALNETO# = IMPORTETOTALNETO# + XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * VALOUNIT)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
    
    IVA1050# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA1050# = IVA1050# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# + XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    IVA2100# = IVA2100# / TIPOCAM#
    IVA1050# = IVA1050# / TIPOCAM#
    COEFIVAINCLU# = 1
    If IMPORTETOTALNETO# > 0.005 Then
      COEFIVAINCLU# = (IMPORTETOTALNETO# + IVA1050# + IVA2100#) / IMPORTETOTALNETO#
    End If
    '
    'TOTALIVAINCLUIDO SE UTILIZA PARA CALCULAR EL DESCUENTO
    If TIFADOC% <> 6 Then
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTSALID * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
      'NOTA DE CREDITO
      TOTALIVAINCLUIDO# = XVALO(VALOBADAMDB("MOVISTO", "SUM(CANTINGRE * ValoUnitIva)", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS"))
      TOTALBRUTOSINIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'VENTA' AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# / TIPOCAM#
    If LETR$ = "B" Then
       TOTALBRUTOSINIVA# = TOTALBRUTOSINIVA# * COEFIVAINCLU#
       IMPORTETOTALNETO# = IMPORTETOTALNETO# * COEFIVAINCLU#
    End If
    '
    BASEIMPO21# = XVALO(VALOBADAMDB(" CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '21.00 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO21# = BASEIMPO21# / TIPOCAM#
    BASEIMPO105# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = '10.50 %'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    BASEIMPO105# = BASEIMPO105# / TIPOCAM#
    INOGRAVADO# = XVALO(VALOBADAMDB("CAJABANC", "BaseImpo", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = XVALO(VALOBADAMDB("CAJABANC", "IDebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# + XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    TOTALDOCUMENTO# = TOTALDOCUMENTO# / TIPOCAM#
    On Error Resume Next
    PORCENTAJEDESCUENTO = 100 * (1 - TOTALBRUTOSINIVA# / TOTALIVAINCLUIDO#)
    On Error GoTo 0
    IMPORTEDESC# = TOTALIVAINCLUIDO# - IMPORTETOTALNETO#
'''''    TOTALBRUTOSINIVACONDESCUENTO# = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN= " & NCBNC%, "NOMESS")) / TIPOCAM#
'''''    On Error Resume Next
'''''    PORCENTAJEDESCUENTO = ROUND((1 - (TOTALDOCUMENTO# / TOTALIVAINCLUIDO#)) * 100, 2)
'''''    On Error GoTo 0
'''''    TOTALBRUTOSINIVACONDESCUENTO# = TOTALBRUTOSINIVA# / (1 - (PORCENTAJEDESCUENTO / 100))
    If TIFADOC <> 6 Then
        PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
        PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IHabeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    Else
       PERCEPIVA# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IVA' AND VARIMOVIM = 'PER-IVA'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
       PERCEPBRUTO# = XVALO(VALOBADAMDB("CAJABANC", "IdebeComp", "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'IBRU' AND VARIMOVIM = 'PERC'  AND NUCLIEN = " & NCBNC%, "NOMESS"))
    End If
    '
    DESCRILAR$ = VALOBADAMDB("MOVISTO", "DESCRI_LAR", "DOSECLAR = '" & NUFACTU$ & "' AND Nume_Clie = " & NCMOV%, "NOMESS")
    CAE$ = VALOBADAMDB("CAJABANC", "CAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    IDCAE$ = VALOBADAMDB("CAJABANC", "IDCAE", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    'SE AGREGA FECHA DE VENCIMIENTO DE CAE
    FECHVENCICAE$ = VALOBADAMDB("CAJABANC", "FechVtoCae", "CODICOM_LAR = '" & NUFACTU$ & "' AND NUCLIEN = " & NCBNC%, "NOMESS")
    
    COBA$ = TRIM$(VALOBADAMDB("CAJABANC", "COD_BARRA", "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    If COBA$ = "" Then COBA$ = ARMACODIGOBARRA(NUFACTU$, CAE$, FECHVENCICAE$)
    
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Observa$ = "CAE: " + CAE$ + " ... Vto Cae: " + FECHVENCICAE$ + " ... Id.Tran: " + IDCAE$ + vbCrLf + VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")
    If InStr(EPAC$, "MECANIZA") > 0 Or InStr(EPAC$, "EDR") > 0 Then Observa$ = VALOBADAMDB("CAJABANC", "NOTAS", "CODICOM_LAR = '" & NUFACTU$ & "'   AND NUCLIEN = " & NCBNC%, "NOMESS")
       
    '
    'NUMERO DE PEDIDO PARA TOMAR DATOS DEL MISMO
    NROP& = XVALO(VALOBADAMDB("MOVISTO", "NuPedCli", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS"))
    NROPEDIDO$ = Str$(NROP&)
    If NROP& > 0 Then
       NUOCOMPRA$ = VALOBADAMDB("PEDENCA", "NroOcom", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
       VENDE% = XVALO(VALOBADAMDB("PEDENCA", "Vended", "NROPED = " & NROP&))
    Else
       'SI NO TIENE NUMERO DE PEDIDO PONE EL VENDEDOR DEL MAESTRO
       VENDE% = VENDCLI%(NCMOV%)
    End If
    NUVENDEX$ = Str$(VENDE%) + " " + ECOARCH$("VENDEDOR", Chr$(VENDE%))
    'CONDICION DE PAGO SI NO TIENE TOMA LA DEL MAESTRO
    CPAG% = XVALO(VALOBADAMDB("CAJABANC", "CondPag", "CodiCom_Lar = '" & NUFACTU$ & "' AND NuClien = " & NCBNC%, "NOMESS"))
    If CPAG% < 1 Then CPAG% = CPAGCLI%(NCMOV%)
    CONDPAG$ = ECOARCH$("CONPAG", Chr$(CPAG%))
    '
    NroRemi$ = VALOBADAMDB("MOVISTO", "DOPRILAR", "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%, "NOMESS")
    PORCENTAJEIVA1$ = "10.50 %"
    PORCENTAJEIVA2$ = "21.00 %"
    IEMBALA = XVALO(VALOBADAMDB("CAJABANC", "IHABECOMP", "CODICOM_LAR = '" & NUFACTU$ & "'AND TIPOMOVIM = 'EMBAL' AND  NUCLIEN = " & NCBNC%, "NOMESS"))
    '
    DOMICI$ = VALOBADAMDB("PEDENCA", "DomiEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    'SI NO TIENE DOMICILIO EL TRANSPORTE LE PONE EL TRANSPORTE DEL MAESTRO
    If TRIM$(DOMICI$) <> "" Then
      TRANSPO$ = VALOBADAMDB("PEDENCA", "TranspEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
      LOCALENT$ = VALOBADAMDB("PEDENCA", "LocaEnt", "NROPED = " & NROP& & "  AND NROCLIE = " & NCMOV%, "NOMESS")
    Else
      DOMICI$ = DTRANCLI$(NCMOV%)
      TRANSPO$ = RTRANCLI$(NCMOV%)
      NT% = NUTRANCLI(NCMOV%)
      TLOCALENT$ = LOCATRANS$(NT%)
    End If
    '
    
    'PARAMETROS DE IMPRESION
    '***********************
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST" 'ESTOS YA ESTABAN COMENTADOS
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    CODTABU1$(27) = "N-AUTENT"
    CODTABU1$(28) = "R-PLANO"
    CODTABU1$(29) = "F-PLANO"
    '
    CACOLTAB1% = 29
    '
    CAITAB1% = 0
    CAITCRU% = 0

    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETR$
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(NroRemi$)
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = TRIM$(NUOCOMPRA$)
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = TRIM$(NROPEDIDO$)
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
    CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = TRIM$(NUVENDEX$)
    CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = TRIM$(CONDPAG$)
    CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORCENTAJEIVA2$
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORCENTAJEIVA1$ 'PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = FORMATNUM$(PORCENTAJEDESCUENTO, "F13.2") 'TRIM$(PORDESCU.TEXT) ' VER CALCULO DE PORCENTAJE DE DESCUENTO
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""    ' Valor en PESOS equivalente a dolares
    '
    CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = TRIM$(DOMICI$) 'LENTREGA(0).TEXT
    CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = TRIM$(LOCALENT$) 'LENTREGA(1).TEXT
    CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = TRIM$(TRANSPO$) 'LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONE$
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
       Incoterms$ = VALOBADAMDB("CAJABANC", "Incoterms", "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCBNC%, "NOMESS")
    CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Incoterms$)          ' VALOR FOB ' ver como conseguir el val
       TIPOCAMM# = TIPOCAM#: If InStr(1, EMPREAC$, "GRAMY") > 0 And MONEMI% = 1 Then TIPOCAMM# = TICAMONE(2)
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = FORMATNUM$(TIPOCAMM#, "F12.4")              ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOCNUE)   '  CONTACTO
    CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(Str$(IEMBALA))                      ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(NCMOV%)))
    CODPARAM$(45) = "ITO-GRAV": DOCPARAM$(45) = ""
    CODPARAM$(46) = "ITO-NOGR": DOCPARAM$(46) = ""
    CODPARAM$(47) = "ITO-GRA1": DOCPARAM$(47) = ""
    CODPARAM$(48) = "ITO-GRA2": DOCPARAM$(48) = ""
    CODPARAM$(51) = "TOT-HABE": DOCPARAM$(51) = FORMATNUM(TOTATOTA#, "F13.2") ' VER ESTE TOTAL COMO SE CALCULA
    CODPARAM$(54) = "LAR-CUTE": DOCPARAM$(54) = CAE$
    CODPARAM$(55) = "LAR-MATE": DOCPARAM$(55) = FECHVENCICAE$
    CODPARAM$(56) = "LAR-CUCO": DOCPARAM$(56) = IDCAE$
    CODPARAM$(57) = "LAR-MACO": DOCPARAM$(57) = COBA$
    '
    'OBTENGO EL VARIDOC PARA TOMAR LA TASA CON ALIPERBRU
    '***************************************************
    VARIDOC% = 1: PIVADOC% = PIVACLI%(NCLIE%)
    If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
    ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
    End If
    TASAX = ALIPERBRU#(NCMOV%, MINIPER#, VARIDOC%)
    CODPARAM$(52) = "IMPO-CHE": DOCPARAM$(52) = FORMATNUM(TASAX, "F5.1")
    CODPARAM$(53) = "NUOPERS": DOCPARAM$(53) = GRUDIVCLI%(NCMOV%)
    '
    'IDCAE + CAE + OBSERVACIONES
    '****************************
    Call SETULTREG("!OBSERVOF", 0)
    If Observa$ <> "" Then
      JR& = 0
      FORMOBSE$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
      ANCHOB% = XVALO(Mid$(FORMOBSE$, 2)): If ANCHOB% < 32 Then ANCHOB% = 32
      Call FRATEXTO(Observa$, ANCHOB%)
      For KKII% = 1 To CARETEX%
        JR& = JR& + 1
        Call GRAREG("!OBSERVOF", RETEX$(KKII%), JR&)
      Next KKII%
      Call SETULTREG("!OBSERVOF", JR&)
      Call CIERRARCH("!OBSERVOF")
    End If
    POVERTI% = XVALO(TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/POSINVER"))))
    '
    CAPARDOC% = 57
    '
    REPLAORI% = 0
      If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
        REPLAORI% = 1
      End If
      If (U& Mod ITPORHOJA%) = 1 Then
        ' PONE ENCABEZADO
        CAITAB1% = 0
        CAHOJAS% = CAHOJAS% + 1
      End If
      '
      U& = 0
      'LO LEO ACA POR QUE SE UTILIZA EL VALOR CUANDO RECORRE EL MOVISTO PARA IMPRIIMIR EL DETALLE
      CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC"))
      '
      'SE AGREGA PARA OBTENER EL VALOR MAXIMO DE ITEMS DE LA FACTURA QUE SE VALIDA MAS ABAJO
      CONDI$ = "DOSECLAR = '" & NUFACTU$ & "'  AND NUME_CLIE = " & NCMOV%
      FIN& = XVALO(VALOBADAMDB("MOVISTO", "COUNT(*)", CONDI$, "NOMESS"))
      
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY NuorItem "
      Set MOVTMP = Stocks.OpenRecordset(SQLX$, 4)
      '
      With MOVTMP
       Do While .EOF = False
         U& = U& + 1
         If (U& Mod ITPORHOJA%) = 1 Then
            ' PONE ENCABEZADO
            CAITAB1% = 0
            CAHOJAS% = CAHOJAS% + 1
         End If
 
         CIPRE% = XVALO(!Cod_Inte)
         CANTID = Abs(XVALO(!CANTSALID) - XVALO(!CANTINGRE))
         CAPEX$ = Str$(CANTID)
         NPEDIDD& = XVALO(!NUPEDCLI)
         SUBCO$ = SUBCODI$(NPEDIDD&, CIPRE%)
         ORDENCOMPRA$ = SAFETEXT$(NumeOrdPed)
         NroRemi$ = SAFETEXT$(!DOPRILAR)
         VALORUNITARIO# = XVALO(!VALOUNITIVA)
           If LETR$ = "B" Then
              VALORUNITARIO# = VALORUNITARIO# / (1 + PORCEIVA(CIPRE%) / 100)
           End If
         CAITAB1% = CAITAB1% + 1
         CODDX$ = CODEXT$(CIPRE%) + SUBCO$ ' en coddx$ trae el codigo externo a partir de cipre(int)
         If REPLAORI% = 1 Then
           CODDX$ = CODORIG$(CIPRE%, NCMOV%) ' + SUBCO$
           'se pasa el codigo interno, mas el nº cliente(NCMOV%) y trae el codigo externo solicitado
        End If
        TETABU1$(1, CAITAB1%) = CODDX$
        DESCRILAR$ = SAFETEXT$(!DESCRI_LAR)
        If CIPRE% > 0 Then
         TEXTO$ = DESCRIT0$(CIPRE%)
        Else
         TEXTO$ = DESCRILAR$
        End If
        TETABU1$(2, CAITAB1%) = TEXTO$
        TETABU1$(3, CAITAB1%) = CAPEX$
        'TETABU1$(4, CAITAB1%) = PRELI$
        'TETABU1$(5, CAITAB1%) = DESCX$
        'TETABU1$(6, CAITAB1%) = PREUX$
        'TETABU1$(7, CAITAB1%) = ITOTX$
        'TETABU1$(8, CAITAB1%) = FENTX$
        TETABU1$(9, CAITAB1%) = Str$(U&)
        TETABU1$(10, CAITAB1%) = UNIMED$(CIPRE%)    ' unidad de medida
       TETABU1$(14, CAITAB1%) = NUPLANO$(CIPRE%) ' numero de plano
       TETABU1$(16, CAITAB1%) = ORDENCOMPRA$  ' numero de o.compra
       TETABU1$(17, CAITAB1%) = NroRemi$  ' numero de remito
      '
      FACIVITEM = 1
      If LETR$ = "B" Then FACIVITEM = 1 + PORCEIVA(CIPRE%) / 100
      '
      VALORUNITARIO# = VALORUNITARIO# * FACIVITEM
      TETABU1$(22, CAITAB1%) = FORMATNUM$(VALORUNITARIO#, "F15.4")  ' precio unitario
      TETABU1$(23, CAITAB1%) = TRIM$(FORMATNUM$(CANTID * VALORUNITARIO#, "F15.4"))   ' total neto
         If TETABU1$(23, CAITAB1%) = "" Then TETABU1$(23, CAITAB1%) = "S/Cargo"
      TETABU1$(24, CAITAB1%) = FORMATNUM$((PORCEIVA(CIPRE%)), "F5.2") ' alicuota iva
      TETABU1$(25, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#) - (VALORUNITARIO#), "F15.4") ' importe iva del item
      TETABU1$(26, CAITAB1%) = FORMATNUM$((FACIVITEM * VALORUNITARIO#), "F15.4") ' total bruto
         If TETABU1$(26, CAITAB1%) = "" Then TETABU1$(26, CAITAB1%) = "S/Cargo"
      TETABU1$(27, CAITAB1%) = Mid$(XXX$, 371, 16) ' numero de pedido cliente
      ' AQUI BUSCA EL NUMERO DE SERIE
        '
        SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
        SQLX$ = SQLX$ + "AND Cod_Inte = " & CIPRE% & " "
        SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
        '
        With CONSITEMP
          .FindFirst SQLX$
25        If .NoMatch = False Then
            SERNU$ = !ReferCor
            If Left$(SERNU$, 5) = "S.No." Then
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = SERNU$
            End If
          End If
        End With
      
      '
      ' AQUI AGREGA DESPACHO DE IMPORTACION 'VER
      If AGREDESP% = 0 Then
        AGREDESP% = (-1)
        If CONTROL$("FACTURA", "AGREDESP") = "SI" Then AGREDESP% = 1
      End If
      If AGREDESP% = 1 Then
        DESPAI$ = TRIM$(ATRIB1$(CIPRE%))
        If DESPAI$ <> "" Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Despacho: " + DESPAI$
        End If
      End If
      '
'SE COMENTO POR QUE ESTA TRAYENDO LA ESPECIFICACION SIN CONFIGURACION Y EN TODO CASO SE HARA POR CONFIGURACION
'      ESPECIFIX$ = ESPECIFIT$(CIPRE%) ' TOMA LA ESPECIFICACION DEL ARTICUL
'      PARAMEXT$ = CODTABU1$(2) ' NOMBRE DEL PARAMETRO
'      Call PRINTEXT(ESPECIFIX$, 2, PARAMEXT$)
      '
      'INTERVENCION DEL CIELO
      NROPEDI& = XVALO(!NUPEDCLI)
      CONDI$ = "NROPED = " & NROPEDI&
      REFERPEDI$ = TRIM$(VALOBADAMDB("PEDENCA", "Referencia", CONDI$, "NOMESS"))
      'If REFERPEDI$ <> "" And InStr(EPAC$, "DEL CIELO") > 0 Then
      If REFERPEDI$ <> "" And CONTROL("FACTURA", "SISTMANT") = "SI" Then
         'TEXTO EJ REFERENCIA: LIMPIEZA DE TANQUE AGOSTO 2015 (1/6)
        POS1% = InStr(REFERPEDI$, "|") ' ME QUEDO SOLAMENTE CON EL MES Y EL NRO DE CUOTA AGOSTO 2015 (1/6)
        If POS1% > 0 Then REFERPEDI$ = TRIM$(Mid$(REFERPEDI$, POS1% + 1))
        Call PRINTEXT(REFERPEDI$, 2, PARAMEXT$)
      End If
      'FIN INTERVENCION DEL CIELO
      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
 
      '
      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
      SUBTOTB# = SUBTOTB# + XVALO(TETABU1$(26, CAITAB1%))
      '
      If (U& Mod ITPORHOJA%) = 0 Then
        If U& < FIN& Then
          If CACOPDOC% > 0 Then
            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
            Call PRINTDOC("@" + FORIPRE$)
            Call CIERRARCH("*.*")
            Call BAJALDISCO
          End If
        End If
      End If
 
       .MoveNext
       Loop
     End With
'''''''''''    '
 
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = FORMATNUM$(TOTALIVAINCLUIDO# + IEMBALA, "F13.2")
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = IVA2100#
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = IVA1050#
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = FORMATNUM$(IMPORTEDESC#, "F15.4") 'Total(1).TEXT
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = IMPORTETOTALNETO#
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = TOTALDOCUMENTO#
     VADOLA# = TOTALDOCUMENTO# * TIPOCAM#
     DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = PERCEPBRUTO#
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = PERCEPIVA#                         'PARA IMPRIMIR PERC-IVA
    ITOGRAVA# = TOTALBRUTOSINIVA#
    DOCPARAM$(45) = TRIM$(FORMATNUM$(ITOGRAVA#, "F13.2"))
    INOGRAVA# = INOGRAVADO#
    DOCPARAM$(46) = TRIM$(FORMATNUM$(INOGRAVA#, "F13.2"))
    DOCPARAM$(47) = BASEIMPO21#
    DOCPARAM$(48) = BASEIMPO105#
 
'''''    If LETR$ = "B" Then
'''''      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * TOTALBRUTOSINIVA# + IEMBALA, "F13.2")
'''''      DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * PORCENTAJEDESCUENTO, "F13.2")
'''''      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2")
'''''      DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
'''''      DOCPARAM$(45) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IMPORTETOTALNETO#, "F13.2"))
'''''      DOCPARAM$(46) = TRIM$(FORMATNUM$(COEFIVAINCLU# * INOGRAVA#, "F13.2"))
'''''      DOCPARAM$(47) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA1050#, "F13.2"))
'''''      DOCPARAM$(48) = TRIM$(FORMATNUM$(COEFIVAINCLU# * IVA2100#, "F13.2"))
'''''    End If
 '''''''''''    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TOTFAC$ = NUMTEX$(TOTALDOCUMENTO#)
    If MONE$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS") ' MODIFICADO POR QUE SIEMPRE IMPRIMIA PESOS
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONE$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
'''''''''''
    PRINTERPORT$ = PORTUFAC$
'    CACOPDOC% = XVALO(CONTROL$("", "COPIAFAC")) LO LEO ARRIBA
    If CACOPDOC% > 0 Then
       Call PRINTDOC("@" + FORIPRE$)
    End If
'''''''''''    '
'    NUCON& = XVALO(CONTROL("ALATALO", FORIPRE$))
'    NUPANT& = XVALO(Text11)
'    If NUCON& > 0 Then
'       If NUCON& < NUPANT& Then
'          Call COMUNI("ATENCION: Alarma de Talonario Proximo a Terminarse.")
'       End If
'    End If
'''''''''''    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub


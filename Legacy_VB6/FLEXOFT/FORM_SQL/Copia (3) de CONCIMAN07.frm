VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form CONCIMAN07 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situación de Conciliación "
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Selección Multiple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6930
      TabIndex        =   26
      ToolTipText     =   "Permite Seleccionar Varios Movimientos y Conciliar con 1 de los Pendientes del Sistema"
      Top             =   5700
      Width           =   1890
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   8925
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8610
      Width           =   1425
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   8925
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5670
      Width           =   1425
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   8925
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2835
      Width           =   1425
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10635
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Conciliar Gastos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      TabIndex        =   18
      Top             =   8595
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   17
      Top             =   8595
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Concilia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3615
      TabIndex        =   16
      Top             =   8595
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8520
      Width           =   10350
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   8895
      Left            =   10400
      TabIndex        =   14
      Top             =   -70
      Width           =   80
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   9015
      Left            =   0
      TabIndex        =   13
      Top             =   -70
      Width           =   80
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "MOVIMIENTOS DEL BANCO (en el sistema) PENDIENTES DE CONCILIACIÓN"
      Top             =   5720
      Width           =   6780
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "LISTA DE MOVIMIENTOS NO CONCILIADOS AUTOMATICAMENTE"
      Top             =   2870
      Width           =   6780
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FBCA9D&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "LISTA DE MOVIMIENTOS CONCICLIADOS AUTOMATICAMENTE"
      Top             =   50
      Width           =   10350
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   8955
      Left            =   10440
      ScaleHeight     =   8895
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   210
         Picture         =   "CONCIMAN07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   210
         Picture         =   "CONCIMAN07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Borra Conciliación Activa"
         Top             =   6090
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Procesar"
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
         Left            =   60
         TabIndex        =   19
         ToolTipText     =   "Procesa y Marca los Movimientos Conciliados"
         Top             =   7665
         Width           =   880
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   8250
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -250
         Picture         =   "CONCIMAN07.frx":0614
         Top             =   8470
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2500
      Left            =   45
      TabIndex        =   6
      Top             =   315
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2500
      Left            =   0
      TabIndex        =   8
      Top             =   3150
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   2500
      Left            =   50
      TabIndex        =   9
      Top             =   6000
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12648384
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -840
      TabIndex        =   7
      Top             =   960
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   16120
      Width           =   870
   End
End
Attribute VB_Name = "CONCIMAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODICONCI$
Public REGSEL1&
Public REGSEL2&
Public REGSEL3&

Function CUENTAGASTO%(NUORD%)
   CUENTAGASTO% = XVALO(VALOBADA("CONCILIA", "CUECONTAI", "CODICONCI = '" & CODICONCI$ & "'ORDCONCILIA = " & NUORD%, "NOMESS"))
End Function
Function FECHADECONCILIA%(NUORD%)
   FECHADECONCILIA% = CVINT(VALOBADA("CONCILIA", "FECHA", "CODICONCI = '" & CODICONCI$ & "' AND ORDCONCILIA = " & NUORD%, "NOMESS"))
End Function

Sub sortflex(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    List1.Clear
    For I& = 1 To .Rows - 1
        T1$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(I&, 3))))
        T2$ = .TextMatrix(I&, 2)
        T3$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(I&, 1))))
        T4$ = .TextMatrix(I&, 4)
        T5$ = .TextMatrix(I&, 5)
        T6$ = .TextMatrix(I&, 6)
        T7$ = .TextMatrix(I&, 7)
        T8$ = .TextMatrix(I&, 8)
        T9$ = .TextMatrix(I&, 9)
        '
        Call REPLA(TX$, T1$, 1, 10)
        Call REPLA(TX$, T2$, 11, 20)
        Call REPLA(TX$, T3$, 31, 10)
        Call REPLA(TX$, T4$, 41, 70)
        Call REPLA(TX$, T5$, 111, 20)
        Call REPLA(TX$, T6$, 131, 20)
        Call REPLA(TX$, T7$, 151, 10)
        Call REPLA(TX$, T8$, 161, 10)
        Call REPLA(TX$, T9$, 171, 10)
        '
        List1.AddItem TX$
    Next I&
    .Rows = 1
    For I& = 1 To List1.ListCount
     TTXX$ = List1.List(I& - 1)
     T3$ = Mid$(TTXX$, 1, 10)
     T2$ = Mid$(TTXX$, 11, 20)
     T1$ = Mid$(TTXX$, 31, 10)
     T4$ = Mid$(TTXX$, 41, 70)
     T5$ = Mid$(TTXX$, 111, 20)
     T6$ = Mid$(TTXX$, 131, 20)
     T7$ = Mid$(TTXX$, 151, 10)
     T8$ = Mid$(TTXX$, 161, 10)
     T9$ = Mid$(TTXX$, 171, 10)
     '
     REG& = REG& + 1
     .Rows = REG& + 1
        FECH_EXTR% = XVALO(T1$)
        COMPRBTE$ = T2$
        FECONTA% = XVALO(T3$)
        CONCEP$ = T4$
        IMPODEBEX# = XVALO(T5$)
        IMPOHABEX# = XVALO(T6$)
        NUORD% = XVALO(T7$)
        DETABLA$ = T8$
        CUEII% = XVALO(T9$)
        '
        .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
        .TextMatrix(REG&, 2) = COMPRBTE$
        .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
        .TextMatrix(REG&, 4) = CONCEP$
        .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
        .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
        .TextMatrix(REG&, 7) = NUORD%
        .TextMatrix(REG&, 8) = DETABLA$
        .TextMatrix(REG&, 9) = ""

    Next I&
    MSF.Redraw = True
  End With
End Sub
Sub MULTISELEC(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    List2.Clear
    For I& = 1 To .Rows - 1
        T1$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(I&, 3))))
        T2$ = .TextMatrix(I&, 2)
        T3$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(I&, 1))))
        T4$ = .TextMatrix(I&, 4)
        T5$ = .TextMatrix(I&, 5)
        T6$ = .TextMatrix(I&, 6)
        T7$ = .TextMatrix(I&, 7)
        T8$ = .TextMatrix(I&, 8)
        T9$ = .TextMatrix(I&, 9)
        '
        Call REPLA(TX$, T1$, 1, 10)
        Call REPLA(TX$, T2$, 11, 20)
        Call REPLA(TX$, T3$, 31, 10)
        Call REPLA(TX$, T4$, 41, 70)
        Call REPLA(TX$, T5$, 111, 20)
        Call REPLA(TX$, T6$, 131, 20)
        Call REPLA(TX$, T7$, 151, 10)
        Call REPLA(TX$, T8$, 161, 10)
        Call REPLA(TX$, T9$, 171, 10)
        '
        List1.AddItem TX$
    Next I&
    .Rows = 1
    For I& = 1 To List1.ListCount
     TTXX$ = List1.List(I& - 1)
     T3$ = Mid$(TTXX$, 1, 10)
     T2$ = Mid$(TTXX$, 11, 20)
     T1$ = Mid$(TTXX$, 31, 10)
     T4$ = Mid$(TTXX$, 41, 70)
     T5$ = Mid$(TTXX$, 111, 20)
     T6$ = Mid$(TTXX$, 131, 20)
     T7$ = Mid$(TTXX$, 151, 10)
     T8$ = Mid$(TTXX$, 161, 10)
     T9$ = Mid$(TTXX$, 171, 10)
     '
     REG& = REG& + 1
     .Rows = REG& + 1
        FECH_EXTR% = XVALO(T1$)
        COMPRBTE$ = T2$
        FECONTA% = XVALO(T3$)
        CONCEP$ = T4$
        IMPODEBEX# = XVALO(T5$)
        IMPOHABEX# = XVALO(T6$)
        NUORD% = XVALO(T7$)
        DETABLA$ = T8$
        CUEII% = XVALO(T9$)
        '
        .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
        .TextMatrix(REG&, 2) = COMPRBTE$
        .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
        .TextMatrix(REG&, 4) = CONCEP$
        .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
        .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
        .TextMatrix(REG&, 7) = NUORD%
        .TextMatrix(REG&, 8) = DETABLA$
        .TextMatrix(REG&, 9) = ""

    Next I&
    MSF.Redraw = True
  End With
End Sub


Sub SUMAMSF(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    ACUMULO# = 0
    For I& = 1 To .Rows - 1
        DEBEX# = XVALO(.TextMatrix(I&, 5))
        HABEX# = XVALO(.TextMatrix(I&, 6))
        SALDO# = DEBEX# - HABEX#
        ACUMULO# = ACUMULO# + SALDO#
    Next I&
    .Redraw = True
    'TOMO EL VALOR DE LA ULTIMA LETRA DEL MSF1, MSF2 O MSF3
    IND% = 3 + XVALO(Right(.Name, 1))
    Text16(IND%) = FORMATNUM$(ACUMULO#, "F14.2")
 End With
End Sub

Private Sub AGenRep_Click()

    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub



Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_2% = 0
    
    With MSF2
      For I& = 1 To .Rows - 1
          If .TextMatrix(I&, 0) = "*" Then
             HAY_SELECCION_2% = 1
             REG2& = I&
           Exit For
          End If
        Next I&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_2% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 2")
       GoTo 99
    End If
    '
    If MSF2.TextMatrix(REG2&, 2) <> "GASTOS" Then
        'VARIABLE CODICONCI DE ALCANCE PUBLICO A NIVEL DE FRM
        ESGASTO% = VALIDA_SI_ES_GASTO%(CODICONCI$, MSF2.TextMatrix(REG2&, 4))
        If ESGASTO% < 1 Then
          OPX% = COMALTER%("ATENCION: Registro Seleccionado No Corresponde A Gastos.\\Cancelar\Ver Tabla de Gastos")
          If OPX% <> 2 Then GoTo 99
          Call OPCHEQ07.CARGA_TABLA_GASEXTRA
          GoTo 99
        End If
        
    End If
    
    'SUBO LA INFO A LISTA 1
    NUORD% = XVALO(MSF2.TextMatrix(REG2&, 7))

    With MSF2
           J& = MSF1.Rows
           MSF1.Rows = J& + 1
           MSF1.TextMatrix(J&, 1) = .TextMatrix(REG2&, 1)
           MSF1.TextMatrix(J&, 2) = "GASTOS"
           MSF1.TextMatrix(J&, 3) = FECHATEX$(FECHADECONCILIA%(NUORD%))
           MSF1.TextMatrix(J&, 4) = .TextMatrix(REG2&, 4)
           MSF1.TextMatrix(J&, 5) = .TextMatrix(REG2&, 5)
           MSF1.TextMatrix(J&, 6) = .TextMatrix(REG2&, 6)
           MSF1.TextMatrix(J&, 7) = .TextMatrix(REG2&, 7)
           MSF1.TextMatrix(J&, 8) = .TextMatrix(REG2&, 8)
           MSF1.TextMatrix(J&, 9) = SAFETEXT$(CUENTAGASTO%(NUORD%))
           Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF2, &HC0E0FF, REG2&, "*", 1)
    End With
    '
    NU0RD% = XVALO(MSF2.TextMatrix(REG2&, 7))
    Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", "ORDCONCILIA  = " & NU0RD%, "GASTOS", REGAF&, "NOMESS")

    'BORRO LA FILA DE LISTA 2
    Call borrafila(MSF2, REG2&)
    '
    'BORRO LA FILA DE LISTA 3
    Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, REG3&, "*", 1)
'    Call borrafila(MSF3, REG3&)
    '
    Call sortflex(MSF1)
    Call sortflex(MSF2)
    Call SUMAMSF(MSF1)
    Call SUMAMSF(MSF2)

    'llevo al final la grilla 1
    TPXI& = MSF1.Rows - 9: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0

99  Command2.Enabled = True

End Sub


Private Sub Command24_Click()
   If DERACCE%("BORRACON", "Esta Seguro de Borrrar la Conciliación Activa") >= 2 Then
      If CantRegistros&("CONCILIA", "STATUS > 0 ", "NOMESS") < 1 Then
        Call BORRAREGISTROS("CONCILIA", "CODICONCI = '" & CODICONCI$ & "'", REGAF&, "NOMESS")
      Else
        Call COMUNI("Imposible Borrar\La Conciliación Activa Se Encuentra Procesada")
      End If
   End If

End Sub

Private Sub Command3_Click()

   Command3.Enabled = False
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_1% = 0
    With MSF1
      For I& = 1 To .Rows - 1
          If .TextMatrix(I&, 0) = "*" Then
             HAY_SELECCION_1% = 1
             REG1& = I&
           Exit For
          End If
        Next I&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_1% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 1")
       GoTo 99
    End If
   '
   '------------------------------------------------------------------
   'TRASLADA LA INFORMACION SELECCIONADA A LA GRILLA 2
   With MSF1
        J& = MSF2.Rows
        MSF2.Rows = J& + 1
        MSF2.TextMatrix(J&, 1) = .TextMatrix(REG1&, 1)
        MSF2.TextMatrix(J&, 2) = ""
        MSF2.TextMatrix(J&, 3) = ""
        MSF2.TextMatrix(J&, 4) = .TextMatrix(REG1&, 4)
        MSF2.TextMatrix(J&, 5) = .TextMatrix(REG1&, 5)
        MSF2.TextMatrix(J&, 6) = .TextMatrix(REG1&, 6)
        MSF2.TextMatrix(J&, 7) = .TextMatrix(REG1&, 7)
        MSF2.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
        MSF2.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
        
        '
        If TRIM$(UCase$(.TextMatrix(REG1&, 2))) = "GASTOS" Then GoTo 70 'BORRA EL COMPROBANTE GASTOS
        '--------------------------------------------------------------
        'AGREGO A LA LISTA 3
        J& = MSF3.Rows
        MSF3.Rows = J& + 1
        MSF3.TextMatrix(J&, 1) = .TextMatrix(REG1&, 1)
        MSF3.TextMatrix(J&, 2) = .TextMatrix(REG1&, 2)
        MSF3.TextMatrix(J&, 3) = .TextMatrix(REG1&, 3)
           NUORD& = XVALO(.TextMatrix(REG1&, 7))
           CONCEP$ = Left$(VALOBADA("CONCILIA", "CONCEPTO_SISTEMA", "ORDCONCILIA = " & NUORD&), 26)
        MSF3.TextMatrix(J&, 4) = CONCEP$ 'CONCEPTO DEL SISTEMA (DET.REF.)
        MSF3.TextMatrix(J&, 5) = .TextMatrix(REG1&, 5)
        MSF3.TextMatrix(J&, 6) = .TextMatrix(REG1&, 6)
        MSF3.TextMatrix(J&, 7) = "" 'NUNERO DE ORDEN VACIO
        MSF3.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
        MSF3.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
        '
        'BORRO LA FILA DE LISTA 1
        'CON ESTO BLANQUEA LA FILA PINTADA
70      Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, REG1&, "*", 1)
        ORDEN& = XVALO(MSF1.TextMatrix(REG1&, 7))
        'BORRO EN EL CONCILIA EL COMPROBANTE Y LA FECHA
        
        Call BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
        'BORRO LA FILA DE LISTA 2
        Call borrafila(MSF1, REG1&)
        
    End With
99  Command3.Enabled = True
    Call sortflex(MSF2)
    Call sortflex(MSF3)
    Call SUMAMSF(MSF1)
    Call SUMAMSF(MSF2)
    Call SUMAMSF(MSF3)

End Sub

Sub BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
       CONDI$ = "ORDCONCILIA = " & ORDEN&
       CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
       Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, "", REGAF&, "NOMESS")
       Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI$, "", REGAF&, "NOMESS")
End Sub


Function borrafila(MSF As MSFlexGrid, REG&)

        For I& = REG& To MSF.Rows - 2
          For J& = 1 To MSF.Cols - 1
            MSF.TextMatrix(I, J) = MSF.TextMatrix(I + 1, J)
          Next
        Next
        MSF.Rows = MSF.Rows - 1
        
End Function

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_2% = 0
    
    With MSF2
      For I& = 1 To .Rows - 1
          If .TextMatrix(I&, 0) = "*" Then
             HAY_SELECCION_2% = 1
             REG2& = I&
           Exit For
          End If
        Next I&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_2% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 2")
       GoTo 99
    End If
    '
    'VALIDO QUE HAYA SELECCIONADO EN LISTA 3
    HAY_SELECCION_3% = 0
    With MSF3
      For I& = 1 To .Rows - 1
          If .TextMatrix(I&, 0) = "*" Then
             HAY_SELECCION_3% = 1
             REG3& = I&
           Exit For
          End If
        Next I&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 3
    If HAY_SELECCION_3% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 3")
       GoTo 99
    End If
    '
    'VALIDO QUE NO SE SELECCIONE VALORES CRUZADOS
    DEBEX2# = XVALO(MSF2.TextMatrix(REG2&, 5))
    HABEX2# = XVALO(MSF2.TextMatrix(REG2&, 6))
    DEBEX3# = XVALO(MSF3.TextMatrix(REG3&, 5))
    HABEX3# = XVALO(MSF3.TextMatrix(REG3&, 6))
    If Abs(DEBEX2# - DEBEX3#) > 0.005 Or Abs(HABEX2# - HABEX3#) > 0.005 Then
      Call COMUNI("Imposible Conciliar Valores Debe y Haber No Coincidente")
      GoTo 99
    End If
    
    'SUBO LA INFO A LISTA 1
    With MSF2
           J& = MSF1.Rows
           MSF1.Cols = 10
           MSF2.Cols = 10
           MSF1.Rows = J& + 1
           MSF1.TextMatrix(J&, 1) = .TextMatrix(REG2&, 1)
           MSF1.TextMatrix(J&, 2) = MSF3.TextMatrix(REG3&, 2) 'COMPROBANTE QUE LO TIENE EN LA LISTA 3
           MSF1.TextMatrix(J&, 3) = MSF3.TextMatrix(REG3&, 3) 'FECHA DE ACREDITACION
           MSF1.TextMatrix(J&, 4) = .TextMatrix(REG2&, 4) 'REFERENCIA DEL PENDIENTE DEL SISTEMA
           MSF1.TextMatrix(J&, 5) = .TextMatrix(REG2&, 5)
           MSF1.TextMatrix(J&, 6) = .TextMatrix(REG2&, 6)
           MSF1.TextMatrix(J&, 7) = .TextMatrix(REG2&, 7)
           MSF1.TextMatrix(J&, 8) = .TextMatrix(REG2&, 8)
           MSF1.TextMatrix(J&, 9) = .TextMatrix(REG2&, 9)
           
           
           'CON ESTO BLANQUE LA FILA PINTADA
           Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF2, &HC0E0FF, REG2&, "*", 1)
           'GRABO EN EL CONCILIA EL COMPROBANTE QUE LEO DE LA GRILLA 3
           COMPROBANTE$ = MSF3.TextMatrix(REG3&, 2)
           ' NUMERO DE ORDEN DEL CONCILIA PARA CONDICIONAR EL UPDATE
           ORDEN& = XVALO(MSF2.TextMatrix(REG2&, 7))
           FE% = FECHANUM(.TextMatrix(REG2&, 1))
           Call GRABA_COMPROBANTE_EN_CONCILIA(COMPROBANTE$, FE%, ORDEN&)
    End With
    '
    'BORRO LA FILA DE LISTA 2
    Call borrafila(MSF2, REG2&)
    '
    'BORRO LA FILA DE LISTA 3
    'CON ESTO BLANQUE LA FILA PINTADA
    Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, REG3&, "*", 1)
    Call borrafila(MSF3, REG3&)
    '
    Call sortflex(MSF1)
    
    Call SUMAMSF(MSF1)
    Call SUMAMSF(MSF2)
    Call SUMAMSF(MSF3)
    'llevo al final la grilla 1
    TPXI& = MSF1.Rows - 9: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0
    '
99  Command5.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    '
    If MSF2.Rows > 1 Then
       Call COMUNI("Se Deberan Conciliar Todos Los Movimientos.")
       GoTo 99
    End If
    
    CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUEBA$ = CODCUE$(CUEII%)
    '
    FIN& = MSF1.Rows - 1
    If FIN& < 2 Then GoTo 99
    For REG1& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      '
      With MSF1
        FECONTA1 = FETEXCOR1$(CVINT(.TextMatrix(REG1&, 1)))
        IDASIE$ = TRIM$(.TextMatrix(REG1&, 2))
        If UCase$(TRIM$(IDASIE$)) = "GASTOS" Then GoTo 29
        FEACRED% = CVINT(.TextMatrix(REG1&, 3))
        FEACRE = FETEXCOR1$(CVINT(.TextMatrix(REG1&, 3)))
        REFEMOV$ = TRIM$(.TextMatrix(REG1&, 4))
        IDEBE# = XVALO(.TextMatrix(REG1&, 5))
        IHABE# = XVALO(.TextMatrix(REG1&, 6))
        NUORD& = XVALO(.TextMatrix(REG1&, 7))
        ORIMO$ = TRIM$(.TextMatrix(REG1&, 8))
        
        If ORIMO$ = "CABAN" Then
            SQLX$ = "SELECT * FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE FECHCONTA = '" & FECONTA1 & "' "
            SQLX$ = SQLX$ + "AND CodiCom_COR = '" & IDASIE$ & "' "
            SQLX$ = SQLX$ + "AND CueContaI = " & CUEII% & " "
            '
            Dim CAJJJABANC As ADODB.Recordset
            Set CAJJJABANC = New ADODB.Recordset
27          On Error Resume Next
            CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 27
            End If
            On Error GoTo 0
            With CAJJJABANC    ' CajaBanc
                CAMOCO% = 0
                Do While Not .EOF
                  If Abs(IDEBE# - !IDEBECOMP) < 0.005 Then
                    If Abs(IHABE# - !IHABECOMP) < 0.005 Then
                      If FEACRED% > 0 Then
                          !MARCONSI = "*"
                          !fechacred = FETEXCOR1$(FEACRED%)
                          CAMOCO% = CAMOCO% + 1
                          If CAMOCO% > 1 Then
                             Call COMUNI("Movimientos Adicionales Conciliados\con Igual Importe e Identificador.")
                          End If
                      Else
                          !MARCONSI = " "
                          CAMOCO% = CAMOCO% + 1
                          If CAMOCO% > 1 Then
                             Call COMUNI("Movimientos Adicionales Des-Conciliados\con Igual Importe e Identificador.")
                          End If
                      End If
                      .Update
                    End If
                  End If
                  .MoveNext
                Loop
              On Error GoTo 0
            End With

        ElseIf ORIMO$ = "DIAR" Then
               SQLX$ = "SELECT * FROM DetaAsien "
               SQLX$ = SQLX$ + "WHERE FechAsien = '" & FECONTA1 & "' "
               ' SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & ""
               SQLX$ = SQLX$ + "AND NuIdAsien = '" & IDASIE$ & "' "
               SQLX$ = SQLX$ + "AND CueCoInt = " & CUEII% & " "
               SQLX$ = SQLX$ + "AND LEFT(RefePase, 28) = '" & REFEMOV$ & "' "
               SQLX$ = SQLX$ + "AND StatPase > 0 "
               '
               Dim DETTTASIEN As ADODB.Recordset
               Set DETTTASIEN = New ADODB.Recordset
28             On Error Resume Next
               DETTTASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                  On Error GoTo 0
                  Call CapturaErrorOpen
                  GoTo 28
               End If
               On Error GoTo 0
               With DETTTASIEN
                 On Error Resume Next
                 .MoveFirst
                 If Err < 1 Then
                   Do While Not .EOF
                     On Error GoTo 0
                     If Abs(IDEBE# - !IDebePase) < 0.005 Then
                       If Abs(IHABE# - !IHabePase) < 0.005 Then
                         If FEACRED% > 0 Then
                             !MARCONSI = "*"
                         Else
                             !MARCONSI = " "
                         End If
                         .Update
                         GoTo 29
                       End If
                     End If
                     .MoveNext
                   Loop
                 End If
                 On Error GoTo 0
               End With
        End If

      End With
29  Next REG1&
    '--------------------------------------------------------------------------
    'GENERACION DE ASIENTOS DE GASTOS
    '--------------------------------------------------------------------------
    NUASI$ = ASIDIAR07.NUPROA$
    FECHAU% = HOY(HO$)
    REFEASI$ = REFECONCI_AUTOMA$(CODICONCI$)

    CANASIEN& = CantRegistros("CONCILIA", "CODICONCI = '" & CODICONCI$ & "' AND COMPROBANTE = 'GASTOS' ")
    If CANASIEN& < 1 Then GoTo 99
    Dim ENCTMP As ADODB.Recordset
    Set ENCTMP = New ADODB.Recordset
    SQLX$ = "Select * from ENCABEASIEN where NuIDAsien = '" & NUASI$ & "'"
    ENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With ENCTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUIDASIEN = NUASI$
          !Refeasien = "CA-GASTOS " + Left$(REFEASI$, 54)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = 0
          !NuDfAsien = 0
          !TIPOASIEN = 0
          !StatAsien = 1  ' activo
          !EditAsien = (-1) 'BLOQUEADO PARA SU MODIFICACION
          !QPasAsien = CANASIEN&
          !DebeAsien = 0
          !Observa = " "
          !Nume_Usuar = USNBR%
          !Niv_Clasi = 0
          .Update
        End With
    ENCTMP.Close
    Set ENCTMP = Nothing
     'FIN DE GRABACION DE ENCABEZADO
    '--------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------
    'DETALLE DE ASIENTO
    '--------------------------------------------------------------------------
    Dim DETATMP As ADODB.Recordset
    Set DETATMP = New ADODB.Recordset
    SQLX$ = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
    DETATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    ORPASE% = 0
    
    REG1& = 0
    For REG1& = 1 To MSF1.Rows - 1
      Call TRACE(20, U&, FIN&)
      If TRIM$(MSF1.TextMatrix(REG1&, 2)) <> "GASTOS" Then GoTo 80
      FECONTA1 = FETEXCOR1$(CVINT(MSF1.TextMatrix(REG1&, 1)))
      FEACRED% = CVINT(MSF1.TextMatrix(REG1&, 3))
      DEPASE$ = MSF1.TextMatrix(REG1&, 4)
      IDEBE# = XVALO(MSF1.TextMatrix(REG1&, 5))
      IHABE# = XVALO(MSF1.TextMatrix(REG1&, 6))
      COCUE$ = TRIM$(MSF1.TextMatrix(REG1&, 9))
      CTA_GASTO% = CUEINT%(COCUE$)
      ACUMU_IDEBE# = 0
      '
      With DETATMP
40           .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !NUIDASIEN = NUASI$
               ORPASE% = ORPASE% + 1
             !OrNuPase = ORPASE%
             !RefePase = Left$(DEPASE$, 64)
             !FECHASIEN = CVDAT(FECHAU%)
               PPXX% = InStr(NUASI$, "-")
             !OrFeAsien = XVALO(Mid$(NUASI$, PPXX% + 1))
             !SucuAsien = 0
             !NuDfAsien = 0
             !StatPase = 1
             !CUECOEXT = COCUE$
             !CUECOINT = CTA_GASTO%
             !IDebePase = IDEBE#
               ACUMU_IDEBE# = ACUMU_IDEBE# + IDEBE#
             !IHabePase = IHABE#
             !MARCONSI = MACON$
             !CenCos = 0
             !Niv_Clasi = 0
             .Update
80           If REG1& = FIN& Then
               'SI YA LLEGO AL FINAL GRABA UN REGISTRO MAS CON EL TOTAL Y CUENTA DEL BANCO
               IDEBE# = 0
               IHABE# = ACUMU_IDEBE#
               CTA_GASTO% = CUEII% 'CUENTA DEL BANCO QUE SE ESTA CONCILIANDO
               COCUE$ = CODCUE$(CUEII%) 'CUENTA EXTERNA DEL BANCO QUE SE ESTA CONCILIANDO
               GoTo 40
             End If
      End With
  Next REG1&
    '
    DETATMP.Close
    Set DETATMP = Nothing
    '
    'ACTUALIZO EL TOTAL EN DETASIEN
    CONDI$ = "NuIDAsien = '" & NUASI$ & "'"
    Call ACTUREGISTRO("ENCABEASIEN", "DebeAsien", CONDI$, FORMATNUM$(ACUMU_IDEBE#, "F16.5"), REGAF&)
    '
    'CAMBIO EL STADO EN CONCILIA DE LO CONCILIADO
    CONDI$ = "CODICONCI = '" & CODICONCI$ & "'"
    Call ACTUREGISTRO("CONCILIA", "STATUS", CONDI$, 1, REGAF&, "NOMESS")
    '
    If FIN& > 1 Then
     Call COMUNI("Se Ha Completado la Conciliacion Bancaria")
    End If
    '
99  Command6.Enabled = True

End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    'valido que haya un item seleccionado
    HAY_SELECCION_1% = 0
    With CONCIMAN07.MSF3
    For I& = 1 To .Rows - 1
          If .TextMatrix(I&, 0) = "*" Then
             HAY_SELECCION_1% = 1
           Exit For
          End If
        Next I&
    End With
    '
    If HAY_SELECCION_1% > 0 Then
       DEBEX# = XVALO(MSF3.TextMatrix(REGSEL3&, 5))
       HABEX# = XVALO(MSF3.TextMatrix(REGSEL3&, 6))
       MULITISEL09.Label9 = FORMATNUM$(DEBEX# - HABEX#, "F12.2")
       MULITISEL09.Show 1
    Else
       Call COMUNI("Falta Seleccionar Movimiento\Movimientos Bancarios (en el sistema)\Pendientes de Conciliación")
       Exit Sub
    End If
    

    Command7.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call APLICACIONSKINS(Me)
   '\\\OT-08-0990-OD-FIN///
   '
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "MM-00 ": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "N.O." 'NUMERO DE ORDEN
   NUCOLU% = 9: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "" 'CUENTA CONTABLE DE GASTOS
   '
   MSF2.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "MM-00 ": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "N.O." 'NUMERO DE ORDEN
   NUCOLU% = 9: TEPRUEBA = "": MSF2.ColWidth(NUCOLU% - 1) = 0: MSF2.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "MM-00 ": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "N.O." 'CUENTA CONTABLE DE GASTOS
   '
   MSF3.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 9: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENTE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF2.TextMatrix(0, NUCOLU% - 1) = "" 'CUENTA CONTABLE DE GASTOS
   '
   REGSEL1& = 0
   REGSEL2& = 0
   REGSEL3& = 0

End Sub



Private Sub MSF1_Click()
   '
   If MSF1.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   I& = MSF1.MouseRow
   J& = MSF1.MouseCol
   If I& = 0 Or I& > MSF1.Rows Then Exit Sub

   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, I&, "*", 1, REGSEL1&)
   REGSEL1& = I& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,
End Sub

Private Sub MSF2_Click()
   '
   If MSF2.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   I& = MSF2.MouseRow
   J& = MSF2.MouseCol
   If I& = 0 Or I& > MSF2.Rows Then Exit Sub

   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF2, &HC0E0FF, I&, "*", 1, REGSEL2&)
   REGSEL2& = I& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,

End Sub
Private Sub MSF3_Click()
   '
   If MSF3.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   I& = MSF3.MouseRow
   J& = MSF3.MouseCol
   If I& = 0 Or I& > MSF3.Rows Then Exit Sub
   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, I&, "*", 1, REGSEL3&)
   REGSEL3& = I& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,

End Sub

Sub COLOREAR_GRILLASQL_CON_MARCA(MSF As MSFlexGrid, COLOR As Variant, REG&, Optional MARCA$, Optional UNICASELECCION%, Optional REGISTRO_a_BLANQUEAR&)
   'CON UN CLIC PINTA TODO EL REGISTRO DEL COLOR PARAMETRIZADO Y LE PONE UN ASTERIZCO EN COLUMNA 0
   'EN EL SEGUNDO CLIC SI YA ESTA PINTADA BORRA EL COLOR Y TAMBIEN EL ASTERIZCO
   '
   FILA_MARCADA% = 0
   With MSF
     
        ASTERI$ = ""
        If TRIM$(MARCA$) <> "" Then ASTERI$ = MARCA$
        '
        'SI ES UNICA SELECCION CUANDO MARCA UNA BORRA LA ANTERIOR PARA QUE SIEMPRE QUEDE SOLO UNA MARCADA
        If UNICASELECCION% > 0 Then
          'SI TIENE VALOR EL 'REGISTROABLANQUEAR' BLANQUEA EL REGISTRO PASADO COMO PARAMETRO
          If REGISTRO_a_BLANQUEAR& > 0 And REGISTRO_a_BLANQUEAR& < MSF.Rows Then
             MSF.Row = REGISTRO_a_BLANQUEAR&
             For J& = 1 To MSF.Cols - 1
                 .Col = J&
                 .CellBackColor = vbWhite
                 If .TextMatrix(REGISTRO_a_BLANQUEAR&, 0) = ASTERI$ Then FILA_MARCADA% = 1
                 .TextMatrix(REGISTRO_a_BLANQUEAR&, 0) = ""
             Next J&
             If FILA_MARCADA% = 1 And REGISTRO_a_BLANQUEAR& = REG& Then Exit Sub
          End If
          GoTo 30 'VA A PINTAR LA CELDA
          
'ESTE BLOQUE DE SOLO UNICA ESTARIA QUEDANDO INUTIL CONSIDERANDO QUE EN ESTE FORMULARIO
'SE ESTA GURDANDO DE MANERA GLOBAL EL REGISTRO SELECCIONADO Y CUANDO CLICKEA EN OTRO
'LIMIPIA SOLAMENTE ESTE.
          For I& = 1 To .Rows - 1
            .Row = I&
            If MSF.TextMatrix(I&, 0) = MARCA$ Then
                For J& = 1 To MSF.Cols - 1
                  If I& <> REG& Then
                    .Col = J&
                    .CellBackColor = vbWhite
                    .TextMatrix(I&, 0) = ""
                  End If
                Next J&
                GoTo 30

            End If
          Next I&
        End If
'FIN SOLO UNICA
        '
        'PINTA LA FILA SELECCIONADA
        
30      If REG& < 1 Then Exit Sub
        .Row = REG&
        .Col = 1
        '
        'SI ES EL MISMO COLOR LA BLANQUEA Y QUITA EL AZTERIZCO CAMBIANDO EL VALOR DE LAS VARIABLES
        If .CellBackColor = COLOR Then
           COLOR = vbWhite: ASTERI$ = ""
        End If
        
        For ICOL% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .Col = ICOL%
          .CellBackColor = COLOR
        Next ICOL%
        '
        'AGREGA O QUITA EL AZTERIZCO
        .TextMatrix(REG&, 0) = ASTERI$
        
   End With
End Sub


Function VALIDA_SI_ES_GASTO%(CODICONCI$, CONCEPTOMSF2$)

    'LA FUNCION VALIDA SI ES UN GASTO PARA ELLO RECORRE LA LISTA DE
    'REFERENCIA DE GASTOS DE LA TABLA GASEXTRA Y SI EXISTE COMPARADO CON EL PASADO COMO PARAMETRO
    'GRABA EN EL CONCILIA (LOS QUE POSEAN ESTE TEXTO ) EL NUMER DE CUENTA Y EL COMPROBANTE GASTOS

    CUEI% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUENTA_SELECCIONADA$ = CODCUE$(CUEI%)
    VALIDA_SI_ES_GASTO% = 0
    '
    RFF$ = RECFILT$("GASEXTRA", 1, AJUSTI$(CUENTA_SELECCIONADA$, 12))
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("GASEXTRA", REG&)
        CONCEP$ = TRIM$(Mid$(XX$, 13, 128))
        CUENTACONT% = CUEINT%(TRIM$(Mid$(XX$, 141, 12)))
        If InStr(UCase$(CONCEP$), UCase$(TRIM$(CONCEPTOMSF2$))) > 0 Then
          CONDI$ = "(COMPROBANTE = '' OR COMPROBANTE IS NULL) AND CONCEPTO LIKE '%" & CONCEP$ & "%' AND STATUS < 1 "
          CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
          Call ACTUREGISTRO("CONCILIA", "CUECONTAI", CONDI$, CUENTACONT%, REGAF&, "NOMESS")
          Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, "GASTOS", REGAF&, "NOMESS")
          VALIDA_SI_ES_GASTO% = 1
        End If
    Next U&

End Function

   Sub GRABA_COMPROBANTE_EN_CONCILIA(COMPROBANTE$, FE%, ORDEN&)
       CONDI$ = "ORDCONCILIA = " & ORDEN&
       CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
       Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, COMPROBANTE$, REGAF&, "NOMESS")
       Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI$, FETEXCOR1$(FE%), REGAF&, "NOMESS")
   End Sub
   
Sub ORDENA_MSF2()
    '
    SQLX$ = "SELECT * FROM CONCILIA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICONCI = '" & CODICONCI$ & "'"
    SQLX$ = SQLX$ + " AND (COMPROBANTE = '' OR COMPROBANTE IS NULL)"
    SQLX$ = SQLX$ + " AND status < 1"
    SQLX$ = SQLX$ + " ORDER BY ORDCONCILIA"
    '
    CONCIMAN07.MSF2.Rows = 1
    '
    CONCIMAN07.MSF2.Redraw = False
    Dim RSTTMP2 As ADODB.Recordset
    Set RSTTMP2 = READSET(SQLX$)
    With RSTTMP2
      Do While Not .EOF
         IMPODEBEX# = XVALO(!DEBE)
         IMPOHABEX# = XVALO(!HABER)
         CUEII% = XVALO(!CUECONTAI)
         IDREFE$ = SAFETEXT$(ID_REFERENCIA)
         FECONTA% = CVINT(!FECHACOMPROBANTE)
         FECH1 = FETEXCOR1$(FECONTA%)
         FECH_EXTR% = CVINT(!FECHA)
         CONCEP$ = SAFETEXT$(!CONCEPTO)
         REFECON$ = SAFETEXT$(!ID_REFERENCIA)
         COMPRBTE$ = SAFETEXT$(!COMPROBANTE)
         NUORD% = XVALO(!ORDCONCILIA)
         DETABLA$ = SAFETEXT$(!COMPROB_DE_TABLA)
         With CONCIMAN07.MSF2
            REG& = REG& + 1
            .Rows = REG& + 1
            .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
            .TextMatrix(REG&, 2) = COMPRBTE$
            .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
            .TextMatrix(REG&, 4) = CONCEP$
            .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
            .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
            .TextMatrix(REG&, 7) = NUORD%
            .TextMatrix(REG&, 8) = DETABLA$
                TX$ = FILTERGETRECORD$("GASEXTRA", 2, AJUSTI$(CONCEP$, 128))
                CODICUENTA$ = TRIM$(Mid$(TX$, 141, 12))
            .TextMatrix(REG&, 9) = CODICUENTA$
         End With
      .MoveNext
      Loop
    End With
    CONCIMAN07.MSF2.Redraw = True
    RSTTMP2.Close
    Set RSTTMP2 = Nothing

End Sub

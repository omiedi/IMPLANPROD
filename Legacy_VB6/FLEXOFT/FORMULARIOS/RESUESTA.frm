VERSION 5.00
Begin VB.Form RESUESTA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes Estadísticos de Ventas"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   33
      Top             =   3360
      Width           =   4110
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   525
         ScaleHeight     =   75
         ScaleWidth      =   3000
         TabIndex        =   34
         Top             =   420
         Width           =   3060
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   27
      Top             =   2415
      Width           =   4110
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Impresión"
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
         Index           =   8
         Left            =   2415
         TabIndex        =   29
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pantalla"
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
         Index           =   4
         Left            =   525
         TabIndex        =   28
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   8820
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   2415
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resúmenes de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   4515
      TabIndex        =   19
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   30
         Top             =   1260
         Width           =   2325
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
            Left            =   -20
            Picture         =   "RESUESTA.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Resúmenes de Entregas por Cliente y Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de En- tregas por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   32
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   24
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command14 
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
            Left            =   -20
            Picture         =   "RESUESTA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Consulta Analítica Entregas por Cliente "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha de Deta- lle por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   26
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   20
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command13 
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
            Left            =   0
            Picture         =   "RESUESTA.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Listado Ventas Totales por Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Ven- tas por Producto"
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
            Left            =   735
            TabIndex        =   22
            Top             =   0
            Width           =   1620
         End
      End
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8820
      Picture         =   "RESUESTA.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   8820
      Picture         =   "RESUESTA.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Terminar - Salir de esta Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ranking de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4515
      TabIndex        =   10
      Top             =   2415
      Width           =   4110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   11
         Top             =   840
         Width           =   2430
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "RESUESTA.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Ranking de Ventas por Articulo por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Ven- tas por Articulo"
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
            Left            =   630
            TabIndex        =   13
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   14
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command6 
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
            Left            =   0
            Picture         =   "RESUESTA.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Ranking de Ventas por Cliente por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking por Cliente"
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
            Left            =   630
            TabIndex        =   16
            Top             =   105
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Análisis Estadistico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   1
         Top             =   1260
         Width           =   2220
         Begin VB.CommandButton Command5 
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
            Left            =   0
            Picture         =   "RESUESTA.frx":15F6
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Análisis de Cobertura de Stocks"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis de Co- bertura de Stocks"
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
            Left            =   630
            TabIndex        =   3
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "RESUESTA.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Análisis de Margen Bruto de Ventas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis Margen Bruto de Ventas"
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
            Left            =   735
            TabIndex        =   6
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   7
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "RESUESTA.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Resumen Estadistico de Ventas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen  Esta- dístico de Ventas"
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
            Left            =   735
            TabIndex        =   9
            Top             =   0
            Width           =   1905
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8260
      Picture         =   "RESUESTA.frx":1F14
      Top             =   3580
      Width           =   2010
   End
End
Attribute VB_Name = "RESUESTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
Private Sub Command1_Click()
    Command1.Enabled = False
    Call RESUENCLI
    Command1.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call REVEPRO
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call FICHENCLI
    Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    Call HELPONLINE("RESUESTA")
    Command29.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = EMPREAC$
    If InStr(UCase$(EPAC$), "SABO") > 0 Then
        'Call CONECRUN("#GESTAVE", "Resumen Estadistico de Ventas")
        Call GESTAVEN
      ElseIf EXISTE%("ESTAVE02.EXE") > 0 Then
        Call CONECRUN("ESTAVE02", "Estadisticas de Ventas")
      ElseIf EXISTE%("ESTAVEN.EXE") > 0 Then
        Call CONECRUN("ESTAVEN", "Estadisticas de Ventas")
      Else
        Call OPNOIN("")
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If InStr(UCase$(EMPRELI$), "SABO") > 0 Then
        PROCESX$ = "#MABRUVE"
        If EXISTE%("ANAMABRU.EXE") > 0 Then
            PROCESX$ = "ANAMABRU"
        End If
        Call CONECRUN(PROCESX$, "Analisis Margen Bruto de Ventas")
      Else
        Call OPNOIN("")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
'MsgBox "LIMITAR A SABO"
    EPAC$ = EMPREAC$
    If InStr(UCase$(EPAC$), "SABO") > 0 Then
        FORCOBERSTO.Show 1
      Else
       Call OPNOIN("")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call RANKCLI
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call RANKPIE
    Command7.Enabled = True
End Sub

Sub GESTAVEN()
    '
    'Dim RMVS As String * 128
    CORRESP$ = TRIM$(CORRAR$("ESTAVEN"))
    If CORRESP$ = "" Then GoTo 200
    '
    TXT$ = "El Archivo de Estadística de Ventas Existente\"
    TXT$ = TXT$ + "Corresponde al Período " + CORRESP$ + ".\"
    FEX = FEGEAR%("ESTAVEN")
    TXT$ = TXT$ + "Generado el " + FECHATEX$(FEX) + " a las " + HORATEX$(HOGEAR%("ESTAVEN")) + " hs.\"
    TXT$ = TXT$ + "\Consultar\Re-Generar\Control Exportación\Análisis Contable\Abandonar"
    X% = COMALTER%(TXT$)
    On X% GoTo 300, 200, 400, 330, 390
    '
200 Call BLANARCH("ITMALCLA")
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = TRIM$(VALACT1$("SELMESA"))
    CORRESP$ = MESANO$(MESASEL$, DES%, HAS%)
    If CORRESP$ = "" Then GoTo 390
    '
    FEX = DES%: DESX$ = FECHATEX$(FEX)
    FEX = HAS%: HASX$ = FECHATEX$(FEX)
    CORR1$ = TRIM$(CORRESP$)
    '
    Screen.MousePointer = 11
    GoSub 130                               ' inicializar RESTAVEN
    GoSub 110                               ' seleccionar facturas
    If CAFAC% < 1 Then                           '       del periodo elegido
       Call COMUNI("No hay Facturación\en el Período Elegido")
       GoTo 200
    End If
    '
    Dim NCI%(16384), TI%(16384), NRO&(16384), FFI%(16384) ' dimensionamiento
    Dim IFAC#(16384), IDET#(16384), SIFAC%(16384)
    Dim OTMO#(5)
    GoSub 111               ' cargar facturas a vector
    GoSub 112               ' coeficiente de precios
    '
    Call BLOQARCH("ESTAVE")
    Call BLOQARCH("ESTAVEN")
    Call BLANARCH("ESTAVEN")
    Call BLANARCH("VENOCLA")
    '
    FIN& = NUMAS%
    'XX% = FILENBR%("MASTER")
    'FIELD #XX%,110 AS MX1$,4 AS MX2$
    'Dim MX0 As String * 128
    '
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       'Get #XX%, I& + 1, MX0$
       MX0$ = REGLEIDO$("MASTER", I& + 1)
       COU = CVS(Mid$(MX0$, 111, 4))
       XS$ = String$(47, 0) + Chr$(1)
       Call REPLA(XS$, MKS$(COU), 21, 4)
       Call GRAREG("ESTAVE", XS$, I&)
    Next I&
    '
    Call CIERRARCH("MASTER")
    XX% = 0
    '
'CI000% = CODINT%("S07216.01N")

    SUNOCLA# = 0
    'FINUMO% = FILENBR%("MOVISTO")
    FIN& = ULTRE&
    For I& = PRIRE& To FIN&
       Call TRACE(20, I& - PRIRE&, FIN& - PRIRE&)
       'Get #FINUMO%, I&, RMVS$
       XS$ = REGLEIDO$("MOVISTO", I&)
       'XS$ = RMVS$
       FF% = CVI(Left$(XS$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           COMO$ = UCase$(Mid$(XS$, 33, 2))
           If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
             NC% = CVI(Mid$(XS$, 84, 2))
             NRO1& = Val(Mid$(XS$, 36, 7))
             GoSub 125
             If COEF > 1.005 Or ENCU% = 0 Then
               DOPRI$ = Mid$(XS$, 33, 10)
               If DOPRI$ <> DOPRIA$ Then
                 'Call COMUNI("Se ha Detectado una Inconsistencia\en el Documento " + DOPRI$)
                 DOPRIA$ = DOPRI$
               End If
             End If
             '
             CI% = CVI(Mid$(XS$, 3, 2))
             If CI% < 1 Or CI% > NUMAS% Then
               DOCU$ = Mid$(XS$, 33, 10)
               CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
               PRE# = CVD(Mid$(XS$, 73, 8)) * COEF
               TOT# = PRE# * CAN#
               MRX% = Int(NC% / 1000): If MRX% < 1 Or MRX% > 5 Then MRX% = 1
               GoSub 150                      ' grabar en VENOCLA.DAT
               GoTo 249
             End If
             '
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
                 PRE# = CVD(Mid$(XS$, 73, 8)) * COEF
                 TOT# = PRE# * CAN#
                 '
                 REGMA& = CI% + 1
                 MX$ = REGLEIDO$("MASTER", REGMA&)
                 TIMAS$ = Mid$(MX$, 69, 2)
                 Call VERIMPUCO(CI%, TIMAS$)
                 'MERCA% = Int(NC% / 1000)   ' esto es la clasificacion de Mingote
                 '
                 'Nueva clasificacion coherente entre detalle y
                 'Resumen estadistico - suprimida 14/12/05
                 'CLAMERCA% = Asc(Mid$(MX$, 71, 1))
                 'If CLAMERCA% = 1 Then
                 '    MERCA% = 2
                 '  ElseIf CLAMERCA% = 2 Then
                 '    MERCA% = 3
                 '  ElseIf CLAMERCA% = 3 Then
                 '    MERCA% = 5
                 '  Else
                 '    Call MENSERR(24, "Codigo " + CODEXT$(CI%) + "Mal Clasificado por Mercado")
                 '    MERCA% = 4
                 'End If
                 '
                 'clasificacion por grupo contable del cliente
                 MERCA% = GRUCOCLI%(NC%)
AMERC1% = Asc(Mid$(MX$, 71, 1))
If AMERC1% = 1 And MERCA% > 2 Then
   TTXX$ = "Codigo " + TRIM$(CODEXT$(CI%)) + " Probablemente mal Clasificado"
   TTXX$ = TTXX$ + "\  \Cliente " + RASOCLI$(NC%) + " --> " + DEGRUVE$(NC%)
   TTXX$ = TTXX$ + "\Material: " + TRIM$(CODEXT$(CI%)) + " --> " + ECOARCH$("CLASI003", MKI$(AMERC1%))
   Call MENSERR(24, TTXX$)
End If
If MERCA% = 5 And AMERC1% <> 3 Then
   TTXX$ = "Codigo " + TRIM$(CODEXT$(CI%)) + " Probablemente mal Clasificado"
   TTXX$ = TTXX$ + "\  \Cliente " + RASOCLI$(NC%) + " --> " + DEGRUVE$(NC%)
   TTXX$ = TTXX$ + "\Material: " + TRIM$(CODEXT$(CI%)) + " --> " + ECOARCH$("CLASI003", MKI$(AMERC1%))
   Call MENSERR(24, TTXX$)
End If
If AMERC1% = 2 And MERCA% <> 3 And MERCA% <> 4 Then
   TTXX$ = "Codigo " + TRIM$(CODEXT$(CI%)) + " Probablemente mal Clasificado"
   TTXX$ = TTXX$ + "\  \Cliente " + RASOCLI$(NC%) + " --> " + DEGRUVE$(NC%)
   TTXX$ = TTXX$ + "\Material: " + TRIM$(CODEXT$(CI%)) + " --> " + ECOARCH$("CLASI003", MKI$(AMERC1%))
   Call MENSERR(24, TTXX$)
End If


'If CI% = CI000% Then
'AAA = BBB
'End If
                 GoSub 131
                 If Abs(TOT#) > 0.005 Then GoSub 160
                 '
                 If CCL% <> 4 Then
                   If CCL% = 0 Then
                      MRX% = MERCA%
                      DOCU$ = Mid$(XS$, 33, 10)
                      GoSub 150                  ' grabar en VENOCLA.DAT
                      GoTo 249
                    Else
                      Call COMUNI("Inconsistencia de Clasificacion.\Consulte Soporte de Sistemas")
                   End If
                   GoTo 249
                 End If
                 '
                 REGES& = CI%
                 Y$ = REGLEIDO$("ESTAVE", REGES&)
                 CAN# = CAN# + CVD(Left$(Y$, 8))
                 TOT# = TOT# + CVD(Mid$(Y$, 9, 8))
                 PROME = 0: If Abs(CAN#) > 0.1 Then PROME = TOT# / CAN#
                 COU = CVS(Mid$(Y$, 21, 4))
                 MATOT# = TOT# - CAN# * COU
                 MAPOR = 0: If Abs(TOT#) > 0.1 Then MAPOR = 100 * MATOT# / TOT#
                 '
                 Call REPLA(Y$, MKD$(CAN#), 1, 8)
                 Call REPLA(Y$, MKD$(TOT#), 9, 8)
                 Call REPLA(Y$, MKS$(PROME), 17, 4)
                 Call REPLA(Y$, MKS$(MAPOR), 25, 4)
                 Call REPLA(Y$, MKD$(MATOT#), 29, 8)
                 Call GRAREG("ESTAVE", Y$, REGES&)
                 '
               End If
             End If
           End If
         End If
       End If
249 Next I&
    '
' AQUI APARENTEMENTE COINCIDEN EL TOTAL DE "ESTAVE" Y EL DE "RESTAVEN"
    FIN& = CAFAC%
    For T% = 1 To CAFAC%
      TL& = T%
      Call TRACE(20, TL&, FIN&)
      If Abs(IDET#(T%)) < 0.5 Then
        If Abs(IFAC#(T%)) > 0.1 Then
          TTXI% = TI%(T%): If TTXI% = 7 Then TTXI% = 0
          SIG% = 1: If TI%(T%) > 3 Then SIG% = (-1)
          DOCU$ = FORMDOC$(NRO&(T%), TTXI%, 1, 1)
          TOT# = (Abs(IFAC#(T%))) * SIG%
          MRX% = Int(NCI%(T%) / 1000): If MRX% < 1 Or MRX% > 5 Then MRX% = 1
          CI% = 0
          GoSub 150                      ' grabar en VENOCLA.DAT
        End If
      End If
    Next T%
    '
    FIN& = NUINV%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
        IX$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(IX$, 17, 2))
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            REGMA& = CI% + 1
            ES$ = REGLEIDO$("ESTAVE", REGMA& - 1)
            Call GRAREG("ESTAVE", String$(36, 0), REGMA& - 1)
            '
            If Abs(CVD(Left$(ES$, 8))) <= 0.05 Then
              If Abs(CVD(Mid$(ES$, 9, 8))) >= 0.005 Then
                Call MENSERR(25, "Código '" + TRIM$(CODEXT$(CI%)) + "' Tiene Importe\Sin Cantidad")
              End If
            End If
            '
            If (Abs(CVD(Left$(ES$, 8))) > 0.05 Or Abs(CVD(Mid$(ES$, 9, 8))) > 0.05) Then
              MX$ = REGLEIDO$("MASTER", REGMA&)
              CLASI$ = Mid$(MX$, 69, 8)
              '
              Z$ = REGBLAN$("ESTAVEN")
              Call REPLA(Z$, CLASI$, 1, 8)
              Call REPLA(Z$, MKI$(CI%), 9, 2)
              Call REPLA(Z$, ES$, 11, 36)       ' CANTIDAD E IMPORTE
              Call REGAPP("ESTAVEN", Z$)
              '
            End If
          End If
        End If
     Next I&
     '
     ' esto es para control
     For KU& = 1 To ULTREG&("ESTAVE")
       ES$ = REGLEIDO$("ESTAVE", KU&)
       If ES$ <> String$(36, 0) Then
         CI% = KU&
         AA1$ = CODEXT$(CI%)
         CANX# = CVD(Left$(ES$, 8))
         IMPX# = CVD(Mid$(ES$, 9, 8))
         If Abs(CANX#) >= 0.005 Or Abs(IMPX#) >= 0.005 Then
           Call MENSERR(24, "Movimiento Sobre Código '" + TRIM$(AA1$) + "'\no Registrado en Lista Invertida")
         End If
       End If
     Next KU&
     '
     GoSub 170          ' AGREGAR DEBITOS Y CREDITOS
     '
     GoSub 140
     Call CIERRARCH("*.*")
     Call GRARGEN("ESTAVEN", TRIM$(CORRESP$))
     '
     GoSub 180     ' generar archivo de analisis CONTABLE ANAIMPUC
     '
     If ULTREG&("ITMALCLA") > 0 Then
       If COMALTER%("Se Encontraron Items Aparentemente Mal Clasificados.\   \¿ Desea Emitir un Listado de Control ?\\Si, Listar\No,Continuar") = 1 Then
         Call FILESORT("ITMALCLA", "", 1, 1)
         Call CONECRUN("*ITMALCLA", "Items Mal Calsificado")
       End If
     End If
     '
300 Screen.MousePointer = 1
    On ALTERNA%("Estadísticas de Detalle\Resumen Estadístico") GoTo 310, 320
    GoTo 390
    '
310 Call COSTAVEN
    Exit Sub
    '
320 Call HEADERS("RESTAVEN", "")
    Call HEADERS("RESTAVEN", "Corresponde a: " + CORRESP$)
    Call HEADERS("RESTAVE", "")
    Call HEADERS("RESTAVE", "Corresponde a: " + CORRESP$)
    Call HEADERS("VENOCLA", "")
    Call HEADERS("VENOCLA", "Corresponde a: " + CORRESP$)
    Call FINAL("*RESTAVE")
    Exit Sub
    '
330 Call FILESORT("ANAIMPUC", "", 1, 5, "ASC")
    Call FINAL("*ANAIMPUC")
    Exit Sub
    '
390 Exit Sub
     '

'*****************************************************************************

110 ' Call BLIMESS("Revisando Archivo de Facturacion ...")
    '
    FIN& = ULTREG&("FACTUR")
    CAFAC% = 0
    PRIRE& = 999999: ULTRE& = 0
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(XS$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CAFAC% = CAFAC% + 1
          If I& < PRIRE& Then PRIRE& = I&
          If I& > ULTRE& Then ULTRE& = I&
        End If
      End If
    Next I&
    'Call BLIMESS("-")
    '
    If CAFAC% < 1 Then
      Call COMUNI("Proceso no Realizable !!!\No hay Facturas Registradas\en el Periodo Elegido.")
    End If
    '
    'If 34 * CAFAC% > FRE(-1) - 1000 Then
    '  Call COMUNI("Proceso no Realizable !!!\Memoria Insuficiente.\-\Consulte Soporte de Sistemas.")
    '  Call FINAL("")
    'End If
    '
Return
'
'
111 CAFAC% = 0
    SUFAC# = 0
    For I& = PRIRE& To ULTRE&
      Call TRACE(20, I& - PRIRE&, ULTRE& - PRIRE&)
      XS$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(XS$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CAFAC% = CAFAC% + 1
          NCI%(CAFAC%) = CVI(Left$(XS$, 2))
          FFI%(CAFAC%) = FF%
          TI%(CAFAC%) = Asc(Mid$(XS$, 13, 1))
          If TI%(CAFAC%) < 0 Or TI%(CAFAC%) > 6 Then TI%(CAFAC%) = 2 ' por default factura
          NRO&(CAFAC%) = CVS(Mid$(XS$, 3, 4))
          '
          If TI%(CAFAC%) = 0 Then                     ' contempla ajustes de saldos
            If CVD(Mid$(XS$, 15, 8)) < 0 Then
              TI%(CAFAC%) = 7
            End If
          End If
          '
          IFAC#(CAFAC%) = Abs(CVD(Mid$(XS$, 15, 8)))
          '
          SIG% = 1: If TI%(CAFAC%) > 3 Then SIG% = (-1)
          SIFAC%(CAFAC%) = SIG%
          SUFAC# = SUFAC# + IFAC#(CAFAC%) * SIG%
        End If
      End If
    Next I&
    'Call BLIMESS("-")
    '
Return
'
'
112 PRIRE& = 999999: ULTRE& = 0
    FIN& = ULTREG&("MOVISTO")
    'FINUMO% = FILENBR%("MOVISTO")
    PRIREMO& = 0
    For I& = 1 To FIN&
      Call TRACE(20, I& - PRIREMO&, FIN& - PRIREMO&)
      'Get #FINUMO%, I&, RMVS$
      XS$ = REGLEIDO$("MOVISTO", I&)
      'XS$ = RMVS$
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If PRIREMO& = 0 Then PRIREMO& = I&
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
            TP1% = 2: If COMO$ = "NC" Then TP1% = 6
            If COMO$ = "ND" Then TP1% = 3
            NRO1& = Val(Mid$(XS$, 36, 7))
            CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
            PRE# = CVD(Mid$(XS$, 73, 8))
            TOT# = Abs(PRE# * CAN#)
            If I& < PRIRE& Then PRIRE& = I&
            If I& > ULTRE& Then ULTRE& = I&
            GoSub 120
          End If
        End If
      End If
    Next I&
    '
Return
'
'
120 For J% = 1 To CAFAC%
      If FFI%(J%) = FF% Then
        If TI%(J%) = TP1% Then
          If NRO&(J%) = NRO1& Then
            IDET#(J%) = IDET#(J%) + TOT#
            GoTo 121
          End If
        End If
      End If
    Next J%
    '
121 Return
'
'
125 COEF = 1: ENCU% = 0
    For J% = 1 To CAFAC%
      If FFI%(J%) = FF% Then
        If NRO&(J%) = NRO1& Then
          ENCU% = 1
          If IDET#(J%) > 0 Then
            COEF = IFAC#(J%) / IDET#(J%)
          End If
          GoTo 129
        End If
      End If
    Next J%
    '
    
129 Return
'
'
130 Call BLANARCH("RESTAVEN")
    For I& = 1 To ULTREG&("CLASI001") + 1
      If I& = ULTREG&("CLASI001") + 1 Then
          COTIPO$ = Chr$(255)
          DETIPO$ = AJUSTI$("TOTAL", 15)
        Else
          XS$ = REGLEIDO$("CLASI001", I&)
          COTIPO$ = Left$(XS$, 1)
          DETIPO$ = Mid$(XS$, 2, 15)
      End If
      '
      If TRIM$(DETIPO$) <> "" Then
        Z$ = String$(2, 255) + Space$(34) + String$(60, 0)
        Call REGAPP("RESTAVEN", Z$)
        For JJ& = 1 To ULTREG&("CLASI002")
          Y$ = REGLEIDO$("CLASI002", JJ&)
          COMARCA$ = Left$(Y$, 1)
          DEMARCA$ = Mid$(Y$, 2, 15)
          If TRIM$(DEMARCA$) <> "" Then
            Z$ = COTIPO$ + COMARCA$ + DETIPO$ + DEMARCA$ + "  U." + String$(60, 0)
            Call REGAPP("RESTAVEN", Z$)
            Z$ = COTIPO$ + COMARCA$ + DETIPO$ + DEMARCA$ + "  $ " + String$(60, 0)
            Call REGAPP("RESTAVEN", Z$)
            Z$ = COTIPO$ + COMARCA$ + DETIPO$ + DEMARCA$ + "    " + String$(60, 0)
            Call REGAPP("RESTAVEN", Z$)
          End If
        Next JJ&
        Z$ = COTIPO$ + Chr$(255) + DETIPO$ + "SUBTOTAL         U." + String$(60, 0)
        If I& = ULTREG&("CLASI001") + 1 Then
          Call REPLA(Z$, "GENERAL   ", 18, 15)
        End If
        Call REGAPP("RESTAVEN", Z$)
        Z$ = COTIPO$ + Chr$(255) + DETIPO$ + "SUBTOTAL         $ " + String$(60, 0)
        If I& = ULTREG&("CLASI001") + 1 Then
          Call REPLA(Z$, "GENERAL   ", 18, 15)
        End If
        Call REGAPP("RESTAVEN", Z$)
        Z$ = String$(2, 255) + Space$(34) + String$(60, 0)
        Call REGAPP("RESTAVEN", Z$)
        '
      End If
      '
    Next I&
    '
Return
'
'
131 CCL% = 0
    If MERCA% < 1 Or MERCA% > 5 Then GoTo 135
    For UU& = 1 To ULTREG&("RESTAVEN")
      ZZ$ = REGLEIDO$("RESTAVEN", UU&)
      If Left$(ZZ$, 2) = TIMAS$ Then
        CANUE# = CAN# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        CATOT# = CAN# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(CANUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(CATOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, UU&)
        '
        ZZ$ = REGLEIDO$("RESTAVEN", UU& + 1)
        IMNUE# = TOT# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        IMTOT# = TOT# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(IMNUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(IMTOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, UU& + 1)
        '
sucan2# = sucan2# + CAN#
SUTOT2# = SUTOT2# + TOT#
        '
        CCL% = CCL% + 1
        GoTo 132
      End If
    Next UU&
    Call MENSERR(24, "Imposible Clasificar " + TRIM$(CODEXT$(CI%)) + "\con Tipo =" + Str$(Asc(Left$(TIMAS$, 1))) + "   y Marca =" + Str$(Asc(Mid$(TIMAS$, 2, 1))))
'MsgBox TTXXx$
    GoTo 135                             ' NO LO ENCONTRO
    '
132 For VV& = UU& + 3 To ULTREG&("RESTAVEN")
      ZZ$ = REGLEIDO$("RESTAVEN", VV&)
      If Left$(ZZ$, 2) = Left$(TIMAS$, 1) + Chr$(255) Then
        CANUE# = CAN# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        CATOT# = CAN# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(CANUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(CATOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, VV&)
        '
        ZZ$ = REGLEIDO$("RESTAVEN", VV& + 1)
        IMNUE# = TOT# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        IMTOT# = TOT# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(IMNUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(IMTOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, VV& + 1)
        '
        CCL% = CCL% + 1
        GoTo 133
      End If
    Next VV&
    GoTo 135
    '
133 For WW& = VV& + 3 To ULTREG&("RESTAVEN")
      ZZ$ = REGLEIDO$("RESTAVEN", WW&)
      If Left$(ZZ$, 2) = Chr$(255) + Mid$(TIMAS$, 2, 1) Then
        CANUE# = CAN# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        CATOT# = CAN# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(CANUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(CATOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, WW&)
        '
        ZZ$ = REGLEIDO$("RESTAVEN", WW& + 1)
        IMNUE# = TOT# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        IMTOT# = TOT# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(IMNUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(IMTOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, WW& + 1)
        '
        CCL% = CCL% + 1
        GoTo 134
      End If
    Next WW&
    GoTo 135
    '
134 For XXL& = WW& + 3 To ULTREG&("RESTAVEN")
      ZZ$ = REGLEIDO$("RESTAVEN", XXL&)
      If Left$(ZZ$, 7) = String$(2, 255) + "TOTAL" Then
        CANUE# = CAN# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        CATOT# = CAN# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(CANUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(CATOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, XXL&)
        '
        ZZ$ = REGLEIDO$("RESTAVEN", XXL& + 1)
        IMNUE# = TOT# + CVD(Mid$(ZZ$, 29 + 8 * MERCA%, 8))
        IMTOT# = TOT# + CVD(Mid$(ZZ$, 77, 8))
        Call REPLA(ZZ$, MKD$(IMNUE#), 29 + 8 * MERCA%, 8)
        Call REPLA(ZZ$, MKD$(IMTOT#), 77, 8)
        Call GRAREG("RESTAVEN", ZZ$, XXL& + 1)
        '
        CCL% = CCL% + 1
        GoTo 135
      End If
    Next XXL&
    '
135 Return
'
'
140 Call BLANARCH("RESTAVE")
    FIN& = ULTREG&("RESTAVEN")
    For U& = 1 To FIN&
      Z$ = REGLEIDO$("RESTAVEN", U&)
      If Mid$(Z$, 35, 2) = "U." Then
        If Abs(CVD(Mid$(Z$, 77, 8))) < 0.5 Then
          For V& = U& To U& + 2
            Call GRAREG("RESTAVEN", String$(96, 0), V&)
          Next V&
        End If
      End If
    Next U&
    '
    V& = 0
    For U& = 1 To FIN&
      Z$ = REGLEIDO$("RESTAVEN", U&)
      If Z$ <> String$(96, 0) Then
        Call REGAPP("RESTAVE", Z$)
      End If
    Next U&
    '
Return
'
'
150 VNC$ = String$(72, 0)
    Call REPLA(VNC$, DOCU$, 1, 18)
    Call REPLA(VNC$, MKI$(CI%), 19, 2)
    Call REPLA(VNC$, MKD$(TOT#), 13 + 8 * MRX%, 8)
    Call REPLA(VNC$, MKD$(TOT#), 61, 8)
    Call REGAPP("VENOCLA", VNC$)
    '
Return
    '
160 If CARUCO3% < 1 Then
      CARUCO3% = ULTREG&("CUECOVEN")
      CAGRUIM% = 0
      Dim VECLASICO$()
      Dim GRUCOD3%(), GRUCLI3%(), CUECON3%()
      ReDim GRUCOD3%(CARUCO3%), GRUCLI3%(CARUCO3%), CUECON3%(CARUCO3%)
      '
      FIN3& = CARUCO3%                     ' cargar vector
      For IL3& = 1 To FIN3&
        XX3$ = REGLEIDO$("CUECOVEN", IL3&)
        JJ3% = IL3&
        GRUCOD3%(JJ3%) = CVI(Left$(XX3$, 2))
        GRUCLI3%(JJ3%) = CVI(Mid$(XX3$, 3, 2))
        CUECON3%(JJ3%) = CVI(Mid$(XX3$, 5, 2))
      Next IL3&
    End If
    '
    If CI% <= 0 Then
        GRUVE3% = CI%
      Else
        GRUVE3% = GRUCOVEI%(CI%)
    End If
    GRUC3% = 0
      If NC% > 0 Then GRUC3% = GRUCOCLI%(NC%)
    TIPO3% = Asc(Left$(TIMAS$, 1))
    MARCA3% = Asc(Mid$(TIMAS$, 2, 1))
    '
    BUSQUE$ = MKI$(GRUVE3%) + MKI$(GRUC3%) + MKI$(TIPO3%) + MKI$(MARCA3%)
    For KKII% = 1 To CAGRUIM%
      CLASICO3$ = VECLASICO$(KKII%)
      If Mid$(CLASICO3$, 3, 8) = BUSQUE$ Then
        If CI% <= 0 Then
          If CVI(Left$(VECLASICO$(KKII%), 2)) <> CUECO3% Then GoTo 161
        End If
        TOTATO3# = CVD(Mid$(CLASICO3$, 11, 8)) + TOT#
        Call REPLA(CLASICO3$, MKD$(TOTATO3#), 11, 8)
        VECLASICO$(KKII%) = CLASICO3$
        GoTo 165
      End If
161 Next KKII%
    '
    If CI% > 0 Then
      CUECO3% = 0
      For KKJJ% = 1 To CARUCO3%
        If GRUCOD3%(KKJJ%) = GRUVE3% Then
          If GRUCLI3%(KKJJ%) = GRUC3% Then
            CUECO3% = CUECON3%(KKJJ%)
            GoTo 162
          End If
        End If
      Next KKJJ%
    End If
    '
162 CAGRUIM% = CAGRUIM% + 1
    ReDim Preserve VECLASICO$(CAGRUIM%)
    CLASICO3$ = String$(18, 0)
    Call REPLA(CLASICO3$, MKI$(CUECO3%), 1, 2)
    Call REPLA(CLASICO3$, BUSQUE$, 3, 8)
    Call REPLA(CLASICO3$, MKD$(TOT#), 11, 8)
    VECLASICO$(CAGRUIM%) = CLASICO3$
    '
165 Return
    '
    '
170 Dim CONCUENTA#(32767), ACTICU3%(32767)
    NUCUEN% = ULTREG&("CUENTAS")
    For KKI% = 1 To CAGRUIM%
      CUEI% = CVI(Left$(VECLASICO$(KKI%), 2))
      If CUEI% > 0 Then
        If CUEI% <= NUCUEN% Then
          CONCUENTA#(CUEI%) = CONCUENTA#(CUEI%) + CVD(Mid$(VECLASICO$(KKI%), 11, 8))
          ACTICU3%(CUEI%) = 1
        End If
      End If
    Next KKI%
    '
    MESB3$ = MESANO$(CORRESP$, DES3%, HAS3%)
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
      XX3$ = REGLEIDO$("ASIDIAR", IL&)
      If Left$(XX3$, 4) = TTXXY$ Then
        CUEI% = CVI(Mid$(XX3$, 5, 2))
        IDEB3# = CVD(Mid$(XX3$, 39, 8))
        If CUEI% > 0 Then
          If CUEI% <= NUCUEN% Then
            If ACTICU3%(CUEI%) > 0 Then
              CONCUENTA#(CUEI%) = CONCUENTA#(CUEI%) + IDEB3#
            End If
          End If
        End If
      End If
    Next IL&
    '
    For CUEI% = 1 To NUCUEN%
      If Abs(CONCUENTA#(CUEI%)) >= 0.005 Then
        CI% = 0
        NC% = 0
        TIMAS$ = String$(2, 0)
        CUECO3% = CUEI%
        TOT# = (-1) * CONCUENTA#(CUEI%)
        GoSub 160
      End If
    Next CUEI%
    '
Return
    '
    '
180 JJ& = 0
    REBLA$ = REGBLAN$("ANAIMPUC")
    For KKII% = 1 To CAGRUIM
      TTXX$ = REBLA$
      CUECO3% = CVI(Left$(VECLASICO$(KKII%), 2))
      GRUVE3% = CVI(Mid$(VECLASICO$(KKII%), 3, 2))
      GRUC3% = CVI(Mid$(VECLASICO$(KKII%), 5, 2))
      TIPO3% = CVI(Mid$(VECLASICO$(KKII%), 7, 2))
      MARCA3% = CVI(Mid$(VECLASICO$(KKII%), 9, 2))
      ITOT3# = CVD(Mid$(VECLASICO$(KKII%), 11, 8))
      '
      If GRUVE3% = 0 Then
           GRUPRO$ = "Z-No Codificados"
         Else
           GRUPRO$ = TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUVE3%)))
      End If
      '
      Call REPLA(TTXX$, CODCUE$(CUECO3%), 1, 16)
      Call REPLA(TTXX$, GRUPRO$, 17, 32)
      Call REPLA(TTXX$, TRIM$(ECOARCH$("GRUCOCLI", MKI$(GRUC3%))), 49, 32)
      Call REPLA(TTXX$, TRIM$(ECOARCH$("CLASI001", Chr$(TIPO3%))), 81, 24)
      Call REPLA(TTXX$, TRIM$(ECOARCH$("CLASI002", Chr$(MARCA3%))), 105, 24)
      Call REPLA(TTXX$, MKD$(ITOT3#), 129, 24)
      '
      JJ& = JJ& + 1
      Call GRAREG("ANAIMPUC", TTXX$, JJ&)
    Next KKII%
    '
    Call SETULTREG("ANAIMPUC", JJ&)
    Call CIERRARCH("ANAIMPUC")
    '
Return


400 Call CONEXPOR
    '
End Sub

Sub VERIMPUCO(CI%, TIMAS$)
   '
   Static REBLA$, MALCLA%(32767)
   If REBLA$ = "" Then
     REBLA$ = REGBLAN$("ITMALCLA")
   End If
   '
   If MALCLA%(CI%) > 0 Then Exit Sub ' PARA NO REPETIR CODIGOS
   '
   TIMA1$ = TRIM$(Str$(Asc(Left$(TIMAS$, 1)))) + "-" + TRIM$(Str$(Asc(Mid$(TIMAS$, 2, 1))))
   '
   GRUCORRE% = 0
   '
   If TIMA1$ = "1-1" Then GRUCORRE% = 1
   If TIMA1$ = "2-1" Then GRUCORRE% = 1
   If TIMA1$ = "4-1" Then GRUCORRE% = 1
   If TIMA1$ = "4-3" Then GRUCORRE% = 1
   If TIMA1$ = "4-4" Then GRUCORRE% = 2   ' RETENES KAKO
   If TIMA1$ = "4-5" Then GRUCORRE% = 1
   If TIMA1$ = "4-6" Then GRUCORRE% = 1
   If TIMA1$ = "11-5" Then GRUCORRE% = 1
   If TIMA1$ = "11-6" Then GRUCORRE% = 1
   '
   If TIMA1$ = "3-2" Then GRUCORRE% = 2
   If TIMA1$ = "4-2" Then GRUCORRE% = 2
   If TIMA1$ = "7-2" Then GRUCORRE% = 2
   If TIMA1$ = "11-2" Then GRUCORRE% = 2
   '
   If TIMA1$ = "6-2" Then GRUCORRE% = 3
   '
   If TIMA1$ = "6-5" Then GRUCORRE% = 4
   '
   If TIMA1$ = "5-2" Then GRUCORRE% = 6
   '
   GRUCLA% = GRUCOVEI%(CI%)
   If GRUCORRE% < 1 Then
     MALCLA%(CI%) = 1
     TTXX$ = "Imposible Determinar Clasificación Contable"
     TTXX$ = TTXX$ + "\Recomendable para Código " + TRIM$(CODEXT$(CI%))
     TTXX$ = TTXX$ + "\   \Cl.Técnica = " + TRIM$(ECOARCH$("CLASI001", Mid$(TIMAS$, 1, 1))) + " + " + TRIM$(ECOARCH$("CLASI002", Mid$(TMSX$, 2, 1)))
     Call MENSERR(24, TTXX$)
     '
     TTXYZ$ = REBLA$
     Call REPLA(TTXYZ$, MKI$(CI%), 1, 2)
     Call REPLA(TTXYZ$, TRIM$(Str$(Asc(Mid$(TIMAS$, 1, 1)))) + "-" + TRIM$(ECOARCH$("CLASI001", Mid$(TIMAS$, 1, 1))), 3, 24)
     Call REPLA(TTXYZ$, TRIM$(Str$(Asc(Mid$(TIMAS$, 2, 1)))) + "-" + TRIM$(ECOARCH$("CLASI002", Mid$(TIMAS$, 2, 1))), 27, 24)
     Call REPLA(TTXYZ$, TRIM$(Str$(GRUCLA%)) + "-" + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUCLA%))), 51, 24)
     Call REGAPP("ITMALCLA", TTXYZ$)
     '
     Exit Sub
   End If
   '
   If GRUCORRE% <> GRUCLA% Then
     TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + " Aparentemente Mal Clasificación Contable."
     TTXX$ = TTXX$ + "\Clas.Técnica = " + TRIM$(ECOARCH$("CLASI001", Mid$(TIMAS$, 1, 1)))
     TTXX$ = TTXX$ + " + " + TRIM$(ECOARCH$("CLASI002", Mid$(TIMAS$, 2, 1)))
     TTXX$ = TTXX$ + "\   \¿ Acepta Aplicar la Siguiente Clasificación Contable ?"
     TTXX$ = TTXX$ + "\" + TRIM$(Str$(GRUCORRE%)) + " - " + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUCORRE%)))
     TTXX$ = TTXX$ + "\\Si, Corregir\No, Dejar cómo está"
     '
'GoTo 20
     If COMALTER%(TTXX$) = 1 Then
       REGMA& = CI% + 1
       MX$ = REGLEIDO$("MASTER", REGMA&)
       Call REPLA(MX$, MKI$(GRUCORRE%), 125, 2)
       Call GRAREG("MASTER", MX$, REGMA&)
       Exit Sub
     End If
     '
20   MALCLA%(CI%) = 1
     TTXYZ$ = REBLA$
     Call REPLA(TTXYZ$, MKI$(CI%), 1, 2)
     Call REPLA(TTXYZ$, TRIM$(Str$(Asc(Mid$(TIMAS$, 1, 1)))) + "-" + TRIM$(ECOARCH$("CLASI001", Mid$(TIMAS$, 1, 1))), 3, 24)
     Call REPLA(TTXYZ$, TRIM$(Str$(Asc(Mid$(TIMAS$, 2, 1)))) + "-" + TRIM$(ECOARCH$("CLASI002", Mid$(TIMAS$, 2, 1))), 27, 24)
     Call REPLA(TTXYZ$, TRIM$(Str$(GRUCLA%)) + "-" + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUCLA%))), 51, 24)
     Call REPLA(TTXYZ$, TRIM$(Str$(GRUCORRE%)) + "-" + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUCORRE%))), 75, 24)
     Call REGAPP("ITMALCLA", TTXYZ$)
     '
   End If
   '
End Sub

Sub COSTAVEN()       ' consultas estadistica de ventas clasificadas
    '
    Static VDEF$
    Dim CLA%(10)
    '
    CORRESP$ = TRIM$(CORRAR$("ESTAVEN"))
    If CORRESP$ = "" Then
       Exit Sub
    End If
    '
    'Call BLOQARCH("TACRICLA")
150 'CLASI$ = String$(8, 0)
    'For I& = 1 To ULTREG&("TACRICLA")
    '  X$ = REGLEIDO$("TACRICLA", I&)
    '  Call REPLA(X$, Mid$(CLASI$, I&, 1), 49, 1)
    '  Call GRAREG("TACRICLA", X$, I&)
    'Next I&
    'Call CIERRARCH("TACRICLA")
    '
200 'VTB% = VENTABU%("CARCLASI/NC/TITUPAN=Estadísticas para " + CORRESP$)
    'If VTB% < 0 Then Exit Sub
    OBX$ = OBJPLA$("SELCLASITEC", VDEF$)
    If OBX$ = "" Then Exit Sub
    VDEF$ = OBX$
    '
300 OBX$ = OBX$ + String$(8, 0)
    TOCLA$ = ""
    For I& = 1 To 8
      X$ = REGLEIDO$("TACRICLA", I&)
      IX% = I&
      CLA%(IX%) = Asc(Mid$(OBX$, IX%, 1))
      If CLA%(IX%) > 0 Then
        CLAX$ = TRIM$(Str$(IX%)): While Len(CLAX$) < 3: CLAX$ = "0" + CLAX$: Wend
        UX& = CLA%(IX%)
        DECLA$ = TRIM$(REGLEIDO$("CLASI" + CLAX$, UX&))
        If DECLA$ <> "" Then
          If TOCLA$ <> "" Then TOCLA$ = TOCLA$ + " + "
          TOCLA$ = TOCLA$ + DECLA$
        End If
      End If
    Next I&
    '
400 Screen.MousePointer = 11
    Call BLANARCH("COSTAVE")
    Call BLANARCH("!COSTAVE")
    FIN& = ULTREG&("ESTAVEN")
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("ESTAVEN", I&)
      For J% = 1 To 8
        If CLA%(J%) > 0 Then
          If Asc(Mid$(X$, J%, 1)) <> CLA%(J%) Then GoTo 490
        End If
      Next J%
      '
      Call REGAPP("COSTAVE", X$)
      Call REGAPP("!COSTAVE", X$)
490 Next I&
    '
    Call CIERRARCH("COSTAVE")
    Call CIERRARCH("!COSTAVE")
    Screen.MousePointer = 1
    '
500 VTB% = VENTABU%("COSTAVE/TITUPAN=" + CORRESP$ + ": " + TOCLA$)
510 AX% = ALTERNA%("Volver\Imprimir\Otra Consulta")
    On AX% GoTo 500, 600, 200
    GoTo 390
    '
600 Call HEADERS("COSTAVE", "")
    Call HEADERS("COSTAVE", "Corresponde a: " + CORRESP$)
    Call HEADERS("COSTAVE", "Seleccion: " + TOCLA$)
    '
    Call FINAL("*LISTAVE")
    GoTo 200
    '
390 End Sub

Sub CONEXPOR()

    Dim NCI%(4096)
    Dim FFI%(4096)
    Dim TI%(4096)
    Dim NRO&(4096)
    Dim IFAC#(4096)
    Dim SIFAC%(4096)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    haynocla% = 0
    FIN& = ULTREG&("FACTUR")
    CAFAC% = 0
    PRIRE& = 999999: ULTRE& = 0
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(XS$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          VA% = Asc(Mid$(XS$, 14, 1))
          If VA% = 0 Then        ' SOLO FACTURAS DE EXPORTACION
            CAFAC% = CAFAC% + 1
            NCI%(CAFAC%) = CVI(Left$(XS$, 2))
            FFI%(CAFAC%) = FF%
            TI%(CAFAC%) = Asc(Mid$(XS$, 13, 1))
            If TI%(CAFAC%) < 0 Or TI%(CAFAC%) > 6 Then TI%(CAFAC%) = 2 ' por default factura
            NRO&(CAFAC%) = CVS(Mid$(XS$, 3, 4))
          End If
        End If
      End If
    Next I&
    '
112 PRIRE& = 999999: ULTRE& = 0
    Dim ACTICO%(32767)
    FIN& = ULTREG&("MOVISTO")
    PRIREMO& = 0
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("MOVISTO", I&)
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
            TP1% = 2: If COMO$ = "NC" Then TP1% = 6
            If COMO$ = "ND" Then TP1% = 3
            NRO1& = Val(Mid$(XS$, 36, 7))
            CI% = CVI(Mid$(XS$, 3, 2))
            For J% = 1 To CAFAC%
              If FFI%(J%) = FF% Then
                If TI%(J%) = TP1% Then
                  If NRO&(J%) = NRO1& Then
                    ACTICO%(CI%) = 1
                    GoTo 120
                  End If
                End If
              End If
            Next J%
          End If
        End If
      End If
120 Next I&
    '
    For KI& = 1 To NUMAS%
      If ACTICO%(KI&) = 1 Then
        XX$ = REGLEIDO$("MASTER", KI& + 1)
        If Asc(Mid$(XX$, 71, 1)) <> 3 Then
          CI% = KI&
          haynocla% = 1
          Call MENSERR(24, "Codigo " + TRIM$(CODEXT$(CI%)) + " no Mercado Exportación")
        End If
      End If
    Next KI&
    '
    Screen.MousePointer = 1
    TTXX$ = "Control Completo."
    If haynocla% = 0 Then
      TTXX$ = TTXX$ + "\No se Detectaron Códigos Mal Clasificados."
    End If
    Call COMUNI(TTXX$)
    '
End Sub

Private Sub Command8_Click()
    Dim TMS$(256), DETIMA$(256)
    DES% = FECHANUM("01/11/05")
    HAS% = FECHANUM("30/11/05")
    FIN& = ULTREG("MOVISTO")
    '
    NMS% = NUMAS%
    CATIMA% = 0
    For I& = 1 To FIN&
       Call TRACE(20, I& - PRIRE&, FIN& - PRIRE&)
       XS$ = REGLEIDO$("MOVISTO", I&)
       FF% = CVI(Left$(XS$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           COMO$ = UCase$(Mid$(XS$, 33, 2))
           If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
             CI% = CVI(Mid$(XS$, 3, 2))
             If CI% > 0 Then
               If CI% <= NMS% Then
                 REGMA& = CI% + 1
                 MX$ = REGLEIDO$("MASTER", REGMA&)
                 TIMAS$ = Mid$(MX$, 69, 2)
                 TMSX$ = MKI$(GRUCOVEI%(CI%)) + TIMAS$
                 '
                 
                 For KU% = 1 To CATIMA%
' TTXX$ = "": For KKI% = 1 To 4: TTXX$ = TTXX$ + Str$(Asc(Mid$(TMSX, KKI%, 1))): Next KKI%
' TTXX$ = TTXX$ + "      "
' For KKI% = 1 To 4: TTXX$ = TTXX$ + Str$(Asc(Mid$(TMS$(KU%), KKI%, 1))): Next KKI%
' MsgBox TTXX$
                  If TMS$(KU%) = TMSX$ Then GoTo 19
                 Next KU%
                 CATIMA% = CATIMA% + 1
                 TMS$(CATIMA%) = TMSX$
                 DETIMA$(CATIMA%) = TRIM$(ECOARCH$("GRUCOVEN", Left$(TMSX$, 2)))
                 DETIMA$(CATIMA%) = DETIMA$(CATIMA%) + " / " + TRIM$(ECOARCH$("CLASI001", Mid$(TMSX$, 3, 1)))
                 DETIMA$(CATIMA%) = DETIMA$(CATIMA%) + " / " + TRIM$(ECOARCH$("CLASI002", Mid$(TMSX$, 4, 1)))
               End If
             End If
           End If
         End If
       End If
19  Next I&

For KKI% = 1 To CATIMA%
TTXX$ = ""
For KKJ% = 1 To 4
 TTXX$ = TTXX$ + Str$(Asc(Mid$(TMS$(KKI%), KKJ%, 1)))
Next KKJ%
Printer.Print "       " + DETIMA$(KKI%) + "    "; TTXX$
Next KKI%
Printer.EndDoc
End Sub

Private Sub Form_Load()
    EPAC$ = EMPREAC$
    If InStr(UCase$(EPAC$), "SABO") < 1 Then
      If EXISTE%("ESTAVEN.EXE") < 1 Then
        Command3.Enabled = False
      End If
    End If
End Sub

VERSION 5.00
Begin VB.Form ESTAVENTB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFF0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadistica de Entregas"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ESTAVENTB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "List"
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
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Análisis de Facturación por Cliente"
      Top             =   3780
      Width           =   855
   End
   Begin VB.ListBox RACLIEN 
      BackColor       =   &H00D0FFF0&
      Height          =   3840
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   1575
      Visible         =   0   'False
      Width           =   9465
   End
   Begin VB.CommandButton Command6 
      Caption         =   "List"
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
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Listado Análisis por Facturación"
      Top             =   2940
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFF0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   4425
      Begin VB.CommandButton Command8 
         Caption         =   "Integral"
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
         Left            =   210
         TabIndex        =   19
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   15
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3045
         TabIndex        =   14
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0FFF0&
         Caption         =   "Número de Cliente:"
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
         Left            =   1050
         TabIndex        =   13
         Top             =   525
         Width           =   1905
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   9870
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   1890
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFF0&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   4830
      TabIndex        =   4
      Top             =   105
      Width           =   4845
      Begin VB.CommandButton Command2 
         Caption         =   "Calc"
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
         Left            =   3885
         Picture         =   "ESTAVENTB.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0FFF0&
         Caption         =   "Cálculo:"
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
         Left            =   105
         TabIndex        =   8
         Top             =   945
         Width           =   750
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   7
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0FFF0&
         Caption         =   "Rango:"
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
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   5
         ToolTipText     =   "Listado Analisis por Facturacion"
         Top             =   420
         Width           =   2745
      End
   End
   Begin VB.CommandButton command1 
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
      Height          =   750
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command7 
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
      Height          =   750
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
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
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":0EBC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprime Ficha de Distribución de Entregas"
      Top             =   4620
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "List"
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
      Left            =   9870
      Picture         =   "ESTAVENTB.frx":1526
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Listado Detalle Estadístico de Entregas Valorizadas"
      Top             =   2100
      Width           =   855
   End
End
Attribute VB_Name = "ESTAVENTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command2_Click()
'
'A PEDIDO DE DOSIVAC, NO TOMA EL PRECIO DE LISTA SINO EL
'DEL PEDIDO - SI NO HAY PEDIDO, EL PRECIO FACTURADO
'VER OBSERVACIONES MAS ABAJO
    '
    Dim PUNMO$(32767), TICLIE%(32767)
    '
'    Dim NCI%(8196)
'    Dim FFI%(8196)
'    Dim TI%(8196)
'    Dim VA%(8196)
'    Dim NRO&(8196)
'    Dim SIFAC%(8196)
'    Dim IFAC#(8196)
'    Dim IDET#(8196)
'    Dim PUMOFAC$(8196)
    Dim NCI%()
    Dim FFI%()
    Dim TI%()
    Dim VA%()
    Dim NRO&()
    Dim SIFAC%()
    Dim IFAC#()
    Dim IDET#()
    Dim PUMOFAC$()
    '
    Dim RMVS As String * 128
    ReDim NCI%(0)
    '
    'Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
    '
    FHCALCOS$ = TRIM$(CONTROL("ANACOSTO", "FECALCOS")) + " - " + TRIM$(CONTROL("ANACOSTO", "HOCALCOS")) + " hs."
    '
    If ULTREG&("MOVIPTA") > 0 Then
      Call BLOQARCH("MOVISTO")
      For U& = 1 To ULTREG&("MOVIPTA")
        XX$ = REGLEIDO$("MOVIPTA", U&)
        CII% = CVI(Mid$(XX$, 3, 2))
        If CII% > 0 Then
          If CII% <= NUMAS% Then
            Call REGAPP("MOVISTO", XX$)
          End If
        End If
      Next U&
      Call CIERRARCH("MOVISTO")
      Call SETULTREG("MOVIPTA", 0)
    End If
    Call CIERRARCH("MOVIPTA")
    '
    VEGRUVE$ = ""
    For U& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO$("GRUCOVEN", U&)
      GRUCO% = CVI(Left$(X$, 2))
      VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
    Next U&
    VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(999)
    '
10  FEX = HOY(HO$)
    HAS% = FEX: HASS$ = FECHATEX$(FEX)
    DES% = FECHANUM("01" + Mid$(HO$, 3))
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    Label1.Caption = PERIO$
    Label1.Refresh
    Refresh
    '
    Screen.MousePointer = 11
    '
    Call HEADERS("DETCLIEN", "")
    Call HEADERS("DETCLIEN", "Periodo: " + PERIO$)
    Call HEADERS("DETCLIEN", "Costos Calculados al: " + FHCALCOS$)
    '
    GoSub 111    ' carga vector de facturas emitidas
    GoSub 112    ' carga vector detalle de facturacion - MOVISTO
    GoSub 113    ' carga vector detalle de facturacion - MOVIREPA
    '
    UREMO& = ULTREG&("MOVISTO")
    UREPA& = ULTREG&("MOVIREPA")
    FIN& = CAFAC%
    J& = 0: J1& = 0: J2& = 0
    '
    For JII% = 1 To CAFAC%
      I& = JII%
      Call TRACE(20, I&, FIN&)
      '
      If TI%(JII%) < 1 Or TI%(JII%) > 6 Then GoTo 28  ' descarta ajustes
      REMIFAC$ = "FC"
      If TI%(JII%) = 3 Then REMIFAC$ = "ND"
      If TI%(JII%) = 6 Then REMIFAC$ = "NC"
      VASS$ = "A": If VA%(JII%) = 2 Then VASS$ = "B"
                   If VA%(JII%) = 3 Then VASS$ = "E"
      REMIFAC$ = REMIFAC$ + "-" + TRIM$(Str$(NRO&(JII%)))
      REMIFAC$ = REMIFAC$ + "-" + VASS$
      '
      COEF# = 1
      If Abs(IDET#(JII%)) >= 0.005 Then
        COEF# = (Abs(IFAC#(JII%))) / (Abs(IDET#(JII%)))
      End If
      '
      For KKI% = 1 To Len(PUMOFAC$(JII%)) Step 4
         REG& = CVS(Mid$(PUMOFAC$(JII%), KKI%, 4))
         If REG& > 0 And REG& <= UREMO& Then
              X$ = REGLEIDO$("MOVISTO", REG&)
            ElseIf REG& < 0 And Abs(REG&) <= UREPA& Then
              X$ = REGLEIDO$("MOVIREPA", Abs(REG&))
            Else
              GoTo 27
         End If
         CLI% = NCI%(JII%)
         FF% = FFI%(JII%)
         NRO1& = NRO&(JII%)
         CII% = CVI(Mid$(X$, 3, 2))
         CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
         VUNI = CVD(Mid$(X$, 73, 8)) * COEF#
         '
' AQUI BUSCAR PRECIO DE LA NOTA DE PEDIDO
PPCC$ = RECFILT$("PEDDETA", 3, MKI$(CLI%))
NREMITO& = Val(Mid$(X$, 26, 7))
For KKX& = 1 To Len(PPCC$) Step 4
  KKY& = CVS(Mid$(PPCC$, KKX&, 4))
  KKZ$ = REGLEIDO$("PEDDETA", KKY&)
  If CVI(Mid$(KKZ$, 33, 2)) = CII% Then
    If CVS(Mid$(KKZ$, 105, 4)) = NREMITO& Then
      PRELI = CVS(Mid$(KKZ$, 71, 4))
      GoTo 18
    End If
  End If
Next KKX&
' NO ENCONTRO PEDIDO - TOMO EL PRECIO FACTURADO
PRELI = VUNI
         'PRELI = PRELISTA#("LIPRE001", CII%)
18       If PRELI < 0.1 Then PRELI = CVD(Mid$(X$, 73, 8))
         If REG& < 0 Then PRELI = VUNI    ' reparaciones
'SI SE FACTURO A PRECIO MAYOR AL PEDIDO TOMA EL FACTURADO
If PRELI < VUNI Then PRELI = VUNI
         VALOUN = VUNI
         VATOT# = CANTU# * VALOUN
         '
         If Abs(PRELI) > 0.00005 Or Abs(VUNI) > 0.00005 Then
            Z$ = REGBLAN$("DETCLIEN")
            Call REPLA(Z$, MKI$(CLI%), 1, 2)
            Call REPLA(Z$, MKI$(FF%), 3, 2)
            Call REPLA(Z$, REMIFAC$, 5, 10)
            Call REPLA(Z$, MKI$(CII%), 15, 2)
            Call REPLA(Z$, MKD$(CANTU#), 17, 8)
            Call REPLA(Z$, MKS$(VUNI), 25, 4)
            Call REPLA(Z$, MKS$(PRELI), 29, 4)
            Call REPLA(Z$, MKD$(VATOT#), 33, 8)
            '
            J& = J& + 1
            Call GRAREG("DETCLIEN", Z$, J&)
            '
            Z$ = REGBLAN$("LIANAFAC")
            GRCC% = GRUCOVEI%(CII%): If GRCC% < 1 Then GRCC% = 999
            If InStr(VEGRUVE$, MKI$(0) + MKI$(GRCC%)) < 1 Then GRCC% = 999
            Call REPLA(Z$, MKI$(1), 1, 2)    ' productos
            If REG& < 0 Then
               Call REPLA(Z$, MKI$(2), 1, 2)    ' reparaciones
            End If
            '
            Call REPLA(Z$, MKI$(GRCC%), 3, 2)
            Call REPLA(Z$, MKI$(CII%), 5, 2)
            Call REPLA(Z$, REMIFAC$, 7, 18)
            Call REPLA(Z$, MKI$(FF%), 25, 2)
            Call REPLA(Z$, MKD$(CANTU#), 27, 8)
            Call REPLA(Z$, MKD$(CANTU# * PRELI), 35, 8)
            Call REPLA(Z$, MKD$(CANTU# * (PRELI - VUNI)), 43, 8)
            PORDE = 0: If PRELI > 0 Then PORDE = 100 * (PRELI - VUNI) / PRELI
            Call REPLA(Z$, MKS$(PORDE), 51, 4)
            Call REPLA(Z$, MKD$(CANTU# * VUNI), 55, 8)
            COU = COSCALCU(CII%, MODOCO%)
'If MODOCO% = 2 Then
'  Call AREXPLO(CII%, 1, 1)
'  If CAIT% > 0 Then
'    MsgBox CODEXT$(CII%)
'  End If
'End If
            '
            If REG& < 0 Then COU = 0
            Call REPLA(Z$, MKD$(CANTU# * COU), 63, 8)
            RENTA# = CANTU# * (VUNI - COU)
            Call REPLA(Z$, MKD$(RENTA#), 71, 8)
            POREN = 0: If Abs(CANTU# * VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
            Call REPLA(Z$, MKS$(POREN), 79, 4)
            '
            J1& = J1& + 1
            Call GRAREG("LIANAFAC", Z$, J1&)
            '
            J2& = J2& + 1
            Call GRAREG("LIANACLI", MKI$(CLI%) + Z$, J2&)
            '
            IFAC#(JII%) = IFAC#(JII%) - CANTU# * VUNI * SIFAC%(JII%)
            '
         End If
         '
27    Next KKI%
      '
      If Abs(IFAC#(JII%)) >= 0.005 Then
         CLI% = NCI%(JII%)
         FF% = FFI%(JII%)
         NRO1& = NRO&(JII%)
         CII% = 0
         CANTU# = SIFAC%(JII%)
         VUNI = Abs(IFAC#(JII%))
         PRELI = VUNI
         VALOUN = VUNI
         VATOT# = CANTU# * VALOUN
         '
         Z$ = REGBLAN$("DETCLIEN")
         Call REPLA(Z$, MKI$(CLI%), 1, 2)
         Call REPLA(Z$, MKI$(FF%), 3, 2)
         Call REPLA(Z$, REMIFAC$, 5, 10)
         Call REPLA(Z$, MKI$(CII%), 15, 2)
         Call REPLA(Z$, MKD$(CANTU#), 17, 8)
         Call REPLA(Z$, MKS$(VUNI), 25, 4)
         Call REPLA(Z$, MKS$(PRELI), 29, 4)
         Call REPLA(Z$, MKD$(VATOT#), 33, 8)
         '
         J& = J& + 1
         Call GRAREG("DETCLIEN", Z$, J&)
         '
         Z$ = REGBLAN$("LIANAFAC")
         GRCC% = 999
         Call REPLA(Z$, MKI$(3), 1, 2)    ' debitos y creditos
         Call REPLA(Z$, MKI$(GRCC%), 3, 2)
         Call REPLA(Z$, MKI$(CII%), 5, 2)
         Call REPLA(Z$, REMIFAC$, 7, 18)
         Call REPLA(Z$, MKI$(FF%), 25, 2)
         Call REPLA(Z$, MKD$(CANTU#), 27, 8)
         Call REPLA(Z$, MKD$(CANTU# * PRELI), 35, 8)
         Call REPLA(Z$, MKD$(CANTU# * (PRELI - VUNI)), 43, 8)
         PORDE = 0: If PRELI > 0 Then PORDE = 100 * (PRELI - VUNI) / PRELI
         Call REPLA(Z$, MKS$(PORDE), 51, 4)
         Call REPLA(Z$, MKD$(CANTU# * VUNI), 55, 8)
         COU = COSCALCU(CII%, MODOCO%)
'If MODOCO% = 2 Then
'  Call AREXPLO(CII%, 1, 1)
'  If CAIT% > 0 Then
'    MsgBox CODEXT$(CII%)
'  End If
'End If
         If REG& < 0 Then COU = 0
         Call REPLA(Z$, MKD$(CANTU# * COU), 63, 8)
         RENTA# = CANTU# * (VUNI - COU)
         Call REPLA(Z$, MKD$(RENTA#), 71, 8)
         POREN = 0: If Abs(VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
         Call REPLA(Z$, MKS$(POREN), 79, 4)
         '
         J1& = J1& + 1
         Call GRAREG("LIANAFAC", Z$, J1&)
         '
         J2& = J2& + 1
         Call GRAREG("LIANACLI", MKI$(CLI%) + Z$, J2&)
         '
      End If
      '
28  Next JII%
    '
    Call SETULTREG("DETCLIEN", J&)
    Call CIERRARCH("DETCLIEN")
    Call LIBARCH("DETCLIEN")
    '
    Call SETULTREG("LIANAFAC", J&)
    Call CIERRARCH("LIANAFAC")
    Call LIBARCH("LIANAFAC")
    '
    Call SETULTREG("LIANACLI", J2&)
    Call CIERRARCH("LIANACLI")
    Call FILESORT("LIANACLI", "", 1, 4)
    Call CIERRARCH("LIANACLI")
    Call LIBARCH("LIANACLI")
    '
    Call HEADERS("LIANACLI", "")
    Call HEADERS("LIANACLI", "Periodo: " + PERIO$)
    Call HEADERS("LIANACLI", "Costos Calculados al: " + FHCALCOS$)
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVAC") > 0 Then
       Call SORTANAFA
       Call LISTANAFA
       GoTo 99
    End If
    '
    FIN& = ULTREG&("DETCLIEN")
    CLIA% = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("DETCLIEN", U&)
      ABUS$ = Space$(41): ABUS2$ = Space$(16)
      CLI% = CVI(Left$(Z$, 2))
      If CLI% <> CLIA% Then
        VENDE% = VENDCLI%(CLI%)
        VENDEX$ = TRIM$(Str$(VENDE%))
        While Len(VENDEX$) < 5: VENDEX$ = "0" + VENDEX$: Wend
        GRUPO% = GRUCOCLI%(CLI%)
        GRUCLI$ = TRIM$(Str$(GRUPO%))
        While Len(GRUCLI$) < 3: GRUCLI$ = "0" + GRUCLI$: Wend
        RACLI$ = RASOCLI$(CLI%)
        CLIA% = CLI%
      End If
      Call REPLA(ABUS$, VENDEX$, 1, 5)
      Call REPLA(ABUS$, GRUCLI$, 6, 3)
      Call REPLA(ABUS$, RACLI$, 9, 30)
      '
      CII% = CVI(Mid$(Z$, 15, 2))
      GRUCO% = GRUCOVEI%(CII%)
      LINEA$ = TRIM$(Str$(GRUCO%))
      While Len(LINEA$) < 3: LINEA$ = "0" + LINEA$: Wend
      Call REPLA(ABUS$, LINEA$, 39, 3)
      Call REPLA(ABUS2$, CODEXT$(CII%), 1, 16)
      '
      CANTI = CVD(Mid$(Z$, 17, 8))
      VATOT# = CVD(Mid$(Z$, 33, 8))
      '
      If VATOT# < 0.005 Then GoTo 29
      '
      For KK% = 1 To RACLIEN.ListCount
        X$ = RACLIEN.List(KK% - 1)
        If Left$(X$, 41) = ABUS$ Then
          If Mid$(X$, 54, 16) = ABUS2$ Then
            RACLIEN.RemoveItem KK% - 1
            GoTo 20
          End If
        End If
      Next KK%
      '
      X$ = Space$(96)
      Call REPLA(X$, ABUS$, 1, 41)
      Call REPLA(X$, "999999999.99", 42, 12)
      Call REPLA(X$, ABUS2$, 54, 16)
      '
20    VASOR# = Val(Mid$(X$, 42, 12)) - VATOT#
      VASOX$ = TRIM$(CSTRING$(MKD$(VASOR#), 4, 12, 2, 2))
      While Len(VASOX$) < 12: VASOX$ = "0" + VASOX$: Wend
      '
      CANTI = CANTI + Val(Mid$(X$, 70, 12))
      CATOX$ = TRIM$(CSTRING$(MKS$(CANTI), 4, 12, 2, 2))
      While Len(CATOX$) < 12: CATOX$ = "0" + CATOX$: Wend
      '
      VATOT# = Val(Mid$(X$, 82, 12)) + VATOT#
      VATOX$ = TRIM$(CSTRING$(MKD$(VATOT#), 4, 12, 2, 2))
      While Len(VATOX$) < 12: VATOX$ = "0" + VATOX$: Wend
      '
      Call REPLA(X$, VASOX$, 42, 12)
      Call REPLA(X$, CATOX$, 70, 12)
      Call REPLA(X$, VATOX$, 82, 12)
      RACLIEN.AddItem X$
      RACLIEN.Refresh
      '
29  Next U&
    '
    JJ& = 0
    VENDEA% = -1
    GRUCLA% = -1
    RACLIA$ = ""
    LINEIA% = (-1)
    FIN& = RACLIEN.ListCount
    For KK% = 1 To RACLIEN.ListCount
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      X$ = RACLIEN.List(KK% - 1)
      '
      VENDE% = Val(Left$(X$, 5))
      If VENDE% <> VENDEA% Then
        If LINEIA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".           TOTAL " + DEGRUVEN$(LINEIA%)
          CATO = CATOLIN
          IMPO# = ITOLIN#
          GoSub 80
          CATOLIN = 0
          ITOLIN# = 0
          LINEIA% = (-1)
        End If
        If RACLIA$ <> "" Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".       TOTAL " + RACLIA$
          CATO = CATOCLI
          IMPO# = ITOCLI#
          GoSub 80
          CATOCLI = 0
          ITOCLI# = 0
          RACLIA$ = ""
        End If
        If GRUCLA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".   TOTAL " + DEGRUCLI$(GRUCLA%)
          CATO = CATOGRU
          IMPO# = ITOGRU#
          GoSub 80
          CATOGRU = 0
          ITOGRU# = 0
          GRUCLA% = (-1)
        End If
        If VENDEA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = "TOTAL " + DEVENDE$(VENDEA%)
          CATO = CATOVEN
          IMPO# = ITOVEN#
          GoSub 80
          CATOVEN = 0
          ITOVEN# = 0
          JJ& = JJ& + 1: Call GRAREG("RACLIEN", String$(128, "="), JJ&)
        End If
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        TTXX$ = "Vendedor: " + DEVENDE$(VENDE%)
        CATO = 0
        IMPO# = 0
        GoSub 80
        VENDEA% = VENDE%
      End If
      '
      GRUCL% = Val(Mid$(X$, 6, 3))
      If GRUCL% <> GRUCLA% Then
        If LINEIA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".           TOTAL " + DEGRUVEN$(LINEIA%)
          CATO = CATOLIN
          IMPO# = ITOLIN#
          GoSub 80
          CATOLIN = 0
          ITOLIN# = 0
          LINEIA% = (-1)
        End If
        If RACLIA$ <> "" Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".       TOTAL " + RACLIA$
          CATO = CATOCLI
          IMPO# = ITOCLI#
          GoSub 80
          CATOCLI = 0
          ITOCLI# = 0
          RACLIA$ = ""
        End If
        If GRUCLA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".   TOTAL " + DEGRUCLI$(GRUCLA%)
          CATO = CATOGRU
          IMPO# = ITOGRU#
          GoSub 80
          CATOGRU = 0
          ITOGRU# = 0
          GRUCLA% = (-1)
          JJ& = JJ& + 1: Call GRAREG("RACLIEN", String$(128, "="), JJ&)
        End If
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        TTXX$ = ".   Grupo Cliente: " + DEGRUCLI$(GRUCL%)
        CATO = 0
        IMPO# = 0
        GoSub 80
        GRUCLA% = GRUCL%
      End If
      '
      RACLI$ = Mid$(X$, 9, 30)
      If RACLI$ <> RACLIA$ Then
        If LINEIA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".           TOTAL " + DEGRUVEN$(LINEIA%)
          CATO = CATOLIN
          IMPO# = ITOLIN#
          GoSub 80
          CATOLIN = 0
          ITOLIN# = 0
          LINEIA% = (-1)
        End If
        If RACLIA$ <> "" Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".       TOTAL " + RACLIA$
          CATO = CATOCLI
          IMPO# = ITOCLI#
          GoSub 80
          CATOCLI = 0
          ITOCLI# = 0
          RACLIA$ = ""
        End If
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        TTXX$ = ".       Cliente: " + RACLI$
        CATO = 0
        IMPO# = 0
        GoSub 80
        RACLIA$ = RACLI$
     End If
      '
      LINEI% = Val(Mid$(X$, 39, 3))
      If LINEI% <> LINEIA% Then
        If LINEIA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".           TOTAL " + DEGRUVEN$(LINEIA%)
          CATO = CATOLIN
          IMPO# = ITOLIN#
          GoSub 80
          CATOLIN = 0
          ITOLIN# = 0
          LINEIA% = (-1)
        End If
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        TTXX$ = ".           Linea: " + DEGRUVEN$(LINEI%): GoSub 80
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        LINEIA% = LINEI%
      End If
      '
      COD$ = Mid$(X$, 54, 16)
      CI% = CODINT%(COD$)
      CATO = Val(Mid$(X$, 70, 12))
      IMPO# = Val(Mid$(X$, 82, 12))
      TTXX$ = ".                " + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT0$(CI%))
      GoSub 80
      CATOLIN = CATOLIN + CATO
      CATOCLI = CATOCLI + CATO
      CATOGRU = CATOGRU + CATO
      CATOVEN = CATOVEN + CATO
      CATOGEN = CATOGEN + CATO
      ITOLIN# = ITOLIN# + IMPO#
      ITOCLI# = ITOCLI# + IMPO#
      ITOGRU# = ITOGRU# + IMPO#
      ITOVEN# = ITOVEN# + IMPO#
      ITOGEN# = ITOGEN# + IMPO#
      '
    Next KK%
        '
        If LINEIA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".           TOTAL " + DEGRUVEN$(LINEIA%)
          CATO = CATOLIN
          IMPO# = ITOLIN#
          GoSub 80
          CATOLIN = 0
          ITOLIN# = 0
          LINEIA% = (-1)
        End If
        If RACLIA$ <> "" Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".       TOTAL " + RACLIA$
          CATO = CATOCLI
          IMPO# = ITOCLI#
          GoSub 80
          CATOCLI = 0
          ITOCLI# = 0
          RACLIA$ = ""
        End If
        If GRUCLA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = ".   TOTAL " + DEGRUCLI$(GRUCLA%)
          CATO = CATOGRU
          IMPO# = ITOGRU#
          GoSub 80
          CATOGRU = 0
          ITOGRU# = 0
          GRUCLA% = (-1)
        End If
        If VENDEA% >= 0 Then
          TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
          TTXX$ = "TOTAL " + DEVENDE$(VENDEA%)
          CATO = CATOVEN
          IMPO# = ITOVEN#
          GoSub 80
          CATOVEN = 0
          ITOVEN# = 0
        End If
        '
        TTXX$ = "": CATO = 0: IMPO# = 0: GoSub 80
        TTXX$ = "TOTAL GENERAL ....................................................."
        CATO = CATOGEN
        IMPO# = ITOGEN#
        GoSub 80
        '
    Call SETULTREG("RACLIEN", JJ&)
    Call CIERRARCH("RACLIEN")
    '
    Call HEADERS("RACLIEN", "")
    Call HEADERS("RACLIEN", "Periodo: " + PERIO$)
    Call HEADERS("RACLIEN", "Costos Calculados al: " + FHCALCOS$)
    Call HEADERS("LIANACLI", "")
    Call HEADERS("LIANACLI", "Periodo: " + PERIO$)
    Call HEADERS("LIANACLI", "Costos Calculados al: " + FHCALCOS$)
    GoTo 99
    '
80  WW$ = REGBLAN$("RACLIEN")
    Call REPLA(WW$, TTXX$, 1, 80)
    Call REPLA(WW$, MKS$(CATO), 81, 4)
    Call REPLA(WW$, MKD$(IMPO#), 85, 8)
    JJ& = JJ& + 1
    Call GRAREG("RACLIEN", WW$, JJ&)
    Return
    '
99  Call CIERRARCH("*.*")
    HOII% = HOY(HOS$)
    Call GRACONTROL("ESTA-VEN", "PERICALC", PERIO$)
    Label1.Caption = CONTROL$("ESTA-VEN", "PERICALC")
    Call GRACONTROL("ESTA-VEN", "FECHACAL", HOS$ + " - " + Time$ + " hs")
    Label3.Caption = CONTROL$("ESTA-VEN", "FECHACAL")
    Call CIERRARCH("*.*")
    '
    Call TORTAVEN
991 Screen.MousePointer = 1
    Exit Sub
    '
    '
    '
111 CAFAC% = 0: SUFAC# = 0: FIN& = ULTREG&("FACTUR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XS$ = REGLEIDO$("FACTUR", I&)
      FF% = CVI(Mid$(XS$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If CAFAC% >= 32767 Then
            Call MENSERR(24, "Imposible Generar Consulta.\Elija Menor Rango de Fechas.")
            GoTo 991
          End If
          '
          CAFAC% = CAFAC% + 1
          If CAFAC% > UBound(NCI%) Then
             ReDim Preserve NCI%(CAFAC%)
             ReDim Preserve FFI%(CAFAC%)
             ReDim Preserve TI%(CAFAC%)
             ReDim Preserve VA%(CAFAC%)
             ReDim Preserve NRO&(CAFAC%)
             ReDim Preserve SIFAC%(CAFAC%)
             ReDim Preserve IFAC#(CAFAC%)
             ReDim Preserve IDET#(CAFAC%)
             ReDim Preserve PUMOFAC$(CAFAC%)
          End If
          NCI%(CAFAC%) = CVI(Left$(XS$, 2))
          FFI%(CAFAC%) = FF%
          TI%(CAFAC%) = Asc(Mid$(XS$, 13, 1))
          If TI%(CAFAC%) < 0 Or TI%(CAFAC%) > 6 Then TI%(CAFAC%) = 2 ' por default factura
          VA%(CAFAC%) = Asc(Mid$(XS$, 14, 1))
          If VA%(CAFAC%) < 1 Or VA%(CAFAC%) > 2 Then VA%(CAFAC%) = 3 ' por default factura
          NRO&(CAFAC%) = CVS(Mid$(XS$, 3, 4))
          '
          If TI%(CAFAC%) = 0 Then                     ' contempla ajustes de saldos
            If CVD(Mid$(XS$, 15, 8)) < 0 Then
              TI%(CAFAC%) = 7
            End If
          End If
          '
          IFAC#(CAFAC%) = Abs(CVD(Mid$(XS$, 15, 8)) + CVD(Mid$(XS$, 49, 8)))
          '
          SIG% = 1: If TI%(CAFAC%) > 3 Then SIG% = (-1)
          SIFAC%(CAFAC%) = SIG%
          SUFAC# = SUFAC# + IFAC#(CAFAC%) * SIG%
        End If
      End If
    Next I&
    '
Return
'
'
112 PRIRE& = 999999: ULTRE& = 0
    FIN& = ULTREG&("MOVISTO")
    FINUMO% = FILENBR%("MOVISTO")
    PRIRESTO& = 0
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      ' Get #FINUMO%, I&, RMVS$
      XS$ = REGLEIDO$("MOVISTO", I&)
      ' XS$ = RMVS$
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If PRIRESTO& = 0 Then PRIRESTO& = I&
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
            If TRIM$(UCase$(Mid$(XS$, 23, 2))) <> "PL" Then
              TP1% = 2: If COMO$ = "NC" Then TP1% = 6
              CLI1% = CVI(Mid$(XS$, 84, 2))
              If COMO$ = "ND" Then TP1% = 3
              NRO1& = Val(Mid$(XS$, 36, 7))
              CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
              PRE# = CVD(Mid$(XS$, 73, 8))
              TOT# = PRE# * CAN#
              If I& < PRIRE& Then PRIRE& = I&
              If I& > ULTRE& Then ULTRE& = I&
              REMOSTO& = I&
              GoSub 120
            End If
          End If
        End If
      End If
    Next I&
    '
Return
'
113 FIN& = ULTREG&("MOVIREPA")
    FINUMO% = FILENBR%("MOVIREPA")
    PRIREREP& = 0
    For I& = 1 To FIN&
      Call TRACE(20, I& - PRIREMO&, FIN& - PRIREMO&)
      ' Get #FINUMO%, I&, RMVS$
      XS$ = REGLEIDO$("MOVIREPA", I&)
      ' XS$ = RMVS$
      FF% = CVI(Left$(XS$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If PRIREREP& = 0 Then PRIREREP& = I&
          COMO$ = UCase$(Mid$(XS$, 33, 2))
          If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
            If TRIM$(UCase$(Mid$(XS$, 23, 2))) <> "PL" Then
              TP1% = 2: If COMO$ = "NC" Then TP1% = 6
              CLI1% = CVI(Mid$(XS$, 84, 2))
              If COMO$ = "ND" Then TP1% = 3
              NRO1& = Val(Mid$(XS$, 36, 7))
              CAN# = CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
              PRE# = CVD(Mid$(XS$, 73, 8))
              TOT# = PRE# * CAN#
              If I& < PRIRE& Then PRIRE& = I&
              If I& > ULTRE& Then ULTRE& = I&
              REMOSTO& = (-1) * I&
              GoSub 120
            End If
          End If
        End If
      End If
    Next I&
    '
Return
'
120 vuelta% = 0
121 For JII% = 1 To CAFAC%
      If TI%(JII%) = TP1% Then
        If NRO&(JII%) = NRO1& Then
          If NCI%(JII%) = CLI1% Or vuelta% = 1 Then
            If FFI%(JII%) = FF% Or vuelta% = 1 Then
              If Abs(IFAC#(JII%)) >= 0.005 Then ' descarta anuladas
                IDET#(JII%) = IDET#(JII%) + TOT#
                PUMOFAC$(JII%) = PUMOFAC$(JII%) + MKS$(REMOSTO&)
              End If
              GoTo 124    ' encontro documento
            End If
          End If
        End If
      End If
    Next JII%
    '
    If vuelta% = 0 Then
      vuelta% = 1: GoTo 121
    End If
    '
    If vuelta% = 1 Then
      If TP1% = 2 Then
        TP1% = 3
        vuelta% = 2
        GoTo 121
      End If
    End If
    '
AAA = BBB
124 Return
    '
End Sub

Private Sub Command3_Click()
   PRINTFORM
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call FINAL("*DETACLI")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Height = 5860
   Top = 1570
   Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
   Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% > 0 Then
      Label8.Caption = TRIM$(Str$(NCXI%))
      Label6.Caption = RASOCLI$(NCXI%)
      ESTAVENTB.Refresh
      Call TORTAVEN
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   EPAC$ = EMPREAC$
   If InStr(EPAC$, "DOSIVAC") > 0 Then
       Call HEADERS("LISTANA", "")
       Call HEADERS("LISTANA", "Periodo: " + TRIM$(Label1.Caption))
       Call ENCACOL("LISTANA", "")
       Call ENCACOL("LISTANA", "R.   Codigo           Comprobante   Fecha         Cant      Bruto    Desc($)   Desc(%)       Neto      Costo    C.M.($)  C.M.(%)")
       Call FINAL("*LISTANA")
     Else
       Call FINAL("*RACLIEN")
   End If
   Command6.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Height = 9000
   Top = 0
   Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
   Label8.Caption = "0"
   Label6.Caption = "Total General Empresa"
   ESTAVENTB.Refresh
   Call TORTAVEN
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("*LIANACLI", "Analisis de Facturación por Cliente")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "NEUMANN") > 0 Then
       Command6.ToolTipText = "Listado de Ranking por Cliente"
    End If
    '
    Label1.Caption = CONTROL$("ESTA-VEN", "PERICALC")
    Label3.Caption = CONTROL$("ESTA-VEN", "FECHACAL")
    Show
    Refresh
    '
    If ULTREG&("DETCLIEN") > 0 Then
       Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
       Label8.Caption = "0"
       Label6.Caption = "Total General Empresa"
       ESTAVENTB.Refresh
       Call TORTAVEN
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub TORTAVEN()
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), DEGRU$(32767)
   '
   VEGRUVE$ = ""
   For U& = 1 To ULTREG&("GRUCOVEN")
     X$ = REGLEIDO$("GRUCOVEN", U&)
     GRUCO% = CVI(Left$(X$, 2))
     VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(GRUCO%)
   Next U&
   VEGRUVE$ = VEGRUVE$ + MKI$(0) + MKI$(999)
   '
   ESTAVENTB.DrawWidth = 2
   ESTAVENTB.ScaleMode = 1
   Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
   TOTAVEN = 0
   NC% = Val(Label8.Caption)
   FIN& = ULTREG&("DETCLIEN")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("DETCLIEN", U&)
      CLI% = CVI(Left$(Z$, 2))
      If CLI% = NC% Or NC% = 0 Then
         CI% = CVI(Mid$(Z$, 15, 2))
         GRUCO% = GRUCOVEI%(CI%)
         If GRUCO% < 1 Then GRUCO% = 999
         If InStr(VEGRUVE$, MKI$(0) + MKI$(GRUCO%)) < 1 Then GRUCO% = 999
         TOVEN(GRUCO%) = TOVEN(GRUCO%) + CVD(Mid$(Z$, 33, 8))
         TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 33, 8))
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUVARI% = 999
   TOVEN(GRUVARI%) = TOTAVEN
   DEGRU$(GRUVARI%) = "Varios"
   For U& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO$("GRUCOVEN", U&)
      GRUCO% = CVI(Left$(X$, 2))
      DEGRU$(GRUCO%) = Mid$(X$, 3, 30)
      If GRUCO% > GRUCOMA% Then GRUCOMA% = GRUCO%
      TOVEN(GRUVARI%) = TOVEN(GRUVARI%) - TOVEN(GRUCO%)
      'List1.AddItem TTT$
   Next U&
   '
   vuelta% = 0
   '
10 OFV = 1680: OFH = 210
   INI = 0: VERTI = 0
   ESTAVENTB.DrawWidth = 2
   ESTAVENTB.ScaleMode = 1
   For UI% = 1 To GRUVARI%
     If TOVEN(UI%) >= 0.005 Then
       ROJO = ROJO + 192
       If ROJO > 255 Then VERDE = VERDE + 192: ROJO = ROJO - 256
       If VERDE > 255 Then AZUL = AZUL + 192: VERDE = VERDE - 256
       If AZUL > 256 Then AZUL = AZUL - 256
       ESTAVENTB.ForeColor = RGB(ROJO, VERDE, AZUL)
       '
       FINI = INI + 2 * 3.14159 * TOVEN(UI%) / TOTAVEN
       For RADIO = 1 To 1500 Step 5
         On Error Resume Next
         ESTAVENTB.Circle (7300, 3500), RADIO, , INI, (FINI + INI) / 2
         ESTAVENTB.Circle (7300, 3500), RADIO, , (FINI + INI) / 2, FINI
         On Error GoTo 0
       Next RADIO
       INI = FINI
       '
       VERTI = VERTI + 200: If NC% = 0 Then VERTI = VERTI - 60
       X1 = 0 + OFH: X2 = 300 + OFH
       Y1 = VERTI + OFV: Y2 = VERTI + 120 + OFV
       Line (X1, Y1)-(X2, Y2), RGB(ROJO, VERDE, AZUL), BF
       TTT$ = CSTRING$(MKI$(UI%), 1, 3, 0, 2) + "  " + AJUSTI$(DEGRU$(UI%), 21) + CSTRING$(MKS$(TOVEN(UI%)), 4, 12, 2, 2)
       CurrentX = 350 + OFH
       CurrentY = VERTI - 50 + OFV
       ForeColor = QBColor(0)
       Print TTT$
     End If
   Next UI%
   '
   VERTI = VERTI + 200
   TTT$ = Space$(26) + "------------"
   CurrentX = 350 + OFH
   CurrentY = VERTI - 50 + OFV
   ForeColor = QBColor(0)
   Print TTT$
   '
   VERTI = VERTI + 200
   CurrentX = 350 + OFH
   CurrentY = VERTI - 50 + OFV
   TTT$ = Space$(5) + AJUSTI$("TOTAL GENERAL", 21) + CSTRING$(MKS$(TOTAVEN), 4, 12, 2, 2)
   Print TTT$
   '
   If vuelta% < 1 Then
     If CurrentY > ScaleHeight - 300 Then
       vuelta% = vuelta% + 1
       HNUE = CurrentY + 700
       If HNUE > 9000 Then HNUE = 9000
       TONUE = Top - (HNUE - Height) / 2
       Height = HNUE
       Top = TONUE
       Line (0, 1680)-(10965, Height), ESTAVENTB.BackColor, BF
       Refresh
       GoTo 10
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub SORTANAFA()
   '
   RACLIEN.Clear
   FIN& = ULTREG&("LIANAFAC")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("LIANAFAC", U&)
     TTXX$ = "* " + AJUSTD$(CSTRING$(Left$(X$, 2), 1, 4, 0, 2), 4)
     TTXX$ = TTXX$ + AJUSTD$(CSTRING$(Mid$(X$, 3, 2), 1, 4, 0, 2), 4)
     TTXX$ = TTXX$ + AJUSTI$(CODEXT$(CVI(Mid$(X$, 5, 2))), 16)
     TTXX$ = TTXX$ + CSTRING$(MKS$(U&), 3, 6, 0, 2)
     '
     For KK& = 1 To RACLIEN.ListCount
       If RACLIEN.List(KK& - 1) > TTXX$ Then
         RACLIEN.AddItem TTXX$, KK& - 1
         GoTo 19
       End If
     Next KK&
     RACLIEN.AddItem TTXX$
     '
19 Next U&
   '
End Sub
'
Sub LISTANAFA()
   '
   Dim SUMACAN#(3), SUMABRU#(3), SUMANET#(3)
   Dim SUMADES#(3), SUMACOS#(3), SUMAREN#(3)
   '
   CODARANT% = 0
   GRUVEANT% = 0
   LINEAANT% = 0
   '
   NIVEINI% = 0: GoSub 60
   '
   JJ& = 0
   URE& = ULTREG&("LIANAFAC")
   FIN& = RACLIEN.ListCount
   For U& = 1 To FIN&
      RXX$ = RACLIEN.List(U& - 1)
      REG& = Val(Mid$(RXX$, 27, 6))
      If REG& > 0 Then
        If REG& <= URE& Then
          X$ = REGLEIDO$("LIANAFAC", REG&)
          LINEA% = CVI(Left$(X$, 2))
          GRUVE% = CVI(Mid$(X$, 3, 2))
          CODAR% = CVI(Mid$(X$, 5, 2))
          '
          If U& > 1 Then
            If LINEA% <> LINEAANT% Then
              GoSub 10
              GoSub 11
              GoSub 12        ' TOTALIZAR POR LINEA
              GRUVEANT% = 0
              CODARANT% = 0
              LINEAANT% = 0
            ElseIf GRUVE% <> GRUVEANT% Then
              GoSub 10
              GoSub 11        ' TOTALIZAR POR GRUPO DE VENTA
              GRUVEANT% = 0
              CODARANT% = 0
            ElseIf CODAR% <> CODARANT% Then
              GoSub 10        ' TOTAL POR ARTICULO
              CODARANT% = 0
            End If
          End If
          '
          If LINEA% <> LINEAANT% Then
            GoSub 20           ' ENCABEZADO DE LINEA
          End If
          If GRUVE% <> GRUVEANT% Then
            GoSub 21           ' ENCABEZADO DE GRUPO DE VENTA
          End If
          '
          TTXX$ = Space$(256)
          Call REPLA(TTXX$, ".", 1, 1)
          If CODAR% <> CODARANT% Then Call REPLA(TTXX$, CODEXT$(CODAR%), 6, 16)
          Call REPLA(TTXX$, Mid$(X$, 7, 18), 23, 12) ' NRO FACTURA
          FEX = CVI(Mid$(X$, 25, 2))
          Call REPLA(TTXX$, FECHATEX$(FEX), 37, 8)
          '
          CANTID# = CVD(Mid$(X$, 27, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(CANTID#), 4, 10, 0, 2), 45, 10)
          For NIVEL% = 0 To 3: SUMACAN#(NIVEL%) = SUMACAN#(NIVEL%) + CANTID#: Next NIVEL%
          '
          BRULIS# = CVD(Mid$(X$, 35, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(BRULIS#), 4, 11, 2, 2), 55, 11)
          For NIVEL% = 0 To 3: SUMABRU#(NIVEL%) = SUMABRU#(NIVEL%) + BRULIS#: Next NIVEL%
          '
          DESCUE# = CVD(Mid$(X$, 43, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(DESCUE#), 4, 11, 2, 2), 66, 11)
          For NIVEL% = 0 To 3: SUMADES#(NIVEL%) = SUMADES#(NIVEL%) + DESCUE#: Next NIVEL%
          '
          DESPOR# = CVS(Mid$(X$, 51, 4))
          Call REPLA(TTXX$, CSTRING$(MKD$(DESPOR#), 4, 10, 1, 2), 77, 10)
          '
          NETOPE# = CVD(Mid$(X$, 55, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(NETOPE#), 4, 11, 2, 2), 87, 11)
          For NIVEL% = 0 To 3: SUMANET#(NIVEL%) = SUMANET#(NIVEL%) + NETOPE#: Next NIVEL%
          '
          COSREP# = CVD(Mid$(X$, 63, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(COSREP#), 4, 11, 2, 2), 98, 11)
          For NIVEL% = 0 To 3: SUMACOS#(NIVEL%) = SUMACOS#(NIVEL%) + COSREP#: Next NIVEL%
          '
          RENTAS# = CVD(Mid$(X$, 71, 8))
          Call REPLA(TTXX$, CSTRING$(MKD$(RENTAS#), 4, 11, 2, 2), 109, 11)
          For NIVEL% = 0 To 3: SUMAREN#(NIVEL%) = SUMAREN#(NIVEL%) + RENTAS#: Next NIVEL%
          '
          RENPOR# = CVS(Mid$(X$, 79, 4))
          Call REPLA(TTXX$, CSTRING$(MKD$(RENPOR#), 4, 9, 1, 2), 120, 9)
          '
          GoSub 50
          LINEAANT% = LINEA%
          GRUVEANT% = GRUVE%
          CODARANT% = CODAR%
          '
        End If
      End If
   Next U&
   GoSub 10
   GoSub 11
   GoSub 12
   GoSub 13         ' TOTAL GENERAL
   Call SETULTREG("LISTANA", JJ&)
   Call CIERRARCH("LISTANA")
   Exit Sub
   '
10 TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TRIM$(CODEXT$(CODARANT%)) + " - " + DESCRIT0$(CODARANT%), 1, 44)
   NIVEINI% = 3: GoSub 70
   TTXX$ = String$(248, "-"): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
11 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TITUGRU$, 1, 44)
   NIVEINI% = 2: GoSub 70
   TTXX$ = String$(248, "="): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
12 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "Tot. " + TITULIN$, 1, 44)
   NIVEINI% = 1: GoSub 70
   TTXX$ = String$(248, "*"): GoSub 50
   TTXX$ = ".": GoSub 50
   GoSub 60
Return
   '
13 'TTXX$ = ".": GoSub 50
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, "T O T A L   G E N E R A L ", 1, 44)
   NIVEINI% = 0: GoSub 70
   TTXX$ = ".": GoSub 50
   TTXX$ = String$(248, "*"): GoSub 50
   GoSub 60
Return
   '
   
20 TITULIN$ = "PRODUCTOS"
   If LINEA% = 2 Then TITULIN$ = "REPARACIONES"
   If LINEA% = 3 Then TITULIN$ = "DEBITOS - CREDITOS - VARIOS"
   '
   TTXX$ = TITULIN$: GoSub 50
   TTXX$ = String$(Len(TITULIN$), "*"): GoSub 50
   TTXX$ = ".": GoSub 50
Return
   '
21 TITUGRU$ = TRIM$(Str$(GRUVE%))
   While Len(TITUGRU$) < 3: TITUGRU$ = "0" + TITUGRU$: Wend
   TITUGRU$ = TITUGRU$ + " - "
   If GRUVE% = 999 Then
        TITUGRU$ = TITUGRU$ + "VARIOS"
      Else
        TITUGRU$ = TITUGRU$ + ECOARCH$("GRUCOVEN", MKI$(GRUVE%))
   End If
   TTXX$ = TITUGRU$: GoSub 50
   TTXX$ = ".": GoSub 50
Return
   '
50 JJ& = JJ& + 1
   Call GRAREG("LISTANA", TTXX$, JJ&)
Return
   '
60 For NIVEL% = NIVEINI% To 3
     SUMACAN#(NIVEL%) = 0
     SUMABRU#(NIVEL%) = 0
     SUMANET#(NIVEL%) = 0
     SUMADES#(NIVEL%) = 0
     SUMACOS#(NIVEL%) = 0
     SUMAREN#(NIVEL%) = 0
   Next NIVEL%
Return
   '
70 Call REPLA(TTXX$, CSTRING$(MKD$(SUMACAN#(NIVEINI%)), 4, 10, 0, 2), 45, 10)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMABRU#(NIVEINI%)), 4, 11, 2, 2), 55, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMADES#(NIVEINI%)), 4, 11, 2, 2), 66, 11)
   '
   DESPORX# = 0: If SUMABRU#(NIVEINI%) > 0 Then DESPORX# = 100 * SUMADES#(NIVEINI%) / SUMABRU#(NIVEINI%)
   Call REPLA(TTXX$, CSTRING$(MKD$(DESPORX#), 4, 10, 1, 2), 77, 10)
   '
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMANET#(NIVEINI%)), 4, 11, 2, 2), 87, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMACOS#(NIVEINI%)), 4, 11, 2, 2), 98, 11)
   Call REPLA(TTXX$, CSTRING$(MKD$(SUMAREN#(NIVEINI%)), 4, 11, 2, 2), 109, 11)
   '
   DESPORX# = 0: If SUMANET#(NIVEINI%) > 0 Then DESPORX# = 100 * SUMAREN#(NIVEINI%) / SUMANET#(NIVEINI%)
   Call REPLA(TTXX$, CSTRING$(MKD$(DESPORX#), 4, 9, 1, 2), 120, 9)
   '
   GoSub 50
   '
Return
   '
End Sub
'
Sub CALCUVIE()
    '
    FIN& = ULTREG&("MOVISTO")
    J& = 0: J1& = 0
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           CLI% = CVI(Mid$(X$, 84, 2))
           If CLI% > 0 Then
             If CLI% <= 32767 Then
               'If GRUCLIEN% = 3 Or TICLIE%(CLI%) = GRUCLIEN% Then
                 TIMO1$ = UCase$(Mid$(X$, 23, 2))
                 TIMO2$ = UCase$(Mid$(X$, 33, 2))
                 If TIMO1$ = "RT" Or TIMO2$ = "RT" Then
                    'PUNMO$(CLI%) = PUNMO$(CLI%) + MKS$(U&)
                    NRO1& = Val(Mid$(X$, 36, 7))
                    If NRO1& < 1 Then NRO1& = Val(Mid$(X$, 26, 7))
                    'GoSub 125
                 End If
               'End If
             End If
           End If
         End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MOVIREPA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVIREPA", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           CLI% = CVI(Mid$(X$, 84, 2))
           If CLI% > 0 Then
             If CLI% <= 32767 Then
               'If GRUCLIEN% = 3 Or TICLIE%(CLI%) = GRUCLIEN% Then
                 TIMO1$ = UCase$(Mid$(X$, 23, 2))
                 TIMO2$ = UCase$(Mid$(X$, 33, 2))
                 If TIMO1$ = "RT" Or TIMO2$ = "RT" Then
                    'PUNMO$(CLI%) = PUNMO$(CLI%) + MKS$((-1) * U&)
                    NRO1& = Val(Mid$(X$, 36, 7))
                    If NRO1& < 1 Then NRO1& = Val(Mid$(X$, 26, 7))
                    'GoSub 125
                 End If
               'End If
             End If
           End If
         End If
       End If
    Next U&
    '
    UREMO& = ULTREG&("MOVISTO")
    UREREPA& = ULTREG&("MOVIREPA")
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Y$ = REGLEIDO$("DEUDOR1", U&)
       CLI% = CVI(Mid$(Y$, 1, 2))
       If CLI% > 0 Then
          If CLI% <= 32767 Then
             'For KKI% = 1 To Len(PUNMO$(CLI%)) Step 4
             '   REG& = CVS(Mid$(PUNMO$(CLI%), KKI%, 4))
                If REG& > 0 And REG& <= UREMO& Then
                      X$ = REGLEIDO$("MOVISTO", REG&)
                   ElseIf REG& < 0 And Abs(REG&) <= UREREPA& Then
                      X$ = REGLEIDO$("MOVIREPA", Abs(REG&))
                   Else
                      GoTo 27
                End If
                CLI% = CVI(Mid$(X$, 84, 2))
                FF% = CVI(Left$(X$, 2))
                NRO1& = Val(Mid$(X$, 36, 7))
                If NRO1& < 1 Then NRO1& = Val(Mid$(X$, 26, 7))
                'GoSub 125
                CII% = CVI(Mid$(X$, 3, 2))
                REMIFAC$ = Mid$(X$, 33, 10)
                If TRIM$(REMIFAC$) = "" Then REMIFAC$ = Mid$(X$, 23, 10)
                CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                VUNI = CVD(Mid$(X$, 73, 8)) * COEF
                PRELI = PRELISTA#("LIPRE001", CII%)
                If PRELI < 0.1 Then PRELI = CVD(Mid$(X$, 73, 8))
                If REG& < 0 Then PRELI = 0

                VALOUN = VUNI
                      ' If VUNI < 0.00005 Then VALOUN = PRELI
                VATOT# = CANTU# * VALOUN
                      '
                Z$ = REGBLAN$("DETCLIEN")
                Call REPLA(Z$, MKI$(CLI%), 1, 2)
                Call REPLA(Z$, MKI$(FF%), 3, 2)
                Call REPLA(Z$, REMIFAC$, 5, 10)
                Call REPLA(Z$, MKI$(CII%), 15, 2)
                Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                Call REPLA(Z$, MKS$(VUNI), 25, 4)
                Call REPLA(Z$, MKS$(PRELI), 29, 4)
                Call REPLA(Z$, MKD$(VATOT#), 33, 8)
                '
                J& = J& + 1
                Call GRAREG("DETCLIEN", Z$, J&)
                '
                Z$ = REGBLAN$("LIANAFAC")
                GRCC% = GRUCOVEI%(CII%): If GRCC% < 1 Then GRCC% = 99
                Call REPLA(Z$, MKI$(1), 1, 2)    ' productos
                If REG& < 0 Then Call REPLA(Z$, MKI$(2), 1, 2)  ' productos
                Call REPLA(Z$, MKI$(GRCC%), 3, 2)
                Call REPLA(Z$, MKI$(CII%), 5, 2)
                Call REPLA(Z$, REMIFAC$, 7, 18)
                Call REPLA(Z$, MKI$(FF%), 25, 2)
                Call REPLA(Z$, MKD$(CANTU#), 27, 8)
                Call REPLA(Z$, MKD$(CANTU# * PRELI), 35, 8)
                Call REPLA(Z$, MKD$(CANTU# * (PRELI - VUNI)), 43, 8)
                PORDE = 0: If PRELI > 0 Then PORDE = 100 * (PRELI - VUNI) / PRELI
                Call REPLA(Z$, MKS$(PORDE), 51, 4)
                Call REPLA(Z$, MKD$(CANTU# * VUNI), 55, 8)
                COU = COSUNI(CII%)
                If REG& < 0 Then COU = 0
                Call REPLA(Z$, MKD$(CANTU# * COU), 63, 8)
                RENTA# = CANTU# * (VUNI - COU)
                Call REPLA(Z$, MKD$(RENTA#), 71, 8)
                POREN = 0: If Abs(VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
                Call REPLA(Z$, MKS$(POREN), 79, 4)
                '
                J1& = J1& + 1
                Call GRAREG("LIANAFAC", Z$, J1&)
                '
27           'Next KKI%
             'PUNMO$(CLI%) = ""
          End If
       End If
    Next U&
    '
    FIN& = 32767
    For U& = 1 To 32767
       Call TRACE(20, U&, FIN&)
       'If PUNMO$(U&) <> "" Then
          'For KKI% = 1 To Len(PUNMO$(U&)) Step 4
             'REG& = CVS(Mid$(PUNMO$(U&), KKI%, 4))
                If REG& > 0 And REG& <= UREMO& Then
                      X$ = REGLEIDO$("MOVISTO", REG&)
                   ElseIf REG& < 0 And Abs(REG&) <= UREREPA& Then
                      X$ = REGLEIDO$("MOVIREPA", Abs(REG&))
                   Else
                      GoTo 37
                End If
                   X$ = REGLEIDO$("MOVISTO", REG&)
                   CLI% = CVI(Mid$(X$, 84, 2))
                   FF% = CVI(Left$(X$, 2))
                   CII% = CVI(Mid$(X$, 3, 2))
                   NRO1& = Val(Mid$(X$, 36, 7))
                   If NRO1& < 1 Then NRO1& = Val(Mid$(X$, 26, 7))
                   'GoSub 125
                   REMIFAC$ = Mid$(X$, 33, 10)
                   If TRIM$(REMIFAC$) = "" Then REMIFAC$ = Mid$(X$, 23, 10)
                   CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                   VUNI = CVD(Mid$(X$, 73, 8)) * COEF
                   PRELI = PRELISTA#("LIPRE001", CII%)
                   If PRELI < 0.1 Then PRELI = CVD(Mid$(X$, 73, 8))
                   If REG& < 0 Then PRELI = 0
                   VALOUN = VUNI
                   ' If VUNI < 0.00005 Then VALOUN = PRELI
                   VATOT# = CANTU# * VALOUN
                   '
                   Z$ = REGBLAN$("DETCLIEN")
                   Call REPLA(Z$, MKI$(CLI%), 1, 2)
                   Call REPLA(Z$, MKI$(FF%), 3, 2)
                   Call REPLA(Z$, REMIFAC$, 5, 10)
                   Call REPLA(Z$, MKI$(CII%), 15, 2)
                   Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                   Call REPLA(Z$, MKS$(VUNI), 25, 4)
                   Call REPLA(Z$, MKS$(PRELI), 29, 4)
                   Call REPLA(Z$, MKD$(VATOT#), 33, 8)
                   '
                   J& = J& + 1
                   Call GRAREG("DETCLIEN", Z$, J&)
                   '
                   Z$ = REGBLAN$("LIANAFAC")
                   GRCC% = GRUCOVEI%(CII%): If GRCC% < 1 Then GRCC% = 99
                   Call REPLA(Z$, MKI$(1), 1, 2)    ' productos
                   If REG& < 0 Then Call REPLA(Z$, MKI$(2), 1, 2)  ' productos
                   Call REPLA(Z$, MKI$(GRCC%), 3, 2)
                   Call REPLA(Z$, MKI$(CII%), 5, 2)
                   Call REPLA(Z$, REMIFAC$, 7, 18)
                   Call REPLA(Z$, MKI$(FF%), 25, 2)
                   Call REPLA(Z$, MKD$(CANTU#), 27, 8)
                   Call REPLA(Z$, MKD$(CANTU# * PRELI), 35, 8)
                   Call REPLA(Z$, MKD$(CANTU# * (PRELI - VUNI)), 43, 8)
                   PORDE = 0: If PRELI > 0 Then PORDE = 100 * (PRELI - VUNI) / PRELI
                   Call REPLA(Z$, MKS$(PORDE), 51, 4)
                   Call REPLA(Z$, MKD$(CANTU# * VUNI), 55, 8)
                   COU = COSUNI(CII%)
                   If REG& < 0 Then COU = 0
                   Call REPLA(Z$, MKD$(CANTU# * COU), 63, 8)
                   RENTA# = CANTU# * (VUNI - COU)
                   Call REPLA(Z$, MKD$(RENTA#), 71, 8)
                   POREN = 0: If Abs(VUNI) > 0 Then POREN = 100 * RENTA# / (CANTU# * VUNI)
                   Call REPLA(Z$, MKS$(POREN), 79, 4)
                   '
                   J1& = J1& + 1
                   Call GRAREG("LIANAFAC", Z$, J1&)
                   '
37        'Next KKI%
       'End If
    Next U&
    '
    Call SETULTREG("DETCLIEN", J&)
    Call CIERRARCH("DETCLIEN")
    Call LIBARCH("DETCLIEN")
    '
    Call SETULTREG("LIANAFAC", J&)
    Call CIERRARCH("LIANAFAC")
    Call LIBARCH("LIANAFAC")
    '
End Sub

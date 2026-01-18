VERSION 5.00
Begin VB.Form AVAFAC04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avance de Facturación y Cobranza"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   20
      Top             =   3465
      Width           =   3900
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
         Height          =   315
         Left            =   1575
         TabIndex        =   21
         Top             =   210
         Width           =   175
      End
      Begin VB.TextBox TEXT2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         TabIndex        =   22
         Text            =   " "
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   -650
         TabIndex        =   24
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   23
         Top             =   210
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Agrupamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   17
      Top             =   4320
      Width           =   3900
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Por Cliente"
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
         Left            =   2205
         TabIndex        =   19
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Por Fecha"
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
         Left            =   630
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   6500
      Left            =   4365
      ScaleHeight     =   6435
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   0
      Width           =   960
      Begin VB.PictureBox Picture16 
         Height          =   5960
         Left            =   0
         ScaleHeight     =   5895
         ScaleWidth      =   1020
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.CommandButton Command5 
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
            Height          =   750
            Left            =   120
            Picture         =   "AVAFAC04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Aprueba y Graba"
            Top             =   4680
            Width           =   650
         End
         Begin VB.Image Image3 
            Height          =   6000
            Left            =   0
            Picture         =   "AVAFAC04.frx":030A
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "AVAFAC04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Detalle de Facturación por Rango de Fechas."
         Top             =   1995
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
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
         Picture         =   "AVAFAC04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Resúmen Operativo"
         Top             =   1155
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   5040
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTREC 
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
         Picture         =   "AVAFAC04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   4200
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "AVAFAC04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AVAFAC04.frx":310A
         Top             =   5420
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   210
      TabIndex        =   6
      Top             =   945
      Width           =   3900
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Notas de Crédito"
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
         Index           =   2
         Left            =   630
         TabIndex        =   11
         Top             =   1470
         Value           =   1  'Checked
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Notas de Débito"
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
         Index           =   1
         Left            =   630
         TabIndex        =   10
         Top             =   1155
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Facturas (Contado y C.Corriente)"
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
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cobranza por Recibo"
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
         Left            =   315
         TabIndex        =   8
         Top             =   1995
         Width           =   2850
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Facturación"
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
         Left            =   315
         TabIndex        =   7
         Top             =   465
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   3
      Top             =   5160
      Width           =   3900
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
         Left            =   630
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
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
         Left            =   2205
         TabIndex        =   4
         Top             =   315
         Width           =   1275
      End
   End
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
      Height          =   360
      Left            =   3930
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox MESAN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1365
      TabIndex        =   0
      Text            =   " "
      ToolTipText     =   "Doble Click para Seleccionar Período"
      Top             =   315
      Width           =   2580
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mes y año:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   -890
      TabIndex        =   2
      Top             =   420
      Width           =   2220
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobarseleccion 
         Caption         =   "Aprobar Selección"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu cn 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu resumenoperativo 
         Caption         =   "Resumen Operativo"
      End
      Begin VB.Menu detalledefacturacionporrangodefechas 
         Caption         =   "Detalle de Facturación por Rango de Fechas."
      End
   End
End
Attribute VB_Name = "AVAFAC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDE%, HASTA%, CORRESP$
Dim DESDAT As Double
Dim HASDAT As Double
'
Sub RESUOPER()
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#(), GASTOS#()
      '
      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      ReDim GASTOS#(URE&)
      '
      i& = 0
      For i& = 1 To URE&
          RELE$ = REGLEIDO$("SELMESA", i&)
          PERIOD$(i&) = MESANO$(RELE$, DES%, HAS%)
          DESDEI%(i&) = DES%
          HASTAI%(i&) = HAS%
          '\\\OT-08-0596-OD-30/07/08///
          If DESDEX% > DES% And DESDEX% < HAS% Then
             PERIOD$(i&) = MESANO$(RELE$, DES%, HAS%) + " (Inc.)"
             DESDEI%(i&) = DESDEX%
          End If
          If HASTAX% > DES% And HASTAX% < HAS% Then
             PERIOD$(i&) = MESANO$(RELE$, DES%, HAS%) + " (Inc.)"
             HASTAI%(i&) = HASTAX%
          End If
          '\\\OT-08-0596-OD-FIN///
      Next i&
      '
      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + "AND  NuClien > 0 "
      '
      Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With AVANCE
          On Error Resume Next
          .MoveFirst
          If Err Then
             On Error GoTo 0
             Screen.MousePointer = 1
             GoTo 10
          End If
          '
          Do While Not .EOF
               IMPO# = !IDEBECOMP - !IHABECOMP
               FFI% = CVINT%(!FechConta)
               '
               J& = 0
               For J& = 1 To URE&
                   
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      '
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + "AND  NuClien > 0 "
       '
       Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 20
           End If
          '
           Do While Not .EOF
                IMPO# = !IHABECOMP - !IDEBECOMP
                FFI% = CVINT%(!FechConta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 30
           End If
           '
           Do While Not .EOF
                IMPO# = !IHABECOMP - !IDEBECOMP
                FFI% = CVINT%(!FechConta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT * FROM CAJABANC "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 40
           End If
           '
           Do While Not .EOF
             IMPO# = !IDEBECOMP - !IHABECOMP
             FFI% = CVINT%(!FechConta)
             '
             J& = 0
             For J& = 1 To URE&
                 If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                    '\\\OT-08-0595-OD-29/07/08///
                    If TICUEPRO%((!NuProve) * (-1)) = 0 Then
                        PAGOS#(J&) = PAGOS#(J&) + IMPO#
                      Else
                        GASTOS#(J&) = GASTOS#(J&) + IMPO#
                    End If
                    '\\\OT-08-0595-OD-FIN///

                    Exit For
                 End If
             Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
40     Call BLANARCH("EVOLMEN")
       Y$ = REGBLAN$("EVOLMEN")
       J& = 0
       '\\\OT-08-0596-OD-30/07/08///
       FEX = DESDEX%
       DES% = COMIENFE%(FEX)
       'DES% = DESDEX%
       HAS% = HASTAX%
       '\\\OT-08-0596-OD-FIN///
       For k& = 1 To URE&
          '
          If DESDEI%(k&) >= DES% And DESDEI%(k&) <= HAS% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             TGAS# = GASTOS#(k&)
             RES1# = 0: RES2# = 0
             '
             '\\\OT-08-0595-OD-29/07/08///
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             Call REPLA(Y$, MKD$(TGAS#), 33, 8)
             RES1# = TVEN# - TCOM# - TGAS#
             Call REPLA(Y$, MKD$(RES1#), 41, 8)
             Call REPLA(Y$, MKD$(TCOB#), 49, 8)
             Call REPLA(Y$, MKD$(TPAG#), 57, 8)
             Call REPLA(Y$, MKD$(TGAS#), 65, 8)
             RES2# = TCOB# - TPAG# - TGAS#
             Call REPLA(Y$, MKD$(RES2#), 73, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 81, 2)
             J& = J& + 1
             Call GRAREG("EVOLMEN", Y$, J&)
            '\\\OT-08-0595-OD-FIN///
          End If
          '
       Next k&
       '
       Call SETULTREG("EVOLMEN", J&)
       '
       Call FILESORT("EVOLMEN", "", 10, 10, "DESC")
       Call HEADERS("EVOLMEN", "")
       Call HEADERS("EVOLMEN", "Corresponde a: " + PERIO$)
       '
       Call FINAL("*EVOLMEN")
       '
End Sub

Private Sub aprobarseleccion_Click()
   Call BOTREC_Click
End Sub

Private Sub BOTREC_Click()
    '
    '\\\OT-05-0668-WAR-12/10/05///
    If VALIMESAN% < 1 Then GoTo 99
    '
    Call BLANARCH("AVAFACB/NOERR")
    GRUPO% = 0
    If Option6.Value And Option2.Value Then    ' por cliente e impresion
       OPTNE% = COMALTER%("Seleccion de Clientes:\\Todos\Nacionales\Extranjeros")
       If OPTNE% < 1 Then GoTo 99
       '
       CantRegistrosX& = ULTREG&("GRUCOCLI")
       If CantRegistrosX& > 0 Then
         Call SELECHO("GRUCOCLI/Seleccion por Grupo (Cancel para Todos)")
         GRUPO% = XVALO(VALACT1$("GRUCOCLI"))
         DENOGRUPO$ = VALACT2$("GRUCOCLI")
       End If
       '
    End If
10  Screen.MousePointer = 11
    'nuevo
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where (INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & ") "
    '
    If Option3.Value = True Then                  ' facturas debitos y creditos
      OP% = COMALTER%("Opciones de Trabajo:\\Importes Brutos\Importes Netos (Sin Iva)")
          If OP% = 1 Then
            SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
          ElseIf OP% = 2 Then
            SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
          Else
            Screen.MousePointer = 1
            GoTo 99
          End If
          Screen.MousePointer = 1
       ElseIf Option4.Value = True Then           ' recibo de cobranza
          SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
    End If
    'ORDEN
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'ESCRITURA
    With AVANCETMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Screen.MousePointer = 1
          Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
          Exit Sub
        End If
        JJ& = 0
        REBLAY$ = REGBLAN$("AVAFACB")
        REBLAX$ = REGBLAN$("AVAFACC")
        '
        Do While Not .EOF
          If !OPCIMOVIM = "FC" And Check1(0).Value <> 1 Then GoTo 49
          If !OPCIMOVIM = "ND" And Check1(1).Value <> 1 Then GoTo 49
          If !OPCIMOVIM = "NC" And Check1(2).Value <> 1 Then GoTo 49
          If TipoMovim = "RCOB" And Option4.Value = False Then GoTo 49
          '
          NC% = !NuClien
          CLIEX$ = ""
          If NC% > 0 Then
             POSIVA% = PIVACLI%(NC%)
             If OPTNE% = 3 And POSIVA% <> 5 Then GoTo 49
             If OPTNE% = 2 And POSIVA% = 5 Then GoTo 49
             '
             NUMVEN% = VENDCLI%(NC%)
             ELEVEN% = XVALO(Text2) 'VENDEDOR ELEGIDO
             GRUPOCLI% = GRUCOCLI%(NC%)
             If (ELEVEN% = 0 Or ELEVEN% = NUMVEN%) And (GRUPOCLI% = GRUPO% Or GRUPO% = 0) Then
             '\\\OT-06-0302-WAR-FIN///
                CLIEX$ = RASOCLI$(NC%)
                NRO& = !NumeComp
                NETO# = !ValAbsComp 'Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
                  If !OPCIMOVIM = "NC" Then NETO# = (-1) * NETO#
                NUSUCUVE% = !CodiSucu
                NNX$ = !CODICOM_LAR
                MONEM% = 0: TICAM = 1
                On Error Resume Next
                MONEM% = !Mone_Emis
                On Error GoTo 0
                '
                TICAM = 1
                On Error Resume Next
                If IsNull(!Tipo_Cam) = False Then
                  TICAM = !Tipo_Cam
                End If
                On Error GoTo 0
                '
                Y$ = REBLAY$
                Call REPLA(Y$, MKI$(CVINT(!FechConta)), 1, 2)
                Call REPLA(Y$, CLIEX$, 3, 28)
                Call REPLA(Y$, NNX$, 31, 18)
                Call REPLA(Y$, MKD$(NETO#), 49, 8)
                '
                If MONEM% > 1 Then
                  If TICAM > 0.1 Then
                    Call REPLA(Y$, SIMBOLPES$(MONEM%), 65, 4)
                    Call REPLA(Y$, MKD$(NETO# / TICAM), 69, 8)
                  End If
                End If
                '
                JJ& = JJ& + 1
                Call GRAREG("AVAFACB", Y$, JJ&)
                '\\\OT-05-0668-WAR-12/10/05///
                XX$ = REBLAX$
                Call REPLA(XX$, MKI(NC%), 1, 2)
                Call REPLA(XX$, CLIEX$, 3, 30)
                Call REPLA(XX$, MKI$(CVINT(!FechConta)), 33, 2)
                Call REPLA(XX$, NNX$, 35, 24)
                Call REPLA(XX$, MKD$(NETO#), 59, 8)
                '
                If MONEM% > 1 Then
                  If TICAM > 0.1 Then
                    Call REPLA(XX$, SIMBOLPES$(MONEM%), 75, 4)
                    Call REPLA(XX$, MKD$(NETO# / TICAM), 79, 8)
                  End If
                End If
                '
                Call GRAREG("AVAFACC", XX$, JJ&)
                '\\\OT-05-0668-WAR-FIN///
             End If
          End If
49      .MoveNext
        Loop
    End With
    '
    Call SETULTREG("AVAFACB", JJ&)
    Call SETULTREG("AVAFACC", JJ&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    VEND1$ = "Todos Los Vendedores"
    ECOVEN1$ = TRIM$(Label23)
    If ELEVEN% > 0 And ECOVEN1$ <> "" Then
      VEND1$ = ECOVEN1$
    End If
    '
    '
    If Option6 Then                          ' agrupa por cliente
       If Option2 Then                       ' salida por impresora
          Call FILESORT("AVAFACC", "AVAFACC", 3, 3, "ASC")
          Call FILESORT("AVAFACC", "AVAFACC", 1, 1, "ASC")
          Call HEADERS("AVAFACC", "")
          Call HEADERS("AVAFACC", "Corresponde a: " & CORRESP$ & " - " & VEND1$)
          If GRUPO% > 0 Then Call HEADERS("AVAFACC", "Grupo de Cliente: " & GRUPO% & " - " & DENOGRUPO$)
          '
          posicion$ = "Todos"
          If OPTNE% = 2 Then posicion$ = "Nacionales"
          If OPTNE% = 3 Then posicion$ = "Extranjeros"
          Call HEADERS("AVAFACC", "Filtro: " & posicion$)
          '
          Call FINAL("*LIAVAFC")
        ElseIf Option1 Then                  ' salida por pantalla
          Call FINAL("*COVECLIE")
       End If
       GoTo 50
    End If
    '
    If Option1.Value = True Then                  ' salida por pantalla
         Call FINAL("*COAVAFB")
       ElseIf Option2.Value = True Then           ' salida por impresora
         Call HEADERS("AVAFACB", "")
         Call HEADERS("AVAFACB", "Corresponde a: " & CORRESP$ & " - " & VEND1$)
         Call FINAL("*LIAVAFB")
    End If
    '
50  Hide
    '
99  End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
  '
  Call SELECHO("VENDEDOR")
  VDEVU$ = VALACT1$("VENDEDOR")
  If TRIM$(VDEVU$) <> "" Then
    Text2.TEXT = VDEVU$
  End If
  '
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.TEXT = MESASEL$
End Sub



Private Sub Command4_Click()
   Command4.Enabled = False
   Call REMEVE04.Command15_Click
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Call BOTREC_Click
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   OPT% = ALTERNA%("Resumen Operativo\Otros Accesos Directos")
   '
   If OPT% = 1 Then
      Call RESUOPER
   End If
   '
   Command8.Enabled = True
End Sub

Private Sub detalledefacturacionporrangodefechas_Click()
    Call Command4_Click
End Sub

Private Sub Form_Load()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Check1(0).BackColor = &HFCD7B6
      Me.Check1(1).BackColor = &HFCD7B6
      Me.Check1(2).BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6
      Me.Option6.BackColor = &HFCD7B6

      MARCOBARRA.Top = 5600
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If
    If CONTROL("FACTURA", "SISTMANT") = "SI" Then
        resumenoperativo.Visible = False
    End If

End Sub

Private Sub MESAN_DblClick()
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   MESAN = PERIO$
   '
End Sub

Private Sub Option3_Click()
   For KI% = 0 To 2
     Check1(KI%).Enabled = True
     Check1(KI%).Value = 1
   Next KI%
End Sub

Private Sub Option4_Click()
   For KI% = 0 To 2
     Check1(KI%).Value = 0
     Check1(KI%).Enabled = False
   Next KI%
End Sub
'\\\OT-05-0668-WAR-12/10/05///
Function VALIMESAN%()
   VMA% = 0
   '
   If Len(MESAN) = 19 Then
     DESDE% = FECHANUM(Left$(MESAN, 8))
     If DESDE% > 0 Then
       HASTA% = FECHANUM(Mid$(MESAN, 12, 8))
       If HASTA% >= DESDE% Then
         VMA% = 1
         DESDAT = Int(CDbl(CVDAT(DESDE%)))
         HASDAT = Int(CDbl(CVDAT(HASTA%)))
         CORRESP$ = MESAN
         GoTo 98
       End If
     End If
   End If
   '
   CORRESP$ = MESANO$(MESAN, DESDE%, HASTA%)
   If CORRESP$ <> "" Then
       VMA% = 1
       DESDAT = Int(CDbl(CVDAT(DESDE%)))
       HASDAT = Int(CDbl(CVDAT(HASTA%)))
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   '
98 VALIMESAN% = VMA%
   '
End Function

Private Sub resumenoperativo_Click()
   Call Command8_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

'\\\OT-05-0668-WAR-FIN///
Private Sub Text2_Change()
  '
  Label23 = TRIM$(ECOARCH$("VENDEDOR", MKI$(XVALO(Text2.TEXT))))
  '
End Sub

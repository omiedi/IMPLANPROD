VERSION 5.00
Begin VB.Form QUERISTO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Consultas y Listados"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8925
   Icon            =   "QUERISTO.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Statist"
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
      Left            =   7930
      Picture         =   "QUERISTO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Análisis Estadístico de Consumos de Stock"
      Top             =   2800
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7820
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   32
      Top             =   3660
      Width           =   965
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Materiales Reservados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   28
      Top             =   1680
      Width           =   2640
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   29
         Top             =   420
         Width           =   1800
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
            Picture         =   "QUERISTO.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Reservas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   31
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Inventario Continuo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   24
      Top             =   240
      Width           =   2640
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   25
         Top             =   420
         Width           =   1800
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
            Picture         =   "QUERISTO.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Material en Depósitos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   27
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   19
      Top             =   3165
      Width           =   2640
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   20
         Top             =   420
         Width           =   1800
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
            Picture         =   "QUERISTO.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Parte Diario de Fabricacion por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Partes de Fabricación"
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
            TabIndex        =   21
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.CommandButton Command27 
      Caption         =   "More"
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
      Left            =   7930
      Picture         =   "QUERISTO.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Complementos, Estadísticas, Despachos de Importación"
      Top             =   1975
      Width           =   750
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
      Height          =   750
      Left            =   7930
      Picture         =   "QUERISTO.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1155
      Width           =   750
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
      Height          =   750
      Left            =   7930
      Picture         =   "QUERISTO.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   345
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Control de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Listados de Control"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1890
         Width           =   4215
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0FF&
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
            Height          =   600
            Left            =   420
            TabIndex        =   16
            Top             =   1320
            Width           =   3375
            Begin VB.OptionButton Option4 
               BackColor       =   &H00E0E0FF&
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
               Left            =   1890
               TabIndex        =   4
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00E0E0FF&
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
               Left            =   315
               TabIndex        =   3
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   14
            Top             =   420
            Width           =   3375
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
               Left            =   2770
               Picture         =   "QUERISTO.frx":1690
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Disponibilidad General de Stocks"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command30 
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
               Left            =   1680
               Picture         =   "QUERISTO.frx":199A
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Existencias Negativas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command20 
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
               Left            =   2215
               Picture         =   "QUERISTO.frx":1CA4
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command19 
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
               Picture         =   "QUERISTO.frx":1FAE
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Existencias Físicas Actuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias Negativos Críticos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   15
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Consultas / Listados Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   12
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   17
            Top             =   420
            Width           =   3375
            Begin VB.CommandButton Command2 
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
               Picture         =   "QUERISTO.frx":22B8
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Listado de Existencias y Ubicaciones"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2760
               Picture         =   "QUERISTO.frx":25C2
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Gestion de Reservas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command9 
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
               Left            =   2200
               Picture         =   "QUERISTO.frx":28CC
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ficha Kardex - Histórico de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - His- torico - Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   18
               Top             =   105
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7320
      Picture         =   "QUERISTO.frx":2BD6
      Top             =   3910
      Width           =   2010
   End
End
Attribute VB_Name = "QUERISTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command14_Click()
    OPRESERVA.Show 1
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call EXISTO2
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call LIUBIEXI
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call CRITSTO2
    If ULTREG&("CRITSTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*CRITSTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*CRITSTL", "")
        End If
    End If
    Command20.Enabled = True
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EXISTE%("TRANSUCU.EXE") > 0 Then
        Call CONECRUN("TRANSUCU", "Transferencias de Sucursal")
      Else
        Call ACCDIR("CONSTOCK")
    End If
    Command27.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    OPXX% = ALTERNA%("Consulta por Día\Totales por Rango de Fechas\Control Desvíos de Consumo")
    On OPXX% GoTo 10, 20, 30
    GoTo 99
    '
10  HOI% = HOY(HO$)
    JJ& = 0
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHINVE$ <> "" Then
        FF% = FECHANUM(FECHINVE$)
        If FF% > 0 Then
            Screen.MousePointer = 11
            FIN& = ULTREG&("MOVISTO")
            For U& = 1 To FIN&
                Call TRACE(20, U&, FIN&)
                X$ = REGLEIDO$("MOVISTO", U&)
                FEMO% = CVI(Left$(X$, 2))
                If FEMO% = FF% Then
                    If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
                        CANTID# = CVD(Mid$(X$, 88, 8))
                        If CANTID# >= 0.05 Then
                            CI% = CVI(Mid$(X$, 3, 2))
                            REFE$ = Mid$(X$, 43, 30)
                            NOFA$ = Mid$(X$, 23, 10)
                            LOTE$ = Mid$(X$, 5, 16)
                            '
                            Y$ = REGBLAN$("PARPROD")
                            Call REPLA(Y$, MKI$(CI%), 1, 2)
                            Call REPLA(Y$, MKD$(CANTID#), 3, 8)
                            Call REPLA(Y$, REFE$, 11, 32)
                            Call REPLA(Y$, NOFA$, 43, 10)
                            Call REPLA(Y$, LOTE$, 53, 16)
                            JJ& = JJ& + 1
                            Call GRAREG("PARPROD", Y$, JJ&)
                        End If
                    End If
                End If
            Next U&
            Call CIERRARCH("PARPROD")
            Screen.MousePointer = 1
            '
            If JJ& < 1 Then
                Call COMUNI("No hay Partes de Producción Registrados\en la Fecha Elegida " + FECHINVE$ + ".")
                GoTo 99
            End If
            '
            Call HEADERS("PARPROD", "")
            Call HEADERS("PARPROD", "Fecha: " + FECHINVE$)
            Call FINAL("*LPARPRO")
            '
        End If
    End If
    GoTo 99
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    Dim CAFA#(32767)
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FEMO% = CVI(Left$(X$, 2))
       If FEMO% >= FF1% Then
          If FEMO% <= FF2% Then
             If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
                CANTID# = CVD(Mid$(X$, 88, 8))
                If CANTID# >= 0.05 Then
                   CI% = CVI(Mid$(X$, 3, 2))
                   If CI% > 0 Then
                      If CI% <= 32767 Then
                         If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                            CODE$ = CODEXT$(CI%)
                            If CODE$ >= cod1$ Then
                               If CODE$ <= cod2$ Then
                                  CAFA#(CI%) = CAFA#(CI%) + CANTID#
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
    Next U&
    '
    Call BLOQARCH("TOTPROD")
    JJ& = 0
    FIN& = NUINV%
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("INVERT", U& + 1)
       CI% = CVI(Mid$(X$, 17, 2))
       If CAFA#(CI%) >= 0.05 Then
          Z$ = REGBLAN$("TOTPROD")
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          Call REPLA(Z$, MKI$(TIMATE%(CI%)), 3, 2)
          Call REPLA(Z$, MKD$(CAFA#(CI%)), 5, 8)
          JJ& = JJ& + 1
          Call GRAREG("TOTPROD", Z$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("TOTPROD", JJ&)
    Call CIERRARCH("TOTPROD")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
       Call COMUNI("No hay Partes de Producción Registrados\en el Rango de Fechas Elegido.")
       GoTo 99
    End If
    '
    Call HEADERS("TOTPROD", "")
    Call HEADERS("TOTPROD", "Rango: " + RANGO$)
    Call FINAL("*LTOTPRO")
    '
    GoTo 99

30  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
32  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 32
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    REBLA0$ = REGBLAN$("COMPACON")
    JJ& = 0
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("MOVISTO", U&)
      FEMO% = CVI(Left$(X$, 2))
      If FEMO% >= FF1% Then
        If FEMO% <= FF2% Then
          If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
            CANTID# = CVD(Mid$(X$, 88, 8))
            If CANTID# >= 0.05 Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                If CI% <= 32767 Then
                  If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                    CODE$ = CODEXT$(CI%)
                    If CODE$ >= cod1$ Then
                      If CODE$ <= cod2$ Then
                        REBLA$ = REBLA0$
                        Call REPLA(REBLA$, MKI$(FEMO%), 1, 2)
                        Call REPLA(REBLA$, Mid$(X$, 23, 10), 3, 10)
                        Call REPLA(REBLA$, MKI$(CI%), 13, 2)
                        REBLA1$ = REBLA$
                        CABAS = CANTID#
                        Call LEEEXPLO(CODE$, CABAS)
                        For KKI% = 1 To CAIT%
                          If Abs(CABAS * USOI(KKI%)) > 0.005 Then
                            YY$ = REBLA$
                            Call REPLA(YY$, MKI$(CIJ%(KKI%)), 15, 2)
                            Call REPLA(YY$, MKS$(CABAS * USOI(KKI%)), 17, 4)
                            Call REPLA(YY$, MKS$(0), 21, 4)
                            Call REPLA(YY$, MKS$(0), 25, 4)
                            JJ& = JJ& + 1
                            Call GRAREG("COMPACON", YY$, JJ&)
                          End If
                        Next KKI%
                      End If
                    End If
                  End If
                End If
              End If
              GoTo 89
            End If
          End If
          '
          If UCase$(Mid$(X$, 21, 2)) = "CF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
            CANTID# = CVD(Mid$(X$, 96, 8))
            If CANTID# >= 0.05 Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                If CI% <= 32767 Then
                  CODE$ = CODEXT$(CI%)
                  YY$ = REBLA1$
                  Call REPLA(YY$, Mid$(X$, 23, 10), 3, 10)
                  Call REPLA(YY$, MKI$(CI%), 15, 2)
                  Call REPLA(YY$, MKS$(0), 17, 4)
                  CARESAL = CANTID#
                  Call REPLA(YY$, MKS$(CARESAL), 21, 4)
                  Call REPLA(YY$, MKS$(0), 25, 4)
                  JJ& = JJ& + 1
                  Call GRAREG("COMPACON", YY$, JJ&)
                End If
              End If
            End If
          End If
        End If
      End If
89  Next U&
    '
    Call SETULTREG("COMPACON", JJ&)
    Call CIERRARCH("COMPACON")
    '
    Call HEADERS("COMPACON", "")
    Call HEADERS("COMPACON", "Rango: " + RANGO$)
    Call FINAL("*SOCOCON")
    GoTo 99
    '
99  Command3.Enabled = True
    '
End Sub
'


Private Sub Command4_Click()
   Command4.Enabled = False
   If EXISTE%("ESTSUC02.EXE") > 0 Then
     OPX% = ALTERNA%("Análisis Estadístico de Consumo\Estadistica de Entregas por Sede\Detalle de Movimiento por Comprobante")
     If OPX% < 1 Or OPX% > 3 Then GoTo 99
     If OPX% = 1 Then GoTo 90
     If OPX% = 3 Then GoTo 80
     Call CONECRUN("ESTSUC02", "Estadística de Entregas por Sucursal")
     GoTo 99
   Else
     OPX% = ALTERNA%("Análisis Estadístico de Consumo\Detalle de Movimiento por Comprobante")
     If OPX% < 1 Or OPX% > 2 Then GoTo 99
     If OPX% = 1 Then GoTo 90
     If OPX% = 2 Then GoTo 80
     GoTo 99
   End If
   '
80 Call DESHASFECHA(DES%, HAS%, PERIO$)
   PERIODO$ = PERIO$
   Screen.MousePointer = 11
   F& = 0
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
   Call TRACE(24, U&, FIN&)
      X$ = REGLEIDO$("MOVISTO", U&)
      FECH% = CVI(Mid$(X$, 1, 2))
      If FECH% >= DES% And FECH% <= HAS% Then
          COMP$ = Mid$(X$, 23, 10)
          CI% = CVI(Mid$(X$, 3, 2))
          CANTENT# = CVD(Mid$(X$, 88, 8))
          CANTSAL# = CVD(Mid$(X$, 96, 8))
          TOTAL# = 0 ' CVI(Mid$(X$, 31, 8))
          '
          F& = F& + 1
          XX$ = REGBLAN$("DETMOVC")
          '
          Call REPLA(XX$, COMP$, 1, 10)
          Call REPLA(XX$, MKI$(FECH%), 11, 2)
          Call REPLA(XX$, MKI$(CI%), 13, 2)
          Call REPLA(XX$, MKD(CANTENT#), 15, 8)
          Call REPLA(XX$, MKD(CANTSAL#), 23, 8)
          Call REPLA(XX$, MKD(TOTAL#), 31, 8)
          Call GRAREG("DETMOVC", XX$, F&)
      End If
   Next U&
   Call HEADERS("DETMOVC", "")
   Call HEADERS("DETMOVC", "Corresponde al Periodo: " + PERIODO$)
   Call SETULTREG("DETMOVC", F&)
   Call FILESORT("DETMOVC", "", 1, 2, "ASC")
   Call CIERRARCH("DETMOVC")
   Screen.MousePointer = 1
   Call CONECRUN("*DETMOVC", "Detalle de Movimiento por Comprobante")
   GoTo 99
   '
90 Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    DISPOSTO.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command30_Click()
    Command30.Enabled = False
    Call NEGASTO2
    If ULTREG&("NEGASTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*NEGASTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*NEGASTL", "")
        End If
    End If
    Command30.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    ACONEC$ = "INVENT02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "INVENTAR"
    End If
    Call CONECRUN(ACONEC$, "Inventarios Físicos y Valorizados")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    OPRESERVA.Show 1
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX", "Ficha Kardex de Movimientos")
      Else
        Call GENKARDEX("")
    End If
    Command9.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(UCase$(EPAC$), "ARCON") > 0 Then
       Command27.ToolTipText = "Movimientos de Sucursal"
    End If
    '
    If EXISTE%("ESTOCK02.EXE") > 0 Then
      Command4.Enabled = True
    End If
    '
    Call GRATITFOR(Caption)
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Label5_Click()
    LU$ = LUDAT$
    If EXISTE%(LU$ + "GETRAMO.DAT") > 0 Then
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("GETRAMO")
            X$ = REGLEIDO$("GETRAMO", U&)
            COD$ = Left$(X$, 16)
            CI% = CODINT%(COD$)
            Y$ = Mid$(X$, 17, 128)
            Call REPLA(Y$, MKI$(CI%), 3, 2)
            Call REGAPP("MOVISTO", Y$)
        Next U&
        Call CIERRARCH("*.*")
        Call FILEKILL(LU$ + "GETRAMO.DAT")
        Screen.MousePointer = 1
    End If
End Sub
'
Sub IMPORNEU2()
   '
   If ULTREG&("MOVISTO") < 1 Then GoTo 3
   LU$ = LUDAT$
   If EXISTE%(LU$ + "MOVISTX.DAT") = 1 Then
      Screen.MousePointer = 11
      Call COPYSTRU("MOVISTO", "MOVISTX")
      FIN& = ULTREG&("MOVISTX")
      For U& = 1 To FIN&
         XXX$ = REGLEIDO$("MOVISTX", U&)
         CIXI% = CVI(Mid$(XXX$, 3, 2))
         Call REGAPP("MOVISTO", XXX$)
         Call CIERRARCH("MOVISTO")
         Call VERISAL(CIXI%)
         Call VESARRAS(CIXI%)
      Next U&
   End If
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
3  Screen.MousePointer = 11
   Call BLANARCH("MOVISTO")
   '
   Dim SALSTO(16384)
   Dim CODEPOS$(64)
   Dim ALIDEP$(128), CODEPX$(128)
   '
   ALIDEP$(1) = "FK": CODEPX$(1) = "5110000"
   ALIDEP$(2) = "DEP.": CODEPX$(2) = "4080000"
   ALIDEP$(3) = "BS": CODEPX$(3) = "5120000"
   ALIDEP$(4) = "510": CODEPX$(4) = "5100000"
   ALIDEP$(5) = "513": CODEPX$(5) = "5130000"
   ALIDEP$(6) = "408000": CODEPX$(6) = "4080000"
   CALI% = 6
   '
   CADEPO% = 0
   For U& = 1 To 64
      If U& <= ULTREG&("SUCURSAL") Then
         CODEPOS$(U&) = TRIM$(Mid$(REGLEIDO$("SUCURSAL", U&), 21, 8))
         CADEPO% = U&
      End If
   Next U&
   '
   'For U& = 1 To ULTREG&("MOVISTO")
   '   X$ = REGLEIDO$("MOVISTO", U&)
   '   CI% = CVI(Mid$(X$, 3, 2))
   '   CAN = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
   '   If CI% > 0 Then
   '      If CI% <= 16384 Then
   '         SALSTO(CI%) = SALSTO(CI%) + CAN
   '      End If
   '   End If
19 'Next U&
   '
   GoSub 95
   '
   HOII% = HOY(HO$)
   NRO& = 0
   CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
   If CMOV$ = "" Then CMOV$ = "AJ"
   '
   ARIMPO$ = "A:\MULTIST"
   PRIBOT% = 0
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
      Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              For KK% = 3 To 6
                 Line Input #NX%, A$
              Next KK%
              CODEPO1$ = TRIM$(Mid$(A$, 2))
              CODEPO$ = TRIM$(REPLACAR$(CODEPO1$, Chr$(34), " "))
              For KK% = 1 To CALI%
                 If CODEPO$ = ALIDEP$(KK%) Then
                    CODEPO$ = CODEPX$(KK%)
                    GoTo 5
                 End If
              Next KK%
              '
5             For KK% = 1 To CADEPO%
                 If CODEPOS$(KK%) = CODEPO$ Then
                    IDEPO% = KK%
                    GoTo 10
                 End If
              Next KK%
              CODEPO$ = CODEPOS$(1)
              IDEPO% = 1
              '
10            For KK% = 7 To 20
                 Line Input #NX%, A$
              Next KK%
              CANTI = Val(Mid$(A$, 2))
              GoSub 90
         End If
      End If
   Wend
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
90 CI% = CODINT%(COD$)
   If CI% > 0 Then
      If Abs(CANTI) > 0.05 Then
         If CI% <= 16384 Then
            SALSTO(CI%) = SALSTO(CI%) + CANTI
         End If
         NRO& = NRO& + 1
         DOSEC$ = "AJ." + Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
         DOPRI$ = "AS." + TRIM$(Str$(NRO&))
         VUNI# = 0
         MONE$ = ""
         NC% = 0
         REFE$ = "ST.IMPORT.AL " + HO$
         LTX$ = "ST.INICIAL"
         DEPO% = IDEPO%
         FF% = HOII%
         '
         Call MOVISTO(FF%, CI%, LTX$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
         '
      End If
   End If
   Return
   '
95 For CI% = 1 To NUMAS%
      U& = CI% + 1
      REBLA$ = String$(36, 0) + MKS$(0.00001)
      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
      Call GRAREG("SALSTOCK", REBLA$, U&)
      REBLA$ = String$(252, 0) + MKS$(0.00001)
      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
      Call GRAREG("SALDEPOS", REBLA$, U&)
   Next CI%
   Return
   '
End Sub

'
Sub IMPORNEU3()
   '
   Dim SALSTO(16384)
   Dim CODEPOS$(64)
   Dim ALIDEP$(128), CODEPX$(128)
   '
   If ULTREG&("UBISTOCK") > 0 Then
      Call MENSERR(24, "Imposible Importar Ubicaciones.\Ya Existe Archivo de Ubicaciones.")
      Call CIERRARCH("*.*")
      GoTo 50
   End If
   '
   ARIMPO$ = "A:\UBIC"
   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   PRIBOT% = 0
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
      Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              For KK% = 3 To 4
                 Line Input #NX%, A$
              Next KK%
              UBICA1$ = TRIM$(Mid$(A$, 2))
              UBICA$ = TRIM$(REPLACAR$(UBICA1$, Chr$(34), " "))
              '
10            For KK% = 5 To 6
                 Line Input #NX%, A$
              Next KK%
              LOECON = Val(Mid$(A$, 2))
              '
20            For KK% = 7 To 8
                 Line Input #NX%, A$
              Next KK%
              PUNPED = Val(Mid$(A$, 2))
              '
              GoSub 90
         End If
      End If
   Wend
   '
50 If ULTREG&("MASTSPEC") > 0 Then
      Call MENSERR(24, "Imposible Importar Especificaciones.\Ya Existe Archivo de Especificaciones.")
      Call CIERRARCH("*.*")
      GoTo 80
   End If
   '
   ARIMPO$ = "A:\ESPECIF"
   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   PRIBOT% = 0
   Screen.MousePointer = 11
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
55    Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              Line Input #NX%, A$
60            Line Input #NX%, TTXX$
              If Left$(TTXX$, 3) = "-1," Then GoTo 55
              GoSub 95
              GoTo 60
         End If
      End If
   Wend
   '
80 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
90 CI% = CODINT%(COD$)
   If CI% > 0 Then
      If TRIM$(UBICA$) <> "" Then
         '
         XX$ = REGBLAN$("UBISTOCK")
         Call REPLA(XX$, MKI$(CI%), 1, 2)
         Call REPLA(XX$, UBICA$, 3, 24)
         Call REGAPP("UBISTOCK", XX$)
         '
      End If
      '
      REGMA& = CI% + 1
      YY$ = REGLEIDO$("MASTER", REGMA&)
      If PUNPED > 0.05 Then
         STMI& = PUNPED
         If STMI& > 30000 Then STMI& = (-1) * PUNPED / 100
         stm1% = STMI&
         Call REPLA(YY$, MKI$(stm1%), 105, 2)
      End If
      If LOECON > 0.05 Then
         LOECO& = LOECON
         If LOECO& > 30000 Then LOECO& = (-1) * LOECON / 100
         lec1% = LOECO&
         Call REPLA(YY$, MKI$(lec1%), 107, 2)
      End If
      Call GRAREG("MASTER", YY$, REGMA&)
      '
   End If
   Return
   '
95 TTYY$ = TRIM$(REPLACAR$(TTXX$, Chr$(34), " "))
   If TRIM$(TTYY$) <> "" Then
      CI% = CODINT%(COD$)
      If CI% > 0 Then
         Call FRATEXTO(TTYY$, 60)
         For KI% = 1 To CARETEX%
            XX$ = REGBLAN$("MASTSPEC")
            Call REPLA(XX$, MKI$(CI%), 1, 2)
            Call REPLA(XX$, RETEX$(KI%), 3, 78)
            Call REGAPP("MASTSPEC", XX$)
         Next KI%
      End If
   End If
   Return
   '
End Sub

Sub EXISTO2()

3   ' EXISTO.QBS    DETERMINACION EXISTENCIAS DE STOCK

    CI1% = PRICODACT%
    CI2% = ULTCODACT%
    TMTX% = 0
    If InStr(EMPREAC$, "DOSIVAC") > 0 Then
       TMTX% = 5
       If HOY(HOS$) = FECHANUM("20/12/01") Then
         'Call CLASIAUTDOSI
       End If
    End If
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(TMTX%)
    
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
       TMT% = 0
       Call REPLA(RACOD$, Chr$(0), 5, 1)
       TIMAT$ = "Material No Clasificado"
    End If
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Call HEADERS("NEGASTO", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    NVII& = NUINV%
    For I& = 1 To NVII&
      Call TRACE(20, I&, NVII&)
      II$ = REGLEIDO$("INVERT", I& + 1)
      CIXI% = CVI(Mid$(II$, 17, 2))
      If Mid$(II$, 16, 1) <> "`" Then
        CDDX$ = CODEXT$(CIXI%)
        If CDDX$ >= cod1$ Then
          If CDDX$ <= cod2$ Then
            TMT1% = TIMATE%(CIXI%)
            If TMT% < 1 Or TMT1% = TMT% Then
              CII& = CIXI%
              SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
              '
              SSTT = 0
              For K% = 1 To 248 Step 4
                SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
              Next K%
              '
              If Abs(SSTT) > 0.05 Then
                TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
                J& = J& + 1
                Call GRAREG("NEGASTO", TTN$, J&)
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        If Option5.Value = True Then
            Call CONECRUN("*EXISTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*EXISTL", "")
        End If
    End If
    '
End Sub
'

Sub NEGASTO2()

3   ' NEGASTO.QBS                     DETERMINACION ITEMS CON STOCK NEGATIVO

    IUI% = -1
    For IU% = 0 To Forms.Count - 1
        XIU$ = UCase$(Forms(IU%).Name)
        If XIU$ = "COBASTO" Or XIU$ = "CONSTOK" Then
            IUI% = IU%
        End If
    Next IU%
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    '
    OPX% = COMALTER%("El Control Puede ser General o Por Depósito\\General\Por Depósito")
    If OPX% = 1 Then
        DEPOSE% = 0
      ElseIf OPX% = 2 Then
        Call SELECHO("TADEPOSI")
        DEPOSE% = Val(VALACT1$("TADEPOSI"))
        If DEPOSE% < 1 Then Exit Sub
      Else
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    NVII& = NUINV%
    SST$ = String$(256, 0)
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CI% > 0 Then
             If CI% <= NUMAS% Then
'                CII& = CI%
'                If DEPOSE% > 0 Then
                SSTT = STODEPOS(CI%, DEPOSE%)
'                SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
'                '
'                SSTT = 0
'                For K% = 1 To 248 Step 4
'                    SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
'                Next K%
'                '
                If SSTT < 0 Then
                    TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
                    J& = J& + 1
                    Call GRAREG("NEGASTO", TTN$, J&)
                End If
             End If
          End If
        End If
     Next I&
     '
     CORRE$ = "Todos los Depósitos"
     If DEPOSE% > 0 Then CORRE$ = ECOARCH$("TADEPOSI", Chr$(DEPOSE%))
     Call HEADERS("NEGASTO", "")
     Call HEADERS("NEGASTO", "Corresponde a: " + CORRE$)
     Call SETULTREG("NEGASTO", J&)
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
         Call COMUNI("No se detectaron items con stock negativo.\Pulse cualquier tecla para volver al menu.")
       Else
         Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Negativos.")
     End If
     '
End Sub
'

Sub CRITSTO2()

3   ' CRITSTO.QBS                     DETERMINACION ITEMS CON STOCK NEGATIVO

    IUI% = 0
    For IU% = 0 To Forms.Count - 1
        XIU$ = UCase$(Forms(IU%).Name)
        If XIU$ = "COBASTO" Or XIU$ = "CONSTOK" Then
            IUI% = IU%
        End If
    Next IU%
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    '
    Screen.MousePointer = 11
    '\\\OT-08-00773-OD-01/09/08///
    Call CARDEPOS 'TRAE EN EL VECTOR DISTO%(AX%) SI EL MATERIAL ESTA DISPONIBLE 0 _
    SI NO DEVUELVE -1
    
    '\\\OT-08-00773-OD-FIN///

    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                CII& = CI%
                SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
                '
                '\\\OT-08-00773-OD-01/09/08///
                '
                'SE MODIFICO LO  SIGUIENTE PARA QUE NO TRAIGA MATERIAL NO DISPONIBLE
                '
                SSTT = 0: AX% = 0
                For K% = 1 To 248 Step 4
                    SALDODEPO = CVS(Mid$(SST$, K%, 4))
                    AX% = AX% + 1
                    If DISTO%(AX%) = 0 Then SSTT = SSTT + SALDODEPO
                Next K%
                '
                '\\\OT-08-00773-OD-FIN///
                '
                STMIN = STOMIN(CI%)
                If SSTT < STMIN Then
                    '
                    COU = COSUNI(CI%)
                    REPOS = STMIN - SSTT
                    If REPOS < LOTESTAN(CI%) Then REPOS = LOTESTAN(CI%)
                    '
                    TTN$ = REGBLAN$("CRITSTO")
                    Call REPLA(TTN$, MKI$(CI%), 1, 2)
                    Call REPLA(TTN$, UMED$, 3, 2)
                    Call REPLA(TTN$, MKS$(SSTT), 5, 4)
                    Call REPLA(TTN$, MKS$(STMIN), 9, 4)
                    Call REPLA(TTN$, MKS$(STMIN - SSTT), 13, 4)
                    Call REPLA(TTN$, MKS$(REPOS), 17, 4)
                    Call REPLA(TTN$, MKS$(COU), 21, 4)
                    Call REPLA(TTN$, MKD$(COU * REPOS), 25, 8)
                    J& = J& + 1
                    Call GRAREG("CRITSTO", TTN$, J&)
                    '
                End If
              End If
           End If
        End If
     Next I&
     '
     Call SETULTREG("CRITSTO", J&)
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
         Call COMUNI("No se detectaron items con stock sub-mínimo.\Pulse cualquier tecla para volver al menu.")
       Else
         Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Sub-Mínimo.")
     End If
     '
End Sub




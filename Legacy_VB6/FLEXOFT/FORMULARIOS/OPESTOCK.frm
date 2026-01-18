VERSION 5.00
Begin VB.Form OPESTOCK 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadística de Stocks"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ultimo Cálculo"
      Height          =   2535
      Left            =   210
      TabIndex        =   10
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   3000
         TabIndex        =   18
         Top             =   1680
         Width           =   3060
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
            Picture         =   "OPESTOCK.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Re-Cálculo de Resumen Mensual para el Mes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Calcular Resumen Mensual de Stocks."
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
            Left            =   840
            TabIndex        =   20
            Top             =   105
            Width           =   2220
         End
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
         TabIndex        =   13
         Text            =   " "
         Top             =   525
         Width           =   2220
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   1
         Left            =   2835
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   0
         Left            =   1365
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -945
         TabIndex        =   17
         Top             =   630
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -945
         TabIndex        =   16
         Top             =   1155
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3675
         TabIndex        =   15
         Top             =   1155
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   1050
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas - Listados"
      Height          =   3165
      Left            =   210
      TabIndex        =   1
      Top             =   2940
      Width           =   5160
      Begin VB.TextBox Text1 
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
         Left            =   4410
         MaxLength       =   3
         TabIndex        =   22
         Text            =   " "
         Top             =   1260
         Width           =   540
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Selección por Tipo de Movimiento .........."
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
         Left            =   525
         TabIndex        =   21
         Top             =   1365
         Width           =   4215
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   4695
         TabIndex        =   8
         Top             =   2730
         Width           =   4750
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   525
         Left            =   4095
         Picture         =   "OPESTOCK.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Emite Reporte según Opciones Elegidas"
         Top             =   1995
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
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
         Left            =   210
         TabIndex        =   4
         Top             =   1890
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2100
            TabIndex        =   5
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valorización de Movimientos de Stocks"
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
         Left            =   525
         TabIndex        =   3
         Top             =   945
         Width           =   4215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Resumen Mensual en Cantidades"
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
         Left            =   525
         TabIndex        =   2
         Top             =   525
         Value           =   -1  'True
         Width           =   3270
      End
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
      Height          =   855
      Left            =   4540
      Picture         =   "OPESTOCK.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Finaliza - Cierra la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "OPESTOCK.frx":075E
      Top             =   2220
      Width           =   2010
   End
End
Attribute VB_Name = "OPESTOCK"
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
   Command2.Enabled = False
   '
   If Option1.Value = True Then
         If Option5.Value = True Then
             Call CONECRUN("*ESTASTO", "Consulta Estadística de Stocks")
           ElseIf Option4.Value = True Then
             Call FIESTOK
         End If
         GoTo 99
       ElseIf Option2.Value = True Then
         If Option5.Value = True Then
             Call MENSERR(24, "Opción no Disponible por Pantalla.\  \Emita Listado Impreso\o Exporte a Archivo de Texto.")
           ElseIf Option4.Value = True Then
             Call FIESVAL
         End If
         GoTo 99
       ElseIf Option3.Value = True Then
         If Option5.Value = True Then
             Call MENSERR(24, "Opción no Disponible por Pantalla.\  \Emita Listado Impreso\o Exporte a Archivo de Texto.")
           ElseIf Option4.Value = True Then
             Call FILMOVE
         End If
         GoTo 99
   End If
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("*CARDESP", "Carga Despachos Importacion por Número")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call ESTASTO
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
   CORRESP$ = CONTROL$("ESTASTO", "ESTASTO")
   PPXX% = InStr(CORRESP$, "- Valores")
   If PPXX% > 0 Then
      CORRESP$ = TRIM$(Left$(CORRESP$, PPXX% - 1))
   End If
   MESAN.TEXT = CORRESP$
   '
   FEX = FEGEAR%("ESTASTO")
   Text4(0).TEXT = FECHATEX$(FEX)
   HHMM% = HOGEAR%("ESTASTO")
   Text4(1).TEXT = Left$(HORATEX$(HHMM%), 5)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
'

Private Sub Option3_Click()
   Text1.Visible = True
End Sub
'

Sub ESTASTO()
    '
    MON1% = 1             ' CALCULA VALORES EN PESOS
    TICO% = 1             ' VALORES INFORMADOS
    TICO1$ = "Valores Informados"
    '
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       Exit Sub
    End If
    MESAN.TEXT = MESASEL$
    Call MESANO$(MESASEL$, DES%, HAS%)
    If DES% < 1 Or HAS% <= DES% Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
700 FEX = DES%: DESX$ = FECHATEX$(FEX)
710 FEX = HAS%: HASX$ = FECHATEX$(FEX)
    '
    Call HEADERS("LIESTOK", "")
    Call HEADERS("LIESTOK", "Corresponde al Mes de: " + CORRESP$)
    Call HEADERS("LIESVAL", "")
    Call HEADERS("LIESVAL", "Corresponde al Mes de: " + CORRESP$)
    Call HEADERS("LIESVAL", TICO1$)
    '
    Call GRACONTROL("ESTASTO", "ESTASTO", TRIM$(MESASEL$) + " - " + TICO1$)
    Call GRARGEN("ESTASTO", TRIM$(MESASEL$) + " - " + TICO1$)
    '
    Call BLOQARCH("ESTAMOV")
    Call BLOQARCH("ESTASTO")
    Call BLOQARCH("ESTAVAL")
    '
    '   Call MENSAJE(11, "INICIALIZANDO ARCHIVO DE ESTADISTICA")
    '
    FIN& = NUMAS%
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CI% = I&
       Call GRAREG("ESTAMOV", MKI$(CI%) + String$(40, 0), I&)
    Next I&
    Call SETULTREG("ESTAMOV", FIN&)
    '
    '   Call MENSAJE(12, "LEYENDO ARCHIVO DE MOVIMIENTOS DE STOCK")
    '
    LU$ = LUDAT$
    CAMO& = ULTREG&("MOVISTO")
    Dim GG As String * 128
    AMOVI% = FILEOPEN%(LU$ + "MOVISTO.DAT", 0, 128)
    For I& = 1 To CAMO&
       Call TRACE(20, I&, CAMO&)
       'GG$ = REGLEIDO$("MOVISTO", I&)
       Get #AMOVI%, I&, GG$
       FECHA% = CVI(Left$(GG$, 2))
       If FECHA% <= HAS% Then
           CI% = CVI(Mid$(GG$, 3, 2))
           If CI% < 1 Or CI% > NUMAS% Then GoTo 799
           '
           IL& = CI%
           XX$ = REGLEIDO$("ESTAMOV", IL&)
           '
           CAN# = CVD(Mid$(GG$, 88, 8)) - CVD(Mid$(GG$, 96, 8))
           COMO$ = UCase$(Mid$(GG$, 21, 2))
           AJU% = 0: If COMO$ = "AS" Or COMO$ = "AJ" Then AJU% = 1
                     If COMO$ = "TR" Or COMO$ = "VR" Then AJU% = 1
           If FECHA% < DES% Then
                X5$ = MKD$(CVD(Mid$(XX$, 3, 8)) + CAN#)
                Call REPLA(XX$, X5$, 3, 8)
             Else
                If AJU% > 0 Then
                    X8$ = MKD$(CVD(Mid$(XX$, 27, 8)) + CAN#)
                    Call REPLA(XX$, X8$, 27, 8)
                 Else
                    If CAN# > 0 Then
                        X6$ = MKD$(CVD(Mid$(XX$, 11, 8)) + CAN#)
                        Call REPLA(XX$, X6$, 11, 8)
                      Else
                        X7$ = MKD$(CVD(Mid$(XX$, 19, 8)) - CAN#)
                        Call REPLA(XX$, X7$, 19, 8)
                    End If
                End If
           End If
           X9$ = MKD$(CVD(Mid$(XX$, 35, 8)) + CAN#)
           Call REPLA(XX$, X9$, 35, 8)
           '
           Call GRAREG("ESTAMOV", XX$, IL&)
           '
       End If
799 Next I&
    Close #AMOVI%: AMOVI% = 0
    '
    '  Call MENSAJE(14, "VALORIZANDO MOVIMIENTOS MENSUALES")
    '
    FIN& = NUINV%
    JJ& = 0
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       IX$ = REGLEIDO$("INVERT", I& + 1)
       If Mid$(IX$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(IX$, 17, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             IL& = CI%
             XX$ = REGLEIDO$("ESTAMOV", IL&)
             If Mid$(XX$, 3, 40) <> String$(40, 0) Then
                COU = COSUNI(CI%)
                JJ& = JJ& + 1
                Call GRAREG("ESTASTO", XX$, JJ&)
                '
                For K% = 3 To 42 Step 8
                    VALO# = COU * CVD(Mid$(XX$, K%, 8))
                    Call REPLA(XX$, MKD$(VALO#), K%, 8)
                Next K%
                Call GRAREG("ESTAVAL", XX$, JJ&)
             End If
           End If
         End If
       End If
    Next I&
    '
    Call SETULTREG("ESTASTO", JJ&)
    Call SETULTREG("ESTAVAL", JJ&)
    Call CIERRARCH("ESTAMOV")
    Call CIERRARCH("ESTAVAL")
    Call CIERRARCH("ESTASTO")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
End Sub

Sub FIESTOK()
   '
10 XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   '
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
      GoTo 99
   End If
   CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   If cod1$ > cod2$ Then GoTo 10
   '
   Call BLOQARCH("LIESTOK")
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("ESTASTO")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      X$ = REGLEIDO$("ESTASTO", IL&)
      CI% = CVI(Left$(X$, 2))
      CODD$ = CODEXT$(CI%)
      If CODD$ >= cod1$ Then
         If CODD$ <= cod2$ Then
            JJ& = JJ& + 1
            Call GRAREG("LIESTOK", X$, JJ&)
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LIESTOK", JJ&)
   Call CIERRARCH("LIESTOK")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIESTOK", "Listado Estadístico Movimientos de Stock")
   '
99 Screen.MousePointer = 1
   '
End Sub

Sub FIESVAL()
   '
10 XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   '
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
      GoTo 99
   End If
   CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   If cod1$ > cod2$ Then GoTo 10
   '
   Call BLOQARCH("LIESVAL")
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("ESTAVAL")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      X$ = REGLEIDO$("ESTAVAL", IL&)
      CI% = CVI(Left$(X$, 2))
      CODD$ = CODEXT$(CI%)
      If CODD$ >= cod1$ Then
         If CODD$ <= cod2$ Then
            JJ& = JJ& + 1
            Call GRAREG("LIESVAL", X$, JJ&)
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LIESVAL", JJ&)
   Call CIERRARCH("LIESVAL")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIESVAL", "Listado Valorizado Estadística de Stock")
   '
99 Screen.MousePointer = 1
   '
End Sub
'

Sub FILMOVE()
   '
   TIMOFI$ = UCase$(TRIM$(Text1.TEXT))
   If Len(TIMOFI$) <> 2 Then
      Call MENSERR(24, "Ingrese Tipo de Movimiento a Listar")
      Text1.TEXT = ""
      Text1.SetFocus
      Exit Sub
   End If
   '
   MESASEL$ = MESAN.TEXT
   Call MESANO$(MESASEL$, DES%, HAS%)
   If DES% < 1 Or HAS% <= DES% Then
      Exit Sub
   End If
   '
   FEX = DES%: DESX$ = FECHATEX$(FEX)
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
    '
10 XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   '
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
      GoTo 99
   End If
   CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   If cod1$ > cod2$ Then GoTo 10
   '
   Call BLOQARCH("MOVISTOL")
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("MOVISTO")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      X$ = REGLEIDO$("MOVISTO", IL&)
      FECHA% = CVI(Left$(X$, 2))
      If FECHA% >= DES% Then
         If FECHA% <= HAS% Then
            TIMO$ = UCase$(Mid$(X$, 21, 2))
            If TIMO$ = TIMOFI$ Then
               CI% = CVI(Mid$(X$, 3, 2))
               CODD$ = CODEXT$(CI%)
               If CODD$ >= cod1$ Then
                  If CODD$ <= cod2$ Then
                     JJ& = JJ& + 1
                     Call GRAREG("MOVISTOL", X$, JJ&)
                  End If
               End If
            End If
         End If
      End If
   Next IL&
   '
   Call SETULTREG("MOVISTOL", JJ&)
   Call CIERRARCH("MOVISTOL")
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SORMOVE", "Listado de Movimientos Seleccionados")
   '
99 Screen.MousePointer = 1
   '
End Sub

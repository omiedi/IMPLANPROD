VERSION 5.00
Begin VB.Form ANAPEPEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Pedidos Pendientes de Entrega"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Condensación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   2940
      TabIndex        =   24
      Top             =   2520
      Width           =   3900
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total General por Artículo"
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
         Index           =   2
         Left            =   525
         TabIndex        =   27
         Top             =   1050
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total por Artículo y Cliente"
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
         Index           =   1
         Left            =   525
         TabIndex        =   26
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Detalle por Pedido"
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
         Index           =   0
         Left            =   525
         TabIndex        =   25
         Top             =   420
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Valorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   210
      TabIndex        =   21
      Top             =   2940
      Width           =   2535
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Cantidades y Precios"
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
         Index           =   1
         Left            =   315
         TabIndex        =   23
         Top             =   635
         Width           =   2115
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Solo Cantidades"
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
         Index           =   0
         Left            =   315
         TabIndex        =   22
         Top             =   370
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFE0&
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
      Height          =   1275
      Left            =   210
      TabIndex        =   17
      Top             =   1575
      Width           =   2535
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Total Pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   900
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Meses Posteriores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   19
         Top             =   635
         Width           =   2010
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Mes en Curso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   370
         Width           =   2010
      End
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
      Height          =   840
      Left            =   7035
      Picture         =   "ANAPEPEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Genera y Emite Reporte Solicitado"
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7035
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   13
      Top             =   2310
      Width           =   865
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   2940
      TabIndex        =   6
      Top             =   210
      Width           =   3900
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2625
         TabIndex        =   10
         Top             =   465
         Width           =   855
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
         Height          =   285
         Left            =   3465
         TabIndex        =   9
         Top             =   465
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2625
         TabIndex        =   8
         Text            =   " "
         Top             =   1365
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Height          =   285
         Left            =   3465
         TabIndex        =   7
         Top             =   1365
         Width           =   175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   1725
         Width           =   3450
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   825
         Width           =   3450
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   510
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   210
         TabIndex        =   11
         Top             =   1410
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Selección"
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
      TabIndex        =   3
      Top             =   210
      Width           =   2535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Todos los Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   28
         Top             =   900
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Compras a Terceros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   5
         Top             =   635
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Fabricación Propia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   370
         Width           =   2115
      End
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
      Height          =   855
      Left            =   7035
      Picture         =   "ANAPEPEN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
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
      Height          =   855
      Left            =   7035
      Picture         =   "ANAPEPEN.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6450
      Picture         =   "ANAPEPEN.frx":075E
      Top             =   3480
      Width           =   2010
   End
End
Attribute VB_Name = "ANAPEPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GENLISPE
   Command2.Enabled = True
End Sub
'

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
       Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " (" + APLINVO$ + ")  -  " + EPAC$
    End If
    '
    Show
    Refresh
    If ULTREG&("DEUDOR1") > 0 Then
      Screen.MousePointer = 11
      NCMIN% = 32767: NCMAX% = 0
      FIN& = ULTREG&("DEUDOR1")
      For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
          If NCX% > NCMAX% Then NCMAX% = NCX%
          If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
      Next U&
      '
      Screen.MousePointer = 1
      Text10 = TRIM$(Str$(NCMIN%))
      Text8 = TRIM$(Str$(NCMAX%))
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = RASOCLI$(NCMIN%)
   End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Text8_Change()
   Label8 = ""
   NCMAX% = Val(Text8.TEXT)
   If Val(NCMAX%) > 0 Then
     Label8 = RASOCLI$(NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Sub GENLISPE()
   '
   If InStr(Caption, "(POR CLIENTE)") > 0 Then
     If Option4(2).Value = True Then
       Call MENSERR(24, "Combinación no Válida")
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0: LL& = 0
   FIN& = ULTREG&("PEDDETA")
   '
   MODO% = 3
   If Option2(0).Value = True Then MODO% = 1
   If Option2(1).Value = True Then MODO% = 2
   FLI1% = -1: FLI2% = FINMESACT%
   If MODO% = 2 Then
       FLI1% = FINMESACT% + 1: FLI2% = 32767
     ElseIf MODO% = 3 Or MODO% = 4 Then
       FLI1% = -1: FLI2% = 32767
   End If
   '
   OPX% = 2
   If Option3(0) = True Then OPX% = 1
   '
   Call HEADERS("LIPEDPE", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   Call HEADERS("LIPEDPE", VARIANTE$)
   '
   NCL1% = Val(Text10)
   NCL2% = Val(Text8)
   If NCL1% > NCL2% Then Exit Sub
   '
   GRCF% = 0
   If Option1(2).Value = True Then GRCF% = 1
   If Option1(1).Value = True Then GRCF% = 2
   '
   VECPEDI$ = ""
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("PEDDETA", U&)
     STATUPE% = Asc(Mid$(X$, 27, 1))
     If STATUPE% <> 9 Then ' no esta anulado
       If STATUPE% <> 3 Then ' no esta entregado
         CI7% = CVI(Mid$(X$, 33, 2))
         If CI7 > 0 And CI7% <= NUMAS% Then
           If GRCF% = 0 Or GRUCOFAB%(CI7%) = GRCF% Then
             NC% = CVI(Mid$(X$, 7, 2))
             If NC% >= NCL1% Then
               If NC% <= NCL2% Then
                 CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                 If CAPEN > 0.05 Then
                   X1$ = X$
                   FENTRESOL% = CVI(Mid$(X$, 25, 2))
                   If FENTRESOL% >= FLI1% Then
                     If FENTRESOL% <= FLI2% Then
                       '
                       Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                       Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                       Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                       Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                       '
                       CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                       If CANOASI < 0 Then CANOASI = 0
                       '
                       PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                       If OPX% <> 2 Then PREUNI = 0
                       ITOTA# = CDbl(CAPEN * PREUNI)
                       INOAS# = CDbl(CANOASI * PREUNI)
                       '
                       Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                       Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                       Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                       Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                       Call REPLA(Y$, Space$(34), 127, 34)
                       '
                       JJ& = JJ& + 1
                       Call GRAREG("LIPEDPE", Y$, JJ&)
                       VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                       '
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
30 If MODO% = 4 Then
     KK& = JJ&: JJ& = 0
     For U& = 1 To KK&
       Y$ = REGLEIDO$("LIPEDPE", U&)
       If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
         JJ& = JJ& + 1
         Call GRAREG("LIPEDPE", Y$, JJ&)
       End If
     Next U&
   End If
   '
   Call SETULTREG("LIPEDPE", JJ&)
   Call CIERRARCH("LIPEDPE")
   '
   Call HEADERS("LIPEDPE", "")
   VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
   '
   If Option4(0).Value = True Then
     If InStr(Caption, "(POR CLIENTE)") > 0 Then
         Call FILESORT("LIPEDPE", "", 1, 3)
         SALI$ = "*LIPENCL"
       Else
         Call FILESORT("LIPEDPE", "", 3, 3)
         SALI$ = "*LIPENMA"
     End If
     Call CIERRARCH("LIPEDPE")
     Call FINAL(SALI$)
     GoTo 99
   End If
   '
   If Option4(1).Value = True Or Option4(2).Value = True Then ' concentrar por articulo y cliente
     FIN& = ULTREG&("LIPEDPE")
     PUCIX$ = ""
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI% = CVI(Mid$(XX$, 1, 2))
       CII% = CVI(Mid$(XX$, 3, 2))
       '
       If CII% <= 0 Then
            PUCIX$ = PUCIX$ + MKI$(0)
         Else
            CAPENPE = CVS(Mid$(XX$, 81, 4)) - CVS(Mid$(XX$, 97, 4))
            If CAPENPE < 0 Then CAPENPE = 0
            For KKU& = 1 To Len(PUCIX$) Step 2
              If CVI(Mid$(PUCIX$, KKU&, 2)) = CII% Then
                YY$ = REGLEIDO$("LIPEDPE", (KKU& + 1) / 2)
                If CVI(Mid$(YY$, 1, 2)) = NCI% Or Option4(2).Value = True Then
                  TOCANTI = CVS(Mid$(YY$, 81, 4)) + CVS(Mid$(XX$, 81, 4))
                  TOENTRE = CVS(Mid$(YY$, 97, 4)) + CVS(Mid$(XX$, 97, 4))
                  TOPENPE = CVS(Mid$(YY$, 101, 4)) + CAPENPE
                  ITOTAL# = CVD(Mid$(YY$, 89, 8)) + CVD(Mid$(XX$, 89, 8))
                  Call REPLA(YY$, MKS$(TOCANTI), 81, 4)
                  Call REPLA(YY$, MKD$(ITOTAL#), 89, 8)
                  Call REPLA(YY$, MKS$(TOENTRE), 97, 4)
                  Call REPLA(YY$, MKS$(TOPENPE), 101, 4)
                  Call GRAREG("LIPEDPE", YY$, (KKU& + 1) / 2)
                  PUCIX$ = PUCIX$ + MKI$(0)
                  Call GRAREG("LIPEDPE", String$(128, 0), U&)
                  GoTo 39
                End If
              End If
            Next KKU&
            '
            PUCIX$ = PUCIX$ + MKI$(CII%)
            Call GRAREG("LIPEDPE", XX$, U&)
       End If
39   Next U&
     '
     JJ& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("LIPEDPE", U&)
       NCI% = CVI(Mid$(XX$, 1, 2))
       CII% = CVI(Mid$(XX$, 3, 2))
       '
       If NCI% > 0 Then
         If CII% > 0 Then
           JJ& = JJ& + 1
           Call GRAREG("LIPEDPE", XX$, JJ&)
         End If
       End If
     Next U&
     '
     Call SETULTREG("LIPEDPE", JJ&)
     '
     If Option4(1).Value = True Then
       If InStr(Caption, "(POR CLIENTE)") > 0 Then
           Call FILESORT("LIPEDPE", "", 1, 3)
           SALI$ = "*LIPENCL"
         Else
           Call FILESORT("LIPEDPE", "", 3, 3)
           SALI$ = "*LIPENMA"
       End If
       Call FINAL(SALI$)
       GoTo 99
     End If
     '
     If Option4(2).Value = True Then
       FIN& = ULTREG&("LIPEDPE")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("LIPEDPE", U&)
         Call REPLA(XX$, String$(56, 0), 105, 56)
         CIXI% = CVI(Mid$(XX$, 3, 2))
         NECEPE = CVS(Mid$(XX$, 101, 4))
         If NECEPE >= 0.005 Then
           DISPONI = STOCKACT(CIXI%)
           If DISPONI < 0 Then DISPONI = 0
           Call REPLA(XX$, MKS$(DISPONI), 105, 4)
           NECEPE = NECEPE - DISPONI
         End If
         DISPONI = STOCOPEN(CIXI%) + STORFAB(CIXI%)
         Call REPLA(XX$, MKS$(DISPONI), 109, 4)
         NECEPE = NECEPE - DISPONI
         If NECEPE >= 0.05 Then
             Call REPLA(XX$, MKS$(NECEPE), 117, 4)
             DISPONI = MAXIFAB(CIXI%)
             Call REPLA(XX$, MKS$(DISPONI), 121, 4)
           Else
             Call REPLA(XX$, MKS$(Abs(NECEPE)), 113, 4)
         End If
         '
         Call GRAREG("LIPEPEN", XX$, U&)
         '
       Next U&
       '
       Call SETULTREG("LIPEPEN", FIN&)
       Call FILESORT("LIPEPEN", "", 3, 3)
       '
       Call HEADERS("LIPEPEN", "")
       If Option1(0).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: Total Pendiente")
       If Option1(1).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Fabricar")
       If Option1(2).Value = True Then Call HEADERS("LIPEPEN", "Corresponde a: A Comprar")
       '
       Call FINAL("*LIPEPEN")
       GoTo 99
     End If
   End If
   '
99 Screen.MousePointer = 1

End Sub

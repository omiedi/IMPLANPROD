VERSION 5.00
Begin VB.Form DISPMRP04 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8115
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
   ScaleHeight     =   4755
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1890
      TabIndex        =   33
      Top             =   3570
      Width           =   1905
      Begin VB.OptionButton Option6 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Disponibilidad"
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
         Left            =   210
         TabIndex        =   35
         Top             =   380
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Cobertura"
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
         Left            =   210
         TabIndex        =   34
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   7575
      Left            =   7245
      ScaleHeight     =   7515
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command12 
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
         Height          =   880
         Left            =   105
         Picture         =   "DISPMRP04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   3650
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   30
         Top             =   2730
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command11 
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
         Height          =   690
         Left            =   105
         Picture         =   "DISPMRP04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1005
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "DISPMRP04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
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
         Picture         =   "DISPMRP04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1785
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DISPMRP04.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3000
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEC0&
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
      Height          =   1485
      Left            =   210
      TabIndex        =   13
      Top             =   1940
      Width           =   3585
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Todos los Tipos"
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
         Left            =   210
         TabIndex        =   18
         Top             =   525
         Value           =   1  'Checked
         Width           =   1695
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
         Height          =   285
         Left            =   3150
         TabIndex        =   15
         Top             =   525
         Width           =   175
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
         Height          =   285
         Left            =   2730
         TabIndex        =   14
         Top             =   525
         Width           =   435
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   17
         Top             =   945
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         TabIndex        =   16
         Top             =   525
         Width           =   750
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Controles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3885
      TabIndex        =   9
      Top             =   3565
      Width           =   3165
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Verificar Movimientos y Saldos de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   420
         TabIndex        =   25
         Top             =   260
         Width           =   2220
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Rango de Códigos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   6840
      Begin VB.CommandButton Command1 
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
         Height          =   300
         Left            =   3150
         TabIndex        =   5
         Top             =   460
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   6
         Top             =   460
         Width           =   2010
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4515
         TabIndex        =   4
         Text            =   " "
         Top             =   460
         Width           =   2010
      End
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
         Height          =   300
         Left            =   6510
         TabIndex        =   3
         Top             =   460
         Width           =   175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3465
         TabIndex        =   12
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   11
         Top             =   900
         Width           =   3255
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
         Left            =   -1155
         TabIndex        =   8
         Top             =   525
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
         Left            =   2100
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEC0&
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
      Height          =   1020
      Left            =   210
      TabIndex        =   1
      Top             =   3565
      Width           =   1590
      Begin VB.OptionButton Option12 
         BackColor       =   &H00EEEEC0&
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
         Height          =   240
         Left            =   210
         TabIndex        =   22
         Top             =   630
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00EEEEC0&
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
         Height          =   240
         Left            =   210
         TabIndex        =   21
         Top             =   380
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Filtrado"
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
      Left            =   3885
      TabIndex        =   0
      Top             =   1940
      Width           =   3165
      Begin VB.OptionButton Option4 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Productos Terminados"
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
         Left            =   420
         TabIndex        =   24
         Top             =   1125
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Todos los Artículos"
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
         Left            =   420
         TabIndex        =   23
         Top             =   855
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Artículos con Movimiento"
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
         Left            =   420
         TabIndex        =   20
         Top             =   585
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Solo Artículos Faltantes"
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
         Left            =   420
         TabIndex        =   19
         Top             =   315
         Value           =   -1  'True
         Width           =   2535
      End
   End
End
Attribute VB_Name = "DISPMRP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IAT&
'

Private Sub Check1_Click()
   Text3.SetFocus
End Sub

Private Sub Check1_LostFocus()
   If Check1.Value = 1 Then
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Command1_Click()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call STODISP
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   If (USNBR% Mod 100) = 1 Then
      On Error Resume Next
      OPDISTK04.Show 1
      On Error GoTo 0
   End If
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Sub ANACOBER()
  '\\\OT-06-0102-WAR-22/03/06///
   Screen.MousePointer = 11
   'ESCRIBO EL ARCHIVO COBEVEST
   JJ& = 0
   Y$ = REGBLAN$("COBEVEST")
   FIN& = ULTREG("STODISP")
   For J& = 1 To FIN&
     Call TRACE(20, J&, FIN&)
     X$ = REGLEIDO$("STODISP", J&)
     CIXJ% = CVI(Mid$(X$, 1, 2)) 'CODIGO
     STACTU1 = CVS(Mid(X$, 5, 4)) 'STOCK ACTUAL
     QC1 = CVS(Mid$(X$, 9, 4)) 'CANTIDAD COMPRAS
     QOF1 = CVS(Mid$(X$, 13, 4)) 'CANTIDAD O/F
     QAI1 = QC1 + QOF1 'CANTIDAD A INGRESAR
     STMIN1 = CVS(Mid$(X$, 17, 4)) 'STOCK MINIMO
     PPENT1 = CVS(Mid$(X$, 21, 4))
     TOTA1 = STACTU1 - PPENT1 + QAI1 'ST. ACT. + CANTIDAD A INGRESAR
     COB1 = 0
     COB2 = 0
     '
     If STMIN1 > 0 Then
        COB1 = (STACTU1 - PPENT1) / STMIN1
        COB2 = TOTA1 / STMIN1
     End If
     '
     Call REPLA(Y$, MKI$(CIXJ%), 1, 2)
     Call REPLA(Y$, MKS$(STMIN1), 3, 4)
     Call REPLA(Y$, MKS$(STACTU1), 7, 4)
     Call REPLA(Y$, MKS$(PPENT1), 11, 4)
     Call REPLA(Y$, MKS$(STACTU1 - PPENT1), 15, 4)
     Call REPLA(Y$, MKS$(COB1), 19, 4)
     Call REPLA(Y$, MKS$(QAI1), 23, 4)
     Call REPLA(Y$, MKS$(TOTA1), 27, 4)
     Call REPLA(Y$, MKS$(COB2), 31, 4)
     JJ& = JJ& + 1
     Call GRAREG("COBEVEST", Y$, JJ&)
     '
   Next J&
   '
   Call SETULTREG("COBEVEST", JJ&)
   Call FILESORT("COBEVEST", "COBEVEST", 7, 7, "ASC")
   Call CIERRARCH("COBEVEST")
   '
   Call FINAL("*COBEVEST")
   '
99 Screen.MousePointer = 1
  ''\\\OT-06-0102-WAR-FIN///
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command6_Click()
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
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16)
    Text1.TEXT = XX$
    CI1% = Abs(CODINT%(XX$))
    Label1.Caption = DESCRIT$(CI1%)
    '
    AAA& = ULTREG&("ECOMAST")
10  XX$ = Left$(REGLEIDO$("ECOMAST", AAA&), 16)
    If TRIM$(XX$) = "" Then
      If AAA& > 1 Then
        AAA& = AAA& - 1
        GoTo 10
      End If
    End If
    Text2.TEXT = XX$
    CI2% = Abs(CODINT%(XX$))
    Label2.Caption = DESCRIT$(CI2%)
    '
    'Call VERNOVTRA
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Text2_Change()
    CI2% = CODINT%(Text2.TEXT)
    If CI2% > 0 Then
         Label2.Caption = DESCRIT$(CI2%)
       Else
         Label2.Caption = ""
    End If
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Text3.TEXT) = "" Then
      Check1.Value = 1
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.TEXT <> "" Then
      If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
        End If
      End If
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
      Check1.Value = 1
    End If
End Sub


Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
        End If
    End If
End Sub

Private Sub Text8_Change()
End Sub

Sub STODISP()
    '
    HOII% = HOY(HOS$)
'    IAT& = 0            ' INDICE ARCHIVO TRANSFERENCIA
'    '
'    VECOPEN$ = PUNOCPEN$
    '
    Call CARDEPOS
    '
    TMT% = XVALO(Text3.TEXT)
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
    End If
    '
    CI1% = CODINT%(Text1.TEXT)
    COD1$ = CODEXT$(CI1%)
    If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CODINT%(Text2.TEXT)
    COD2$ = CODEXT$(CI2%)
    If TRIM$(COD2$) = "" Then Exit Sub
    '
    NIVEFIL% = 1
    If Option2.Value = True Then NIVEFIL% = 2
    If Option1.Value = True Then NIVEFIL% = 3
    If Option4.Value = True Then NIVEFIL% = 4
    '
    Screen.MousePointer = 11
    Dim EXIS(16384, 6)
    Dim PUNOF$(16384)
    Dim PROTER%(16384)      ' MARCA PRODUCTO TERMINADO
    '
    UREMAS& = ULTREG&("MASTER")
    User% = USNBR% Mod 100
    If User% < 1 Then User% = 1
    '
    FIN& = NUMAS%
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CIIX% = I&
       If Check2.Value = 1 Then
          Call VESARRAS(CIIX%)
       End If
       EXIS(I&, 1) = STOCKACT(CIIX%)
       EXIS(I&, 4) = STOMIN(CIIX%)
    Next I&
    '
    Call ABRECOMPRAS
    SQLX$ = " SELECT * FROM OCOMDETA "
      'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "WHERE Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
    SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
    SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
    '
    Set OcoTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OcoTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NUORSE& = !N_OrServ
          If NUORSE& < 1 Then
            COEFCAN = !Coef_Unids
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = !Cant_Ocom * COEFCAN
            CAREC = !Cant_Rec * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              CIIX% = !Cod_Inte
              EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    '\\\OT-06-0149-WAR-25/04/06///
    Call ABRECOMPRAS
    SQLX$ = " SELECT * FROM AUTOENT "
      'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "WHERE Stat_Aten < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
    SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
    SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
    '
    Set AtenTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With AtenTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NUORSE& = 0 ' !N_OrServ
          If NUORSE& < 1 Then
            COEFCAN = 1 ' !Coef_Unids
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = !Cant_Aten * COEFCAN
            CAREC = !Acuen_Aten * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              CIIX% = !Cod_Inte
              EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '\\\OT-06-0149-WAR-FIN///
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    '
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COBRECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While (.EOF = False)
          CACOM = !CANTNOM
          CAREC = !CanToApro
          CAPEN = CACOM - CAREC
          If CAPEN > 0 Then
            CIIX% = !Cod_Inte
            EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
    '
    Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OfaTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          CACOM = !CanProy
          CAREC = !CanApro
          CAPEN = CACOM - CAREC
          If CAPEN > 0 Then
            CIIX% = !Cod_Inte
            EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    Call ABRESERVA
    SQLX$ = " SELECT * FROM RESERVA "
    '  SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
    '  SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
      '
    Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ReserTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          TORE# = !CantRes
          CACO# = !CantCons
          '
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
             TORE# = TORE# / 1000
             CACO# = CACO# / 1000
          End If
          '
          SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
          SARE = TORE# - CACO#
          If SARE * SIGNO% < 0 Then
            SARE = 0
          End If
          CAPEN = SARE
          If Abs(CAPEN) > 0 Then
            CIIX% = !Cod_Inte
            EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    Call ABREPEDCLI
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + "WHERE Status < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
    SQLX$ = SQLX$ + "AND VenRep = 'V' "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NROX& = !NroPed
          NCI% = !NROCLIE
          CAPEN = !CanPed - !CantEnt
          '
          If CAPEN > 0 Then
            CIIX% = !CodiInt
            EXIS(CIIX%, 5) = EXIS(CIIX%, 5) + CAPEN
          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    If NIVEFIL% = 4 Then GoTo 100 ' situacion productos terminados
    '
    '
50  J& = 0
    NVII& = NUMAS%
    For I& = 1 To NVII&
      Call TRACE(20, I&, NVII&)
      CIIX% = I&
      If STATUIT%(CIIX%) > 0 Then
        CODE$ = CODEXT$(CIIX%)
        If CODE$ >= COD1$ Then
          If CODE$ <= COD2$ Then
            If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
              X$ = REGBLAN$("STODISP")
              Call REPLA(X$, MKI$(CIIX%), 1, 2)
              Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
              SUMA = 0: HAYMO% = 0
              For K% = 1 To 6
                SIG% = 1: If K% >= 4 Then SIG% = (-1)
                SUMA = SUMA + SIG% * EXIS(CIIX%, K%)
                If Abs(EXIS(CIIX%, K%)) >= 0.05 Then HAYMO% = 1
                Call REPLA(X$, MKS$(EXIS(CIIX%, K%)), 4 * K% + 1, 4)
              Next K%
              If SUMA > 0.05 Then
                  Call REPLA(X$, MKS$(SUMA), 29, 4)
                ElseIf SUMA < (-0.05) Then
                  Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                  'Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
              End If
              '
              If (NIVEFIL% = 1) Or (NIVEFIL% = 2 And HAYMO% = 1) Or (NIVEFIL% = 3 And SUMA <= (-0.05)) Then
                Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                J& = J& + 1
                Call GRAREG("STODISP", X$, J&)
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
60  Call SETULTREG("STODISP", J&)
    Call CIERRARCH("STODISP")
    Call FILESORT("STODISP", "", 1, 1, "ASC")
    '
    Call CIERRARCH("*.*")
    '
    If Option5.Value = True Then ' COBERTURA
      Call ANACOBER
      Exit Sub
    End If
    '
    If Option11.Value = True Then
         Call CONECRUN("*STODISC", "")
       ElseIf Option12.Value = True Then
         Call CONECRUN("*STODISL", "")
    End If
    Exit Sub
    '
100 Call ABRESTRUC
    With Formulas
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          CIIX% = !CODICONJ
          If CIIX% > 0 Then
            PROTER%(CIIX%) = 1
          End If
        .MoveNext
        Loop
      End If
      '
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          CIIX% = !CODIELEM
          If CIIX% > 0 Then
            PROTER%(CIIX%) = 0
          End If
        .MoveNext
        Loop
      End If
      '
    End With
    '
    J& = 0
    NVII& = NUMAS%
    For I& = 1 To NVII&
      Call TRACE(20, I&, NVII&)
      CIIX% = I&
      If STATUIT%(CIIX%) > 0 Then
        If PROTER%(CIIX%) = 1 Then
          CODE$ = CODEXT$(CIIX%)
          If CODE$ >= COD1$ Then
            If CODE$ <= COD2$ Then
              If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
                X$ = REGBLAN$("STODISP")
                Call REPLA(X$, MKI$(CIIX%), 1, 2)
                Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
                SUMA = 0: HAYMO% = 0
                For K% = 1 To 6
                  SIG% = 1: If K% >= 4 Then SIG% = (-1)
                  SUMA = SUMA + SIG% * EXIS(CIIX%, K%)
                  If Abs(EXIS(CIIX%, K%)) >= 0.05 Then HAYMO% = 1
                  Call REPLA(X$, MKS$(EXIS(CIIX%, K%)), 4 * K% + 1, 4)
                Next K%
                If SUMA > 0.05 Then
                    Call REPLA(X$, MKS$(SUMA), 29, 4)
                  ElseIf SUMA < (-0.05) Then
                    Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                    'Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
                End If
                Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                J& = J& + 1
                Call GRAREG("STODISP", X$, J&)
                '
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
    GoTo 60
    '
End Sub
'

VERSION 5.00
Begin VB.Form ANACOB04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Cobertura MRP-II"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8250
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
   ScaleHeight     =   5835
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fecha de Corte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3885
      TabIndex        =   33
      Top             =   4890
      Width           =   3270
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
         Left            =   2625
         TabIndex        =   34
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label14 
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
         Left            =   -315
         TabIndex        =   36
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Height          =   300
         Left            =   1575
         TabIndex        =   35
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   3165
      Left            =   3885
      TabIndex        =   11
      Top             =   1680
      Width           =   3270
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   105
         ScaleHeight     =   375
         ScaleWidth      =   2910
         TabIndex        =   31
         Top             =   420
         Width           =   2970
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   " Nro. Tipo de Material"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   105
            TabIndex        =   32
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.ListBox LICOPEN 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   30
         ToolTipText     =   "Tilde los Vendedores a Considerar"
         Top             =   840
         Width           =   2970
      End
      Begin VB.Line Line2 
         X1              =   3075
         X2              =   3075
         Y1              =   450
         Y2              =   2940
      End
      Begin VB.Line Line3 
         X1              =   105
         X2              =   3080
         Y1              =   2940
         Y2              =   2940
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Último Cálculo"
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
      TabIndex        =   18
      Top             =   210
      Width           =   6945
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro:"
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
         Left            =   -420
         TabIndex        =   27
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label10 
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
         Height          =   510
         Left            =   975
         TabIndex        =   26
         Top             =   735
         Width           =   5790
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
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
         Left            =   2940
         TabIndex        =   25
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4335
         TabIndex        =   24
         Top             =   360
         Width           =   2430
      End
      Begin VB.Label Label123 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
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
         Left            =   -420
         TabIndex        =   23
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   2070
         TabIndex        =   22
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label121 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hs"
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
         Left            =   2940
         TabIndex        =   21
         Top             =   435
         Width           =   330
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   2205
         TabIndex        =   20
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   975
         TabIndex        =   19
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   7575
      Left            =   7350
      ScaleHeight     =   7515
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture2 
         Height          =   6000
         Left            =   0
         ScaleHeight     =   5940
         ScaleWidth      =   1005
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   5895
            Left            =   0
            Picture         =   "DISPMRP04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
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
         Height          =   650
         Left            =   105
         Picture         =   "DISPMRP04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   3255
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   4935
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   825
         Left            =   105
         Picture         =   "DISPMRP04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   3990
         Width           =   650
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
         Picture         =   "DISPMRP04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "DISPMRP04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DISPMRP04.frx":2C40
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5250
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
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
      Height          =   2745
      Left            =   210
      TabIndex        =   1
      Top             =   1680
      Width           =   3585
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   1620
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
         Left            =   3150
         TabIndex        =   2
         Top             =   1620
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
         Left            =   210
         TabIndex        =   10
         Top             =   2055
         Width           =   3150
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
         Left            =   210
         TabIndex        =   9
         Top             =   900
         Width           =   3150
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
         TabIndex        =   7
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
         Left            =   -1260
         TabIndex        =   6
         Top             =   1680
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
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
      Height          =   1170
      Left            =   210
      TabIndex        =   0
      Top             =   4515
      Width           =   3585
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   13
         Top             =   735
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   12
         Top             =   450
         Width           =   2535
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu revertirseleccion 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "ANACOB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IAT&

Private Sub aprobar_Click()
   Call Command12_Click
End Sub

'
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
   Call ANACOBE
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command3.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   Label15 = VALACT1$("SELFECHA")
End Sub

'
Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture2.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.MARCOBARRA.Top = 5500
      Set MARCOBARRA.Container = Me.Picture2
      MARCOBARRA.ZOrder 0
    End If

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
    FECHOR$ = Control$("CANECOM", "FECHOCAL")
    Label7.Caption = Left$(FECHOR$, 8): Label7.Refresh
    Label95.Caption = Mid$(FECHOR$, 11, 5): Label95.Refresh
    '
    Label8 = Control$("CANECOM", "MODOCAL")   ' presenta el modo del último cálculo
    Label10 = Control$("CANECOM", "CLISELE")  ' presenta el modo del último cálculo
    '
    For U& = 1 To ULTREG&("TATIMAT")
       X$ = REGLEIDO$("TATIMAT", U&)
       NRO$ = Asc(Mid$(X$, 1, 1))
       NOMBRE$ = Mid$(X$, 2, 30)
       NRO1$ = Space(3 - Len(NRO$)) & NRO$
       TTXX$ = NRO1$ & Space(2) & NOMBRE$
       LICOPEN.AddItem TTXX$
    Next U&
    '
    ' Call VERNOVTRA
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub Label15_DblClick()
   Call Command9_Click
End Sub

Private Sub revertirseleccion_Click()
  Call Command3_Click
End Sub

Private Sub salir_Click()
   Call Command10_Click
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

Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Sub ANACOBE()
    '
    FECORTE% = FECHANUM(Label15)
    If FECORTE% < 1 Then FECORTE% = 32767
    '
    If DIASEM%(FECORTE%) <> 1 Then
       Call MENSERR(24, "En programación Semanal, la Fecha de Corte\Debe Corresponder a un Día LUNES.")
       Exit Sub
    End If
    '
    CI1% = CODINT%(Text1.TEXT)
    COD1$ = CODEXT$(CI1%)
    If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CODINT%(Text2.TEXT)
    COD2$ = CODEXT$(CI2%)
    If TRIM$(COD2$) = "" Then Exit Sub
    '
    NIVEFIL% = 0
    If Option1.Value = True Then NIVEFIL% = 1
    '
    Screen.MousePointer = 11
    Dim EXIS(32767, 10), TIENEMOVI%(32767)
    Dim ACTI%(256)
    '
    For U& = 1 To LICOPEN.ListCount
      If LICOPEN.Selected(U& - 1) = True Then
        TTXX$ = LICOPEN.List(U& - 1)
        TMT% = XVALO(Left$(TTXX$, 3))
        If TMT% > 0 Then
          If TMT% <= 255 Then
            ACTI%(TMT%) = 1
          End If
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("COBERFA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("COBERFA", U&)
       FEPRO% = CVI(Left$(X$, 2))
       If FEPRO% > FECORTE% Then
          GoTo 29
       End If
       CIXI% = CVI(Mid$(X$, 3, 2))
       TIMO% = CVI(Mid$(X$, 51, 2))
       CANTU = CVS(Mid$(X$, 53, 4)) - CVS(Mid$(X$, 57, 4))
       TIENEMOVI%(CIXI%) = 1
       Select Case TIMO%
         Case 10
            EXIS(CIXI%, 1) = EXIS(CIXI%, 1) + CANTU
         Case 11
            EXIS(CIXI%, 2) = EXIS(CIXI%, 2) + CANTU
         Case 13                            ' ORDENES DE COMPRA
            EXIS(CIXI%, 3) = EXIS(CIXI%, 3) + CANTU
         Case 12                            ' ORDENES DE FABRICACION
            EXIS(CIXI%, 4) = EXIS(CIXI%, 4) + CANTU
         Case 21                            ' PEDIDOS PENDIENTES DE ENTREGA
            EXIS(CIXI%, 5) = EXIS(CIXI%, 5) - CANTU
         Case 22                            ' RESERVAS O/F Y CONJUNTOS
            EXIS(CIXI%, 6) = EXIS(CIXI%, 6) - CANTU
         Case 20                            ' STOCK MINIMO DEL ITEM
            EXIS(CIXI%, 7) = EXIS(CIXI%, 7) - CANTU
         Case 15                            ' ORDENES DE FABRICACION A EMITIR
            EXIS(CIXI%, 8) = EXIS(CIXI%, 8) + CANTU
         Case 14                            ' ORDENES DE COMPRA A EMITIR
            EXIS(CIXI%, 9) = EXIS(CIXI%, 9) + CANTU
         Case Else
            Call MENSERR(24, "Codigo de Movimiento Incorrecto (" + TRIM$(Str$(TIMO%)) + ").")
            EXIS(CIXI%, 1) = EXIS(CIXI%, 1) + CANTU
       End Select
29  Next U&
    '
    JJ& = 0
    NMS% = NUMAS%
    For CIXI% = 1 To NUMAS%
      If TIENEMOVI%(CIXI%) > 0 Then
        CODE$ = CODEXT$(CIXI%)
        If CODE$ >= COD1$ Then
          If CODE$ <= COD2$ Then
            TMT% = TIMATE%(CIXI%)
            If TMT% = 0 Or ACTI%(TMT%) > 0 Then
              NETO = 0
              NETO = NETO + EXIS(CIXI%, 1) + EXIS(CIXI%, 2) + EXIS(CIXI%, 3) + EXIS(CIXI%, 4)
              NETO = NETO - EXIS(CIXI%, 5) - EXIS(CIXI%, 6) - EXIS(CIXI%, 7)
              If NETO > 0 Then EXIS(CIXI%, 10) = NETO
              '
              If NIVEFIL% < 1 Or EXIS(CIXI%, 8) + EXIS(CIXI%, 9) >= 0.05 Then
                 X$ = REGBLAN$("STODISP")
                 Call REPLA(X$, MKI$(CIXI%), 1, 2)
                 Call REPLA(X$, Chr$(TIMATE%(CIXI%)), 3, 1)
                 For K% = 1 To 10
                   Call REPLA(X$, MKS$(EXIS(CIXI%, K%)), 4 * K% + 1, 4)
                 Next K%
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("STODISP", X$, JJ&)
              End If
            End If
          End If
        End If
      End If
    Next CIXI%
    '
    Call SETULTREG("STODISP", JJ&)
    Call FILESORT("STODISP", "", 1, 1)
    '
    Call HEADERS("STODISP", "")
    Call HEADERS("STODISP", "Filtro: " + Label10)
    Call FINAL("*ANACOBE")
    '
End Sub
'

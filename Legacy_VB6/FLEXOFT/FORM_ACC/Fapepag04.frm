VERSION 5.00
Begin VB.Form FAPEPAG04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obligaciones Pendientes de Pago"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9240
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
   ScaleHeight     =   3345
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Grupo Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   105
      TabIndex        =   29
      Top             =   2415
      Width           =   4005
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
         Height          =   285
         Left            =   1260
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   2535
      End
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
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Top             =   360
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
         Height          =   285
         Left            =   360
         MaxLength       =   2
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3350
      Left            =   8295
      ScaleHeight     =   3285
      ScaleWidth      =   915
      TabIndex        =   23
      Top             =   0
      Width           =   975
      Begin VB.PictureBox Picture2 
         Height          =   4005
         Left            =   0
         ScaleHeight     =   3945
         ScaleWidth      =   1005
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   3450
            Left            =   0
            Picture         =   "Fapepag04.frx":0000
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1065
         End
      End
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
         Left            =   120
         Picture         =   "Fapepag04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Obligaciones a Pagar por Rango de Fecha."
         Top             =   1695
         Width           =   675
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2415
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   27
            Top             =   0
            Width           =   12000
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
         Height          =   690
         Left            =   140
         Picture         =   "Fapepag04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   140
         Picture         =   "Fapepag04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "Fapepag04.frx":2AF6
         Top             =   2760
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   105
      TabIndex        =   10
      Top             =   90
      Width           =   4005
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
         Height          =   360
         Left            =   2625
         TabIndex        =   16
         Top             =   420
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
         Height          =   360
         Left            =   3465
         TabIndex        =   15
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
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
         Left            =   315
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3330
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
         Height          =   360
         Left            =   2625
         TabIndex        =   13
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
         Height          =   360
         Left            =   3465
         TabIndex        =   12
         Top             =   1365
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   315
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1785
         Width           =   3330
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
         TabIndex        =   18
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         TabIndex        =   17
         Top             =   1470
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
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
      Height          =   765
      Left            =   4200
      TabIndex        =   7
      Top             =   2450
      Width           =   4005
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
         Height          =   650
         Left            =   3360
         Picture         =   "Fapepag04.frx":4A18
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   105
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   9
         Top             =   420
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1890
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4200
      TabIndex        =   6
      Top             =   1620
      Width           =   4005
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   21
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   3570
         TabIndex        =   20
         Top             =   270
         Width           =   175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Vencimiento:"
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
         TabIndex        =   22
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   4200
      TabIndex        =   1
      Top             =   90
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Fecha de Vencimiento"
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
         Index           =   3
         Left            =   420
         TabIndex        =   5
         Top             =   1080
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Alfabético"
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
         Left            =   2205
         TabIndex        =   4
         Top             =   720
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numérico"
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
         Left            =   735
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Proveedor"
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
         Left            =   420
         TabIndex        =   2
         Top             =   360
         Width           =   1905
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu consulistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu generareporte 
         Caption         =   "Facturas Pendientes de Pago (según Filtro en Pantalla)"
      End
      Begin VB.Menu obligacionapagar 
         Caption         =   "Obligaciones a Pagar por Rango de Fecha."
      End
   End
End
Attribute VB_Name = "FAPEPAG04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TICUE%(32767)
Dim MODOMONEI%

Private Sub Command1_Click()
   Call SELECHO("GRUCOPRO")
   VDEVU$ = VALACT1$("GRUCOPRO")
   If TRIM$(VDEVU$) <> "" Then
      Text1.TEXT = VDEVU$
   End If
End Sub

Private Sub Command2_Click()
    Call FAPEPAG2
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COPAPE1")
            Exit Sub
          Else
            Call FINAL("*COPAPE2")
            Exit Sub
         End If
       Else                          ' impresora
         If MODOMONEI% = 2 Then
             If Option1(3).Value = False Then
                Call FINAL("*LIPAPED1")
                Exit Sub
              Else
                Call FINAL("*LIPAPED2")
                Exit Sub
             End If
           Else
             If Option1(3).Value = False Then
                Call FINAL("*LIPAPE1")
                Exit Sub
              Else
                Call FINAL("*LIPAPE2")
                Exit Sub
             End If
         End If
    End If
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPEPAG1
        CARFAPEN% = 1
    End If
    '
End Sub


Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If DES% < 1 Then GoTo 99
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    CORRESP$ = PERIO$
    Call HEADERS("VENXRAFE", "")
    Call HEADERS("VENXRAFE", "Corresponde a: " + CORRESP$)
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & HAS1# & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))>= " & DES1# & " "
    SQLX$ = SQLX$ + "AND (SaldAcre)>= 0.005 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set VEXRAFETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VEXRAFETMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Vencimientos en el Rango de Fecha Seleccionado")
        GoTo 99
      End If
      On Error GoTo 0
      J& = 0
      Y$ = REGBLAN("VENXRAFE")
      Do While .EOF = False
          '
          NC% = !Nuprove
          NFS$ = !CodiCom_Lar
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = !SaldAcre - !SaldDebe
          '
          Call REPLA(Y$, MKI$(NC%), 1, 2)
          Call REPLA(Y$, CODICLI$((-1) * NC%), 3, 24)
          Call REPLA(Y$, RASOCLI$((-1) * NC%), 27, 30)
          Call REPLA(Y$, NFS$, 57, 18)
          Call REPLA(Y$, MKI$(FVI%), 75, 2)
          Call REPLA(Y$, MKD$(SALDEU#), 77, 8)
          '
          J& = J& + 1
          Call GRAREG("VENXRAFE", Y$, J&)
          '
        .MoveNext
        '
      Loop
      '
      Call SETULTREG("VENXRAFE", J&)
      Call FILESORT("VENXRAFE", "VENXRAFE", 5, 5, "ASC")
      Call FILESORT("VENXRAFE", "VENXRAFE", 1, 1, "ASC")
      '
    End With
    '
    Call FINAL("*LIVERAFE")
    '
99  Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("FAPEPAGO")
End Sub

Private Sub Command8_Click()
    Call SELECHO("ACREDOR")
    NCMAX% = XVALO(VALACT1$("ACREDOR"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ACREDOR")
    NCMIN% = XVALO(VALACT1$("ACREDOR"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture2.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Option1(0).BackColor = &HFCD7B6
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      Me.Option1(3).BackColor = &HFCD7B6
      Me.Option1(4).BackColor = &HFCD7B6
      Me.Option1(8).BackColor = &HFCD7B6
      MARCOBARRA.Top = 3000
      Set MARCOBARRA.Container = Me.Picture2
      MARCOBARRA.ZOrder 0
    End If
    
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
End Sub

Private Sub generareporte_Click()
    Call Command2_Click
End Sub

Private Sub obligacionapagar_Click()
    Call Command5_Click
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub salir_Click()
     Call Command3_Click
End Sub

Private Sub Text1_Change()
   ECO$ = ECOARCH$("GRUCOPRO", MKI$(XVALO(Text1)))
   '
   If ECO$ <> "" Then
        Text2 = ECO$
      Else
        Text2 = "No Definido"
   End If
End Sub

Private Sub Text1_DblClick()
   Call Command1_Click
End Sub

Private Sub Text10_Change()
    NCMIN% = XVALO(Text10.TEXT)
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text8_Change()
    NCMAX% = XVALO(Text8.TEXT)
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    Text7.Refresh
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ACREDOR")
    NCMIN% = XVALO(VALACT1$("ACREDOR"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("ACREDOR")
    NCMAX% = XVALO(VALACT1$("ACREDOR"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    Dim NROPO$(8196), CLIPO%(8196)
    Dim TICUE%(32767), SARRACLI#(32767), SUDEUCLI#(32767)
    'Dim TIS$(10)
    'TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    'For I% = 0 To 7
    '    TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    'Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("CUECORR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCLIE% = CVI(Mid$(X$, 5, 2))
      If NCLIE% > 0 Then
        If TICUE%(NCLIE%) = 0 Then
          FEDOC% = CVI(Left$(X$, 2))
          If FEDOC% <= FLIM% Then
            IMPO# = CVD(Mid$(X$, 123, 8))
            IMPO# = IMPO# - CVD(Mid$(X$, 115, 8))
            SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
            SARRATO# = SARRATO# + IMPO#
            '
            SALDEX$ = Mid$(X$, 139, 60)
            For KKI% = 1 To 6
              FVI% = CVI(Mid$(SALDEX$, 10 * KKI% - 9, 2))
              SALDO# = CVD(Mid$(SALDEX$, 10 * KKI% - 7, 8))
              If FVI% > 0 Then
                If Abs(SALDO#) >= 0.005 Then
                  FF% = CVI(Left$(X$, 2))
                  NC% = CVI(Mid$(X$, 5, 2))
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
                  ES$ = "    "
                  If SALDO# < 0 Then
                      ES$ = " CR "
                    Else
                      If FVI% <= HOO% + 8 Then ES$ = " ** "
                      If FVI% <= HOO% Then ES$ = "!!!!"
                  End If
                  '
                  Z$ = REGBLAN$("FAPEPAG")
                  Call REPLA(Z$, MKI$(NC%), 1, 2)
                  Call REPLA(Z$, NFS$, 3, 18)
                  Call REPLA(Z$, MKI$(FF%), 21, 2)
                  Call REPLA(Z$, MKI$(FVI%), 23, 2)
                  Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                  Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
                  Call REPLA(Z$, ES$, 41, 4)
                  '
                  UJ& = UJ& + 1
                  Call GRAREG("FAPEPAG", Z$, UJ&)
                  SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDO#
                  SUDEUTO# = SUDEUTO# + SALDO#
                  '
                End If
              End If
            Next KKI%
          End If
          If FEDOC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO$(CAMOPO%) = Mid$(X$, 7, 18)
            CLIPO%(CAMOPO%) = NCLIE%
          End If
        End If
      End If
    Next U&
    '
    '
    ' LEER APLIPAG - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLIPAG")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLIPAG", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            OPOSTE$ = Mid$(X$, 5, 18)
            DOCUORI$ = Mid$(X$, 55, 18)
            'If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   'AAA = BBB
            'End If
            For KKI% = 1 To CAMOPO%
              If NROPO$(KKI%) = DOCUORI$ Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            'GoTo 29
            '
24          Z$ = REGBLAN$("FAPEPAG")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
            NFS$ = DOCUORI$
            Call REPLA(Z$, NFS$, 3, 18)
            FF% = FLIM%
            Call REPLA(Z$, MKI$(FF%), 21, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 23, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPEPAG", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
   '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          RFF$ = RECFILT$("CUECORR", 4, MKI$(CLIAPLI%))
          For U& = Len(RFF$) To 1 Step -4
            REG& = CVS(Mid$(RFF$, U& - 3, 4))
            If REG& > 0 Then
              If REG& <= ULTREG&("CUECORR") Then
                X$ = REGLEIDO$("CUECORR", REG&)
                If CVI(Left$(X$, 2)) <= FLIM% Then
                  FF% = CVI(Left$(X$, 2))
                  FVI% = FLIM%
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
                  ES$ = "    "
                  If Abs(IMORIG# - DIFESAL#) < 0.005 Then
                    SALDO# = IMORIG#
                    '
                    Z$ = REGBLAN$("FAPEPAG")
                    Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
                    Call REPLA(Z$, NFS$, 3, 18)
                    Call REPLA(Z$, MKI$(FF%), 21, 2)
                    Call REPLA(Z$, MKI$(FVI%), 23, 2)
                    Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                    Call REPLA(Z$, MKD$(Abs(IMORIG#)), 33, 8)
                    Call REPLA(Z$, ES$, 41, 4)
                    '
                    UJ& = UJ& + 1
                    Call GRAREG("FAPEPAG", Z$, UJ&)
                    SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDO#
                    SUDEUTO# = SUDEUTO# + SALDO#
                    GoTo 39
                    '
                  End If
                End If
              End If
            End If
          Next U&
      End If
39  Next UI&
          
              
     For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
           
          Z$ = REGBLAN$("FAPEPAG")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 18)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 21, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 23, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
          IMORIG# = Abs(DIFESAL#)
          Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPEPAG", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          SUDEUTO# = SUDEUTO# + SALDEU#
          '
      End If
49  Next UI&
    '
50  Call SETULTREG("FAPEPAG", UJ&)
    Call CIERRARCH("FAPEPAG")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    Screen.MousePointer = 11
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPEPAG")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPEPAG", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 23, 2))
            '
            TLISTA$ = Space$(36)
            If MODORDE% = 1 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
              ElseIf MODORDE% = 2 Then
                 Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
              ElseIf MODORDE% = 3 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
            End If
            Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            List1.AddItem TLISTA$
            '
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = XVALO(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPEPAG", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 18)
        FF% = CVI(Mid$(X$, 21, 2))
        FVI% = CVI(Mid$(X$, 23, 2))
        SALDEU# = CVD(Mid$(X$, 25, 8))
        IMORIG# = CVD(Mid$(X$, 33, 8))
        ES$ = Mid$(X$, 41, 4)
        '
        Y$ = REGBLAN$("FAPEPAGX")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$((-1) * NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$((-1) * NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 18)
        Call REPLA(Y$, MKI$(FF%), 67, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 69, 8)
        Call REPLA(Y$, MKI$(FVI%), 77, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
        Call REPLA(Y$, ES$, 95, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPEPAGX", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    '
    FEX = FLIM%
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIPAPE1")
    '
    Screen.MousePointer = 1

End Sub


Sub FAPEPAG1()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
      If NCX% > 0 Then
        If NCX% > NCMAX% Then NCMAX% = NCX%
        If NCX% < NCMIN% Then NCMIN% = NCX%
        TICUE%(NCX%) = 0
      End If
    Next U&
    Text10 = TRIM$(Str$(NCMIN%))
    Text9 = RASOCLI$((-1) * NCMIN%)
    Text8 = TRIM$(Str$(NCMAX%))
    Text7 = RASOCLI$((-1) * NCMAX%)
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    FLIM = HOY(HOS$)
    With SACTEMP
      .MoveFirst
      If Err < 1 Then
        FLIM = CVINT(!FEVENCIM)
      End If
    End With
    On Error GoTo 0
    '
    Text6 = FECHATEX$(FLIM)
    Text6.Refresh
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub FAPEPAG2()
    '
    Screen.MousePointer = 11
    FLIM = CDbl(CVDAT(Int(FECHANUM(Text6.TEXT))))
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCOSEL% = XVALO(Text1)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    SQLX$ = "SELECT * FROM SACREPEN "
    SQLX$ = SQLX$ + "WHERE Status < 3 "
    SQLX$ = SQLX$ + "AND NuProve >= " & NCMIN% & " "
    SQLX$ = SQLX$ + "AND NuProve <= " & NCMAX% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FeVencim))<= " & FLIM & " "
    SQLX$ = SQLX$ + "AND ((SaldAcre + SaldDebe)>= 0.005 "
    SQLX$ = SQLX$ + "OR (SaldAcre_Dola + SaldDebe_Dola)>= 0.005) "
    SQLX$ = SQLX$ + "ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With SACTEMP
      .MoveLast
     If Err < 1 Then
        FIN& = .RecordCount
        On Error GoTo 0
        .MoveFirst
        UJ& = 0
        Do While Not .EOF
          Call TRACE(20, UJ&, FIN&)
          NC% = !Nuprove
          If TICUEPRO%((-1) * NC%) = 0 Then ' SOLO CUENTAS CORRIENTES
            If GRUCOSEL% < 1 Or GRUCOCLI%((-1) * NC%) = GRUCOSEL% Then
              NFS$ = !CodiCom_Lar
              FF% = CVINT(!FechEmisi)
              FVI% = CVINT(!FEVENCIM)
              SALDEU# = !SaldAcre - !SaldDebe
              SALDEUDOL# = 0: MEMI% = 0
              On Error Resume Next
              SALDEUDOL# = !SaldAcre_Dola - !SaldDebe_Dola
              MEMI% = !Mone_Emis: If MEMI% < 0 Then MEMI% = 0
              If MEMI% < 2 Then SALDEUDOL# = 0
              '
              On Error GoTo 0
              ' ESTA ES PARA DESCARTAR LOS CASOS DE FACTURAS EN PESOS QUE QUEDARON
              ' CON ALGÚN SALDITO EN DOLARES - RECLAMO CLAUDIA MATSUSAKA - CT-100294-FL
              If Abs(SALDEU#) < 0.005 Then
                If Abs(SALDEUDOL#) < 0.005 Then
                  GoTo 29
                End If
              End If
              IMORIG# = !IBruComp
              ES$ = "    "
              '
              Y$ = REGBLAN$("FAPEPAGX")
              Call REPLA(Y$, MKI$(NC%), 1, 2)
              Call REPLA(Y$, RASOCLI$((-1) * NC%), 3, 30)
              Call REPLA(Y$, TEL1CLI$((-1) * NC%), 33, 16)
              Call REPLA(Y$, NFS$, 49, 18)
              Call REPLA(Y$, MKI$(FF%), 67, 2)
              Call REPLA(Y$, MKD$(IMORIG#), 69, 8)
              Call REPLA(Y$, MKI$(FVI%), 77, 2)
              Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
              Call REPLA(Y$, ES$, 95, 4)
              '\\\OT-06-0104-WAR-23/03/06///
              LEYMAN1$ = ".........."
              Call REPLA(Y$, LEYMAN1$, 99, 10)
              Call REPLA(Y$, LEYMAN1$, 109, 10)
              Call REPLA(Y$, LEYMAN1$, 119, 10)
              '\\\OT-06-0104-WAR-FIN///
              UJ& = UJ& + 1
              Call GRAREG("FAPEPAGX", Y$, UJ&)
              '
              '//OT-07-0408
              If MODOMONEI% = 2 Then
                 Z$ = REGBLAN$("FAPEPAGD")
                 Call REPLA(Z$, Left$(Y$, 94), 1, 94)
                 Call REPLA(Z$, Mid$(Y$, 95, 34), 115, 34)
                 Call REPLA(Z$, SIMBOLPES$(MEMI%), 95, 4)
                 Call REPLA(Z$, MKD$(SALDEUDOL#), 99, 8)
                 If MEMI% > 1 Then
                    TIPOCAM# = TICAMISORP(NFS$, FF%, NC%)
                    Call REPLA(Z$, MKD$(TIPOCAM#), 149, 8)
                 End If
                 Call GRAREG("FAPEPAGD", Z$, UJ&)
               '//OT-07-0408
              End If
            End If
          End If
29      .MoveNext
        Loop
        '
     End If '*
    End With
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Rango: Cuentas " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)) + " ... Vencim.Hasta: " + Text6)
    '
    ' ORDENA POR numero de comprobante
    Call FILESORT("FAPEPAGX", "", 4, 4, "ASC")
    ' ORDENA POR FECHA DE VENCIMIENTO
    Call FILESORT("FAPEPAGX", "", 7, 7, "ASC")
    ' AHORA ORDENA POR MODO SELECCIONADO
    Select Case MODORDE%
      Case 1
        Call FILESORT("FAPEPAGX", "", 1, 1, "ASC")
      Case 2
        Call FILESORT("FAPEPAGX", "", 2, 2, "ASC")
    End Select
    '
    If MODOMONEI% = 2 Then
       Call SETULTREG("FAPEPAGD", UJ&)
       Call CIERRARCH("FAPEPAGD")
       '
       Call HEADERS("FAPEPAGD", "")
       Call HEADERS("FAPEPAGD", "Rango: Cuentas " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)) + " ... Vencim.Hasta: " + Text6)
       '
       ' ORDENA POR numero de comprobante
       Call FILESORT("FAPEPAGD", "", 4, 4, "ASC")
       ' ORDENA POR FECHA DE VENCIMIENTO
       Call FILESORT("FAPEPAGD", "", 7, 7, "ASC")
       ' AHORA ORDENA POR MODO SELECCIONADO
       Select Case MODORDE%
         Case 1
           Call FILESORT("FAPEPAGD", "", 1, 1, "ASC")
         Case 2
           Call FILESORT("FAPEPAGD", "", 2, 2, "ASC")
       End Select
    End If
    Screen.MousePointer = 1
    '
End Sub
'



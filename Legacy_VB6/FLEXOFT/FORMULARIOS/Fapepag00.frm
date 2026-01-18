VERSION 5.00
Begin VB.Form FAPEPAG00 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obligaciones Pendientes de Pago"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
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
   ScaleHeight     =   3585
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8610
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   25
      Top             =   2625
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   12000
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
      Left            =   8610
      Picture         =   "Fapepag00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   420
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
      Left            =   8610
      Picture         =   "Fapepag00.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   105
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
      Height          =   2220
      Left            =   210
      TabIndex        =   12
      Top             =   315
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      Height          =   810
      Left            =   210
      TabIndex        =   9
      Top             =   2625
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
         Height          =   500
         Left            =   3255
         Picture         =   "Fapepag00.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   210
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      Height          =   1040
      Left            =   4410
      TabIndex        =   8
      Top             =   2415
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
         TabIndex        =   23
         Top             =   420
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
         TabIndex        =   22
         Top             =   420
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
         TabIndex        =   24
         Top             =   525
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
      Height          =   2010
      Left            =   4410
      TabIndex        =   3
      Top             =   315
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
         TabIndex        =   7
         Top             =   1470
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
         TabIndex        =   6
         Top             =   840
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
         TabIndex        =   5
         Top             =   840
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
         TabIndex        =   4
         Top             =   420
         Width           =   1905
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8085
      Picture         =   "Fapepag00.frx":075E
      Top             =   2940
      Width           =   2010
   End
End
Attribute VB_Name = "FAPEPAG00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TICUE%(32767)

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
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
         If Option1(3).Value = False Then
            Call FINAL("*LIPAPE1")
            Exit Sub
          Else
            Call FINAL("*LIPAPE2")
            Exit Sub
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
    NCMAX% = Val(VALACT1$("ACREDOR"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ACREDOR")
    NCMIN% = Val(VALACT1$("ACREDOR"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text10_Change()
    NCMIN% = Val(Text10.TEXT)
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
End Sub

Private Sub Text8_Change()
    NCMAX% = Val(Text8.TEXT)
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ACREDOR")
    NCMIN% = Val(VALACT1$("ACREDOR"))
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
    NCMAX% = Val(VALACT1$("ACREDOR"))
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
If NC% = 434 Then
AAA = BBB
End If
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
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
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
        REG& = Val(Mid$(List1.TEXT, 31, 8))
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
    Text10.TEXT = TRIM$(Str$(NCMIN%))
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    Text8.TEXT = TRIM$(Str$(NCMAX%))
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    FLIM = HOY(HO$)
    FIN& = ULTREG&("CUECORR")
    RFF$ = ""
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCX% = CVI(Mid$(X$, 5, 2))
      If NCX% > 0 Then
        If TICUE%(NCX%) = 0 Then
          SALDEU$ = Mid$(X$, 139, 60)
          ATX% = 0
          For KKI% = 1 To 6
            FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
            SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
            If Abs(SALDO#) > 0 Then
              ATX% = 1
              If FEVE > FLIM Then FLIM = FEVE
            End If
          Next KKI%
          If ATX% > 0 Then
            RFF$ = RFF$ + MKS$(U&)
          End If
        End If
      End If
    Next U&
    Text6.TEXT = FECHATEX$(FLIM)
    Text6.Refresh
    '
    UJ& = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        SALDEU$ = Mid$(X$, 139, 60)
        For KKI% = 1 To 6
           FVI% = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
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
                 '
              End If
           End If
        Next KKI%
    Next U&
    '
    Call SETULTREG("FAPEPAG", UJ&)
    Screen.MousePointer = 1
    '
End Sub
'

Sub FAPEPAG2()
    '
    Screen.MousePointer = 11
    FLIM = FECHANUM(Text6.TEXT): If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
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
            If FVI% <= FLIM Then
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
        REG& = Val(Mid$(List1.TEXT, 31, 8))
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
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Vencimiento Hasta: " + FECHATEX$(FLIM))
    Call HEADERS("FAPEPAGX", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
    Screen.MousePointer = 1
    '
End Sub
'



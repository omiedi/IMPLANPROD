VERSION 5.00
Begin VB.Form OPOCAB04 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra Abiertas - Proveedores Externos"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00006060&
      Height          =   4845
      Left            =   4200
      ScaleHeight     =   4785
      ScaleWidth      =   1005
      TabIndex        =   21
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton Command19 
         Caption         =   "N.Orig"
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
         Picture         =   "OPOCAB04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Códigos Equivalentes de Proveedor"
         Top             =   2835
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   3240
         Visible         =   0   'False
         Width           =   555
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
         Picture         =   "OPOCAB04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   1005
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
         Picture         =   "OPOCAB04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command18 
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
         Picture         =   "OPOCAB04.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1695
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   125
         ScaleHeight     =   75
         ScaleWidth      =   495
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3675
         Width           =   555
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   23
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "OPOCAB04.frx":09E0
         Stretch         =   -1  'True
         Top             =   3990
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Gestión de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   210
      TabIndex        =   3
      Top             =   105
      Width           =   3795
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Autorizaciones de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   210
         TabIndex        =   29
         Top             =   2730
         Width           =   3375
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2745
            TabIndex        =   30
            Top             =   450
            Width           =   2800
            Begin VB.CommandButton Command2 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   2205
               Picture         =   "OPOCAB04.frx":2902
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Listado de Autorizaciones de Entrega Pendientes"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   0
               Picture         =   "OPOCAB04.frx":2C0C
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Autorización de Entrega"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Autorización de Entrega"
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
               TabIndex        =   31
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Ordenes de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   3375
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   715
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   5
            Top             =   1250
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
               Left            =   0
               Picture         =   "OPOCAB04.frx":2F16
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Anulación de Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación de O/Compras"
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
               TabIndex        =   6
               Top             =   105
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   515
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   7
            Top             =   850
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Picture         =   "OPOCAB04.frx":3220
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir"
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
               TabIndex        =   8
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   450
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
               Picture         =   "OPOCAB04.frx":352A
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Modificar Orden de Compra no Anulada (Sólo Precios)"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificar O/C"
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
               TabIndex        =   10
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pedidos de Suministro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   630
      TabIndex        =   11
      Top             =   5460
      Visible         =   0   'False
      Width           =   3165
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   18
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command6 
            Enabled         =   0   'False
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
            Picture         =   "OPOCAB04.frx":3674
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Generar un Pedido de Suministro"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Pedido de Suministro"
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
            TabIndex        =   20
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   795
         Width           =   2325
         Begin VB.CommandButton Command17 
            Enabled         =   0   'False
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
            Picture         =   "OPOCAB04.frx":397E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Adjudicación de Pedido de Suministro a Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignacion de Proveedor"
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
            TabIndex        =   17
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   1260
         Width           =   2325
         Begin VB.CommandButton Command21 
            Enabled         =   0   'False
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
            Picture         =   "OPOCAB04.frx":3C88
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Pedidos de Suministro Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "PS's Pendientes"
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
            TabIndex        =   14
            Top             =   105
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "OPOCAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub


Sub Command12_Click()
    '
    Command12.Enabled = False
    OPX% = COMALTER%("Detalle de Ultimas Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command12.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call OPLISCO04.Command12_Click
    Else
      Call OPLISCO04.Command20_Click
    End If
    Command12.Enabled = True
    '
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
    Call MODIOCOM
99  Command13.Enabled = True
    '
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   Call OCOMANU
   Command14.Enabled = True
End Sub

Sub Command15_Click()
   Command15.Enabled = False
   Call DEVOLCONSI
   Command15.Enabled = True
   '
   'If EXISTE%("ORTRAB04.EXE") > 0 Then
   '    Call MENSERR(24, "No Compatible con Módulo de Ordenes de Trabajo.")
   '  Else
   '    Call OPNOIN("")
   'End If
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: K& = 0: OCOA& = 0
    '
    For I& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", I&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             K& = K& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, K&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next I&
    '
    Call SETULTREG("OCOMPRA", K&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, UNIMED$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(RASOCLI$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If Val(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Sub Command18_Click()
    '
    Command18.Enabled = False
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      Call CONECRUN("FSETUP04/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    '
    Command19.Enabled = False
    FORCOMP04.Nuorig.Enabled = False
    FORCOMP04.Nuorig1.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
      FORCOMP04.Nuorig.Enabled = True
      FORCOMP04.Nuorig1.Enabled = True
    End If
    '
    Command18.Enabled = True
    '
End Sub

Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORPR")
   Command19.Enabled = True
End Sub

Sub Command2_Click()
  Command2.Enabled = False
    Call SELECHO("FECHA/Fecha Límite Hasta")
    VDEVU$ = VALACT1$("FECHA")
    FLIM% = FECHANUM(VDEVU$)
    FM1# = CDbl(CVDAT(FLIM%))
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM AUTOENT "
    SQLX$ = SQLX$ + " WHERE STAT_ATEN = 0"
    If FLIM% > 0 Then
      SQLX$ = SQLX$ + " AND CDBL(Fentre_Sol) <= " & FM1# & ""
    End If
    Set ATENTMP = Compras.OpenRecordset(SQLX$, 4)
    With ATENTMP
      On Error Resume Next
      .MoveFirst
      If Err = 1 Then
        Call COMUNI("No Existen Entregas Pendientes")
        On Error GoTo 0
        GoTo 99
      End If
      On Error GoTo 0
      Y$ = REGBLAN$("LIAUTOEN")
      JJ& = 0
      Do While .EOF = False
       CI% = !Cod_inte
       FES1% = CVINT%(!Fentre_Sol)
       QA1# = !Cant_Aten
       QP1# = !Sald_Aten
       NC% = !Npro_Aten
       FEAE1% = CVINT%(!Femi_Aten)
       '
       Call REPLA(Y$, MKS$(!Nume_Aten), 1, 4)
       Call REPLA(Y$, MKS$(!Nume_Ocom), 5, 4)
       Call REPLA(Y$, MKI$(FEAE1%), 9, 2)
       Call REPLA(Y$, MKI$(NC%), 11, 2)
       Call REPLA(Y$, MKI$(CI%), 13, 2)
       Call REPLA(Y$, MKS$(QA1#), 15, 4)
       Call REPLA(Y$, MKI$(FES1%), 19, 2)
       Call REPLA(Y$, MKS$(QP1#), 21, 4)
       Call REPLA(Y$, !Observa_Aten, 25, 30)
       JJ& = JJ& + 1
       Call GRAREG("LIAUTOEN", Y$, JJ&)
       .MoveNext
      Loop
    End With
    Call SETULTREG("LIAUTOEN", JJ&)
    '
    Call FILESORT("LIAUTOEN", "LIAUTOEN", 9, 9, "ASC") 'FECHA
    Call FILESORT("LIAUTOEN", "LIAUTOEN", 6, 6, "ASC") 'ART
    Call FILESORT("LIAUTOEN", "LIAUTOEN", 4, 4, "ASC") 'PROV
    '
    Call FINAL("*LIAUTOEN")
    '
99 Command2.Enabled = True
End Sub


Sub Command20_Click()
    '
    Command20.Enabled = False
    OPX% = COMALTER%("Ranking de Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command20.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call RANKPRO
    Else
      Call RANKART
    End If
    Command20.Enabled = True
    '
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

Sub Command22_Click()
    '
    Command22.Enabled = False
    Call LISOCAB04.Command22_Click
    Command22.Enabled = True
    '
End Sub

Sub Command23_Click()
   Command23.Enabled = False
   Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command23.Enabled = True
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
        CONTROCOS.Show 1
      Else
20      Call OPNOIN("Control Costos de Reposición")
    End If
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
   Command4.Enabled = False

'PROCESO DE AUTORIZACIÓN DE ENTREGA
'******* ** ************ ** *******

'1.- SELECCION DE LA OC ABIERTA
5     Screen.MousePointer = 11
    Call ABRECOMPRAS
    SQLX$ = " SELECT  DISTINCT ENCOCABI.Nume_Ocom, ENCOCABI.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM ENCOCABI "
    SQLX$ = SQLX$ + " INNER JOIN  DETOCABI "
    SQLX$ = SQLX$ + " ON (ENCOCABI.Nume_Ocom = DETOCABI.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE DETOCABI.Stat_Ocom <8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND ENCOCABI.Stat_Ocom < 1 "
    SQLX$ = SQLX$ + " AND ENCOCABI.Stat_Ocom >= 0 "
    SQLX$ = SQLX$ + " ORDER BY ENCOCABI.Nume_Ocom  ASC "
    Set OCANULTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCANULTEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No hay Ordenes de Compra Abiertas\Pendientes que Puedan Autorizarse.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        J& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
            J& = J& + 1
            Call GRAREG("!INDIOCPE", XXY$, J&)
            .MoveNext
        Loop
    End With
    '
    Call SETULTREG("!INDIOCPE", J&)
    Call CIERRARCH("!INDIOCPE")
    Call BAJALDISCO
    Call FRESHECHO("!INDIOCPE")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra Abiertas\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("!INDIOCPE/Indice de Ordenes de Compra Abiertas Pendientes")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
'2.- SELECCIONAR ARTICULOS DE LA O/C
15  AUTOENT04.Label14 = NUOCO&
    '
    AUTOENT04.LICOPEN.ListIndex = 0
    AUTOENT04.Label18 = Str$(AUTOENT04.PRONUAE&)
    HOII% = HOY(HOS$)
    AUTOENT04.Text14 = FECHATEX$(HOII%)
    AUTOENT04.Show 1
    '
999 Command4.Enabled = True
End Sub
   '
Function PRONUAE&()
   '
   PRONU& = 1
   Call ABRECOMPRAS
   '
   SQLX$ = "SELECT Nume_Aten FROM AutoEnt "
   SQLX$ = SQLX$ + "ORDER BY Nume_Aten DESC;"
   Set AutoEnTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With AutoEnTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Aten
     End If
   End With
   On Error GoTo 0
   '
   PRONUAE& = PRONU&
   '
End Function
Private Sub Command6_Click()
    Command6.Enabled = False
    Call PEDISUMI
    'Call OPNOIN("Pedido de Suministro")
    Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") < 1 Then GoTo 20
    OPX% = ALTERNA%("Control Costos Reposición\Otros Accesos Directos")
    If OPX% = 1 Then
        CONTROCOS.Show 1
      ElseIf OPX% = 2 Then
20      Call ACCDIR("OCOMPRA")
    End If
    Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Sub PEDISUMI()             ' NUEVO PEDIDO DE SUMINISTRO
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
    X0$ = REGBLAN$("OCOMPRA")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 9, 2)
    Call REPLA(X0$, "S", 57, 1)
    '
101 OBJE$ = "PEDISUMI"
    '
    X1$ = OBJPLA$(OBJE$, X0$)
    If X1$ = "" Then GoTo 299
    '
    X0$ = X1$
    NRO& = CVS(Left$(X1$, 4))
    If NRO& < 1 Then GoTo 101
    '
    FEMI% = CVI(Mid$(X1$, 5, 2))
    If FEMI% < 1 Then
        Call MENSERR(24, "Falta Fecha de Emisión")
        GoTo 101
    End If
    '
    NPX$ = MKI$(0)
    NPRO% = CVI(NPX$)
    '
    Call BLANARCH("!CARDETOC")
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - Uso Interno"
    VTB% = VENTABU%("CARPESU1/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    CAICO% = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       CANTU = CVS(Mid$(X$, 47, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       CAICO% = J&
       '
       Call REPLA(X$, UNIMED$(CI0%), 3, 4)
       If CVI(Mid$(X$, 67, 2)) < HO% Then
          Call REPLA(X$, MKI$(HO%), 67, 2)
       End If
       Call REPLA(X$, "  ", 69, 2)
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!CARDETOC", X$, J&)
       '
    Next J&
    '
    If CAICO% < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
103 VTB% = VENTABU%("CARPESU2/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       '
       If FENTRE% < FEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FEMI%), 9, 2)
          GoTo 103
       End If
       '
    Next J&
    '
    Call CIERRARCH("!CARDETOC")
    '
    Call BLOQARCH("OCOMPRA")
    Call BLOQARCH("MACONSIG")
    '
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PREUNID# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = X1$
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       'Call REPLA(X2$, MKS$(0), 58, 4)
       'Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For I% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(I%)
             CANTI# = CDbl(CANTU * USOI(I%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, UNIMED$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
               End If
             End If
          Next I%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Pedido de Suministro"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CAPARDOC% = 6
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CACOLTAB1% = 10
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PREUNID# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
        '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    
'consignacion inmediata si/no

    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CoCoPen()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = Asc(Mid$(XXX$, 122, 1))
         If MONEMI% <= 1 Then MONEX$ = "    "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         PREUNID# = CVD(Mid$(XXX$, 79, 8)) * TICAMBIO
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
         Call REPLA(XXY$, MKD$(PREUNID#), 133, 8)
         Call REPLA(XXY$, MKD$(PREUNID# * CAPEN), 141, 8)
         JJ& = JJ& + 1
         Call GRAREG("OPENDETA", XXY$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   Screen.MousePointer = 1
   '
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   OPX% = COMALTER%("Elija su Opción de Control:\\Cumplimiento\Calidad\Cancelar")
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  GoTo 10
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("BOLRECEP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
            CAPEN(K&) = CAPEN(K&) - CAREC
            GoTo 19
          End If
        End If
      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
      Z0$ = REGBLAN$("COCOBREC")
      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
      FENTREPRO = CVI(Mid$(X$, 95, 2))
      '
      If Len(PUBOL$(K&)) < 4 Then
           DIATRA% = 0
           If FENTREPRO < HOII Then
             DIATRA% = DIENTRE%(FENTREPRO, HOII)
           End If
           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
           NCOCO& = NCOCO& + 1
           Call GRAREG("COCOBREC", Z0$, NCOCO&)
         Else
           For II& = 1 To Len(PUBOL$(K&)) Step 4
             Z$ = Z0$
             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
             FERECI = CVI(Mid$(Y$, 81, 2))
             DIATRA% = 0
             If FENTREPRO < FERECI Then
               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
             End If
             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
             NCOCO& = NCOCO& + 1
             Call GRAREG("COCOBREC", Z$, NCOCO&)
           Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("BOLRECEP")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                CI% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(CI%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                JJ& = JJ& + 1
                Call GRAREG("COCAREC", Z$, JJ&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
    '
    Command19.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    If EXISTE%("ESTOCK02.EXE") > 0 Then
      Command23.Enabled = True
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      Command3.Enabled = True
    End If
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub MODIOCOM()         ' MODIFICAR ORDEN DE COMPRA
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORMOCOM"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    AVALIDA$ = "MASTER"
    '
    'Command22.Enabled = False
    'Cargo la lista con las O/C
    Call ABRECOMPRAS
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI "
    SQLX$ = SQLX$ + "WHERE Stat_Ocom < 9 "   ' EXCLUYE ANULADAS
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set VRIMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 999
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INDIOCOM", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Screen.MousePointer = 1
    '
    Call FRESHECHO("!INDIOCOM")
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas)")
    NUOCO& = Val(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NUOCO& < 1 Then GoTo 999
    '
    Call ABRECOMPRAS
    With EncoCabi
      .FindFirst "Nume_Ocom = " & NUOCO& & "AND Stat_Ocom < 9 "
      If .NoMatch = True Then
        Call MENSERR(24, "Imposible Abrir la O/Compra " & NUOCO& & " - Consulte !")
        GoTo 3
      End If
      '
      If !Stat_Ocom > 8 Then ' esta anulada
        Call MENSERR(24, "O/Compra Anulada - Imposible Modificar")
        GoTo 3
      End If
      '
      NPRO% = !Npro_Ocom
      NC1% = (-1) * NPRO%
      PVCL% = PIVACLI%(NC1%)
      If PVCL% = 2 Or PVCL% = 4 Or PVCL% = 5 Then
        ALIVX = 0
      End If
      '
      FECHEMI% = CVINT%(!Femi_Ocom)
      FENTREGEN% = 32767
      NRO& = NUOCO&
      CPAG% = !Cpag_Ocom
      EMBAL% = 0
      On Error Resume Next
      EMBAL% = !Embala_Ocom
      On Error GoTo 0
      FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
      If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
      ATEN$ = !AtSr_Ocom
      '
      MONEMI% = !Mone_Emis
      If MONEMI% < 1 Or ECOARCH$("TAMONED", Chr$(MONEMI%)) = "" Then MONEMI% = 1
      MONECO% = !Mone_Cos
      If MONECO% < 1 Or ECOARCH$("TAMONED", Chr$(MONECO%)) = "" Then MONECO% = 1
      '
      FOCAMOD04.Label11 = DEMONECO$(MONEMI%)
      FOCAMOD04.LText1 = NPRO%
      FOCAMOD04.Text6 = ATEN$
      FOCAMOD04.TEXT21(0) = !Cpag_Ocom
      FOCAMOD04.TEXT21(1) = EMBAL%
      FOCAMOD04.Text10 = !Lentrega
      FOCAMOD04.Text13 = Obse_Ocom
      FOCAMOD04.Text7 = ""    ' destino
      FOCAMOD04.Text18 = ""   ' sector
      FOCAMOD04.Text14 = ""   ' fecha de entrega
      FOCAMOD04.Text15 = !Mone_Cos
      FOCAMOD04.Text11 = NUOCO&
      FOCAMOD04.Text9 = FECHATEX$(FECHEMI%)
      'FOCAMOD04.PORDESCU = !Pdes_Ocom
      'FOCAMOD04.Total(4) = FORMATNUM(!Iotr1_Ocom, "F10.2")
       ' PORPER = 0
       ' On Error Resume Next
        'PORPER = 100 * Val(FOCAMOD04.Total(4)) / (!ToBru_Ocom - !IDes_Ocom)
        'On Error GoTo 0
       ' If PORPER > 99.9 Then PORPER = 99.9
      'FOCAMOD04.Text17 = TRIM$(FORMATNUM(PORPER, "F5.1"))
      FOCAMOD04.Text13 = !Obse_Ocom
      On Error Resume Next
      FOCAMOD04.Label13 = !Anex_Ocom
      FOCAMOD04.Label15 = TRIM$(Str$(!Stat_Ocom))
      '\\\OT-06-0082-WAR-06/03/06///
      NROREV% = 0
      On Error Resume Next
        NROREV% = XVALO(!OCOMREVIS)
        If !Stat_Ocom >= 0 Then 'si no esta aprobada no cambia revisión(0).
          NROREV% = NROREV% + 1
        End If
      On Error GoTo 0
      FOCAMOD04.Label16 = TRIM$(Str$(NROREV%)) 'FORMATNUM$(NROREV1%, "F2.0") LO COMENTE XQ DEVUELVE "" SI NUMERO =0
      '\\\OT-06-0082-WAR-FIN///
      On Error GoTo 0
      '
    End With
    '
    Screen.MousePointer = 11
    REDETA$ = ""
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG("!OCOMDETA", 0)
    Call SETULTREG(ARCHILIB$, 0)
    '
    REBLA$ = REGBLAN$("OCOMDETA")
    Call REPLA(REBLA$, MKS$(NUOCO&), 1, 4)
    Call REPLA(REBLA$, MKI$(FECHEMI%), 5, 2)
    Call REPLA(REBLA$, MKI$(NPRO%), 7, 2)
    Call REPLA(REBLA$, Chr$(MONECO%), 121, 1)
    Call REPLA(REBLA$, Chr$(MONEMI%), 122, 1)
      TICAMBIO = 1
      If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
    '
    SQLX$ = "SELECT * FROM DETOCABI "
    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    SQLX$ = SQLX$ + "ORDER BY NuOrd_Item ASC"
    Set COMDET = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COMDET
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "Imposible cargar Detalle de O/Compra")
        On Error GoTo 0
        GoTo 3
      End If
      On Error GoTo 0
      Do While Not .EOF
        '
        XXY$ = REBLA$
        CAPEN = !Cant_Ocom - !Cant_Rec
        If CAPEN < 0.05 Then CAPEN = 0
        PREUNID# = !PrUn_Neto * TICAMBIO
        '
        CIXI% = !Cod_inte
        Call REPLA(XXY$, MKI$(CIXI%), 9, 2)
        If CIXI% <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
        On Error Resume Next
        UNIMA$ = TRIM$(!Uni_Mas)
        On Error GoTo 0
        If UNIMA$ = "" Then UNIMA$ = TRIM$(UNIMED$(CIXI%))
        Call REPLA(XXY$, UNIMA$, 11, 4)
        '\\\OT-05-0693-WAR-24/10/05///
          If IsNull(!Sect_Ocom) Then
            SECT1$ = ""
          Else
            SECT1$ = !Sect_Ocom
          End If
          Call REPLA(XXY$, SECT1$, 15, 16)
            CCU% = !IContab
            CCU1$ = CODCUE$(CCU%)
          Call REPLA(XXY$, CCU1$, 105, 12)
        '\\\OT-05-0693-WAR-FIN///
        Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
        If IsNull(!Uni_Com) Then
          UNICOMOC$ = "  "
        Else
          UNICOMOC$ = !Uni_Com
        End If
        Call REPLA(XXY$, UNICOMOC$, 59, 4)
        If TRIM$(Mid$(XXY$, 11, 4)) = "" Then Call REPLA(XXY$, UNICOMOC$, 11, 4)
        '
        Call REPLA(XXY$, MKS$(!Coef_Unids), 63, 4)
        Call REPLA(XXY$, MKD$(!Pre_Unit), 67, 8)
        Call REPLA(XXY$, MKS$(!Desc_Item), 75, 4)
        Call REPLA(XXY$, MKD$(!PrUn_Neto), 79, 8)
        Call REPLA(XXY$, MKD$(!Cant_Ocom * !PrUn_Neto), 87, 8)
        Call REPLA(XXY$, MKI$(CVINT%(!FVig_Ocom)), 95, 2)
        FENTREPAR% = CVINT(!FVig_Ocom)
        If FENTREPAR% > 0 Then
          If FENTREPAR% < FENTREGEN% Then
            FENTREGEN% = FENTREPAR%
          End If
        End If
        '
        If IsNull(!CONSIMAT) Then
          CONSIMATOC$ = "  "
        Else
          CONSIMATOC$ = !CONSIMAT
        End If
        Call REPLA(XXY$, CONSIMATOC$, 97, 2)
        Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
          STATUI% = !Stat_Ocom: While STATUI% < 0: STATUI% = STATUI% + 256: Wend
        Call REPLA(XXY$, Chr$(STATUI%), 124, 1)
        Call REPLA(XXY$, MKS$(!NuOrd_Item), 125, 4)
        '
          NUECUEIMPO$ = CODCUE$(!IContab)
        Call REPLA(XXY$, NUECUEIMPO$, 105, 12)
          'SECTOROC$ = !Sect_Ocom '***' ESTA LINEA NO ESTABA ATAJADA
          'Call REPLA(XXY$, SECTOROC$, 15, 12)
          '
        Call REGAPP("!OCOMDETA", XXY$)
          '
          DELIBX$ = TRIM$(!Delibre)
          If DELIBX$ = "" Then DELIBX$ = DESCRIT0$(CIXI%)
        Call REGAPP(ARCHILIB$, DELIBX$)
        '
        FOCAMOD04.Text7 = CODCUE$(!IContab)
        On Error Resume Next
        FOCAMOD04.Text18 = !Sect_Ocom
        On Error GoTo 0
        '
      .MoveNext
      Loop
    End With
    '
    If FENTREGEN% = 32767 Or FENTREGEN% < FECHEMI% Then FENTREGEN% = FECHEMI%
    FOCAMOD04.Text14 = FECHATEX$(FENTREGEN%)
    '
    Call FOCAMOD04.CARCOANT
    Call FOCAMOD04.CARLISCO
    Call FOCAMOD04.CALTOCOM
    '\\\OT-06-0149-WAR-19/04/06///
    FOCAMOD04.Text122 = FORMATNUM(1, "F11.4")
    HOII% = HOY(HOS$)
    FOCAMOD04.Label21(5) = HOS$ 'COTIZACION DEL DIA DE LA FECHA
    TICAMBX# = TICAMONE(MONEMI%)
    FOCAMOD04.Text122 = FORMATNUM(TICAMBX#, "F11.4")
    '\\\OT-06-0149-WAR-FIN///
    FOCAMOD04.Show 1
    GoTo 3
    '
999 Screen.MousePointer = 1
    '
End Sub
'
Sub ACTCOSTANSER(CI0%, NOPERX%, NPRO%, PREUNINE#, MONEMI%, FECHEMI%)
   RFF$ = RECFILT$("COSOPTER", 1, MKI$(CI0%))
   For UK& = 1 To Len(RFF$) Step 4
     RECOSO& = CVS(Mid$(RFF$, UK&, 4))
     If RECOSO& > 0 Then
       If RECOSO& <= ULTREG&("COSOPTER") Then
         ZXC$ = REGLEIDO$("COSOPTER", RECOSO&)
         If CVI(Mid$(ZXC$, 3, 2)) = NOPERX% Then GoTo 20
       End If
     End If
   Next UK&
   '
   RECOSO& = 1 + ULTREG&("COSOPTER")
   ZXC$ = String$(32, 0)
   Call REPLA(ZXC$, MKI$(CI0%), 1, 2)
   Call REPLA(ZXC$, MKI$(NOPERX%), 3, 2)
   '
20 NPROX% = CVI(Mid$(ZXC$, 5, 2))
   If NPROX% = 0 Or NPROX% = NPRO% Then
     If PREUNINE# >= 0.00005 Then
       MONEI% = MONEMI%: If MONEI% < 1 Then MONEI% = 1
       Call REPLA(ZXC$, MKI$(NPRO%), 5, 2)
       Call REPLA(ZXC$, MKD$(PREUNINE#), 7, 8)
       Call REPLA(ZXC$, MKI$(MONEMI%), 15, 2)
       Call REPLA(ZXC$, MKI$(FECHEMI%), 17, 2)
       Call GRAREG("COSOPTER", ZXC$, RECOSO&)
       Call CIERRARCH("COSOPTER")
     End If
   End If
   '
End Sub

Sub OCOMANU()
    '
    If COMALTER%("Realmente Desea\ANULAR Ordenes de Compra Abiertas?\\No, Abandonar\Si, Anular O/C's") <> 2 Then
      Exit Sub
    End If
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE COMPRA"
    '
    'Cargo en el list las O/C que puedan ser Anuladas.
5   Screen.MousePointer = 11
    Call ABRECOMPRAS
    SQLX$ = " SELECT  DISTINCT ENCOCABI.Nume_Ocom, ENCOCABI.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM ENCOCABI "
    SQLX$ = SQLX$ + " INNER JOIN  DETOCABI "
    SQLX$ = SQLX$ + " ON (ENCOCABI.Nume_Ocom = DETOCABI.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE DETOCABI.Stat_Ocom <8 " ' Excluyo las que estan anuladas O COMPLIDAS
    'SQLX$ = SQLX$ + " AND DETOCABI.Cant_Rec < DETOCABI.Cant_Ocom "
    SQLX$ = SQLX$ + " AND ENCOCABI.Stat_Ocom < 1 "
    SQLX$ = SQLX$ + " ORDER BY ENCOCABI.Nume_Ocom  ASC "
    Set OCANULTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCANULTEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No hay Ordenes de Compra Abiertas\Pendientes que Puedan Anularse.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        J& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
            J& = J& + 1
            Call GRAREG("!INDIOCPE", XXY$, J&)
            .MoveNext
        Loop
    End With
    '
    Call SETULTREG("!INDIOCPE", J&)
    Call CIERRARCH("!INDIOCPE")
    Call BAJALDISCO
    Call FRESHECHO("!INDIOCPE")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra Abiertas\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    Call FRESHECHO("!INDIOCPE")
10  YY$ = REGSEL$("!INDIOCPE/Indice de Ordenes de Compra Abiertas Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    'nuevo marcado de anulacion
    With EncoCabi
      ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Stat_Ocom < 8"
      RFXX$ = "(anul) "
      .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 9                          'cambio status
        !Refe_OcoM = RFXX$ + !Refe_OcoM         'cambio referen
        .Update
      End If
    End With
    '
    YAPRIREC% = 0
    With DetoCabi
      ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Stat_Ocom < 8"
20    If YAPRIREC% = 0 Then
          .FindFirst ABUS$
          YAPRIREC% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 9
        .Update
        GoTo 20
      End If
    End With
'Marco como cumplidas las autorizaciones de entrega
    With AutoEnt
     ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Sald_Aten > 0 "
     ABUS$ = ABUS$ + " AND Stat_Aten <> 9  "
     .FindFirst ABUS$
30   If .NoMatch = False Then
        .Edit
        !Stat_Aten = 9
        !Sald_Aten = 0
        .Update
        .FindNext ABUS$
        GoTo 30
      End If
    End With
    'EncoCabi.Close
    'DetoCabi.Close
    '
    Screen.MousePointer = 11
    '
    J& = 0
    For I& = 1 To ULTREG&("!INDIOCPE")
      XX$ = REGLEIDO$("!INDIOCPE", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("!INDIOCPE", XX$, J&)
      End If
    Next I&
    Call SETULTREG("!INDIOCPE", J&)
    '
    ' borrar de lista de materiales
    NOCOBUS$ = TRIM$(Str$(NUOCO&))
    While Len(NOCOBUS$) < 6: NOCOBUS$ = "0" + NOCOBUS$: Wend
    NOCOBUS$ = "OC-" + Left$(NOCOBUS$, 6) + "-"

    Call ABREORTRA
    ABUS$ = "Nro_Ocom LIKE '" & NOCOBUS$ & "*'"
    With ListaMat
40    .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Nro_Ocom = ""
        .Update
        GoTo 40
      End If
    End With

    ' CANCELAR RESERVAS
    '
'    JJ& = 0
'    For J& = 1 To ULTREG&("MACONSIG")
'      XX$ = REGLEIDO$("MACONSIG", J&)
'      If Val(Mid$(XX$, 24, 6)) <> NUOCO& Then
'           JJ& = JJ& + 1
'           Call GRAREG("MACONSIG", XX$, JJ&)
'         Else
'           CAENTRE# = CVD(Mid$(XX$, 63, 8))
'           CADEVUE# = CVD(Mid$(XX$, 111, 8))
'           If Abs(CAENTRE# - CADEVUE#) >= 0.05 Then
'             JJ& = JJ& + 1
'             Call GRAREG("MACONSIG", XX$, JJ&)
'           End If
'      End If
'    Next J&
'    Call SETULTREG("MACONSIG", JJ&)
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FRESHECHO("!INDIOCOM")
    Screen.MousePointer = 1
    GoTo 5
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CONVERCOM()
   '
   ' conversion de formato de ordenes de compra
   '
   If ULTREG&("OCOMPRA") > 0 Then
     '
     Call BLOQARCH("OCOMPRA")
     Call BLOQARCH("OCOMENCA")
     Call BLOQARCH("OCOMDETA")
     '
     Screen.MousePointer = 11
     OCOA& = 0: NPROA% = 0: ENCAO$ = ""
     FIN& = ULTREG&("OCOMPRA")
     '
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       NUOCO& = CVS(Left$(XX$, 4))
       FEMI% = CVI(Mid$(XX$, 5, 2))
       NPRO% = CVI(Mid$(XX$, 7, 2))
       FENTRE% = CVI(Mid$(XX$, 9, 2))
       CI0% = CVI(Mid$(XX$, 11, 2))
       CAOR = CVS(Mid$(XX$, 13, 4))
       OBSER$ = Mid$(XX$, 17, 30)
       PREUNID# = CVS(Mid$(XX$, 49, 4))
       DESCUNI = CVS(Mid$(XX$, 53, 4))
       TISU% = Asc(Mid$(XX$, 57, 1))
       CARE = CVS(Mid$(XX$, 58, 4))
       STAT% = Asc(Mid$(XX$, 62, 1))
       '
       If NUOCO& <> OCOA& Or NPRO% <> NPROA% Then
         If OCOA& <> 0 Then
           TOTIVA# = 0
           PVCL% = PIVACLI%((-1) * NPRO%)
           If PVCL% <> 2 Then
             If PVCL% <> 4 Then
               TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
             End If
           End If
           Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
           Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
           Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 145, 8)
           Call REGAPP("OCOMENCA", ENCAO$)
         End If
         '
         OCOA& = NUOCO&: NPROA% = NPRO%
         ENCAO$ = REGBLAN$("OCOMENCA")
         '
         Call REPLA(ENCAO$, MKS$(NUOCO&), 1, 4)
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
         Call REPLA(ENCAO$, REFE$, 5, 44)
         Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
         Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
         Call REPLA(ENCAO$, Chr$(TISU%), 53, 1)
         Call REPLA(ENCAO$, Chr$(STAT%), 54, 1)
         '
         HO% = HOY(HOS$)
         HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
         Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
         Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
         Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
         '
         If STAT% = 9 Then
           Call REPLA(ENCAO$, MKI$(HO%), 67, 2)
           Call REPLA(ENCAO$, MKI$(Int(HP)), 69, 2)
           Call REPLA(ENCAO$, MKI$(USNBR%), 71, 2)
           Call REPLA(ENCAO$, "Conversion BD", 73, 16)
         End If
         '
         CPAG% = CPAGCLI%((-1) * NPRO%)
         Call REPLA(ENCAO$, MKI$(CPAG%), 153, 2)
         REDETA& = 1 + ULTREG&("OCOMDETA")
         Call REPLA(ENCAO$, MKS$(REDETA&), 155, 4)
         Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
         '
         Call REPLA(ENCAO$, OBSER$, 193, 64)
         '
       End If
       '
       REDET$ = REGBLAN$("OCOMDETA")
       Call REPLA(REDET$, MKS$(NUOCO&), 1, 4)
       Call REPLA(REDET$, MKI$(FEMI%), 5, 2)
       Call REPLA(REDET$, MKI$(NPRO%), 7, 2)
       Call REPLA(REDET$, MKI$(CI0%), 9, 2)
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(REDET$, UNIMAS$, 11, 4)
       Call REPLA(REDET$, MKS$(CAOR), 55, 4)
       Call REPLA(REDET$, UNIMED$(CI0%), 59, 4)
       Call REPLA(REDET$, MKS$(1), 63, 4)
       Call REPLA(REDET$, MKD$(PREUNID#), 67, 8)
       Call REPLA(REDET$, MKD$(PREUNID#), 79, 8)
       Call REPLA(REDET$, MKD$(PREUNID# * CAOR), 87, 8)
       Call REPLA(REDET$, MKI$(FENTRE%), 95, 2)
       If TISU% > 1 Then
         Call REPLA(REDET$, " *", 97, 2)
       End If
       Call REPLA(REDET$, MKS$(CARE), 99, 4)
       Call REPLA(REDET$, Chr$(STAT%), 124, 1)
       RENCA& = 1 + ULTREG&("OCOMENCA")
       Call REPLA(REDET$, MKS$(RENCA&), 125, 4)
       Call REGAPP("OCOMDETA", REDET$)
       '
       TOTANE# = TOTANE# + PREUNID# * CAOR
       REDETA& = ULTREG&("OCOMDETA")
       Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
       '
       Call GRAREG("OCOMPRA", String$(128, 0), U&)
       '
     Next U&
     '
     If OCOA& <> 0 Then
       TOTIVA# = 0
       PVCL% = PIVACLI%((-1) * NPRO%)
       If PVCL% <> 2 Then
         If PVCL% <> 4 Then
           TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
         End If
       End If
       Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
       Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
       Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 137, 8)
       Call REGAPP("OCOMENCA", ENCAO$)
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
   End If
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Sub SELCOMRAFE()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   REBLA$ = REGBLAN$("COPRECOM")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     FECHEMI% = CVI(Mid$(XXX$, 5, 2))
     '
     STAT% = Asc(Mid$(XXX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       If FECHEMI% >= DES% Then
         If FECHEMI% <= HAS% Then
           '
           NPRO% = CVI(Mid$(XXX$, 7, 2))
           CIXI% = CVI(Mid$(XXX$, 9, 2))
           DOPE$ = ""
           NOPE% = CVI(Mid$(XXX$, 103, 2))
           If NOPE% > 0 Then
             DOPE$ = DESCRIOP$(CIXI%, NOPE%)
           End If
           NUOC& = CVS(Left$(XXX$, 4))
           UMEDI$ = Mid$(XXX$, 59, 4)
           If TRIM$(UMEDI$) = "" Then
             UMEDI$ = UNIMED$(CIXI%)
           End If
           CACOM = CVS(Mid$(XXX$, 55, 4))
           '
           If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
             CACOM = CVS(Mid$(XXX$, 99, 4))  ' LA CANT COMPRADA = A LA RECIBIDA
           End If
           '
           PREUNI# = CVD(Mid$(XXX$, 79, 8))
           '
           MONEMI% = Asc(Mid$(XXX$, 122, 1))
           If MONEMI% <= 1 Then MONEX$ = "    "
           If MONEMI% = 2 Then MONEX$ = "U$S "
           If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
           TICAMBIO = 1
           If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
           '
           YYY$ = REBLA$
           Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
           Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
           Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
           Call REPLA(YYY$, DOPE$, 11, 22)
           Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
           Call REPLA(YYY$, UMEDI$, 37, 4)
           Call REPLA(YYY$, MKS$(CACOM), 41, 4)
           Call REPLA(YYY$, MONEX$, 49, 4)
           Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
           ITOCOM# = PREUNI# * CACOM
           Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
           Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
           Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
           JJ& = JJ& + 1
           Call GRAREG("COPRECOM", YYY$, JJ&)
         End If
       End If
     End If
     '
   Next U&
   '
   Call SETULTREG("COPRECOM", JJ&)
   Call CIERRARCH("COPRECOM")
   Screen.MousePointer = 1
   '
End Sub
'
Sub DESANULOC()
   '
'3  Call COPYSTRU("INDIOCOM", "IOCANUL")
'   J& = 0: K& = 0: OCOA& = 0
'   '
'   FIN& = ULTREG&("OCOMENCA")
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMENCA", U&)
'      If Asc(Mid$(X$, 54, 1)) = 9 Then ' esta anulada
'5       Y$ = Left$(X$, 48) + MKS$(U&)
'        J& = J& + 1
'        Call GRAREG("IOCANUL", Y$, J&)
'      End If
'9   Next U&


    'CARGO LA LISTA DE O/C QUE SE PUEDEN RESTAURAR

3   Call COPYSTRU("INDIOCOM", "IOCANUL")
    Screen.MousePointer = 11
    Call ABRECOMPRAS
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom = 9 " ' Si esta anulada
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set OCRESTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCRESTEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No  hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      J& = 0
      Do While (.EOF = False)
                XXY$ = REGBLAN$("IOCANUL")
                Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
                Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
                J& = J& + 1
                Call GRAREG("IOCANUL", XXY$, J&)
                .MoveNext
      Loop
    End With
    Call SETULTREG("IOCANUL", J&)
    Call CIERRARCH("IOCANUL")
    Call BAJALDISCO
    Call FRESHECHO("IOCANUL")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
     Call MENSERR(24, "No hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("IOCANUL/Indice de Ordenes de Compra Anuladas para Restauración")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
     
    'nuevo marcado de anulacion
     
     With OcomEnca
      ABUS$ = "Nume_Ocom = " & NUOCO& & "AND Stat_Ocom = 9"
      .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 0                      'cambio status
        FECHEMII$ = Format(!Femi_Ocom, "dd/mm/yy")
        PRORAI$ = !RaSo_Prov
        REFEI$ = FECHEMII$ & " - " & PRORAI$
        !Refe_OcoM = REFEI$    'cambio referencia
        .Update
      End If
    End With
    '
    YAPRIREC% = 0
    With OcomDeta
      ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Stat_Ocom = 9"
20    If YAPRIREC% = 0 Then
          .FindFirst ABUS$
          YAPRIREC% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Edit
        STATUIX% = 0
        If !Cant_Rec > 0 Then STATUIX% = 2
        If !Cant_Rec >= !Cant_Ocom - 0.05 Then STATUIX% = 3
        !Stat_Ocom = STATUIX%
        .Update
        GoTo 20
      End If
    End With
   OcomEnca.Close
   OcomDeta.Close
    
     '''
'    RENCA& = CVS(Mid$(YY$, 49, 4))
'    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
'      Call MENSERR(24, "Error de Indices - Consulte")
'      GoTo 999
'    End If
'    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
'    If CVS(Left$(XX$, 4)) <> NUOCO& Then
'      Call MENSERR(24, "Error de Indices - Consulte")
'      GoTo 999
'    End If
    '
    Screen.MousePointer = 11
'    FIN& = ULTREG&("OCOMENCA")
'    For J& = 1 To FIN&
'      Call TRACE(20, J&, FIN&)
'      XX$ = REGLEIDO$("OCOMENCA", J&)
'      If CVS(Left$(XX$, 4)) = nuoco& Then
'          FECHEMII% = CVI(Mid$(XX$, 49, 2))
'          NPROI% = CVI(Mid$(XX$, 51, 2))
'          FEX = FECHEMII%
'          REFEI$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPROI%)
'          Call REPLA(XX$, Chr$(0), 54, 1)
'          Call REPLA(XX$, REFEI$, 5, 44)
'          Call GRAREG("OCOMENCA", XX$, J&)
'          YADESA% = 1
'        Else
'          If YADESA% > 0 Then GoTo 19
'      End If
'    Next J&
'19  Call CIERRARCH("OCOMENCA")
'    '
'    FIN& = ULTREG&("OCOMDETA")
'    YADESA% = 0
'    For J& = FIN& To 1 Step -1
'      Call TRACE(20, FIN& - J& + 1, FIN&)
'      XX$ = REGLEIDO$("OCOMDETA", J&)
'      If CVS(Left$(XX$, 4)) = nuoco& Then
'        If Asc(Mid$(XX$, 124, 1)) = 9 Then
'          CAPEDI1 = CVS(Mid$(XX$, 55, 4))
'          CARECI1 = CVS(Mid$(XX$, 99, 4))
'          STATUIX% = 0
'          If CARECI1 > 0 Then STATUIX% = 2
'          If CARECI1 >= CAPEDI1 - 0.05 Then STATUIX% = 3
'          Call REPLA(XX$, Chr$(STATUIX%), 124, 1)
'          Call GRAREG("OCOMDETA", XX$, J&)
'        End If
'      End If
'    Next J&
'    Call CIERRARCH("OCOMDETA")
    '
    J& = 0
    For I& = 1 To ULTREG&("IOCANUL")
      XX$ = REGLEIDO$("IOCANUL", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("IOCANUL", XX$, J&)
      End If
    Next I&
    Call SETULTREG("IOCANUL", J&)
    '
    GoTo 3
    '
999 Call CIERRARCH("OCOMENCA")
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
End Sub

Sub ITEMCUMPLI()
    '
    TITU$ = "DAR POR CUMPLIDOS ITEMS DE COMPRAS"
    '
2   Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCUM")
    Call ABRECOMPRAS
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < ( OCOMDETA.Cant_Ocom - 0.05 ) "
    SQLX$ = SQLX$ + " AND OCOMDETA.Cod_Inte > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom DESC "
    Set OCCUMTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCCUMTEMP
      On Error Resume Next
      If Err Then
        Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      .MoveFirst
      J& = 0
      Do While (.EOF = False)
        XXY$ = REGBLAN$("INDIOCUM")
        Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
        Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
        J& = J& + 1
        Call GRAREG("INDIOCUM", XXY$, J&)
      .MoveNext
      Loop
    End With
    '
'    OCOA& = 0: VOCO$ = ""
'    FIN& = ULTREG&("OCOMDETA")
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMDETA", U&)
'      NOCO& = CVS(Left$(X$, 4))
'      If NOCO& <> OCOA& Then
'        If Asc(Mid$(X$, 124, 1)) < 8 Then ' esta abierta y pendiente
'          If CVS(Mid$(X$, 99, 4)) < CVS(Mid$(X$, 55, 4)) - 0.05 Then
'            VOCO$ = VOCO$ + MKS$(NOCO&)
'            OCOA& = NOCO&
'          End If
'        End If
'      End If
'    Next U&
'
'3   J& = 0: K& = 0: OCOA& = 0
'    '
'    FIN& = ULTREG&("OCOMENCA")
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMENCA", U&)
'      If Asc(Mid$(X$, 54, 1)) < 9 Then ' esta abierta y pendiente
'        NOCO& = CVS(Left$(X$, 4))
'        For KKL& = 1 To Len(VOCO$) Step 4
'          If CVS(Mid$(VOCO$, KKL&, 4)) = NOCO& Then GoTo 5
'        Next KKL&
'        GoTo 9
'        '
'5       Y$ = Left$(X$, 48) + MKS$(U&)
'        J& = J& + 1
'        Call GRAREG("INDIOCUM", Y$, J&)
'      End If
'9   Next U&
    Call SETULTREG("INDIOCUM", J&)
    Call CIERRARCH("INDIOCUM")
    Call FRESHECHO("INDIOCUM")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("INDIOCUM/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    DAPORCU04.Label14 = NUOCO&
    '
    DAPORCU04.Show 1
    '
    ' A PARTIR DE AQUI MARCAR REGISTROS QUE SE DAN POR CUMPLIDOS
    Call ABRECOMPRAS
    If Val(DAPORCU04.Label14) = NUOCO& Then
      Screen.MousePointer = 11
      'URE& = ULTREG&("OCOMDETA")
      For U& = 1 To DAPORCU04.LICOPEN.ListCount
        If DAPORCU04.LICOPEN.Selected(U& - 1) = True Then
          TTXX$ = DAPORCU04.LICOPEN.List(U& - 1)
          NUORIT% = Val(TRIM$(Mid$(TTXX$, 121, 8)))
          '
          SQLX$ = "SELECT * FROM OCOMDETA "
          SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
          SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
          SQLX$ = SQLX$ + "AND NuOrd_Item = " & NUORIT% & " "
          '
          Set MARCTEMP = Compras.OpenRecordset(SQLX$, dbOpenDynaset)
          '
          With MARCTEMP
            .MoveFirst
            Do While (.EOF = False)
              COEFCAN = !Coef_Unids
              If COEFCAN < 0.005 Then COEFCAN = 1
               CAORI = !Cant_Ocom * COEFCAN
               CAREC = !Cant_Rec
               CAPEN = CAORI - CAREC
              If CAPEN >= 0.005 Then
               .Edit
               !Stat_Ocom = 8
               .Update
              End If
              .MoveNext
            Loop
          End With
         End If
    Next U&
End If
'          If REOC& > 0 Then
'            If REOC& <= URE& Then
'              X$ = REGLEIDO$("OCOMDETA", REOC&)
'              STAT% = Asc(Mid$(X$, 124, 1))
'              If STAT% < 9 Then   ' NO ESTA ANULADA
'                COEFCAN = CVS(Mid$(X$, 63, 4))
'                  If COEFCAN < 0.005 Then COEFCAN = 1
'                CAORI = CVS(Mid$(X$, 55, 4)) * COEFCAN
'                CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
'                CAPEN = CAORI - CAREC
'                If CAPEN >= 0.005 Then
'                  Call REPLA(X$, Chr$(8), 124, 1)
'                  Call GRAREG("OCOMDETA", X$, REOC&)
'                End If
'              End If
'            End If
'          End If
'        End If
'     Next U&
'   End If
    Unload DAPORCU04
    Screen.MousePointer = 1
    '
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub

Sub DEVOLCONSI()
        '
1105    Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 1
        If NPROVE% > 0 Then
             EBOL$ = MKI$(NPROVE%)
          Else
             EBOL$ = OBJPLA$("SELPROVE", "")
             If EBOL$ = "" Then
                GoTo 9999
             End If
        End If
        '
        NC% = CVI(EBOL$)
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            NPROVE% = 0
            GoTo 1105
        End If
        '
        NC1% = (-1) * NC%
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        Screen.MousePointer = 11
        HAYOC% = 0
        Call ABRESTOCKS
        SQLX$ = "SELECT * FROM MACONSIG "
        SQLX$ = SQLX$ + " WHERE (Cant_Entreg - Cant_Rendid)>0 "
        SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
        Set MACONSITMP = Stocks.OpenRecordset(SQLX$, 4)
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
          Do While .EOF = False
            CAPENE# = !Cant_Entreg - !Cant_Rendid
            NUOC& = Val(Mid$(!Nro_Ocom, 4, 6))
            CI% = !Cod_inte
            '
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, UNIMED$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
            Call REPLA(Y$, MKS$(NUOC&), 151, 4)
            Call REPLA(Y$, MKS$(I&), 189, 4)
            HAYOC% = 1
            Call REGAPP("!BOLREDET", Y$)
            .MoveNext
          Loop
          End If
        End With
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Rendición Materiales Consignados - " + TRIM$(RASOCLI$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
        '
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            LOTE$ = Mid$(YY$, 69, 12)
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                If CAN# > 0.005 Then
                  NUOC& = CVS(Mid$(YY$, 151, 4))
                  DOSEC$ = ""
                  With MaConsig
                    SQLX$ = " Val(Mid$(Nro_Ocom, 4, 6)) =  " & NUOC & " "
                    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & ""
                    SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
                    .FindFirst SQLX$
                    If .NoMatch = False Then
                      .Edit
                      DOSEC$ = !Nro_Ocom
                      TOENTRE# = !Cant_Rendid + CAN#
                      !Cant_Rendid = TOENTRE#
                      .Update
                    End If
                  End With
                End If
              End If
            End If
            '
          Next YX&
          Call CIERRARCH("MACONSIG")
          '
          Screen.MousePointer = 1
          '
1199    End If
        '
9999    Call CIERRARCH("*.*")
        '
End Sub

Sub RANKPRO()
'\\\OT-05-0381-WAR-07/06/05///
    Dim TOTCOMPRA#(32767)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTCOMPRA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "    ' NO ANULADAS
    '
    Set CABATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
        UK& = UK& + 1
        Call TRACE(20, UK&, FIN&)
        NPRO% = !Npro_Ocom
        If NPRO% > 0 Then
          If NPRO% > PROMAX% Then PROMAX% = NPRO%
          IMPNETO# = (!PrUn_Neto * !Cant_Ocom)
          TOTCOMPRA#(NPRO%) = TOTCOMPRA#(NPRO%) + (IMPNETO# * TICAMONE(!Mone_Emis))
        End If
        .MoveNext
      Loop
    End With
    '
    For I% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTCOMPRA#(I%)
    Next I%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTCOMPRA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTCOMPRA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = Val(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANPRO")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0381-WAR-FIN///
End Sub

Sub RANKART()
'\\\OT-05-0382-WAR-07/06/05
    '
    Dim TOTART#(32767)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTART#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "  ' NO ANULADAS
    '
    Set CABATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
         UK& = UK& + 1
         Call TRACE(20, UK&, FIN&)
         NART% = !Cod_inte
         IMPNETO# = (!PrUn_Neto * !Cant_Ocom)
         If NART% > 0 Then
           If NART% > PROMAX% Then PROMAX% = NART%
           TOTART#(NART%) = TOTART#(NART%) + (IMPNETO# * TICAMONE(!Mone_Emis))
         Else
           TOTART#(0) = TOTART#(0) + (IMPNETO# * TICAMONE(!Mone_Emis))
         End If
         .MoveNext
      Loop
    End With
    '
    For I% = 0 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTART#(I%)
    Next I%
    '
    List1.Clear
    For NCI% = 0 To 32766
       If Abs(TOTART#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTART#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTART#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTART#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = Val(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANART")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0382-WAR-FIN
End Sub


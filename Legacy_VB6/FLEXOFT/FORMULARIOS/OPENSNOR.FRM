VERSION 5.00
Begin VB.Form OPENSNOR 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayos Normalizados - Opciones"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5625
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
   ScaleHeight     =   4605
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   800
      Left            =   4620
      Picture         =   "OPENSNOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Informes y Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   210
      TabIndex        =   9
      Top             =   2310
      Width           =   4215
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   13
         Top             =   525
         Width           =   3795
         Begin VB.CommandButton Command64 
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
            Left            =   3150
            Picture         =   "OPENSNOR.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Pedidos Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command62 
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
            Left            =   2550
            Picture         =   "OPENSNOR.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Pedidos Pendientes por Cliente"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
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
            Picture         =   "OPENSNOR.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Encabezados de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command63 
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
            Left            =   1950
            Picture         =   "OPENSNOR.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Detalle por Número de Pedido / Listado Pedidos sin Facturar"
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   210
         TabIndex        =   10
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option1 
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
            Left            =   525
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
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
            Left            =   2205
            TabIndex        =   11
            Top             =   270
            Width           =   1170
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4620
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   22
      Top             =   3360
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ensayos Normalizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   2
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   20
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command5 
            Height          =   600
            Left            =   0
            Picture         =   "OPENSNOR.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ensayos Normalizados"
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
            TabIndex        =   21
            Top             =   105
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   18
         Top             =   480
         Width           =   2430
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
            Picture         =   "OPENSNOR.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Alta de Nuevo Item de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipos de Ensayos"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1530
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
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
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
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
      Height          =   800
      Left            =   4620
      Picture         =   "OPENSNOR.frx":12FE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1155
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   4300
      X2              =   5650
      Y1              =   4460
      Y2              =   4460
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4050
      Picture         =   "OPENSNOR.frx":1608
      Top             =   3960
      Width           =   2010
   End
End
Attribute VB_Name = "OPENSNOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    End
End Sub


Private Sub Command2_Click()
'Sub COPEFAC()
   '
   RECU% = 0: If HOY(HOS$) <= FECHANUM("24/10/01") Then RECU% = 1
   HOII% = HOY(HOS$)
   JJ& = 0
   FIN& = ULTREG("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETA", U&)
     NC% = CVI(Mid$(XX$, 7, 2))
     If NC% > 0 Then
       If Asc(Mid$(XX$, 27, 1)) <> 9 Then ' no esta anulado
         CI% = CVI(Mid$(XX$, 33, 2))
         If CI% > 0 And CI% <= NUMAS% Then
           FF% = CVI(Mid$(XX$, 5, 2))
           NF& = CVS(Left$(XX$, 4))
           DOCU$ = "PC-" + TRIM$(Str$(NF&))
           CAPED = CVS(Mid$(XX$, 79, 4))
           SALPE = CAPED - CVS(Mid$(XX$, 95, 4))
           '
           If RECU% = 1 Then
             If NF& = 77 Then
               If CI% = CODINT%("065008T-0000") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1681
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 85 Then
               If CI% = CODINT%("042261T-PHH0") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1685
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 21 Then
               If CI% = CODINT%("44132/2K") Then
                 TOENTREGA = 5
                 SALPENDIEN = 0
                 REMI = 1628
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
             If NF& = 81 Then
               If CI% = CODINT%("R410104-PVH0") Then
                 TOENTREGA = 1
                 SALPENDIEN = 0
                 REMI = 1652
                 Call REPLA(XX$, MKS$(TOENTREGA), 95, 4)
                 Call REPLA(XX$, MKS$(SALPENDIEN), 99, 4)
                 Call REPLA(XX$, MKI$(HOII%), 103, 2)
                 Call REPLA(XX$, MKS$(REMI), 105, 4)
                 Call GRAREG("PEDDETA", XX$, U&)
               End If
             End If
             '
           End If
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(SALPE), 19, 4)
           Call REPLA(YY$, MKS$(CAPED), 27, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
       End If
     End If
   Next U&
   '
   FIN& = ULTREG("MOVISTO")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MOVISTO", U&)
      DOPRI$ = Mid$(XX$, 23, 10): DOSEC$ = Mid$(XX$, 33, 10)
      CMX1$ = Left$(DOPRI$, 2): CMX2$ = Left$(DOSEC$, 2)
      If CMX1$ = "FC" Or CMX2$ = "FC" Or CMX1$ = "RT" Or CMX2$ = "RT" Then
         NC% = CVI(Mid$(XX$, 84, 2))
         If NC% > 0 Then
           CI% = CVI(Mid$(XX$, 3, 2))
           FF% = CVI(Mid$(XX$, 1, 2))
           DOCU$ = DOPRI$
           CAENT = CVD(Mid$(XX$, 96, 8))
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(CAENT), 31, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
      End If
   Next U&
   '
   If EXISTE%(LUDAT$ + "MOVIREPA.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "MOVIREPA")
   End If
   '
   FIN& = ULTREG("MOVIREPA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MOVIREPA", U&)
      DOPRI$ = Mid$(XX$, 23, 10): DOSEC$ = Mid$(XX$, 33, 10)
      CMX1$ = Left$(DOPRI$, 2): CMX2$ = Left$(DOSEC$, 2)
      If CMX1$ = "FC" Or CMX2$ = "FC" Or CMX1$ = "RT" Or CMX2$ = "RT" Then
         NC% = CVI(Mid$(XX$, 84, 2))
         If NC% > 0 Then
           CI% = CVI(Mid$(XX$, 3, 2))
           FF% = CVI(Mid$(XX$, 1, 2))
           DOCU$ = DOPRI$
           CAENT = CVD(Mid$(XX$, 96, 8))
           '
           YY$ = REGBLAN$("COPEFAC")
           Call REPLA(YY$, MKI$(NC%), 1, 2)
           Call REPLA(YY$, MKI$(CI%), 3, 2)
           Call REPLA(YY$, MKI$(FF%), 5, 2)
           Call REPLA(YY$, DOCU$, 7, 12)
           Call REPLA(YY$, MKS$(CAENT), 31, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("COPEFAC", YY$, JJ&)
           '
         End If
      End If
   Next U&
   '
   Call SETULTREG("COPEFAC", JJ&)
   Call CIERRARCH("COPEFAC")
   Call CIERRARCH("*.*")
   Call FINAL("*SOPEFAC")
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call ACCDIR("PEDIDO02")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    Call TRALOCAL("TABTIEN", "")
10
    VTB% = VENTABU("TABTIEN")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TABTIEN")
        X$ = REGLEIDO$("!TABTIEN", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        Call REPLA(X$, CD$, 1, 2)
        Call GRAREG("!TABTIEN", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo
        If Len(CD$) < 0 Then
          Call MENSERR(24, "Código de Tipo de Ensayo Nulo")
          GoTo 10
        End If
        ' Valida que posea descripcion
        DESC$ = TRIM$(Mid$(X$, 3, 18))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Codigo de Ensayo sin Descripción")
          GoTo 10
        End If
        ' Valida que el codigo no se repita
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TABTIEN", U2&)
          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
            Call MENSERR(24, "Código de Ensayo Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TABTIEN", "")
    End If
    '
99  Command4.Enabled = True
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Revertir Anulación de Pedido Itemizado")
       Exit Sub
    End If
    '
    If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call DESANUPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Command43.Enabled = True
End Sub

Private Sub Command45_Click()
    Command45.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Anulación de Pedido Itemizado")
       GoTo 99
    End If
    '
    If DERACCE%("ANULPEDI", "Anular Pedidos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call BLOQARCH("ODESEMI")
    Call ANULAPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Call LIBARCH("ODESEMI")
99  Command45.Enabled = True
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call MENSERR(24, "Imposible Modificar Pedido\para Formato Itemizado.\ \Anule y Cargue un Nuevo Pedido.")
       GoTo 99
    End If
    '
    If DERACCE%("MODIPEDI", "Modificar Pedidos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    Call COPYSTRU("PEDENCA", "PEDICLIE")
    Call BLOQARCH("PEDICLIE")
    Call CAMBIAPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
99  Command46.Enabled = True
End Sub

Sub CAMBIAPED()
     '
2100 If YAINDI% = 0 Then
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        RFF$ = RECFILT$("PEDDETA", 7, Chr$(1)) ' SOLO PENDIENTES DE ENTREGA Y NO DESPACHADOS
        For UU& = 1 To Len(RFF$) Step 4
          REPRESU& = CVS(Mid$(RFF$, UU&, 4))
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            '
            If Asc(Mid$(X$, 27, 1)) < 3 Then
              If CAPENDENT < 0.05 Then
                Call REPLA(X$, Chr$(3), 27, 1)
                Call GRAREG("PEDDETA", X$, REPRESU&)
              End If
            End If
            '
            If CAPENDENT >= 0.05 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
          '
        Next UU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NUPEDIDO& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       UU& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("PEDDETA", UU&)
       If CVS(Mid$(XX$, 75, 4)) >= 0.05 Then
         Call MENSERR(24, "Imposible Modificar.\Hay Ordenes de Despacho en Curso.")
         GoTo 2110
       End If
     Next UKKK&
     '
     RFQ$ = RECFILT$("DESPAPED", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       UU& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("DESPAPED", UU&)
       If Asc(Mid$(XX$, 27, 1)) < 9 Then
         If CVS(Mid$(XX$, 95, 4)) > 0 Then
           Call MENSERR(24, "Imposible Modificar.\Ya se Despachó Material.")
           GoTo 2110
         End If
       End If
     Next UKKK&
     '
     RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       REDE& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("PEDDETA", REDE&)
       If Asc(Mid$(XX$, 27, 1)) <> 9 Then ' no esta anulado
         CADESPACH = CVS(Mid$(XX$, 95, 4))
         If CADESPACH >= 0.05 Then
           Call MENSERR(24, "Imposible Modificar este Pedido.\  \Ya hay Entregas Parciales por Remito.")
           GoTo 2120
         End If
       End If
     Next UKKK&
     '
     'VERIFICAR CONSISTENCIA
     '
2120 If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
        Call MENSERR(24, "Imposible Modificar este Pedido")
        GoTo 2110
     End If
     '
     Screen.MousePointer = 11
     PDC$ = REGLEIDO$("PEDENCA", REGENCA&)
     MODPEDCLI.Label9.Caption = TRIM$(Str$(REGENCA&))
     FEX = CVI(Mid$(PDC$, 53, 2))
     MODPEDCLI.Text9.TEXT = FECHATEX$(FEX)
     MODPEDCLI.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
     NCLIE% = CVI(Mid$(PDC$, 55, 2))
     MODPEDCLI.Text1.TEXT = TRIM$(Str$(NCLIE%))
     '
     NC% = NCLIE%
     MODPEDCLI.Text2.TEXT = RASOCLI$(NC%)
     MODPEDCLI.Text3.TEXT = DOMICLI$(NC%)
     MODPEDCLI.Text4.TEXT = CPOSCLI$(NC%)
     MODPEDCLI.Text5.TEXT = LOCACLI$(NC%)
     MODPEDCLI.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
     MODPEDCLI.Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
     CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
     MODPEDCLI.Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then MODPEDCLI.Text8.ForeColor = RGB(255, 0, 0)
     MODPEDCLI.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
     '
     MODPEDCLI.Text10.TEXT = Mid$(PDC$, 57, 16)
     MODPEDCLI.Text21(0).TEXT = TRIM$(Str$(Asc(Mid$(PDC$, 268, 1))))
     ' If Val(MODPEDCLI.TEXT21(0).TEXT) < 1 Then MODPEDCLI.TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
     ' deshabilitado para permitir precios concertados fijos
     MODPEDCLI.Text21(1).TEXT = TRIM$(Str$(Asc(Mid$(PDC$, 269, 1))))
     If Val(MODPEDCLI.Text21(1).TEXT) < 1 Then MODPEDCLI.Text21(1).TEXT = Str$(CPAGCLI%(NC%))
     MODPEDCLI.Text21(2).TEXT = TRIM$(Str$(Asc(Mid$(PDC$, 270, 1))))
     If Val(MODPEDCLI.Text21(2).TEXT) < 1 Then MODPEDCLI.Text21(2).TEXT = Str$(VENDCLI%(NC%))
     '
     If Val(MODPEDCLI.Text21(2).TEXT) < 1 Then
        MODPEDCLI.Eco21(2).TEXT = "GERENCIA"
     End If
     '
     MODPEDCLI.PORDESCU.TEXT = DESCUCLI$(NC%)
     If Mid$(PDC$, 275, 14) <> String$(14, 0) Then
        MODPEDCLI.PORDESCU.TEXT = TRIM$(Mid$(PDC$, 275, 14))
     End If
     '
     LENTREG$ = Mid$(PDC$, 73, 96)
     MODPEDCLI.LENTREGA(0) = Left$(LENTREG$, 32)
     MODPEDCLI.LENTREGA(1) = Mid$(LENTREG$, 33, 32)
     MODPEDCLI.LENTREGA(2) = Mid$(LENTREG$, 65, 32)
     '
     MODPEDCLI.Text13.TEXT = TRIM$(Mid$(PDC$, 169, 32) + Mid$(PDC$, 201, 32) + Mid$(PDC$, 233, 32))
     MODPEDCLI.Label11.ForeColor = RGB(0, 0, 255)
     MODPEDCLI.Label11 = "Venta"
     '
     Call SETULTREG("!CARDETPE", 0)
     UREPRE& = ULTREG&("PEDDETA")
     JJ& = 0
     RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       UU& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("PEDDETA", UU&)
       If CVS(Left$(XX$, 4)) = NUPEDIDO& Then
         If Asc(Mid$(XX$, 27, 1)) < 9 Then
            CAN = CVS(Mid$(XX$, 79, 4))
            CADESPACH = CVS(Mid$(XX$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              '
              MONEMA% = Asc(Mid$(XX$, 52, 1))
              If JJ& < 1 Then
                MONEMI% = Asc(Mid$(XX$, 52, 1))
                If MONEMI% < 1 Then MONEMI% = 1
                If MONEMI% = 2 Then MODPEDCLI.MONEMIS = "Dólares" Else MODPEDCLI.MONEMIS = "Pesos"
              End If
              COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
              '
              ZZ$ = REGBLAN$("!CARDETPE")
              CIXI% = CVI(Mid$(XX$, 33, 2))
              Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
              PRELI = CVS(Mid(XX$, 71, 4))
              Call REPLA(ZZ$, MKS$(PRELI * COEFCON), 39, 4)
              PREUNI = CVS(Mid$(XX$, 83, 4))
              Call REPLA(ZZ$, MKS$(PREUNI * COEFCON), 65, 4)
              '
              ' verificar que esto no arme despelote en DOSIVAC
              If PREUNI > PRELI Then
                  Call REPLA(ZZ$, MKS$(PREUNI), 39, 4)
                Else
                  If PRELI > 0.005 Then
                    DESCUIN = 100 * (PRELI - PREUNI) / PRELI
                    Call REPLA(ZZ$, MKS$(DESCUIN), 69, 4)
                  End If
              End If
              '
              CAPE = CAPENDENT
              Call REPLA(ZZ$, MKS$(CAPE), 47, 4)   ' CANTIDAD
                'DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
                'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
              FENTRESO% = CVI(Mid$(XX$, 25, 2))
              Call REPLA(ZZ$, MKI$(FENTRESO%), 63, 2)
              CAVERX$ = Mid$(XX$, 113, 10)
              Call REPLA(ZZ$, CAVERX$, 113, 10)
              Call REPLA(ZZ$, MKS$(0), 125, 4)
              JJ& = JJ& + 1
              Call GRAREG("!CARDETPE", ZZ$, JJ&)
              PORCOMIPE = CVI(Mid$(XX$, 31, 2)) / 10
              If PORCOMIPE < 0 Then PORCOMIPE = 0
              If PORCOMIPE > 99.9 Then PORCOMIPE = 99.9
              MODPEDCLI.Text44.TEXT = CSTRING$(MKS$(PORCOMIPE), 3, 4, 1, 2)
              '
              If UCase$(Mid$(XX$, 51, 1)) = "R" Then
                MODPEDCLI.Label11.ForeColor = RGB(255, 0, 0)
                MODPEDCLI.Label11 = "Reparación"
              End If
              '
            End If
         End If
       End If
     Next UKKK&
     '
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("!CARDETPE")
     '
     RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
     URN& = ULTREG&("NOTACLI")
     MODPEDCLI.Text6.TEXT = ""
     '
     For U& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, U&, 4))
       If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
           TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
           If MODPEDCLI.Text6.TEXT <> "" Then
             MODPEDCLI.Text6.TEXT = MODPEDCLI.Text6.TEXT + Chr$(13) + Chr$(10)
           End If
           MODPEDCLI.Text6.TEXT = MODPEDCLI.Text6.TEXT + RTrim$(TTYY$)
         End If
       End If
     Next U&
     MODPEDCLI.MARCONOTA.Visible = True
     On Error Resume Next
     MODPEDCLI.Command6.SetFocus
     On Error GoTo 0
     Screen.MousePointer = 1
     '
     MODPEDCLI.Show 1
     GoTo 2110
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'

Sub ANULAPED()
     '
2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        RFF$ = RECFILT$("PEDDETA", 7, Chr$(1)) ' SOLO PENDIENTES DE ENTREGA Y NO DESPACHADOS
        For UU& = 1 To Len(RFF$) Step 4
          REPRESU& = CVS(Mid$(RFF$, UU&, 4))
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            '
            If Asc(Mid$(X$, 27, 1)) < 3 Then
              If CAPENDENT < 0.05 Then
                Call REPLA(X$, Chr$(3), 27, 1)
                Call GRAREG("PEDDETA", X$, REPRESU&)
              End If
            End If
            '
            If CAPENDENT >= 0.05 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
          '
        Next UU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        Call FRESHECHO("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NUPEDIDO& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       UU& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("PEDDETA", UU&)
       If CVS(Mid$(XX$, 75, 4)) >= 0.05 Then
         Call MENSERR(24, "Imposible Anular.\Hay Ordenes de Despacho en Curso.")
         GoTo 2210
       End If
     Next UKKK&
     '
     RFQ$ = RECFILT$("DESPAPED", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
       UU& = CVS(Mid$(RFQ$, UKKK&, 4))
       XX$ = REGLEIDO$("DESPAPED", UU&)
       If Asc(Mid$(XX$, 21, 1)) < 9 Then
         If CVS(Mid$(XX$, 95, 4)) > 0 Then
           Call MENSERR(24, "Imposible Anular.\Ya se Despachó Material.")
           GoTo 2210
         End If
       End If
     Next UKKK&
     '
     'VERIFICAR CONSISTENCIA
     '
     If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
        Call MENSERR(24, "Imposible Anular este Pedido")
        GoTo 2210
     End If
     '
     RPENCA$ = REGLEIDO$("PEDENCA", REGENCA&)
     If CVS(Left$(RPENCA$, 4)) <> NUPEDIDO& Then
        Call MENSERR(24, "Imposible Anular este Pedido")
        GoTo 2210
     End If
     '
     NROPE& = CVS(Left$(X$, 4))
     If COMALTER%("Confirme la Anulación del Pedido Número " + TRIM$(Str$(NROPE&)) + "\Correspondiente a: " + TRIM$(Mid$(X$, 5, 48)) + "\\Cancelar\Confirmar Anulación") <> 2 Then
        GoTo 2210
     End If
     '
     ' RGL% = Int(REGENCA&)
     '
     Screen.MousePointer = 11
     Call BLOQARCH("PEDDETA")
     Call BLOQARCH("PEDDETP")
     '
     REPARAX% = 0
     RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
        I& = CVS(Mid$(RFQ$, UKKK&, 4))
        X$ = REGLEIDO$("PEDDETA", I&)
        If CVS(Left$(X$, 4)) = NROPE& Then
           If UCase$(Mid$(X$, 51, 1)) = "R" Then REPARAX% = 1
           Call REPLA(X$, Chr$(9), 27, 1)
           Call GRAREG("PEDDETA", X$, I&)
        End If
     Next UKKK&
     Call CIERRARCH("PEDDETA")
     '
     RFQ$ = RECFILT$("PEDDETP", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
        I& = CVS(Mid$(RFQ$, UKKK&, 4))
        X$ = REGLEIDO$("PEDDETP", I&)
        If CVS(Left$(X$, 4)) = NROPE& Then
           X$ = REGBLAN$("PEDDETP")
           Call GRAREG("PEDDETP", X$, I&)
        End If
     Next UKKK&
     Call CIERRARCH("PEDDETP")
     '
     RFQ$ = RECFILT$("DESPAPED", 1, MKS$(NUPEDIDO&))
     For UKKK& = 1 To Len(RFQ$) Step 4
        I& = CVS(Mid$(RFQ$, UKKK&, 4))
        X$ = REGLEIDO$("DESPAPED", I&)
        If CVS(Left$(X$, 4)) = NROPE& Then
           Call REPLA(X$, Chr$(9), 27, 1)
           Call GRAREG("DESPAPED", X$, I&)
        End If
     Next UKKK&
     Call CIERRARCH("DESPAPED")
     '
     X$ = REGLEIDO$("PEDENCA", REGENCA&)
     Call REPLA(X$, Chr$(9), 267, 1)
     Call GRAREG("PEDENCA", X$, REGENCA&)
     Call CIERRARCH("PEDENCA")
     '
     J& = 0
     FIN& = ULTREG&("ODESEMI")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       Y$ = REGLEIDO$("ODESEMI", I&)
       If CVS(Mid$(Y$, 1, 4)) = NROPE& Then
         Y$ = REGBLAN$("ODESEMI")
         Call GRAREG("ODESEMI", Y$, I&)
       End If
     Next I&
     Call CIERRARCH("ODESEMI")
     '
     HOII% = HOY(HOS$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, HOS$ + "  " + "Anulación Interactiva", 1, 56)
     Call REPLA(TTXX$, USERTER$, 69, 24)
     Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
     Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
     Call CIERRARCH("ANOTAPED")
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
     GoTo 2000
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'

Sub DESANUPED()
     '
     YAINDI% = 0
     Call COPYSTRU("INDIPEP", "INDIPEQ")
     '
2300 Screen.MousePointer = 11
     JJ& = 0
     Call FRESHECHO("INDIPEQ")
     RFF$ = RECFILT$("PEDENCA", 14, Chr$(9))
     For UU& = 1 To Len(RFF$) Step 4
       I& = CVS(Mid$(RFF$, UU&, 4))
       X$ = REGLEIDO$("PEDENCA", I&)
       If Asc(Mid$(X$, 267, 1)) = 9 Then
          Call VERIREFPED(X$, I&)
          JJ& = JJ& + 1
          Call GRAREG("INDIPEQ", Left$(X$, 52) + MKS$(I&), JJ&)
       End If
     'Next I&
     Next UU&
     Call SETULTREG("INDIPEQ", JJ&)
     Call CIERRARCH("INDIPEQ")
     Screen.MousePointer = 1
     '
     If ULTREG&("INDIPEQ") < 1 Then
        Call COMUNI("No hay Pedidos Anulados\que Puedan Re-Abrirse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2310 X$ = REGSEL$("INDIPEQ/Pedidos Anulados - Revertir Anulación")
     If Len(X$) <= 4 Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     NROPE& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     RGL% = Int(REGENCA&)
     Call BLOQARCH("PEDDETP")
     '
     Screen.MousePointer = 11
     X$ = REGLEIDO$("PEDENCA", REGENCA&)
     Call REPLA(X$, Chr$(1), 267, 1)
     Call GRAREG("PEDENCA", X$, REGENCA&)
     Call CIERRARCH("PEDENCA")
     '
     J& = 0
     FIN& = ULTREG&("PEDDETA")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("PEDDETA", I&)
        PENDE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
        If PENDE < 0 Then PENDE = 0
        Call REPLA(X$, MKS$(PENDE), 99, 4)
        If CVS(Left$(X$, 4)) = NROPE& Then
        Call REPLA(X$, Chr$(1), 27, 1)
        Call GRAREG("PEDDETA", X$, I&)
        '
        PDC$ = X$
        CANOR = CVS(Mid$(X$, 79, 4))
        Call REPLA(PDC$, Chr$(1), 27, 1)
        Call REPLA(PDC$, String$(38, 0), 71, 38)
        Call REPLA(PDC$, MKS$(CANOR), 75, 4)
        Call REPLA(PDC$, MKS$(CANOR), 99, 4)
        Call REGAPP("DESPAPED", PDC$)
        '
        Call REPLA(X$, MKS$(I&), 109, 4)       ' puntero a PEDDETA
        Call REGAPP("PEDDETP", X$)
        End If
     Next I&
     Call CIERRARCH("PEDDETP")
     Call CIERRARCH("PEDDETA")
     '
     HOII% = HOY(HOS$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, HOS$ + "  Reversión de Anulación Interactiva", 1, 56)
     Call REPLA(TTXX$, USERTER$, 69, 24)
     Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
     Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
     Call CIERRARCH("ANOTAPED")
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     Screen.MousePointer = 1
     '
     GoTo 2300
     '
End Sub
'

Sub ANULITPED()
     '
2000 YAINDI% = 0
     '
2100 If YAINDI% = 0 Then
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        RFF$ = RECFILT$("PEDDETA", 7, Chr$(1)) ' SOLO PENDIENTES DE ENTREGA Y NO DESPACHADOS
        FIN& = Len(RFF$)
        For UU& = 1 To Len(RFF$) Step 4
          Call TRACE(20, UU&, FIN&)
          REPRESU& = CVS(Mid$(RFF$, UU&, 4))
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            '
            If Asc(Mid$(X$, 27, 1)) < 3 Then
              If CAPENDENT < 0.05 Then
                Call REPLA(X$, Chr$(3), 27, 1)
                Call GRAREG("PEDDETA", X$, REPRESU&)
              End If
            End If
            '
            If CAPENDENT >= 0.05 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
          '
        Next UU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        Call FRESHECHO("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación por Items")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NUPEDIDO& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     DPEDI$ = TRIM$(Str$(NUPEDIDO&)) + " del " + TRIM$(Mid$(X$, 5, 48))
     UREODES& = ULTREG&("PEDDETA")
     JJKK& = 0
     '   If TIDOC% <> 1 Then GoTo 1211
     '
     CAPESEL% = 1
     TTXX$ = DPEDI$
     NUODSEL& = Val(Mid$(TTXX$, 15, 10))
     REPRESU& = Val(Mid$(TTXX$, 121, 8))
     '
1205 TINI = Timer
     '
     SELITPED.List2.Clear
     '
     RFF$ = RECFILT$("PEDDETA", 1, MKS$(NUPEDIDO&)) ' SOLO PENDIENTES DE ENTREGA Y NO DESPACHADOS
     For UU& = 1 To Len(RFF$) Step 4
       UJ& = CVS(Mid$(RFF$, UU&, 4))
       REGPEDA& = UJ&
       X$ = REGLEIDO$("PEDDETA", REGPEDA&)
       If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
         CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
         If CAN >= 0.05 Then
           CI% = CVI(Mid$(X$, 33, 2))
           TTYY$ = Space$(128)
           Call REPLA(TTYY$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTYY$, DESCRIT0$(CI%), 19, 38)
           Call REPLA(TTYY$, FORMATNUM$(CAN, "F10.1"), 57, 10)
           Call REPLA(TTYY$, FORMATNUM$(REGODES&, "F8.0"), 113, 8)
           Call REPLA(TTYY$, FORMATNUM$(REGPEDA&, "F8.0"), 121, 8)
           SELITPED.List2.AddItem TTYY$
           '
         End If
       End If
1208 Next UU&
     '
1209 SELITPED.Caption = "Anulacion por Items Pedido " + DPEDI$
     SELITPED.Show 1
     DoEvents
     '
     If SELITPED.List2.ListCount > 0 Then
       Screen.MousePointer = 11
       For U& = 1 To SELITPED.List2.ListCount
         TTXX$ = SELITPED.List2.List(U& - 1)
         If Mid$(TTXX$, 71, 5) = "anul." Then
           REGPEDA& = Val(Mid$(TTXX$, 121, 8))
           If REGPEDA& > 0 Then
             If REGPEDA& <= ULTREG&("PEDDETA") Then
               XX$ = REGLEIDO$("PEDDETA", REGPEDA&)
               Call REPLA(XX$, Chr$(9), 27, 1)
               Call GRAREG("PEDDETA", XX$, REGPEDA&)
             End If
           End If
         End If
       Next U&
       Call CIERRARCH("PEDDETA")
       '
       Call BLOQARCH("PEDDETP")
       Call RECAPEDPE
       Call CIERRARCH("PEDDETP")
       Screen.MousePointer = 1
       '
       HOII% = HOY(HOS$)
       TTXX$ = Space$(96)
       Call REPLA(TTXX$, HOS$ + "  " + "Anulación Parcial de Items Pedidos", 1, 56)
       Call REPLA(TTXX$, USERTER$, 69, 24)
       Call REGAPP("ANOTAPED", Left$(MKS$(NROPE&) + TTXX$, 96))
       Call REGAPP("ANOTAPED", MKS$(NROPE&) + String$(88, "="))
       Call CIERRARCH("ANOTAPED")
       '
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     DoEvents
     '
     GoTo 2000
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub

Private Sub Command5_Click()
    ENSANOR.Show 1
End Sub

'
Private Sub Command62_Click()
    Command62.Enabled = False
    'If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
    '    Call LIPEDIT(3)
    '    GoTo 99
    'End If
    '
    If Option1.Value = True Then
           Call RECAPEDPE
           SALI$ = "*COPEDCL"
           If FORMAPED% = 2 Then SALI$ = "*COQEDCL"
           Call FINAL(SALI$)
        Else
           If EXISTE%("ANAPEP02.EXE") > 0 Then
             Call CONECRUN("ANAPEP02/por Cliente", "Analisis Pedidos Pendientes por Cliente")
             GoTo 99
           End If
           '
           Screen.MousePointer = 11
           JJ& = 0: LL& = 0
           FIN& = ULTREG&("PEDDETA")
           '
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
           If MODO% < 1 Then GoTo 99
           FLI1% = -1: FLI2% = FINMESACT%
           If MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
10             FNMA = FINMESACT%
               FLI1% = DIPOST(FNMA): FLI2% = 32767
               VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
               OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
               If OBX$ = "" Then GoTo 99
               '
               FLI1% = CVI(Left$(OBX$, 2))
               FLI2% = CVI(Mid$(OBX$, 3, 2))
               If FLI1% > FLI2% Then GoTo 10
               FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
               FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           Call HEADERS("LIPEDPE", "")
           VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
           If MODO% = 4 Then VARIANTE$ = "Pendientes sin Orden de Despacho"
           Call HEADERS("LIPEDPE", VARIANTE$)
           '
           Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
           If NCL1% < 1 Then GoTo 99
           '
           VECPEDI$ = ""
           For U& = 1 To FIN&
             Call TRACE(20, U&, FIN&)
             X$ = REGLEIDO$("PEDDETA", U&)
             NC% = CVI(Mid$(X$, 7, 2))
             If NC% >= NCL1% Then
              If NC% <= NCL2% Then
               If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
                If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
                 CI% = CVI(Mid$(X$, 33, 2))
                 If CI% > 0 And CI% <= NUMAS% Then
                   CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                   If CAPEN < 0 Then CAPEN = 0
                   If CAPEN > 0.05 Then
                     X1$ = X$
                     'Call REPLA(X1$, MKS$(U&), 109, 4)
                     'LL& = LL& + 1
                     'Call GRAREG("PEDDETP", X1$, LL&)
                     FENTRESOL% = CVI(Mid$(X$, 25, 2))
                     If FENTRESOL% >= FLI1% Then
                      If FENTRESOL% <= FLI2% Then
                        If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                          Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                          Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                          Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                          Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                          '
                          CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                          If CAPEN < 0 Then CAPEN = 0
                          CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                          If CANOASI < 0 Then CANOASI = 0
                          '
                          NC% = CVI(Mid$(X$, 7, 2))
                          PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                          If OPX% <> 2 Then PREUNI = 0
                          ITOTA# = CDbl(CAPEN * PREUNI)
                          INOAS# = CDbl(CANOASI * PREUNI)
                          '
                          Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                          Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                          Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                          Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                          Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                          Call REPLA(Y$, Space$(34), 127, 34)
                          '
                          JJ& = JJ& + 1
                          Call GRAREG("LIPEDPE", Y$, JJ&)
                          VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
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
           'Call SETULTREG("PEDDETP", LL&)
           'Call CIERRARCH("PEDDETP")
           GoTo 30
           '
           '
           '
           '
           FIN& = ULTREG&("ORDESPA")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ORDESPA", U&)
             CADESPAX = CVS(Mid$(X$, 79, 4))
             If CADESPAX >= 0.05 Then
               ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
               For W& = 1 To Len(VECPEDI$) / 8
                 If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                   Y$ = REGLEIDO$("LIPEDPE", W&)
                   ODX$ = TRIM$(Mid$(Y$, 127, 34))
                   If ODX$ = "" Then
                       ODX$ = "OD"
                     Else
                       ODX$ = ODX$ + ","
                   End If
                   ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
                   Call REPLA(Y$, ODX$, 127, 34)
                   '
                   If MODO% = 4 Then
                      CADESPAX = CVS(Mid$(X$, 79, 4))
                      CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
                      CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
                      Call REPLA(Y$, MKS$(CAENTX), 97, 4)
                      Call REPLA(Y$, MKS$(CASALX), 101, 4)
                   End If
                   '
                   Call GRAREG("LIPEDPE", Y$, W&)
                   GoTo 19
                 End If
               Next W&
             End If
19         Next U&
           '
           FIN& = ULTREG&("ODESEMI")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ODESEMI", U&)
             ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
             For W& = 1 To Len(VECPEDI$) / 8
               If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                 Y$ = REGLEIDO$("LIPEDPE", W&)
                 ODX$ = TRIM$(Mid$(Y$, 127, 34))
                 If ODX$ = "" Then
                     CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
                     If Val(CASIGNA$) >= 0.05 Then
                        Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                        Call GRAREG("LIPEDPE", Y$, W&)
                     End If
                 End If
                 GoTo 29
               End If
             Next W&
29         Next U&
           '
30         Call HEADERS("LIPEDPE", "")
           If MODO% = 2 Then
              Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
           End If
           '
           If MODO% = 4 Then
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
           VARIANTE$ = "Totales por Entregar (con / sin Orden de Despacho)"
           If MODO% = 4 Then VARIANTE$ = "Pedidos Pendientes sin Orden de Despacho"
           Call HEADERS("LIPEDPE", "Corresponde a: " + VARIANTE$)
           If VENDSEL% > 0 Then
             Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
           End If
           '
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENCL"
           If FORMAPED% = 2 Then SALI$ = "*SOQENCL"
           Call FINAL(SALI$)
    End If
99  Screen.MousePointer = 1
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    'If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
    '    Call LIPEDIT(2)
    '    GoTo 99
    'End If
    '
    If Option2.Value = True Then
      If InStr(EMPREAC$, "AMBEST") > 0 Then
        Call LIDETPE
        GoTo 99
      End If
    End If
    '
    If Option1.Value = True And InStr(EMPREAC$, "DOSIVA") < 1 Then
           SALI$ = "*COPEDNO"
           If FORMAPED% = 2 Then SALI$ = "*COQEDNO"
           Call FINAL(SALI$)
        Else
           Call LIPENOFA
    End If
99  Command63.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    'If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
    '    Call LIPEDIT(4)
    '    GoTo 99
    'End If
    '
    If Option1.Value = True Then
           Call RECAPEDPE
           SALI$ = "*COPEDMA"
           If FORMAPED% = 2 Then SALI$ = "*COQEDMA"
           Call FINAL(SALI$)
        Else
           If EXISTE%("ANAPEP02.EXE") > 0 Then
             Call CONECRUN("ANAPEP02/por Articulo", "Analisis Pedidos Pendientes por Cliente")
             GoTo 99
           End If
           '
           Screen.MousePointer = 11
           JJ& = 0: LL& = 0
           FIN& = ULTREG&("PEDDETA")
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
           If MODO% < 1 Then GoTo 99
           FLI1% = -1: FLI2% = FINMESACT%
           '
           If MODO% = 2 Then
10             FNMA = FINMESACT%
               FLI1% = DIPOST(FNMA): FLI2% = 32767
               VDEF$ = MKI$(FLI1%) + MKI$(FLI2%)
               OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
               If OBX$ = "" Then GoTo 99
               '
               FLI1% = CVI(Left$(OBX$, 2))
               FLI2% = CVI(Mid$(OBX$, 3, 2))
               If FLI1% > FLI2% Then GoTo 10
               FEX = FLI1%: PERIO$ = FECHATEX$(FEX)
               FEX = FLI2%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
           If NCL1% < 1 Then GoTo 99
           '
           VECPEDI$ = ""
           For U& = 1 To FIN&
             Call TRACE(20, U&, FIN&)
             X$ = REGLEIDO$("PEDDETA", U&)
             NC% = CVI(Mid$(X$, 7, 2))
             If NC% >= NCL1% Then
              If NC% <= NCL2% Then
               If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
                If Asc(Mid$(X$, 27, 1)) <> 3 Then ' no esta entregado
                 CI% = CVI(Mid$(X$, 33, 2))
                 If CI% > 0 And CI% <= NUMAS% Then
                   CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                   If CAPEN < 0 Then CAPEN = 0
                   If CAPEN > 0.05 Then
                     X1$ = X$
                     'Call REPLA(X1$, MKS$(U&), 109, 4)
                     'LL& = LL& + 1
                     'Call GRAREG("PEDDETP", X1$, LL&)
                     FENTRESOL% = CVI(Mid$(X$, 25, 2))
                     If FENTRESOL% >= FLI1% Then
                      If FENTRESOL% <= FLI2% Then
                       If VENDSEL% < 1 Or VENDCLI%(NC%) = VENDSEL% Then
                         Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                         Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                         Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                         Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                         '
                         CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                         If CAPEN < 0 Then CAPEN = 0
                         CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                         If CANOASI < 0 Then CANOASI = 0
                         '
                         NC% = CVI(Mid$(X$, 7, 2))
                         PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                         If OPX% <> 2 Then PREUNI = 0
                         ITOTA# = CDbl(CAPEN * PREUNI)
                         INOAS# = CDbl(CANOASI * PREUNI)
                         
                         Call REPLA(Y$, MKS$(CAPEN), 101, 4)
                         Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                         Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                         Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                         Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                         Call REPLA(Y$, Space$(34), 127, 34)
                         '
                         JJ& = JJ& + 1
                         Call GRAREG("LIPEDPE", Y$, JJ&)
                         VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
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
           'Call SETULTREG("PEDDETP", LL&)
           'Call CIERRARCH("PEDDETP")
           '
           GoTo 30
           '
           '
           '
           '
           FIN& = ULTREG&("ORDESPA")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ORDESPA", U&)
             CADESPAX = CVS(Mid$(X$, 79, 4))
             If CADESPAX >= 0.05 Then
               ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
               For W& = 1 To Len(VECPEDI$) / 8
                 If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                   Y$ = REGLEIDO$("LIPEDPE", W&)
                   ODX$ = TRIM$(Mid$(Y$, 127, 34))
                   If ODX$ = "" Then
                       ODX$ = "OD"
                     Else
                       ODX$ = ODX$ + ","
                   End If
                   ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
                   Call REPLA(Y$, ODX$, 127, 34)
                   '
                   If MODO% = 4 Then
                      CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
                      CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
                      Call REPLA(Y$, MKS$(CAENTX), 97, 4)
                      Call REPLA(Y$, MKS$(CASALX), 101, 4)
                   End If
                   '
                   Call GRAREG("LIPEDPE", Y$, W&)
                   GoTo 19
                 End If
               Next W&
             End If
19         Next U&
           '
           FIN& = ULTREG&("ODESEMI")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ODESEMI", U&)
             ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
             For W& = 1 To Len(VECPEDI$) / 8
               If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                 Y$ = REGLEIDO$("LIPEDPE", W&)
                 ODX$ = TRIM$(Mid$(Y$, 127, 34))
                 If ODX$ = "" Then
                    CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
                      If Val(CASIGNA$) > 0.05 Then
                        Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                        Call GRAREG("LIPEDPE", Y$, W&)
                      End If
                 End If
                 GoTo 29
               End If
             Next W&
29         Next U&
           '
30         Call HEADERS("LIPEDPE", "")
           If MODO% = 2 Then
              Call HEADERS("LIPEDPE", "Entrega Solicitada: " + PERIO$)
           End If
           '
           If MODO% = 4 Then
             KK& = JJ&: JJ& = 0
             For U& = 1 To KK&
               Y$ = REGLEIDO$("LIPEDPE", U&)
               If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
                 JJ& = JJ& + 1
                 Call GRAREG("LIPEDPE", Y$, JJ&)
               End If
             Next U&
             Call HEADERS("LIPEDPE", "Corresponde a: Pedidos Pendientes sin Orden de Despacho")
             If VENDSEL% > 0 Then
               Call HEADERS("LIPEDPE", "Vendedor: " + ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
             End If
           End If
           '
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENMA"
           If FORMAPED% = 2 Then SALI$ = "*SOQENMA"
           Call FINAL(SALI$)
    End If
99  Screen.MousePointer = 1
    Command64.Enabled = True
End Sub

Function VERIPEDPE%(REGPE$)
    '
    Static CTX&, CTY&
    If CTX& = 0 Then
        CTX& = ULTREG&("PEDENCA")
        '
        If BLOQRESULT%("PEDIFAC") <> 1 Then
           TMENSA$ = "Asignación de Pedidos"
           If TRIM$(UCase$(App.EXEName)) = "ADMIPED" Then
              TMENSA$ = "Facturación"
           End If
           Call MENSERR(24, "Proceso Imposible en este Momento.\Proceso de " + TMENSA$ + " Abierto.\  \Resuelva y Vuelva a Intentarlo.")
           Call FINAL("")
        End If
        '
    End If
    '
    If CTY& = 0 Then
        CTY& = ULTREG&("PEDDETA")
    End If
    '
    VPP% = 0
    '
    X1$ = REGPE$
    NPED0& = CVS(Left$(X1$, 4))
    REGA& = CVS(Mid$(X1$, 109, 4))
    Call REPLA(X1$, String$(4, 0), 109, 4)
    Call REPLA(X1$, String$(2, 0), 31, 2)
10  If REGA& > 0 Then
       If REGA& <= CTY& Then
          X2$ = REGLEIDO$("PEDDETA", REGA&)
          REGENCA& = CVS(Mid$(X2$, 109, 4))
          Call REPLA(X2$, String$(4, 0), 109, 4)
          Call REPLA(X2$, String$(2, 0), 31, 2)
          If X1$ = X2$ Then
             NPEDI& = CVS(Left$(X2$, 4))
             If REGENCA& > 0 Then
                If REGENCA& <= CTX& Then
                   X3$ = REGLEIDO$("PEDENCA", REGENCA&)
                   NPEDJ& = CVS(Left$(X3$, 4))
                   If NPEDJ& = NPEDI& Then
                      VPP% = 1
                      GoTo 999
                   End If
                   MRX$ = "(1)"
                End If
                MRX$ = "(2)"
             End If
             MRX$ = "(3)"
          End If
          MRX$ = "(4)"
       End If
       MRX$ = "(5)"
    End If
    MRX$ = "(6)"
    '
    Screen.MousePointer = 1
    MENSA$ = "FLEXOFT Debe Ahora Reorganizar\"
    MENSA$ = MENSA$ + "la Base de Datos de Pedidos Pendientes.\"
    MENSA$ = MENSA$ + "Error " + MRX$ + " en Pedido " + TRIM$(Str$(NPED0&)) + ".\  \"
    MENSA$ = MENSA$ + "Pinche sobre 'Aceptar Mensaje'\"
    MENSA$ = MENSA$ + "para Iniciar el Proceso,\"
    MENSA$ = MENSA$ + "y Espere Hasta su Finalización."
    Call COMUNI(MENSA$)
    '
    Screen.MousePointer = 11
    '
    Dim NPI&(65535)
    If CTZ& = 0 Then
       CTZ& = ULTREG&("PEDENCA")
       For U& = 1 To CTZ&
           Call TRACE(20, U&, CTZ&)
           X3$ = REGLEIDO$("PEDENCA", U&)
           NPI&(U&) = CVS(Left$(X3$, 4))
       Next U&
    End If
    '
    Call BLOQARCH("PEDDETP")
    J& = 0
    FIN& = ULTREG&("PEDDETA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X1$ = REGLEIDO$("PEDDETA", U&)
       NPEDI& = CVS(Left$(X1$, 4))
       If NPEDI& > 0 Then
          REGENCA& = CVS(Mid$(X1$, 109, 4))
          If REGENCA& > 0 Then
             If REGENCA& <= CTZ& Then
                X3$ = REGLEIDO$("PEDENCA", REGENCA&)
                If CVS(Left$(X3$, 4)) = NPEDI& Then
                   GoTo 100
                End If
             End If
          End If
          '
          For UI& = 1 To CTZ&
             If NPI&(UI&) = NPEDI& Then
                REGENCA& = UI&
                Call REPLA(X1$, MKS$(REGENCA&), 109, 4)
                Call GRAREG("PEDDETA", X1$, U&)
                GoTo 100
             End If
          Next UI&
          '
          If NPEDI& <> NPEDIA& Then
             Call COMUNI("No se Encontró el Encabezado\del Pedido Número" + Str$(NPEDI&) + ".\Este Pedido Queda Como 'Anulado'\   \Consulte a su Soporte de Sistemas FLEXOFT.")
             NPEDIA& = NPEDI&
          End If
          Call REPLA(X1$, Chr$(9), 27, 1)
          Call GRAREG("PEDDETA", X1$, U&)
       End If
       '
100    CAPED = CVS(Mid$(X1$, 79, 4))
       CAENT = CVS(Mid$(X1$, 95, 4))
       CAPEN = CAPED - CAENT: If CAPEN < 0 Then CAPEN = 0
       If CVS(Mid$(X1$, 99, 4)) <> CAPEN Then
          Call REPLA(X1$, MKS$(CAPEN), 99, 4)
          Call GRAREG("PEDDETA", X1$, U&)
       End If
       '
       If Asc(Mid$(X1$, 27, 1)) <= 1 Then
            If CAPEN >= 0.05 Then
                Call REPLA(X1$, MKS$(U&), 109, 4)
                J& = J& + 1
                Call GRAREG("PEDDETP", X1$, J&)
            End If
       End If
       '
    Next U&
    '
    Call SETULTREG("PEDDETP", J&)
    Call CIERRARCH("PEDDETP")
    Screen.MousePointer = 1
    '
999 VERIPEDPE% = VPP%
    '
End Function

Sub VERIREFPED(RPENCA$, REGENCA&)
    '
    REFE$ = TRIM$(Mid$(RPENCA$, 5, 48))
    If Left$(REFE$, 1) = "-" Then
        NCI% = CVI(Mid$(RPENCA$, 55, 2))
        RASO$ = TRIM$(RASOCLI$(NCI%))
        If InStr(REFE$, RASO$) < 1 Then
            REFE$ = RASO$ + " " + Mid$(RPENCA$, 5, 48)
            Call REPLA(RPENCA$, REFE$, 5, 48)
            Call GRAREG("PEDENCA", RPENCA$, REGENCA&)
        End If
    End If
    '
End Sub
'
Sub LIPENOFA()
    '
    Dim NUPE&(4096), FEPEDIDO%(4096), CLIENTE%(4096), TOPEDIDO#(4096), FACTURADO#(4096), RPEFA$(4096)
    CAPEDIPEN% = 0
    Dim REMIL&(4096), NCREMI%(4096), CIREMI%(4096), CAREMI(4096)
    CAREMIPE% = 0
    '
    Screen.MousePointer = 11
    Call RECAPEDPE
    Screen.MousePointer = 11
    '
    FIN& = ULTREG&("CONSICLI")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CONSICLI", U&)
      ncx8% = CVI(Left$(XX$, 2))
      CIXI% = CVI(Mid$(XX$, 35, 2))
      If CIXI% > 0 Then
        If CIXI% <= NUMAS% Then
          CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
          If CACON > 0 Then
              CAREMIPE% = CAREMIPE% + 1
              REMIL&(CAREMIPE%) = Val(Mid$(XX$, 26, 9))
              NCREMI%(CAREMIPE%) = ncx8%
              CIREMI%(CAREMIPE%) = CIXI%
              CAREMI(CAREMIPE%) = CACON
            Else
              REMIZ& = Val(Mid$(XX$, 26, 9))
              For CRPX% = 1 To CAREMIPE%
                If REMIL&(CRPX%) = REMIZ& Or REMIZ& < 1 Then
                  If NCREMI%(CRPX%) = ncx8% Then
                    If CIREMI%(CRPX%) = CIXI% Then
                      If CAREMI(CRPX%) > 0.05 Then
                        CAREMI(CRPX%) = CAREMI(CRPX%) + CACON
                        GoTo 19
                      End If
                    End If
                  End If
                End If
              Next CRPX%
          End If
        End If
      End If
19  Next U&
    '
    'For UI% = 1 To CAREMIPE%
    'CACA = CAREMI(UI%)
    'If Abs(CACA) > 0 Then
    '  cccc$ = CODEXT$(CIREMI%(UI%))
    '  NREMI = REMIL&(UI%)
    '  NCLIE% = NCREMI(UI%)
    '  AAA = bbb
    'End If
    'Next UI%
    '
    JJ& = 0: LL& = 0
    FIN& = ULTREG&("PEDDETA")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETA", U&)
      NUPEDID& = CVS(Left$(X$, 4))
      FEPEXI% = CVI(Mid$(X$, 5, 2))
      NC% = CVI(Mid$(X$, 7, 2))
      If NC% > 0 Then
        If Asc(Mid$(X$, 27, 1)) < 9 Then ' no esta anulado NI ENTREGADO
          CI% = CVI(Mid$(X$, 33, 2))
          If CI% > 0 And CI% <= NUMAS% Then
             CATOPED = CVS(Mid$(X$, 79, 8))
             CAENTRE = CVS(Mid$(X$, 95, 4))
             CAPEN = CATOPED - CAENTRE
             NREMIPE& = CVS(Mid$(X$, 105, 4))
             REMIPEFA$ = ""
             If NREMIPE& > 0 Then GoSub 50
             '
             If CAPEN < 0 Then CAPEN = 0
             If CAPEN > 0.05 Then
               PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
               ITOTA# = CDbl(CATOPED * PREUNI)
               IFACTU# = CDbl(CAENTRE * PREUNI)
               If IFACTU# > ITOTA# Then IFACTU# = ITOTA#
               IPENDE# = CDbl(CAPEN * PREUNI)
               '
               If IPENDE# > 0.005 Or CAPEN > 0.05 Then
' AJUSTE PARA DOSIVAC DEL 17/09/02
If Asc(Mid$(X$, 27, 1)) = 3 Then
  If NREMIPE& < 1 Then
    Call REPLA(X$, Chr$(9), 27, 1)
    Call GRAREG("PEDDETA", X$, U&)
    GoTo 29
    ' SI MARCADO COMO ENTREGADO Y NO PENDIENTE LO ANULA
  End If
End If
                 If NUPEDID& = NUPE&(CAPEDIPEN%) Then
                      TOPEDIDO#(CAPEDIPEN%) = TOPEDIDO#(CAPEDIPEN%) + ITOTA#
                      FACTURADO#(CAPEDIPEN%) = FACTURADO(CAPEDIPEN%) + IFACTU#
                    Else
                      CAPEDIPEN% = CAPEDIPEN% + 1
                      NUPE&(CAPEDIPEN%) = NUPEDID&
                      FEPEDIDO%(CAPEDIPEN%) = FEPEXI%
                      CLIENTE%(CAPEDIPEN%) = NC%
                      TOPEDIDO#(CAPEDIPEN%) = ITOTA#
                      FACTURADO#(CAPEDIPEN%) = IFACTU#
                      RPEFA$(CAPEDIPEN%) = REMIPEFA$
                 End If
               End If
             End If
          End If
        End If
      End If
29  Next U&
    '
    For U& = 1 To CAPEDIPEN%
      X$ = String$(64, 0)
      Call REPLA(X$, MKS$(NUPE&(U&)), 1, 4)
      Call REPLA(X$, MKI$(FEPEDIDO%(U&)), 5, 2)
      Call REPLA(X$, MKI$(CLIENTE%(U&)), 7, 2)
      Call REPLA(X$, MKD$(TOPEDIDO#(U&)), 9, 8)
      Call REPLA(X$, MKD$(FACTURADO#(U&)), 17, 8)
      Call REPLA(X$, MKD$(TOPEDIDO#(U&) - FACTURADO#(U&)), 25, 8)
      Call REPLA(X$, RPEFA$(U&), 33, 32)
      Call GRAREG("LIPENOFA", X$, U&)
      LL& = U&
   Next U&
   '
   Call SETULTREG("LIPENOFA", LL&)
   Call CIERRARCH("LIPENOFA")
   '
   Call FINAL("*SOPENOF")
   Screen.MousePointer = 1
   Exit Sub
   '
50 For CRPX% = 1 To CAREMIPE%
     If REMIL&(CRPX%) = NREMIPE& Then
       If NCREMI%(CRPX%) = NC% Then
         If CIREMI%(CRPX%) = CI% Then
           CAPEN = CAPEN + CAREMI(CRPX%)
           CAENTRE = CAENTRE - CAREMI(CRPX%)
           If InStr(" " + REMIPEFA$, " " + TRIM$(Str$(NREMIPE&))) < 1 Then
             If REMIPEFA$ <> "" Then REMIPEFA$ = REMIPEFA$ + ", "
             REMIPEFA$ = REMIPEFA$ + TRIM$(Str$(NREMIPE&))
           End If
         End If
       End If
     End If
   Next CRPX%
   Return
   '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("ODESPA02", "Ordenes de Despacho Pedidos Clientes")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    DoEvents
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    '
End Sub

Sub RECAPEDPE()
   '
   IFIX$ = INDEXSTRI$("PEDDETA", 7)
   RFF$ = "": POIN& = 1
   For IQQ& = 1 To Len(IFIX$)
      STATI% = Asc(Mid$(IFIX$, IQQ&, 1))
      If STATI% <> 9 Then
        If STATI% <> 3 Then
          If POIN& > Len(RFF$) Then RFF$ = RFF$ + String$(4096, 0)
          Mid$(RFF$, POIN&, 4) = MKS$(IQQ&)
          POIN& = POIN& + 4
        End If
      End If
   Next IQQ&
   RFF$ = Left$(RFF$, POIN& - 1)
   FIREFI& = Len(RFF$): V1& = 0
   '
   V1& = 0
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   Call BLOQARCH("PEDDETP")
   For UU& = 1 To FIREFI& Step 4
      '
      U& = CVS(Mid$(RFF$, UU&, 4))
      Call TRACE(20, UU&, FIREFI&)
      X$ = REGLEIDO$("PEDDETA", U&)
      '
      If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
        '
        CAPEDIDA = CVS(Mid$(X$, 79, 4))
        If CAPEDIDA < 0.05 Then       ' marca como anulado
          Call REPLA(X$, Chr$(9), 27, 1)
          Call GRAREG("PEDDETA", X$, U&)
        End If
        '
        If CAPEDIDA >= 0.05 Then
          '
          CADESPACH = CVS(Mid$(X$, 95, 4))
          CAPENDENT = CAPEDIDA - CADESPACH
          '
          If Asc(Mid$(X$, 27, 1)) < 3 Then
            If CAPENDENT < 0.05 Then   ' marca como entregado
              Call REPLA(X$, Chr$(3), 27, 1)
              Call GRAREG("PEDDETA", X$, U&)
            End If
          End If
          '
          If CAPENDENT >= 0.05 Then
            '
            NCLIE% = CVI(Mid$(X$, 7, 2))
            NPED& = CVS(Left$(X$, 4))
            '
            If NPED& > 0 Then
              '
              CI% = CVI(Mid$(X$, 33, 2))
              REGENCA& = CVS(Mid$(X$, 109, 4))
              '
              Z$ = REGLEIDO$("PEDENCA", REGENCA&)
              NPEDA& = CVS(Left$(Z$, 4))
              NCLIA% = CVI(Mid$(Z$, 55, 2))
              '
              If NPEDA& <> NPED& Or NCLIA% <> NCLIE% Then
                Call MENSERR(24, "Punteros Incorrectos en Pedido " + TRIM$(Str$(NPED&)) + "/" + TRIM$(Str$(NPEDA&)) + "\Consulte a su Soporte Técnico")
                GoTo 75
              End If
              '
              If Asc(Mid$(Z$, 267, 1)) > 1 Then
                Call REPLA(X$, Chr$(9), 27, 1)  ' anula detalle
                Call GRAREG("PEDDETA", X$, U&)
                GoTo 90
              End If
              '
              ' RE-GENERA CANTIDAD PENDIENTE
              If Abs(CVS(Mid$(X$, 99, 4)) - CAPENDENT) > 0.05 Then
                Call REPLA(X$, MKS$(CAPENDENT), 99, 4)
                Call GRAREG("PEDDETA", X$, U&)
              End If
              '
75            Y$ = X$
              Call REPLA(Y$, MKS$(U&), 109, 4) ' puntero a PEDDETA
              V1& = V1& + 1
              Call GRAREG("PEDDETP", Y$, V1&)
              '
            End If
          End If
        End If
      End If
      '
90 Next UU&
   '
   Call SETULTREG("PEDDETP", V1&)
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("PEDDETA")
   Call BAJALDISCO
   '
   Screen.MousePointer = SCMA
   '
End Sub

Sub LIDETPE()
   '
   Call COPYSTRU("PEDDETA", "PEDDETX")
   JJ& = 0
   FIN& = ULTREG&("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("PEDDETA", U&)
     If Asc(Mid$(X$, 27, 1)) < 9 Then ' NO ESTA ANULADO
       NRO& = CVS(Mid$(X$, 1, 4))
       FEPE% = CVI(Mid$(X$, 5, 2))
       NCLI% = CVI(Mid$(X$, 7, 2))
       CIXI% = CVI(Mid$(X$, 33, 2))
       CAPE = CVS(Mid$(X$, 79, 4))
       PRUN = CVS(Mid$(X$, 83, 4))
       If PRUN < 0.005 Then
         PRUN = PRECALCU(CIXI%)
         Call REPLA(X$, MKS$(PRUN), 83, 4)
       End If
       ITOX# = CDbl(PRUN * CAPE)
       Call REPLA(X$, MKD$(ITOX#), 87, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("PEDDETX", X$, JJ&)
     End If
   Next U&
   '
   Call SETULTREG("PEDDETX", JJ&)
   Call CIERRARCH("PEDDETX")
   Call CIERRARCH("*.*")
   '
   Call FILESORT("PEDDETX", "", 1, 1, "ASC")
   Call CONECRUN("*LIPEDET", "Listado Detalle de Pedidos")
   '
End Sub

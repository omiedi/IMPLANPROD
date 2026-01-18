VERSION 5.00
Begin VB.Form MOVIMSTO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5955
   Icon            =   "MOVIMSTO.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Anul."
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
      Left            =   4960
      Picture         =   "MOVIMSTO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Anulación de Movimientos de Stock"
      Top             =   2650
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4890
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   26
      Top             =   3780
      Width           =   965
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4935
      TabIndex        =   25
      ToolTipText     =   "Recuperación Movimientos Perdidos"
      Top             =   4725
      Visible         =   0   'False
      Width           =   750
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
      Height          =   750
      Left            =   4960
      Picture         =   "MOVIMSTO.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1850
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
      Left            =   4960
      Picture         =   "MOVIMSTO.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
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
      Left            =   4960
      Picture         =   "MOVIMSTO.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   250
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0EEC0&
      Caption         =   "Movimientos de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   210
      TabIndex        =   7
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Expedición y Despacho"
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
         TabIndex        =   10
         Top             =   1785
         Width           =   4110
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   3630
            TabIndex        =   15
            Top             =   420
            Width           =   3690
            Begin VB.CommandButton Command10 
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
               Left            =   3070
               Picture         =   "MOVIMSTO.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Lista de Empaque"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command23 
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
               Picture         =   "MOVIMSTO.frx":11B4
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Remito de Ventas - Clientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command8 
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
               Left            =   2510
               Picture         =   "MOVIMSTO.frx":14BE
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Remito de Traslado"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command22 
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
               Left            =   560
               Picture         =   "MOVIMSTO.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Remito en Consignacion - Proveedores de Servicios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Comprobantes de  Despacho"
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
               Left            =   1200
               TabIndex        =   23
               Top             =   105
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Fábrica - Ajustes - Transferencias"
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
         TabIndex        =   9
         Top             =   3150
         Width           =   4110
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   3630
            TabIndex        =   16
            Top             =   420
            Width           =   3690
            Begin VB.CommandButton Command17 
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
               Left            =   3070
               Picture         =   "MOVIMSTO.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Descartes de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command24 
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
               Left            =   560
               Picture         =   "MOVIMSTO.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Transferencias de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command25 
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
               Picture         =   "MOVIMSTO.frx":20E6
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command21 
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
               Left            =   2520
               Picture         =   "MOVIMSTO.frx":2230
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Vale de Consumo de Materiales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos     Varios"
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
               Left            =   1260
               TabIndex        =   17
               Top             =   105
               Width           =   1110
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Entradas a Depósito"
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
         TabIndex        =   8
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   13
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command18 
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
               Picture         =   "MOVIMSTO.frx":237A
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Producción"
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
               TabIndex        =   14
               Top             =   105
               Width           =   1110
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   11
            Top             =   420
            Width           =   1800
            Begin VB.CommandButton Command28 
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
               Picture         =   "MOVIMSTO.frx":2684
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción Proveedores"
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
               TabIndex        =   12
               Top             =   105
               Width           =   1110
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
      Left            =   4380
      Picture         =   "MOVIMSTO.frx":298E
      Top             =   4210
      Width           =   2010
   End
End
Attribute VB_Name = "MOVIMSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CONECRUN("LISEMPAQ", "Generación de Listas de Empaque")
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If CLACONTRA% = 1 Then
       If EXISTE%(LUDAT$ + "MOVISTK.DAT") > 0 Then
         Call COMUNI("Iniciando Recuperación Remitos 9554 a 9599")
         Call COPYSTRU("MOVISTO", "MOVISTK")
         Call BLOQARCH("MOVISTO")
         '
         Screen.MousePointer = 11
         FIN& = ULTREG&("MOVISTO")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("MOVISTO", U&)
           If TRIM$(Mid$(XX$, 21, 2)) = "RT" Then
             YY$ = TRIM$(Mid$(XX$, 23, 10))
             NREX& = Val(Mid$(YY$, 4))
             If NREX& >= 9554 Then
               If NREX& <= 9599 Then
                 XX$ = String$(256, 0)
                 Call GRAREG("MOVISTO", XX$, U&)
               End If
             End If
           End If
         Next U&
         '
         FIN& = ULTREG&("MOVIREPA")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("MOVIREPA", U&)
           If TRIM$(Mid$(XX$, 21, 2)) = "RP" Then
             YY$ = TRIM$(Mid$(XX$, 23, 10))
             NREX& = Val(Mid$(YY$, 4))
             If NREX& >= 9554 Then
               If NREX& <= 9599 Then
                 XX$ = String$(256, 0)
                 Call GRAREG("MOVIREPA", XX$, U&)
               End If
             End If
           End If
         Next U&
         '
         FIN& = ULTREG&("MOVISTK")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("MOVISTK", U&)
           Call REGAPP("MOVISTO", XX$)
           Call REPLA(XX$, "RP", 21, 2)
           Call REPLA(XX$, "PL.       ", 23, 10)
           Call REPLA(XX$, "Packing-List " + Mid$(XX$, 43, 30), 43, 30)
           Call REPLA(XX$, Mid$(XX$, 96, 8), 88, 8)
           Call REPLA(XX$, String$(8, 0), 96, 8)
           Call REGAPP("MOVISTO", XX$)
         Next U&
         Call CIERRARCH("MOVISTO")
         Call CIERRARCH("MOVISTK")
         '
         FIN& = ULTREG&("PDOCLST")
         NREY& = 9553
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("PDOCLST", U&)
           YY$ = Mid$(XX$, 5)
           If Left$(YY$, 20) = "Remito Nro.0001-0000" Then
             NREX& = Val(Mid$(YY$, 21, 4))
             If NREX& >= 9554 Then
               If NREX& <= 9599 Then
                 NREY& = NREY& + 1
                 ZZ$ = TRIM$(Str$(NREY&))
                 Call REPLA(XX$, ZZ$, 25, 4)
                 Call GRAREG("PDOCLST", XX$, U&)
               End If
             End If
           End If
         Next U&
         Call CIERRARCH("PDOCLST")
             
         Screen.MousePointer = 1
         '
         Call LIBARCH("MOVISTO")
         Call COMUNI("Recuperacion completada")
         '
       End If
     End If
   End If
   '
   Command11.Enabled = True
End Sub

Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'If Button = 2 Then
   '  If Shift = 1 Then
   '    Call RECUMOVI
   '  End If
   'End If
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
      EPAC$ = UCase$(TRIM$(EMPREAC$))
      If InStr(EPAC$, "SABO") > 0 Then
        If EXISTE%(LUDAT$ + "TAOBSOLF.DAT") > 0 Then
          If COMALTER%("Puede Importar Tabla de Obsoletos\\No, Continuar\Importar Obsoletos") = 2 Then
            Call IMPOROBSOL
            GoTo 99
          End If
        End If
      End If
      Call DESCRAP
    End If
99  Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    Call RECPROP
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If DERACCE%("ANUMOSTK", "Anulación de Movimientos de Stock") = 2 Then
       MOVANU$ = TRIM$(InputBox$("Comprobante", "Número de Comprobante a Anular"))
       If MOVANU$ <> "" Then
          Dim CIMO%(32766)
          Screen.MousePointer = 11
          FIN& = ULTREG&("MOVISTO")
          yaprese% = 0
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            XX$ = REGLEIDO$("MOVISTO", U&)
            If TRIM$(Mid$(XX$, 23, 10)) = MOVANU$ Then
              If yaprese% < 1 Then
                FEX = CVI(Left$(XX$, 2))
                TTXX$ = "Confirma Anular TODOS los Movimientos\Asociados al Comprobante '" + MOVANU$ + "' del " + FECHATEX$(FEX) + " ? "
                If COMALTER%(TTXX$ + "\\No, Conservar\Si, Anular") <> 2 Then GoTo 99
                yaprese% = 1
              End If
              CIMO%(CVI(Mid$(XX$, 3, 2))) = 1
              Call GRAREG("MOVISTO", String$(128, 0), U&)
            End If
          Next U&
          Call CIERRARCH("MOVISTO")
          '
          If yaprese% < 1 Then
            Call MENSERR(24, "Comprobante no Encontrado")
            GoTo 99
          End If
          '
          FIN& = NUMAS%
          For CI1% = 1 To NUMAS%
            U& = CI1%
            Call TRACE(20, U&, FIN&)
            If CIMO%(CI1%) > 0 Then
               Call VERISAL(CI1%)
            End If
          Next CI1%
          Call COMUNI("Anulación Completa")
       End If
    End If
99  Command2.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    If DERACCE%("VALEMAT", "Vale de Consumo de Materiales") >= 2 Then
       Call VALEMAT
    End If
    Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    Call ENTRECONSI(0)
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO02"
    If EXISTE%("REMITO02.EXE") < 1 Then
       MODU$ = "REMSTAN"
    End If
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("TRASTOCK", "Transferencias de Stock") >= 2 Then
      TRANSTOK.Show 1
   End If
   Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTOCK.Show 1
    End If
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
   Call CONECRUN("FLSETUP/OPTRAJUST", "Configuración de Funcionamiento")
   Call FINAL("")
   'OPTRAJUST.Show 1
End Sub

Private Sub Command28_Click()
    Command28.Visible = False
    ACONEC$ = "RECEPMA"
    If EXISTE%("BRECEP02.EXE") > 0 Then ACONEC$ = "BRECEP02"
    If EXISTE%("BRECEP03.EXE") > 0 Then ACONEC$ = "BRECEP03"
    Call CONECRUN(ACONEC$, "Recepcion de Materiales")
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If EXISTE%("RETRAS02.EXE") > 0 Then
        Call CONECRUN("RETRAS02", "Emision de Remito de Traslado")
      Else
        Call CONECRUN("REMITRAS", "Emision de Remito de Traslado")
    End If
    Command8.Enabled = True
End Sub

Private Sub Form_Activate()
    TEXPE$ = "Remito de Ventas - Clientes"
    If CONTROL$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Command23.ToolTipText = TEXPE$
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
    Call GRATITFOR(Caption)
    '
    If EXISTE%("REMITO02.EXE") > 0 Or EXISTE%("REMSTAN.EXE") > 0 Then Command23.Enabled = True
    If EXISTE%("ORCOMP02.EXE") > 0 Or EXISTE%("OCOMPRA.EXE") > 0 Then Command22.Enabled = True
    If EXISTE%("RETRAS02.EXE") > 0 Or EXISTE%("REMITRAS.EXE") > 0 Then Command8.Enabled = True
    If EXISTE%("LISEMPAQ.EXE") > 0 Then Command10.Enabled = True
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

Sub IMPOROBSOL()
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
        If InStr(EPAC$, "DOSIVA") > 0 Then
          DEPORIG% = 91
        End If
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(CONTROL$("", "CODSCRAP"))
        If CMO$ = "" Then CMO$ = "SC"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
          If U& < 0.9 * FIN& Then GoTo 1002
          Call TRACE(20, FIN& - U& + 1, FIN&)
          XX$ = REGLEIDO$("MOVISTO", U&)
          If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
            UNRO& = Val(Mid$(XX$, 26, 7))
            GoTo 1002
          End If
        Next U&
        UNRO& = 0
        '
1002    NROVAL& = UNRO& + 1
        Call FRESHALL
        '
1105    Call CIERRARCH("*.*")
        '
1111    HOII% = HOY(HOS$)
        VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DEPORIG%) + Chr$(0) + AJUSTI$("BAJA OBSOLETOS AL 31/12/02", 32) + USERNAME$
        PEDY$ = OBJPLA$("DESCARSCRAP", VDEF$)
        If PEDY$ = "" Then GoTo 9999
        FF% = CVI(Mid$(PEDY$, 5, 2))
        DEPORIG% = Asc(Mid$(PEDY$, 7, 1))
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        REFE$ = Mid$(PEDY$, 9, 32)
        If TRIM$(REFE$) = "" Then
           Call MENSERR(24, "Referencia o Motivo es Mandatoria")
           GoTo 1111
        End If
        If DEPORIG% < 1 Or DEPORIG% >= 99 Or TRIM$(DEPOSI$) = "" Then
           Call MENSERR(24, "Depósito de Salida no Válido")
           GoTo 1111
        End If
        '
1112    Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 11
        '
        If EXISTE%(LUDAT$ + "TAOBSOLF.DAT") > 0 Then
          FIN& = ULTREG&("TAOBSOLF")
          JKL& = 0
          For U& = 1 To FIN&
            XX$ = REGLEIDO$("TAOBSOLF", U&)
            CIXI% = CVI(Left$(XX$, 2))
            CAOB# = CVD(Mid$(XX$, 3, 8))
            YY$ = REGBLAN$("BOLREDET")
            Call REPLA(YY$, MKI$(CIXI%), 83, 2)
            Call REPLA(YY$, MKD$(CAOB#), 89, 8)
            Call REPLA(YY$, UBISTO$(CIXI%), 135, 16)
            '
            If CAOB# > STOCKACT(CIXI%) Then
              Call MENSERR(24, "Atención: Cantidad Obsoleta Mayor a Stock Actual\en Código " + TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%))
              Call REPLA(YY$, "St.Act." + TRIM$(Str$(STOCKACT(CIXI%))), 119, 16)
            End If
            '
            JKL& = JKL& + 1
            Call GRAREG("!BOLREDET", YY$, JKL&)
            '
          Next U&
          Call SETULTREG("!BOLREDET", JKL&)
          Call CIERRARCH("!BOLREDET")
        End If
        '
        FFF$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
1113    VTB% = VENTABU%("DESCRAP/NC/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + FFF$)
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            GoTo 9999
        End If
        '
        OPXX% = COMALTER%("Opciones de Trabajo\\Imprimir Listado\Procesar")
        If OPXX% = 1 Then
            Call TRACENTRAL("BOLREDET", "")
            Call FINAL("*LIOBSOL")
            GoTo 1113
          ElseIf OPXX% = 2 Then
            GoTo 11141
          Else
            GoTo 9999
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          Call CIERRARCH("MOVISTO")
          Call BLOQARCH("MOVISTO")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             LOTE$ = Mid$(YY$, 69, 12)
             REMI$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    DOSEC$ = ""
1120                CANTI = (-1) * CAN#
                    Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, REFE$, 0, "$", 0, DEPORIG%, CANTI)
                 End If
               End If
             End If
             '
          Next YX&
        End If
        '
        Call CIERRARCH("MOVISTO")
        '
        Screen.MousePointer = 1
        '
9999    Screen.MousePointer = 1
        Call CIERRARCH("*.*")
        '
End Sub



VERSION 5.00
Begin VB.Form MOVISTBAS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5955
   Icon            =   "MOVISTBAS.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4890
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   9
      Top             =   2600
      Width           =   965
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   10
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
      Height          =   750
      Left            =   4960
      Picture         =   "MOVISTBAS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Picture         =   "MOVISTBAS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Height          =   3270
      Left            =   210
      TabIndex        =   3
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEC0&
         Caption         =   "Movimientos Varios"
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
         TabIndex        =   5
         Top             =   1830
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   12
            Top             =   420
            Width           =   1695
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
               Picture         =   "MOVISTBAS.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes de Stock"
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
               TabIndex        =   15
               Top             =   105
               Width           =   795
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   8
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command8 
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
               Picture         =   "MOVISTBAS.frx":08A8
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Remito de Traslado"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Despacho Mercaderia"
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
               Width           =   1320
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
         TabIndex        =   4
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   6
            Top             =   420
            Width           =   2115
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
               Picture         =   "MOVISTBAS.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción de Materiales"
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
               TabIndex        =   7
               Top             =   105
               Width           =   1320
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4380
      Picture         =   "MOVISTBAS.frx":0EBC
      Top             =   2840
      Width           =   2010
   End
End
Attribute VB_Name = "MOVISTBAS"
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


Private Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
       Call DESCRAP
    End If
    Command17.Enabled = True
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
      Call TRASTOCK
   End If
   Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTO00.Show 1
    End If
    Command25.Enabled = True
End Sub


Private Sub Command28_Click()
    Command28.Visible = False
    Call RECECOMO
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   AJUSTOBAS.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    FORECOMO.Show 1
    Command8.Enabled = True
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

Sub RECECOMO() ' recepcion de materiales Comodoro
        '
   'Call BLIMESS("Seleccionando ...")
   FIN& = ULTREG&("ATRASLA")
   Call BLANARCH("!ATRASLA")
   CABOPE% = 0
   FF% = HOY(FECHAREC$)
   '
   For I& = 1 To FIN&
     X$ = REGLEIDO$("ATRASLA", I&)
     If CVD(Mid$(X$, 105, 8)) >= 0.05 Then
       If CVD(Mid$(X$, 117, 8)) < 0.05 Then
         CANOM# = CVD(Mid$(X$, 105, 8))
         Call REPLA(X$, MKD$(CANOM#), 117, 8)
         Call REPLA(X$, MKS$(I&), 125, 4) ' registro
         Call REGAPP("!ATRASLA", X$)
       End If
     End If
   Next I&
   '
   If ULTREG&("!ATRASLA") < 1 Then
     Call COMUNI("No hay Recepciones\Pendientes de Ingreso.")
     Exit Sub
   End If
   '
   FERECEP$ = REPLACAR$(FECHAREC$, "/", "-")
   VTB% = VENTABU%("RECECOM/NC/TITUPAN=Boleta de Recepción - " + FERECEP$)
   '
   If VTB% >= 0 Then
     If COMALTER%("Confirme Ingreso de Materiales Recibidos\\Cancelar\Confirmar Ingreso") = 2 Then
       AINDI% = 0
       For U& = 1 To ULTREG&("!ATRASLA")
         X$ = REGLEIDO$("!ATRASLA", U&)
         APRO# = CVD(Mid$(X$, 117, 8))
         If APRO# > 0.05 Then
           Call REPLA(X$, MKI$(FF%), 115, 2)
           REGORI& = CVS(Mid$(X$, 125, 4))
           If REGORI& > 0 Then
             Y$ = REGLEIDO$("ATRASLA", REGORI&)
           End If
           If REGORI& < 1 Or Left$(Y$, 112) <> Left$(X$, 112) Then
               Call REGAPP("ATRASLA", X$)
             Else
               Call GRAREG("ATRASLA", X$, REGORI&)
           End If
           '
           CI% = CVI(Mid$(X$, 39, 2))
           REFE$ = "Recepción " + FECHAREC$
           AINDI% = AINDI% + 1
           LOTE$ = TRIM$(Mid$(X$, 3, 14)) + "-" + TRIM$(Str$(AINDI%))
           DOPRI$ = Mid$(X$, 3, 14)
           DOSEC$ = "BR." + FECHAREC$
           DEPX% = 1
           CANTI = APRO#
           Call MOVISTO(FF%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
         End If
       Next U&
     End If
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub


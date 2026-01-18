VERSION 5.00
Begin VB.Form LIMOSTK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos - Consultas y Listados"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9600
   Icon            =   "LIMOSTK07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000080&
      Height          =   7800
      Left            =   8715
      ScaleHeight     =   7740
      ScaleWidth      =   1005
      TabIndex        =   14
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "LIMOSTK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda en Línea"
         Top             =   850
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "LIMOSTK07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "LIMOSTK07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -255
         Picture         =   "LIMOSTK07.frx":0BA0
         Stretch         =   -1  'True
         Top             =   3700
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
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
      Height          =   4020
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4155
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Ajustes / Transferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   250
         TabIndex        =   6
         Top             =   2220
         Width           =   3615
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2550
            TabIndex        =   7
            Top             =   600
            Width           =   2615
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
               Left            =   2000
               Picture         =   "LIMOSTK07.frx":2AC2
               Style           =   1  'Graphical
               TabIndex        =   31
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
               Picture         =   "LIMOSTK07.frx":2DCC
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos Varios"
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
               TabIndex        =   13
               Top             =   105
               Width           =   1425
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Entradas / Salidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   250
         TabIndex        =   9
         Top             =   420
         Width           =   3615
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2550
            TabIndex        =   10
            Top             =   600
            Width           =   2615
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
               Left            =   2000
               Picture         =   "LIMOSTK07.frx":2F16
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Vale de Consumo de Materiales"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "LIMOSTK07.frx":3060
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Recepción de Proveedores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos Externos"
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
               Width           =   1275
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Control de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   4155
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0FF&
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
         Height          =   735
         Left            =   240
         TabIndex        =   27
         Top             =   3150
         Width           =   3660
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0FF&
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
            Left            =   480
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0FF&
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
            Left            =   2040
            TabIndex        =   28
            Top             =   360
            Width           =   1220
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Consultas / Listados Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         Left            =   250
         TabIndex        =   2
         Top             =   420
         Width           =   3615
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   3
            Top             =   1820
            Width           =   2300
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
               Picture         =   "LIMOSTK07.frx":336A
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Gestion de Reservas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Gestión de Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   840
               TabIndex        =   4
               Top             =   105
               Width           =   1740
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   21
            Top             =   1340
            Width           =   2300
            Begin VB.CommandButton Command30 
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
               Picture         =   "LIMOSTK07.frx":3674
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Listado de Existencias Negativas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Negativos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   840
               TabIndex        =   23
               Top             =   45
               Width           =   1740
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   24
            Top             =   880
            Width           =   2300
            Begin VB.CommandButton Command20 
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
               Picture         =   "LIMOSTK07.frx":397E
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Items con Stock Sub-Minimo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   735
               TabIndex        =   25
               Top             =   45
               Width           =   1740
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2235
            TabIndex        =   18
            Top             =   420
            Width           =   2300
            Begin VB.CommandButton Command19 
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
               Picture         =   "LIMOSTK07.frx":3C88
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Existencias Físicas Actuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias FÍsi- cas Actuales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   735
               TabIndex        =   19
               Top             =   45
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "LIMOSTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GEN%
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub Command14_Click()
   Command14.Enabled = False
   '
   If Option5 Then 'PANTALLA
     GEN% = 0
     Call GENRESDAT
     If GEN% > 0 Then
        Call CONECRUN("*CORESE3", "")
     End If
   ElseIf Option4 Then 'IMPRESION
     GEN% = 0
     Call GENRESDAT
     'PARA QUE NO SE EJECUTE LA CONSULTA
     'CON EL ARCHIVO VACIO.
     If GEN% > 0 Then
        'Call CONECRUN("*CORESE2", "")
        Call TOMATRES
     End If
   End If
   '
   Command14.Enabled = True
End Sub

Sub Command19_Click()
  Command19.Enabled = False
    Call EXISTO2
  Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call LIUBIEXI
    Command2.Enabled = True
End Sub

 Sub Command20_Click()
    Command20.Enabled = False
    Call CRITSTO2
    If ULTREG&("CRITSTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*CRITSTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*CRITSTL", "")
        End If
    End If
    Command20.Enabled = True
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    Call ACCDIR("CONSTO00")
    Command27.Enabled = True
End Sub

 Sub Command21_Click()
    Command21.Enabled = False
    If DERACCE%("VALEMAT", "Vale de Consumo de Materiales") >= 2 Then
       Call VALEMAT
    End If
    Command21.Enabled = True
End Sub

Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("TRASTOCK", "Transferencias de Stock") >= 2 Then
      TRANSTK07.Command4.Visible = False
      TRANSTK07.Show 1
   End If
   Command24.Enabled = True
End Sub

Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUMAQ07.Show 1
    End If
    Command25.Enabled = True
End Sub

Sub Command28_Click()
  '
  FOREMAQ07.Show 1
  '
End Sub

Private Sub Command29_Click()
  '
  Call HELPONLINE("LISMOSKT")
  '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
      OPMOMAQ07.Show 1
99  Command3.Enabled = True
    '
End Sub
Private Sub Command5_Click()
    Command5.Enabled = False
    DISPOSTO.Show 1
    Command5.Enabled = True
End Sub

 Sub Command30_Click()
    Command30.Enabled = False
    Call NEGASTO2
    If ULTREG&("NEGASTO") > 0 Then
      Call CIERRARCH("*.*")
      If Option5.Value = True Then
           Call CONECRUN("*NEGASTS", "")
         ElseIf Option4.Value = True Then
           Call CONECRUN("*NEGASTL", "")
      End If
    End If
    Command30.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    ACONEC$ = "INVENT02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "INVENTAR"
    End If
    Call CONECRUN(ACONEC$, "Inventarios Físicos y Valorizados")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    OPRESERVA.Show 1
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX", "Ficha Kardex de Movimientos")
      Else
        'Call GENKARDEX("")
    End If
    Command9.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
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
   Call BAJALDISCO
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
'
'Sub IMPORNEU2()
'   '
'   If ULTREG&("MOVISTO") < 1 Then GoTo 3
'   LU$ = LUDAT$
'   If EXISTE%(LU$ + "MOVISTX.DAT") = 1 Then
'      Screen.MousePointer = 11
'      Call COPYSTRU("MOVISTO", "MOVISTX")
'      FIN& = ULTREG&("MOVISTX")
'      For U& = 1 To FIN&
'         XXX$ = REGLEIDO$("MOVISTX", U&)
'         CIXI% = CVI(Mid$(XXX$, 3, 2))
'         Call REGAPP("MOVISTO", XXX$)
'         Call CIERRARCH("MOVISTO")
'         Call VERISAL(CIXI%)
'         Call VESARRAS(CIXI%)
'      Next U&
'   End If
'   Call CIERRARCH("*.*")
'   Screen.MousePointer = 1
'   Exit Sub
'   '
'3  Screen.MousePointer = 11
'   Call BLANARCH("MOVISTO")
'   '
'   Dim SALSTO(16384)
'   Dim CODEPOS$(64)
'   Dim ALIDEP$(128), CODEPX$(128)
'   '
'   ALIDEP$(1) = "FK": CODEPX$(1) = "5110000"
'   ALIDEP$(2) = "DEP.": CODEPX$(2) = "4080000"
'   ALIDEP$(3) = "BS": CODEPX$(3) = "5120000"
'   ALIDEP$(4) = "510": CODEPX$(4) = "5100000"
'   ALIDEP$(5) = "513": CODEPX$(5) = "5130000"
'   ALIDEP$(6) = "408000": CODEPX$(6) = "4080000"
'   CALI% = 6
'   '
'   CADEPO% = 0
'   For U& = 1 To 64
'      If U& <= ULTREG&("SUCURSAL") Then
'         CODEPOS$(U&) = TRIM$(Mid$(REGLEIDO$("SUCURSAL", U&), 21, 8))
'         CADEPO% = U&
'      End If
'   Next U&
'   '
'   'For U& = 1 To ULTREG&("MOVISTO")
'   '   X$ = REGLEIDO$("MOVISTO", U&)
'   '   CI% = CVI(Mid$(X$, 3, 2))
'   '   CAN = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
'   '   If CI% > 0 Then
'   '      If CI% <= 16384 Then
'   '         SALSTO(CI%) = SALSTO(CI%) + CAN
'   '      End If
'   '   End If
'19 'Next U&
'   '
'   GoSub 95
'   '
'   HOII% = HOY(HO$)
'   NRO& = 0
'   CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
'   If CMOV$ = "" Then CMOV$ = "AJ"
'   '
'   ARIMPO$ = "A:\MULTIST"
'   PRIBOT% = 0
'   NX% = FILEOPEN%(ARIMPO$, 1, 0)
'   While Not EOF(NX%)
'      Line Input #NX%, A$
'      If Left$(A$, 3) = "BOT" Then
'         If PRIBOT% = 0 Then
'              PRIBOT% = 1
'           Else
'              For KK% = 1 To 2
'                 Line Input #NX%, A$
'              Next KK%
'              cod1$ = TRIM$(Mid$(A$, 2))
'              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
'              For KK% = 3 To 6
'                 Line Input #NX%, A$
'              Next KK%
'              CODEPO1$ = TRIM$(Mid$(A$, 2))
'              CODEPO$ = TRIM$(REPLACAR$(CODEPO1$, Chr$(34), " "))
'              For KK% = 1 To CALI%
'                 If CODEPO$ = ALIDEP$(KK%) Then
'                    CODEPO$ = CODEPX$(KK%)
'                    GoTo 5
'                 End If
'              Next KK%
'              '
'5             For KK% = 1 To CADEPO%
'                 If CODEPOS$(KK%) = CODEPO$ Then
'                    IDEPO% = KK%
'                    GoTo 10
'                 End If
'              Next KK%
'              CODEPO$ = CODEPOS$(1)
'              IDEPO% = 1
'              '
'10            For KK% = 7 To 20
'                 Line Input #NX%, A$
'              Next KK%
'              CANTI = Val(Mid$(A$, 2))
'              GoSub 90
'         End If
'      End If
'   Wend
'   '
'   Call CIERRARCH("*.*")
'   Screen.MousePointer = 1
'   Exit Sub
'   '
'90 CI% = CODINT%(COD$)
'   If CI% > 0 Then
'      If Abs(CANTI) > 0.05 Then
'         If CI% <= 16384 Then
'            SALSTO(CI%) = SALSTO(CI%) + CANTI
'         End If
'         NRO& = NRO& + 1
'         DOSEC$ = "AJ." + Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
'         DOPRI$ = "AS." + TRIM$(Str$(NRO&))
'         VUNI# = 0
'         MONE$ = ""
'         NC% = 0
'         REFE$ = "ST.IMPORT.AL " + HO$
'         LTX$ = "ST.INICIAL"
'         DEPO% = IDEPO%
'         FF% = HOII%
'         '
'         Call Movisto(FF%, CI%, LTX$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
'         '
'      End If
'   End If
'   Return
'   '
'95 For CI% = 1 To NUMAS%
'      U& = CI% + 1
'      REBLA$ = String$(36, 0) + MKS$(0.00001)
'      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
'      Call GRAREG("SALSTOCK", REBLA$, U&)
'      REBLA$ = String$(252, 0) + MKS$(0.00001)
'      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
'      Call GRAREG("SALDEPOS", REBLA$, U&)
'   Next CI%
'   Return
'   '
'End Sub
'
''
'Sub IMPORNEU3()
'   '
'   Dim SALSTO(16384)
'   Dim CODEPOS$(64)
'   Dim ALIDEP$(128), CODEPX$(128)
'   '
'   If ULTREG&("UBISTOCK") > 0 Then
'      Call MENSERR(24, "Imposible Importar Ubicaciones.\Ya Existe Archivo de Ubicaciones.")
'      Call CIERRARCH("*.*")
'      GoTo 50
'   End If
'   '
'   ARIMPO$ = "A:\UBIC"
'   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
'   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
'      Call CIERRARCH("*.*")
'      Exit Sub
'   End If
'   '
'   Screen.MousePointer = 11
'   PRIBOT% = 0
'   NX% = FILEOPEN%(ARIMPO$, 1, 0)
'   While Not EOF(NX%)
'      Line Input #NX%, A$
'      If Left$(A$, 3) = "BOT" Then
'         If PRIBOT% = 0 Then
'              PRIBOT% = 1
'           Else
'              For KK% = 1 To 2
'                 Line Input #NX%, A$
'              Next KK%
'              cod1$ = TRIM$(Mid$(A$, 2))
'              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
'              For KK% = 3 To 4
'                 Line Input #NX%, A$
'              Next KK%
'              UBICA1$ = TRIM$(Mid$(A$, 2))
'              UBICA$ = TRIM$(REPLACAR$(UBICA1$, Chr$(34), " "))
'              '
'10            For KK% = 5 To 6
'                 Line Input #NX%, A$
'              Next KK%
'              LOECON = Val(Mid$(A$, 2))
'              '
'20            For KK% = 7 To 8
'                 Line Input #NX%, A$
'              Next KK%
'              PUNPED = Val(Mid$(A$, 2))
'              '
'              GoSub 90
'         End If
'      End If
'   Wend
'   '
'50 If ULTREG&("MASTSPEC") > 0 Then
'      Call MENSERR(24, "Imposible Importar Especificaciones.\Ya Existe Archivo de Especificaciones.")
'      Call CIERRARCH("*.*")
'      GoTo 80
'   End If
'   '
'   ARIMPO$ = "A:\ESPECIF"
'   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
'   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
'      Call CIERRARCH("*.*")
'      Exit Sub
'   End If
'   '
'   PRIBOT% = 0
'   Screen.MousePointer = 11
'   NX% = FILEOPEN%(ARIMPO$, 1, 0)
'   While Not EOF(NX%)
'55    Line Input #NX%, A$
'      If Left$(A$, 3) = "BOT" Then
'         If PRIBOT% = 0 Then
'              PRIBOT% = 1
'           Else
'              For KK% = 1 To 2
'                 Line Input #NX%, A$
'              Next KK%
'              cod1$ = TRIM$(Mid$(A$, 2))
'              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
'              Line Input #NX%, A$
'60            Line Input #NX%, TTXX$
'              If Left$(TTXX$, 3) = "-1," Then GoTo 55
'              GoSub 95
'              GoTo 60
'         End If
'      End If
'   Wend
'   '
'80 Call CIERRARCH("*.*")
'   Screen.MousePointer = 1
'   Exit Sub
'   '
'90 CI% = CODINT%(COD$)
'   If CI% > 0 Then
'      If TRIM$(UBICA$) <> "" Then
'         '
'         XX$ = REGBLAN$("UBISTOCK")
'         Call REPLA(XX$, MKI$(CI%), 1, 2)
'         Call REPLA(XX$, UBICA$, 3, 24)
'         Call REGAPP("UBISTOCK", XX$)
'         '
'      End If
'      '
'      REGMA& = CI% + 1
'      YY$ = REGLEIDO$("MASTER", REGMA&)
'      If PUNPED > 0.05 Then
'         STMI& = PUNPED
'         If STMI& > 30000 Then STMI& = (-1) * PUNPED / 100
'         stm1% = STMI&
'         Call REPLA(YY$, MKI$(stm1%), 105, 2)
'      End If
'      If LOECON > 0.05 Then
'         LOECO& = LOECON
'         If LOECO& > 30000 Then LOECO& = (-1) * LOECON / 100
'         lec1% = LOECO&
'         Call REPLA(YY$, MKI$(lec1%), 107, 2)
'      End If
'      Call GRAREG("MASTER", YY$, REGMA&)
'      '
'   End If
'   Return
'   '
'95 TTYY$ = TRIM$(REPLACAR$(TTXX$, Chr$(34), " "))
'   If TRIM$(TTYY$) <> "" Then
'      CI% = CODINT%(COD$)
'      If CI% > 0 Then
'         Call FRATEXTO(TTYY$, 60)
'         For KI% = 1 To CARETEX%
'            XX$ = REGBLAN$("MASTSPEC")
'            Call REPLA(XX$, MKI$(CI%), 1, 2)
'            Call REPLA(XX$, RETEX$(KI%), 3, 78)
'            Call REGAPP("MASTSPEC", XX$)
'         Next KI%
'      End If
'   End If
'   Return
'   '
'End Sub
Sub EXISTO2()
    '
    TMTX% = 0: DPSX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%) + Chr$(DPSX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD4", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
       TMT% = 0
       Call REPLA(RACOD$, Chr$(0), 5, 1)
       TIMAT$ = "Material No Clasificado"
    End If
    DPS% = Asc(Mid$(RACOD$, 6, 1))
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Call HEADERS("NEGASTO", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, URE&)
      CIXI% = I&
      CDDX$ = CODEXT$(CIXI%)
      If CDDX$ >= COD1$ Then
        If CDDX$ <= COD2$ Then
          TMT1% = TIMATE%(CIXI%)
          If TMT% < 1 Or TMT1% = TMT% Then
            '
            CII& = CIXI%
            Call VESARRAS(CIXI%)
            'SSTT = STOCKACT(CIXI%)
            SSTT = STODEPOS(CIXI%, DPS%)
            If Abs(SSTT) > 0.05 Then
              TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
              J& = J& + 1
              Call GRAREG("NEGASTO", TTN$, J&)
            End If
            '
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call FILESORT("NEGASTO", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        If Option5.Value = True Then
            Call CONECRUN("*EXISTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*EXISTL", "")
        End If
    End If
    '
End Sub
'\\\OT-05-0379-WAR-FIN///
'
Sub NEGASTO2()    'DETERMINACION ITEMS CON STOCK NEGATIVO
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    '
    FIN& = NUMAS%
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      SSTT = STOCKACT(CI%)
      If SSTT < 0 Then
        TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
        J& = J& + 1
        Call GRAREG("NEGASTO", TTN$, J&)
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se detectaron items con stock negativo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Negativos.")
    End If
    '
End Sub
'

Sub CRITSTO2()    'DETERMINACION ITEMS CON STOCK SUB MINIMO
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    FIN& = NUMAS%
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      SSTT = STOCKACT(CI%)
      '
      STMIN = STOMIN(CI%)
      If SSTT < STMIN Then
        '
        COU = COSUNI(CI%)
        REPOS = STMIN - SSTT
        If REPOS < LOTESTAN(CI%) Then REPOS = LOTESTAN(CI%)
        '
        TTN$ = REGBLAN$("CRITSTO")
        Call REPLA(TTN$, MKI$(CI%), 1, 2)
        Call REPLA(TTN$, UMED$, 3, 2)
        Call REPLA(TTN$, MKS$(SSTT), 5, 4)
        Call REPLA(TTN$, MKS$(STMIN), 9, 4)
        Call REPLA(TTN$, MKS$(STMIN - SSTT), 13, 4)
        Call REPLA(TTN$, MKS$(REPOS), 17, 4)
        Call REPLA(TTN$, MKS$(COU), 21, 4)
        Call REPLA(TTN$, MKD$(COU * REPOS), 25, 8)
        J& = J& + 1
        Call GRAREG("CRITSTO", TTN$, J&)
        '
      End If
    Next I&
    '
    Call SETULTREG("CRITSTO", J&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se detectaron items con stock sub-mínimo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Sub-Mínimo.")
    End If
    '
End Sub
Sub LIUBIEXI()
    '
    Dim PRIRE&(32767)
    Static NUSU%
    '
    TXT$ = "El Listado de Artículos y Ubicaciones Existente\"
    FEX = FEGEAR%("UBIEXI")
    If FEX < 1 Then GoTo 310
    TXT$ = TXT$ + "Corresponde al " + FECHATEX$(FEX) + " a las " + HORATEX$(HOGEAR%("UBIEXI")) + " hs.\"
    TXT$ = TXT$ + "\Consultar\Re-Generar\Abandonar"
    X% = COMALTER%(TXT$)
    On X% GoTo 320, 310
    Exit Sub
    '
310 Screen.MousePointer = 11
    NMS% = NUMAS%
    '
'    If NUSU% = 0 Then
'        NUSU% = 1
'        For I& = 1 To ULTREG&("SUCURSAL")
'            XXX$ = REGLEIDO$("SUCURSAL", I&)
'            If I& <= 62 Then
'                If Asc(Left$(XXX$, 1)) > 32 Then
'                    NUSU% = I&
'                End If
'            End If
'        Next I&
'        Call CIERRARCH("SUCURSAL")
'    End If
'    '
    J& = 0
    FIN& = ULTREG&("UBISTOCK")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("UBISTOCK", I&)
      CI% = CVI(Left$(XX$, 2))
      UBI$ = AJUSTI$(Mid$(XX$, 3, 24), 24) + AJUSTI$(Mid$(XX$, 31, 30), 30)
      If TRIM$(UBI$) <> "" Then
        If CI% > 0 Then
          If CI% <= NMS% Then
            YY$ = REGBLAN$("UBIEXI")
            Call REPLA(YY$, MKI$(CI%), 1, 2)
            Call REPLA(YY$, UBI$, 7, 54)
            J& = J& + 1
            Call GRAREG("UBIEXI", YY$, J&)
            If PRIRE&(CI%) < 1 Then
              PRIRE&(CI%) = J&
            End If
          End If
        End If
      End If
    Next I&
    '
311 FIN& = NUMAS%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      CAN = STOCKACT(CI%)
      '
      If Abs(CAN) >= 0.05 Then
        If PRIRE&(CI%) > 0 Then
            YY$ = REGLEIDO$("UBIEXI", PRIRE&(CI%))
            Call REPLA(YY$, MKS$(CAN), 3, 4)
            Call GRAREG("UBIEXI", YY$, PRIRE&(CI%))
          Else
            YY$ = REGBLAN$("UBIEXI")
            Call REPLA(YY$, MKI$(CI%), 1, 2)
            Call REPLA(YY$, MKS$(CAN), 3, 4)
            J& = J& + 1
            Call GRAREG("UBIEXI", YY$, J&)
            PRIRE&(CI%) = J&
        End If
      End If
    Next I&
    '
    Call SETULTREG("UBIEXI", J&)
    Call CIERRARCH("UBIEXI")
    Call GRARGEN("UBIEXI", "")
    '
320 Screen.MousePointer = 1
    FIN& = ULTREG&("UBIEXI")
    If FIN& < 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
    End If
    '
330 VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD", VDEF$)
    If OBX$ <> "" Then
      '
      Screen.MousePointer = 11
      CI1% = CVI(Left$(OBX$, 2))
      COD1$ = CODEXT$(CI1%)
      CI2% = CVI(Mid$(OBX$, 3, 2))
      COD2$ = CODEXT$(CI2%)
      '
      J& = 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        YY$ = REGLEIDO$("UBIEXI", U&)
        CIII% = CVI(Left$(YY$, 2))
        CODX$ = CODEXT$(CIII%)
        If CODX$ >= COD1$ Then
          If CODX$ <= COD2$ Then
            J& = J& + 1
            Call GRAREG("LUBIEXI", YY$, J&)
          End If
        End If
      Next U&
      Call SETULTREG("LUBIEXI", J&)
      Call CIERRARCH("LUBIEXI")
      Call CIERRARCH("*.*")
      Screen.MousePointer = 1
      Call FINAL("*SUBIEXI")
    End If
    '
End Sub
'

Sub GENRESDAT()
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RESERTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        GEN% = 1
        '
        Call REPLA(Y$, MKI$(!Cod_Inte), 1, 2)
        Call REPLA(Y$, !IdSubOr, 3, 12)
        Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(!CantRes), 19, 8)
        Call REPLA(Y$, MKD$(!CantCons), 27, 8)
        Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
        Call REPLA(Y$, !IdenLote, 43, 16)
        Call REPLA(Y$, Chr$(!Deposito), 60, 1)
        Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("RESERVA", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("RESERVA", JJ&)
    Call CIERRARCH("RESERVA")
    '
    Call ABREORFAB
    'ACA HAY QUE REVISAR EN VEZ DE ORDEFABR LA BASE DE
    'ORDENES DE MANTENIMIENTO
'    SQLX$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR "
'    SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
'    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
'    '
'    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With ORDEFABTMP
'      On Error Resume Next
'      .MoveFirst
'      If Err Then Call MENSERR(24, "NO Existen Ordenes de Fabricación. Imposible Continuar la Carga.")
'      On Error GoTo 0
'      JJ& = 0
'      Do While (.EOF = False)
'        Y$ = REGBLAN("ECORDEP")
'        '
'        Call REPLA(Y$, !IdSubOr, 1, 12)
'        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
'        '
'        JJ& = JJ& + 1
'        Call GRAREG("ECORDEP", Y$, JJ&)
'        '
'      .MoveNext
'      Loop
'    End With
    Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
End Sub
'
Sub TOMATRES()
    '
    Dim CARESE#(32767)
    FIN& = ULTREG&("RESERVA")
    If FIN& < 1 Then
      Call MENSERR(24, "No hay Reservas Registradas")
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("RESERVA", I&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVD(Mid$(XX$, 19, 8))
          CACO# = CVD(Mid$(XX$, 27, 8))
          SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
          SARE# = TORE# - CACO#
          If SARE# * SIGNO% < 0 Then
              SARE# = 0
          End If
          Call REPLA(XX$, MKD$(SARE#), 35, 8)
          Call GRAREG("RESERVA", XX$, I&)
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
    Next I&
    '
    FIN& = ULTREG&("REPUREPA")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("REPUREPA", I&)
      CI% = CVI(Mid$(XX$, 7, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVS(Mid$(XX$, 21, 4))
          CACO# = CVS(Mid$(XX$, 25, 4))
          '
          If UCase$(UNIMED$(CIIX%)) = "GK" Then
            TORE# = TORE# / 1000
            CACO# = CACO# / 1000
          End If
          '
          SARE# = TORE# - CACO#
          If SARE# < 0.05 Then GoTo 29
          '
          NOREPARA& = CVS(Left$(XX$, 4))
          REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
          If REREP& > 0 Then
            YY$ = REGLEIDO$("PRESREPA", REREP&)
            If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 29 ' O/Rep Cerrada
            If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 29 ' falta O/Venta
          End If
          '
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
29  Next I&
    '
    JJ& = 0
    FIN& = NUMAS%: If FIN& < 1 Then FIN& = 1
    '
    For I& = 1 To NUMAS%
      II$ = REGLEIDO$("INVERT", I& + 1)
      CI% = I&
      If CI% > 0 Then
        If Abs(CARESE#(CI%)) > 0.05 Then
          STOCK# = STOCKACT(CI%)
          XX$ = REGBLAN$("TOMATRES")
          Call REPLA(XX$, MKI$(CI%), 1, 2)
          Call REPLA(XX$, MKD$(CARESE#(CI%)), 3, 8)
          Call REPLA(XX$, MKD$(STOCK#), 11, 8)
          FALTA# = CARESE#(CI%) - STOCK#
          If FALTA# < 0 Then FALTA# = 0
          Call REPLA(XX$, MKD$(FALTA#), 35, 8)
          OBSE$ = "      "
          If FALTA# > 0 Then OBSE$ = "CRIT. "
          If STOCK# - STOMIN(CI%) < 0 Then OBSE$ = "CRIT. "
          If STOCK# < 0 Then OBSE$ = "NEG.  "
          Call REPLA(XX$, OBSE$, 43, 6)
          JJ& = JJ& + 1
          Call GRAREG("TOMATRES", XX$, JJ&)
        End If
      End If
    Next I&
    '
    Call SETULTREG("TOMATRES", JJ&)
    Call CIERRARCH("TOMATRES")
    Call FILESORT("TOMATRES", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    '
    If Option4.Value = True Then
         Call FINAL("*LITOTRE")       ' impresion
       Else
         Call FINAL("*COTOTRE")       ' pantalla
    End If
    '
End Sub
'


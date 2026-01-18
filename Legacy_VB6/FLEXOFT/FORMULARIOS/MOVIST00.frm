VERSION 5.00
Begin VB.Form MOVIST00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5955
   Icon            =   "MOVIST00.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "MOVIST00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1760
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4890
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   16
      Top             =   2600
      Width           =   965
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   17
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
      Picture         =   "MOVIST00.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1010
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
      Picture         =   "MOVIST00.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   3
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
      TabIndex        =   5
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
         TabIndex        =   7
         Top             =   1830
         Width           =   4110
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   620
            ScaleHeight     =   585
            ScaleWidth      =   2790
            TabIndex        =   12
            Top             =   420
            Width           =   2850
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
               Picture         =   "MOVIST00.frx":0BA0
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Remito de Traslado"
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
               Left            =   1680
               Picture         =   "MOVIST00.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Vale de Consumo de Materiales"
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
               Left            =   1120
               Picture         =   "MOVIST00.frx":0FF4
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Transferencias de Stock"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2240
               Picture         =   "MOVIST00.frx":12FE
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Descartes de Stock"
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
               Left            =   560
               Picture         =   "MOVIST00.frx":1608
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
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
         TabIndex        =   6
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   10
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
               Picture         =   "MOVIST00.frx":1752
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
               TabIndex        =   11
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
            TabIndex        =   8
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
               Picture         =   "MOVIST00.frx":1A5C
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
               TabIndex        =   9
               Top             =   105
               Width           =   1110
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4380
      Picture         =   "MOVIST00.frx":1D66
      Top             =   2840
      Width           =   2010
   End
End
Attribute VB_Name = "MOVIST00"
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

Private Sub Command18_Click()
    Call RECPROP
End Sub

Private Sub COMMAND21_CLICK()
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

Private Sub Command26_Click()
    'OPTRAJUST.Show 1
    Call CONECRUN("FLSETUP/OPTRAJUST", "Configuración de Funcionamiento")
    Call FINAL("")
End Sub

Private Sub Command28_Click()
    Command28.Visible = False
    If EXISTE%("BRECEP02.EXE") > 0 Then
         Call CONECRUN("BRECEP02", "Recepcion de Materiales")
       Else
         Call CONECRUN("RECEPMA", "Recepcion de Materiales")
    End If
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    ACONEC$ = "REMITO02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "RETRAS02"
      If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        ACONEC$ = "REMITRAS"
      End If
    End If
    Call CONECRUN(ACONEC$, "Emision de Remito de Traslado")
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


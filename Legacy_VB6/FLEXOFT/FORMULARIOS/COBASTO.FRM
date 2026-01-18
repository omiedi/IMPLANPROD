VERSION 5.00
Begin VB.Form COBASTO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Stock"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control de Expedición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   4830
      TabIndex        =   28
      Top             =   2280
      Width           =   4110
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   35
         Top             =   1680
         Width           =   2220
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
            Picture         =   "COBASTO.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Estadisticas de Ventas y Análisis de Margen Bruto"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resúmenes Esta- dísticos de Venta"
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
            TabIndex        =   37
            Top             =   105
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   32
         Top             =   1155
         Width           =   2220
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
            Picture         =   "COBASTO.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   " "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta Despa- chos p/Factura"
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
            Left            =   630
            TabIndex        =   34
            Top             =   105
            Width           =   1620
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   29
         Top             =   630
         Width           =   2220
         Begin VB.CommandButton Command15 
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
            Picture         =   "COBASTO.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   " Remitos Pendientes de Facturacion - Listado por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Remitos Pendien- tes de Factura"
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
            TabIndex        =   31
            Top             =   105
            Width           =   1515
         End
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Depos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9135
      Picture         =   "COBASTO.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4120
      Width           =   855
   End
   Begin VB.CommandButton Command27 
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
      Height          =   855
      Left            =   9135
      Picture         =   "COBASTO.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2270
      Width           =   855
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
      Height          =   855
      Left            =   9135
      Picture         =   "COBASTO.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3195
      Width           =   855
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
      Height          =   855
      Left            =   9135
      Picture         =   "COBASTO.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1345
      Width           =   855
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
      Height          =   855
      Left            =   9135
      Picture         =   "COBASTO.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1800
      Left            =   4830
      TabIndex        =   16
      Top             =   315
      Width           =   4110
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   22
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command23 
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
            Picture         =   "COBASTO.frx":1608
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Expedicion - Des- pacho por Remito"
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
            TabIndex        =   24
            Top             =   105
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   25
         Top             =   525
         Width           =   2430
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
            Picture         =   "COBASTO.frx":1912
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste Directo de Existencias"
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
            TabIndex        =   27
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   0
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   4635
      Left            =   210
      TabIndex        =   13
      Top             =   315
      Width           =   4425
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
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
         Height          =   2430
         Left            =   210
         TabIndex        =   15
         Top             =   1995
         Width           =   4005
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   315
            ScaleHeight     =   75
            ScaleWidth      =   3315
            TabIndex        =   38
            Top             =   2100
            Width           =   3375
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Salida"
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
            Left            =   315
            TabIndex        =   19
            Top             =   1260
            Width           =   3375
            Begin VB.OptionButton Option4 
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
               Left            =   1890
               TabIndex        =   3
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
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
               Left            =   315
               TabIndex        =   2
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   17
            Top             =   420
            Width           =   3375
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
               Left            =   2200
               Picture         =   "COBASTO.frx":1A5C
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Listado de Existencias Negativas"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2760
               Picture         =   "COBASTO.frx":1D66
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "COBASTO.frx":2070
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Existencias Físicas Actuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - Ne- gativos y Críticos"
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
               Left            =   630
               TabIndex        =   18
               Top             =   105
               Width           =   1740
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Consultas por Pantalla"
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
         TabIndex        =   14
         Top             =   525
         Width           =   4005
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   20
            Top             =   420
            Width           =   2955
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
               Height          =   600
               Left            =   0
               Picture         =   "COBASTO.frx":237A
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Consulta de Disponibilidad Física Actual y Ubicaciones"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command9 
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
               Left            =   2340
               Picture         =   "COBASTO.frx":24C4
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ficha Kardex - Histórico de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Disponibilidad y      Mov. Historicos"
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
               TabIndex        =   21
               Top             =   105
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3045
      TabIndex        =   12
      Top             =   5250
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "COBASTO.frx":27CE
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   10920
   End
End
Attribute VB_Name = "COBASTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call FINAL("")
End Sub


Private Sub Command15_Click()
    Command15.Enabled = False
    Call FINAL("*SOREMPE")
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Call CARTADEPOSI
    Command16.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call EXISTO2
    If ULTREG&("NEGASTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*EXISTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*EXISTL", "")
        End If
    End If
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    If Left$(TRIM$(UCase$(EMPREAC$)), 4) = "SABO" Then
         CONEXI.Show
       Else
         Call CARPEDPE
         CODISTOK.Show
    End If
    Command2.Enabled = True
    '
End Sub

Private Sub Command20_Click()
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

Private Sub Command23_Click()
    Command23.Enabled = False
    MODUCO$ = "REMSTAN"
    If FileLen("REMSTAN.EXE") < 400000 Then
        MODUCO$ = "#REMSTAN"
    End If
    Call CONECRUN(MODUCO$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
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
    OPTRAJUST.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("CONSTOCK")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("*COMOFAC", "")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESUESTA.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command30_Click()
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


Private Sub Command9_Click()
    Command9.Enabled = False
    'Call CONECRUN("*KARDEX", "")
    Call GENKARDEX("")
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
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

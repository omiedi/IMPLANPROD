VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form QUERSTK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Consultas y Listados"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8085
   Icon            =   "QUERSTK07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   7200
      ScaleHeight     =   7545
      ScaleWidth      =   1665
      TabIndex        =   30
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   41
         Top             =   3840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "QUERSTK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1660
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Disp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "QUERSTK07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Sub-Sistema de Disponibilidad General de Stocks e Insumos"
         Top             =   2260
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Statist"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "QUERSTK07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Modulo de Estadísticas de Stock"
         Top             =   2950
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "QUERSTK07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   300
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "QUERSTK07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   900
         Width           =   650
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Statist"
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
      Height          =   750
      Left            =   105
      Picture         =   "QUERSTK07.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Análisis Estadístico de Consumos de Stock"
      Top             =   6090
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "MATERIALES RESERVADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4440
      TabIndex        =   23
      Top             =   1530
      Width           =   2640
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   1770
         TabIndex        =   24
         Top             =   420
         Width           =   1800
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "QUERSTK07.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Reservas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   26
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "INVENTARIO CONTINUO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4440
      TabIndex        =   19
      Top             =   120
      Width           =   2640
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   1770
         TabIndex        =   20
         Top             =   420
         Width           =   1800
         Begin VB.CommandButton Command6 
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
            Left            =   0
            Picture         =   "QUERSTK07.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Material en Depósitos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   22
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "VARIOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4440
      TabIndex        =   15
      Top             =   2950
      Width           =   2640
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2370
         TabIndex        =   16
         Top             =   420
         Width           =   2400
         Begin VB.CommandButton Command10 
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
            Left            =   1830
            Picture         =   "QUERSTK07.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Compras-Recepción-Consumos"
            Top             =   0
            Width           =   560
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
            Height          =   645
            Left            =   0
            Picture         =   "QUERSTK07.frx":1CA4
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Parte Diario de Fabricacion por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Productos en Proceso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   600
            TabIndex        =   17
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROL DE EXISTENCIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4100
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "LISTADOS DE CONTROL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   3855
         Begin VB.Frame Frame10 
            BackColor       =   &H00CDDADA&
            Caption         =   "SALIDA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   3375
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00CDDADA&
               Caption         =   "Impresión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   1890
               TabIndex        =   3
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
               Appearance      =   0  'Flat
               BackColor       =   &H00CDDADA&
               Caption         =   "Pantalla"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   315
               TabIndex        =   2
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   3345
            TabIndex        =   10
            Top             =   420
            Width           =   3375
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
               Height          =   645
               Left            =   1695
               Picture         =   "QUERSTK07.frx":1FAE
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Listado de Lotes y Ubicaciones"
               Top             =   0
               Visible         =   0   'False
               Width           =   560
            End
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
               Height          =   645
               Left            =   2215
               Picture         =   "QUERSTK07.frx":22B8
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
               Height          =   645
               Left            =   2730
               Picture         =   "QUERSTK07.frx":25C2
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
               Top             =   0
               Width           =   660
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
               Height          =   645
               Left            =   0
               Picture         =   "QUERSTK07.frx":28CC
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Existencias Actuales: Cantidades/Valor/Control Existencias por N.Serie"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias Negativos Críticos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   720
               TabIndex        =   11
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONSULTAS / LISTADOS GENERALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3855
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   3345
            TabIndex        =   13
            Top             =   420
            Width           =   3375
            Begin VB.CommandButton Command9 
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
               Height          =   645
               Left            =   2245
               Picture         =   "QUERSTK07.frx":2BD6
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Stocks en Proceso "
               Top             =   0
               Width           =   560
            End
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
               Height          =   645
               Left            =   0
               Picture         =   "QUERSTK07.frx":2D20
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Listado de Existencias y Ubicaciones"
               Top             =   0
               Width           =   560
            End
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
               Height          =   645
               Left            =   2800
               Picture         =   "QUERSTK07.frx":302A
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Gestion de Reservas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - Histórico - Reservas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   600
               TabIndex        =   14
               Top             =   75
               Width           =   1620
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "QUERSTK07.frx":3334
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "QUERSTK07.frx":3568
      TabIndex        =   37
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Acceso a Estadisticas de Stocks"
      Height          =   750
      Left            =   1050
      TabIndex        =   29
      Top             =   6090
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Acceso a Disponibilidad General de Stocks"
      Height          =   645
      Left            =   945
      TabIndex        =   28
      Top             =   5040
      Width           =   1485
   End
End
Attribute VB_Name = "QUERSTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub


Sub GENERAXCEL(NombreDeArchivo$)
   '
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) WHERE CANSALDO > 0.05 AND (DEPOUBIC=5 OR DEPOUBIC=6) ORDER BY DEPOUBIC"
   Dim Rst As ADODB.Recordset
   Set Rst = READSET(SQLX$)
   
   XLSHEET.Cells(1, 1) = "Codigo"
   XLSHEET.Cells(1, 2) = "U-M"
   XLSHEET.Cells(1, 3) = "Deposito"
   XLSHEET.Cells(1, 4) = "Lote"
   XLSHEET.Cells(1, 5) = "Ubicacion"
   XLSHEET.Cells(1, 6) = "Saldo"
   XLSHEET.Cells(1, 7) = "Ancho mm."
   XLSHEET.Cells(1, 8) = "Sd.Mt.Lin"

   JJ& = 1
   With Rst
        Do While Not .EOF
             JJ& = JJ& + 1
                CIXI% = XVALO(!cod_inte)
                Codigo$ = CODEXT$(CIXI%)
             XLSHEET.Cells(JJ&, 1) = Codigo$
                UMED$ = VALOBADA("MASTER", "UMEDIDA", "CODINT=" & CIXI%)
             XLSHEET.Cells(JJ&, 2) = UMED$
             XLSHEET.Cells(JJ&, 3) = SAFETEXT$(!DepoUbic)
             XLSHEET.Cells(JJ&, 4) = SAFETEXT$(!idenlote)
             XLSHEET.Cells(JJ&, 5) = SAFETEXT$(!UBICFIS)
                SALDO$ = FORMATNUM$(XVALO(!CanSaldo), "F12.3")
             XLSHEET.Cells(JJ&, 6) = XVALO(SALDO$)
                AnchoMili$ = FORMATNUM$(XVALO(!AnchoMili), "F12.3")
             XLSHEET.Cells(JJ&, 7) = XVALO(AnchoMili$)
                SALMETRO$ = FORMATNUM$(XVALO(!SALMETRO), "F12.3")
             XLSHEET.Cells(JJ&, 8) = XVALO(SALMETRO$)
             .MoveNext
        Loop
   End With
     
   On Error Resume Next
   Rst.Close
   Set Rst = Nothing
   
   HOY_HORAFECHA0$ = CStr(Now)
   HOY_HORAFECHA1$ = Mid$(HOY_HORAFECHA0$, 1, 17)
   HOY_HORAFECHA1$ = REPLACAR(HOY_HORAFECHA1$, " ", "_")
   HOY_HORAFECHA1$ = REPLACAR(HOY_HORAFECHA1$, "/", "")
   HOY_HORAFECHA1$ = REPLACAR(HOY_HORAFECHA1$, ":", "")
   NombreDeArchivo$ = NombreDeArchivo$ & "_" & HOY_HORAFECHA1$ & ".XLS"
   XLSHEET.SaveAs NombreDeArchivo$
   'XLAPP.Visible = True
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado

   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   'Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub

Private Sub Command10_ORIGINAL_Click()
    Command10.Enabled = False
    'Call CARGA_VECTORES_DE_MATERIALES_REMITIDOS
    Exit Sub
    '
    'ESTE CODIGO COMENTADO ERA PARA QUE EN EL CASO QUE SE REPITIERA EN EL ENCABEZADO UN MISMO TITULO
    'SOLO IMPRIMIA UNO AGRUPANDO LA CELDA CON LA DE AL LADO, PERO DESPUES TRAJO PROBLEMAS CON LA EXPORTACION A EXCEL
    'CARGO ENCABEZADO DE UNA GRILLA
    'establecemos el tipo de agrupación para combinar celdas
'    MUESTRADATOS.GRID.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS.GRID, "Código/A16;Descripción/A32;U. (COMPRA)/F10;$ (COMPRA)/F10;U. (INGR.)/F10;$ (INGR.)/F10;U. (SALE)/F10;$ (SALE)/F10;U. (INGR.-SALE)/F16;$ (INGR.-SALE)/F16;U. (COMP.-SALE)/F16;$ (COMP.-SALE)/F16", MUESTRADATOS, 1, 1)
'    MUESTRADATOS.GRID.MergeCol(0) = True
'    MUESTRADATOS.GRID.MergeCol(1) = True
'    MUESTRADATOS.GRID.MergeCol(2) = True
'    MUESTRADATOS.GRID.Rows = 3
'    MUESTRADATOS.GRID.MergeCells = 1 'Estilo libre
'    'Se combinan las celdas de la 1ra y 2da fila que tengan el mismo contenido
'    MUESTRADATOS.GRID.MergeRow(0) = True
'    MUESTRADATOS.GRID.MergeRow(1) = True
    '
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    MUESTRADATOS.MSF_2.Rows = 3
    MUESTRADATOS.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS.MSF_2, "Código/A16;Descripción/A32;U. (COMPRA)/F10;$ (COMPRA)/F10;U. (INGR.)/F10;$ (INGR.)/F10;U. (SALE)/F10;$ (SALE)/F10;U. (INGR.-SALE)/F16;$ (INGR.-SALE)/F16;U. (COMP.-SALE)/F16;$ (COMP.-SALE)/F16", MUESTRADATOS, 0, 1)

    HOII% = HOY(HOS$)
    FEX = HOY(HO$)
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)

10  OBX$ = OBJPLA$("DESHASFECHMAT", MKI$(Des%) + MKI$(HOII%))
    If OBX$ <> "" Then
      Des% = CVI(Left$(OBX$, 2))
      Has% = CVI(Mid$(OBX$, 3, 2))
      TMAT% = Asc(Mid$(OBX$, 5, 1))
    Else
      GoTo 99
    End If
    '
    If Des% > Has% Then
       Call COMUNI("Verifique Rango de Fechas!!")
       GoTo 99
    End If
    '
    J& = 0
    Screen.MousePointer = 11
    '
    J& = 0
    MUESTRADATOS.GRID.Rows = 2
    SQLX$ = "SELECT CODINT,CODIGO,DESCRIPCION FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STATUS >= 0"
    If TMAT% > 0 Then SQLX$ = SQLX$ + " AND TIMATE = " & TMAT%
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
           i& = i& + 1
           Call TRACE(20, i&, XVALO(NUMAS%))
           CI1% = XVALO(!CODINT)
           CANCOMPRA = CANTICOMPRADA(CI1%, TOTCOMPRA#, Des%, Has%)
           COMPRAPESOS# = TOTCOMPRA#
           CANRECI = CANTIRECIBIDA(CI1%, Des%, Has%)
           CANCON = CANTICONSU(CI1%, Des%, Has%)
           If (CANCOMPRA + CANRECI + CANCON) > 0 Then
              COU# = COSUNI(CI1%)
              J& = J& + 1
              MUESTRADATOS.GRID.Rows = J&
              MUESTRADATOS.GRID.Rows = MUESTRADATOS.GRID.Rows + 1
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 1) = SAFETEXT$(!Codigo)
              MUESTRADATOS.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descripcion)
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 3) = FORMATNUM$(CANCOMPRA, "F10.1")
              MUESTRADATOS.GRID.TextMatrix(J&, 4) = FORMATNUM$(COMPRAPESOS#, "F12.2")
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 5) = FORMATNUM$(CANRECI, "F10.1")
                 RECIPESOS# = COU# * CANRECI
              MUESTRADATOS.GRID.TextMatrix(J&, 6) = FORMATNUM$(RECIPESOS#, "F12.2")
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 7) = FORMATNUM$(CANCON, "F10.1")
                 CONSUPESOS# = COU# * CANCON
              MUESTRADATOS.GRID.TextMatrix(J&, 8) = FORMATNUM$(COU# * CANCON, "F12.2")
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 9) = FORMATNUM$(CANRECI - CANCON, "F10.1")
              MUESTRADATOS.GRID.TextMatrix(J&, 10) = FORMATNUM$(RECIPESOS# - CONSUPESOS#, "F12.2")
              '
              MUESTRADATOS.GRID.TextMatrix(J&, 11) = FORMATNUM$(CANCOMPRA, "F10.1")
              MUESTRADATOS.GRID.TextMatrix(J&, 12) = FORMATNUM$(COMPRAPESOS# - CONSUPESOS#, "F12.2")
              '
           End If
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS.GRID, &HE0E0E0, 0, 2)
    Call COLOREAR_GRILLASQL(MUESTRADATOS.GRID, &HE0E0E0)
    MUESTRADATOS.Caption = "Informe de Ingresos, Consumos, Salidas por Cantidad, Valorizados"
    Screen.MousePointer = 1
    MUESTRADATOS.Show 1
    '
99  Command10.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
'    Call CARGA_VECTORES_DE_MATERIALES_REMITIDOS
'    Exit Sub
    '
    'ESTE CODIGO COMENTADO ERA PARA QUE EN EL CASO QUE SE REPITIERA EN EL ENCABEZADO UN MISMO TITULO
    'SOLO IMPRIMIA UNO AGRUPANDO LA CELDA CON LA DE AL LADO, PERO DESPUES TRAJO PROBLEMAS CON LA EXPORTACION A EXCEL
    'CARGO ENCABEZADO DE UNA GRILLA
    'establecemos el tipo de agrupación para combinar celdas
'    MUESTRADATOS.GRID.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS.GRID, "Código/A16;Descripción/A32;U. (COMPRA)/F10;$ (COMPRA)/F10;U. (INGR.)/F10;$ (INGR.)/F10;U. (SALE)/F10;$ (SALE)/F10;U. (INGR.-SALE)/F16;$ (INGR.-SALE)/F16;U. (COMP.-SALE)/F16;$ (COMP.-SALE)/F16", MUESTRADATOS, 1, 1)
'    MUESTRADATOS.GRID.MergeCol(0) = True
'    MUESTRADATOS.GRID.MergeCol(1) = True
'    MUESTRADATOS.GRID.MergeCol(2) = True
'    MUESTRADATOS.GRID.Rows = 3
'    MUESTRADATOS.GRID.MergeCells = 1 'Estilo libre
'    'Se combinan las celdas de la 1ra y 2da fila que tengan el mismo contenido
'    MUESTRADATOS.GRID.MergeRow(0) = True
'    MUESTRADATOS.GRID.MergeRow(1) = True
    '
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    MUESTRADATOS.MSF_2.Rows = 3
    MUESTRADATOS.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS.MSF_2, "Código/A16;Descripción/A32;U. (COMPRA)/F10;$ (COMPRA)/F10;U. (INGR.)/F10;$ (INGR.)/F10;U. (SALE)/F10;$ (SALE)/F10;U. (INGR.-SALE)/F16;$ (INGR.-SALE)/F16;U. (COMP.-SALE)/F16;$ (COMP.-SALE)/F16", MUESTRADATOS, 0, 1)

    HOII% = HOY(HOS$)
    FEX = HOY(HO$)
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then
      Screen.MousePointer = 1:
      Command10.Enabled = True
      Exit Sub
    End If
    '
    If Des% > Has% Then
       Call COMUNI("Verifique Rango de Fechas!!")
       GoTo 99
    End If
    '
    J& = 0
    Screen.MousePointer = 11
    '
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
    
    MUESTRADATOS.GRID.Rows = 2
    '''''''''''''''''''''''''''''''''''''''''''
    'COMIENZA PROCESO SE RECORREN LOS MOVIMIENTOS RT (REMITOS) SEGUN EL RANGO DE FECHAS
    'ACUMULADOS POR CODIGO
    'ABRO LA ESTRUCUTRA DE CADA UNO Y ACUMULO LOS CONSUMOS EN UN VECTOR  GUARDANDO EL CONSUMO
    Dim CONSUMO_COMPONENTE()
    ReDim Preserve CONSUMO_COMPONENTE(NUMAS%)
    SQLX$ = "SELECT COD_INTE, SUM(CANTSALID) AS CANSAL FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT'"
    SQLX$ = SQLX$ + " AND CANTSALID > 0 "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND FECHMOV >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FECHMOV <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY  COD_INTE"
    SQLX$ = SQLX$ + " ORDER BY COD_INTE"
    Set rs = READSET(SQLX$)
    With rs
     Do While Not .EOF
           CI1% = XVALO(!cod_inte)
           CANTISALE = XVALO(!CANSAL)
           'Call AREXPLO(CI1%, 1, 1)
           Call EXPLOCOM(CI1%, 1)
           For KKJ% = 1 To CONTAMUL%
             CONSUMO_COMPONENTE(CIJ%(KKJ%)) = CONSUMO_COMPONENTE(CIJ%(KKJ%)) + (USOI(KKJ%) * CANTISALE)
           Next KKJ%
           .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'RECORRO LOS VECTORES CON INFORMACION Y PRESENTO LOS COMONENTES CONSUMIDOS SEGUN ESTRUCUTRA DE CADA PRODUCTO DEL REMITO
    '
 A = CONSUMO_COMPONENTE(1)
 A = CONSUMO_COMPONENTE(2)
 A = CONSUMO_COMPONENTE(3)
 A = CONSUMO_COMPONENTE(4)
 
    FIN& = NUMAS%
    For JJ% = 1 To FIN&
        i& = i& + 1
        Call TRACE(20, i&, FIN&)
        CI1% = JJ%
        CANCON = CONSUMO_COMPONENTE(JJ%)
        If CANCON > 0 Then
           CANCOMPRA = CANTICOMPRADA(CI1%, TOTCOMPRA#, Des%, Has%)
           COMPRAPESOS# = TOTCOMPRA#
           CANRECI = CANTIRECIBIDA(CI1%, Des%, Has%)

           COU# = COSUNI(CI1%)
           J& = J& + 1
           MUESTRADATOS.GRID.Rows = J&
           MUESTRADATOS.GRID.Rows = MUESTRADATOS.GRID.Rows + 1
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
           MUESTRADATOS.GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 3) = FORMATNUM$(CANCOMPRA, "F10.1")
           MUESTRADATOS.GRID.TextMatrix(J&, 4) = FORMATNUM$(COMPRAPESOS#, "F12.2")
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 5) = FORMATNUM$(CANRECI, "F10.1")
              RECIPESOS# = COU# * CANRECI
           MUESTRADATOS.GRID.TextMatrix(J&, 6) = FORMATNUM$(RECIPESOS#, "F12.2")
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 7) = FORMATNUM$(CANCON, "F10.1")
              CONSUPESOS# = COU# * CANCON
           MUESTRADATOS.GRID.TextMatrix(J&, 8) = FORMATNUM$(COU# * CANCON, "F12.2")
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 9) = FORMATNUM$(CANRECI - CANCON, "F10.1")
           MUESTRADATOS.GRID.TextMatrix(J&, 10) = FORMATNUM$(RECIPESOS# - CONSUPESOS#, "F12.2")
           '
           MUESTRADATOS.GRID.TextMatrix(J&, 11) = FORMATNUM$(CANCOMPRA - CANCON, "F10.1")
           MUESTRADATOS.GRID.TextMatrix(J&, 12) = FORMATNUM$(COMPRAPESOS# - CONSUPESOS#, "F12.2")
           '
        End If
        If JJ% >= 32767 Then Exit For
    Next JJ%
    '
    'LIMPIO LOS VECTORES.
    Erase CIMOVIM%, CAMOVIM

    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS.GRID, &HE0E0E0, 0, 2)
    Call COLOREAR_GRILLASQL(MUESTRADATOS.GRID, &HE0E0E0)
    MUESTRADATOS.Caption = "Informe de Ingresos, Consumos, Salidas por Cantidad, Valorizados"
    Screen.MousePointer = 1
    MUESTRADATOS.Show 1
    '
99  Command10.Enabled = True
    Screen.MousePointer = 1

End Sub

Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
        TXTX$ = " CRITERIO DE REPORTE COMPARATIVO DE COMPRAS - INGRESOS - SALIDAS (CANTIDAD Y VALORIZADO)" + vbCrLf
        TXTX$ = TXTX$ + "------------------------------------------------------------------------------------------------------------------------" + vbCrLf
        TXTX$ = TXTX$ + " Captura de los datos de compras en el rango selccionado" + vbCrLf
        TXTX$ = TXTX$ + "(tomando como fecha la de la generacion de la o.c.)" + vbCrLf
        TXTX$ = TXTX$ + " Generando una fila para cada item," + vbCrLf
        TXTX$ = TXTX$ + " Imprimiendo el total de unidades y acumulando el valor  del calculo cantidad * importe" + vbCrLf
        TXTX$ = TXTX$ + " Respetando el precio de cada compra." + vbCrLf
        TXTX$ = TXTX$ + " Captura de los datos de la recepcion en dicho rango (tomando como fecha la de la recepcion)" + vbCrLf
        TXTX$ = TXTX$ + " Calculando el importe que viene segun el costo de reposición " + vbCrLf
        TXTX$ = TXTX$ + " Captura de la salida de estos materiales segun el consumo (por orden de fabricacion, remito o transferencia o vale de materiales)" + vbCrLf
        TXTX$ = TXTX$ + " Valorizados segun el costo de reposicion actual." + vbCrLf
        TXTX$ = TXTX$ + " Resultado en ultimas columnas arroja la Diferencia entre Ingresa y Sale "
        TXTX$ = TXTX$ + " y Compra y Sale"
        MsgBox TXTX$
   End If
   '
End Sub

 Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("ORFABR07/OPRESERV", "Gestión de Ordenes de Fabricación")
   Command14.Enabled = True
End Sub



 Sub Command19_Click()
    Command19.Enabled = False
    '\\\OT-05-0379-WAR-06/05/05///
    XX1% = COMALTER%("Opciones de Listado: \\Solo Cantidades\Valorizado a última Compra\Control de Existencias por Nº.Serie")
    Select Case XX1%
      Case 1
        Call EXISTO2
        GoTo 99
      Case 2
        Call EXISTO4
        GoTo 99
      Case 3
        GoTo 10 ' LISTADO DE CONTROL DE EXISTENCAS POR NUMEROS DE SERIE
      Case Else
        GoTo 99
   End Select
   '\\\OT-05-0379-WAR-FIN///
   '
10 Screen.MousePointer = 11 ' LISTADO DE CONTROL DE EXISTENCAS POR NUMEROS DE SERIE
   Call CARDEPOS
   '
   U& = 0
   Call SETULTREG("CONUMSE", 0)
   '
   For i& = 1 To NUMAS%
     URE& = NUMAS%
     Call TRACE(20, i&, URE&)
     '
     CI% = i&
     If ESTRAZABLE%(CI%) = 2 Then
       For J% = 1 To NUSU%
         NRODEPO% = COSU%(J%)
         CONDI$ = "(NUMREMITO = '' OR NUMREMITO IS NULL)"
         CONDI$ = CONDI$ + " AND DEPONUMSE = " & NRODEPO% & ""
         CONDI$ = CONDI$ + " AND COD_INTE = " & CI% & ""
         '
         STANUMSE& = XVALO(VALOBADA("TANUMSE", "COUNT(*)", CONDI$))
         STMOVISTO& = XVALO(STODEPOS(CI%, NRODEPO%))
         If STANUMSE& <> STMOVISTO& Then
           Call REPLA(TX$, MKI$(CI%), 1, 2)
           Call REPLA(TX$, MKI$(NRODEPO%), 3, 2)
           Call REPLA(TX$, MKS$(STANUMSE&), 5, 4)
           Call REPLA(TX$, MKS$(STMOVISTO&), 9, 4)
           Diferencia& = STANUMSE& - STMOVISTO&
           '
           If Abs(Diferencia&) > 0 Then
              Call REPLA(TX$, MKS$(Diferencia&), 13, 4)
              U& = U& + 1
              Call GRAREG("CONUMSE", TX$, U&)
           End If
           '
         End If
       Next J%
     End If
   Next i&
   '
   Call SETULTREG("CONUMSE", U&)
   Call CIERRARCH("CONUMSE")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   If U& > 0 Then
         Call FINAL("*CONUMSE")
       Else
         Call COMUNI("No se Han Detectado Diferencias\entre las Bases de Datos de Stocks\y de Numeros de Serie.")
   End If
   '
99 Command19.Enabled = True
End Sub

 Sub Command2_Click()
    Command2.Enabled = False
    Call LIUBIEXI
    Command2.Enabled = True
End Sub

 Sub Command20_Click()
     Command20.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "NEUQUEN") > 0 Then
       Call StockNegativoNeuquen
       'LISTADO ESPECIAL PARA NEUQUEN EL STOCK MINIMO EN TABLA DATASQL LISTA =  'StMinNeu'
       GoTo 99
    End If
    
    Call CRITSTO2
    If ULTREG&("CRITSTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*CRITSTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*CRITSTL", "")
        End If
    End If
99  Command20.Enabled = True
End Sub
Sub StockNegativoNeuquen()
   'Stock Negativo considerando Lote indiv ventA
        Campos$ = "Codigo/A20" '1
        Campos$ = Campos$ + ";Descripción/A32" '2
        Campos$ = Campos$ + ";Stock Minimo/F12.1" '3
        Campos$ = Campos$ + ";Stock Nqn./F12.1" '4
        Campos$ = Campos$ + ";Reposición/F12.1" '5
    
        J& = 0
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        '
        SQLX$ = " SELECT D.CAMPO1,COALESCE(D.CAMPO3, 0) AS CAMPO3,M.DESCRIPCION ,M.CODINT "
        SQLX$ = SQLX$ + " FROM DATASQL D WITH(NOLOCK)"
        SQLX$ = SQLX$ + " INNER JOIN MASTER M ON M.CODIGO = D.CAMPO1"
        SQLX$ = SQLX$ + "  WHERE D.LISTA = 'StMinNeu'  AND  COALESCE(D.CAMPO3, 0) > 0"
        SQLX$ = SQLX$ + "  AND M.STATUS >=0"
        SQLX$ = SQLX$ + " ORDER BY D.CAMPO1"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             
             SALDO = SaldoFinal(XVALO(!CODINT))
'             If SALDO < XVALO(!CAMPO3) Then para que muestre todos los de la lista de StMinNeu
                J& = J& + 1
                GRILLAGRAL.GRID.Rows = J& + 1
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = SAFETEXT$(!CAMPO1)
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descripcion)
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!CAMPO3)
                GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(XVALO(SALDO), "F12.1")
                Diferencia = XVALO(!CAMPO3) - XVALO(SALDO)
                If Diferencia > 0 Then Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, J&, 5)
                GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(Diferencia), "F12.1")
'             End If
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        
        
   
   
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        TX$ = "Items con Saldo por debajo del Stock Mínimo"
        
        GRILLAGRAL.Caption = TX$
        Screen.MousePointer = 1
        GRILLAGRAL.Show 1
        
        NORFA$ = RESP_ZELE_VECT(1)
        
        Dim MSF As msFlexGrid
        If GRILLAGRAL.MENU1 > 0 Then
             Set MSF = GRILLAGRAL.GRID
             If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
             Call CARGA_TABLA_IMPRE(MSF, Campos$, "DEBASMIN", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
             Call FINAL("?" & RET_NAMECONS$)
             GRILLAGRAL.MENU1 = 0
             Screen.MousePointer = 1
             Exit Sub
        End If

   
End Sub


Private Sub Command23_Click()
   '
End Sub

Sub Command27_Click()
    Command27.Enabled = False
    '\\\OT-5-0428-FG-28/06/05///
    If EXISTE%("ESTOCK07.EXE") > 0 Then
        OPX% = ALTERNA%("Estadísticas de Stock\Control Estadístico de Consumo")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        If OPX% = 1 Then GoTo 80
        If OPX% = 2 Then
           Call CONECRUN("ESTOCK07", "Control Estadístico de Consumo")
           GoTo 99
        End If
    End If
80  ESTASTK07.Show 1
    Command27.Enabled = True
    Exit Sub
    '\\\OT-5-0428-FG-28/06/05///
    'Call CONECRUN("ESTOCK04", "Modulo de Estadísticas de Stock")
    HOII% = HOY(HOS$)
10  OBX$ = OBJPLA$("DESHESFECHREC", MKI$(HOII%) + MKI$(HOII%))
    If OBX$ <> "" Then
      Des% = CVI(Left$(OBX$, 2))
      Has% = CVI(Mid$(OBX$, 3, 2))
      If Des% <= Has% Then
        COMOVI$ = TRIM$(UCase$(InputBox$("Codigo de Movimiento", "Selección Tipo de Movimiento")))
        Screen.MousePointer = 11
        
        'Call ABRESTOCKS
        'DESDE = CDbl(Int(CVDAT(DES%)))
        'HASTA = CDbl(Int(CVDAT(HAS%)))
        'SQLX$ = "SELECT * FROM MOVISTO "
        'SQLX$ = SQLX$ + "WHERE Cod_Inte >0 "
        'SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) >= " & DESDE & " "
        'SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
        'If COMOVI$ <> "" Then
        '  SQLX$ = SQLX$ + "AND CodiMovi = '" & COMOVI$ & "' "
        'End If
        'SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        '
        Call ABRESTOCKS
        Desde = FETEXCOR1$(Des%)
        Hasta = FETEXCOR1$(Has%)
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE Cod_Inte >0 "
        SQLX$ = SQLX$ + "AND FechMov >= " & Desde & " "
        SQLX$ = SQLX$ + "AND FechMov <= " & Hasta & " "
        If COMOVI$ <> "" Then
          SQLX$ = SQLX$ + "AND CodiMovi = '" & COMOVI$ & "' "
        End If
        SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        
        
        JJ& = 0
        Call COPYSTRU("COKARDEX", "LIMOVSTO")
        REBLA$ = REGBLAN$("COKARDEX")
        'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'        Dim MoviTemp As ADODB.Recordset
'        Set MoviTemp = New ADODB.Recordset
'        MoviTemp.CursorType = adOpenDynamic
'        MoviTemp.LockType = adLockReadOnly
'        MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
        Set MoviTemp = READSET(SQLX$)
        Screen.MousePointer = 11
        With MoviTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              YY$ = REBLA$
              Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
              Call REPLA(YY$, MKI$(!cod_inte), 3, 2)
              Call REPLA(YY$, !idenlote, 5, 16)
              Call REPLA(YY$, !CODIMOVI, 21, 2)
              Call REPLA(YY$, !dopricor, 23, 10)
              Call REPLA(YY$, !DoSecCor, 33, 10)
              Call REPLA(YY$, !ReferCor, 43, 30)
              Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
                MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
              Call REPLA(YY$, MONEX$, 81, 2)
              Call REPLA(YY$, Chr$(!depomovi), 83, 1)
              Call REPLA(YY$, MKI$(!NUME_CLIE), 84, 2)
              Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
              Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
              Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
              Call REPLA(YY$, MKD$(0), 104, 8)
              Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
              Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
              
              JJ& = JJ& + 1
              Call GRAREG("LIMOVSTO", YY$, JJ&)
            .MoveNext
            Loop
          End If
        End With
        MoviTemp.Close
        Set MoviTemp = Nothing
        
        'Movisto.Close
        '
        Call SETULTREG("LIMOVSTO", JJ&)
        '
        Call FILESORT("LIMOVSTO", "", 2, 2, "ASC")
        Call FINAL("*LIMOVSTO")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    '
99  Command27.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("ORFABR07/PRODUPRO", "Gestión de Ordenes de Fabricación")
99  Command3.Enabled = True
End Sub
'
Private Sub Command4_Click()
   Command4.Enabled = False
   'Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command4.Enabled = True
End Sub

Private Sub Command43_Click()

End Sub

Sub Command5_Click()
    Command5.Enabled = False
    DISPSTK07.Show 1
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
 Sub Command6_Click()
    Command6.Enabled = False
    Call CONECRUN("INVENT07", "Inventarios Físicos y Valorizados")
    Command6.Enabled = True
End Sub

 Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("ORFABR07/OPRESERV", "Gestión de Ordenes de Fabricación")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call APARTIMA
  Screen.MousePointer = 1
  '\\\OT-08-0999-OD-31/10/08///
  If InStr(EMPREAC$, "ACHERNAR") > 0 Then
     OPX% = ALTERNA("Listado de Lotes y Ubicaciones\Planilla Diaria de Stock\Listado de Ubicaciones Automático")
     If OPX% = 0 Then GoTo 99
     If OPX% = 1 Then GoTo 10
     If OPX% = 2 Then
        Call Cargar_PLADIASTO
        GoTo 99
     End If
     If OPX% = 3 Then
        If CLACONTRA% = 0 Then GoTo 99
        OPX% = COMALTER%("¿Desea Generar Reporte En Esta Terminal?\\Si, Asignar\No, Cancelar")
        If OPX% = 1 Then
            ARCHI_OMISION$ = TRIM$(Control("CONSTO07", "LIAUTUBI"))
            Archivo$ = InputBox("Ingrese la Ubicación y Nombre del Archivo a Generar", "Listado de Ubicaciones", ARCHI_OMISION$)
            If TRIM$(Archivo$) <> "" Then
                HORA_OMISION$ = TRIM$(Control("CONSTO07", "HOAUTUBI"))
                HORA$ = InputBox("Ingrese el Horario de Ejecución", "Listado de Ubicaciones", HORA_OMISION$)
                HORA$ = Format(HORA$, "hh:mm")
                If XVALO(Mid$(HORA$, 1, 2)) < 24 And XVALO(Mid$(HORA$, 4, 2)) < 60 Then
                    Call GRACONTROL("CONSTO07", "LIAUTUBI", TRIM$(Archivo$))
                    Call GRACONTROL("CONSTO07", "HOAUTUBI", TRIM$(HORA$))
                    Call GRACONTROL("CONSTO07", "USLIAUUB", IDENTER$)
                    Call COMUNI("Datos de Configuración Registrados")
                End If
            End If
            GoTo 99
        End If
     End If
  End If
  '
10 EPAC$ = UCase$(TRIM$(EMPREAC$))
  OPCIONESX$ = "Formato de Salida\\Por Depósito\Total por Código"
  If InStr(EPAC$, "ASEMBLI") > 0 Then OPCIONESX$ = OPCIONESX$ + "\Diagnóstico Lote Vs. Saldos\Control de Sobre-Embarque"
  '
  OPX% = COMALTER%(OPCIONESX$)
  If OPX% = 1 Then
    Call FINAL("?UBICALOTE") ' LISTADO DE LOTES Y UBICACIONES
  ElseIf OPX% = 2 Then
    Call FINAL("?UBILOTOT") ' LISTADO POR LOTES Y ACUMULADOS POR ARTICULOS
  ElseIf OPX% = 3 Then
       Call LOTES_VS_SALDOS
  ElseIf OPX% = 4 Then
     Call CONTROL_SOBREEMBARQUE
  End If
  
  '\\\OT-08-0999-OD-FIN///
99  Command8.Enabled = True
End Sub

 Sub Command9_Click()
    Command9.Enabled = False
    'Call SITSTOK07.Command3_Click
    CFGSALSTO.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
   
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    
    If InStr(1, EMPREAC$, "ALEJPA") > 0 Then
        Command9.Enabled = True
    End If

    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    '
    '\\\OT-08-0999-OD-31/10/08///
    'If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Or EXISTE%("TRANSE07.EXE") > 0 Then
      Command8.Visible = True ' BOTON DE LISTADO LOTES Y UBICACIONES
      Label12.Left = 600 ' LABEL DE INFORMACION
    'End If
    '\\\OT-08-0999-OD-FIN///
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    Call GRATITFOR(Caption)
    '
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

'\\\OT-05-0379-WAR-06/05/05///
Sub EXISTO4()
'Dim PUN1#(32767), MOPUN1%(32767), FEPUN1#(32767)
    '
    TMTX% = 0
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
    IncluyeStockEnCero% = 0
    OpcionIncStockCero% = COMALTER%("¿Desea Incluir a los Artículos con Stock En Cero?\\No Incluir\Si Incluir")
    If OpcionIncStockCero% = 2 Then IncluyeStockEnCero% = 1
    '
    Call HEADERS("NEGASVAL", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
'    Call ABRECOMPRAS
'    SQLX$ = "SELECT * FROM OCOMDETA "
'    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 " 'NO ANULADA
'    Set STKVALTMP = Compras.OpenRecordset(SQLX$, 4)
'    With STKVALTMP
'      On Error Resume Next
'      .MoveFirst
'      If Err Then
'        On Error GoTo 0
'        GoTo 11
'      End If
'      On Error GoTo 0
'      .MoveLast
'      TOT& = .RecordCount
'      .MoveFirst
'      RA& = 0
'      Do While .EOF = False
'        FEPUN1#(!Cod_Inte) = CDbl(!Femi_ocom)
'        RA& = RA& + 1
'        Call TRACE(20, RA&, TOT&)
'        .MoveNext
'      Loop
'      .MoveFirst
'      RA& = 0
'      Do While .EOF = False
'        If FEPUN1#(!Cod_Inte) <= CDbl(!Femi_ocom) Then
'          FEPUN1#(!Cod_Inte) = CDbl(!Femi_ocom)
'          MOPUN1%(!Cod_Inte) = !Mone_Emis
'          PUN1#(!Cod_Inte) = !PrUn_Neto '!Pre_Unit
'        End If
'        RA& = RA& + 1
'        Call TRACE(20, RA&, TOT&)
'        .MoveNext
'      Loop
'    End With
    '
    Y$ = REGBLAN("NEGASVAL")
    '
11  For i& = 1 To NUMAS%
      URE& = NUMAS%
      Call TRACE(20, i&, URE&)
      CIXI% = i&
      CDDX$ = CODEXT$(CIXI%)
      If CDDX$ >= COD1$ Then
        If CDDX$ <= COD2$ Then
          TMT1% = TIMATE%(CIXI%)
          If TMT% < 1 Or TMT1% = TMT% Then
            '
            CII& = CIXI%
            'Call VESARRAS(CIXI%)
            'SSTT = STOCKACT(CIXI%)
            SSTT = STODEPOS(CIXI%, DPS%)
            '
            If Abs(SSTT) > 0.05 Or IncluyeStockEnCero% = 1 Then
              Call REPLA(Y$, MKI$(CIXI%), 1, 2)
              Call REPLA(Y$, MKS$(SSTT), 3, 4)
                PUN1# = PreUltCompra(CIXI%, Moneda%, Femiocom)
              Call REPLA(Y$, MKD$(PUN1#), 9, 8)
              Call REPLA(Y$, MKI$(Moneda%), 17, 2)
              
              'Call REPLA(Y$, MKI$(CVINT((CII&))), 19, 2)
              Call REPLA(Y$, MKI$(CVINT(Femiocom)), 19, 2)
                IMPORTE1# = SSTT * PUN1#
              Call REPLA(Y$, MKD$(IMPORTE1#), 21, 8)
                'IMPORTE2# = IMPORTE1# * TICAMONE(MOPUN1%(CII&)) / 1
                IMPORTE2# = IMPORTE1# * TICAMONE(Moneda%) / 1
              Call REPLA(Y$, MKD$(IMPORTE2#), 29, 8)
              J& = J& + 1
              Call GRAREG("NEGASVAL", Y$, J&)
            End If
            '
          End If
        End If
      End If
    Next i&
    '
    Call SETULTREG("NEGASVAL", J&)
    Call FILESORT("NEGASVAL", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        Call FINAL("*LISTVAL")
    End If
    '
End Sub
Function PreUltCompra(CI%, Moneda%, Femiocom)
PreUltCompra = 0
Moneda% = 1

strSQL$ = "select prun_neto, Mone_Emis, Femi_Ocom from ocomdeta WITH(NOLOCK) where cod_inte=" & CI%
strSQL$ = strSQL$ + " and stat_ocom<9 order by Femi_ocom desc"
'Dim MIRS As ADODB.Recordset
'Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL$))
Set MIRS = READSET(strSQL$)
If Not (MIRS.EOF And MIRS.BOF) Then
   MIRS.MoveFirst
   Moneda% = MIRS!Mone_Emis
   PreUltCompra = MIRS!prun_neto '!Pre_Unit
   Femiocom = MIRS!FEMI_OCOM
End If
MIRS.Close
Set MIRS = Nothing
End Function

'\\\OT-05-0379-WAR-FIN///
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
    IncluyeStockEnCero% = 0
    OpcionIncStockCero% = COMALTER%("¿Desea Incluir a los Artículos con Stock En Cero?\\No Incluir\Si Incluir")
    If OpcionIncStockCero% = 2 Then IncluyeStockEnCero% = 1
    '
    Call HEADERS("NEGASTO", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
    For i& = 1 To NUMAS%
      Call TRACE(20, i&, URE&)
      CIXI% = i&
      CDDX$ = CODEXT$(CIXI%)
      If CDDX$ >= COD1$ Then
        If CDDX$ <= COD2$ Then
          TMT1% = TIMATE%(CIXI%)
          If TMT% < 1 Or TMT1% = TMT% Then
            '
            CII& = CIXI%
            'Call VESARRAS(CIXI%) 'jandy
            'SSTT = STOCKACT(CIXI%)
            SSTT = STODEPOS(CIXI%, DPS%)
            If Abs(SSTT) > 0.05 Or IncluyeStockEnCero% = 1 Then
              TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
              J& = J& + 1
              Call GRAREG("NEGASTO", TTN$, J&)
            End If
            '
          End If
        End If
      End If
    Next i&
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
'
Sub NEGASTO2()    'DETERMINACION ITEMS CON STOCK NEGATIVO
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    '\\\OT-08-0775-OD-08/09/08///
    NuDepo% = 0
    OPX% = COMALTER%("El Control Puede ser General o Por Depósito\\General\Por Depósito")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 2 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
           NuDepo% = VALACT1$("TADEPOSI")
        End If
    End If
    '\\\OT-08-0775-OD-FIN///

    Screen.MousePointer = 11
    '
    FIN& = NUMAS%
    For i& = 1 To NUMAS%
      Call TRACE(20, i&, FIN&)
      CI% = i&
      'Call VESARRAS(CI%)
      '
      '\\\OT-08-0775-OD-08/09/08///
      'SSTT = STOCKACT(CI%)
      SSTT = STODEPOS(CI%, NuDepo%)
      '\\\OT-08-0775-OD-FIN///
      '
      If SSTT < (-0.05) Then
        TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
        J& = J& + 1
        Call GRAREG("NEGASTO", TTN$, J&)
      End If
    Next i&
     '\\\OT-08-0775-OD-08/09/08///
     '
     CORRE$ = "Todos los Depósitos"
     If NuDepo% > 0 Then CORRE$ = ECOARCH$("TADEPOSI", Chr$(NuDepo%))
     Call HEADERS("NEGASTO", "")
     Call HEADERS("NEGASTO", "Corresponde a: " + CORRE$)
     Call SETULTREG("NEGASTO", J&)
     Call CIERRARCH("*.*")
     Call FILESORT("NEGASTO", "", 1, 1, "ASC")
     Screen.MousePointer = 1
     '
     '\\\OT-08-0775-OD-FIN///
     
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
    OBX$ = OBJPLA$("TIPOMATE", VDEF$)
    If OBX$ = "" Then TM% = 0: GoTo 10
    '
    TM% = Asc(Left$(OBX$, 1))
    
10  Screen.MousePointer = 11
    FIN& = NUMAS%
    For i& = 1 To NUMAS%
      Call TRACE(20, i&, FIN&)
      CI% = i&
      If TIMATE%(CI%) = TM% Or TM% = 0 Then
          'Call VESARRAS(CI%) 'jandy
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
      End If
    Next i&
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
310 OPX2% = COMALTER%("Opciones de Trabajo\\Solo Items con Stock\Todos")

    Screen.MousePointer = 11
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
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("UBISTOCK", i&)
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
    Next i&
    '
311 FIN& = NUMAS%
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      CI% = i&
      CAN = STOCKACT(CI%)
      '
      If Abs(CAN) >= 0.05 Or OPX2% = 2 Then
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
    Next i&
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
    ENCABE2$ = ""
    FILXLIKE% = 0
    FITROX$ = ""
    OPX% = COMALTER%("Opciones de Filtro: \\Por Rango de Codigos\Por Filtro de Descripción")
    
    If OPX% = 2 Then
       FILTROX$ = UCase$(InputBox("Ingrese Texto Filtro (Utilice '%' como Comodin)", "Busqueda por Filtro", "Ejemplo%"))
       If FILTROX$ <> "" Then
          POS1% = Len(FILTROX$)
          If InStrRev(FILTROX$, "%") > 0 Then
            FILTROX$ = Left$(FILTROX$, POS1% - 1)
            FILXLIKE% = 1: ENCABE2$ = "Filtro: Que la Descripción Contenga el Texto '" & FILTROX$ & "'"
            GoTo 343
          Else
            FILXLIKE% = 0: ENCABE2$ = "Filtro: Que la Descripción Comience con el Texto '" & FILTROX$ & "'"
            GoTo 343
          End If
       End If
    End If
    
330 VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD", VDEF$)
340 If OBX$ <> "" Then
      'INCLUIR O NO ARTCULOS DADOS DE BAJA
      EXCLBAJA% = COMALTER%("Opciones de Listado: \\Incluir Artículos Dados de Baja\Excluir Artículos Dados de Baja")
      '
      Screen.MousePointer = 11
      CI1% = CVI(Left$(OBX$, 2))
      COD1$ = CODEXT$(CI1%)
      CI2% = CVI(Mid$(OBX$, 3, 2))
      COD2$ = CODEXT$(CI2%)
      '
      J& = 0
343   For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        YY$ = REGLEIDO$("UBIEXI", U&)
        CIII% = CVI(Left$(YY$, 2))
        CODX$ = CODEXT$(CIII%)
        'BLOQUE FILTRO POR DESCRIPCION
        If FILTROX$ <> "" Then
           DESC$ = UCase$(DESCRIT0$(CIII%))
           If FILXLIKE% > 0 Then
              If InStr(DESC$, FILTROX$) > 0 Then
                 GoTo 345
              Else
                 GoTo 350
              End If
           Else
              If InStr(Left$(DESC$, POS1%), FILTROX$) > 0 Then
                 GoTo 345
              Else
                 GoTo 350
              End If
           End If
        End If
        'FIN FILTRO POR DESCRIPCION
        If CODX$ >= COD1$ Then
          If CODX$ <= COD2$ Then
345         CIX% = CODINT(CODX$)
            If EXCLBAJA% > 1 And CIX% < 1 Then GoTo 350
            J& = J& + 1
            Call GRAREG("LUBIEXI", YY$, J&)
          End If
        End If
350   Next U&
      Call SETULTREG("LUBIEXI", J&)
      Call CIERRARCH("LUBIEXI")
      Call CIERRARCH("*.*")
      '
      Call HEADERS("LUBIEXI", "")
      ENCABE1$ = "Modo: Incluye Artículos Dados de Baja"
      If EXCLBAJA% > 1 Then ENCABE1$ = "Modo: Excluye Artículos Dados de Baja"
      Call HEADERS("LUBIEXI", "")
      Call HEADERS("LUBIEXI", ENCABE1$)
      Call HEADERS("LUBIEXI", ENCABE2$)
      '
      Screen.MousePointer = 1
      Call FINAL("*SUBIEXI")
    End If
    '
    Screen.MousePointer = 1
End Sub
'
Sub Cargar_PLADIASTO()
    'CREAR CAMPOS
    Call ABRECONEXION
    Call CREACAMPO("STOCKS", "PLADIASTO", "DepoUbic", 3, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "Cod_Inte", 3, 2) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "Codigo", 10, 24) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "FechAlta", 8, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "IdenLote", 10, 16) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "NuOrProd", 10, 16)
    Call CREACAMPO("STOCKS", "PLADIASTO", "SalMetro", 6, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "CanAlta", 7, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "CanConsu", 7, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "CanSaldo", 7, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "UbicFis", 10, 32) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "Causa", 3, 0)
    Call CREACAMPO("STOCKS", "PLADIASTO", "Cajas", 4, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "EtiqxCaja", 4, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "RolxCaja", 4, 0) 'ok
    Call CREACAMPO("STOCKS", "PLADIASTO", "EtiqxRollo", 4, 0) 'ok
    '
    Call BORRAREGISTROS("PLADIASTO", "COD_INTE <> 0 ", REG&)
    '
    SELECHO ("TADEPOSI")
    DEP% = XVALO(VALACT1$("TADEPOSI"))
    If DEP% = 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    Dim PLADIASTO As ADODB.Recordset
    Set PLADIASTO = New ADODB.Recordset
    SQLX$ = "SELECT * FROM PLADIASTO WHERE COD_INTE < 0 "
    PLADIASTO.Open FILTSQL(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    Dim LOTES As ADODB.Recordset
    Set LOTES = New ADODB.Recordset
    '
    SQLX$ = "SELECT * FROM LOTINGRE WHERE DEPOUBIC = " & DEP% & " AND CanConsu < CanAlta"
    Set LOTES = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTES
      Do While Not (.EOF)
         CI% = XVALO(!cod_inte)
         Codigo$ = CODEXT$(CI%)
         If TIMATE%(CI%) = 5 Then
            '
            'calculo el valor de cajas
            On Error Resume Next
               SALDO# = XVALO(!CanAlta) - XVALO(!CanConsu)
               CAJ# = SALDO# / XVALO(!EtiqxCaja)
               CAJENT# = Int(CAJ#)
               CAJRES# = CAJ# - CAJENT#
            If Err > 0 Then
               CAJ# = 0: CAJENT# = 0: CAJRES# = 0
            End If
            On Error GoTo 0
            '
            'calculo el valor de cantidades
            CantPri# = CAJENT# * XVALO(!EtiqxCaja) 'CANTIDAD Etiquetas por caja, DEL PRIMER REGISTRO
            CANTRES# = SALDO# - CantPri#  'CANTIDAD RESTANTE
            '
            On Error Resume Next
               CantRolxCajEnt# = XVALO(!EtiqxCaja) / XVALO(!EtiqxRollo)
               CantRolXCajRes# = CANTRES# / XVALO(!EtiqxRollo)
            If Err > 0 Then
               CantRolxCajEnt# = 0
               CantRolXCajRes# = 0
            End If
            On Error GoTo 0
            '
            'Calculo SalMetro
            CONDI$ = "CODINT=" & XVALO(!cod_inte)
            M2# = XVALO(VALOBADA("Master", "PESUNI", CONDI$))
            SalMetEnt# = M2# * CantPri#
            SalMetRes# = M2# * CANTRES#
            '
            'completo informacion en campos
            PLADIASTO.AddNew 'agregar primer registro
            PLADIASTO!DepoUbic = XVALO(!DepoUbic)
            PLADIASTO!cod_inte = XVALO(!cod_inte)
            PLADIASTO!Codigo = TRIM$(Codigo$)
            PLADIASTO!FechAlta = !FechAlta
            PLADIASTO!CanAlta = XVALO(!CanAlta)
            PLADIASTO!UBICFIS = SAFETEXT(!UBICFIS)
            PLADIASTO!CODIEMPR = CODIEMPR%
            PLADIASTO!idenlote = SAFETEXT(!idenlote)
            PLADIASTO!SALMETRO = SalMetEnt#
            PLADIASTO!CanSaldo = CantPri#
            PLADIASTO!CanConsu = XVALO(!CanConsu)
            CONDI$ = "COD_INTE = " & XVALO(!cod_inte)
            CONDI$ = CONDI$ & " AND IdSubOr ='" & SAFETEXT(!idenlote) & "'"
            OF$ = VALOBADA("ORDEFABR", "Referen", CONDI$)
            PLADIASTO!NUORPROD = Left(OF$, 16)
            
            CONDI$ = "COD_INTE=" & XVALO(!cod_inte)
            CONDI$ = CONDI$ & " AND IDENLOTE='" & SAFETEXT(!idenlote) & "'"
            CONDI$ = CONDI$ & " AND CAUSA <> 0"
            CAUSAX% = XVALO(VALOBADA("MOVISTO", "CAUSA", CONDI$))
            PLADIASTO!CAUSA = CAUSAX%
            '
            PLADIASTO!CAJAS = CAJENT#
            PLADIASTO!EtiqxCaja = XVALO(!EtiqxCaja)
            PLADIASTO!ROLXCAJA = CantRolxCajEnt#
            PLADIASTO!EtiqxRollo = LOTES!EtiqxRollo
            PLADIASTO.Update
            '
            'SE CALCULA SI HAY CAJA CON RESTO
             If CANTRES# > 0 Then
                'AGREGAR REGISTRO POR LA DIFERENCIA
                PLADIASTO.AddNew
                 'completo informacion en campos
                PLADIASTO!CODIEMPR = CODIEMPR%
                PLADIASTO!idenlote = SAFETEXT(!idenlote)
                PLADIASTO!SALMETRO = SalMetRes#
                PLADIASTO!CanSaldo = CANTRES#
                PLADIASTO!CanConsu = XVALO(!CanConsu)
                PLADIASTO!NUORPROD = Left(OF$, 16)
                PLADIASTO!DepoUbic = XVALO(!DepoUbic)
                PLADIASTO!cod_inte = XVALO(!cod_inte)
                PLADIASTO!Codigo = TRIM$(Codigo$)
                PLADIASTO!FechAlta = !FechAlta
                PLADIASTO!CanAlta = XVALO(!CanAlta)
                PLADIASTO!UBICFIS = SAFETEXT(!UBICFIS)
                '
                PLADIASTO!CAJAS = 1
                PLADIASTO!EtiqxCaja = CANTRES#
                PLADIASTO!ROLXCAJA = CantRolXCajRes#
                PLADIASTO!EtiqxRollo = XVALO(!EtiqxRollo)
                PLADIASTO.Update
             End If
         End If
50       .MoveNext
      Loop
      PLADIASTO.Close
      .Close
    End With
    Call FINAL("?PLADIASTO")
99  Screen.MousePointer = 1
 
End Sub

Sub LOTES_VS_SALDOS()
    Dim SASTO(32767), SALOT(32767)
    '
    TMTX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
       GoTo 99
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
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Screen.MousePointer = 11
    FIN& = CantRegistros&("MOVISTO", "Cod_Inte > 0")
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK)"
    '
    Set MoviTemp = READSET(SQLX$)
    CAMOVI& = 0: JJ& = 0: CIMAX% = 0
    With MoviTemp
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           CI% = !cod_inte
           If CI% > 0 Then
              SASTO(CI%) = SASTO(CI%) + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
              If TRIM$(!idenlote) <> "" Then
                SALOT(CI%) = SALOT(CI%) + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
              End If
              If CI% > CIMAX% Then CIMAX% = CI%
           End If
         .MoveNext
         Loop
    End With
    '
    On Error Resume Next
    MoviTemp.Close
    Set MoviTemp = Nothing
    On Error GoTo 0
    '
    FIN& = CIMAX%: JJ& = 0: REBLA$ = REGBLAN$("COSTNEG")
    For CI% = 1 To CIMAX%
       JJ& = JJ& + 1
       Call TRACE(20, JJ&, FIN&)
       If TMT% < 1 Or TIMATE%(CI%) = TMT% Then
         CODE$ = CODEXT$(CI%)
         'If STATUIT%(CI%) > 0 Then
            If CODE$ >= COD1$ Then
              If CODE$ <= COD2$ Then
                SALPLO = SALPORLOTES(CI%)
                If Abs(SALPLO - SASTO(CI%)) > 0.05 Then
                  KK& = KK& + 1
                  XX$ = REBLA$
                  Call REPLA(XX$, MKI$(CI%), 1, 2)
                  Call REPLA(XX$, Unimed$(CI%), 3, 2)
                  Call REPLA(XX$, MKS$(SASTO(CI%)), 5, 4)
                  Call REPLA(XX$, MKS$(SALOT(CI%)), 21, 4)
                  Call REPLA(XX$, MKS$(SASTO(CI%) - SALOT(CI%)), 25, 4)
                  Call GRAREG("COSTNEG", XX$, KK&)
                End If
              End If
            End If
         'End If
       End If
    Next CI%
     
    Call SETULTREG("COSTNEG", KK&)
    Call FILESORT("COSTNEG", "", 1, 1)
    '
    Call HEADERS("COSTNEG", "")
    Call HEADERS("COSTNEG", "Tipo de Material: " + TIMAT$)
    '
    Screen.MousePointer = 1
    '
    If KK& < 1 Then
         Call COMUNI("No se Detectaron Diferencias entre Stock Total\y Stock por Lotes Para el Tipo de Material Elegido.")
      Else
         Call FINAL("*LISTLOT")
    End If
    '
99  Screen.MousePointer = 1

End Sub


Sub CONTROL_SOBREEMBARQUE()
   JJ& = 1
   FIN& = CantRegistros&("LOTINGRE", "COD_INTE > 0", "NOMESS")
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
   Set LOTINGTEMP = READSET(SQLX$)
   '
   With LOTINGTEMP
     '
     Do While Not .EOF
       Call TRACE(24, JJ&, FIN&)
       CI% = XVALO(!cod_inte)
       LOTTE$ = SAFETEXT$(!idenlote)
       Call ACSALOTE(CI%, LOTTE$)
       JJ& = JJ& + 1
     .MoveNext
     Loop
   End With
   '
    On Error Resume Next
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    On Error GoTo 0
   '
   Call MENSERR(24, "Fin de Control de Sobre-Embarque")
99 Screen.MousePointer = 1
   Command8.Enabled = True
End Sub
Function SALPORLOTES(CIXI%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    Set LOTINGTEMP = READSET(SQLX$)
    '
    SALPLO = 0
    With LOTINGTEMP
         Do While Not .EOF
             SALFIS = XVALO(!CanAlta) - XVALO(!CanConsu): If SALFIS < 0 Then SALFIS = 0
             SALPLO = SALPLO + SALFIS
         .MoveNext
         Loop
    End With
    '
    SALPORLOTES = SALPLO
    '
End Function


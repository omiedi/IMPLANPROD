VERSION 5.00
Begin VB.Form QUERSTK04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Consultas y Listados"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8745
   Icon            =   "QUERSTK04.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   7575
      Left            =   7875
      ScaleHeight     =   7515
      ScaleWidth      =   1635
      TabIndex        =   31
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command5 
         Caption         =   "Disp."
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
         Picture         =   "QUERSTK04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Sub-Sistema de Disponibilidad General de Stocks e Insumos"
         Top             =   2000
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Statist"
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
         Picture         =   "QUERSTK04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Modulo de Estadísticas de Stock"
         Top             =   2690
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
         Picture         =   "QUERSTK04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   300
         Width           =   650
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
         Left            =   105
         Picture         =   "QUERSTK04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   990
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   32
         Top             =   3675
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "QUERSTK04.frx":0D72
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3990
         Width           =   1515
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
      Picture         =   "QUERSTK04.frx":2C94
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Análisis Estadístico de Consumos de Stock"
      Top             =   6090
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Materiales Reservados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   24
      Top             =   1680
      Width           =   2640
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   25
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
            Height          =   600
            Left            =   0
            Picture         =   "QUERSTK04.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
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
            Left            =   630
            TabIndex        =   27
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Inventario Continuo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   20
      Top             =   220
      Width           =   2640
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   21
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
            Height          =   600
            Left            =   0
            Picture         =   "QUERSTK04.frx":32A8
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Material en Depósitos"
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
            TabIndex        =   23
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   16
      Top             =   3165
      Width           =   2640
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   17
         Top             =   420
         Width           =   1800
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
            Picture         =   "QUERSTK04.frx":35B2
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Parte Diario de Fabricacion por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Productos en Proceso"
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
            TabIndex        =   18
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Control de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   240
         TabIndex        =   10
         Top             =   1890
         Width           =   4215
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0FF&
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
            Left            =   420
            TabIndex        =   13
            Top             =   1320
            Width           =   3375
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
               Left            =   1890
               TabIndex        =   4
               Top             =   250
               Width           =   1275
            End
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
               Left            =   315
               TabIndex        =   3
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   11
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
               Height          =   600
               Left            =   1660
               Picture         =   "QUERSTK04.frx":38BC
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Control Diagnóstico de Lotes Vs. Saldos."
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
               Height          =   600
               Left            =   2215
               Picture         =   "QUERSTK04.frx":3BC6
               Style           =   1  'Graphical
               TabIndex        =   6
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
               Picture         =   "QUERSTK04.frx":3ED0
               Style           =   1  'Graphical
               TabIndex        =   7
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
               Picture         =   "QUERSTK04.frx":41DA
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Existencias Actuales: Cantidades/Valor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias Negativos Críticos"
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
               TabIndex        =   12
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0FF&
         Caption         =   "Consultas / Listados Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   14
            Top             =   420
            Width           =   3375
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
               Picture         =   "QUERSTK04.frx":44E4
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
               Height          =   600
               Left            =   2760
               Picture         =   "QUERSTK04.frx":47EE
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Gestion de Reservas"
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
               Left            =   2200
               Picture         =   "QUERSTK04.frx":4AF8
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ficha Kardex - Histórico de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - His- torico - Reservas"
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
               TabIndex        =   15
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
   Begin VB.Label Label2 
      Caption         =   "Acceso a Estadisticas de Stocks"
      Height          =   750
      Left            =   1050
      TabIndex        =   30
      Top             =   6090
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Acceso a Disponibilidad General de Stocks"
      Height          =   645
      Left            =   945
      TabIndex        =   29
      Top             =   5040
      Width           =   1485
   End
End
Attribute VB_Name = "QUERSTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

 Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("ORFABR04/OPRESERV", "Gestión de Ordenes de Fabricación")
   Command14.Enabled = True
End Sub

 Sub Command19_Click()
    Command19.Enabled = False
    '\\\OT-05-0379-WAR-06/05/05///
    XX1% = COMALTER%("Opciones de Listado: \\Solo Cantidades\Valorizado a última Compra")
    Select Case XX1%
      Case 1
        Call EXISTO2
      Case 2
        Call EXISTO4
    End Select
    '\\\OT-05-0379-WAR-FIN///
    Command19.Enabled = True
End Sub

 Sub Command2_Click()
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

 Sub Command27_Click()
    Command27.Enabled = False
    '\\\OT-5-0428-FG-28/06/05///
    ESTASTK04.Show 1
    Command27.Enabled = True
    Exit Sub
    '\\\OT-5-0428-FG-28/06/05///
    'Call CONECRUN("ESTOCK04", "Modulo de Estadísticas de Stock")
    HOII% = HOY(HOS$)
10  OBX$ = OBJPLA$("DESHESFECHREC", MKI$(HOII%) + MKI$(HOII%))
    If OBX$ <> "" Then
      DES% = CVI(Left$(OBX$, 2))
      HAS% = CVI(Mid$(OBX$, 3, 2))
      If DES% <= HAS% Then
        COMOVI$ = TRIM$(UCase$(InputBox$("Codigo de Movimiento", "Selección Tipo de Movimiento")))
        Screen.MousePointer = 11
        Call ABRESTOCKS
        DESDE = CDbl(Int(CVDAT(DES%)))
        HASTA = CDbl(Int(CVDAT(HAS%)))
        SQLX$ = "SELECT * FROM MOVISTO "
        SQLX$ = SQLX$ + "WHERE Cod_Inte >0 "
        SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) >= " & DESDE & " "
        SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
        If COMOVI$ <> "" Then
          SQLX$ = SQLX$ + "AND CodiMovi = '" & COMOVI$ & "' "
        End If
        SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        '
        JJ& = 0
        Call COPYSTRU("COKARDEX", "LIMOVSTO")
        REBLA$ = REGBLAN$("COKARDEX")
        Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
        Screen.MousePointer = 11
        With MoviTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              YY$ = REBLA$
              Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
              Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
              Call REPLA(YY$, !IdenLote, 5, 16)
              Call REPLA(YY$, !CodiMovi, 21, 2)
              Call REPLA(YY$, !DoPriCor, 23, 10)
              Call REPLA(YY$, !DoSecCor, 33, 10)
              Call REPLA(YY$, !ReferCor, 43, 30)
              Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
                MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
              Call REPLA(YY$, MONEX$, 81, 2)
              Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
              Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
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
        Set MoviTemp = Nothing
        Movisto.Close
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
    Command27.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Sub Command3_Click()
    COMMAND3.Enabled = False
    Call CONECRUN("ORFABR04/PRODUPRO", "Gestión de Ordenes de Fabricación")
99  COMMAND3.Enabled = True
End Sub
'
Private Sub Command4_Click()
   Command4.Enabled = False
   'Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    DISPSTK04.Show 1
    Command5.Enabled = True
End Sub

 Sub Command30_Click()
    Command30.Enabled = False
    Call SETULTREG("NEGASTO", 0)
    OPX% = COMALTER%("Listado de Existencias Negativas\\Listado General\Listado por Depósito")
20  If OPX% = 1 Then
        Call NEGASTO2
       ElseIf OPX% = 2 Then
        Call NEGASTO3
    End If
    '
98  If ULTREG&("NEGASTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*NEGASTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*NEGASTL", "")
        End If
        If OPX% = 2 Then
          Call SETULTREG("NEGASTO", 0)
          GoTo 20
        End If
    End If
99  Command30.Enabled = True
End Sub

Sub Command6_Click()
    Command6.Enabled = False
    Call CONECRUN("INVENT04", "Inventarios Físicos y Valorizados")
    Command6.Enabled = True
End Sub

 Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("ORFABR04/OPRESERV", "Gestión de Ordenes de Fabricación")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    
    If EXISTE%("TRAZAB04.EXE") < 1 Then
       Call COMUNI("Opción Solo Disponible Para Versión con Trazabilidad")
       GoTo 99
    End If
    
    OPX% = ALTERNA%("Diagnóstico Lote Vs. Saldos\Control de Sobre-Embarque")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 2 Then GoTo 98
    
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
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO"
    '
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    CAMOVI& = 0: JJ& = 0: CIMAX% = 0
    With MoviTemp
       On Error Resume Next
       '.MoveLast
       If Err < 1 Then
         On Error GoTo 0
         FIN& = .RecordCount
         .MoveFirst
         Do While Not .EOF
           JJ& = JJ& + 1
           Call TRACE(20, JJ&, FIN&)
           CI% = !COD_INTE
           If CI% > 0 Then
              SASTO(CI%) = SASTO(CI%) + !CANTINGRE - !CANTSALID
              If TRIM$(!IdenLote) <> "" Then
                SALOT(CI%) = SALOT(CI%) + !CANTINGRE - !CANTSALID
              End If
              If CI% > CIMAX% Then CIMAX% = CI%
           End If
         .MoveNext
         Loop
       End If
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
                  Call REPLA(XX$, UNIMED$(CI%), 3, 2)
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
    GoTo 99
    '
98 JJ& = 1
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   On Error Resume Next
   LotIngTemp.MoveLast
   On Error GoTo 0
   FIN& = LotIngTemp.RecordCount
   With LotIngTemp
    
     On Error Resume Next
     .MoveFirst
     If Err Then
         On Error GoTo 0
     End If
     '
     Do While Not .EOF
       Call TRACE(24, JJ&, FIN&)
       CI% = !COD_INTE
       LOTTE$ = !IdenLote
       Call ACSALOTE(CI%, LOTTE$)
       JJ& = JJ& + 1
     .MoveNext
     Loop
   End With
   '
    On Error Resume Next
    LotIngTemp.Close
    Set LotIngTemp = Nothing
    On Error GoTo 0
   '
   Call MENSERR(24, "Fin de Control de Sobre-Embarque")
99 Screen.MousePointer = 1
   Command8.Enabled = True
End Sub

 Sub Command9_Click()
    Command9.Enabled = False
    Call SITSTOK04.Command3_Click
    Command9.Enabled = True
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
    '
    If EXISTE%("TRAZAB04.EXE") > 0 Or INTERPRE% = 1 Then
       Command8.Visible = True
       Label12.Left = 640
    End If
    '
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

'\\\OT-05-0379-WAR-06/05/05///
Sub EXISTO4()
Dim PUN1#(32767), MOPUN1%(32767), FEPUN1#(32767)
    '
    TMTX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%) + Chr$(DPSX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD4", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    '
    OPX% = COMALTER%("Opciones de Inclusión:\\Solo con Stock\Todos los Artículos")
    '
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
    Call HEADERS("NEGASVAL", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    j& = 0
    Screen.MousePointer = 11
    NVII& = NUINV%
    '
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 " 'NO ANULADA
    Set STKVALTMP = compras.OpenRecordset(SQLX$, 4)
    With STKVALTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 11
      End If
      On Error GoTo 0
      .MoveLast
      TOT& = .RecordCount
      .MoveFirst
      RA& = 0
      Do While .EOF = False
        FEPUN1#(!COD_INTE) = CDbl(!Femi_Ocom)
        RA& = RA& + 1
        Call TRACE(20, RA&, TOT&)
        .MoveNext
      Loop
      .MoveFirst
      RA& = 0
      Do While .EOF = False
        If FEPUN1#(!COD_INTE) <= CDbl(!Femi_Ocom) Then
          FEPUN1#(!COD_INTE) = CDbl(!Femi_Ocom)
          MOPUN1%(!COD_INTE) = !Mone_Emis
          PUN1#(!COD_INTE) = !PrUn_Neto '!Pre_Unit
        End If
        RA& = RA& + 1
        Call TRACE(20, RA&, TOT&)
        .MoveNext
      Loop
    End With
    '
    y$ = REGBLAN("NEGASVAL")
    '
11  For I& = 1 To NUMAS%
      URE& = NUMAS%
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
            '
            If Abs(SSTT) > 0.05 Or OPX% = 2 Then
              Call REPLA(y$, MKI$(CIXI%), 1, 2)
              Call REPLA(y$, MKS$(SSTT), 3, 4)
              Call REPLA(y$, MKD$(PUN1#(CII&)), 9, 8)
              Call REPLA(y$, MKI$(MOPUN1%(CII&)), 17, 2)
              Call REPLA(y$, MKI$(CVINT(FEPUN1#(CII&))), 19, 2)
              IMPORTE1# = SSTT * PUN1#(CII&)
              Call REPLA(y$, MKD$(IMPORTE1#), 21, 8)
              IMPORTE2# = IMPORTE1# * TICAMONE(MOPUN1%(CII&)) / 1
              Call REPLA(y$, MKD$(IMPORTE2#), 29, 8)
              j& = j& + 1
              Call GRAREG("NEGASVAL", y$, j&)
            End If
            '
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASVAL", j&)
    Call FILESORT("NEGASVAL", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        Call FINAL("*LISTVAL")
    End If
    '
End Sub
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
    '
    OPX% = COMALTER%("Opciones de Inclusión:\\Solo con Stock\Todos los Articulos")
    '
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
    j& = 0
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
            If Abs(SSTT) > 0.05 Or OPX% = 2 Then
              TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
              j& = j& + 1
              Call GRAREG("NEGASTO", TTN$, j&)
            End If
            '
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", j&)
    Call FILESORT("NEGASTO", "", 1, 1, "ASC")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
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
    j& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    '
    FIN& = NUMAS%
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      Call VESARRAS(CI%)
      SSTT = STOCKACT(CI%)
      If SSTT < (-0.05) Then
        TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
        j& = j& + 1
        Call GRAREG("NEGASTO", TTN$, j&)
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", j&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
        Call COMUNI("No se detectaron items con stock negativo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call COMUNI("Se Detectaron " + TRIM$(Str$(j&)) + " Casos\de Stocks Negativos.")
    End If
    '
End Sub
'


Sub NEGASTO3()    'DETERMINACION ITEMS CON STOCK NEGATIVO
    '
    j& = 0
    '
    Call SELECHO("TADEPOSI")
    DEPOSI% = XVALO(VALACT1$("TADEPOSI"))
    If DEPOSI% = 0 Then Exit Sub
    FIN& = NUMAS%
    '
    Screen.MousePointer = 11
    '
    For I& = 1 To NUMAS%
      Call TRACE(20, I&, FIN&)
      CI% = I&
      XXX = STODEPOS(CI%, DEPOSI%)
      'XXX = STOCKACT(CI%)
      If XXX < (-0.05) Then
        TTN$ = MKI$(CI%) + MKS$(XXX) + String$(4, 0)
        j& = j& + 1
        Call GRAREG("NEGASTO", TTN$, j&)
      End If
      
    Next I&
    '
    Call SETULTREG("NEGASTO", j&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
        Call COMUNI("No se Detectaron Items con Stock Negativo\en el Deposito " + TRIM(Str(DEPOSI%)) + ".\Pulse Cualquier Tecla Para Volver al Menu.")
      Else
        Call COMUNI("Se Detectaron " + TRIM$(Str$(j&)) + " Casos\de Stocks Negativos en el Deposito " + TRIM(Str(DEPOSI%)))
    End If
    '
End Sub

Sub CRITSTO2()    'DETERMINACION ITEMS CON STOCK SUB MINIMO
    '
    TMTX% = 0: DPSX% = 0
    VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(TMTX%) + Chr$(DPSX%)
    '
10  RACOD$ = OBJPLA$("DESHASCOD4", VDEF$)
    If RACOD$ = "" Then
       Exit Sub
    End If
    '
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
       TMT% = 0
       Call REPLA(RACOD$, Chr$(0), 5, 1)
       TIMAT$ = "Sin Filtro por Tipo de Material"
    End If
    DPS% = Asc(Mid$(RACOD$, 6, 1))
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    
    j& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    FIN& = NUMAS%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CI% = I&
      If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
         CODE$ = CODEXT$(CI%)
         If CODE$ < COD1$ Or CODE$ > COD2$ Then GoTo 90
         '
         Call VESARRAS(CI%)
         SSTT = STODEPOS(CI%, DPS%)
         '
         STMIN = STOMIN(CI%)
         If SSTT < STMIN Then
           '
           COU = COSUNI(CI%)
           REPOS = STMIN - SSTT
           If REPOS < LOTESTAN(CI%) Then REPOS = LOTESTAN(CI%)
           If LINDIVIS(CI%) > 0 Then
              CANAREPO = LINDIVIS(CI%)
              While CANAREPO < REPOS: CANAREPO = CANAREPO + LINDIVIS(CI%): Wend
              REPOS = CANAREPO
           End If
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
           j& = j& + 1
           Call GRAREG("CRITSTO", TTN$, j&)
           '
         End If
      End If
90  Next I&
    '
    Call SETULTREG("CRITSTO", j&)
    Call CIERRARCH("*.*")
    Call HEADERS("CRITSTO", "")
    Screen.MousePointer = 1
    '
    If j& < 1 Then
        Call COMUNI("No se detectaron items con stock sub-mínimo.\Pulse cualquier tecla para volver al menu.")
      Else
        Call HEADERS("CRITSTO", "Material: " + TIMAT$)
        Call COMUNI("Se Detectaron " + TRIM$(Str$(j&)) + " Casos\de Stocks Sub-Mínimo.")
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
    x% = COMALTER%(TXT$)
    On x% GoTo 320, 310
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
    j& = 0
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
            j& = j& + 1
            Call GRAREG("UBIEXI", YY$, j&)
            If PRIRE&(CI%) < 1 Then
              PRIRE&(CI%) = j&
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
            j& = j& + 1
            Call GRAREG("UBIEXI", YY$, j&)
            PRIRE&(CI%) = j&
        End If
      End If
    Next I&
    '
    Call SETULTREG("UBIEXI", j&)
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
      j& = 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        YY$ = REGLEIDO$("UBIEXI", U&)
        CIII% = CVI(Left$(YY$, 2))
        CODX$ = CODEXT$(CIII%)
        If CODX$ >= COD1$ Then
          If CODX$ <= COD2$ Then
            j& = j& + 1
            Call GRAREG("LUBIEXI", YY$, j&)
          End If
        End If
      Next U&
      Call SETULTREG("LUBIEXI", j&)
      Call CIERRARCH("LUBIEXI")
      Call CIERRARCH("*.*")
      Screen.MousePointer = 1
      Call FINAL("*SUBIEXI")
    End If
    '
End Sub
'



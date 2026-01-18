VERSION 5.00
Begin VB.Form TRANSTK04 
   BackColor       =   &H00E0E8FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias de Stocks"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   5100
      Left            =   4540
      ScaleHeight     =   5040
      ScaleWidth      =   1035
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   5100
         Left            =   0
         Picture         =   "TRANSTK04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   5265
      Left            =   4560
      ScaleHeight     =   5205
      ScaleWidth      =   2475
      TabIndex        =   21
      Top             =   0
      Width           =   2535
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   0
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   28
         Top             =   4200
         Width           =   850
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Kdx"
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
         Picture         =   "TRANSTK04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1800
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "TRANSTK04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3360
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "TRANSTK04.frx":2AE4
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2490
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
         Picture         =   "TRANSTK04.frx":2DEE
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Picture         =   "TRANSTK04.frx":2F38
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   990
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "TRANSTK04.frx":3242
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4455
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E8FF&
      Caption         =   "Operaciones"
      Height          =   1275
      Left            =   150
      TabIndex        =   11
      Top             =   210
      Width           =   4170
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   12
         Top             =   420
         Width           =   2535
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
            Picture         =   "TRANSTK04.frx":5164
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Informar Trans- ferencias de Stock"
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
            TabIndex        =   13
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E8FF&
      Caption         =   "Consultas - Listados"
      Height          =   3225
      Left            =   150
      TabIndex        =   9
      Top             =   1680
      Width           =   4170
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   2100
         TabIndex        =   17
         Top             =   420
         Width           =   1905
         Begin VB.CommandButton BOTSEL1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1575
            Picture         =   "TRANSTK04.frx":546E
            TabIndex        =   6
            Top             =   1260
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1575
            Picture         =   "TRANSTK04.frx":5778
            TabIndex        =   4
            Top             =   630
            Width           =   175
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   525
            TabIndex        =   3
            Text            =   " "
            Top             =   630
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   525
            TabIndex        =   5
            Text            =   " "
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label Label2 
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
            Left            =   0
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            TabIndex        =   18
            Top             =   1000
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Rango Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   1800
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   525
            TabIndex        =   2
            Top             =   1260
            Width           =   1065
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   525
            TabIndex        =   1
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            TabIndex        =   16
            Top             =   1000
            Width           =   750
         End
         Begin VB.Label Label9 
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
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E8FF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   10
         Top             =   2310
         Width           =   3795
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
            Left            =   3150
            Picture         =   "TRANSTK04.frx":5A82
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   95
            Width           =   645
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E8FF&
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
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E8FF&
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
            Left            =   1680
            TabIndex        =   8
            Top             =   315
            Width           =   1275
         End
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu informatransferenciadestock 
         Caption         =   "Informar Transferencia de Stock"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu listadodetransferenciasrealizadas 
         Caption         =   "Listado de Transferencias Realizadas"
         Begin VB.Menu pantalla 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora 
            Caption         =   "Impresora"
         End
      End
   End
End
Attribute VB_Name = "TRANSTK04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RUSANJUA$

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub BOTSEL1_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Command1_Click()
    Hide
    Call FRESHALL
End Sub

Private Sub Command2_Click()
    DESNUM& = 1: If XVALO(Text7.TEXT) > 0 Then DESNUM& = XVALO(Text7.TEXT)
    DESNUX$ = "TR." + TRIM$(Str$(DESNUM&))
    HASNUM& = 999999
    If TRIM$(Text8.TEXT) <> "" Then
        If XVALO(Text8.TEXT) >= DESNUM& Then HASNUM& = XVALO(Text8.TEXT)
    End If
    HASNUX$ = "TR." + TRIM$(Str$(HASNUM&))
    DESFE% = 0: If FECHANUM(TRIM$(Text2.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text2.TEXT))
    HASFE% = 32767
    If TRIM$(Text1.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text1.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text1.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call SETULTREG("LISTRAJ", 0)
    '
    FEC1# = CDbl(CVDAT(DESFE%))
    FEC2# = CDbl(CVDAT(HASFE%))
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  'TR' "
    SQLX$ = SQLX$ + " and DoPriCor >= '" & DESNUX$ & "' "
    SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & FEC1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & FEC2# & " "
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) ASC"
    '
    Set AJUSMOVTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With AJUSMOVTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Transferencias Registradas segun Filtro")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      Do While .EOF = False
      REFE$ = !ReferCor
      CI1% = !COD_INTE
      DP1% = !DEPOMOVI
      CAN# = !CANTINGRE
      FF% = CVINT(!FECHMOV)
      NRO& = Mid(!DoPriCor, 4)
      y$ = REGBLAN$("LISTRAJ")
      Call REPLA(y$, MKI$(FF%), 1, 2)
      Call REPLA(y$, REFE$, 3, 48)
      Call REPLA(y$, MKS$(NRO&), 51, 4)
      Call REPLA(y$, MKI$(CI1%), 55, 2)
      Call REPLA(y$, UNIMED$(CI1%), 57, 4)
      If CAN# > 0 Then
          Call REPLA(y$, MKD$(CAN#), 61, 8)
          Call REPLA(y$, MKD$(!CANTSALID), 69, 8)
        Else
          Call REPLA(y$, MKD$(CAN#), 61, 8)
          Call REPLA(y$, MKD$(!CANTSALID), 69, 8)
      End If
      Call REPLA(y$, Chr$(DP1%), 80, 1)
      Call REGAPP("LISTRAJ", y$)
      
      .MoveNext
9      Loop
    End With
    '
    Call CIERRARCH("LISTRAJ")
    Screen.MousePointer = 1
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        'Call OPNOIN("")
        Call COMUNI("Solo Disponible para Impresión. Genere el \ Listado y Exportelo a un Archivo de Texto")
        GoTo 10
      ElseIf OPCI% = 2 Then
10      Call CONECRUN("*LITRAST", "")
    End If
    '
End Sub

Private Sub Command26_Click()
   'Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   OPTRAJU04.Show 1

   Call FINAL("")
   'OPMOMAQ04.Show 1
End Sub

Sub Command27_Click()
    Command27.Enabled = False
    OPX% = ALTERNA%("Control de Depósitos\Otros Accesos Directos")
    If OPX% = 1 Then
          Call CONDEPOPRE
       ElseIf OPX% = 2 Then
          Call ACCDIR("TRANSFE")
    End If
    Command27.Enabled = True
End Sub

Sub CONDEPOPRE()
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    F1# = CDbl(CVDAT(DES%))
    F2# = CDbl(CVDAT(HAS%))
    If F1# > F2# Then
      Call COMUNI("Rango de Fecha no valido")
      GoTo 10
    End If
    '
    Call FRESHALL
    Call ABRESTOCKS
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CDBL( FechMov ) >= " & F1# & " "
    SQLX$ = SQLX$ + " AND CDBL( FechMov ) <= " & F2# & " "
    SQLX$ = SQLX$ + "AND DepoMovi = 99 "
    '
    Set RST1TMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    With RST1TMP
      On Error Resume Next
      .MoveLast
      If Err Then
        Call COMUNI("No Existen Movimientos Recuperables\en el Rango de Fechas Seleccionado.")
        Screen.MousePointer = 1
        Exit Sub
      End If
      '
      ULT1& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      '
      Do While Not .EOF
         CIXI% = !COD_INTE
         .Edit
         !DEPOMOVI = DEPOPRE%(CIXI%)
         .Update
      .MoveNext
      Loop
    End With
    '
    Screen.MousePointer = 1
    Call COMUNI("Control Completo")
    '
End Sub
'
Sub Command3_Click()
    Call TRASTOCK
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call SITSTOK04.Command3_Click
    'Call CONECRUN("*KARDEX", "")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call OPNOIN("")
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("TRANSFE")
End Sub

Private Sub Form_Load()
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      Me.Option5.BackColor = &HFCD7B6
      MARCOBARRA.Top = 4500
      Set MARCOBARRA.Container = Me.Picture6
      
      
      MARCOBARRA.ZOrder 0
       
       
  End If

   Call TRANSLABELS(Name)
   '
End Sub

Private Sub impresora_Click()
   Option4.Value = True
   Call Command2_Click
End Sub

Private Sub pantalla_Click()
    Option5.Value = True
    Call Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

'
Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Sub TRASTOCK()
    '
    Static CONTRAZA%
    '
    If CONTRAZA% = 0 Then
       CONTRAZA% = (-1)
       If EXISTE%("TRAZAB04.EXE") > 0 Then
          CONTRAZA% = 1
       End If
    End If
    '
    HOOI% = HOY(HO$)
    FORIPRE$ = TRIM$(CONTROL$("", "FORTRAST"))
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Transferencia:\  \'" + RMCC$ + "'.")
        Exit Sub
      End If
    End If
    '
100 CMOV$ = TRIM$(CONTROL$("", "CODTRAST"))
    If CMOV$ = "" Then CMOV$ = "TR"
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMOV$ & "'"
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC "
    '
    UNUM& = 0
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    MoviTemp.MoveFirst
    If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
    End If
    On Error GoTo 0
    NRO& = 1 + UNUM&
    MoviTemp.Close
    '
    TIPODOC$ = "Vale Transferencia de Stock"
    NUMEDOC$ = TRIM$(Str$(NRO&))
    '
110 Screen.MousePointer = 1
    '
    VDEF$ = String$(6, 0) + MKI$(HOOI%) + Space$(16) + MKS$(0) + Space$(32)
    OPX% = COMALTER%("Opciones de Procesamiento:\\Item único\Lista Múltiple\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 2 Then
      '\\\OT-5-0447-WAR-08/07/05///
      FORIPRE$ = TRIM$(CONTROL$("", "FORTRSTM"))
      If FORIPRE$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
          FORIPRE$ = ""
        End If
      End If
      '
115   x$ = OBJPLA$("TRANSTOMUL", VDEF$)
      If x$ = "" Then
        Exit Sub                        ' abandonar
      End If
      '
      DEPORIG% = Asc(Mid$(x$, 3, 1)): DEPOSI$ = ECOARCH$("ECODEP", Chr$(DEPORIG%))
      DEPODES% = Asc(Mid$(x$, 6, 1)): DEPODE$ = ECOARCH$("ECODEP", Chr$(DEPODES%))
      FEMOVI% = CVI(Mid$(x$, 7, 2))
      REFMO$ = TRIM$(Mid$(x$, 29, 32))
      If REFMO$ = "" Then
        Call MENSERR(24, "Referencia o Motivo es Mandatoria")
        VDEF$ = x$
        GoTo 115
      End If
      '
1112  Call SETULTREG("!BOLREDET", 0)
      Call CIERRARCH("*.*")
      '
      Screen.MousePointer = 11
      '
      FFF$ = REPLACAR$(FECHATEX$(FEMOVI%), "/", "-")
1113  VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Transferencia de Stock - De " + TRIM$(DEPOSI$) + " a " + TRIM$(DEPODE$) + " - " + FFF$)
      If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
         GoTo 110
      End If
      '
      CONTA% = ULTREG&("!BOLREDET")
      '
1115  If CONTA% > 0 Then
        Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
        '
        '\\\OT-05-0730-WAR-10/11/05///
        CAITMO% = 0
        For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                   '
                   If CONTRAZA% = 1 Then
                      LOTE$ = TRIM$(Mid$(YY$, 69, 12))
                        If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
                      If LOTE$ = "" Then
                        Call MENSERR(24, "Falta Identificar Lote\en Código '" + TRIM$(CODEXT$(CI%)) + "'")
                        GoTo 1113
                      End If
                      '
                      If SALOTE(CI%, LOTE$, 2) < CAN# - 0.05 Then
                        Call MENSERR(24, "Saldo Insuficiente en\Código '" + TRIM$(CODEXT$(CI%)) + "' - Lote '" + LOTE$ + "'.")
                        GoTo 1113
                      End If
                   End If
                   '
                 End If
               End If
             End If
        Next YX&
        If PUEDEMOVI%(DEPORIG%) < 0 Then GoTo 1113 ' CANCELA SALIDA
        '\\\OT-05-0730-WAR-FIN///
        '
        AINDI% = 0
        NUORIT% = (-1)
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          LOTE$ = Mid$(YY$, 69, 12)
            If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
          REMI$ = "TR." + TRIM$(Str$(NRO&))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                DOSEC$ = ""
1120            CANTI = (-1) * CAN#
                NUORIT% = NUORIT% + 1
                DPX1% = DEPORIG%: If DEPORIG% = 99 Then DPX1% = DEPOPRE%(CI%)
                Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, 0, 0, 0, DPX1%, CANTI)
                NUORIT% = NUORIT% + 1
                DPX2% = DEPODES%: If DEPODES% = 99 Then DPX2% = DEPOPRE%(CI%)
                Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, 0, 0, 0, DPX2%, (-1) * CANTI)
              End If
            End If
          End If
        Next YX&
        Call CIERRARCH("MOVISTO")
         '\\\OT-5-0447-WAR-08/07/05///
        'IMPRESION DEL VALE DE TRANSFERENCIA MULTIPLE
        If FORIPRE$ <> "" Then
          CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = CStr(NRO&)
          CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHATEX(FEMOVI%)
          CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSI$
          CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DEPODE$
          CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = REFMO$
          CAPARDOC% = 5
          '
          CODTABU1$(1) = "COD-MPRI"
          CODTABU1$(2) = "DES-MPRI"
          CODTABU1$(3) = "UNI-MPRI"
          CODTABU1$(4) = "CAN-MPRI"
          CACOLTAB1% = 4
          '
          CAITAB1% = 0
          '
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            '
            CAITAB1% = CAITAB1% + 1
              CODMAT$ = TRIM$(CODEXT(CI%))
            TETABU1$(1, CAITAB1%) = CODMAT$
              DESMAT$ = DESCRIT$(CI%)
            TETABU1$(2, CAITAB1%) = DESMAT$
              UNIMAT$ = UNIMED$(CI%)
            TETABU1$(3, CAITAB1%) = UNIMAT$
              CANTMAT$ = FORMATNUM(CAN#, "F12.2")
            TETABU1$(4, CAITAB1%) = CANTMAT$
          Next YX&
          '
          Call PRINTDOC("FORTRSTM")
        End If
        '\\\OT-5-0447-WAR-FIN///
        Call COMUNI("Transferencia Procesada")
        '
      End If
      GoTo 110
    End If
    '
140 Screen.MousePointer = 1
150 x$ = OBJPLA$("TRANSFESTOCK", VDEF$)
    If x$ = "" Then
        Exit Sub                        ' abandonar
    End If
    '
    VDEF$ = x$
    TRANSFE$ = x$
    '
    CI1% = CVI(Left$(TRANSFE$, 2))
    DP1% = Asc(Mid$(TRANSFE$, 3, 1))
    CI2% = CVI(Mid$(TRANSFE$, 4, 2))
    DP2% = Asc(Mid$(TRANSFE$, 6, 1))
    FECHA% = CVI(Mid$(TRANSFE$, 7, 2))
      If FECHA% < 1 Then FECHA% = HOY(HO$)
    LOTE$ = Mid$(TRANSFE$, 9, 16)
    CAN = CVS(Mid$(TRANSFE$, 25, 4))
    REFE0$ = TRIM$(Mid$(TRANSFE$, 29, 32))
    '
    If CI1% > 0 Then
      If CI1% <= NUMAS% Then
        If DP1% > 0 Then
           If CI2% > 0 Then
              If CI2% <= NUMAS% Then
                If DP2% > 0 Then
                   If CAN >= 0.1 Then
                      GoTo 160
                   End If
                End If
              End If
           End If
        End If
      End If
    End If
    '
    Call MENSERR(24, "Datos Incompletos !!!\Transferencia Abortada.")
    GoTo 150
    '
160 CAITMO% = 0
    '
    EXILO% = 0: CANMA = 0
    If CONTRAZA% = 1 Then            ' VERIFICA TRAZABILIDAD
       '
       If CI1% <> CI2% Then
          Call MENSERR(24, "En Transferencia con Trazabilidad\Deben Coincidir Código Origen y Destino.")
          GoTo 150
       End If
       '
       If TRIM$(LOTE$) = "" Then
          Call MENSERR(24, "Mandatorio Identificar el Lote de Origen.")
          GoTo 150
       End If
       '
       Call ARMALILOT(CI1%)
       For KU& = 1 To ULTREG&("!LILOTSE")
         XY$ = REGLEIDO$("!LILOTSE", KU&)
         If TRIM$(Left$(XY$, 16)) = TRIM$(LOTE$) Then
           GoTo 162
         End If
       Next KU&
       '
       Call MENSERR(24, "No se Ha Encontrado\el Lote de Origen Especificado '" + TRIM$(LOTE$) + "'")
161    XY$ = Left$(REGSEL$("!LILOTSE"), 16)
       Call REPLA(VDEF$, XY$, 9, 16)
       GoTo 150
       '
162    CANMA = SALOTE(CI1%, LOTE$, 2)
       If CAN > CANMA + 0.05 Then
         Call MENSERR(24, "La Cantidad Informada Supera\el Saldo Físico del Lote '" + TRIM$(LOTE$) + "'")
         GoTo 161
       End If
    End If
    '
    For YX& = 1 To 1
         CI% = CI1%
         CAN = CAN
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN > 0.005 Then
               CAITMO% = CAITMO% + 1
               CIMOVIM%(CAITMO%) = CI%
               CAMOVIM(CAITMO) = (-1) * CAN
             End If
           End If
         End If
    Next YX&
    If PUEDEMOVI%(DP1%) < 0 Then GoTo 150 ' CANCELA SALIDA
    '\\\OT-05-0730-WAR-FIN///
    NC% = 0
    VUNI# = 0
    MONE$ = ""
    '
    DOPRI$ = "TR." + TRIM$(Str$(NRO&))
    DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
                              While Len(DOSEC$) < 10
                                DOSEC$ = DOSEC$ + " "
                              Wend
                              '
    COD$ = LTrim$(RTrim$(CODEXT$(CI2%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. A " + COD$ + " - D." + Str$(DP2%)
    End If
    '
    CANTI = (-1) * CAN
    DPX1% = DP1%: If DP1% = 99 Then DPX1% = DEPOPRE%(CI1%)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 0, VUNI#, MONEI%, NC%, DPX1%, CANTI)
    '
    COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. DE " + COD$ + " - D." + Str$(DP1%)
    End If
    '
    CANTI = CAN
    DPX2% = DP2%: If DP2% = 99 Then DPX2% = DEPOPRE%(CI2%)
    Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 1, VUNI#, MONEI%, NC%, DPX2%, CANTI)
    Call CIERRARCH("MOVISTO")
    '
'IMPRIMIR FORMULARIO
237 If TRIM$(FORIPRE$) <> "" Then
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI1%)), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI2%)), "COD-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(DESCRIT$(CI1%), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(DESCRIT$(CI2%), "DES-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(UNIMED$(CI1%), "UNIMED", 0, YAC)
               Call PRINTFORWIN(UNIMED$(CI2%), "UNIMED", 0, YAC + INTERLI)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP1%)), "DEPOSITO", 0, YAC)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP2%)), "DEPOSITO", 0, YAC + INTERLI)
               Call PRINTFORWIN(LOTE$, "NRO-LOTE", 0, YAC)
               CANS$ = CSTRING$(MKS$(CAN), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(REFE0$, "REF-MAT1", 0, YAC)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Transferencia de Stock" + Str$(NRO&), OKX%)
    End If
    '
299 Call COMUNI("Transferencia Completa\y Verificada !!!")
    '
    TRANSFE$ = ""
    GoTo 100
    '
End Sub
'



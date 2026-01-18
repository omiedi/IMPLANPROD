VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORLIPRO07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Proveedores - Cuentas Corrientes y Gastos"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9225
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
   ScaleHeight     =   3420
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "AVANCE DE PROCESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2415
      Width           =   8085
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   7425
         Picture         =   "FORLIPRO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   130
         Width           =   640
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   375
         Width           =   7050
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C7D9CC&
      Caption         =   "RANGO DE CUENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   4200
      TabIndex        =   4
      Top             =   105
      Width           =   4005
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   315
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   7
         Text            =   " "
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   6
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   315
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   12
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   4005
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Todas las Cuentas (Corrientes y Gastos)"
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
         Height          =   555
         Left            =   420
         TabIndex        =   23
         Top             =   1560
         Width           =   2445
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Cuenta de Gastos"
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
         Height          =   435
         Left            =   420
         TabIndex        =   22
         Top             =   1050
         Width           =   3165
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Cuenta Corriente de Proveedoes"
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
         Height          =   435
         Left            =   420
         TabIndex        =   21
         Top             =   420
         Value           =   -1  'True
         Width           =   3165
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "ORDENAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   4005
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Código Postal"
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
         Height          =   360
         Index           =   3
         Left            =   2835
         TabIndex        =   15
         Top             =   380
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Alfabético"
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
         Height          =   240
         Index           =   2
         Left            =   1470
         TabIndex        =   2
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Numérico"
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
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   8340
      ScaleHeight     =   4605
      ScaleWidth      =   1035
      TabIndex        =   17
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   26
         Top             =   3000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   27
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Height          =   690
         Left            =   120
         Picture         =   "FORLIPRO07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "FORLIPRO07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   900
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FORLIPRO07.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORLIPRO07.frx":0992
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FORLIPRO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Call GELISPRO
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("PROVED1")
    NCMAX% = XVALO(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = rasocli$(NCMAX% * (-1))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("PROVED1")
    NCMIN% = XVALO(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
        TEXT10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = rasocli$(NCMIN% * (-1))
    End If
End Sub
   
Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    Dim Provetemp As ADODB.Recordset
    Set Provetemp = New ADODB.Recordset
    Provetemp.CursorType = adOpenKeyset
    Provetemp.LockType = adLockReadOnly
    '
    Call APERBASDAT("PROVE")
    pricli& = 999999: ultcli& = 0
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM PROVED12 "
    SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
    SQLX$ = SQLX$ + "ORDER BY Nume_Cli ASC"
    Provetemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    With Provetemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then pricli& = !nume_cli
      .MoveLast
      If Err < 1 Then ultcli& = !nume_cli
      On Error GoTo 0
    End With
    Provetemp.Close
    Set Provetemp = Nothing
    '
    If pricli& < 999999 Then TEXT10 = TRIM$(Str$(pricli&))
    If ultcli& > 0 Then Text8 = TRIM$(Str$(ultcli&))
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    If XVALO(TEXT10.TEXT) > 32767 Then
       TEXT10.TEXT = Left$(TRIM$(TEXT10.TEXT), 4)
       Exit Sub
    End If
    If XVALO(TEXT10.TEXT) < 1 Then TEXT10.TEXT = "": Exit Sub
    NCMIN% = XVALO(TEXT10.TEXT)
    Text9.TEXT = rasocli$((-1) * NCMIN%)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If XVALO(Text1.TEXT) < 256 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If XVALO(Text1.TEXT) < 999 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text8_Change()
    If XVALO(Text8.TEXT) > 32767 Then
       Text8.TEXT = Left$(TRIM$(Text8.TEXT), 4)
       Exit Sub
    End If
    If XVALO(Text8.TEXT) < 1 Then Text8.TEXT = "": Exit Sub
    NCMAX% = XVALO(Text8.TEXT)
    Text7.TEXT = rasocli$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        TEXT10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = rasocli$(NCMIN%)
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = rasocli$(NCMAX%)
    End If
End Sub

Private Sub Text9_Click()
    TEXT10.SetFocus
End Sub

Sub GELISPRO()
    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    '
    NCMIN = XVALO(TEXT10.TEXT)
    NCMAX = XVALO(Text8.TEXT)
    '
    If Option2.Value = True Then TIPOCTA% = 0
    If Option3.Value = True Then TIPOCTA% = 1
    If Option4.Value = True Then TIPOCTA% = -1
    
    '
    If NCMIN < 1 Then NCMIN = PrimerClienteActivo&("PROVEEDOR")
    If NCMAX < 1 Then NCMAX = UltimoClienteActivo&("PROVEEDOR")
    '
    If TIPOCTA% >= 0 Then
      'CON FILTRO POR TIPO DE CUENTA
     If EXISTESQR_ENSQL%("R-LIPROVE1") > 0 Then
      Call FINAL("?R-LIPROVE1(" & NCMIN & "," & NCMAX & "," & TIPOCTA% & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If
    Else
     If EXISTESQR_ENSQL%("R-LIPROVE") > 0 Then
      Call FINAL("?R-LIPROVE(" & NCMIN & "," & NCMAX & ")") '
      Screen.MousePointer = 1
     Else
      GoTo 10
     End If
    End If
    Exit Sub
    
10  If Option2.Value = True Then
      MODFILT% = 0
     ElseIf Option3.Value = True Then
      MODFILT% = 1
     ElseIf Option4.Value = True Then
      MODFILT% = 2
    End If
    '
    MODORDE% = 1
    If FORLIPRO07.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLIPRO07.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    SQLX$ = "SELECT * FROM Proved12 WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX
    '
    If MODFILT% < 2 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
        Case 1
        SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
        Case 2
        SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
        Case 3
        SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    '
    Dim PROVTEMP As ADODB.Recordset
    Set PROVTEMP = New ADODB.Recordset
    Set PROVTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With PROVTEMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         Screen.MousePointer = 1
         On Error GoTo 0
         GoTo 90
       End If
       FIN& = .RecordCount
       On Error GoTo 0
    End With
    '
    With PROVTEMP
        .Move first
        Do While (.EOF = False)
            NP% = XVALO(!nume_cli)
            NC% = (-1) * NP%
             '
            Label3 = rasocli$(NC%)
            Label3.Refresh
            Y$ = REGBLAN$("LIPROVE")
            Call REPLA(Y$, MKI$(NP%), 1, 2)
            Call REPLA(Y$, Label3, 3, 32)
            Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
            Call REPLA(Y$, CuitCli$(NC%), 51, 16)
            Call REPLA(Y$, POSIVAC$(NC%), 67, 10)
            CUECOX% = TICUE%(NP%)
            Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
            Call REPLA(Y$, DETICUE$(CUECOX%), 79, 32)
            Call REPLA(Y$, CUEMADRE$(NC%), 111, 16)
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            If TICUE%(NP%) = 0 Then
               Y$ = REGBLAN$("LIPROVE")
               Call REPLA(Y$, MKI$(NP%), 1, 2)
               Call REPLA(Y$, DOMICLI$(NC%), 3, 32)
               'Call REPLA(Y$, FAXCLI$(NC%), 35, 16)
               Call REPLA(Y$, NIBRCLI$(NC%), 51, 16)
               CTRX$ = TRIM$(CATIBRU$(NC%))
               CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                                    If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                                    If CTRX$ = "2" Then CATIBX$ = "C.Multil."
               Call REPLA(Y$, CATIBX$, 67, 10)
               CPAG% = CPAGCLI%(NC%)
               Call REPLA(Y$, MKI$(CPAG%), 77, 2)
               Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
               ULI& = ULI& + 1
               Call GRAREG("LIPROVE", Y$, ULI&)
               '
               Y$ = REGBLAN$("LIPROVE")
               Call REPLA(Y$, MKI$(NP%), 1, 2)
               DOMI$ = TRIM$(CPOSCLI$(NC%))
               If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
               DOMI$ = DOMI$ + LOCACLI$(NC%)
               Call REPLA(Y$, DOMI$, 3, 32)
               'Call REPLA(Y$, EMAICLI$(NC%), 35, 16)
               ULI& = ULI& + 1
               Call GRAREG("LIPROVE", Y$, ULI&)
            End If
            '
            Y$ = REGBLAN$("LIPROVE")
            ULI& = ULI& + 1
            Call GRAREG("LIPROVE", Y$, ULI&)
            '
            .MoveNext
        Loop
    End With
    '
    Call SETULTREG("LIPROVE", ULI&)
    Call CIERRARCH("LIPROVE")
    Screen.MousePointer = 1
    Call FINAL("*LIPROVE")
    '
    PROVTEMP.Close
    Set PROVTEMP = Nothing

90 End Sub
'

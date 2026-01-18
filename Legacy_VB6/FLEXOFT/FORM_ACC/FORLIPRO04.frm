VERSION 5.00
Begin VB.Form FORLIPRO04 
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listados de Proveedores - Cuentas Corrientes y Gastos"
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
   Begin VB.PictureBox Picture6 
      Height          =   4635
      Left            =   8340
      ScaleHeight     =   4575
      ScaleWidth      =   1035
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   3800
         Left            =   0
         Picture         =   "FORLIPRO04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   14
      Top             =   2415
      Width           =   4005
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   3350
         Picture         =   "FORLIPRO04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   105
         Width           =   640
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   16
         Top             =   375
         Width           =   2850
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
      BackColor       =   &H00D8D8C0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
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
         Height          =   360
         Left            =   3465
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
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
         Height          =   360
         Left            =   3465
         TabIndex        =   6
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
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
         Left            =   210
         TabIndex        =   12
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         Left            =   210
         TabIndex        =   11
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         BackColor       =   &H00D8D8C0&
         Caption         =   "Todas las Cuentas (Corriente y de Gasto)"
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
         Left            =   420
         TabIndex        =   25
         Top             =   1680
         Width           =   3165
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Cuentas de Gastos"
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
         Left            =   420
         TabIndex        =   24
         Top             =   1050
         Width           =   3165
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Cuentas Corrientes de Proveedores"
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
         Left            =   420
         TabIndex        =   23
         Top             =   420
         Width           =   3165
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   0
      Top             =   2415
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D8D8C0&
         Caption         =   "C.Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2835
         TabIndex        =   15
         Top             =   420
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Alfabético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1470
         TabIndex        =   2
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D8D8C0&
         Caption         =   "Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      BackColor       =   &H00800000&
      Height          =   4635
      Left            =   8340
      ScaleHeight     =   4575
      ScaleWidth      =   1005
      TabIndex        =   17
      Top             =   0
      Width           =   1065
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
         Height          =   690
         Left            =   120
         Picture         =   "FORLIPRO04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
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
         Height          =   640
         Left            =   120
         Picture         =   "FORLIPRO04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   19
         Top             =   2520
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   20
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "FORLIPRO04.frx":2AF6
         Stretch         =   -1  'True
         Top             =   2835
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FORLIPRO04"
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
    'nuevo
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
        Text7.TEXT = RASOCLI$(NCMAX% * (-1))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("PROVED1")
    NCMIN% = XVALO(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN% * (-1))
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      FORLIPRO04.Frame1.BackColor = &HFCD7B6
      FORLIPRO04.Frame2.BackColor = &HFCD7B6
      FORLIPRO04.Frame4.BackColor = &HFCD7B6
      FORLIPRO04.Frame5.BackColor = &HFCD7B6
      FORLIPRO04.Option1(1).BackColor = &HFCD7B6
      FORLIPRO04.Option1(2).BackColor = &HFCD7B6
      FORLIPRO04.Option1(3).BackColor = &HFCD7B6
      FORLIPRO04.Option2.BackColor = &HFCD7B6
      FORLIPRO04.Option3.BackColor = &HFCD7B6
      FORLIPRO04.Option4.BackColor = &HFCD7B6
      
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    Call APERBASDAT("PROVE")
    PRICLI& = 999999: ULTCLI& = 0
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM PROVED12 "
    SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
    SQLX$ = SQLX$ + "ORDER BY Nume_Cli ASC"
    Set PROVETEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PROVETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then PRICLI& = !Nume_Cli
      .MoveLast
      If Err < 1 Then ULTCLI& = !Nume_Cli
      On Error GoTo 0
    End With
    '
    If PRICLI& < 999999 Then Text10 = TRIM$(Str$(PRICLI&))
    If ULTCLI& > 0 Then Text8 = TRIM$(Str$(ULTCLI&))
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
    If XVALO(Text10.TEXT) > 32767 Then
       Text10.TEXT = Left$(TRIM$(Text10.TEXT), 4)
       Exit Sub
    End If
    If XVALO(Text10.TEXT) < 1 Then Text10.TEXT = "": Exit Sub
    NCMIN% = XVALO(Text10.TEXT)
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DblClick()
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
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
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
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub GELISPRO()

    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    If Option2.Value = True Then
      MODFILT% = 0
     ElseIf Option3.Value = True Then
      MODFILT% = 1
     ElseIf Option4.Value = True Then
      MODFILT% = 2
    End If
    '
    MODORDE% = 1
    If FORLIPRO04.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLIPRO04.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    SQLX$ = "SELECT * FROM Proved12 "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN% & " "
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX% & " "
    If MODFILT% = 0 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
    If MODFILT% = 1 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
    If MODFILT% = 2 Then SQLX$ = SQLX$
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
        Case 1
        SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
        Case 2
        SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
        Case 3
        SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set PROVTEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
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
       On Error GoTo 0
    End With
       '
    With PROVTEMP
        .Move first
        Do While (.EOF = False)
            NP% = !Nume_Cli
            NC% = (-1) * NP%
             '
            Y$ = REGBLAN$("LIPROVE")
            Call REPLA(Y$, MKI$(NP%), 1, 2)
            Call REPLA(Y$, RASOCLI$(NC%), 3, 32)
            Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
            Call REPLA(Y$, CUITCLI$(NC%), 51, 16)
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
    

90 End Sub
'

'RUTINA ANTIGUA

'
'    '
'    Dim TICUE%(32767)
'    Dim DETICUE$(2)
'    Screen.MousePointer = 11
'    '
''    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
''    DETICUE$(0) = "Cuenta Corriente"
''    DETICUE$(1) = "Gastos"
''    FIN& = ULTREG&("ACREDOR")
'    '
''    For U& = 1 To FIN&
''        Call TRACE(20, U&, FIN&)
''        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
''        If NCX% > 0 Then
''            TICUE%(NCX%) = 0
''        End If
''    Next U&
'    '
'    NCMIN% = XVALO(Text10.TEXT)
'    NCMAX% = XVALO(Text8.TEXT)
'    '
''    GRUCO% = XVALO(Text1)
''    TICUENTA% = XVALO(Text3)
'    '
'    If Option2.Value = True Then
'      MODFILT% = 0
'     ElseIf Option3.Value = True Then
'      MODFILT% = 1
'     ElseIf Option4.Value = True Then
'      MODFILT% = 2
'    End If
'    '
'    MODORDE% = 1
'    If FORLIPRO04.Option1(2) = True Then
'        MODORDE% = 2
'      ElseIf FORLIPRO04.Option1(3) = True Then
'        MODORDE% = 3
'    End If
'    'CONSULTA SQL
'    SQLX$ = "SELECT * FROM Proved12 "
'    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN% & " "
'    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX% & " "
'    If MODFILT% = 0 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
'    If MODFILT% = 1 Then SQLX$ = SQLX$ + "AND Tcpr_Cli= " & MODFILT% & ""
'    If MODFILT% = 2 Then SQLX$ = SQLX$
'    'Agrego Clausula order by a la consulta SQL
'    Select Case MODORDE%
'        Case 1
'        SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
'        Case 2
'        SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
'        Case 3
'        SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
'    End Select
'    Set PROVTEMP = Badaprov.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With PROVTEMP
'       On Error Resume Next
'       .MoveFirst
'       'Valida que la tabla no este vacia
'       If Err Then
'         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'         Screen.MousePointer = 1
'         On Error GoTo 0
'         GoTo 90
'       End If
'       On Error GoTo 0
'    End With
'       '
''       FIN& = ULTREG&("PROVED1")
''       UJ& = 0
''       FORLIPROVE.List1.Clear
''    '
''    For U& = 1 To FIN&
'        '
''        Call TRACE(20, U&, FIN&)
''        X$ = REGLEIDO$("PROVED1", U&)
''        NC% = CVI(Left$(X$, 2))
''        '
''        If NC% >= NCMIN% Then
''          If NC% <= NCMAX% Then
''             If TICUENTA% = 2 Or TICUE%(NC%) = TICUENTA% Then
''                '
''                TLISTA$ = Space$(36)
''                If MODORDE% = 1 Then
''                     Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
''                  ElseIf MODORDE% = 2 Then
''                     Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
''                  ElseIf MODORDE% = 3 Then
''                     Call REPLA(TLISTA$, "A" + AJUSTD$(CPOSCLI$((-1) * NC%), 8), 1, 9)
''                End If
''                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
''                FORLIPROVE.List1.AddItem TLISTA$
''                '
''             End If
''          End If
''        End If
''        '
''    Next U&
''    '
''    ULI& = 0
''    FIN& = FORLIPROVE.List1.ListCount
''    '
'
' 'Falta definir el origen de el tipo de cta y
' 'El grupo de proveedores
'
'    With PROVTEMP
'        .Move First
'        Do While (.EOF = False)
'
'       'For U& = 1 To FIN&
'       'Call TRACE(20, U&, FIN&)
'       'FORLIPROVE.List1.ListIndex = U& - 1
'       ' REG& = XVALO(Mid$(FORLIPROVE.List1.TEXT, 31, 8))
'       ' X$ = REGLEIDO$("PROVED1", REG&)
'       ' NP% = CVI(Left$(X$, 2))
'
'            NP% = !Nume_Cli
'            NC% = (-1) * NP%
'             '
'            Y$ = REGBLAN$("LIPROVE")
'            Call REPLA(Y$, MKI$(NP%), 1, 2)
'            Call REPLA(Y$, RASOCLI$(NC%), 3, 32)
'            Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
'            Call REPLA(Y$, CUITCLI$(NC%), 51, 16)
'            Call REPLA(Y$, POSIVAC$(NC%), 67, 10)
'            CUECOX% = TICUE%(NP%)
'            Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
'            Call REPLA(Y$, DETICUE$(CUECOX%), 79, 32)
'            Call REPLA(Y$, CUEMADRE$(NC%), 111, 16)
'            ULI& = ULI& + 1
'            Call GRAREG("LIPROVE", Y$, ULI&)
'            '
'            If TICUE%(NP%) = 0 Then
'               Y$ = REGBLAN$("LIPROVE")
'               Call REPLA(Y$, MKI$(NP%), 1, 2)
'               Call REPLA(Y$, DOMICLI$(NC%), 3, 32)
'               'Call REPLA(Y$, FAXCLI$(NC%), 35, 16)
'               Call REPLA(Y$, NIBRCLI$(NC%), 51, 16)
'               CTRX$ = TRIM$(CATIBRU$(NC%))
'               CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
'                                    If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
'                                    If CTRX$ = "2" Then CATIBX$ = "C.Multil."
'               Call REPLA(Y$, CATIBX$, 67, 10)
'               CPAG% = CPAGCLI%(NC%)
'               Call REPLA(Y$, MKI$(CPAG%), 77, 2)
'               Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
'               ULI& = ULI& + 1
'               Call GRAREG("LIPROVE", Y$, ULI&)
'               '
'               Y$ = REGBLAN$("LIPROVE")
'               Call REPLA(Y$, MKI$(NP%), 1, 2)
'               DOMI$ = TRIM$(CPOSCLI$(NC%))
'               If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
'               DOMI$ = DOMI$ + LOCACLI$(NC%)
'               Call REPLA(Y$, DOMI$, 3, 32)
'               'Call REPLA(Y$, EMAICLI$(NC%), 35, 16)
'               ULI& = ULI& + 1
'               Call GRAREG("LIPROVE", Y$, ULI&)
'            End If
'            '
'            Y$ = REGBLAN$("LIPROVE")
'            ULI& = ULI& + 1
'            Call GRAREG("LIPROVE", Y$, ULI&)
'            '
'            .MoveNext
'        Loop
'        'Next U&
'    End With
'    '
'    Call SETULTREG("LIPROVE", ULI&)
'    Call CIERRARCH("LIPROVE")
'    Screen.MousePointer = 1
'    Call FINAL("*LIPROVE")
'    '
'
'
'90 End Sub
''

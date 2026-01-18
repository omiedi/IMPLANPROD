VERSION 5.00
Begin VB.Form FORLICLI04 
   BackColor       =   &H00C0D0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listados de Clientes - Cuentas Corrientes -"
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
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   3800
         Left            =   0
         Picture         =   "FORLICLI04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0D0F0&
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
      TabIndex        =   22
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
         Picture         =   "FORLICLI04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   26
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
         TabIndex        =   24
         Top             =   375
         Width           =   2850
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   3465
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0D0F0&
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
      TabIndex        =   12
      Top             =   105
      Width           =   4005
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0D0F0&
         Caption         =   "Excluir  de Baja"
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
         Left            =   315
         TabIndex        =   31
         Top             =   400
         Width           =   1065
      End
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   20
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
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
         Left            =   210
         TabIndex        =   19
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0F0&
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
      Begin VB.TextBox Text4 
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1680
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
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
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2940
         TabIndex        =   8
         Text            =   " "
         Top             =   1260
         Width           =   540
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   840
         Width           =   3330
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   5
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   2940
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   525
         TabIndex        =   11
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente:"
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
         Left            =   525
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0D0F0&
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
         BackColor       =   &H00C0D0F0&
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
         TabIndex        =   23
         Top             =   420
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0D0F0&
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
         BackColor       =   &H00C0D0F0&
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
      BackColor       =   &H00000080&
      Height          =   4635
      Left            =   8340
      ScaleHeight     =   4575
      ScaleWidth      =   1005
      TabIndex        =   25
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
         Picture         =   "FORLICLI04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "FORLICLI04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   29
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
         TabIndex        =   27
         Top             =   2520
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   28
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "FORLICLI04.frx":2AF6
         Stretch         =   -1  'True
         Top             =   2835
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FORLICLI04"
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
    Call GELISCLI
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
    Call SELECHO("DEUDOR1")
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      FORLICLI04.Frame1.BackColor = &HFCD7B6
      FORLICLI04.Frame2.BackColor = &HFCD7B6
      FORLICLI04.Frame4.BackColor = &HFCD7B6
      FORLICLI04.Frame5.BackColor = &HFCD7B6
      FORLICLI04.Option1(1).BackColor = &HFCD7B6
      FORLICLI04.Option1(2).BackColor = &HFCD7B6
      FORLICLI04.Option1(3).BackColor = &HFCD7B6
      Check1.BackColor = &HFCD7B6
    End If

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    Call APERBASDAT("CLIEN")
    PRICLI& = 999999: ULTCLI& = 0
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM Deudor12 "
    SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
    SQLX$ = SQLX$ + "ORDER BY Nume_Cli ASC"
    Set PROVETEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PROVETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then PRICLI& = !Nume_Cli
      .MoveLast
      If Err < 1 Then ULTCLI& = !Nume_Cli
      On Error GoTo 0
    End With
    '
    If PRICLI& < 999999 Then
       Text10 = TRIM$(Str$(PRICLI&))
       Text9 = RASOCLI$(PRICLI&)
    End If
    If ULTCLI& > 0 Then
       Text8 = TRIM$(Str$(ULTCLI&))
       Text7.TEXT = RASOCLI$(ULTCLI&)
    End If
    '
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
    Text9.TEXT = RASOCLI$(NCMIN%)
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
    Text7.TEXT = RASOCLI$(NCMAX%)
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

Sub GELISCLI()

    '
    Screen.MousePointer = 11
    '
    NCMIN% = XVALO(Text10.TEXT)
    NCMAX% = XVALO(Text8.TEXT)
    '
    GRUCO% = XVALO(Text1)
    VENDE% = XVALO(Text3)
    '
    MODORDE% = 1
    If FORLICLI04.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLICLI04.Option1(3) = True Then
        MODORDE% = 3
    End If
    'CONSULTA SQL
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM DEUDOR12 "
    SQLX$ = SQLX$ + "WHERE Nume_Cli>=" & NCMIN% & " "
    SQLX$ = SQLX$ + "AND Nume_Cli<=" & NCMAX% & " "
    If TRIM$(Text1) <> "" Then SQLX$ = SQLX$ + "AND GrCo_Cli =  " & GRUCO% & "  "
    If TRIM$(Text3) <> "" Then SQLX$ = SQLX$ + "AND VEND_Cli = " & VENDE% & " "
    If Check1.Value = 1 Then SQLX$ = SQLX$ + "AND Stat_Cli = 1 "
    'Agrego Clausula order by a la consulta SQL
    Select Case MODORDE%
      Case 1
      SQLX$ = SQLX$ + " ORDER BY Nume_Cli ASC "
      Case 2
      SQLX$ = SQLX$ + " ORDER BY Raso_Cli ASC "
      Case 3
      SQLX$ = SQLX$ + " ORDER BY Cpos_Cli ASC "
    End Select
    Set PROVTEMP = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
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
      '
      .Move first
      Do While (.EOF = False)
        '
        NC% = !Nume_Cli
        '
        Label3 = RASOCLI$(NC%)
        Label3.Refresh
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NC%), 3, 32)
        Call REPLA(Y$, TEL1CLI$(NC%), 35, 16)
        Call REPLA(Y$, CUITCLI$(NC%), 51, 16)
        Call REPLA(Y$, POSIVAC$(NC%), 67, 10)
        '\\\OT-06-0081-WAR-06/03/06///
        LPRX% = LIPRCLI%(NC%)
        Call REPLA(Y$, MKI$(LPRX%), 77, 2)
        DDLX$ = TRIM$(DELIPRE$(LPRX%))
        DCCL$ = TRIM$(DESCUCLI$(NC%))
        If DCCL$ <> "" Then
           If InStr(DCCL$, "%") < 1 Then
              DCCL$ = DCCL$ + " %"
           End If
           If Len(DCCL$) > 28 Then DCCL$ = Left$(DCCL$, 28)
           DDLX$ = Left$(DDLX$, 29 - Len(DCCL$))
           DDLX$ = DDLX$ + " (" + DCCL$ + ")"
        End If
        Call REPLA(Y$, DDLX$, 79, 32)
        '\\\OT-06-0081-WAR-FIN///
        Call REPLA(Y$, CUEMADRE$(NC%), 111, 16)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, DOMICLI$(NC%), 3, 32)
        Call REPLA(Y$, FAXCLI$(NC%), 35, 16)
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
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        DOMI$ = TRIM$(CPOSCLI$(NC%))
        If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
        DOMI$ = DOMI$ + LOCACLI$(NC%)
        Call REPLA(Y$, DOMI$, 3, 32)
        Call REPLA(Y$, EMAILCLI$(NC%), 35, 16)
          CUECOX% = VENDCLI%(NC%)
        Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
        Call REPLA(Y$, DEVENDE$(CUECOX%), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
      .MoveNext
      Loop
    End With
    '
    Call HEADERS("LICLIEN", "")
    '
    If Check1.Value = 1 Then
      Call HEADERS("LICLIEN", "Excluidos los Clientes Dados de Baja ")
     Else
      Call HEADERS("LICLIEN", "Todos ")
    End If
    '
    Call SETULTREG("LICLIEN", ULI&)
    Call CIERRARCH("LICLIEN")
    Screen.MousePointer = 1
    Call FINAL("*LICLIEN")
    '
90 End Sub
'

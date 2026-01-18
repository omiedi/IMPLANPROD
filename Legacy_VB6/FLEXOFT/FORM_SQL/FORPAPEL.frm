VERSION 5.00
Begin VB.Form FORPAPEL 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de Código de Artículo por Rubro"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   5535
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros Insumos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   30
      Top             =   2400
      Width           =   4500
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Varios (V)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ribbons (R)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   660
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nucleos (N)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cajas (C)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Foils (F)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   105
      TabIndex        =   17
      Top             =   4770
      Width           =   4500
      Begin VB.CommandButton Command2 
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
         Height          =   280
         Left            =   2040
         Picture         =   "FORPAPEL.frx":0000
         TabIndex        =   36
         Top             =   1470
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   735
         Width           =   4125
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1680
         TabIndex        =   26
         Top             =   1470
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Material:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   25
         Top             =   1530
         Width           =   1485
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   22
         Top             =   1470
         Width           =   1995
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2625
         TabIndex        =   21
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   20
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   495
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Terminados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   12
      Top             =   3615
      Width           =   4500
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
         Height          =   330
         Left            =   4095
         Picture         =   "FORPAPEL.frx":030A
         TabIndex        =   13
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   16
         Top             =   315
         Width           =   2430
      End
      Begin VB.Label Label13 
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
         Height          =   330
         Left            =   2940
         TabIndex        =   15
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   14
         Top             =   630
         Width           =   4110
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Papel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4500
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
         Height          =   330
         Left            =   4095
         Picture         =   "FORPAPEL.frx":0614
         TabIndex        =   29
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro - Subrubro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   6
         Top             =   315
         Width           =   2430
      End
      Begin VB.Label Label11 
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
         Height          =   330
         Left            =   2940
         TabIndex        =   2
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   1
         Top             =   630
         Width           =   4110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   6870
      Left            =   4725
      ScaleHeight     =   6810
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command3 
         Height          =   640
         Left            =   105
         Picture         =   "FORPAPEL.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   640
      End
      Begin VB.CommandButton BOTOK 
         Height          =   1170
         Left            =   105
         Picture         =   "FORPAPEL.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Graba Nuevo Código"
         Top             =   4875
         Width           =   640
      End
      Begin VB.CommandButton Command5 
         Height          =   630
         Left            =   105
         Picture         =   "FORPAPEL.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Tabla de tipo Papel"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Height          =   640
         Left            =   105
         Picture         =   "FORPAPEL.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Tabla de Tipo de Tinta"
         Top             =   1470
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4455
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FORPAPEL.frx":1006
         Top             =   6240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tinta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   7
      Top             =   1260
      Width           =   4500
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   4095
         Picture         =   "FORPAPEL.frx":2F28
         TabIndex        =   11
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   4110
      End
      Begin VB.Label Label12 
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
         Height          =   330
         Left            =   2940
         TabIndex        =   9
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro - Subrubro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   8
         Top             =   315
         Width           =   2430
      End
   End
End
Attribute VB_Name = "FORPAPEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OTRINSU%

Private Sub BOTOK_Click()
    '
    BOTOK.Enabled = False
    'DDDDD
    '
    CODINUE$ = Label4
    If OTRINSU% < 1 Then
       X$ = FILTERGETRECORD$("TIPOPAPE", 1, TRIM$(Label4))
       If Len(X$) < 4 Then
         Call COMUNI("Código de Familia No Válido")
         GoTo 99
       End If
    End If
    '
    If TRIM$(Text1) = "" Then
      Call COMUNI("Falta Descripción del Item")
      GoTo 99
    End If
    '
    TIMA% = XVALO(Label7)
    If TIMA% < 1 Then
       Call MENSERR(24, "Falta Tipo de Material")
       BOTOK.Enabled = True
       Exit Sub
    End If
    
    DESCRINUE$ = Text1
    UNID$ = "U."
    If TIMA% = 1 Then UNID$ = "M2"
    If TIMA% = 2 Then UNID$ = "Kg"
    ETRAZA% = 1
    
    Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%, ETRAZA%)
    '/////////***********/////////////////
    If OKX% < 1 Then
      Call MENSERR(24, "El Alta de Codigo\no pudo Completarse.\  \Ejecute Control de Maestro\y Consulte a su Soporte Técnico.")
    Else
      AMASTOK07.Text1 = CODINUE$
    End If
    '
99  Screen.MousePointer = 1
    BOTOK.Enabled = True
    Call LIMPIAR: Call LIMPIAR2
    '
End Sub



Private Sub Command1_Click()
   Call SELECHO("TIPOTINT")
   VDEV$ = TRIM$(VALACT1$("TIPOTINT"))
   If VDEV$ <> "" Then
     X$ = FILTERGETRECORD$("TIPOTINT", 1, TRIM$(VDEV$))
     Call LIMPIAR: Call LIMPIAR2
     Label12 = Mid$(X$, 1, 4)
     Label9 = VALACT2("TIPOTINT")
     Text1 = VALACT2("TIPOTINT")
     Label7 = 2
     Text1.SetFocus
     Text1.SelStart = Len(Text1)
     Label4 = "T" + Label12 + "-" ' se antecede la T en tintas
     Call ArmaCodigo
   End If
End Sub


Private Sub Command2_Click()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Label7 = VALONUE$
   End If
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command4_Click()
   Call SELECHO("DEUDOR1")
   VDEV$ = TRIM$(VALACT1$("DEUDOR1"))
   If VDEV$ <> "" Then
     Call LIMPIAR: Call LIMPIAR2
     Label13 = VALACT1("DEUDOR1")
     Label1 = VALACT2("DEUDOR1")
     Text1 = VALACT2("DEUDOR1")
     Label7 = 5
     Text1.SetFocus
     Text1.SelStart = Len(Text1)
     ABUS$ = TRIM$(Label13)
     While Len(ABUS$) < 4: ABUS$ = "0" + ABUS$: Wend
     Label4 = ABUS$ + "-"
     Call ArmaCodigo
   End If
End Sub


Private Sub Command5_Click()
    '
    Call BLANARCH("!TIPOPAPE")
    For i& = 1 To ULTREG&("TIPOPAPE")
        X$ = REGLEIDO$("TIPOPAPE", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!TIPOPAPE", X$)
        End If
    Next i&
    Call CIERRARCH("TIPOPAPE")
    '
    VTB% = VENTABU%("TIPOPAPE/NC/TITUPAN=Tipo de Papeles")
    If VTB% < 0 Then Exit Sub
    j& = 0
    For i& = 1 To ULTREG&("!TIPOPAPE")
        X$ = REGLEIDO$("!TIPOPAPE", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            j& = j& + 1
            Call GRAREG("TIPOPAPE", X$, j&)
        End If
    Next i&
    Call SETULTREG("TIPOPAPE", j&)
    Call CIERRARCH("TIPOPAPE")
    '


End Sub

Private Sub Command6_Click()
    Call BLANARCH("!TIPOTINT")
    For i& = 1 To ULTREG&("TIPOTINT")
        X$ = REGLEIDO$("TIPOTINT", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!TIPOTINT", X$)
        End If
    Next i&
    Call CIERRARCH("TIPOTINT")
    '
    VTB% = VENTABU%("TIPOTINT/NC/TITUPAN=Tipo de Tintas")
    If VTB% < 0 Then Exit Sub
    j& = 0
    For i& = 1 To ULTREG&("!TIPOTINT")
        X$ = REGLEIDO$("!TIPOTINT", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            j& = j& + 1
            Call GRAREG("TIPOTINT", X$, j&)
        End If
    Next i&
    Call SETULTREG("TIPOTINT", j&)
    Call CIERRARCH("TIPOTINT")

End Sub

Private Sub Command9_Click()
   Call SELECHO("TIPOPAPE")
   VDEV$ = TRIM$(VALACT1$("TIPOPAPE"))
   If VDEV$ <> "" Then
      X$ = FILTERGETRECORD$("TIPOPAPE", 1, TRIM$(VDEV$))
      Call LIMPIAR: Call LIMPIAR2
      Label11 = Mid$(X$, 1, 4)
      Label8 = VALACT2("TIPOPAPE")
      Text1 = VALACT2("TIPOPAPE")
      Label7 = 1
      Label4 = Label11 + "-"
      Text1.SetFocus
      Text1.SelStart = Len(Text1)
      Call ArmaCodigo
   End If
End Sub
'
Private Sub LIMPIAR()
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label1.Caption = ""
    Label4.Caption = ""
    Label7.Caption = ""
    Label29.Caption = ""
    Text1.TEXT = ""
    OTRINSU% = 1: Command2.Visible = True
End Sub

Sub LIMPIAR2()
    Option1.Value = 0
    Option2.Value = 0
    Option3.Value = 0
    Option4.Value = 0
    Option5.Value = 0
    OTRINSU% = 0: Command2.Visible = False
End Sub
'
Sub ArmaCodigo()
    PRONULI& = 1
    ABUS$ = UCase$(TRIM$(Label4))
    '
    Call ABREMASTER
    If OTRINSU% < 1 Then
         SQLX$ = "SELECT * FROM MASTER WHERE TIMATE = " & TRIM$(Label7)
       Else
         SQLX$ = "SELECT * FROM MASTER WHERE CODIGO LIKE '" & ABUS$ & "'"
    End If
    '
    Dim ARTLITMP As ADODB.Recordset
    Set ARTLITMP = New ADODB.Recordset
    ARTLITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With ARTLITMP
        On Error Resume Next
        .MoveFirst
        If Err Then
           On Error GoTo 0
           GoTo 20
        End If
        On Error GoTo 0
        '
        Do Until .EOF = True
          On Error Resume Next
          CODILEIDO$ = UCase$(TRIM$(!Codigo)) 'Mid$(XX$, 77, 15)))
          If Err Then
            On Error GoTo 0
            GoTo 19
          End If
          On Error GoTo 0
          If XVALO(Label7) <> 5 And Left$(CODILEIDO$, Len(ABUS$)) = ABUS$ Then
             NUITN& = XVALO(Mid$(CODILEIDO$, 1 + Len(ABUS$)))
             If NUITN& >= PRONULI& Then PRONULI& = 1 + NUITN&
          End If
19        .MoveNext
        Loop
    End With
20  ARTLITMP.Close
    Set ARTLITMP = Nothing
    '
25  PRONUX$ = TRIM$(Str$(PRONULI&))
    While Len(PRONUX$) < 4
      PRONUX$ = "0" + PRONUX$
    Wend
    CODINUE$ = ABUS$ + PRONUX$
    If CODINT%(CODINUE$) <> 0 Then ' YA EXISTE EL CÓDIGO
        PRONULI& = PRONULI& + 1
        GoTo 25
    End If
    '
    Label4 = CODINUE$
End Sub

Private Sub Form_Load()
  Call LIMPIAR: Call LIMPIAR2
  Call APLICACIONSKINS(Me, Picture1)
End Sub


Private Sub Label11_DblClick()
  Command9_Click
End Sub

Private Sub Label12_DblClick()
  Command1_Click
End Sub

Private Sub Label13_DblClick()
   Call Command4_Click
End Sub

Private Sub Label7_Change()
   VALONU = XVALO(Label7)
   If VALONU < 0 Or VALONU > 255 Then
      Label7 = "": Exit Sub
   End If
   Label29.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
End Sub

Private Sub Option1_Click()
   Call LIMPIAR
   Text1 = Option1.Caption
   '      Label7 = 2
   Text1.SetFocus
   Text1.SelStart = Len(Text1)
   Label4 = "F0000" + "-"   ' para foils
   Call ArmaCodigo
End Sub

Private Sub Option2_Click()
   Call LIMPIAR
   Text1 = Option2.Caption
   '      Label7 = 2
   Text1.SetFocus
   Text1.SelStart = Len(Text1)
   Label4 = "C0000" + "-"   ' para CAJAS
   Call ArmaCodigo
End Sub

Private Sub Option3_Click()
   Call LIMPIAR
   Text1 = Option3.Caption
   '      Label7 = 2
   Text1.SetFocus
   Text1.SelStart = Len(Text1)
   Label4 = "N0000" + "-"   ' para NUCLEOS
   Call ArmaCodigo
End Sub

Private Sub Option4_Click()
   Call LIMPIAR
   Text1 = Option4.Caption
   '      Label7 = 2
   Text1.SetFocus
   Text1.SelStart = Len(Text1)
   Label4 = "R0000" + "-"   ' para RIBBONS
   Call ArmaCodigo
End Sub

Private Sub Option5_Click()
   Call LIMPIAR
   Text1 = Option5.Caption
   '      Label7 = 2
   Text1.SetFocus
   Text1.SelStart = Len(Text1)
   Label4 = "V0000" + "-"   ' para VARIOS
   Call ArmaCodigo
End Sub

Private Sub Text1_Change()
Text1.MaxLength = 64
End Sub

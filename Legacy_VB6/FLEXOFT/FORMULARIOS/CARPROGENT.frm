VERSION 5.00
Begin VB.Form CARPROGENT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización Programa de Entregas"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7365
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
   ScaleHeight     =   6315
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   790
      Left            =   4515
      TabIndex        =   23
      Top             =   5295
      Width           =   2640
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   915
         Left            =   -105
         Picture         =   "CARPROGENT.frx":0000
         Stretch         =   -1  'True
         Top             =   -50
         Width           =   10815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grabar"
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
      Left            =   5460
      Picture         =   "CARPROGENT.frx":13166
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4410
      Width           =   750
   End
   Begin VB.CommandButton BOTPRINT 
      Caption         =   "Print"
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
      Left            =   6405
      Picture         =   "CARPROGENT.frx":13470
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4410
      Width           =   750
   End
   Begin VB.CommandButton BOTERASE 
      Caption         =   "Borrar"
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
      Left            =   4515
      Picture         =   "CARPROGENT.frx":1377A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4410
      Width           =   750
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1695
      Left            =   210
      TabIndex        =   22
      Top             =   4410
      Width           =   4005
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   0
         TabIndex        =   5
         Top             =   210
         Width           =   4005
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         Caption         =   " Fecha                 Cant.             Acum."
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   4005
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ingresar Programa"
      Height          =   960
      Left            =   210
      TabIndex        =   19
      Top             =   3255
      Width           =   6945
      Begin VB.CommandButton Command7 
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
         Left            =   2310
         TabIndex        =   27
         Top             =   420
         Width           =   175
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
         Height          =   465
         Left            =   5880
         Picture         =   "CARPROGENT.frx":13A84
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   360
         Left            =   3990
         TabIndex        =   3
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   1155
         TabIndex        =   2
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2940
         TabIndex        =   21
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   420
         TabIndex        =   20
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección"
      Height          =   1800
      Left            =   210
      TabIndex        =   14
      Top             =   1260
      Width           =   6945
      Begin VB.CommandButton Command6 
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
         Left            =   2730
         TabIndex        =   26
         Top             =   1050
         Width           =   175
      End
      Begin VB.CommandButton Command5 
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
         Left            =   1890
         TabIndex        =   25
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   3150
         TabIndex        =   18
         Text            =   " "
         Top             =   1050
         Width           =   3585
      End
      Begin VB.TextBox Text3 
         Height          =   390
         Left            =   1155
         TabIndex        =   1
         Top             =   1050
         Width           =   1590
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   2310
         TabIndex        =   17
         Top             =   525
         Width           =   4425
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   1155
         TabIndex        =   0
         Text            =   " "
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo:"
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
         Left            =   -315
         TabIndex        =   16
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   -315
         TabIndex        =   15
         Top             =   630
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opciones de Tabajo"
      Height          =   855
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   5055
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Programa de Entregas por Artículo y Cliente"
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
         Left            =   2730
         TabIndex        =   13
         Top             =   315
         Width           =   2220
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Estimado de   Ventas por Material"
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
         Left            =   525
         TabIndex        =   12
         Top             =   315
         Width           =   2115
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
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
      Left            =   5355
      Picture         =   "CARPROGENT.frx":13D8E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   210
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
      Left            =   6300
      Picture         =   "CARPROGENT.frx":14098
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "CARPROGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PGE$(256)
Dim CI As Integer
Dim NC As Integer
'

Private Sub BOTERASE_Click()
    PPX% = List1.ListIndex
    If PPX% >= 0 Then
        If PPX% <= List1.ListCount - 1 Then
            List1.RemoveItem (PPX%)
        End If
    End If
    '
    CAMO% = List1.ListCount
    For II% = 1 To CAMO%
       PGE$(II%) = Left$(List1.List(II% - 1), 19)
    Next II%
    '
    List1.Clear
    ACUM = 0
    For II% = 1 To CAMO%
       ACUM = ACUM + Val(Mid$(PGE$(II%), 9, 11))
       CQ$ = CSTRING$(MKS$(ACUM), 3, 11, 1, 2)
       List1.AddItem PGE$(II%) + CQ$
    Next II%
    '
    If PPX% >= List1.ListCount Then
        PPX% = List1.ListCount - 1
    End If
    '
    List1.ListIndex = PPX%
    '
End Sub

Private Sub BOTPRINT_Click()
    PrintForm
End Sub

Private Sub Command1_Click()
    CARPROGENT.Hide
    Call FINAL("*SPROGEN")
End Sub

Private Sub Command2_Click()
    OPCI% = 1
    If CARPROGENT.Option2.Value = True Then OPCI% = 2
    '
    If OPCI% = 1 Then
        Call OPNOIN("")
        CARPROGENT.Option2.Value = True
      ElseIf OPCI% = 2 Then
        Text1.SetFocus
    End If
    '
End Sub

Private Sub Command3_Click()
    '

    FF$ = TRIM$(Text5.Text)
    FEX = FECHANUM(FF$)
    If FEX < 1 Then
        Text5.SetFocus
        Exit Sub
    End If
    FF$ = FECHATEX$(FEX)
    CANTI = Val(Text6.Text)
    If CANTI < 0.1 Then
        Text6.SetFocus
        Exit Sub
    End If
    '
    CAMO% = List1.ListCount + 1
    For II% = 1 To CAMO% - 1
       PGE$(II%) = Left$(List1.List(II% - 1), 19)
    Next II%
    '
    CP$ = CSTRING$(MKS$(CANTI), 3, 11, 1, 2)
    PGE$(CAMO%) = FF$ + CP$
    PGE0$ = FF$ + CP$
    '
    For II% = 1 To CAMO%
        For JJ% = II% + 1 To CAMO%
            If FECHANUM(Left$(PGE$(II%), 8)) > FECHANUM(Left$(PGE$(JJ%), 8)) Then
                PGX$ = PGE$(II%)
                PGE$(II%) = PGE$(JJ%)
                PGE$(JJ%) = PGX$
            End If
        Next JJ%
    Next II%
    '
    List1.Clear
    ACUM = 0
    For II% = 1 To CAMO%
       ACUM = ACUM + Val(Mid$(PGE$(II%), 9, 11))
       CQ$ = CSTRING$(MKS$(ACUM), 3, 11, 1, 2)
       List1.AddItem PGE$(II%) + CQ$
       If PGE$(II%) = PGE0$ Then
           List1.ListIndex = II% - 1
       End If
    Next II%
    '
    Text5.SetFocus
    '
End Sub

Private Sub Command4_Click()
   '
   Screen.MousePointer = 11
   J& = 0
   For U& = 1 To ULTREG&("PROGENT")
      XXX$ = REGLEIDO$("PROGENT", U&)
      If CVI(Left$(XXX$, 2)) <> CI% Or CVI(Mid$(XXX$, 5, 2)) <> NC% Then
          J& = J& + 1
          Call GRAREG("PROGENT", XXX$, J&)
      End If
   Next U&
   '
   For II% = 0 To List1.ListCount - 1
      '
      FEXI% = FECHANUM(Left$(List1.List(II%), 8))
      CANTI = Val(Mid$(List1.List(II%), 9, 11))
      XXX$ = REGBLAN$("PROGENT")
      Call REPLA(XXX$, MKI$(CI%), 1, 2)
      Call REPLA(XXX$, MKI$(FEXI%), 3, 2)
      Call REPLA(XXX$, MKI$(NC%), 5, 2)
      Call REPLA(XXX$, MKS$(CANTI), 7, 4)
      '
      J& = J& + 1
      If J& <= ULTREG&("PROGENT") Then
          Call GRAREG("PROGENT", XXX$, J&)
        Else
          Call REGAPP("PROGENT", XXX$)
      End If
      '
   Next II%
   '
   For U& = J& + 1 To ULTREG&("PROGENT")
      Call GRAREG("PROGENT", String$(128, 0), U&)
   Next U&
   '
   Call CIERRARCH("PROGENT")
   List1.Clear
   Text5.Text = ""
   Text6.Text = ""
   Screen.MousePointer = 1
   Text1.SetFocus
   '
End Sub

Private Sub Command5_Click()
    If Option2.Value = True Then
        Call SELECHO("DEUDOR1")
        NCC$ = VALACT1$("DEUDOR1")
        If TRIM$(NCC$) <> "" Then
            Text1.Text = NCC$
            Text2.Text = VALACT2$("DEUDOR1")
            Text1.Refresh
            Text2.Refresh
        End If
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODINT%(CODD$) > 0 Then
        Text3.Text = VALACT1$("MASTER")
        Text3.Refresh
        Text4.Text = VALACT2$("MASTER")
        Text4.Refresh
    End If
End Sub

Private Sub Command7_Click()
    Call SELECHO("SELFECHA")
    FFXX$ = VALACT1$("SELFECHA")
    If TRIM$(FFXX$) <> "" Then
        Text5.Text = FFXX$
        Text6.SetFocus
    End If
End Sub

Private Sub Form_Load()
   Option2.Value = True
End Sub

Private Sub List1_DblClick()
    Text5.Text = Left$(List1.Text, 8)
    Text6.Text = TRIM$(Mid$(List1.Text, 9, 11))
End Sub

Private Sub Option1_Click()
    Option2.Value = False
    Option1.Value = True
    Text1.Text = ""
    Text2.Text = "Estimado Clientes Varios"
    Text3.SetFocus
End Sub

Private Sub Text1_DblClick()
    If Option2.Value = True Then
        Call SELECHO("DEUDOR1")
        NCC$ = VALACT1$("DEUDOR1")
        If TRIM$(NCC$) <> "" Then
            Text1.Text = NCC$
            Text2.Text = VALACT2$("DEUDOR1")
            Text1.Refresh
            Text2.Refresh
        End If
      Else
        Text3.SetFocus
    End If
End Sub

Private Sub Text1_GotFocus()
    If Option1.Value = True Then
        Text3.SetFocus
    End If
End Sub

Private Sub Text1_LostFocus()
    List1.Clear
    ACUM = 0
    For U& = 1 To ULTREG&("PROGENT")
       XXX$ = REGLEIDO$("PROGENT", U&)
       If CVI(Mid$(XXX$, 1, 2)) = CI% Then
          If CVI(Mid$(XXX$, 5, 2)) = NC% Then
             FF$ = FECHATEX$(CVI(Mid$(XXX$, 3, 2)))
             XXZ$ = Mid$(XXX$, 7, 4)
             CP$ = CSTRING$(XXZ$, 3, 11, 1, 2)
             ACUM = ACUM + Val(CP$)
             CQ$ = CSTRING$(MKS$(ACUM), 3, 11, 1, 2)
             List1.AddItem FF$ + CP$ + CQ$
          End If
        End If
    Next U&
End Sub

Private Sub Text2_GotFocus()
    Text3.SetFocus
End Sub

Private Sub Text3_Change()
    CIIX% = CODINT%(Text3.Text)
    If CIIX% > 0 Then
        Text4.Text = DESCRIT$(CIIX%)
        Text5.SetFocus
        List1.Clear
    End If
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODINT%(CODD$) > 0 Then
        Text3.Text = VALACT1$("MASTER")
        Text3.Refresh
        Text4.Text = VALACT2$("MASTER")
        Text4.Refresh
    End If
End Sub

Private Sub Text3_GotFocus()
    CARPROGENT.Refresh
End Sub

Private Sub Text3_LostFocus()
    List1.Clear
    ACUM = 0
    For U& = 1 To ULTREG&("PROGENT")
       XXX$ = REGLEIDO$("PROGENT", U&)
       If CVI(Mid$(XXX$, 1, 2)) = CI% Then
          If CVI(Mid$(XXX$, 5, 2)) = NC% Then
             FF$ = FECHATEX$(CVI(Mid$(XXX$, 3, 2)))
             XXZ$ = Mid$(XXX$, 7, 4)
             CP$ = CSTRING$(XXZ$, 3, 11, 1, 2)
             ACUM = ACUM + Val(CP$)
             CQ$ = CSTRING$(MKS$(ACUM), 3, 11, 1, 2)
             List1.AddItem FF$ + CP$ + CQ$
          End If
        End If
    Next U&
End Sub

Private Sub Text4_GOTFOCUS()
    Text3.SetFocus
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    FFXX$ = VALACT1$("SELFECHA")
    If TRIM$(FFXX$) <> "" Then
        Text5.Text = FFXX$
        Text6.SetFocus
    End If
End Sub

Private Sub Text5_GotFocus()
    '
    If List1.ListCount > 0 Then
        Exit Sub
    End If
    '
    COD$ = AJUSTI$(Text3.Text, 16)
    CI% = CODINT%(COD$)
    If CI% < 1 Then
        Text3.SetFocus
        Exit Sub
    End If
    NUCLI = Val(Text1.Text)
    NUCLIMI = 0
    If Option2.Value = True Then NUCLIMI = 1
    If NUCLI < NUCLIMI Or NUCLI > 32767 Then
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
    End If
    NC% = NUCLI
    If Option2.Value = True Then
        If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
            Text1.SetFocus
            Exit Sub
        End If
    End If
    '
    List1.Clear
    ACUM = 0
    For U& = 1 To ULTREG&("PROGENT")
       XXX$ = REGLEIDO$("PROGENT", U&)
       If CVI(Mid$(XXX$, 1, 2)) = CI% Then
          If CVI(Mid$(XXX$, 5, 2)) = NC% Then
             FF$ = FECHATEX$(CVI(Mid$(XXX$, 3, 2)))
             XXZ$ = Mid$(XXX$, 7, 4)
             CP$ = CSTRING$(XXZ$, 3, 11, 1, 2)
             ACUM = ACUM + Val(CP$)
             CQ$ = CSTRING$(MKS$(ACUM), 3, 11, 1, 2)
             List1.AddItem FF$ + CP$ + CQ$
          End If
        End If
    Next U&
    '
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
        Text6.SetFocus
    End If
End Sub

Private Sub Text6_Click()
    COD$ = AJUSTI$(Text3.Text, 16)
    CI% = CODINT%(COD$)
    If CI% < 1 Then
        Text3.SetFocus
        Exit Sub
    End If
    NUCLI = Val(Text1.Text)
    NUCLIMI = 0
    If Option2.Value = True Then NUCLIMI = 1
    If NUCLI < NUCLIMI Or NUCLI > 32767 Then
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
    End If
    NC% = NUCLI
    If Option2.Value = True Then
        If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
            Text1.SetFocus
            Exit Sub
        End If
    End If
    '
    If FECHANUM(Text5.Text) < 1 Then
        Text5.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
        Command3.SetFocus
    End If
End Sub


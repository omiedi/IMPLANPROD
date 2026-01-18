VERSION 5.00
Begin VB.Form FORSELECHO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Courier New"
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
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox LINOSORT 
      Height          =   540
      ItemData        =   "FORSELECHO.frx":0000
      Left            =   210
      List            =   "FORSELECHO.frx":0007
      TabIndex        =   56
      Top             =   1995
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.VScrollBar SCROLLBA 
      Height          =   1065
      Left            =   4410
      TabIndex        =   55
      Top             =   0
      Width           =   225
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   15
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   15
         Left            =   315
         TabIndex        =   54
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   15
         Left            =   630
         TabIndex        =   53
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   14
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   14
         Left            =   315
         TabIndex        =   51
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   14
         Left            =   630
         TabIndex        =   50
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   13
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   13
         Left            =   315
         TabIndex        =   48
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   13
         Left            =   630
         TabIndex        =   47
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   12
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   12
         Left            =   315
         TabIndex        =   45
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   12
         Left            =   630
         TabIndex        =   44
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   11
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   11
         Left            =   315
         TabIndex        =   42
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   11
         Left            =   630
         TabIndex        =   41
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   10
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   10
         Left            =   315
         TabIndex        =   39
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   10
         Left            =   630
         TabIndex        =   38
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   9
      Left            =   105
      TabIndex        =   34
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   9
         Left            =   315
         TabIndex        =   36
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   9
         Left            =   630
         TabIndex        =   35
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   8
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   8
         Left            =   315
         TabIndex        =   33
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   8
         Left            =   630
         TabIndex        =   32
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   7
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   7
         Left            =   315
         TabIndex        =   30
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   7
         Left            =   630
         TabIndex        =   29
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   6
         Left            =   315
         TabIndex        =   27
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   6
         Left            =   630
         TabIndex        =   26
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   5
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   5
         Left            =   315
         TabIndex        =   24
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   5
         Left            =   630
         TabIndex        =   23
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   4
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   4
         Left            =   315
         TabIndex        =   21
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   4
         Left            =   630
         TabIndex        =   20
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   3
         Left            =   315
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   3
         Left            =   630
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   2
         Left            =   315
         TabIndex        =   15
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   2
         Left            =   630
         TabIndex        =   14
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   1
         Left            =   315
         TabIndex        =   12
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   1
         Left            =   630
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4635
      Begin VB.ListBox TABLA 
         Height          =   1020
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         Height          =   1020
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Top             =   210
         Visible         =   0   'False
         Width           =   3690
      End
   End
   Begin VB.ListBox LISORT 
      Height          =   540
      ItemData        =   "FORSELECHO.frx":0015
      Left            =   2415
      List            =   "FORSELECHO.frx":001C
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1995
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton BOTCANCE 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4305
      TabIndex        =   5
      Top             =   2835
      Width           =   330
   End
   Begin VB.CommandButton BOTERA 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3885
      TabIndex        =   4
      Top             =   2835
      Width           =   500
   End
   Begin VB.CommandButton BOTINSE 
      Caption         =   "Ins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3465
      TabIndex        =   3
      Top             =   2835
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   945
      TabIndex        =   0
      Top             =   2800
      Width           =   2115
   End
   Begin VB.Line LINVER 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   2100
      Y2              =   2730
   End
   Begin VB.Line LINHOR 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   4725
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   2895
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   240
      Left            =   105
      TabIndex        =   1
      Top             =   1785
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "FORSELECHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTCANCE_Click()
    '
    If TASEL(ITABLA%).Visible = True Then
       Text1.TEXT = ""
       TASEL(ITABLA%).Visible = False
       If TASEL(ITABLA%).ListCount > 0 Then
           TASEL(ITABLA%).ListIndex = 0
       End If
       TABLA(ITABLA%).Visible = True
       TASEL(ITABLA%).Clear
       Exit Sub
       '
     Else
       '
       VD1$(ITABLA%) = ""
       VD2$(ITABLA%) = ""
       FORSELECHO.Hide
       Call FRESHALL
       '
   End If
   '
End Sub

Private Sub BOTCANCE_KeyDown(Keycode As Integer, Shift As Integer)
   Call REPLATECLA(Keycode)
   Call ENTRATECLA(Keycode)
End Sub

Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
   Call REPLATECLA(Keycode)
   Call ENTRATECLA(Keycode)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VD1$(ITABLA%) = ""
    VD2$(ITABLA%) = ""
End Sub

Private Sub SCROLLBA_Change()
   If CONTROBARRA% >= 0 Then
     If TABLA(ITABLA%).Visible = True Then
         REGTOPE&(ITABLA%) = FORSELECHO.SCROLLBA.Value
       ElseIf TASEL(ITABLA%).Visible = True Then
         TASEL(ITABLA%).TopIndex = SCROLLBA.Value - 1
     End If
   End If
   '
   If TABLA(ITABLA%).Visible = True Then
     Call CARGATABLA
   End If
End Sub

Private Sub SCROLLBA_KeyDown(Keycode As Integer, Shift As Integer)
   Call REPLATECLA(Keycode)
   Call ENTRATECLA(Keycode)
End Sub

Private Sub TABLA_Click(Index As Integer)
   REGACTI&(ITABLA%) = REGTOPE&(ITABLA%) + TABLA(ITABLA%).ListIndex
   If REGTOPE&(ITABLA%) >= SCROLLBA.Min Then
      On Error Resume Next
      SCROLLBA.Value = REGTOPE&(ITABLA%)
      On Error GoTo 0
   End If
End Sub

Private Sub TABLA_DblClick(Index As Integer)
    TASELEC$ = FORSELECHO.TABLA(Index).TEXT
    VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
    VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index)))
    ARCHU$ = ARCHA$(ITABLA%)
    If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
        REGG& = FORSELECHO.TABLA(Index).ListIndex + REGTOPE&(ITABLA%)
        Call SETREG(ARCHU$, REGG&)
        IACTI&(Index) = REGG& - 1
    End If
    FORSELECHO.MARCOTA(Index).Visible = False
    FORSELECHO.Hide
    Call FRESHALL
End Sub
'

Private Sub TABLA_KeyDown(Index As Integer, Keycode As Integer, Shift As Integer)
   '
   Call REPLATECLA(Keycode)
   Call ENTRATECLA(Keycode)
   '
End Sub

Private Sub TASEL_Click(Index As Integer)
   Static TOPINANT
   If TASEL(ITABLA%).TopIndex <> TOPINANT Then
     FORSELECHO.SCROLLBA.Value = TASEL(ITABLA%).TopIndex + 1
     TOPINANT = TASEL(ITABLA%).TopIndex
   End If
End Sub

Private Sub TASEL_DblClick(Index As Integer)
    TASELEC$ = FORSELECHO.TASEL(Index).TEXT
    VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
    VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index), ACOL2%(Index)))
    ARCHU$ = ARCHA$(ITABLA%)
    If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
        REGG& = Val(Mid$(TASELEC$, 3 + ACOL1%(Index) + ACOL2%(Index)))
        Call SETREG(ARCHU$, REGG&)
        IACTI&(Index) = REGG& - 1
    End If
    FORSELECHO.TASEL(Index).Visible = False
    FORSELECHO.TABLA(Index).Visible = False
    FORSELECHO.Hide
    Call FRESHALL
End Sub
'

Private Sub TASEL_KeyDown(Index As Integer, Keycode As Integer, Shift As Integer)
    Call REPLATECLA(Keycode)
    Call ENTRATECLA(Keycode)
End Sub

Private Sub Text1_Change()
   '
   Call CARTEBUS
   Call CARTASEL
   ABUSCA$(ITABLA%) = Text1.TEXT
   '
   If TABLA(ITABLA%).Visible = True Then
      FORSELECHO.SCROLLBA.Min = 1
      MASCROLL& = ULTREG&(ARCHA$(ITABLA%)) - 10
      If MASCROLL& < 1 Then MASCROLL& = 1
      FORSELECHO.SCROLLBA.Max = MASCROLL&
      If REGTOPE&(ITABLA%) > 0 Then
        If REGTOPE&(ITABLA%) <= MASCROLL& Then
          FORSELECHO.SCROLLBA.Value = REGTOPE&(ITABLA%)
        End If
      End If
    ElseIf TASEL(ITABLA%).Visible = True Then
      FORSELECHO.SCROLLBA.Min = 1
      MASCROLL& = TASEL(ITABLA%).ListCount - 10
      If MASCROLL& < 1 Then MASCROLL& = 1
      FORSELECHO.SCROLLBA.Max = MASCROLL&
      If TASEL(ITABLA%).TopIndex >= 0 Then
        If TASEL(ITABLA%).TopIndex < MASCROLL& Then
          FORSELECHO.SCROLLBA.Value = TASEL(ITABLA%).TopIndex + 1
        End If
      End If
   End If
   FORSELECHO.SCROLLBA.LargeChange = 11
   '
End Sub

Private Sub Text1_DblClick()
    Text1.SelText = Text1.TEXT
End Sub

Private Sub Text1_KeyDown(Keycode As Integer, Shift As Integer)
    If Keycode <> 32 Then
      If Keycode = 13 Or Keycode >= 27 Then
        If Keycode < 45 Then
          Call ENTRATECLA(Keycode)
        End If
      End If
    End If
End Sub

Sub REPLATECLA(Keycode)
    If Keycode = 16 Then
      Text1.SetFocus
    End If
    '
    If Keycode < 112 Then
      If Keycode > 105 Then Keycode = Keycode - 64
      If Keycode > 95 Then Keycode = Keycode - 48
    End If
End Sub


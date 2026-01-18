VERSION 5.00
Begin VB.Form FORMUSEL2 
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
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   15
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   14
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   13
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   12
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   11
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   10
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   9
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   8
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   7
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   6
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   5
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   4
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   3
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   2
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   225
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   960
      Index           =   0
      Left            =   4410
      TabIndex        =   40
      Top             =   1785
      Width           =   225
   End
   Begin VB.ListBox LINOSORT 
      Height          =   540
      Left            =   420
      TabIndex        =   39
      Top             =   1995
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox LISORT 
      Height          =   540
      Left            =   2415
      Sorted          =   -1  'True
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
      Top             =   2835
      Width           =   435
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   15
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   14
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   13
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   12
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   11
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   10
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   9
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   8
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   7
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   5
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   4
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   3
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   2
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   15
      Left            =   0
      TabIndex        =   19
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   14
      Left            =   0
      TabIndex        =   18
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   13
      Left            =   0
      TabIndex        =   17
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   12
      Left            =   0
      TabIndex        =   16
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   11
      Left            =   0
      TabIndex        =   15
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   10
      Left            =   0
      TabIndex        =   14
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   9
      Left            =   0
      TabIndex        =   13
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   8
      Left            =   0
      TabIndex        =   12
      Top             =   735
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TASEL 
      BackColor       =   &H00D0FFFF&
      Height          =   1020
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   945
      TabIndex        =   1
      Top             =   2800
      Width           =   2115
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   7
      Left            =   105
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.ListBox TABLA 
      Height          =   1020
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Image Image1 
      Height          =   10000
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10000
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
      TabIndex        =   10
      Top             =   2895
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   240
      Left            =   105
      TabIndex        =   9
      Top             =   1785
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "FORMUSEL2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XMOUSE%, YMOUSE%

Private Sub BOTCANCE_Click()
    '
    If TASEL(ITABLA%).Visible = True Then
       Text1.Text = ""
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
       FORMUSEL.Hide
       Call FRESHALL
       '
   End If
   '
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub


Private Sub TABLA_DblClick(Index As Integer)
    TASELEC$ = FORMUSEL.TABLA(Index).Text
    VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
    VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index)))
    ARCHU$ = ARCHA$(ITABLA%)
    If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
        REGG& = FORMUSEL.TABLA(Index).ListIndex + 1
        Call SETREG(ARCHU$, REGG&)
        IACTI&(Index) = REGG& - 1
    End If
    FORMUSEL.TABLA(Index).Visible = False
    FORMUSEL.Hide
    Call FRESHALL
End Sub
'

Private Sub TABLA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
       FORMUSEL.Hide
       Call FRESHALL
       Exit Sub
     ElseIf KeyAscii = 13 Then
       If FORMUSEL.TABLA(Index).ListIndex < FORMUSEL.TABLA(Index).TopIndex Then Exit Sub
       If FORMUSEL.TABLA(Index).ListIndex > FORMUSEL.TABLA(Index).TopIndex + 10 Then Exit Sub
       '
       TASELEC$ = FORMUSEL.TABLA(Index).Text
       VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
       VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index)))
       ARCHU$ = ARCHA$(ITABLA%)
       If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
           REGG& = FORMUSEL.TABLA(Index).ListIndex + 1
           Call SETREG(ARCHU$, REGG&)
           IACTI&(Index) = REGG& - 1
       End If
       FORMUSEL.TABLA(Index).Visible = False
       FORMUSEL.Hide
       Call FRESHALL
   End If
End Sub

Private Sub TABLA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      YACT% = Int(Y / TEPRUEBA.Height)
      TABLA(Index).ListIndex = TABLA(Index).TopIndex + YACT%
      ARCHU$ = TRIM$(UCase$(ARCHA$(ITABLA%)))
      If TABLA(Index).ListIndex >= 0 Then
         If TABLA(Index).ListIndex < TABLA(Index).ListCount Then
            ARGUBUS$ = TRIM$(Left$(TABLA(Index).Text, ACOL1%(Index)))
            If ARCHU$ = "ECOMAST" Or Left$(ARCHU$, 5) = "LIPRE" Then
               Call MUESTRADATCOD(ARGUBUS$)
              ElseIf ARCHU$ = "PROVED1" Or ARCHU$ = "ACREDOR" Then
               Call MUESTRADATPRO(ARGUBUS$)
            End If
         End If
      End If
   End If
End Sub

Private Sub TASEL_DblClick(Index As Integer)
    TASELEC$ = FORMUSEL.TASEL(Index).Text
    VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
    VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index), ACOL2%(Index)))
    ARCHU$ = ARCHA$(ITABLA%)
    If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
        REGG& = Val(Mid$(TASELEC$, 3 + ACOL1%(Index) + ACOL2%(Index)))
        Call SETREG(ARCHU$, REGG&)
        IACTI&(Index) = REGG& - 1
    End If
    FORMUSEL.TASEL(Index).Visible = False
    FORMUSEL.TABLA(Index).Visible = False
    FORMUSEL.Hide
    Call FRESHALL
End Sub
'

Private Sub TASEL_KeyPress(Index As Integer, KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
       Text1.Text = ""
       TASEL(ITABLA%).Visible = False
       If TASEL(ITABLA%).ListCount > 0 Then
           TASEL(ITABLA%).ListIndex = 0
       End If
       TABLA(ITABLA%).Visible = True
       TASEL(ITABLA%).Clear
       Exit Sub
     ElseIf KeyAscii = 13 Then
       If FORMUSEL.TASEL(Index).ListIndex >= FORMUSEL.TASEL(Index).TopIndex Then
          If FORMUSEL.TASEL(Index).ListIndex <= FORMUSEL.TASEL(Index).TopIndex + 10 Then
             TASELEC$ = FORMUSEL.TASEL(Index).Text
             VD1$(Index) = TRIM$(Left$(TASELEC$, ACOL1%(Index)))
             VD2$(Index) = TRIM$(Mid$(TASELEC$, 3 + ACOL1%(Index)))
             ARCHU$ = ARCHA$(ITABLA%)
             If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
                REGG& = Val(Mid$(TASELEC$, 3 + ACOL1%(Index) + ACOL2%(Index)))
                Call SETREG(ARCHU$, REGG&)
                IACTI&(Index) = REGG& - 1
             End If
             FORMUSEL.TASEL(Index).Visible = False
             FORMUSEL.Hide
             Call FRESHALL
          End If
       End If
   End If
   '
End Sub

Private Sub Text1_DBLCLICK()
    Text1.SelText = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   '
   KK% = KeyCode
   If KK% = 33 Or KK% = 34 Or KK% = 38 Or KK% = 40 Then
      If TASEL(ITABLA%).Visible = True Then
          TASEL(ITABLA%).SetFocus
        Else
          TABLA(ITABLA%).SetFocus
      End If
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   '
   KK% = KeyAscii
   If KK% = 27 Then
     If TASEL(ITABLA%).Visible = True Then
       Text1.Text = ""
       TASEL(ITABLA%).Visible = False
       If TASEL(ITABLA%).ListCount > 0 Then
           TASEL(ITABLA%).ListIndex = 0
       End If
       TABLA(ITABLA%).Visible = True
       TASEL(ITABLA%).Clear
       Exit Sub
     Else
       FORMUSEL.Hide
       Call FRESHALL
       Exit Sub
     End If
   End If
   '
   If KK% = 13 Then
       '
       If TRIM$(Text1.Text) = "" Then
           '
           TASEL(ITABLA%).Visible = False
           If TASEL(ITABLA%).ListCount > 0 Then
               TASEL(ITABLA%).ListIndex = 0
           End If
           TABLA(ITABLA%).Visible = True
           TASEL(ITABLA%).Clear
           '
         Else
           '
           ABUS$ = UCase$(TRIM$(Text1.Text))
           ASELEC$(ITABLA%) = ABUS$
           TASEL(ITABLA%).Clear
           TABLA(ITABLA%).Visible = False
           For UU& = 0 To TABLA(ITABLA%).ListCount - 1
               TABLA(ITABLA%).ListIndex = UU&
               TXTT$ = Left$(TABLA(ITABLA%).Text + Space$(80), 2 + ACOL1%(ITABLA%) + ACOL2%(ITABLA%))
               TXTT$ = TXTT$ + Str$(UU& + 1)
               If InStr(UCase$(TXTT$), ABUS$) > 0 Then
                   TASEL(ITABLA%).AddItem TXTT$
               End If
           Next UU&
           Screen.MousePointer = 1
           TABLA(ITABLA%).Visible = False
           If TASEL(ITABLA%).ListCount > 0 Then
               TASEL(ITABLA%).ListIndex = 0
           End If
           TASEL(ITABLA%).Visible = True
           '
           If TASEL(ITABLA%).ListCount > 0 Then
              TASEL(ITABLA%).SetFocus
           End If
           '
       End If
       '
   End If
   '
End Sub

Private Sub VScroll1_Change(Index As Integer)
   Call CARGATABLA
End Sub

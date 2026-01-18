VERSION 5.00
Begin VB.Form AJUCOLOR 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Colores"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4905
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
   ScaleHeight     =   2055
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Configuración de Colores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   2
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   1300
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   1
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   912
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   0
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   525
         Width           =   330
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   2
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   4
         Top             =   1300
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   1
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   3
         Top             =   912
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Height          =   960
         Left            =   315
         ScaleHeight     =   900
         ScaleWidth      =   1005
         TabIndex        =   2
         Top             =   525
         Width           =   1065
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   0
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   1
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   225
         Index           =   2
         Left            =   4095
         TabIndex        =   10
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label Label1 
         Height          =   225
         Index           =   1
         Left            =   4095
         TabIndex        =   9
         Top             =   945
         Width           =   435
      End
      Begin VB.Label Label1 
         Height          =   225
         Index           =   0
         Left            =   4095
         TabIndex        =   8
         Top             =   525
         Width           =   435
      End
   End
End
Attribute VB_Name = "AJUCOLOR"
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
    If TRIM$(Text1.TEXT) = "" Then Exit Sub
    If TRIM$(ECOARCH$("PADMAQ", TRIM$(Text1.TEXT))) = "" Then Exit Sub
    '
    COLO$ = ""
    For KK% = 0 To 2
      COLO$ = COLO$ + Chr$(HScroll1(KK%).Value Mod 256)
    Next KK%
    '
    For U& = 1 To ULTREG("PADMAQ")
       X$ = REGLEIDO$("PADMAQ", U&)
       If TRIM$(Left$(X$, 6)) = TRIM$(Text1.TEXT) Then
          Call REPLA(X$, COLO$, 97, 3)
          Call GRAREG("PADMAQ", X$, U&)
       End If
    Next U&
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("PADMAQ")
    VDEVU$ = VALACT1$("PADMAQ")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("PADMAQ")
        Call CARCOLOR
    End If
End Sub
'

Private Sub Form_Load()
    '
    Picture2(0).BackColor = RGB(255, 0, 0)
    Picture2(1).BackColor = RGB(0, 255, 0)
    Picture2(2).BackColor = RGB(0, 0, 255)
    '
End Sub

Private Sub HScroll1_Change(Index As Integer)
   Picture1.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
   INIFLEX04.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
   Label1(Index).Caption = HScroll1(Index).Value
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("PADMAQ")
    VDEVU$ = VALACT1$("PADMAQ")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("PADMAQ")
        Call CARCOLOR
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Sub CARCOLOR()
    '
    If TRIM$(Text1.TEXT) = "" Then Exit Sub
    If TRIM$(ECOARCH$("PADMAQ", TRIM$(Text1.TEXT))) = "" Then Exit Sub
    '
    COLO$ = ""
    For U& = 1 To ULTREG("PADMAQ")
       X$ = REGLEIDO$("PADMAQ", U&)
       If TRIM$(Left$(X$, 6)) = TRIM$(Text1.TEXT) Then
          COLO$ = Mid$(X$, 97, 3)
       End If
    Next U&
    '
    For KK% = 0 To 2
      HScroll1(KK%).Value = Asc(Mid$(COLO$, KK% + 1))
    Next KK%
    '
End Sub


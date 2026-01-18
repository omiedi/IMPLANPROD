VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COLOMAQ 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colores por Máquina"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5700
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
   ScaleHeight     =   4110
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Configuración de Colores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   7
      Top             =   2205
      Width           =   4215
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   2
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   1300
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   1
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   912
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Height          =   185
         Index           =   0
         Left            =   3675
         ScaleHeight     =   120
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   525
         Width           =   330
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   2
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   11
         Top             =   1300
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   1
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   10
         Top             =   912
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Height          =   960
         Left            =   315
         ScaleHeight     =   900
         ScaleWidth      =   1005
         TabIndex        =   9
         Top             =   525
         Width           =   1065
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   170
         Index           =   0
         LargeChange     =   16
         Left            =   1785
         Max             =   255
         TabIndex        =   8
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Equipo / Puesto de Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   4215
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   6
         Text            =   " "
         Top             =   525
         Width           =   1380
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
         Left            =   3780
         TabIndex        =   5
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Máquina:"
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
         Left            =   0
         TabIndex        =   3
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4620
      Picture         =   "Colomaq.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
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
      Height          =   750
      Left            =   4620
      Picture         =   "Colomaq.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "Colomaq.frx":0454
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "Colomaq.frx":0688
      TabIndex        =   16
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "COLOMAQ"
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
    Call APLICACIONSKINS(Me)
End Sub

Private Sub HScroll1_Change(Index As Integer)
   Picture1.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
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

VERSION 5.00
Begin VB.Form PRIORIENT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analisis de Prioridades de Entregas"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5520
   Icon            =   "PRIORIENT.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   8
      Top             =   2520
      Width           =   4110
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2205
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   12
         Top             =   420
         Width           =   1695
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
            Picture         =   "PRIORIENT.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Analisis de Prioridad"
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
            Index           =   0
            Left            =   630
            TabIndex        =   14
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   9
         Top             =   420
         Width           =   1695
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
            Left            =   0
            Picture         =   "PRIORIENT.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Consulta por Liquidacion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Nivel de Prioridad"
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
            Left            =   600
            TabIndex        =   11
            Top             =   120
            Width           =   1100
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Re-Cálculo General"
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
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ultimo Calculo"
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
         Left            =   105
         TabIndex        =   15
         Top             =   1260
         Width           =   3900
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2940
            TabIndex        =   19
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   945
            TabIndex        =   18
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
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
            Index           =   12
            Left            =   1260
            TabIndex        =   17
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label2 
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
            Index           =   13
            Left            =   -1365
            TabIndex        =   16
            Top             =   420
            Width           =   2220
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command28 
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
            Picture         =   "PRIORIENT.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Recálculo de Si- tuación de Stocks y Pedidos"
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
            Left            =   735
            TabIndex        =   7
            Top             =   0
            Width           =   1635
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4515
      ScaleHeight     =   75
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   2940
      Width           =   750
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command29 
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
      Height          =   750
      Left            =   4545
      Picture         =   "PRIORIENT.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1120
      Width           =   750
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
      Height          =   750
      Left            =   4545
      Picture         =   "PRIORIENT.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   320
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3960
      Picture         =   "PRIORIENT.frx":0EBC
      Top             =   3255
      Width           =   2010
   End
End
Attribute VB_Name = "PRIORIENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*CANAPRIO")
    Command2.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    Dim STOACT(32767, 3), PPENT(32767)
    Screen.MousePointer = 11
    FIN& = NUMAS%
    For KU& = 1 To NUMAS%
      Call TRACE(20, KU&, FIN&)
      CIXI% = KU&
      STOACT(CIXI%, 0) = STODEPOS(CIXI%, 1) + STODEPOS(CIXI%, 2) + STODEPOS(CIXI%, 3)
      STOACT(CIXI%, 1) = STODEPOS(CIXI%, 9)
      STOACT(CIXI%, 2) = STODEPOS(CIXI%, 10)
    Next KU&
    '
    Call OPPEDIDO.RECAPEDPE
    FIN& = ULTREG&("PEDDETP")
    For KU& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PEDDETP", KU&)
      CIXI% = CVI(Mid$(XX$, 33, 2))
      STOACT(CIXI%, 3) = STOACT(CIXI%, 3) + CVS(Mid$(XX$, 79, 4)) - CVS(Mid$(XX$, 95, 4))
    Next KU&
    '
    REBLA$ = REGBLAN$("ANAPRIO")
    JJ& = 0
    FIN& = NUMAS%
    For KU& = 1 To NUMAS%
      Call TRACE(20, KU&, FIN&)
      If STOACT(KU&, 3) >= 0.05 Then   ' SOLO LOS QUE TIENEN PEDIDOS
        'For KII% = 0 To 3
        '  If (STOACT(KU&, KII%)) >= 0.05 Then GoTo 25
        'Next KII%
        'GoTo 29
        '
25      XX$ = REBLA$
        CIXI% = KU&
        Call REPLA(XX$, MKI$(CIXI%), 1, 2)
        SUSTO = 0
        For KII% = 0 To 2
          SUSTO = SUSTO + STOACT(KU&, KII%)
          Call REPLA(XX$, MKS$(STOACT(KU&, KII%)), 4 * KII% + 3, 4)
        Next KII%
        Call REPLA(XX$, MKS$(SUSTO), 15, 4)
        Call REPLA(XX$, MKS$(STOACT(KU&, 3)), 19, 4) ' PEDIDOS
        STONE = SUSTO - STOACT(KU&, KII%)
        Call REPLA(XX$, MKS$(STONE), 23, 4)
        STOMI = STOMIN(CIXI%)
        Call REPLA(XX$, MKS$(STOMI), 27, 4)
        '
        If STONE <= 0 Then
             PRIO% = 10
             PRIO2% = 16
           Else
             PRIO% = 0
             PRIO2% = 0
             If STOMI > 0 Then
               PRIO% = Int(10 - STONE / STOMI)
               If PRIO% < 0 Then PRIO% = 0
               PRIO2% = PRIO%
             End If
        End If
        '
        Call REPLA(XX$, MKI$(PRIO%), 31, 2)
        Call REPLA(XX$, String$(2 * PRIO2%, "*"), 33, 32)
        '
        JJ& = JJ& + 1
        Call GRAREG("ANAPRIO", XX$, JJ&)
      End If
29  Next KU&
    '
    Call SETULTREG("ANAPRIO", JJ&)
    Call GRARGEN("ANAPRIO", "")
    Call FILESORT("ANAPRIO", "", 1, 1)
    '
    Screen.MousePointer = 1
    Call COMUNI("Calculo Completo")
    '
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call FINAL("*LANAPRIO")
   Command3.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FEX = FEGEAR%("ANAPRIO")
    Label1 = FECHATEX$(FEX)
    If FEX < 1 Then Label1 = "00/00/00"
    HOX% = HOGEAR%("ANAPRIO")
    Label4 = HORATEX$(HOX%)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


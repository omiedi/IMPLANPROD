VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DIARPAN07 
   BackColor       =   &H00D0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario General"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   10080
      TabIndex        =   2
      Top             =   0
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Busqueda por Texto"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   6800
      Left            =   11470
      TabIndex        =   16
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   6800
      Left            =   9975
      TabIndex        =   15
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   6800
      Left            =   8470
      TabIndex        =   14
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   6800
      Left            =   4935
      TabIndex        =   13
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6800
      Left            =   3465
      TabIndex        =   12
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6800
      Left            =   2500
      TabIndex        =   11
      Top             =   1000
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6800
      Left            =   1155
      TabIndex        =   10
      Top             =   1000
      Width           =   15
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   105
      TabIndex        =   7
      Top             =   1470
      Width           =   11670
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   11610
      TabIndex        =   8
      Top             =   1050
      Width           =   11670
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"DIARPAN07.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   9
         Top             =   105
         Width           =   11565
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10560
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   600
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DIARPAN07.frx":00A0
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "DIARPAN07.frx":00B0
      TabIndex        =   1
      Top             =   10000
      Width           =   11655
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "DIARPAN07.frx":0A83
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "DIARPAN07.frx":0CB7
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE DE LA EMPRESA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   6
      Top             =   210
      Width           =   5565
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EJERCICIO CONTABLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   4455
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   7920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":0D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":104B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":1365
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":167F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":1999
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":1CB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":1FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":22E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":2601
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":291B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":2C35
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":2E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":2F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN07.frx":323B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Libro Diario General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5475
   End
   Begin VB.Menu Files 
      Caption         =   "Archivos"
      Begin VB.Menu Print 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edicion"
      Begin VB.Menu Busca 
         Caption         =   "Nueva Busqueda"
      End
      Begin VB.Menu rebusca 
         Caption         =   "Repetir Ultima        <F3>"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "DIARPAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
'
Private Sub Command19_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command18_Click()
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   
   Command1.Enabled = True
End Sub

Private Sub DBGrid1_DblClick()
   'MsgBox DBGrid1.TEXT
   If DbGrid1.Col = 1 Then
     'ASIDIAR.Hide
     ACTICARGA% = 1
     ASIDIAR.Ltext2 = DbGrid1.TEXT
     On Error Resume Next
     ASIDIAR.Show 1
     If Err Then
       Hide
     End If
     On Error GoTo 0
     'ASIDIAR.Show 1
   End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
      Call rebusca_Click
    End If
End Sub

Private Sub Exit_Click()
    Call CIERRATO
End Sub

Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Caption = Caption + " - " + TRIM$(EJACT$)
    '
    PPXX% = InStr(EPAC$, " -"): If PPXX% < 1 Then PPXX% = 1 + Len(EPAC$)
    Label1 = Left$(EPAC$, PPXX% - 1)
    Label2 = DENOEJ$
    Call APLICACIONSKINS(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Private Sub Busca_Click()
   Call Search
End Sub



Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
     Call rebusca_Click
   End If
End Sub

Private Sub rebusca_Click()
   If ARBUS$ <> "" Then
     VUELTA% = 0
     INI& = List1.ListIndex + 2
     '
10   For U& = INI& To List1.ListCount
       If InStr(UCase$(List1.List(U& - 1)), ARBUS$) > 0 Then
         List1.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     '
     If VUELTA% < 1 Then
       VUELTA% = 1
       INI& = 1
       GoTo 10
     End If
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   '
   Select Case UCase$(Button.Key)
      Case "PRINT"
        Call LIDIARPRIN
        'Call List'comente esta linea xq no compilaba, xq no encuentra el sub "list"
      Case "SEARCH"
        Call Search
      Case "EXIT"
        Call CIERRATO
   End Select
   '
End Sub

Sub LIDIARPRIN()
   '
   DES% = FECHANUM(DIARGEN07.Label5)
   HAS% = FECHANUM(DIARGEN07.Label6)
   CORRESP$ = FECHATEX$(DES%) + " al " + FECHATEX$(HAS%)
   '
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     XX$ = REGBLAN$("LIDIARG")
       FEXI% = FECHANUM(Left$(TTXX$, 8))
     Call REPLA(XX$, MKI$(FEXI%), 1, 2)
       IDASIENTO$ = Mid$(TTXX$, 11, 12)
     Call REPLA(XX$, IDASIENTO$, 3, 12)
       NUAS& = Val(Mid$(TTXX$, 24, 7))
     Call REPLA(XX$, MKS$(NUAS&), 15, 4)
       CUENTA$ = Mid$(TTXX$, 33, 12)
     Call REPLA(XX$, CUENTA$, 19, 16)
       REFE$ = Mid$(TTXX$, 47, 32)
     Call REPLA(XX$, REFE$, 35, 32)
       IDEBE# = XVALO(Mid$(TTXX$, 80, 14))  'SE CAMBIO POR XVALO PORQUE EL IMPORTE ESTA FORMATEADO CON EL FORMATO "G"
     Call REPLA(XX$, MKD$(IDEBE#), 67, 8)
       IHABE# = XVALO(Mid$(TTXX$, 94, 14))  'SE CAMBIO POR XVALO PORQUE EL IMPORTE ESTA FORMATEADO CON EL FORMATO "G"
     Call REPLA(XX$, MKD$(IHABE#), 75, 8)
       DEMOVI$ = Mid$(TTXX$, 129, 40)
     Call REPLA(XX$, DEMOVI$, 87, 40)
     If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Or TRIM$(IDASIENTO$) = "" Then
       ' LO QUE SIGUE ES PARA SUPRIMIR MOVIMIENTOS ANULADOS EN LISTADO
       If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Then
         JJ& = JJ& + 1
         Call GRAREG("LIDIARG", XX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("LIDIARG", JJ&)
   Call CIERRARCH("LIDIARG")
   '
   Call HEADERS("LIDIARG", "")
   Call HEADERS("LIDIARG", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIDIARG", "Libro Diario General")
   '
End Sub

Sub Search()
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     List1.ListIndex = (-1)
     For U& = 1 To List1.ListCount
       If InStr(UCase$(List1.List(U& - 1)), ABUSQUE$) > 0 Then
         List1.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Sub CIERRATO()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

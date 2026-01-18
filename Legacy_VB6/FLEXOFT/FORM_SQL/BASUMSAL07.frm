VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form BASUMSAL07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario General"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   10080
      TabIndex        =   1
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
            Key             =   "Print"
            Object.ToolTipText     =   "Emite Sumas y Saldos por Impresora"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Help"
            Object.ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Cierra y Abandona la Aplicación"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox MARCOBARRA 
      Height          =   195
      Left            =   10320
      ScaleHeight     =   135
      ScaleWidth      =   1185
      TabIndex        =   15
      Top             =   550
      Width           =   1245
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   16
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   105
      ScaleHeight     =   420
      ScaleWidth      =   11625
      TabIndex        =   11
      Top             =   840
      Width           =   11655
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"BASUMSAL07.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   140
         Width           =   11595
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6150
      ItemData        =   "BASUMSAL07.frx":00DA
      Left            =   105
      List            =   "BASUMSAL07.frx":00DC
      TabIndex        =   10
      Top             =   1260
      Width           =   11670
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Texto;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "BASUMSAL07.frx":00DE
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "BASUMSAL07.frx":0312
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   9
      Top             =   7455
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   10290
      TabIndex        =   8
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   8925
      TabIndex        =   7
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   7560
      TabIndex        =   6
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   4830
      TabIndex        =   5
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   6195
      TabIndex        =   4
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EJERCICIO CONTABLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE DE LA EMPRESA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   9480
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
            Picture         =   "BASUMSAL07.frx":038C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":06A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":09C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":0FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":130E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":1628
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":1942
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":1C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":1F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":246A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":257C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BASUMSAL07.frx":2896
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
      Caption         =   "Balance de Sumas y Saldos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
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
Attribute VB_Name = "BASUMSAL07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$
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
   If DbGrid1.COL = 1 Then
     'ASIDIAR.Hide
     ACTICARGA% = 1
     ASIDIAR07.Ltext2 = DbGrid1.TEXT
     Hide
     'ASIDIAR.Show 1
   End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If ARBUS$ <> "" Then
     If KeyCode = 114 Then
        Data1.Recordset.FindNext ARBUS$
     End If
   End If
End Sub

Private Sub Exit_Click()
   Call CIERRATO
End Sub

Private Sub Form_Load()
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    PPXX% = InStr(EPAC$, " -"): If PPXX% < 1 Then PPXX% = 1 + Len(EPAC$)
    Label1 = Left$(EPAC$, PPXX% - 1)
    Label2 = OPBALAN07.Text6

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Private Sub Busca_Click()
   Call Search
End Sub


Private Sub rebusca_Click()
   If ARBUS$ <> "" Then
     Data1.Recordset.FindNext ARBUS$
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   '
   Select Case UCase$(Button.Key)
      Case "SEARCH"
        Call Search
      Case "PRINT"
        Call PRIBASUSAL
      Case "EXIT"
        ' TTXX$ = "": For U& = 0 To 6: TTXX$ = TTXX$ + Str$(DBGrid1.Columns(U&).Width): Next
        Call CIERRATO
   End Select
   '
End Sub

Sub PRIBASUSAL()
   '
'   JJ& = 0
'   FIN& = BALASEC1.RecordCount
'   '
'   With BALASEC1   ' Data1.Recordset
'     For U& = 1 To FIN&
'       If U& = 1 Then
'            .MoveFirst
'          Else
'            .MoveNext
'       End If
'       '
'       ZZ$ = REGBLAN$("BASUSAL")
'       Call REPLA(ZZ$, .Fields(0).Value, 1, 12)
'       Call REPLA(ZZ$, .Fields(1).Value, 13, 52)
'       Call REPLA(ZZ$, MKD$(.Fields(3).Value), 65, 8)
'       Call REPLA(ZZ$, MKD$(.Fields(4).Value), 73, 8)
'       Call REPLA(ZZ$, MKD$(.Fields(5).Value), 81, 8)
'       Call REPLA(ZZ$, MKD$(.Fields(6).Value), 89, 8)
'       Call REPLA(ZZ$, MKD$(.Fields(7).Value), 97, 8)
'       JJ& = JJ& + 1
'       Call GRAREG("BASUSAL", ZZ$, JJ&)
'     Next U&
'   End With
'   '
'   Call SETULTREG("BASUSAL", JJ&)
'   Call CIERRARCH("BASUSAL")
'   '
   MESASEL$ = OPBALAN07.Text6
   Call TRACENTRAL("BASUSAL", "")
   Call HEADERS("BASUSAL", "")
   Call HEADERS("BASUSAL", "Corresponde a: " + MESASEL$)
   '
   Screen.MousePointer = 1
   '
   Call CONECRUN("*BASUSAL", "Balance de Sumas y Saldos")
   '
End Sub
'
Sub Search()
   ABUSQUE$ = InputBox$("Busqueda por Texto")
   If ABUSQUE$ <> "" Then
     Call MENSERR(24, "Imposible Completar la Transacción Solicitada")
     Exit Sub
     '
     Data1.Recordset.MoveFirst
     URECORD& = Data1.Recordset.RecordCount
     '
     ARBUS$ = "CODEXCUE LIKE '*" & ABUSQUE$ & "*' "
     ARBUS$ = ARBUS$ + "OR DENOCUE LIKE '*" & ABUSQUE$ & "*' "
     Data1.Recordset.FindNext ARBUS$
   End If
End Sub

Sub CIERRATO()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

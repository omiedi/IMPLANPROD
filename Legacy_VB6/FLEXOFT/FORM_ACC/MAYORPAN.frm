VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MAYORPAN 
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes Contables"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   9975
      TabIndex        =   17
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
            Object.ToolTipText     =   "Listado por Rango de Fechas"
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
   Begin VB.Frame Frame6 
      Caption         =   "Frame2"
      Height          =   6230
      Left            =   2550
      TabIndex        =   20
      Top             =   1390
      Width           =   15
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Concilia"
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
      Left            =   8820
      TabIndex        =   19
      ToolTipText     =   "Conciliación de Cartera de Valores"
      Top             =   225
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame9"
      Height          =   6230
      Left            =   11520
      TabIndex        =   18
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6230
      Left            =   1100
      TabIndex        =   16
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6230
      Left            =   2785
      TabIndex        =   15
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   6230
      Left            =   10020
      TabIndex        =   11
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   6230
      Left            =   7020
      TabIndex        =   12
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   6230
      Left            =   8520
      TabIndex        =   10
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6230
      Left            =   3700
      TabIndex        =   9
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   7185
      Left            =   150
      TabIndex        =   0
      Top             =   1080
      Width           =   11670
      Begin VB.PictureBox Picture1 
         Height          =   460
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   11595
         TabIndex        =   2
         Top             =   6510
         Width           =   11655
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8415
            TabIndex        =   4
            Top             =   60
            Width           =   1380
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6930
            TabIndex        =   3
            Top             =   60
            Width           =   1380
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "MAYORPAN.frx":0000
         Left            =   0
         List            =   "MAYORPAN.frx":0002
         TabIndex        =   5
         Top             =   400
         Width           =   11670
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   11595
         TabIndex        =   6
         Top             =   0
         Width           =   11655
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   $"MAYORPAN.frx":0004
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   7
            Top             =   105
            Width           =   11355
         End
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11655
         Y1              =   6970
         Y2              =   6970
      End
      Begin VB.Line Line2 
         X1              =   11655
         X2              =   11655
         Y1              =   0
         Y2              =   6980
      End
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   8925
      Top             =   0
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
            Picture         =   "MAYORPAN.frx":00A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":03BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":06D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":09F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":0D0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":1027
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":1341
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":165B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":1975
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":1C8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":1FA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":2183
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":2295
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN.frx":25AF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Libro Mayor por Cuenta"
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
      TabIndex        =   14
      Top             =   120
      Width           =   5475
   End
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6615
      TabIndex        =   13
      Top             =   630
      Width           =   5175
   End
   Begin VB.Label Ltext2 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6615
      TabIndex        =   8
      Top             =   225
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   1
      Top             =   270
      Width           =   1275
   End
End
Attribute VB_Name = "MAYORPAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Sub CONCILIAVIE()
   Screen.MousePointer = 11
10 For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To List1.ListCount
         TTXX2$ = List1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           If TXD$ = TXH$ Then
             GoSub RECUSALCHE
           End If
           Call REPLA(TTXX1$, Space$(14), 66, 14)
           List1.RemoveItem U& - 1
           List1.AddItem TTXX1$, U& - 1
           Call REPLA(TTXX2$, Space$(14), 80, 14)
           List1.RemoveItem U2& - 1
           List1.AddItem TTXX2$, U2& - 1
           GoTo 10
         End If
       Next U2&
     End If
19 Next U&
   '
20 TRANSPOR# = 0
   For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
     If TRIM$(Mid$(List1.List(U& - 1), 66, 28)) = "" Then
         List1.RemoveItem (U& - 1)
         GoTo 20
       Else
         TTXX1 = List1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         List1.RemoveItem U& - 1
         List1.AddItem TTXX1$, U& - 1
     End If
   Next U&
   Screen.MousePointer = 1
   List1.Refresh
   DoEvents
   Exit Sub
   '
RECUSALCHE:
   IMCHEX$ = TRIM$(IDEX$)
   IDEVA$ = TRIM$(Mid$(TTXX1$, 129, 64))
   FESALI% = FECHANUM(Left$(TTXX2$, 8))
   DOCUSA$ = TRIM$(Mid$(TTXX2$, 10, 14))
   DESTIN$ = TRIM$(Mid$(TTXX2$, 35, 30))
   HOII% = HOY(HOS$)
   '
   SQLX$ = "DescriMov = '" & IDEVA$ & "' "
   SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
   SQLX$ = SQLX$ + "AND Status = 1"
   '
   Call ABRECABAN
            VUELTA% = 0: ENCUEN% = 0
   With CajaBanc
     .FindFirst SQLX$
     If .NoMatch = False Then
       .Edit
       !FEDEPOSENT = CVDAT(FESALI%)
       !Document = DOCUSA$
       !DESTICHEQ = DESTIN
       !REFESALCHE = "Conciliación Documento " + DOCUSA$
       !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
       !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
       !Status = 2                     ' Marca Salida de Cartera
       !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
       .Update
       Call COMUNI("Se ha Conciliado el Valor '" + IDEVA$ + "'")
    End If
  End With
Return
   '
End Sub

Private Sub Command1_Click()
   '
   CUEII% = CUEINT%(Ltext2)
   If CUEII% < 1 Then Exit Sub
   '
   If RECLEN("LIMAYOAC") < 128 Then
      Call MENSERR(24, "Proceso Imposible. Controlador 'LIMAYOAC.RFS'\con Longitud de Registro Incorrecta. Consulte.")
      Exit Sub
   End If
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = !IDEBEPASE
       IHABE# = !IHABEPASE
       Z$ = REGBLAN$("LIMAYOAC")
       Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
       Call REPLA(Z$, !NUIDASIEN, 3, 12)
         MACON$ = ""
         On Error Resume Next
         MACON$ = TRIM$(!MARCONSI)
         On Error GoTo 0
       Call REPLA(Z$, MACON$, 15, 1)
       Call REPLA(Z$, "D", 16, 1)
         DEPASE$ = !REFEPASE
       Call REPLA(Z$, DEPASE$, 17, 32)
       Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
       Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         CENCO$ = !CenCos
       Call REPLA(Z$, CENCO$, 65, 8)
       JJ& = JJ& + 1
       Call GRAREG("LIMAYOAC", Z$, JJ&)
     .MoveNext
     Loop
   End With
   '
30 On Error GoTo 0
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true)) "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = !IDEBECOMP
       IHABE# = !IHABECOMP
       Z$ = REGBLAN$("LIMAYOAC")
       Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
       Call REPLA(Z$, !CodiCom_Cor, 3, 12)
         MACON$ = ""
         On Error Resume Next
         MACON$ = TRIM$(!MARCONSI)
         On Error GoTo 0
       Call REPLA(Z$, MACON$, 15, 1)
       Call REPLA(Z$, "C", 16, 1)
         DEPASE$ = !RefeCom
         DE2PASE$ = !DescriMov
       Call REPLA(Z$, DEPASE$, 17, 32)
       Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
       Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         CENCO$ = !CenCosto
       Call REPLA(Z$, CENCO$, 65, 8)
       Call REPLA(Z$, DE2PASE$, 73, 56)
       JJ& = JJ& + 1
       Call GRAREG("LIMAYOAC", Z$, JJ&)
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("LIMAYOAC", JJ&)
   Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   '
   MAYORPAN.List1.Clear
   Z$ = Space$(128 + 64)
     FEX = DIANTE(COMEJ%)
     FEXA$ = FECHATEX$(FEX)
   Call REPLA(Z$, FEXA$, 1, 8)
     DEPASE$ = "TRANSPORTE .............................."
   Call REPLA(Z$, DEPASE$, 35, 30)
   Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
   MAYORPAN.List1.AddItem Z$
   MAYORPAN.List1.AddItem Space$(95) + " =========== "
   '
   SUDEBE1# = 0: SUHABE1# = 0
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("LIMAYOAC", U&)
     Z$ = Space$(128 + 64 + 8)
       FEX = CVI(Left$(X$, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 1, 8)
       COCOM$ = TRIM$(Mid$(X$, 3, 12))
     Call REPLA(Z$, COCOM$, 10, 14)
       MACON$ = TRIM$(Mid$(X$, 15, 1))
     Call REPLA(Z$, MACON$, 24, 2)
       ORIMO$ = TRIM$(Mid$(X$, 16, 1))
     Call REPLA(Z$, ORIMO$, 128, 1)
       DEPASE$ = TRIM$(Mid$(X$, 17, 32))
       DE2PASE$ = TRIM$(Mid$(X$, 73, 56))
     Call REPLA(Z$, DEPASE$, 35, 30)
     Call REPLA(Z$, DE2PASE$, 129, 64)
       CENCO$ = TRIM$(Mid$(X$, 65, 8))
     Call REPLA(Z$, CENCO$, 193, 8)
       IDEBE# = CVD(Mid$(X$, 49, 8))
       IHABE# = CVD(Mid$(X$, 57, 8))
     If IDEBE# > 0 Then
         Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 14, 2, 2), 66, 14)
         SUDEBE1# = SUDEBE1# + IDEBE#
       Else
         Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 14, 2, 2), 80, 14)
         SUHABE1# = SUHABE1# + IHABE#
     End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
     Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
     MAYORPAN.List1.AddItem Z$
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   On Error GoTo 0
   If MAYORPAN.List1.ListCount > 0 Then
     MAYORPAN.List1.ListIndex = MAYORPAN.List1.ListCount - 1
   End If
   MAYORPAN.List1.ToolTipText = ""
   '
   MAYORPAN.Label21 = TRIM$(FORMATNUM$(SUDEBE1#, "F11.2"))
   MAYORPAN.Label22 = TRIM$(FORMATNUM$(SUHABE1#, "F11.2"))
   Screen.MousePointer = 1
   '
   Screen.MousePointer = 11
50 For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To List1.ListCount
         TTXX2$ = List1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           If TXD$ = TXH$ Then
             Call REPLA(TTXX1$, Space$(14), 66, 14)
             List1.RemoveItem U& - 1
             List1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, Space$(14), 80, 14)
             List1.RemoveItem U2& - 1
             List1.AddItem TTXX2$, U2& - 1
             GoTo 50
           End If
         End If
       Next U2&
     End If
19 Next U&
   '
55 For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To List1.ListCount
         TTXX2$ = List1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           TTXX$ = "Conciliación Imperfecta:\" + TRIM$(Mid$(TTXX1$, 129, 64)) + "\" + TRIM$(Mid$(TTXX2$, 129, 64))
           If COMALTER%(TTXX$ + "\\Conciliar\Continuar") = 1 Then
             Call REPLA(TTXX1$, Space$(14), 66, 14)
             List1.RemoveItem U& - 1
             List1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, Space$(14), 80, 14)
             List1.RemoveItem U2& - 1
             List1.AddItem TTXX2$, U2& - 1
             DoEvents
             List1.Refresh
             GoTo 59
           End If
         End If
       Next U2&
     End If
59 Next U&
   '
60 TRANSPOR# = 0
   For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
63   If TRIM$(Mid$(List1.List(U& - 1), 66, 28)) = "" Then
         On Error Resume Next
         List1.RemoveItem (U& - 1)
         If Err Then GoTo 89
         DoEvents
         List1.Refresh
         GoTo 63
       Else
         TTXX1 = List1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         List1.RemoveItem U& - 1
         List1.AddItem TTXX1$, U& - 1
     End If
   Next U&
89 On Error GoTo 0
   Screen.MousePointer = 1
   List1.Refresh
   DoEvents
   '
   Call GESVAL04.GENLICHECA(3)
   '
   For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To ULTREG&("!LICARCHE")
         TTXX2$ = REGLEIDO$("!LICARCHE", U2&)
         IHAX$ = TRIM$(FORMATNUM$(CVD(Mid$(TTXX2$, 49, 8)), "F14.2"))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 1, 40), " ", "")
           If TXD$ = TXH$ Then
             Call REPLA(TTXX1$, Space$(14), 66, 14)
             List1.RemoveItem U& - 1
             List1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, MKD$(0), 49, 8)
             Call GRAREG("!LICARCHE", TTXX2$, U2&)
             DoEvents
             List1.Refresh
             GoTo 91
           End If
         End If
       Next U2&
     End If
91 Next U&
   '
   TRANSPOR# = 0
   For U& = 1 To List1.ListCount
     TTXX1$ = List1.List(U& - 1)
92   If TRIM$(Mid$(List1.List(U& - 1), 66, 28)) = "" Then
         On Error Resume Next
         List1.RemoveItem (U& - 1)
         If Err Then GoTo 93
         DoEvents
         List1.Refresh
         GoTo 92
       Else
         TTXX1 = List1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         List1.RemoveItem U& - 1
         List1.AddItem TTXX1$, U& - 1
         DoEvents
         List1.Refresh
     End If
   Next U&
93 On Error GoTo 0
   Screen.MousePointer = 1
   List1.Refresh
   DoEvents
   '
   For U2& = 1 To ULTREG&("!LICARCHE")
     TTXX2$ = REGLEIDO$("!LICARCHE", U2&)
     If Abs(CVD(Mid$(TTXX2$, 49, 8))) > 0.005 Then
       IHAX$ = TRIM$(FORMATNUM$(CVD(Mid$(TTXX2$, 49, 8)), "F14.2"))
       BANUFE$ = TRIM$(Left$(TTXX2$, 40))
       Call COMUNI("Valor en Cartera\sin Entrada en Cuenta Activa:\  \" + BANUFE$ + " - " + IHAX$)
     End If
   Next U2&
   '
   If List1.ListCount > 0 Then
       Call COMUNI("Hay Movimientos en Cuenta Activa\No Conciliables con Cartera.")
     Else
       Call COMUNI("Fin de Control")
   End If
   '
End Sub

Private Sub Label2_Click()

End Sub




Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
    End If

End Sub

Private Sub List1_Click()
   List1.ToolTipText = TRIM$(Mid$(List1.TEXT, 129, 56))
End Sub

Private Sub List1_DblClick()
   ORIMO$ = TRIM$(Mid$(List1.TEXT, 128, 1))
   If ORIMO$ = "C" Then
     Call MUEFACO
   End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  '\\\OT-07-0046-RG-01/02/07///
   If KeyCode = 114 Then
       If ARBUS$ <> "" Then
          VUELTA% = 0
          INI& = List1.ListIndex + 2
    '
10        For U& = INI& To List1.ListCount
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
   End If
   '\\\OT-07-0046-RG-FIN///
End Sub

Private Sub Ltext2_Change()
   If Ltext2 = TRIM$(CUECARCHE$) Then
     Command1.Visible = True   ' PARA HABILITAR CONCILIACIóN DE CARTERA
   End If
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   '
   Select Case UCase$(Button.Key)
      Case "SEARCH"
        Call Search
      Case "PRINT"
        Call LIMAYORPRIN
      Case "EXIT"
        Unload Me
   End Select
   '
End Sub
Sub Search()
  '\\\OT-07-0046-RG-01/02/07///
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     List1.ListIndex = (-1)
     For U& = 1 To List1.ListCount
       If InStr(UCase$(List1.List(U& - 1)), ABUSQUE$) > 0 Then
         List1.ListIndex = U& - 1
         List1.SetFocus
         Exit Sub
       End If
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
   '\\\OT-07-0046-RG-FIN///
End Sub
'
Sub LIMAYORPRIN()
   '
10 VDEF$ = MKI$(COMEJ%) + MKI$(FINEJ%)
   RANFE$ = OBJPLA$("DESHASFECHA", VDEF$)
   If RANFE$ = "" Then Exit Sub
   '
   DES% = CVI(Left$(RANFE$, 2))
   HAS% = CVI(Mid$(RANFE$, 3, 2))
   If HAS% < DES% Then
     Call MENSERR(24, "Rango de Fechas no procesable.")
     GoTo 10
   End If
   CORRESP$ = FECHATEX$(DES%) + " al " + FECHATEX$(HAS%) + " - " + TRIM$(Label4) + " (" + TRIM$(Ltext2) + ")"
   '
   If ECOARCH$("CUEBANC", TRIM$(Ltext2)) <> "" Then
     OPXX% = COMALTER%("La Cuenta Elegida Figura en\la Tabla de Cuentas Bancarias.\Elija el Formato de Impresión:\\Mayor por Cuenta\Libro de Bancos")
     If OPXX% = 1 Then
         GoTo 15
       ElseIf OPXX% = 2 Then
         Call LIBROBAN(DES%, HAS%, CORRESP$)
         Exit Sub
       Else
         Exit Sub
     End If
   End If
   '
15 JJ& = 0
   YATRAN% = 0
   TRANSPO# = 0
   '
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     FEMO% = FECHANUM(Left$(TTXX$, 8))
     IDEBE# = Val(Mid$(TTXX$, 66, 14))
     IHABE# = Val(Mid$(TTXX$, 80, 14))
     If U& = 1 Then
       If TRIM$(Mid$(TTXX$, 66, 28)) = "" Then
         IDEBE# = Val(Mid$(TTXX$, 94, 14))
         If IDEBE# < 0 Then
           IHABE# = Abs(IDEBE#)
           IDEBE# = 0
         End If
       End If
     End If
     '
     If FEMO% < DES% Then
         TRANSPO# = TRANSPO# + IDEBE# - IHABE#
         GoTo 29
       ElseIf FEMO% <= HAS% Then
         If YATRAN% < 1 Then
           If Abs(TRANSPO#) >= 0.005 Then
             XX$ = REGBLAN$("LIMAYOW")
               DESDE = DES%
             Call REPLA(XX$, FECHATEX$(DIANTE(DESDE)), 1, 8)
             Call REPLA(XX$, "Transporte ........................................................................", 37, 48)
             Call REPLA(XX$, MKD$(TRANSPO#), 101, 8)
             JJ& = JJ& + 1
             Call GRAREG("LIMAYOW", XX$, JJ&)
             '
             XX$ = REGBLAN$("LIMAYOW")
             JJ& = JJ& + 1
             Call GRAREG("LIMAYOW", XX$, JJ&)
             YATRAN% = 1
           End If
         End If
         '
         ' LO QUE SIGUE ES PARA SUPRIMIR LOS COMPROBANTES ANULADOS
         If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Then
             FFAX$ = Left$(TTXX$, 8)
           Call REPLA(XX$, FFAX$, 1, 8)
             COCOM$ = Mid$(TTXX$, 10, 16)
           Call REPLA(XX$, COCOM$, 9, 16)
             CENCO$ = TRIM$(Mid$(TTXX$, 193, 8))
           Call REPLA(XX$, CENCO$, 25, 8)
             REFE$ = Mid$(TTXX$, 35, 30)
           Call REPLA(XX$, REFE$, 37, 48)
           Call REPLA(XX$, MKD$(IDEBE#), 85, 8)
           Call REPLA(XX$, MKD$(IHABE#), 93, 8)
             ITRAN# = Val(Mid$(TTXX$, 94, 14))
           Call REPLA(XX$, MKD$(ITRAN#), 101, 8)
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOW", XX$, JJ&)
         End If
     End If
29 Next U&
   '
   Call SETULTREG("LIMAYOW", JJ&)
   Call CIERRARCH("LIMAYOW")
   '
   Call HEADERS("LIMAYOW", "")
   Call HEADERS("LIMAYOW", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIMAYOW", "Libro Mayor por Cuenta")
   '
End Sub

Sub LIBROBAN(DES%, HAS%, CORRESP$)
   '
   FEINI = Int(CDbl(CVDAT(COMEJ%)))
   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   CUEII% = CUEINT%(Ltext2)
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   'SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   Call COPYSTRU("LIBROBAN", "LIBROBAX")
   '
   JJ& = 0
   TRANSPOR# = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       NUIDAS$ = !NUIDASIEN
       Encasien.FindFirst "NuIdAsien = '" & NUIDAS$ & "' AND StatAsien = 1 "
       If Encasien.NoMatch = False Then
         If Encasien!TIPOASIEN < 1 Then
           IDEBE# = !IDEBEPASE
           IHABE# = !IHABEPASE
           If Int(CDbl(!FECHASIEN)) < FEINI Then
               TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
             ElseIf Int(CDbl(!FECHASIEN)) <= FEFIN Then
               Z$ = REGBLAN$("LIBROBAN")
               Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
               Call REPLA(Z$, !NUIDASIEN, 3, 12)
               Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
                 DEPASE$ = !REFEPASE
               Call REPLA(Z$, DEPASE$, 17, 32)
                 CHENUME$ = Space$(48)
               Call REPLA(Z$, CHENUME$, 49, 48)
               Call REPLA(Z$, MKD$(IDEBE#), 97, 8)
               Call REPLA(Z$, MKD$(IHABE#), 105, 8)
                 MACON$ = ""
                 On Error Resume Next
                 MACON$ = TRIM$(!MARCONSI)
                 On Error GoTo 0
               If MACON$ <> "" Then
                 Call REPLA(Z$, "Concilia", 121, 8)
               End If
               ' LO QUE SIGUE ES PARA SUPRIMIR LOS COMPROBANTES ANULADOS
               If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Then
                 JJ& = JJ& + 1
                 Call GRAREG("LIBROBAX", Z$, JJ&)
               End If
           End If
         End If
       End If
     .MoveNext
     Loop
   End With
   '
30 On Error GoTo 0
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CueContaI = " & CUEII% & " "
   'SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true)) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(FechConta) <= " & FEFIN & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !IDEBECOMP
         IHABE# = !IHABECOMP
       If Int(CDbl(!FechConta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
           Z$ = REGBLAN$("LIBROBAN")
           Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!FechAcred)), 15, 2)
             DEPASE$ = !RefeCom
           Call REPLA(Z$, DEPASE$, 17, 32)
             CHENUME$ = !DescriMov
           Call REPLA(Z$, CHENUME$, 49, 48)
           Call REPLA(Z$, MKD$(IDEBE#), 97, 8)
           Call REPLA(Z$, MKD$(IHABE#), 105, 8)
             MACON$ = ""
             On Error Resume Next
             MACON$ = TRIM$(!MARCONSI)
             On Error GoTo 0
           If MACON$ <> "" Then
             Call REPLA(Z$, "Concilia", 121, 8)
           End If
           ' LO QUE SIGUE ES PARA SUPRIMIR LOS COMPROBANTES ANULADOS
           If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Then
             JJ& = JJ& + 1
             Call GRAREG("LIBROBAX", Z$, JJ&)
           End If
       End If
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("LIBROBAX", JJ&)
   Call FILESORT("LIBROBAX", "", 1, 1, "ASC")
   '
55 JJ& = 0
   YATRAN% = 0
   TRANSPO# = TRANSPOR#: ITRAN# = TRANSPOR#
   '
   For U& = 1 To ULTREG&("LIBROBAX")
     XX$ = REGLEIDO$("LIBROBAX", U&)
     FEMO% = CVI(Left$(XX$, 2))
     IDEBE# = CVD(Mid$(XX$, 97, 8))
     IHABE# = CVD(Mid$(XX$, 105, 8))
     If FEMO% < DES% Then
         TRANSPO# = TRANSPO# + IDEBE# - IHABE#
         ITRAN# = TRANSPO#
         GoTo 59
       ElseIf FEMO% <= HAS% Then
         If YATRAN% < 1 Then
           ITRAN# = TRANSPO#
           If Abs(TRANSPO#) >= 0.005 Then
             YY$ = REGBLAN$("LIBROBAN")
               DESDE = DES%
             Call REPLA(YY$, MKI$(DIANTE(DESDE)), 1, 2)
             Call REPLA(YY$, "Transporte ........................................................................", 49, 48)
             Call REPLA(YY$, MKD$(TRANSPO#), 113, 8)
             JJ& = JJ& + 1
             Call GRAREG("LIBROBAN", YY$, JJ&)
             '
             YY$ = REGBLAN$("LIBROBAN")
             JJ& = JJ& + 1
             Call GRAREG("LIBROBAN", YY$, JJ&)
             YATRAN% = 1
           End If
           YATRAN% = 1
         End If
         '
         ITRAN# = ITRAN# + IDEBE# - IHABE#
         Call REPLA(XX$, MKD$(ITRAN#), 113, 8)
         JJ& = JJ& + 1
         Call GRAREG("LIBROBAN", XX$, JJ&)
     End If
59 Next U&
   '
   Call SETULTREG("LIBROBAN", JJ&)
   Call CIERRARCH("LIBROBAN")
   '
   Call HEADERS("LIBROBAN", "")
   Call HEADERS("LIBROBAN", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIBROBAN", "Libro de Bancos")
   '
End Sub

Sub MUEFACO()
   '
   NUDOCOR$ = TRIM$(Mid$(List1.TEXT, 10, 17)) 'NUMERO DE DOC COR
   FECHDOCU$ = TRIM$(Left$(List1.TEXT, 9))
   FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   '
   Call ABRECABAN
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CDBL(FechConta) = " & CDbl(FECDOC1) & " "
   SQLX$ = SQLX$ + " AND CodiCom_Cor= '" & NUDOCOR$ & "' "
   Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = TIPODOCU$ + " " + !CodiCom_Lar + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CueContaI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DescriMov
        IDEBE1# = !IDEBECOMP
        IHABE1# = !IHABECOMP
        If Abs(IDEBE1#) < 0.005 And Abs(IHABE1#) < 0.005 Then GoTo 19 ' ELIMINA REGISTROS EN BLANCO
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        SHOWFACO.List1.AddItem ADDLIS1$
19   .MoveNext
     Loop
   End With
   
   SHOWFACO.Show 1
   '
End Sub



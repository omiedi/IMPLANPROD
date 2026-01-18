VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form MAYORPAN07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes Contables"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   9975
      TabIndex        =   14
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
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   840
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame2"
      Height          =   6230
      Left            =   3130
      TabIndex        =   17
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
      TabIndex        =   16
      ToolTipText     =   "Conciliación de Cartera de Valores"
      Top             =   225
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame9"
      Height          =   6230
      Left            =   11520
      TabIndex        =   15
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6230
      Left            =   1050
      TabIndex        =   13
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6230
      Left            =   3430
      TabIndex        =   12
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   6230
      Left            =   10020
      TabIndex        =   8
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   6230
      Left            =   7020
      TabIndex        =   9
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   6230
      Left            =   8520
      TabIndex        =   7
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6230
      Left            =   3900
      TabIndex        =   6
      Top             =   1390
      Width           =   15
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6540
      Left            =   150
      TabIndex        =   0
      Top             =   1080
      Width           =   11670
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         ItemData        =   "MAYORPAN12.frx":0000
         Left            =   0
         List            =   "MAYORPAN12.frx":0002
         TabIndex        =   2
         Top             =   400
         Width           =   11655
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   11625
         TabIndex        =   3
         Top             =   0
         Width           =   11655
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"MAYORPAN12.frx":0004
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
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   11595
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MAYORPAN12.frx":00CB
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   600
      OleObjectBlob   =   "MAYORPAN12.frx":02FF
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   150
      ScaleHeight     =   435
      ScaleWidth      =   11625
      TabIndex        =   23
      Top             =   7610
      Width           =   11655
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6900
         TabIndex        =   25
         Top             =   60
         Width           =   1450
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8415
         TabIndex        =   24
         Top             =   60
         Width           =   1450
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   945
      TabIndex        =   18
      Top             =   630
      Width           =   5370
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
            Picture         =   "MAYORPAN12.frx":0379
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":0693
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":09AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":0CC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":0FE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":12FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":1615
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":192F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":1C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":1F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":227D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":2457
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":2569
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAYORPAN12.frx":2883
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
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   5475
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6615
      TabIndex        =   10
      Top             =   630
      Width           =   5175
   End
   Begin VB.Label Ltext2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6615
      TabIndex        =   5
      Top             =   225
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5230
      TabIndex        =   1
      Top             =   320
      Width           =   1275
   End
End
Attribute VB_Name = "MAYORPAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
Sub CONCILIAVIE()
   Screen.MousePointer = 11
10 For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To LIST1.ListCount
         TTXX2$ = LIST1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           If TXD$ = TXH$ Then
             GoSub RECUSALCHE
           End If
           Call REPLA(TTXX1$, Space$(14), 66, 14)
           LIST1.RemoveItem U& - 1
           LIST1.AddItem TTXX1$, U& - 1
           Call REPLA(TTXX2$, Space$(14), 80, 14)
           LIST1.RemoveItem U2& - 1
           LIST1.AddItem TTXX2$, U2& - 1
           GoTo 10
         End If
       Next U2&
     End If
19 Next U&
   '
20 TRANSPOR# = 0
   For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
     If TRIM$(Mid$(LIST1.List(U& - 1), 66, 28)) = "" Then
         LIST1.RemoveItem (U& - 1)
         GoTo 20
       Else
         TTXX1 = LIST1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         LIST1.RemoveItem U& - 1
         LIST1.AddItem TTXX1$, U& - 1
     End If
   Next U&
   Screen.MousePointer = 1
   LIST1.Refresh
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
   'Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   'Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
   PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   JJ& = 0
   TRANSPOR# = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = !idebepase
       IHABE# = !ihabepase
       Z$ = REGBLAN$("LIMAYOAC")
       Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
       Call REPLA(Z$, !nuidasien, 3, 12)
         MACON$ = ""
         On Error Resume Next
         MACON$ = TRIM$(!MARCONSI)
         On Error GoTo 0
       Call REPLA(Z$, MACON$, 15, 1)
       Call REPLA(Z$, "D", 16, 1)
         DEPASE$ = !RefePase
       Call REPLA(Z$, DEPASE$, 17, 32)
       Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
       Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         CENCO$ = SAFETEXT(!CenCos)
       Call REPLA(Z$, CENCO$, 65, 8)
       JJ& = JJ& + 1
       Call GRAREG("!LIMAYOAC", Z$, JJ&)
     .MoveNext
     Loop
   End With
   Set PASEASIEN = Nothing
   '
30 On Error GoTo 0
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   'SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true)) "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL ))"
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = !Idebecomp
       IHABE# = !IHABECOMP
       Z$ = REGBLAN$("LIMAYOAC")
       Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
       Call REPLA(Z$, !CodiCom_Cor, 3, 12)
         MACON$ = ""
         On Error Resume Next
         MACON$ = TRIM$(!MARCONSI)
         On Error GoTo 0
       Call REPLA(Z$, MACON$, 15, 1)
       Call REPLA(Z$, "C", 16, 1)
         DEPASE$ = !RefeCom
         DE2PASE$ = !DESCRIMOV
       Call REPLA(Z$, DEPASE$, 17, 32)
       Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
       Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         CENCO$ = SAFETEXT(!CenCosto)
       Call REPLA(Z$, CENCO$, 65, 8)
       Call REPLA(Z$, DE2PASE$, 73, 56)
       JJ& = JJ& + 1
       Call GRAREG("!LIMAYOAC", Z$, JJ&)
     .MoveNext
     Loop
   End With
   Set CABATEMP = Nothing
   '
40 Call SETULTREG("!LIMAYOAC", JJ&)
   'Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   '\\\OT-08-0609-OD-29/07/08///
   Call TRACENTRAL("LIMAYOAC", "")
   Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   Call TRALOCAL("LIMAYOAC", "")
   '\\\OT-08-0609-OD-FIN///
   '
   MAYORPAN07.LIST1.Clear
   Z$ = Space$(128 + 64)
     FEX = DIANTE(COMEJ%)
     FEXA$ = FECHATEX$(FEX)
   Call REPLA(Z$, FEXA$, 1, 8)
     DEPASE$ = "TRANSPORTE .............................."
   Call REPLA(Z$, DEPASE$, 35, 30)
   Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
   MAYORPAN07.LIST1.AddItem Z$
   MAYORPAN07.LIST1.AddItem Space$(95) + " =========== "
   '
   SUDEBE1# = 0: SUHABE1# = 0
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("!LIMAYOAC", U&)
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
     MAYORPAN07.LIST1.AddItem Z$
   Next U&
   Call CIERRARCH("!LIMAYOAC")
   '
   On Error GoTo 0
   If MAYORPAN07.LIST1.ListCount > 0 Then
     MAYORPAN07.LIST1.ListIndex = MAYORPAN07.LIST1.ListCount - 1
   End If
   MAYORPAN07.LIST1.ToolTipText = ""
   '
   MAYORPAN07.Label21 = TRIM$(FORMATNUM$(SUDEBE1#, "F11.2"))
   MAYORPAN07.Label22 = TRIM$(FORMATNUM$(SUHABE1#, "F11.2"))
   Screen.MousePointer = 1
   '
   Screen.MousePointer = 11
50 For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To LIST1.ListCount
         TTXX2$ = LIST1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           If TXD$ = TXH$ Then
             Call REPLA(TTXX1$, Space$(14), 66, 14)
             LIST1.RemoveItem U& - 1
             LIST1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, Space$(14), 80, 14)
             LIST1.RemoveItem U2& - 1
             LIST1.AddItem TTXX2$, U2& - 1
             GoTo 50
           End If
         End If
       Next U2&
     End If
19 Next U&
   '
55 For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
     IDEX$ = TRIM$(Mid$(TTXX1$, 66, 14))
     If Val(IDEX$) > 0 Then
       For U2& = 1 To LIST1.ListCount
         TTXX2$ = LIST1.List(U2& - 1)
         IHAX$ = TRIM$(Mid$(TTXX2$, 80, 14))
         If IDEX$ = IHAX$ Then
           TXD$ = REPLACAR$(Mid$(TTXX1$, 129, 64), " ", "")
           TXH$ = REPLACAR$(Mid$(TTXX2$, 129, 64), " ", "")
           TTXX$ = "Conciliación Imperfecta:\" + TRIM$(Mid$(TTXX1$, 129, 64)) + "\" + TRIM$(Mid$(TTXX2$, 129, 64))
           If COMALTER%(TTXX$ + "\\Conciliar\Continuar") = 1 Then
             Call REPLA(TTXX1$, Space$(14), 66, 14)
             LIST1.RemoveItem U& - 1
             LIST1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, Space$(14), 80, 14)
             LIST1.RemoveItem U2& - 1
             LIST1.AddItem TTXX2$, U2& - 1
             DoEvents
             LIST1.Refresh
             GoTo 59
           End If
         End If
       Next U2&
     End If
59 Next U&
   '
60 TRANSPOR# = 0
   For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
63   If TRIM$(Mid$(LIST1.List(U& - 1), 66, 28)) = "" Then
         On Error Resume Next
         LIST1.RemoveItem (U& - 1)
         If Err Then GoTo 89
         DoEvents
         LIST1.Refresh
         GoTo 63
       Else
         TTXX1 = LIST1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         LIST1.RemoveItem U& - 1
         LIST1.AddItem TTXX1$, U& - 1
     End If
   Next U&
89 On Error GoTo 0
   Screen.MousePointer = 1
   LIST1.Refresh
   DoEvents
   '
   Call GESVAL07.GENLICHECA(3)
   '
   For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
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
             LIST1.RemoveItem U& - 1
             LIST1.AddItem TTXX1$, U& - 1
             Call REPLA(TTXX2$, MKD$(0), 49, 8)
             Call GRAREG("!LICARCHE", TTXX2$, U2&)
             DoEvents
             LIST1.Refresh
             GoTo 91
           End If
         End If
       Next U2&
     End If
91 Next U&
   '
   TRANSPOR# = 0
   For U& = 1 To LIST1.ListCount
     TTXX1$ = LIST1.List(U& - 1)
92   If TRIM$(Mid$(LIST1.List(U& - 1), 66, 28)) = "" Then
         On Error Resume Next
         LIST1.RemoveItem (U& - 1)
         If Err Then GoTo 93
         DoEvents
         LIST1.Refresh
         GoTo 92
       Else
         TTXX1 = LIST1.List(U& - 1)
         IDEBE# = Val(Mid$(TTXX1$, 66, 14))
         IHABE# = Val(Mid$(TTXX1$, 80, 14))
         TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         Call REPLA(TTXX1$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
         LIST1.RemoveItem U& - 1
         LIST1.AddItem TTXX1$, U& - 1
         DoEvents
         LIST1.Refresh
     End If
   Next U&
93 On Error GoTo 0
   Screen.MousePointer = 1
   LIST1.Refresh
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
   If LIST1.ListCount > 0 Then
       Call COMUNI("Hay Movimientos en Cuenta Activa\No Conciliables con Cartera.")
     Else
       Call COMUNI("Fin de Control")
   End If
   '
End Sub

Private Sub Label2_Click()

End Sub





Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub

Private Sub GRID2_Click()
   NGRID = GRID2.Row
   NLIST = LIST1.ListCount
   If (NLIST) >= NGRID - 1 Then
     LIST1.ListIndex = NGRID - 1
     Call LIST1_Click
   End If
End Sub

Private Sub GRID2_DblClick()
   NGRID = GRID2.Row
   NLIST = LIST1.ListCount
   If (NLIST) >= NGRID - 1 Then
     LIST1.ListIndex = NGRID - 1
     Call List1_DblClick
   End If
End Sub

Private Sub GRID2_KeyDown(KeyCode As Integer, Shift As Integer)
   NGRID = GRID2.Row
   NLIST = LIST1.ListCount
   If (NLIST) >= NGRID - 1 Then
     LIST1.ListIndex = NGRID - 1
     Call List1_KeyDown(KeyCode, Shift)
   End If
   '
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
     If ARBUS$ <> "" Then
        VUELTA% = 0
        j& = GRID2.Row
    
        h& = 0: If GRID2.COL < GRID2.Cols Then h& = GRID2.COL + 1
10            For U& = j& To GRID2.Rows - 1
           HI& = h&: If U& > j& Then HI& = 0
           For i& = HI& To GRID2.Cols - 1
            If InStr(UCase$(GRID2.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
              GRID2.Row = U&: GRID2.COL = i&
              On Error Resume Next
               VISIROWS = (GRID2.Height / GRID2.RowHeight(j&)) - 4
              If Err > 0 Then
               VISIROWS = 1
              End If
              On Error GoTo 0
               If VISIROWS > 0 Then
                  If GRID2.TopRow < GRID2.Row - VISIROWS Then
                    GRID2.TopRow = GRID2.Row - VISIROWS
                  End If
               End If
              GRID2.SetFocus
              Exit Sub
            End If
            Next i&
        Next U&
        '
        Call MENSERR(24, "Argumento no Encontrado.")
     End If
  End If

End Sub

Private Sub LIST1_Click()
   LIST1.ToolTipText = TRIM$(Mid$(LIST1.TEXT, 129, 56))
End Sub

Private Sub List1_DblClick()
   ORIMO$ = TRIM$(Mid$(LIST1.TEXT, 128, 1))
   If ORIMO$ = "C" Then
     Call MUEFACO
   End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  '\\\OT-07-0046-RG-01/02/07///
   If KeyCode = 114 Then
       If ARBUS$ <> "" Then
          VUELTA% = 0
          INI& = LIST1.ListIndex + 2
    '
10        For U& = INI& To LIST1.ListCount
              If InStr(UCase$(LIST1.List(U& - 1)), ARBUS$) > 0 Then
                 LIST1.ListIndex = U& - 1
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
   
'   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
'   If ABUSQUE$ <> "" Then
'     ARBUS$ = ABUSQUE$
'     For U& = 1 To GRID2.Rows - 1
'      For I& = 0 To GRID2.Cols - 1
'       If InStr(UCase$(GRID2.TextMatrix(U&, I&)), ABUSQUE$) > 0 Then
'        GRID2.Row = U&: GRID2.Col = I&
'        If GRID2.TopRow > GRID2.Row Then GRID2.TopRow = GRID2.Row
'        On Error Resume Next
'        VISIROWS = (GRID2.Height / GRID2.RowHeight(J&)) - 4
'        If Err > 0 Then
'         VISIROWS = 1
'        End If
'        On Error GoTo 0
'         If VISIROWS > 0 Then
'            If GRID2.TopRow < GRID2.Row - VISIROWS Then
'              GRID2.TopRow = GRID2.Row - VISIROWS
'            End If
'         End If
'        GRID2.SetFocus
'        Exit Sub
'       End If
'       Next I&
'     Next U&
'     '
'     Call MENSERR(24, "Argumento no Encontrado.")
'   End If

  '\\\OT-07-0046-RG-01/02/07///
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     LIST1.ListIndex = (-1)
     For U& = 1 To LIST1.ListCount
       If InStr(UCase$(LIST1.List(U& - 1)), ABUSQUE$) > 0 Then
         LIST1.ListIndex = U& - 1
         LIST1.SetFocus
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
   If Label4 = "Selección Multiple" Then
      Call LIMAYORPRIN_SELMUL
      Exit Sub
   End If

10 VDEF$ = MKI$(COMEJ%) + MKI$(FINEJ%)
   RANFE$ = OBJPLA$("DESHASFECHA", VDEF$)
   If RANFE$ = "" Then Exit Sub
   '
   Des% = CVI(Left$(RANFE$, 2))
   Has% = CVI(Mid$(RANFE$, 3, 2))
   If Has% < Des% Then
     Call MENSERR(24, "Rango de Fechas no procesable.")
     GoTo 10
   End If
   CORRESP$ = FECHATEX$(Des%) + " al " + FECHATEX$(Has%) + " - " + TRIM$(Label4) + " (" + TRIM$(Ltext2) + ")"
   '
   If ECOARCH$("CUEBANC", TRIM$(Ltext2)) <> "" Then
     OPXX% = COMALTER%("La Cuenta Elegida Figura en\la Tabla de Cuentas Bancarias.\Elija el Formato de Impresión:\\Mayor por Cuenta\Libro de Bancos")
     If OPXX% = 1 Then
         GoTo 15
       ElseIf OPXX% = 2 Then
         Call LIBROBAN(Des%, Has%, CORRESP$)
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
   For U& = 1 To LIST1.ListCount
     TTXX$ = LIST1.List(U& - 1)
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
     If FEMO% < Des% Then
         TRANSPO# = TRANSPO# + IDEBE# - IHABE#
         GoTo 29
       ElseIf FEMO% <= Has% Then
         If YATRAN% < 1 Then
           If Abs(TRANSPO#) >= 0.005 Then
             XX$ = REGBLAN$("LIMAYOW")
               Desde = Des%
             Call REPLA(XX$, FECHATEX$(DIANTE(Desde)), 1, 8)
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
             'COCOM$ = Mid$(TTXX$, 10, 16)
             COCOM$ = Mid$(TTXX$, 10, 18): TIPOCOM$ = Mid$(COCOM$, 1, 5)
             If InStr(1, "FCNCND", Mid$(TIPOCOM$, 1, 2)) > 0 Then
                PP$ = TRIM$(REPLACAR$(Mid$(COCOM$, 1, 5), "-", ""))
                COCOM$ = PP$ & TRIM$(Mid$(COCOM$, 6))
             End If
           Call REPLA(XX$, COCOM$, 9, 16)
             CENCO$ = TRIM$(Mid$(TTXX$, 193, 8))
           Call REPLA(XX$, CENCO$, 25, 8)
             'REFE$ = Mid$(TTXX$, 35, 30)
             REFE$ = Mid$(TTXX$, 37, 28)
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

Sub LIMAYORPRIN_SELMUL()

   '
   Des% = FECHANUM(INIMAY07.Label5): If FEX1% < COMEJ% Then FEX1% = COMEJ%
   Has% = FECHANUM(INIMAY07.Label6): If FEX2% > FINEJ% Then FEX2% = FINEJ%

   CORRESP$ = FECHATEX$(Des%) + " al " + FECHATEX$(Has%) + " - " + TRIM$(Label4) + " (" + TRIM$(Ltext2) + ")"
   SALTOLINEA% = 0

   For U& = 1 To LIST1.ListCount
     TTXX$ = LIST1.List(U& - 1)
     REFE$ = Mid$(TTXX$, 37, 28)
     
     If SALTOLINEA% = 1 Then
        SALTOLINEA% = 0
        JJ& = JJ& + 1
        Call GRAREG("LIMAYOW", String$(256, "="), JJ&)
        GoTo 29
     End If
     If InStr(1, UCase(REFE$), "TOTAL") > 0 And (U& + 2) < LIST1.ListCount Then
        SALTOLINEA% = 1
     End If
     
     FEMO% = FECHANUM(Left$(TTXX$, 8))
     COCOM$ = Mid$(TTXX$, 10, 18): TIPOCOM$ = Mid$(COCOM$, 1, 5)
     IDEBE# = Val(Mid$(TTXX$, 66, 14))
     IHABE# = Val(Mid$(TTXX$, 80, 14))
     If InStr(1, "FCNCND", Mid$(TIPOCOM$, 1, 2)) > 0 Then
       PP$ = TRIM$(REPLACAR$(Mid$(COCOM$, 1, 5), "-", ""))
       COCOM$ = PP$ & TRIM$(Mid$(COCOM$, 6))
     End If
     Call REPLA(XX$, Left$(TTXX$, 8), 1, 8)
     Call REPLA(XX$, COCOM$, 9, 16)
      CENCO$ = TRIM$(Mid$(TTXX$, 193, 8))
     Call REPLA(XX$, CENCO$, 25, 8)
      REFE$ = Mid$(TTXX$, 37, 28)
     Call REPLA(XX$, REFE$, 37, 48)
     Call REPLA(XX$, MKD$(IDEBE#), 85, 8)
     Call REPLA(XX$, MKD$(IHABE#), 93, 8)
      ITRAN# = Val(Mid$(TTXX$, 94, 14))
     Call REPLA(XX$, MKD$(ITRAN#), 101, 8)
     JJ& = JJ& + 1
     Call GRAREG("LIMAYOW", XX$, JJ&)
     'Debug.Print XX$
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


Sub LIBROBAN(Des%, Has%, CORRESP$)
   '
   FEINI = Int(CDbl(CVDAT(COMEJ%)))
   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   FEINI1 = FETEXCOR1$(COMEJ%)
   FEFIN1 = FETEXCOR1$(FINEJ%)
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   CUEII% = CUEINT%(Ltext2)
   NICONFLIM% = NIVCONFI%
   FEX1% = Des%: If FEX1% < COMEJ% Then FEX1% = COMEJ%
   FEX2% = Has%: If FEX2% > FINEJ% Then FEX2% = FINEJ%
   FEINI = FETEXCOR1$(FEX1%)
   FEFIN = FETEXCOR1$(FEX2%)
   CONSOLI% = 0
   ' SUCUI% = UNINEGO%
   ' If INIMAY07.Check1.Value = 1 Then CONSOLI% = 1
   '
   Call ABRECABAN
   Dim PASEASIEN As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
   '
   TRANSPOR# = 0
   SQLX$ = "SELECT sum(idebepase) AS TRADEBE, SUM(ihabepase) AS TRAHABE FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien < '" & FEINI & "'"
   SQLX$ = SQLX$ + " AND statpase > 0 ) "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
   Set PASEASIEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   TRANSPOR# = XVALO(PASEASIEN!TRADEBE) - XVALO(PASEASIEN!TRAHABE)
   PASEASIEN.Close
   '
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   '
   SQLX$ = "SELECT sum(IDEBECOMP) AS TRADEBE, SUM(IHABECOMP) AS TRAHABE FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE (CueContaI = " & CUEII%
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL)) "
   SQLX$ = SQLX$ + " AND FechConta < '" & FEINI & "' "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND CODIEMPR = " & SUCUI% & " "
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   TRANSPOR# = TRANSPOR# + XVALO(CABATEMP!TRADEBE) - XVALO(CABATEMP!TRAHABE)
   CABATEMP.Close
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND FechAsien >= '" & FEINI & "' "
   SQLX$ = SQLX$ + " AND Fechasien <= '" & FEFIN & "'"
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   '
   Set PASEASIEN = New ADODB.Recordset
   PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   Call COPYSTRU("LIBROBAN", "LIBROBAX")
   '
   JJ& = 0
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
        NUIDAS$ = !nuidasien
        TIASIX% = XVALO(VALOBADA("EncabeAsien", "TIPOASIEN", "NuIdAsien = '" & NUIDAS$ & "' AND StatAsien = 1 "))
        If TIASIX% < 1 Then
           IDEBE# = !idebepase
           IHABE# = !ihabepase
           Z$ = REGBLAN$("LIBROBAN")
           Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
           Call REPLA(Z$, !nuidasien, 3, 12)
           Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 15, 2)
             DEPASE$ = !RefePase
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
     .MoveNext
     Loop
   End With
   PASEASIEN.Close
   Set PASEASIEN = Nothing
   '
30 On Error GoTo 0
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND FechConta >= '" & FEINI & "' "
   SQLX$ = SQLX$ + " AND FechConta <= '" & FEFIN & "' "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND CODIEMPR = " & SUCUI% & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        GoTo 40
     End If
     On Error GoTo 0
     Do While Not .EOF
         IDEBE# = !Idebecomp
         IHABE# = !IHABECOMP
         Z$ = REGBLAN$("LIBROBAN")
         Call REPLA(Z$, MKI$(CVINT(!fechconta)), 1, 2)
         Call REPLA(Z$, !CodiCom_Cor, 3, 12)
         Call REPLA(Z$, MKI$(CVINT(!fechacred)), 15, 2)
           DEPASE$ = !RefeCom
         Call REPLA(Z$, DEPASE$, 17, 32)
           CHENUME$ = !DESCRIMOV
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
     .MoveNext
     Loop
   End With
   CABATEMP.Close
   Set CABATEMP = Nothing
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
       If YATRAN% < 1 Then
         ITRAN# = TRANSPO#
         If Abs(TRANSPO#) >= 0.005 Then
           YY$ = REGBLAN$("LIBROBAN")
             Desde = Des%
           Call REPLA(YY$, MKI$(DIANTE(Desde)), 1, 2)
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
       IDEBE# = XVALO(FORMATNUM$(IDEBE#, "F12.2"))
       IHABE# = XVALO(FORMATNUM$(IHABE#, "F12.2"))
       '
       ITRAN# = ITRAN# + IDEBE# - IHABE#
       Call REPLA(XX$, MKD$(ITRAN#), 113, 8)
       JJ& = JJ& + 1
       Call GRAREG("LIBROBAN", XX$, JJ&)
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
   NUDOCOR$ = TRIM$(Mid$(LIST1.TEXT, 10, 18)) 'NUMERO DE DOC COR
   FECHDOCU$ = TRIM$(Left$(LIST1.TEXT, 9))
   'FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   FECDOC1 = FETEXCOR1(FECHANUM(TRIM$(Mid$(LIST1.TEXT, 1, 8)))) 'FECHA DEL DOC.
   DEPAS$ = TRIM$(Mid$(LIST1.TEXT, 35, 30))
   CONDI$ = "FechConta = '" & FECDOC1 & "' "
   CONDI$ = CONDI$ + " AND (CodiCom_Lar= '" & NUDOCOR$ & "' OR CodiCom_Cor='" & NUDOCOR$ & "') "
   'CONDI$ = CONDI$ + " AND CodiCom_Cor= '" & NUDOCOR$ & "' "
   'se cambia criterio de consulta reemplazando = por like por que en el list trunca a 30 carcteres
   'y en la base de datos puede tener mas caracteres.
   'CONDI$ = CONDI$ + " AND RefeCom = '" & DEPAS$ & "'"
   CONDI$ = CONDI$ + " AND RefeCom LIKE '" & DEPAS$ & "%'"
   NUMCLI& = XVALO(VALOBADA("CAJABANC", "NuClien", CONDI$, "NOMESS"))
   NUMPRO& = XVALO(VALOBADA("CAJABANC", "NuProve", CONDI$, "NOMESS"))
   '
   Call ABRECABAN
   Dim SHOFACOTMP As ADODB.Recordset
   SQLX$ = " SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + " WHERE CDBL(FechConta) = " & CDbl(FECDOC1) & " "
   SQLX$ = SQLX$ + " WHERE FechConta = '" & FECDOC1 & "' "
   'SQLX$ = SQLX$ + " AND CodiCom_Cor= '" & NUDOCOR$ & "' "
   SQLX$ = SQLX$ + " AND (CodiCom_Lar= '" & NUDOCOR$ & "' OR CodiCom_Cor='" & NUDOCOR$ & "') "
   If NUMCLI& > 0 Then SQLX$ = SQLX$ + " AND NuClien = " & NUMCLI& & " "
   If NUMPRO& > 0 Then SQLX$ = SQLX$ + " AND Nuprove = " & NUMPRO& & " "
   '
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
   Set SHOFACOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With SHOFACOTMP
     'On Error Resume Next
     '.MoveFirst
     'If Err Then
     If (.EOF And .BOF) Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        'On Error GoTo 0
        Exit Sub
     End If
     'On Error GoTo 0
     .MoveFirst
     SHOWFACO.Caption = TIPODOCU$ + " " + !codicom_lar + " del " + FECHDOCU$
     SHOWFACO.LIST1.Clear
     NP2& = XVALO(!NUPROVE)
     NC2& = XVALO(!NUCLIEN)
     FEXI% = CVINT(!FECHEMISI)
     FECDOC2 = FETEXCOR1$(FEXI%)
     NUDOCOR2$ = SAFETEXT$(!codicom_lar)
     SQLZ$ = " WHERE FechEmisi = '" & FECDOC2 & "' "
     SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR2$ & "' "
     SQLZ$ = SQLZ$ + " AND NuClien =  " & NC2& & " "
     SQLZ$ = SQLZ$ + " AND NuProve =  " & NP2& & " "
     '
     ' boton de recuperacion
     If Left$(!codicom_lar, 2) = "OP" Then
        SHOWFACO.Label16 = TRIM$(Mid$(!codicom_lar, 13, 6))
        SHOWFACO.Label18 = TRIM$(!codicom_lar)
        SHOWFACO.Label17 = FECDOC1
        'SHOWFACO.Command1.Visible = True
     End If
     '
     SHOWFACO.Label19 = ""
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     Do While .EOF = False
        If XVALO(!NUCLIEN) <> NUMCLI& Then GoTo 19
        If XVALO(!NUPROVE) <> NUMPRO& Then GoTo 19
        '
        QQQ1$ = CODCUE$(!CUECONTAI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = SAFETEXT$(!DESCRIMOV)
        REFE2$ = TRIM$(SAFETEXT$(!OBSERGEN))
        If InStr(REFE1$, NUDOCOR$) > 0 Then
           If REFE2$ <> "" Then
              REFE1$ = REFE2$
           End If
        End If
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        If Abs(IDEBE1#) < 0.005 And Abs(IHABE1#) < 0.005 Then GoTo 19 ' ELIMINA REGISTROS EN BLANCO
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.LIST1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        TMVX$ = UCase$(TRIM$(!TipoMovim))
        If TMVX$ = "FCOM" Or TMVX$ = "OPAG" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = Abs((!IHABECOMP - !Idebecomp)) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - Abs((!IHABECOMP - !Idebecomp))
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
        End If
        SHOWFACO.Label14 = !OBSERGEN
        '\\\OT-08-988-29/10/08-FL
        USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
        SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
        '\\\OT-08-988-FIN-FL
        '
19   .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub


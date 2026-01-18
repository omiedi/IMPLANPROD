VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DAPORCU07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra Pendientes - Marcar Items Cumplidos"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11745
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   23
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LICOPEN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Items que Dará por Cumplidos"
      Top             =   2040
      Width           =   10545
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   10515
      TabIndex        =   6
      Top             =   1630
      Width           =   10545
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   32
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Recep."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "U/S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   29
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   10800
      ScaleHeight     =   5745
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   645
         TabIndex        =   25
         Top             =   3600
         Width           =   675
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         Picture         =   "DAPORCU07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   4030
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         Picture         =   "DAPORCU07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   4755
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         Picture         =   "DAPORCU07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   240
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "DAPORCU07.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "DAPORCU07.frx":0992
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   22
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   9560
      TabIndex        =   21
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   9680
      TabIndex        =   20
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de O/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7800
      TabIndex        =   19
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Emisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7680
      TabIndex        =   18
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      Height          =   285
      Left            =   5760
      TabIndex        =   17
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   5760
      TabIndex        =   16
      Top             =   450
      Width           =   1695
   End
   Begin VB.Label Label9 
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
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   1110
      Width           =   4095
   End
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   780
      Width           =   3255
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label6 
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
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   450
      Width           =   4095
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
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1110
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   220
      Left            =   120
      TabIndex        =   9
      Top             =   855
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Posición IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   260
      Left            =   5790
      TabIndex        =   8
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5790
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7050
      Y1              =   4340
      Y2              =   4340
   End
End
Attribute VB_Name = "DAPORCU07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   If COMALTER%("Acepta Dar por Cumplidos\los Items Tildados ?\\No, Cancelar\Si, Dar por Cumplidos") = 2 Then
     Hide
   End If
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Label14_Change()
'JANDY SE PODRIA HACER UN JOIN ENTRE OCOMENCA Y OCOMDETA WHERE OCOMDETA<9 AND NUME_OCOM=NUOCO&
   
   '
   NUOCO& = XVALO(Label14)
   If NUOCO& > 0 Then
     Call APERBASDAT("COMPRAS")
     SQLX$ = " SELECT Npro_Ocom, Femi_Ocom, AtSr_Ocom "
     SQLX$ = SQLX$ + " FROM OCOMENCA "
     SQLX$ = SQLX$ + " WHERE Nume_Ocom =   " & NUOCO& & "     "
     'Set COMPLETEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim COMPLETEMP As ADODB.Recordset
     Set COMPLETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With COMPLETEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "Imposible Presentar O/C " + Str$(NUOCO&))
         Exit Sub
       End If
       '
       NPRO% = XVALO(!Npro_Ocom)
       Label15 = FECHATEX$(CVINT(!FEMI_OCOM))
       Label9 = SAFETEXT$(!AtSr_Ocom)
     End With
     COMPLETEMP.Close
     Set COMPLETEMP = Nothing
     '
     Label3 = NPRO%
     NPRO1% = (-1) * NPRO%
     '
     Label4 = rasocli(NPRO1%)
     Label6 = DOMICLI(NPRO1%)
     Label7 = CPOSCLI(NPRO1%)
     Label8 = LOCACLI(NPRO1%)
     Label10 = PIVACLI(NPRO1%)
     Label11 = CUITCLI(NPRO1%)
     '
     SQLX$ = "SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + " WHERE  Nume_Ocom =   " & NUOCO& & " "
     SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
     'Set ITETEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim ITETEMP As ADODB.Recordset
     Set ITETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     With ITETEMP
       On Error Resume Next
       .MoveFirst
       LICOPEN.Clear
       Do While (.EOF = False)
         CII% = XVALO(!Cod_Inte)
         DEMAT$ = DESCRIT0$(CII%)
           If CII% <= 0 Then              ' PARA CONSIDERAR TEXTO LIBRE
             DEMAT$ = TRIM$(Left$(SAFETEXT$(!DELIBRE), 64))
           End If
'         COEFCAN = !Coef_Unids
'           If COEFCAN < 0.005 Then COEFCAN = 1
         CAORI = XVALO(!Cant_Ocom) '* COEFCAN se comenta para que tome el valor de cant_ocom.
         'Se quita la variable COEFCAN porque la cantidad recibida no esta expresada
         'segun la unidad de compra, que es para lo que se utiliza COEFCAN
         'CAREC = !Cant_Rec * COEFCAN
         CAREC = XVALO(!Cant_Rec)
         CAPEN = CAORI - CAREC: If CAPEN < 0 Then CAPEN = 0
         UNICO$ = SAFETEXT$(!Uni_Mas)
         COEFUNI = XVALO(!Coef_Unids)
         STAT% = XVALO(!Stat_Ocom)
         NUORIT% = XVALO(!NuOrd_Item)
         '
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
         Call REPLA(TTXX$, DEMAT$, 18, 30)
         Call REPLA(TTXX$, UNICO$, 49, 4)
         Call REPLA(TTXX$, FORMATNUM$(CAORI, "F10.1"), 54, 10)
         Call REPLA(TTXX$, FORMATNUM$(CAREC, "F10.1"), 64, 10)
         Call REPLA(TTXX$, FORMATNUM$(CAPEN, "F10.1"), 74, 10)
         Call REPLA(TTXX$, Str$(NUORIT%), 121, 8)  ' numero de orden del item
         '
         If STAT% = 8 Then
           Call REPLA(TTXX$, "Cumplido", 76, 10)
         End If
         LICOPEN.AddItem TTXX$
       .MoveNext
       Loop
     End With
     ITETEMP.Close
     Set ITETEMP = Nothing
     '
   End If
   '
End Sub


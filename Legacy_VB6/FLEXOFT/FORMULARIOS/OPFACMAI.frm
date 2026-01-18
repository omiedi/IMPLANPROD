VERSION 5.00
Begin VB.Form OPFACMAI 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Facturación"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6660
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
   ScaleHeight     =   5145
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5600
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   4095
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   5160
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Acopio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3570
         TabIndex        =   15
         Top             =   1050
         Width           =   1065
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Anulación de Comprobantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   11
         Top             =   4200
         Width           =   3900
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Devoluciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2940
         TabIndex        =   10
         Top             =   3150
         Width           =   1800
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   8
         Top             =   3150
         Width           =   1800
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato de Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   6
         Top             =   2100
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato Itemizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   4
         Top             =   1050
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Anulaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   12
         Top             =   3675
         Width           =   2850
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   9
         Top             =   2625
         Width           =   2850
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas de Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   7
         Top             =   1575
         Width           =   2850
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Facturación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   420
         TabIndex        =   5
         Top             =   525
         Width           =   2850
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   5600
      Picture         =   "OPFACMAI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
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
      Left            =   5600
      Picture         =   "OPFACMAI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   855
      Left            =   5600
      Picture         =   "OPFACMAI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   335
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5355
      Picture         =   "OPFACMAI.frx":075E
      Stretch         =   -1  'True
      Top             =   4425
      Width           =   5250
   End
End
Attribute VB_Name = "OPFACMAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   Hide
   If Option1.Value = True Then       ' Facturación Itemizada
       FORFACMAI.Show
     ElseIf Option2.Value = True Then ' Notas de Débito de Texto Libre
        FORFACMAI.Hide
        FORDELIB.Show
     ElseIf Option3.Value = True Then ' Notas de Crédito de Texto Libre
        FORFACMAI.Hide
        FORCRELIB.Show
     ElseIf Option4.Value = True Then ' Notas de Crédito por Devoluciones
        FORFACMAI.Hide
        FORCREDEV.Show
     ElseIf Option5.Value = True Then ' Anulacion de Comprobantes
       Call ANUFAC
     ElseIf Option6.Value = True Then ' Facturas de Acopio
        FORFACMAI.Hide
        FORFACOPIO.Show
     Else
       FORFACMAI.Show
       Option1.Value = True
       Call MENSERR(24, "La Opción Elegida no se Encuentra Disponible por el Momento.")
   End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("FACTUMAI")
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call VERIMOVPER
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Option1_dblClick() ' Facturacion Itemizada
   Hide
   FORFACMAI.Show
End Sub

Private Sub Option2_dblClick() ' Notas de Débito de Texto Libre
   Hide
   FORFACMAI.Hide
   FORDELIB.Show
End Sub

Private Sub Option3_DBLClick() ' Notas de Crédito de Texto Libre
   Hide
   FORFACMAI.Hide
   FORCRELIB.Show
End Sub

Private Sub Option4_DBLClick() ' Notas de Crédito por Devoluciones
   Hide
   FORFACMAI.Hide
   FORCREDEV.Show
End Sub

Private Sub Option5_DblClick() ' Anulacion de Comprobantes
   Hide
   Call ANUFAC
End Sub

Private Sub Option6_DblClick()    ' Factura de Acopio
   Hide
   FORFACMAI.Hide
   FORFACOPIO.Show
End Sub
'
Sub VERIMOVPER()
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("BOLREPEN")
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("MOVISTO")
   INI& = FIN& - 4096: If INI& < 1 Then INI& = 1
   '
   For U& = INI& To FIN&
     X1$ = REGLEIDO$("MOVISTO", U&)
     If TRIM$(Mid$(X1$, 21, 52)) = "" Then GoTo 10
   Next U&
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("BOLREPEN")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
   '
10 Dim IDENBO$(32767)
   CABOL& = 0
   RECUPE% = 0
   CONAJU$ = ""
   PUNHUECO$ = ""
   PUNCODIMO$ = String$(32767, "0")
   '
   FIN& = ULTREG&("MOVISTO")
   UREBK& = ULTREG&("BKMOVIST")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     If U& <= UREBK& Then
       X1$ = REGLEIDO$("MOVISTO", U&)
       X2$ = REGLEIDO$("BKMOVIST", U&)
       If Mid$(X1$, 21, 52) <> Mid$(X2$, 21, 52) Then
         If Mid$(X2$, 21, 52) <> String$(52, "-") Then
           Call GRAREG("MOVISTO", X2$, U&)
           RECUPE% = 1
           CI% = CVI(Mid$(X2$, 3, 2)): If CI% > 0 Then Mid$(PUNCODIMO$, CI%) = "1"
         End If
       End If
     End If
     '
     X1$ = REGLEIDO$("MOVISTO", U&)
     If UCase$(Mid$(X1$, 21, 2)) = "AJ" Then
        NUAJU& = Val(Mid$(X1$, 26, 7))
        If CONAJU$ = "" Then
           CONAJU$ = String$(NUAJU&, "*")
        End If
        While Len(CONAJU$) < NUAJU&
           CONAJU$ = CONAJU$ + " "
        Wend
        Mid$(CONAJU$, NUAJU&, 1) = "*"
     End If
     '
     If UCase$(Mid$(X1$, 21, 2)) = "BR" Then
        CABOL& = CABOL& + 1
        IDENBO$(CABOL&) = TRIM$(Mid$(X1$, 5, 16))
     End If
     '
     If TRIM$(Mid$(X1$, 21, 52)) = "" Then
        PUNHUECO$ = PUNHUECO$ + MKS$(U&)
     End If
     '
   Next U&
   '
   CABOPE& = 0
   For U& = 1 To ULTREG&("BOLRECEP")
      X$ = REGLEIDO$("BOLRECEP", U&)
      IDENBOL$ = TRIM$(Mid$(X$, 69, 12))
      For CA& = 1 To CABOL&
        If IDENBO$(CA&) = IDENBOL$ Then GoTo 59
      Next CA&
      CABOPE& = CABOPE& + 1
      Call REPLA(X$, String$(24, 0), 163, 8)
      Call REPLA(X$, MKS$(U&), 189, 4)  ' puntero a BOLRECEP
      Call GRAREG("BOLREPEN", X$, CABOPE&)
      RECUPE% = 1
      CI% = CVI(Mid$(X$, 83, 2)): If CI% > 0 Then Mid$(PUNCODIMO$, CI%) = "1"
59 Next U&
   Call SETULTREG("BOLREPEN", CABOPE&)
   Call CIERRARCH("BOLREPEN")
   Call CIERRARCH("BOLRECEP")
   '
   For U& = 1 To ULTREG&("AJUSTOCK")
      X$ = REGLEIDO$("AJUSTOCK", U&)
      NUAJU& = CVS(Mid$(X$, 3, 4))
      If Mid$(CONAJU$, NUAJU&, 1) = " " Then
         GoSub RECUAJU
      End If
   Next U&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("BOLREPEN")
   Call CIERRARCH("*.*")
   If RECUPE% > 0 Then
      GoSub GENLISRECU
   End If
   Screen.MousePointer = 1
   Exit Sub
   '
RECUAJU:
   FECHA% = CVI(Left$(X$, 2))
   CI% = CVI(Mid$(X$, 7, 2))
   LOTE$ = Mid$(X$, 9, 16)
   CMOV$ = "AJ"
   DOPRIX$ = "AS." + TRIM$(Str$(NUAJU&))
   DOSECX$ = "AJ."
   REFE$ = Mid$(X$, 34, 47)
   VUNI# = 0
   DEPO% = 0
   NCLIE% = 0
   NPROV% = 0
   ENTRADA# = CVD(Mid$(X$, 25, 8)): SALIDA# = 0
   If ENTRADA# < 0 Then
      SALIDA# = Abs(ENTRADA#)
      ENTRADA# = 0
   End If
   '
   RECORD$ = REGBLAN$("MOVISTO")
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(DEPO%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(ENTRADA#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(USER%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   If Len(PUNHUECO$) >= 4 Then
       REGMO& = CVS(Left$(PUNHUECO$, 4))
       PUNHUECO$ = Mid$(PUNHUECO$, 5)
       Call GRAREG("MOVISTO", RECORD$, REGMO&)
     Else
       Call REGAPP("MOVISTO", RECORD$)
   End If
   RECUPE% = 1
   If CI% > 0 Then Mid$(PUNCODIMO$, CI%) = "1"
   '
Return
   '
GENLISRECU:
   '
   JJ& = 0
   FIN& = NUINV%
   For U& = 1 To NUINV%
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("INVERT", U& + 1)
      CI% = CVI(Mid$(X$, 17, 2))
      If CI% > 0 Then
         If Mid$(PUNCODIMO$, CI%, 1) <> "0" Then
            JJ& = JJ& + 1
            Call GRAREG("LISRECU", MKI$(CI%), JJ&)
         End If
      End If
   Next U&
   Call SETULTREG("LISRECU", JJ&)
   Call CIERRARCH("LISRECU")
   '
   If JJ& > 0 Then
      HOII% = HOY(HOS$)
      Call HEADERS("LISRECU", "")
      Call HEADERS("LISRECU", "Generado el: " + HOS$ + " - " + Time$ + " hs.")
      Call MENSERR(24, "Se Han Detectado Diferencias Recuperadas\en la Situación General de Stocks.\   \Pulse 'Aceptar' para Emitir el Listado\de los Códigos que Deben Revisarse.")
      Call CONECRUN("*LISRECU", "Listado de Códigos a Revisar")
   End If
   '
Return
   '
End Sub

VERSION 5.00
Begin VB.Form SENDCOMIS 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar Comisiones por Mail"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selección"
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
      Left            =   315
      TabIndex        =   10
      Top             =   2415
      Width           =   4635
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle Deuda"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   420
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comisiones"
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
         Left            =   630
         TabIndex        =   11
         Top             =   420
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   315
      ScaleHeight     =   75
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   735
      Width           =   1170
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12000
      End
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
      Height          =   330
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   4200
      Picture         =   "SENDCOMIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Genera Resumen y Envía por Mail"
      Top             =   630
      Width           =   750
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
      Height          =   330
      Left            =   3820
      TabIndex        =   0
      Top             =   630
      Width           =   175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   1605
      Width           =   2220
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   315
      TabIndex        =   5
      Top             =   1890
      Width           =   4635
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   315
      TabIndex        =   4
      Top             =   1155
      Width           =   3690
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3285
      TabIndex        =   3
      Top             =   630
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
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
      Index           =   10
      Left            =   945
      TabIndex        =   1
      Top             =   735
      Width           =   2220
   End
End
Attribute VB_Name = "SENDCOMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click()
   '
   BOTREC.Enabled = False
   NVENDE% = Val(Label1)
   If NVENDE% < 1 Or NVENDE% > 255 Then GoTo 99
   If TRIM$(ECOARCH$("VENDEDOR", Chr$(NVENDE%))) = "" Then
     Call MENSERR(24, "Vendedor Inexistente")
     GoTo 99
   End If
   '
   DIRMAI$ = TRIM$(Label4)
   If DIRMAI$ = "" Or InStr(DIRMAI$, Chr$(64)) < 2 Then
     Call MENSERR(24, "Direccion de Mail no Válida")
     GoTo 99
   End If
   '
   If Check1.Value < 1 Then
     If Check2.Value < 1 Then
       Call MENSERR(24, "No se Han Seleccionado Datos")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   If Check1.Value = 1 Then
      JJ& = 0
      N1% = FILEOPEN%("C:\FLEXOFT\COMISION.TXT", 2, 128)
      FIN& = ULTREG&("LICOMIS")
      For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("LICOMIS", IL&)
        If Asc(Left$(X$, 1)) = NVENDE% Then
          TTXX$ = Space$(80)
          NC% = CVI(Mid$(X$, 2, 2))
          NRECIB& = CVS(Mid$(X$, 7, 4))
          FEX = CVI(Mid$(X$, 4, 2))
          IMPO# = CVD(Mid$(X$, 21, 8))
          ICOM# = CVD(Mid$(X$, 33, 8))
          '
          Call REPLA(TTXX$, RASOCLI$(NC%), 1, 24)
          Call REPLA(TTXX$, FORMATNUM$(NRECIB&, "F8.0"), 25, 8)
          Call REPLA(TTXX$, FECHATEX$(FEX), 35, 8)
          Call REPLA(TTXX$, FORMATNUM$(IMPO#, "F12.2"), 43, 12)
          Call REPLA(TTXX$, FORMATNUM$(ICOM#, "F12.2"), 55, 12)
          '
          If JJ& < 1 Then
             Print #N1%, "Cliente                   Recibo  Fecha        Importe    Comision"
             Print #N1%, "------------------------------------------------------------------"
          End If
          JJ& = JJ& + 1
          Print #N1%, TTXX$
        End If
      Next IL&
      Print #N1%, "------------------------------------------------------------------"
      Print #N1%, "  "
      Print #N1%, "  "
      Close #N1%
   End If
   '
   ' *******************************************
   '
   If Check2.Value = 1 Then
      FIN& = ULTREG&("FAPECOB")
      UJ& = 0
      Call COPYSTRU("FAPECOB", "FAPECOX")
      '
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If VENDCLI%(NC%) = NVENDE% Then
          UJ& = UJ& + 1
          Call GRAREG("FAPECOX", X$, UJ&)
        End If
      Next U&
      Call SETULTREG("FAPECOX", UJ&)
      Call FILESORT("FAPECOX", "", 1, 1)
      '
      JJ& = 0
      N1% = FILEOPEN%("C:\FLEXOFT\CODEUCLI.TXT", 2, 128)
      FIN& = ULTREG&("FAPECOX")
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOX", U&)
        '
        TTXX$ = Space$(80)
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
        '
        Call REPLA(TTXX$, RASOCLI$(NC%), 1, 24)
        Call REPLA(TTXX$, NFS$, 27, 10)
        FEX = FF%
        Call REPLA(TTXX$, FECHATEX$(FEX), 39, 8)
        Call REPLA(TTXX$, FORMATNUM$(IMORIG#, "F12.2"), 47, 12)
        FEX = FVI%
        Call REPLA(TTXX$, FECHATEX$(FEX), 61, 8)
        Call REPLA(TTXX$, FORMATNUM$(SALDEU#, "F12.2"), 69, 12)
        '
        If JJ& < 1 Then
          Print #N1%, "Cliente                     Factura     Fecha      Importe   Vencim.       Saldo"
          Print #N1%, "--------------------------------------------------------------------------------"
        End If
        '
        JJ& = JJ& + 1
        Print #N1%, TTXX$
      Next U&
      Print #N1%, "--------------------------------------------------------------------------------"
      Print #N1%, ""
      Print #N1%, ""
      Close #N1%
   End If
   '
   '*******************************************************
   '
   MAI_DESTI$ = TRIM$(Label3)
   MAI_DIREL$ = TRIM$(Label4)
   MAI_ASUNT$ = "Comisiones y Detalle de Deuda"
   MAI_TEXTO$ = "Adjuntamos Detalle de Comisiones / Composicion de Deuda - Cordiales Saludos - FUINYTER S.H."
   '
   JKI% = 0
   If Check1.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\COMISION.TXT"
   If Check2.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\CODEUCLI.TXT"
   '
   Call GENERAMAIL
   Screen.MousePointer = 1
   MsgBox "Colocado en Bandeja de Salida"
   '
99 BOTREC.Enabled = True
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("VENDEDOR")
    Label1 = VALACT1$("VENDEDOR")
    Cancel = 0
    Command3.Enabled = True
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label1_Change()
   Label3 = ECOARCH$("VENDEDOR", Chr$(Val(Label1)))
   Label4 = ""
   RFF$ = RECFILT$("DIMAIVEN", 1, Chr$(Val(Label1)))
   If Len(RFF$) >= 4 Then
     REDI& = CVS(Left$(RFF$, 4))
     Label4 = TRIM$(Mid$(REGLEIDO$("DIMAIVEN", REDI&), 2))
   End If
End Sub

Private Sub Label1_DblClick()
    Call SELECHO("VENDEDOR")
    Label1 = VALACT1$("SELMESA")
    Cancel = 0
End Sub

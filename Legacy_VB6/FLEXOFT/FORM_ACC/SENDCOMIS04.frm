VERSION 5.00
Begin VB.Form SENDCOMIS04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar Comisiones por Mail"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Barrer"
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
      TabIndex        =   21
      ToolTipText     =   "Barre Todos los Vendedores"
      Top             =   1150
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cuenta Corriente Vendedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   315
      TabIndex        =   13
      Top             =   3465
      Width           =   4635
      Begin VB.CommandButton Command6 
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
         Height          =   285
         Left            =   1995
         TabIndex        =   18
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         TabIndex        =   17
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
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
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   3150
         TabIndex        =   15
         Top             =   420
         Width           =   1065
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Incluir Cuenta Corriente Vendedores"
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
         Left            =   105
         TabIndex        =   14
         Top             =   0
         Width           =   3480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   -1470
         TabIndex        =   20
         Top             =   480
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   2205
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
   End
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
      Height          =   500
      Left            =   4200
      Picture         =   "SENDCOMIS04.frx":0000
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
Attribute VB_Name = "SENDCOMIS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click_ant()
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
       If Check3.Value < 1 Then
          Call MENSERR(24, "No se Han Seleccionado Datos")
          GoTo 99
       End If
     End If
   End If
   '
   INCLUCUEC% = 0
   If Check3.Value > 0 Then
      INCLUCUEC% = 1
      DES% = FECHANUM(Text6)
      HAS% = FECHANUM(Text5)
      If DES% < 1 Or HAS% < DES% Then
         Call MENSERR(24, "Rango de Fechas no Procesable")
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
      FIN& = ULTREG&("FAPECOBA")
      UJ& = 0
      Call COPYSTRU("FAPECOBA", "FAPECOX")
      '
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOBA", U&)
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
        NFS$ = Mid$(X$, 49, 12)
        FF% = CVI(Mid$(X$, 61, 2))
        FVI% = CVI(Mid$(X$, 71, 2))
        SALDEU# = CVD(Mid$(X$, 73, 8))
        IMORIG# = CVD(Mid$(X$, 63, 8))
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
   If Check3.Value = 1 Then
      Call FILTERFILE("CUECOMIS", "CUECOMIM", 1, Chr$(NVENDE%))
      Call FILESORT("CUECOMIM", "", 2, 2, "ASC")
      FIN& = ULTREG&("CUECOMIM")
      '
      VACU# = 0
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECOMIM", U&)
        VACU# = VACU# + CVD(Mid$(X$, 41, 8)) - CVD(Mid$(X$, 49, 8))
        Call REPLA(X$, MKD$(VACU#), 57, 8)
        Call GRAREG("CUECOMIM", X$, U&)
      Next U&
      '
      JJ& = 0
      N1% = FILEOPEN%("C:\FLEXOFT\CUECOVEN.TXT", 2, 128)
      FIN& = ULTREG&("CUECOMIM")
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECOMIM", U&)
        '
        FF% = CVI(Mid$(X$, 2, 2))
        If FF% >= DES% Then
          If FF% <= HAS% Then
            TTXX$ = Space$(128)
            CLIX$ = Mid$(X$, 9, 32)
            NFS$ = Mid$(X$, 65, 32)
            IDEBE# = CVD(Mid$(X$, 41, 8))
            IHABE# = CVD(Mid$(X$, 49, 8))
            ISALD# = CVD(Mid$(X$, 57, 8))
            '
            FEX = FF%
            Call REPLA(TTXX$, FECHATEX$(FF%), 1, 8)
            Call REPLA(TTXX$, NFS$, 11, 32)
            Call REPLA(TTXX$, CLIX$, 45, 32)
            Call REPLA(TTXX$, FORMATNUM$(IDEBE#, "F12.2"), 87, 12)
            Call REPLA(TTXX$, FORMATNUM$(IHABE#, "F12.2"), 99, 12)
            Call REPLA(TTXX$, FORMATNUM$(ISALD#, "F12.2"), 111, 12)
            '
            If JJ& < 1 Then
              Print #N1%, "Fecha     Referencia                        Comprobante                                     Comis.        Pago       Saldo"
              Print #N1%, "--------------------------------------------------------------------------------------------------------------------------"
            End If
            '
            JJ& = JJ& + 1
            Print #N1%, TTXX$
            '
          End If
        End If
      Next U&
      '
      Print #N1%, "--------------------------------------------------------------------------------------------------------------------------"
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
   MAI_TEXTO$ = "Adjuntamos Detalle de Comisiones / Composicion de Deuda - Cordiales Saludos - "
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "SILCAR") > 0 Or InStr(EPAC$, "TELSO") > 0 Then
       MAI_TEXTO$ = MAI_TEXTO$ + "SILCAR S.A. / TELSO S.A."
     Else
       POS1% = InStr(EMPREAC$, " - ") 'PARA QUE NO PONGA EL USUARIO
       MAI_TEXTO$ = MAI_TEXTO$ + Mid$(EPAC$, 1, POS1% - 1) '"FUINYTER S.H."
   End If
   '
   JKI% = 0
   If Check1.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\COMISION.TXT"
   If Check2.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\CODEUCLI.TXT"
   If Check3.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\CUECOVEN.TXT"
   '
   Call GENERAMAIL
   Screen.MousePointer = 1
   If Command2.Enabled = True Then
      MsgBox "Colocado en Bandeja de Salida"
   End If
   '
99 BOTREC.Enabled = True
   '
End Sub
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
       If Check3.Value < 1 Then
          Call MENSERR(24, "No se Han Seleccionado Datos")
          GoTo 99
       End If
     End If
   End If
   '
   INCLUCUEC% = 0
   If Check3.Value > 0 Then
      INCLUCUEC% = 1
      DES% = FECHANUM(Text6)
      HAS% = FECHANUM(Text5)
      If DES% < 1 Or HAS% < DES% Then
         Call MENSERR(24, "Rango de Fechas no Procesable")
         GoTo 99
      End If
   End If
   '
   Screen.MousePointer = 11
   '
   If Check1.Value = 1 Then
      JJ& = 0
      N1% = FILEOPEN%("C:\FLEXOFT\COMISION.TXT", 2, 128)
      FIN& = ULTREG&("LICOMIN")
      For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("LICOMIN", IL&)
        If CVI(Mid$(X$, 19, 2)) = NVENDE% Then
          TTXX$ = Space$(80)
          NC% = CVI(Mid$(X$, 1, 2))
          NRECIB& = CVS(Mid$(X$, 3, 4))
          FEX = CVI(Mid$(X$, 7, 2))
          IMPO# = CVD(Mid$(X$, 9, 8))
          ICOM# = CVD(Mid$(X$, 49, 8))
          '
          Call REPLA(TTXX$, RASOCLI$(NC), 1, 24)
          Call REPLA(TTXX$, FORMATNUM$(NRECIB&, "F8.0"), 25, 8)
          Call REPLA(TTXX$, FECHATEX$(FEX), 35, 8)
          Call REPLA(TTXX$, FORMATNUM$(IMPO#, "F12.2"), 43, 12)
          Call REPLA(TTXX$, FORMATNUM$(ICOM#, "F12.2"), 55, 12)
'          Debug.Print TTXX$ & "  " & ACUMUIMPO#
'          ACUMUIMPO# = ACUMUIMPO# + IMPO#
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
      FIN& = ULTREG&("FAPECOBA")
      UJ& = 0
      Call COPYSTRU("FAPECOBA", "FAPECOX")
      '
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOBA", U&)
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
        NFS$ = Mid$(X$, 49, 12)
        FF% = CVI(Mid$(X$, 61, 2))
        FVI% = CVI(Mid$(X$, 71, 2))
        SALDEU# = CVD(Mid$(X$, 73, 8))
        IMORIG# = CVD(Mid$(X$, 63, 8))
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
   If Check3.Value = 1 Then
      Call FILTERFILE("CUECOMIS", "CUECOMIM", 1, Chr$(NVENDE%))
      Call FILESORT("CUECOMIM", "", 2, 2, "ASC")
      FIN& = ULTREG&("CUECOMIM")
      '
      VACU# = 0
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECOMIM", U&)
        VACU# = VACU# + CVD(Mid$(X$, 41, 8)) - CVD(Mid$(X$, 49, 8))
        Call REPLA(X$, MKD$(VACU#), 57, 8)
        Call GRAREG("CUECOMIM", X$, U&)
      Next U&
      '
      JJ& = 0
      N1% = FILEOPEN%("C:\FLEXOFT\CUECOVEN.TXT", 2, 128)
      FIN& = ULTREG&("CUECOMIM")
      For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECOMIM", U&)
        '
        FF% = CVI(Mid$(X$, 2, 2))
        If FF% >= DES% Then
          If FF% <= HAS% Then
            TTXX$ = Space$(128)
            CLIX$ = Mid$(X$, 9, 32)
            NFS$ = Mid$(X$, 65, 32)
            IDEBE# = CVD(Mid$(X$, 41, 8))
            IHABE# = CVD(Mid$(X$, 49, 8))
            ISALD# = CVD(Mid$(X$, 57, 8))
            '
            FEX = FF%
            Call REPLA(TTXX$, FECHATEX$(FF%), 1, 8)
            Call REPLA(TTXX$, NFS$, 11, 32)
            Call REPLA(TTXX$, CLIX$, 45, 32)
            Call REPLA(TTXX$, FORMATNUM$(IDEBE#, "F12.2"), 87, 12)
            Call REPLA(TTXX$, FORMATNUM$(IHABE#, "F12.2"), 99, 12)
            Call REPLA(TTXX$, FORMATNUM$(ISALD#, "F12.2"), 111, 12)
            '
            If JJ& < 1 Then
              Print #N1%, "Fecha     Referencia                        Comprobante                                     Comis.        Pago       Saldo"
              Print #N1%, "--------------------------------------------------------------------------------------------------------------------------"
            End If
            '
            JJ& = JJ& + 1
            Print #N1%, TTXX$
            '
          End If
        End If
      Next U&
      '
      Print #N1%, "--------------------------------------------------------------------------------------------------------------------------"
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
   MAI_TEXTO$ = "Adjuntamos Detalle de Comisiones / Composicion de Deuda - Cordiales Saludos - "
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "SILCAR") > 0 Or InStr(EPAC$, "TELSO") > 0 Then
       MAI_TEXTO$ = MAI_TEXTO$ + "SILCAR S.A. / TELSO S.A."
     Else
       POS1% = InStr(EMPREAC$, " - ") 'PARA QUE NO PONGA EL USUARIO
       MAI_TEXTO$ = MAI_TEXTO$ + Mid$(EPAC$, 1, POS1% - 1) '"FUINYTER S.H."
   End If
   '
   JKI% = 0
   If Check1.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\COMISION.TXT"
   If Check2.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\CODEUCLI.TXT"
   If Check3.Value = 1 Then JKI% = JKI% + 1: MAI_ADJUN$(JKI%) = "C:\FLEXOFT\CUECOVEN.TXT"
   '
'   Call GENERAMAIL
   Screen.MousePointer = 1
   If Command2.Enabled = True Then
      MsgBox "Colocado en Bandeja de Salida"
   End If
   '
99 BOTREC.Enabled = True
   '
End Sub


Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check3.Value = 0
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Check1.Value = 0
   End If
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If Check1.Value < 1 Then
     If Check2.Value < 1 Then
       If Check3.Value < 1 Then
          Call MENSERR(24, "No se Han Seleccionado Datos")
          GoTo 99
       End If
     End If
   End If
   '
   INCLUCUEC% = 0
   If Check3.Value > 0 Then
      INCLUCUEC% = 1
      DES% = FECHANUM(Text6)
      HAS% = FECHANUM(Text5)
      If DES% < 1 Or HAS% < DES% Then
         Call MENSERR(24, "Rango de Fechas no Procesable")
         GoTo 99
      End If
   End If
   '
   For U& = 1 To ULTREG&("VENDEDOR")
      X$ = REGLEIDO$("VENDEDOR", U&)
      Label1 = Asc(Left$(X$, 1))
      Call label1_Change
      If TRIM$(Label4) <> "" Then
         Call BOTREC_Click
      End If
   Next U&
   '
   MsgBox "Colocado en Bandeja de Salida"
   '
99 Command2.Enabled = True

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
   '
   Height = 3795
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "SILCAR") > 0 Or InStr(EPAC$, "TELSO") > 0 Then
      Check3.Value = 1
      Check1.Value = 0
      Height = 4965
      Text6 = FECHATEX$(COMESANT)
      Text5 = FECHATEX$(DIANTE(COMESACT))
   End If
   '
   Call CENTRAFORM(Me)
   TTXX$ = "Recuerde que el Envío de Composición de Deuda\Corresponde al Último Listado Emitido,\el que Puede Diferir de la Situación Actual.\  \Si Desea Refrescar Dicha Información\Vuelva a Pedir Aquel Listado."
   If COMALTER%(TTXX$ + "\\Continuar\Cancelar") <> 1 Then
      Call FINAL("")
   End If
   '
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub label1_Change()
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

Private Sub Text5_DblClick()
   Call Command10_Click
End Sub

Private Sub Text6_DblClick()
   Call Command6_Click
End Sub

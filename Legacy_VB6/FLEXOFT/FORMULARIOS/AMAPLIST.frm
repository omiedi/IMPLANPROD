VERSION 5.00
Begin VB.Form AMAPLIST 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Clientes - Consultas y Listados"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4370
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   2205
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12000
      End
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
      Left            =   4370
      Picture         =   "AMAPLIST.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   160
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1260
         Width           =   2325
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
            Left            =   -20
            Picture         =   "AMAPLIST.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Deudas y Límites de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   3
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   7
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command5 
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
            Left            =   -20
            Picture         =   "AMAPLIST.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Cuentas Inactivas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   9
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   10
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command8 
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
            Left            =   -20
            Picture         =   "AMAPLIST.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado General   y Clasificado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   12
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salida"
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
         Left            =   420
         TabIndex        =   4
         Top             =   2050
         Width           =   3165
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Impresión"
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
            Left            =   1785
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pantalla"
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
            Index           =   0
            Left            =   315
            TabIndex        =   5
            Top             =   250
            Width           =   1380
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4095
      Picture         =   "AMAPLIST.frx":0A68
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   5250
   End
End
Attribute VB_Name = "AMAPLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call LILIMICRE
   Command3.Enabled = True
End Sub

Private Sub Command5_Click()
'Exit Sub
    Command5.Enabled = False
    Call SELECHO("SELFECHA")
    FECHLI$ = VALACT1$("SELFECHA")
    FLIM% = FECHANUM(FECHLI$)
    If FLIM% > 0 Then
       Screen.MousePointer = 11
       Dim ACTI%(32767)
       FIN& = ULTREG&("FACTUR")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$("FACTUR", U&)
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If NC% <= 32767 Then
             If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
               ACTI%(NC%) = CVI(Mid$(X$, 7, 2))
             End If
           End If
         End If
       Next U&
       '
       RELISA& = 0
       CX& = ULTREG&("DEUDOR1")
       For IL& = 1 To CX&
         Call TRACE(20, IL&, CX&)
         NC% = CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
         If ACTI%(NC%) < FLIM% Then
            Y$ = REGBLAN$("LISACLI")
            SSCC# = SALDCLI#(NC%)
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, MKD$(SSCC#), 3, 8)
            Call REPLA(Y$, MKD$(SSCC#), 11, 8)
            FEX = ACTI%(NC%)
            TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
            Call REPLA(Y$, TTXX$, 19, 30)
            '
            RELISA& = RELISA& + 1
            Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next IL&
       '
       Call SETULTREG("LISACLI", RELISA&)
       Call CIERRARCH("*.*")
       Screen.MousePointer = 1
       Call CONECRUN("*DEUINAC", "Listado de Cuentas Inactivas")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command8_Click()
    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLICLIEN.Show 1
    Command8.Enabled = True
End Sub

Sub LILIMICRE()
   '
   Screen.MousePointer = 11
   '
   Dim DEUCUECO#(16384), CHEPEAC#(16384)
   '
   FIN& = ULTREG&("SVD202")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      SS$ = REGLEIDO$("SVD202", IL&)
      NCI% = CVI(Left$(SS$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            TI% = Asc(Mid$(SS$, 3, 1))
            If TI% > 7 Then TI% = 2
            SALDEU# = CVD(Mid$(SS$, 11, 8))
            If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
            DEUCUECO#(NCI%) = DEUCUECO#(NCI%) + SALDEU#
            GoTo 19
         End If
      End If
19 Next IL&
   '
   HOII% = HOY(HOS$)
   FIN& = ULTREG&("CHECAR")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("CHECAR", IL&)
      NCI% = CVI(Mid$(XX$, 3, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
               CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + CVD(Mid$(XX$, 33, 8))
            End If
         End If
      End If
   Next IL&
   '
   FIN& = ULTREG&("DEUDOR1")
   JJ& = 0
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", IL&)
      NCI% = CVI(Left$(XX$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
               YY$ = REGLEIDO$("DEUDOR2", IL&)
               CRELIM# = CVS(Mid$(YY$, 121, 4))
               DISPO# = CRELIM# - DEUCUECO#(NCI%) - CHEPEAC#(NCI%)
               EXCES# = 0
               If DISPO# < 0 Then
                  EXCES# = Abs(DISPO#)
                  DISPO# = 0
               End If
               ZZ$ = REGBLAN$("LILICRE")
               Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
               Call REPLA(ZZ$, MKD$(CRELIM#), 3, 8)
               Call REPLA(ZZ$, MKD$(DEUCUECO#(NCI%)), 11, 8)
               Call REPLA(ZZ$, MKD$(CHEPEAC#(NCI%)), 19, 8)
               Call REPLA(ZZ$, MKD$(DISPO#), 27, 8)
               Call REPLA(ZZ$, MKD$(EXCES#), 35, 8)
               JJ& = JJ& + 1
               Call GRAREG("LILICRE", ZZ$, JJ&)
            End If
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LILICRE", JJ&)
   Call CIERRARCH("LILICRE")
   '
   OPSALI$ = "IMPRE"
   On Error GoTo 40
   If Option1(0).Value = True Then OPSALI$ = "SCREEN"
   GoTo 41
   '
40 Resume 41
41 On Error GoTo 50
   If Option1(0).Value = True Then OPSALI$ = "SCREEN"
   GoTo 51
   '
50 Resume 51
51 On Error GoTo 0
   If OPSALI$ = "SCREEN" Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
   End If
   Screen.MousePointer = 1
   '
End Sub


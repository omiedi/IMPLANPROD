VERSION 5.00
Begin VB.Form ANUFORMU 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulacion de Comprobantes"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Cancela Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2140
      Picture         =   "ANUFORMU.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Cancela la Operación de Anulación"
      Top             =   3570
      Width           =   1800
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Confirma Anulación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   210
      Picture         =   "ANUFORMU.frx":030A
      TabIndex        =   7
      ToolTipText     =   "Anula el Número de Comprobante Mostrado"
      Top             =   3570
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Número de Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   210
      TabIndex        =   4
      Top             =   2310
      Width           =   3750
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1575
         TabIndex        =   10
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   735
         TabIndex        =   6
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   1410
         TabIndex        =   5
         Top             =   420
         Width           =   150
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tipo de Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   3750
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas de Exportación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   9
         Top             =   1050
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas, Débitos, Créditos 'B'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   3
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   2
         Top             =   1365
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Facturas, Débitos, Créditos 'A'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   3060
      End
   End
End
Attribute VB_Name = "ANUFORMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub NUFACTURA()
   '
   Static FLINUMERA%
   Static NUMECORR%
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   NUMECORR% = 1
   VARIDOC% = 1
   If Option1(1).Value = True Then VARIDOC% = 2
   If Option1(2).Value = True Then VARIDOC% = 3
   '
   If Option1(3).Value = True Then
      NROPRO& = 0
      Call NUREMIFA(NROPRO&)
      NUMAXI& = NROPRO& - 1
      GoTo 99
   End If
   '
20 NUMAXI& = 0: KKK% = 0
   For MXC& = ULTREG&("FACTUR") To 1 Step -1
       F$ = REGLEIDO$("FACTUR", MXC&)
       TIX% = Asc(Mid$(F$, 13, 1)): VAX% = Asc(Mid$(F$, 14, 1))
       SUCUVEX% = CVI(Mid$(F$, 47, 2)) '\\\OT-06-0077-WAR-03/03/06///
       '
       If SUCUVEX% = Val(Label12) Then '\\\OT-06-0077-WAR-03/03/06///
        If TIX% > 0 And TIX% < 7 Then
          If TIX% = TIFADOC% Or NUMECORR% = 1 Then
            If VAX% = VARIDOC% Or (VARIDOC% = 3 And VAX% = 0) Then
              If CVI(Mid$(F$, 7, 2)) < FLINUMERA% And KKK% > 32 Then GoTo 99
              If CVS(Mid$(F$, 3, 4)) > NUMAXI& Then
                NUMAXI& = CVS(Mid$(F$, 3, 4))
              End If
              KKK% = KKK% + 1: If KKK% > 128 Then GoTo 99
            End If
          End If
        End If
       End If
       '
   Next MXC&
   '
99 NPP$ = TRIM$(Str$(NUMAXI& + 1))
   While Len(NPP$) < 8: NPP$ = "0" + NPP$: Wend
   Label1 = TRIM$(NPP$)
   Label1.Refresh
   '
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Unload ANUFORMU
End Sub

Private Sub Command13_Click()
   '
   VARIDOC% = 1
   NUMECOM& = Val(Label1)
   FF% = HOY(HOS$)
   '
   If Option1(1).Value = True Then VARIDOC% = 2
   If Option1(2).Value = True Then VARIDOC% = 3
   '
   If Option1(3).Value = True Then
      CI% = 0
      CMOV$ = "RT"
      LOTE$ = ""
      DOPRI$ = "RT." + TRIM$(Str$(NUMECOM&))
      REFE$ = "Anulacion de Comprobante"
      CANTAA = 1
      Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, "", REFE$, PRE#, "$", NC%, DEPOSALI%, CANTAA)
      Call CIERRARCH("*.*")
      Call NUFACTURA
      Exit Sub
   End If
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(32767), 1, 2)
   Call REPLA(X$, MKS$(NUMECOM&), 3, 4)
   Call REPLA(X$, MKI$(FF%), 7, 2)
   Call REPLA(X$, Chr$(2) + Chr$(VARIDOC%), 13, 2)
   '\\\OT-06-0060-WAR-22/02/06///
   SUCUVEN% = Val(Label12)
   Call REPLA(X$, MKI$(SUCUVEN%), 47, 2)
   '\\\OT-06-0060-WAR-FIN///
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("*.*")
   Call NUFACTURA
   Exit Sub
   '
End Sub

Private Sub Form_Load()
   Call NUFACTURA
End Sub

Private Sub Label12_Change()
  '\\\OT-06-0077-WAR-03/03/06///
  Call NUFACTURA
  '\\\OT-06-0077-WAR-FIN///
End Sub

Private Sub Label12_DBLClick()
  '\\\OT-06-0077-WAR-03/03/06///
  If Label12 = "0001" Then
      Label12 = "0002"
    ElseIf Label12 = "0002" Then
      Label12 = "0003"
    ElseIf Label12 = "0003" Then
      Label12 = "0004"
    ElseIf Label12 = "0004" Then
      Label12 = "0005"
    ElseIf Label12 = "0005" Then
      Label12 = "0001"
  End If
  '\\\OT-06-0077-WAR-FIN///
End Sub

Private Sub Option1_Click(Index As Integer)
   Call NUFACTURA
End Sub

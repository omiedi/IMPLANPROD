VERSION 5.00
Begin VB.Form XSORTAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulo de Ordenamiento Automático "
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5385
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   945
      ScaleHeight     =   75
      ScaleWidth      =   3360
      TabIndex        =   7
      Top             =   2100
      Width           =   3420
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   105
      Top             =   945
   End
   Begin VB.ListBox LISORT 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   1785
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   8625
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3780
      TabIndex        =   5
      Text            =   " "
      Top             =   1155
      Width           =   1420
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3780
      TabIndex        =   2
      Text            =   " "
      Top             =   210
      Width           =   1420
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   210
      TabIndex        =   1
      Text            =   " "
      Top             =   1575
      Width           =   4980
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   0
      Text            =   " "
      Top             =   630
      Width           =   4980
   End
   Begin VB.Label Label3 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo de Trabajo:"
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
      Left            =   1365
      TabIndex        =   4
      Top             =   315
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo de Salida:"
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
      Left            =   1260
      TabIndex        =   3
      Top             =   1260
      Width           =   2325
   End
End
Attribute VB_Name = "XSORTAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RGFJ%(32), TFF%(32)
Dim VFMIN$(32), VFMAX$(32)
Dim CALIMI As Integer
Dim CALIMS As Integer
Dim ARCHISAL$

Private Sub BOTEXIT_Click()

End Sub

Private Sub Form_Activate()
   If Label3 <> "" Then
     Refresh
     DoEvents
     Call SORTABU
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Private Sub Label3_Change()
   Visible = False
   If Label3 <> "" Then
     Call CARDATSOR
   End If
End Sub

Sub CARDATSOR()
     '
     FORMA$ = Label3
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     '
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     ARCHITRA$ = TRIM$(UCase$(Mid$(YYY$, 57, 8)))
     Text1.Text = ARCHITRA$
     Text3.Text = TITUARCH$(ARCHITRA$)
     ARCHISAL$ = TRIM$(UCase$(Mid$(YYY$, 65, 8)))
     If ARCHISAL$ <> ARCHITRA$ Then
         Call COPYSTRU(ARCHITRA$, ARCHISAL$)
     End If
     Text4.Text = ARCHISAL$
     Text2.Text = TITUARCH$(ARCHISAL$)
     Refresh
     '
     CALIMI% = Asc(Mid$(YYY$, 73, 1))
     MORDE% = Asc(Mid$(YYY$, 74, 1))
     CALIMS% = Asc(Mid$(YYY$, 75, 1))
        If CALIMS% < CALIMI% Then CALIMS% = CALIMI%
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     '
5803 S1$ = Left$(YYY$, 48)
     SALI$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = Asc(Mid$(COD$, 7))
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: RG% = 2
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
     If RG% = 2 Then
         TVL%(TCL%) = 1
         LDL%(TCL%) = 0
     End If
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
     If RG% = 2 Then
         PR%(TCL%) = 0
         LF%(TCL%) = 2
     End If
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     PE%(NCL%) = Asc(Mid$(YYY$, 73))
     If RG% = 2 Then PE%(NCL%) = 2
5815 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
5816 GoTo 5805
5817 '
5820 ATOT% = NCL% - 1
     CLXX% = OFS% + 2
     For KCL% = 1 To NCL%
        LC%(KCL%) = ACL%(KCL%)
        ATOT% = ATOT% + ACL%(KCL%)
        CL%(KCL%) = CLXX%
        CLXX% = CLXX% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NCL%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
 '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 88)) < 1 Then GoTo 5847   ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
     RGFJ%(NFT%) = Asc(Mid$(YYY$, 84))
5843 If RGFJ%(NFT%) - 1 < RG0% Then RG0% = RGFJ%(NFT%) - 1
5844 ECL$(NFT%) = TRIM$(Mid$(YYY$, 37, 32))
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70)): LEF%(NFT%) = Asc(Mid$(YYY$, 86)): CEF%(NFT%) = Asc(Mid$(YYY$, 87)): TFF%(NFT%) = Asc(Mid$(YYY$, 88))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8): VFMIN$(NFT%) = Mid$(YYY$, 113, 8): VFMAX$(NFT%) = Mid$(YYY$, 121, 8)
5847 GoTo 5835
 '
5850 S1$ = TRIM$(S1$)
     '
5860 Close #N15%: N15% = 0
     '
End Sub
'
Sub SORTABU()
    '
     'Timer1.Enabled = False
     'Refresh
     Screen.MousePointer = 11
     URETRA& = ULTREG&(ARCHITRA$)
     If URETRA& < 1 Then GoTo 199
     '
     FINDEX% = FILEINDEX%(ARCHITRA$)
     FINUME% = FILENBR%(ARCHITRA$)
     LOREGI% = LOREGIS%(FINDEX%)
     '
     LISORT.Clear
     For RECNO& = 1 To URETRA&
        '
        Call TRACE(20, RECNO&, URETRA&)
        'XXX$ = REGLEIDO$(ARCHITRA$, RECNO&)
        XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
        '
        ARGU$ = "": LARGU% = Len(ARGU$)
        For J% = CALIMI% To CALIMS%
           XXZ$ = Mid$(XXX$, PR%(J%), LF%(J%))
           TV1% = TVL%(J%): If TV1% = 5 Or TV1% = 8 Then TV1% = 1
           LF1% = LF%(J%)
           AC1% = LF%(J%)
           JT1% = 1: LD1% = LDL%(J%)
           '
           If TV1% = 1 Or TV1% = 2 Then
                  JT1% = 2: LD1% = 0
                  If LF1% < 2 Then AC1% = 3 Else AC1% = 5
               ElseIf TV1% = 3 Or TV1% = 4 Then
                  JT1% = 2
                  If LF1% < 8 Then AC1% = 7 Else AC1% = 12
                  If LD1% > 0 Then AC1% = AC1% + 2
               ElseIf TV1% = 5 Then
                  AC1% = 8: LD1% = 0
               ElseIf TV1% = 7 Then
                  AC1% = 16: LD1% = 0
           End If
           '
           ARGX$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
           ARGU$ = ARGU$ + ARGX$
        Next J%
        '
        If LARGU% = 0 Then
            LARGU% = Len(ARGU$)
        End If
        '
        ARGU$ = ARGU$ + CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
        LISORT.AddItem ARGU$
        '
    Next RECNO&
    '
    PPX% = LARGU% + 1
    ARCHINTER$ = ARCHISAL$
    If ARCHISAL$ = ARCHITRA$ Then
        ARCHINTER$ = "!" + ARCHITRA$
    End If
    '
    JJ& = 0
    FINDEXINT% = FILEINDEX%(ARCHINTER$)
    FINUINTER% = FILENBR%(ARCHINTER$)
    FINPRO& = LISORT.ListCount - 1
    For U& = 0 To LISORT.ListCount - 1
       Call TRACE(20, U&, FINPRO&)
       RECNO& = Val(Mid$(LISORT.List(U&), PPX% + 1))
       If RECNO& > 0 Then
          If RECNO& <= URETRA& Then
             'XXX$ = REGLEIDO$(ARCHITRA$, RECNO&)
             XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
             JJ& = JJ& + 1
             'Call GRAREG(ARCHINTER$, XXX$, JJ&)
             POSGRABA& = (JJ& - 1) * LOREGI% + 1
             Put #FINUINTER%, POSGRABA&, XXX$
          End If
       End If
    Next U&
    '
    LISORT.Clear
    FINPRO& = JJ&
    If ARCHISAL$ = ARCHITRA$ Then
       JJ& = 0
       For U& = 1 To FINPRO&
          Call TRACE(20, U&, FINPRO&)
          'XXX$ = REGLEIDO$(ARCHINTER$, U&)
          XXX$ = LEEREGIS$(FINDEXINT%, FINUINTER%, LOREGI%, U&)
          JJ& = JJ& + 1
          POSGRABA& = (JJ& - 1) * LOREGI% + 1
          Put #FINUME%, POSGRABA&, XXX$
          'Call GRAREG(ARCHITRA$, XXX$, JJ&)
       Next U&
       Call SETULTREG(ARCHITRA$, JJ&)
    End If
    Call SETULTREG(ARCHINTER$, FINPRO&)
    '
    Call CIERRARCH(ARCHINTER$)
    Call CIERRARCH(ARCHITRA$)
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    LISORT.Clear
    Screen.MousePointer = 1
    '
    Hide
    '
End Sub
'


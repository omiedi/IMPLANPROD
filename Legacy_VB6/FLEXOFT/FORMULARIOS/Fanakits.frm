VERSION 5.00
Begin VB.Form FANAKITS 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Pedidos de KITS"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Situación al:"
      Height          =   990
      Left            =   210
      TabIndex        =   16
      Top             =   315
      Width           =   4125
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   320
         Left            =   1080
         TabIndex        =   18
         Top             =   440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   13
         Left            =   -1260
         TabIndex        =   20
         Top             =   525
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Index           =   12
         Left            =   1380
         TabIndex        =   19
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   320
         Left            =   3120
         TabIndex        =   17
         Top             =   440
         Width           =   735
      End
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4500
      Picture         =   "Fanakits.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Recálculo - Actualización"
      Height          =   1455
      Left            =   210
      TabIndex        =   9
      Top             =   1470
      Width           =   4110
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   10
         Top             =   525
         Width           =   2235
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
            Picture         =   "Fanakits.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Revisar Situa- ción de Stocks"
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
            TabIndex        =   11
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consultas - Listados"
      Height          =   3525
      Left            =   210
      TabIndex        =   2
      Top             =   3150
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Necesidad Mensual de Piezas para Kits"
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
         Left            =   525
         TabIndex        =   22
         Top             =   1785
         Width           =   4275
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Necesidad de Componentes para Kits"
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
         Left            =   525
         TabIndex        =   21
         Top             =   840
         Width           =   4275
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Conjuntos tipo Kit a Componer"
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
         Left            =   525
         TabIndex        =   15
         Top             =   1470
         Width           =   4275
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   4695
         TabIndex        =   13
         Top             =   3150
         Width           =   4750
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
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4095
         Picture         =   "Fanakits.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2385
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pedidos de Kits sin Entrega - Falta de Stock"
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
         Left            =   525
         TabIndex        =   7
         Top             =   525
         Value           =   -1  'True
         Width           =   4110
      End
      Begin VB.Frame Frame6 
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
         Left            =   210
         TabIndex        =   4
         Top             =   2310
         Width           =   3690
         Begin VB.OptionButton Option5 
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
            Left            =   315
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
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
            Left            =   2100
            TabIndex        =   5
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Componentes Faltantes para Armado de Kits"
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
         Left            =   520
         TabIndex        =   3
         Top             =   1155
         Width           =   4500
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
      Height          =   855
      Left            =   4500
      Picture         =   "Fanakits.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3960
      Picture         =   "Fanakits.frx":0A68
      Top             =   2400
      Width           =   2010
   End
End
Attribute VB_Name = "FANAKITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    OPCI% = 1
    If Option2.Value = True Then OPCI% = 2
    If Option1.Value = True Then OPCI% = 3
    If Option7.Value = True Then OPCI% = 4
    If Option3.Value = True Then OPCI% = 5
    If OPCI% = 1 Then
        If Option5.Value = True Then
             Call CONECRUN("*COFAKIT", "")
           Else
             Call CONECRUN("*LIFAKIT", "")
        End If
      ElseIf OPCI% = 2 Then
        If Option5.Value = True Then
             Call CONECRUN("*CONEKIT", "")
           Else
             Call CONECRUN("*LINEKIT", "")
        End If
      ElseIf OPCI% = 3 Then
        If Option5.Value = True Then
             Call CONECRUN("*COTOKIT", "")
           Else
             Call CONECRUN("*LITOKIT", "")
        End If
      ElseIf OPCI% = 4 Then
        If Option5.Value = True Then
             Call CONECRUN("*COARKIT", "")
           Else
             Call CONECRUN("*LIARKIT", "")
        End If
      ElseIf OPCI% = 5 Then
        If Option5.Value = True Then
             Call CONECRUN("*CONEPIE", "")
           Else
             Call CONECRUN("*LINEPIE", "")
        End If
      Else
        Hide
    End If
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call REVIPEDKIT
    ' Call OPNOIN("")
    Command3.Enabled = True
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("ANAKITS")
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Label21.Caption = CONTROL$("", "FEREKITS")
    Label22.Caption = CONTROL$("", "HOREKITS")
    '
End Sub

Sub REVIPEDKIT()
   '
   Screen.MousePointer = 11
   '
   Dim KIT%(16384)
   Dim CAPE(16384)
   Dim VAPE#(16384)
   Dim TOCO(16384)
   Dim RESER(16384)
   '
   FIN& = NUMAS%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      CIII% = U&
      If InStr(UCase$(DESCRIT$(CIII%)), "KIT") = 1 Then
         KIT%(CIII%) = 1
      End If
   Next U&
   '
   FIN& = ULTREG&("PEDDETP")
   '
   MODO% = 2
   If MODO% = 2 Then
      FLI1% = -1: FLI2% = 32767
   End If
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETP", U&)
      CISX% = CVI(Mid$(X$, 33, 2))
      If CISX% > 0 Then
         If CISX% <= 16384 Then
            If KIT%(CISX%) = 1 Then
               FENTRESOL% = CVI(Mid$(X$, 25, 2))
               If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                     CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                     If CAPEN < 0 Then CAPEN = 0
                     If CAPEN > 0 Then
                        NC% = CVI(Mid$(X$, 7, 2))
                        PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        '
                        CAPE(CISX%) = CAPE(CISX%) + CAPEN
                        VAPE#(CISX%) = VAPE#(CISX%) + ITOTA#
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   NMS& = NUMAS%
   JJ& = 0: KK& = 0: LL& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If KIT%(CISX%) = 1 Then
         If CAPE(CISX%) >= 0.05 Then
            STACT = STODEPOS(CISX%, DEPOEXPE%)
            FALTA = 0
            If STACT < CAPE(CISX%) Then
               FALTA = CAPE(CISX%) - STACT
            End If
            Y$ = REGBLAN$("PEDIKIT")
            Call REPLA(Y$, MKI$(CISX%), 1, 2)
            Call REPLA(Y$, MKS$(STACT), 3, 4)
            Call REPLA(Y$, MKS$(CAPE(CISX%)), 7, 4)
            Call REPLA(Y$, MKD$(VAPE#(CISX%)), 11, 8)
            Call REPLA(Y$, MKS$(FALTA), 19, 4)
            Call REPLA(Y$, MKD$(VAPE#(CISX%) * FALTA / CAPE(CISX%)), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("PEDIKIT", Y$, JJ&)
            '
            If FALTA >= 0.05 Then
               CODKI$ = CODEXT$(CISX%)
               Call LEEEXPLOMUL(CODKI$, 1, 2)
               If CAIT% < 1 Then
                  Call MENSERR(24, "Kit " + TRIM$(CODKI$) + "\" + DESCRIT$(CISX%) + "\   \Falta Fórmula de Composición de este Kit.")
               End If
               '
               Z$ = REGBLAN$("NECOMKIT")
               Call REPLA(Z$, MKI$(CISX%), 1, 2)
               Call REPLA(Z$, DESCRIT$(CISX%), 3, 64)
               Call REPLA(Z$, MKD$(CDbl(FALTA)), 67, 8)
               KK& = KK& + 1
               Call GRAREG("NECOMKIT", Z$, KK&)
               '
               Call REPLA(Z$, Space$(64), 3, 64)
               Call REPLA(Z$, MKD$(0), 67, 8)
               KK& = KK& + 1
               Call GRAREG("NECOMKIT", Z$, KK&)
               '
               MAXAR# = 99999999
               For ITI% = 1 To CAIT%
                  NECOM# = FALTA * USOI(ITI%)
                  STOAC# = STODEPOS(CIJ%(ITI%), DEPOEXPE%)
                  STASI# = STOCKASI(CIJ%(ITI%))
                  STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
                  FALCO# = NECOM# - STDIS#: If FALCO# < 0 Then FALCO# = 0
                  TOCO(CIJ%(ITI%)) = TOCO(CIJ%(ITI%)) + FALTA * USOI(ITI%)
                  '
                  Call REPLA(Z$, CODEXT$(CIJ%(ITI%)), 3, 16)
                  Call REPLA(Z$, DESCRIT$(CIJ%(ITI%)), 19, 48)
                  Call REPLA(Z$, MKD$(NECOM#), 67, 8)
                  Call REPLA(Z$, MKD$(STOAC#), 75, 8)
                  Call REPLA(Z$, MKD$(STASI#), 83, 8)
                  Call REPLA(Z$, MKD$(STDIS#), 91, 8)
                  Call REPLA(Z$, MKD$(FALCO#), 99, 8)
                  '
                  KK& = KK& + 1
                  Call GRAREG("NECOMKIT", Z$, KK&)
                  '
                  MAXAS# = (STDIS# - RESER(CIJ%(ITI%))) / USOI(ITI%)
                  If MAXAS# < MAXAR# Then
                     MAXAR# = MAXAS#
                  End If
               Next ITI%
               '
               MAXAR# = Int(MAXAR#)
               ARMAR# = FALTA
               If MAXAR# < ARMAR# Then
                  ARMAR# = MAXAR#
               End If
               ARMAR# = Int(ARMAR#)
               '
               T$ = REGBLAN$("ARMAKIT")
               Call REPLA(T$, MKI$(CISX%), 1, 2)
               Call REPLA(T$, MKD$(CDbl(FALTA)), 3, 8)
               Call REPLA(T$, MKD$(MAXAR#), 11, 8)
               Call REPLA(T$, MKD$(ARMAR#), 19, 8)
               LL& = LL& + 1
               Call GRAREG("ARMAKIT", T$, LL&)
               '
               For ITI% = 1 To CAIT%
                  RESER(CIJ%(ITI%)) = RESER(CIJ%(ITI%)) + ARMAR# * USOI(ITI%)
               Next ITI%
               '
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("PEDIKIT", JJ&)
   Call CIERRARCH("PEDIKIT")
   Call SETULTREG("NECOMKIT", KK&)
   Call CIERRARCH("NECOMKIT")
   Call SETULTREG("ARMAKIT", LL&)
   Call CIERRARCH("ARMAKIT")
   '
   NMS& = NUMAS%
   JJ& = 0: KK& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If TOCO(CISX%) >= 0.05 Then
         NECOM# = TOCO(CISX%)
         STOAC# = STODEPOS(CISX%, DEPOEXPE%)
         STASI# = STOCKASI(CISX%)
         STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
         FALCO# = NECOM# - STDIS#: If FALCO# < 0 Then FALCO# = 0
         '
         Z$ = REGBLAN$("TOCOMKIT")
         Call REPLA(Z$, MKI$(CISX%), 1, 2)
         Call REPLA(Z$, MKD$(NECOM#), 3, 8)
         Call REPLA(Z$, MKD$(STOAC#), 11, 8)
         Call REPLA(Z$, MKD$(STASI#), 19, 8)
         Call REPLA(Z$, MKD$(STDIS#), 27, 8)
         Call REPLA(Z$, MKD$(FALCO#), 35, 8)
         '
         KK& = KK& + 1
         Call GRAREG("TOCOMKIT", Z$, KK&)
      End If
   Next U&
   '
   Call SETULTREG("TOCOMKIT", KK&)
   Call CIERRARCH("TOCOMKIT")
   '
   For II% = 1 To NUMAS%
      CAPE(II%) = 0
      TOCO(II%) = 0
      RESER(II%) = 0
   Next II%
   '
   CAMESES = 6
   HOII = HOY(HO$)
   M0% = Val(Mid$(HO$, 4, 2)): A0% = Val(Mid$(HO$, 7, 2))
   A2% = A0%: M2% = M0% - 1: If M2% < 1 Then M2% = M2% + 12: A2% = A2% - 1
   A1% = A2%: M1% = M2% - CAMESES + 1: If M1% < 1 Then M1% = M1% + 12: A1% = A1% - 1
   If A2% < 0 Then A2% = A2% + 100
   If A1% < 0 Then A1% = A1% + 100
   '
   MS1$ = TRIM$(Str$(M1%)): While Len(MS1$) < 2: MS1$ = "0" + MS1$: Wend
   AS1$ = TRIM$(Str$(A1%)): While Len(AS1$) < 2: AS1$ = "0" + AS1$: Wend
   DXS$ = "01/" + MS1$ + "/" + AS1$: DES% = FECHANUM(DXS$)
   '
   MS2$ = TRIM$(Str$(M2%)): While Len(MS2$) < 2: MS2$ = "0" + MS2$: Wend
   AS2$ = TRIM$(Str$(A2%)): While Len(AS2$) < 2: AS2$ = "0" + AS2$: Wend
   HXS$ = "31/" + MS2$ + "/" + AS2$: HAS% = FECHANUM(HXS$)
   '
   FIN& = ULTREG&("MOVISTO")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      MVX$ = REGLEIDO$("MOVISTO", I&)
      X$ = MVX$
      '
      FF% = CVI(Left$(X$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CISX% = CVI(Mid$(X$, 3, 2))
          If CISX% > 0 Then
            If CISX% <= NUMAS% Then
              If KIT%(CISX%) = 1 Then
                COMO$ = UCase$(Mid$(X$, 33, 2))
                If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                  CAPE(CISX%) = CAPE(CISX%) + CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                End If
              End If
            End If
          End If
        End If
      End If
   Next I&
   '
   NMS& = NUMAS%
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If KIT%(CISX%) = 1 Then
        If CAPE(CISX%) >= 0.05 Then
          '
          FALTA = CAPE(CISX%) / CAMESES
          CODKI$ = CODEXT$(CISX%)
          Call LEEEXPLOMUL(CODKI$, 1, 2)
          If CAIT% < 1 Then
             Call MENSERR(24, "Kit " + TRIM$(CODKI$) + "\" + DESCRIT$(CISX%) + "\   \Falta Fórmula de Composición de este Kit.")
          End If
          '
          For ITI% = 1 To CAIT%
            TOCO(CIJ%(ITI%)) = TOCO(CIJ%(ITI%)) + FALTA * USOI(ITI%)
'If TRIM$(CODEXT$(CIJ%(ITI%))) = "S02544.00V" Then
'FFF = FALTA * USOI(ITI%)
'End If
          Next ITI%
          '
        End If
      End If
   Next U&
   '
   KK& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If TOCO(CISX%) >= 0.05 Then
        NECOM# = TOCO(CISX%)
        STOAC# = STODEPOS(CISX%, DEPOEXPE%)
        STASI# = STOCKASI(CISX%)
        STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
        '
        Z$ = REGBLAN$("NEMENPIE")
        Call REPLA(Z$, MKI$(CISX%), 1, 2)
        Call REPLA(Z$, MKD$(NECOM#), 3, 8)
        Call REPLA(Z$, MKD$(STOAC#), 11, 8)
        Call REPLA(Z$, MKD$(STASI#), 19, 8)
        Call REPLA(Z$, MKD$(STDIS#), 27, 8)
        COBE = STDIS# / NECOM#: If COBE > 999 Then COBE = 999
        Call REPLA(Z$, MKS$(COBE), 35, 4)
        '
        KK& = KK& + 1
        Call GRAREG("NEMENPIE", Z$, KK&)
        '
      End If
   Next U&
   '
   Call SETULTREG("NEMENPIE", KK&)
   Call CIERRARCH("NEMENPIE")
   Call CIERRARCH("*.*")
   Call FINAL("*SORKIT1")
   '
   Call GRACONTROL("", "FEREKITS", HO$)
   Call GRACONTROL("", "HOREKITS", Left$(Time$, 5))
   FANAKITS.Label21.Caption = HO$
   FANAKITS.Label22.Caption = Left$(Time$, 5)
   
   Screen.MousePointer = 1
   '
End Sub
'


VERSION 5.00
Begin VB.Form FLEXOFTA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT (r) Soluciones Informáticas Integrales - Versión 2002 Gold"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10905
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8326.541
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   13198.18
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      FillStyle       =   2  'Horizontal Line
      Height          =   2894
      Left            =   5670
      ScaleHeight     =   2835
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2527.918
      TabIndex        =   6
      Top             =   210
      Visible         =   0   'False
      Width           =   5040
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2920
         Left            =   -99
         TabIndex        =   7
         Top             =   -70
         Width           =   5122
         Begin VB.Line Line4 
            X1              =   5070
            X2              =   5070
            Y1              =   105
            Y2              =   2940
         End
         Begin VB.Line Line3 
            X1              =   105
            X2              =   5100
            Y1              =   2890
            Y2              =   2890
         End
         Begin VB.Image Image1 
            Height          =   315
            Left            =   1050
            Picture         =   "FLEXOFTA.frx":0000
            Stretch         =   -1  'True
            Top             =   1260
            Width           =   375
         End
         Begin VB.Image Image2 
            Height          =   375
            Left            =   1050
            Picture         =   "FLEXOFTA.frx":030A
            Stretch         =   -1  'True
            Top             =   1995
            Width           =   420
         End
         Begin VB.Image Image3 
            Height          =   375
            Left            =   1050
            Picture         =   "FLEXOFTA.frx":0614
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " Sistema Preconfigurado FLEXOFT "
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
            Left            =   105
            TabIndex        =   14
            Top             =   210
            Width           =   4845
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   945
            TabIndex        =   13
            Top             =   420
            Width           =   3060
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Estudios C.I.M.E. - S.O.I."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   315
            TabIndex        =   12
            Top             =   735
            Width           =   4425
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Bauness 2170 - 5to 'D' C1431DNP Buenos Aires República Argentina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   1785
            TabIndex        =   11
            Top             =   1155
            Width           =   2745
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "00-54-11-4524-1284"
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
            Left            =   1785
            TabIndex        =   10
            Top             =   1995
            Width           =   2220
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "(15) 4473-1634"
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
            Left            =   1785
            TabIndex        =   9
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "4541-1515"
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
            Left            =   2730
            TabIndex        =   8
            Top             =   2205
            Width           =   1695
         End
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3045
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   525
      Top             =   3465
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillStyle       =   2  'Horizontal Line
      Height          =   2894
      Left            =   182
      ScaleHeight     =   2835
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2695.431
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   5370
      Begin VB.Line Line2 
         X1              =   2689.847
         X2              =   2689.847
         Y1              =   0
         Y2              =   2835
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2718.274
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones Informaticas Integrales "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   3150
         TabIndex        =   15
         Top             =   120
         Width           =   2010
      End
      Begin VB.Image Image4 
         Height          =   3135
         Left            =   0
         Picture         =   "FLEXOFTA.frx":091E
         Top             =   -300
         Width           =   5370
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones Informáticas Integrales "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2835
         TabIndex        =   4
         Top             =   105
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   525
      Top             =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "y/o imitación mediante prácticas de ingenieria inversa, salvo autorizacion formal y expresa del autor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      TabIndex        =   2
      Top             =   6615
      Visible         =   0   'False
      Width           =   7890
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FLEXOFTA.frx":377D4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1575
      TabIndex        =   1
      Top             =   6090
      Visible         =   0   'False
      Width           =   7890
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1785
      TabIndex        =   0
      Top             =   5250
      Visible         =   0   'False
      Width           =   7365
   End
End
Attribute VB_Name = "FLEXOFTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    '
    Static CTX%
    '
    If CTX% = 0 Then
        FLEXOFTA.Picture1.Top = FLEXOFTA.ScaleHeight / 2 - FLEXOFTA.Picture1.Height / 2
        FLEXOFTA.Picture1.Left = FLEXOFTA.ScaleWidth / 2 - FLEXOFTA.Picture1.Width / 2
        Picture1.Visible = True
        Label6.Visible = True
        Timer1.Enabled = True
        CTX% = 1
    End If
    '
    MENUPRIN$ = TRIM$(Command$)
    If MENUPRIN$ <> "" Then
       Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
    End If
    '
End Sub
'

Private Sub Form_Load()
    '
    ABC$ = LOADFILE$("LANGUAGE.DRV")
    LANGUA% = 0
    If Val(ABC$) > 0 Then
      If Val(ABC$) <= 5 Then
        LANGUA% = Val(ABC$)
      End If
    End If
    '
    EPAC$ = TRIM$(EMPRELI$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Timer1_Timer()
    Static VECES, CT1%, TINI, CT2%
    VECES = VECES + 1
    '
    If CT1% > 0 Then GoTo 10
    '
    If VECES >= 12 Then
        FLEXOFTA.Label1.Visible = True
        FLEXOFTA.Label2.Visible = True
        FLEXOFTA.Label3.Visible = True
    End If
    '
    If VECES > 22 Then
       Timer2.Enabled = True
       CT1% = 1
    End If
    '
10  'If Frame1.Visible = True Then
    If Picture2.Visible = True Then
       '
       If TINI = 0 Then
          TINI = Timer
          '
          Call VERIDIR
          EMPREAC$ = EMPRELI$
          Call SELUDAT
          Call VERITER
          '
          IMPREACTU.Show 1  ' MUESTRA IMPRESORA PREDETERMINADA
          Call FRESHALL
          '
          Call CARDATRAM
          If VERCD% <> 1 Then
            If EXISTE%(LUDAT$ + "SVD147.EXE") < 1 Then
              DATEMPRE.Show 1
              Call FRESHALL
            End If
          End If
          '
          If VERCD% = 1 Then
            VERDEMO% = 1  ' (-1)
          End If
          '
          If VERDEMO% = 0 Then
              EPL$ = UCase$(EMPRELI$)
              If Left$(EPL$, 4) = "SABO" Then VERDEMO% = (-1)
              If Left$(EPL$, 7) = "LODIGIA" Then VERDEMO% = (-1)
              If Left$(EPL$, 4) = "NOVO" Then VERDEMO% = (-1)
              If Left$(EPL$, 7) = "YALTRES" Then VERDEMO% = (-1)
              If Left$(EPL$, 6) = "AMBEST" Then VERDEMO% = (-1)
              If Left$(EPL$, 5) = "FORJA" Then VERDEMO% = (-1)
              If Left$(EPL$, 7) = "TALEMEC" Then VERDEMO% = (-1)
              If Left$(EPL$, 5) = "FERVI" Then VERDEMO% = (-1)
              If Left$(EPL$, 5) = "STADT" Then VERDEMO% = (-1)
              If Left$(EPL$, 7) = "NEUMANN" Then VERDEMO% = (-1)
              If InStr(EPL$, "DOSIVAC") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "VDR") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "GUIDI") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "FUINY") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "ARCON") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "DOVAL") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "DEVOTO") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "CIME-SOI") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "FAEM") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "BORENS") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "MAINIERI") > 0 Then VERDEMO% = (-1)
              If Left$(EPL$, 6) = "CENTRO" Then VERDEMO% = (-1)
              If InStr(EPL$, "DIPRO") > 0 Then VERDEMO% = (-1)
              If InStr(EPL$, "GRISFER") > 0 Then VERDEMO% = (-1)
          End If
          '
          CLARAN$ = RANDSTRING$(6)
          CONTRA$ = CLARAN$
          If VERDEMO% = (-1) Then
              CONTRA$ = ENCRIPTA$(CLARAN$)
          End If
          '
          Call GRARAM(293, CLARAN$)
          Call GRARAM(299, CONTRA$)
          Call GRARAM(305, TRIM$(Str$(LANGUA%)))
          RSW% = 0
          If CONTROL$(IDENTER$, "RESIWIN") = "SI" Then
            RSW% = 1
          End If
          Call GRARAM(306, TRIM$(Str$(RSW%)))
          '
          Call VERIFLIM(0)
          '
          VDX% = VERDEMO%
          If VDX% <> (-1) Then
            TXT$ = "La Presente Versión FLEXOFT es un\Ejemplar Limitado  para el Fin de\una Demostración.\  \Tiene Restricciones de Uso ya que\"
            TXT$ = TXT$ + "sólo Puede Utilizarse para Evaluar\su Adquisición."
            OPX% = COMALTER%(TXT$ + "\\Continuar\Terminar\Habilitar")
            If OPX% <> 1 Then
                If OPX% = 3 Then
                    OPY% = COMALTER%("Para Habilitar Contacte a\FLEXOFT Soluciones:\  \Telefax: 00-54-11-4541-1515/4524-1284\\Aceptar")
                End If
                Call FUERASYS
                Close: End
            End If
          End If
       End If
       '
       '
       If Timer > TINI + 5 Then
           If CT2% = 0 Then
              FLEXOFTA.MousePointer = 1
              If MULTIEM% = 1 Then
                 ' FORMULTI.Show 1
              End If
              FLESCRIT.Show
              Hide
              CT2% = 1
           End If
        End If
    End If
    
End Sub

Private Sub Timer2_Timer()
    '
    Static vmove, hmove, LLEGO%
    TOPEFIN = 100: IZQFIN = 100
    '
    If vmove = 0 Then
        vmove = (Picture1.Top - TOPEFIN) / 75
    End If
    '
    If hmove = 0 Then
        hmove = (Picture1.Left - IZQFIN) / 75
    End If
    '
    LLEGO% = 1
    If Picture1.Left > 200 Then
       Picture1.Move Picture1.Left - hmove, Picture1.Top
       LLEGO% = 0
    End If
    '
    If Picture1.Top > 200 Then
       Picture1.Move Picture1.Left, Picture1.Top - vmove
       LLEGO% = 0
    End If
    '
    If LLEGO% = 0 Then
        Exit Sub
    End If
    '
    'Frame1.Top = Picture1.Top
    Picture2.Top = Picture1.Top
    'Frame1.Visible = True
    Picture2.Visible = True
    'Frame1.Refresh
    Picture2.Refresh
    Timer2.Enabled = False
'        Call CARDATRAM
'        Label4.Caption = DENOSYS$
'        Label5.Caption = "Serie Nº: " + SERINUM$

End Sub


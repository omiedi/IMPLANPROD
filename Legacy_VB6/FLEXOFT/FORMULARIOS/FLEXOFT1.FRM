VERSION 5.00
Begin VB.Form FORM1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Visual Basic (r) - V.5.0"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   5670
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "4-473-1634"
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
         TabIndex        =   11
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "00-54-11-4-541-1515"
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
         Top             =   2100
         Width           =   2220
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Diaz Colodrero 3519 1431 Buenos Aires Argentina"
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
         TabIndex        =   9
         Top             =   1155
         Width           =   2325
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
         TabIndex        =   8
         Top             =   735
         Width           =   4425
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
         TabIndex        =   7
         Top             =   420
         Width           =   3060
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
         TabIndex        =   6
         Top             =   210
         Width           =   4845
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   1050
         Picture         =   "FLEXOFT1.frx":0000
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1050
         Picture         =   "FLEXOFT1.frx":030A
         Stretch         =   -1  'True
         Top             =   1995
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   1050
         Picture         =   "FLEXOFT1.frx":0614
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   375
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   525
      Top             =   3465
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   2895
      Left            =   210
      Picture         =   "FLEXOFT1.frx":091E
      ScaleHeight     =   2895
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2565.99
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   5055
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
         Left            =   2520
         TabIndex        =   4
         Top             =   105
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
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
      Caption         =   $"FLEXOFT1.frx":502EC
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public termina%


Private Sub Form_Activate()
    Static CTX%
    If CTX% = 0 Then
        Call CARDATRAM
        Form1.Picture1.Top = Form1.ScaleHeight / 2 - Form1.Picture1.Height / 2
        Form1.Picture1.Left = Form1.ScaleWidth / 2 - Form1.Picture1.Width / 2
        Picture1.Visible = True
        Label6.Visible = True
        Timer1.Enabled = True
        CTX% = 1
    End If
    '
    MENU$ = TRIM$(Command$)
    If MENU$ <> "" Then
       Call GRARAM(13, AJUSTI$(MENU$, 8))
    End If
    '
End Sub
'

Private Sub Timer1_Timer()
    Static veces, CT1%, tini, CT2%
    veces = veces + 1
    '
    If CT1% > 0 Then GoTo 10
    '
    If veces >= 12 Then
        Form1.Label1.Visible = True
        Form1.Label2.Visible = True
        Form1.Label3.Visible = True
    End If
    '
    If veces > 22 Then
       Timer2.Enabled = True
       CT1% = 1
    End If
    '
10  If Frame1.Visible = True Then
       If tini = 0 Then
          tini = Timer
          Form1.MousePointer = 11
          Call VERIDIR("C:\FLEXOFT")
          Call VERIPIF
          VDX% = VERDEMO%
          If VDX% <> (-1) Then
            TXT$ = "La Presente Versión FLEXOFT es un\Ejemplar Limitado  para el Fin de\una Demostración.\  \Tiene Restricciones de Uso ya que\"
            TXT$ = TXT$ + "sólo Puede Utilizarse para Evaluar\su Adquisición."
            OPX% = COMALTER%(TXT$ + "\\Continuar\Terminar\Habilitar")
            If OPX% <> 1 Then
                If OPX% = 3 Then
                    OPY% = COMALTER%("Para Habilitar Contacte al\Estudio CIME-SOI:\  \Telefax: 00-54-11-4541-1515\\Aceptar")
                End If
                Call FUERASYS
                Close: End
            End If
          End If
       End If
       '
       '
       If Timer > tini + 5 Then
           If CT2% = 0 Then
              Form1.MousePointer = 1
              If MULTIEM% = 1 Then
                 ' FORMULTI.Show 1
              End If
              FORM2.Show
              CT2% = 1
           End If
        End If
    End If
    
End Sub

Private Sub Timer2_Timer()
    Static vmove, hmove
    TOPEFIN = 100: IZQFIN = 100
    If vmove = 0 Then vmove = (Picture1.Top - TOPEFIN) / 75
    If hmove = 0 Then hmove = (Picture1.Left - IZQFIN) / 75
    termina% = 1
    If Picture1.Left > 200 Then
       Picture1.Move Picture1.Left - hmove, Picture1.Top
       termina% = 0
    End If
    '
    If Picture1.Top > 200 Then
       Picture1.Move Picture1.Left, Picture1.Top - vmove
       termina% = 0
    End If
    '
    If termina% = 1 Then
        Frame1.Top = Picture1.Top
        Frame1.Visible = True
        Label4.Caption = DENOSYS$
        Label5.Caption = "Serie Nº: " + SERINUM$
    End If
End Sub


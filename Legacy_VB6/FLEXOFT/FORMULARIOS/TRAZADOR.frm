VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TRAZADOR 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trazador de Procesos"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   360
      ScaleHeight     =   75
      ScaleWidth      =   4875
      TabIndex        =   19
      Top             =   2400
      Width           =   4935
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   420
      TabIndex        =   12
      Top             =   720
      Width           =   2295
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   1360
         TabIndex        =   18
         Top             =   825
         Width           =   800
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   1360
         TabIndex        =   17
         Top             =   360
         Width           =   800
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   1360
         TabIndex        =   16
         Top             =   580
         Width           =   800
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faltan:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   15
         Top             =   825
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Procesados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   13
         Top             =   580
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tiempos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   2415
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transcurrido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   11
         Top             =   580
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estimado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faltante:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -880
         TabIndex        =   9
         Top             =   825
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "00.00.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1360
         TabIndex        =   8
         Top             =   580
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "00.00.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1360
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "00.00.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1360
         TabIndex        =   6
         Top             =   825
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TRAZADOR.frx":0000
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "TRAZADOR.frx":0234
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   280
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2570
      Width           =   495
   End
End
Attribute VB_Name = "TRAZADOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    Call APLICACIONSKINS(Me)
    '
End Sub

Sub TRACE1(RG%, i&, FIN&)
    '
    Static FINA&, COEF, TANT, VISI%, TINI
    '
    If FIN& <> FINA& Then
        FINA& = FIN&
        If FIN& > 0 Then
            COEF = 1
            COEF = TRAZADOR.MARCOBARRA.Width / FIN&
            TRAZADOR.Label8.Caption = TRIM$(Str$(FIN&))
            TINI = Timer
            TRAZADOR.Label9.Caption = ""
            TRAZADOR.Label10.Caption = ""
            TRAZADOR.Label11.Caption = ""
            TRAZADOR.Label15.Caption = ""
            TRAZADOR.Label16.Caption = ""
            TRAZADOR.Label17.Caption = ""
        End If
        TRAZADOR.BARRATRAZA.Left = 0
    End If
    '
    If FIN& < 1 Then Exit Sub
    '
    If i& >= FIN& Then
        FINA& = 0
        VISI% = 0
        TRAZADOR.Hide
        TRAZADOR.Label3.Caption = TTT$ + " ..."
        Call FRESHALL
        Exit Sub
    End If
    '
    If Int(Timer) <> TANT Then
        TANT = Int(Timer)
        TRACU = Timer - TINI
        TRAZADOR.Label9.Caption = HORMINSE$(TRACU)
        TRAZADOR.Label16.Caption = TRIM$(Str$(FIN&))
        If i& > 0 Then
            TRAZADOR.Label10.Caption = HORMINSE$(TRACU * FIN& / i&)
            TRAZADOR.Label11.Caption = HORMINSE$(TRACU * FIN& / i& - TRACU)
        TRAZADOR.Label15.Caption = TRIM$(Str$(i&))
        TRAZADOR.Label17.Caption = TRIM$(Str$(FIN& - i&))
        End If
        TRAZADOR.Label9.Refresh
        TRAZADOR.Label10.Refresh
        TRAZADOR.Label11.Refresh
        TRAZADOR.Label15.Refresh
        TRAZADOR.Label16.Refresh
        TRAZADOR.Label17.Refresh
    End If
    '
    If VISI% < 1 Then
        TRAZADOR.Show
        TRAZADOR.Refresh
        VISI% = 1
    End If
    TRAZADOR.BARRATRAZA.Left = COEF * i&
    TRAZADOR.MARCOBARRA.Refresh
    TRAZADOR.BARRATRAZA.Refresh
    '
End Sub

Sub BLIMESS(Texto$)
    '
    TTT$ = TRIM$(Texto$)
    If Left$(TTT$, 1) = "-" Then
        TRAZADOR.Label3.Caption = "Proceso Completo !"
      Else
        TRAZADOR.Label3.Caption = TTT$ + " ..."
    End If
    '
End Sub




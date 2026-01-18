VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANEXOCOM 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emisión Nota de Pedido - Documentos Anexos "
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5415
   ControlBox      =   0   'False
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
   ScaleHeight     =   3195
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   4560
      ScaleHeight     =   3750
      ScaleWidth      =   1425
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "ANEXOCOM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "DOCUMENTOS A ANEXAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4320
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Plan de Inspección "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   735
         TabIndex        =   9
         Top             =   525
         Value           =   1  'Checked
         Width           =   3060
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "INSTRUCCIONES - PROCEDIMIENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   210
         TabIndex        =   4
         Top             =   1050
         Width           =   3900
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Instrucciones de Almacenamiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   525
            TabIndex        =   8
            Top             =   420
            Value           =   1  'Checked
            Width           =   3270
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Compras y Abastecimientos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   525
            TabIndex        =   7
            Top             =   735
            Value           =   1  'Checked
            Width           =   3000
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Control de Reposición"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   525
            TabIndex        =   6
            Top             =   1050
            Value           =   1  'Checked
            Width           =   3000
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Instrucciones de Ensayos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   525
            TabIndex        =   5
            Top             =   1365
            Value           =   1  'Checked
            Width           =   3000
         End
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   4620
      Picture         =   "ANEXOCOM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   645
   End
   Begin VB.CommandButton command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   4620
      Picture         =   "ANEXOCOM.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2350
      Width           =   645
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANEXOCOM.frx":091E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "ANEXOCOM.frx":0B52
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label CONTROLABEL 
      Height          =   225
      Left            =   4620
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Menu menu 
      Caption         =   "Menú "
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "ANEXOCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call Command1_Click
End Sub

Private Sub Command1_Click()
   CONTROLABEL.Caption = "OK"
   Screen.MousePointer = 11
   UREPRO& = ULTREG&("CODIPROC")
   '
   On Error Resume Next
   Set FORINSTRUC = INSTRUCCI
   If Err > 0 Then
     On Error GoTo 0 'VER QUE FRM VA A UTILIZAR -- VALIDA SI EXISTE EL FRM INSTRUCCI O INSTRUI07
     Set FORINSTRUC = INSTRUI07
   End If
   On Error GoTo 0
   '
   For J& = 1 To ULTREG&("!OCOMDETA")
     X$ = REGLEIDO$("!OCOMDETA", J&)
     CI0% = CVI(Mid$(X$, 9, 2))
     If CI0% > 0 Then
       If CI0% <= NUMAS% Then
         If Check2.Value = 1 Then
           PLAINSPEC.Label28.Caption = str$(CI0%)
           PLAINSPEC.DOCUORI.Caption = DOCUORIG$ = "N/Ped-" + NUMEDOC$
           PLAINSPEC.COMMLAB.Caption = "PRINT"
         End If
         '
         YACAPUN% = 0
         YY$ = String$(128, 0)
         For KKI% = 0 To 3
           If Check1(KKI%).Value = 1 Then
             '
             If YACAPUN% = 0 Then
               RFF$ = RECFILT$("CODIPROC", 1, MKI$(CI0%))
               If Len(RFF$) >= 4 Then
                 RECODIPRO& = CVS(Left$(RFF$, 4))
                 If RECODIPRO& > 0 Then
                   If RECODIPRO& <= UREPRO& Then
                     YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
                   End If
                 End If
               End If
               YACAPUN% = 1
             End If
             '
             CODINSTRU% = CVI(Mid$(YY$, 3 + 2 * KKI%, 2))
             If CODINSTRU% > 0 Then
               FORINSTRUC.Label28.Caption = str$(CODINSTRU%)
               FORINSTRUC.DOCUORI.Caption = DOCUORIG$ = "N/Ped-" + NUMEDOC$
               FORINSTRUC.COMMLAB.Caption = "PRINT"
             End If
           End If
         Next KKI%
       End If
     End If
   Next J&
   '
   Screen.MousePointer = 1
   Hide
   '
End Sub

Private Sub Command2_Click()
   Call Command1_Click
End Sub

Private Sub Command7_Click()
   CONTROLABEL.Caption = ""
   Hide
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub salir_Click()
    Call Command7_Click
End Sub

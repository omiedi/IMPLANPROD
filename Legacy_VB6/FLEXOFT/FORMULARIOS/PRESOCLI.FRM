VERSION 5.00
Begin VB.Form PRESOCLI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Material de Texto Libre"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Documentación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   4830
      TabIndex        =   20
      Top             =   2415
      Width           =   2640
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
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
         Index           =   4
         Left            =   -1050
         TabIndex        =   31
         Top             =   795
         Width           =   1920
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   30
         Top             =   735
         Width           =   1485
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
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
         Index           =   3
         Left            =   -1050
         TabIndex        =   29
         Top             =   375
         Width           =   1920
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   21
         Top             =   315
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   6945
      Left            =   7665
      ScaleHeight     =   6885
      ScaleWidth      =   1005
      TabIndex        =   13
      Top             =   -120
      Width           =   1065
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "PRESOCLI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   315
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "PRESOCLI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Aprueba la Selección de Items a Comprar"
         Top             =   2100
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   2940
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "PRESOCLI.frx":0614
         Stretch         =   -1  'True
         Top             =   3255
         Width           =   1410
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   210
      TabIndex        =   10
      Top             =   2415
      Width           =   4485
      Begin VB.OptionButton Option2 
         Caption         =   "Rechazado"
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
         Left            =   2730
         TabIndex        =   36
         Top             =   780
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aprobado"
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
         Left            =   1260
         TabIndex        =   35
         Top             =   780
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.Frame Frame4 
         Caption         =   "Precio Unit."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2310
         TabIndex        =   32
         Top             =   0
         Width           =   2170
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   22
            Top             =   250
            Width           =   1170
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$"
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
            Left            =   -210
            TabIndex        =   33
            Top             =   290
            Width           =   960
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2310
         Y1              =   685
         Y2              =   685
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   4515
         X2              =   0
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Control:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -840
         TabIndex        =   34
         Top             =   820
         Width           =   1920
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente:"
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
         Index           =   0
         Left            =   -420
         TabIndex        =   12
         Top             =   315
         Width           =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1575
         TabIndex        =   11
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   7215
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         Left            =   -210
         TabIndex        =   28
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   420
         TabIndex        =   27
         Top             =   570
         Width           =   1920
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2415
         TabIndex        =   26
         Top             =   255
         Width           =   4650
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
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
         Left            =   1365
         TabIndex        =   9
         Top             =   250
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   7215
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº B/R:"
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
         Left            =   2520
         TabIndex        =   25
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3570
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1140
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/C:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5955
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Left            =   4920
         TabIndex        =   5
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1140
         TabIndex        =   3
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2295
         TabIndex        =   1
         Top             =   615
         Width           =   4755
      End
   End
End
Attribute VB_Name = "PRESOCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    'ALTA EN BOLRECEP
    NC% = Val(Label5)
    NC1% = (-1) * NC%
    NUOC& = Val(Label4)
    CAPROB# = XVALO(Label7): CRECHA# = 0
    If Option2.Value = True Then
       CAPROB# = 0: CRECHA# = XVALO(Label7)
    End If
    NBORE& = Val(Label12)
    REMI$ = TRIM$(Label22)
    NFAC$ = TRIM$(Label18)
    FRECEP% = FECHANUM(Label2)
    CIXI% = CODINT%(Label16)
    MONEV$ = TRIM$(Label19)
    PRUNI = Val(Text1)
    '
    TTXX$ = REGBLAN$("BOLRECEP")
    Call REPLA(TTXX$, MKS$(NBORE&), 1, 4)
    Call REPLA(TTXX$, Label15, 5, 64)
    Call REPLA(TTXX$, MKI$(FRECEP%), 81, 2)
    Call REPLA(TTXX$, MKI$(CIXI%), 83, 2)
    Call REPLA(TTXX$, "U.", 85, 4)
    Call REPLA(TTXX$, MKD$(CAPROB# + CRECHA#), 89, 8)
    Call REPLA(TTXX$, MKI$(NC%), 105, 2)
    Call REPLA(TTXX$, REMI$, 119, 16)
    Call REPLA(TTXX$, NFAC$, 135, 16)
    Call REPLA(TTXX$, MKS$(NUOC&), 151, 4)
    Call REPLA(TTXX$, MONEV$, 155, 4)           ' MONEDA VALOR UNITARIO TIPO 6 EN 155,4
    Call REPLA(TTXX$, MKS$(PRUNI), 159, 4)      ' VALOR UNITARIO TIPO 3 EN 159,4
    Call REPLA(TTXX$, MKD$(CAPROB#), 163, 8)
    Call REPLA(TTXX$, MKD$(CRECHA#), 171, 8)
    Call REPLA(TTXX$, MKI$(FRECEP%), 187, 2)
    Call REGAPP("BOLRECEP", TTXX$)
    Call CIERRARCH("BOLRECEP")
    '
    ZZ$ = FILTERGETRECORD$("OCOMDETA", 1, MKS$(NUOC&))
    CATORE = CAPROB#
    Call REPLA(ZZ$, MKS$(CATORE), 99, 4)
    CAPENDIEN = CVS(Mid$(ZZ$, 55, 4)) - CVS(Mid$(ZZ$, 99, 4))
    If CAPENDIEN < 0 Then CAPENDIEN = 0
    If CAPENDIEN <= 0 Then Call REPLA(ZZ$, Chr$(3), 124, 1)
    Call FILTERPUTRECORD("OCOMDETA", 1, MKS$(NUOC&), ZZ$)
    Call CIERRARCH("OCOMDETA")
    '
    ZZ$ = FILTERGETRECORD$("TXOCOLIB", 1, MKS$(NUOC&))
    Call REPLA(ZZ$, MKS$(NBORE&), 69, 4)
    Call FILTERPUTRECORD("TXOCOLIB", 1, MKS$(NUOC&), ZZ$)
    Call CIERRARCH("TXOCOLIB")
    '
    X$ = REGBLAN$("INDIBOL")
    Call REPLA(X$, MKS$(NBORE&), 1, 4)
    FEX = FRECEP%
    Call REPLA(X$, FECHATEX$(FEX), 5, 8)
    Call REPLA(X$, REMI$, 15, 16)
    Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
    Call REGAPP("INDIBOL", X$)
    '
    DOPRI$ = "BR." + TRIM$(Str$(NBORE&))
    
DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
    CANTI = CAPROB# + CRECHA#
    REFE$ = "Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
    VUNI# = XVALO(Text1)
    MONE$ = TRIM$(Label19)
    NC1% = (-1) * NC%
    DOSEC$ = TRIM$(REMI$)
       While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
       If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
    Call MOVISTO(FRECEP%, CIXI%, "", "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
    '
    If CRECHA# > 0 Then
      CANTI = (-1) * CRECHA#
      REFE$ = "Rechazo Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
      VUNI# = XVALO(Text1)
      MONE$ = TRIM$(Label19)
      NC1% = (-1) * NC%
      DOSEC$ = TRIM$(REMI$)
         While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
         If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
      Call MOVISTO(FRECEP%, CIXI%, "", "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
    End If
    '
    X$ = REGLEIDO$("TATIREC", 1)
    FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
    If TRIM$(FORIPRE$) <> "" Then
       '
       CANCELPRINT% = 0
       If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
       '
       TIPODOC$ = "Boleta de Recepción"
       DESTIDOC% = (-1) * NC%
       NUMEDOC$ = Str$(NBORE&)
       FECHDOC$ = Label2
       '
       CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
       CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
       CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
       CAPARDOC% = 3
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "COLOR"
       CODTABU1$(4) = "UNIMED"
       CODTABU1$(5) = "NRO-LOTE"
       CODTABU1$(6) = "CANTIDAD"
       CODTABU1$(7) = "N-OCOMPR"
       CODTABU1$(8) = "CAKILOS"
       CODTABU1$(9) = "F-PLANO"
       CODTABU1$(10) = "FECH-REC"
       '\\\OT-05-0543-WAR-26/07/05///
       CODTABU1$(11) = "PAR-MPRI"
       '\\\OT-05-0543-WAR-FIN///
       CACOLTAB1% = 11
       '
       CAITAB1% = 1
       YY$ = TTXX$
       TETABU1$(1, CAITAB1%) = Label16
          CI% = CIXI%
          UNID$ = TRIM$(Mid$(YY$, 85, 4))
          If UNID$ = "" Then UNID$ = UNIMED$(CI%)
          CAN# = CAPROB#
          CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
          KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
          FEX = CVI(Mid$(YY$, 155, 2))
          FESOLEN$ = FECHATEX$(FEX)
          FEPLA$ = FEPLANO$(CI%)
                      'NULOTE$ = TRIM$(Mid$(YY$, 107, 12)) '¿Esto estaBA BieN?
          NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
          If NULOTE$ = "" Then
               NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
          End If
          NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
          LOTPROV$ = TRIM$(Mid$(YY$, 107, 12))
          DESMAT$ = Label15
       TETABU1$(2, CAITAB1%) = DESMAT$
       TETABU1$(3, CAITAB1%) = COLOR$(CI%)
       TETABU1$(4, CAITAB1%) = UNID$
       TETABU1$(5, CAITAB1%) = NULOTE$
       TETABU1$(6, CAITAB1%) = CANS$
       TETABU1$(7, CAITAB1%) = NUOCO$
       TETABU1$(8, CAITAB1%) = KILAJE$
       TETABU1$(9, CAITAB1%) = FEPLA$
       TETABU1$(10, CAITAB1%) = FESOLEN$
       TETABU1$(11, CAITAB1%) = LOTPROV$
       '
       Call PRINTDOC("FORMBREC")
       '
    End If
    '
    Call COMUNI("Recepcion Procesada con Éxito.")
90  Unload Me
    '
99  Command2.Enabled = True
    '
End Sub


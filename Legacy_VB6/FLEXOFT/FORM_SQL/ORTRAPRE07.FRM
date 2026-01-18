VERSION 5.00
Begin VB.Form ORTRAPRE07 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Orden de Trabajo Preventivo"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ordenes de Trabajo Preventivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3945
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   2325
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "ORTRAPRE07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Material en Consignación"
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
            TabIndex        =   10
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
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
            Picture         =   "ORTRAPRE07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Pendientes por Material"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Material"
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
            TabIndex        =   7
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H80000016&
         Height          =   650
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2625
         TabIndex        =   3
         Top             =   360
         Width           =   2685
         Begin VB.CommandButton Command15 
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
            Left            =   0
            Picture         =   "ORTRAPRE07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Orden de Trabajo Preventivo"
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
            Left            =   720
            TabIndex        =   4
            Top             =   60
            Width           =   1815
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   5055
      Left            =   4200
      ScaleHeight     =   4995
      ScaleWidth      =   1125
      TabIndex        =   0
      Top             =   0
      Width           =   1185
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   100
         Picture         =   "ORTRAPRE07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   160
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "ORTRAPRE07.frx":0D60
         Top             =   960
         Width           =   1515
      End
   End
End
Attribute VB_Name = "ORTRAPRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
Command14.Enabled = False
    STPMANPRE07.Show 1
Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    '
    Call FRESHECHO("ABMSIS")
    Call SELECHO("ABMSIS")
    LINESIS$ = TRIM$(VALACT1$("ABMSIS"))
    If LINESIS$ = "" Then Exit Sub
    '
    HOII% = HOY(HOS$)
    FECHDOC$ = HOS$
    NUMEDOC$ = ""
    '
    For II& = 1 To ULTREG&("ABMSIS")
        XX$ = REGLEIDO$("ABMSIS", II&)
        If TRIM$(Mid$(XX$, 1, 64)) = LINESIS$ Then
            CRITICI$ = TRIM$(Mid$(XX$, 65, 32))
            REVISIO$ = TRIM$(Mid$(XX$, 97, 16))
            NUMEPOE$ = TRIM$(Mid$(XX$, 113, 32))
            Exit For
        End If
    Next II&
    '
    CONDI$ = "LINEA_EQU ='" & LINESIS$ & "'"
    PLANT$ = VALOBADA("EQUIPOS", "PLANTA_EQU", CONDI$, "NOMESS")
    CODPARAM$(1) = "MEDIDA-1": DOCPARAM$(1) = PLANT$ & " - " & TRIM$(ECOARCH$("PLANTPRO", PLANT$))         'EDIFICIO
    CODPARAM$(2) = "MEDIDA-2": DOCPARAM$(2) = LINESIS$  'SISTEMA
    CODPARAM$(3) = "Z-PLANO": DOCPARAM$(3) = CRITICI$   'CRITICIDAD
    CODPARAM$(4) = "DOCUORIG": DOCPARAM$(4) = NUMEPOE$  'POE
    CODPARAM$(5) = "R-PLANO": DOCPARAM$(5) = REVISIO$   'REVISION
    CODPARAM$(6) = "EQUIPO-1": DOCPARAM$(6) = ""        'EQUIPO
    CAPARDOC% = 6
    '
    CODTABU1$(1) = "REF-MAT1"
    CODTABU1$(2) = "REF-MAT2"
    CODTABU1$(3) = "REF-MAT3"
    CODTABU1$(4) = "REF-MAT4"
    CODTABU1$(5) = "REF-MAT5"
    CODTABU1$(6) = "REF-MAT6"
    CACOLTAB1% = 6
    DESTIDOC% = 0
    CAITAB1% = 0
    CAITCRU% = 0
    '
    Dim LINEQUTMP As ADODB.Recordset
    SQLX$ = " SELECT * FROM EQUIPOS "
    SQLX$ = SQLX$ + "where LINEA_EQU = '" & TRIM$(LINESIS$) & "'"
    Set LINEQUTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    '
    With LINEQUTMP
       Do While Not .EOF
          EQUIPP$ = SAFETEXT$(!codigo_equ)
          DEQUIPP$ = SAFETEXT$(!Descri_Equ)
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = String$(32, "-")
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = DEQUIPP$
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = ""
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          '
          Dim EQUIPTMP As ADODB.Recordset
          SQLX$ = " SELECT * FROM SUBCONJUNTOS "
          SQLX$ = SQLX$ + "where CODIGO_EQU = '" & EQUIPP$ & "'"
          SQLX$ = SQLX$ + "ORDER BY NRO_SUB"
          Set EQUIPTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
          '
          With EQUIPTMP
             Do While Not EQUIPTMP.EOF
                NUSUBCONJ = XVALO(!NRO_SUB)
                DESUBCONJ$ = SAFETEXT(!DESCRI_SUB)
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = String$(32, "-")
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = ""
                TETABU1$(2, CAITAB1%) = DESUBCONJ$
                TETABU1$(3, CAITAB1%) = ""
                TETABU1$(4, CAITAB1%) = ""
                TETABU1$(5, CAITAB1%) = ""
                TETABU1$(6, CAITAB1%) = ""
                '
                Dim SUBCOTMP As ADODB.Recordset
                SQLX$ = " SELECT * FROM PLAMANPREV "
                SQLX$ = SQLX$ + "where NRO_SUB = " & NUSUBCONJ
                SQLX$ = SQLX$ + " AND CODIGO_EQU = '" & EQUIPP$ & "'"
                SQLX$ = SQLX$ + "ORDER BY orden_tar"
                Set SUBCOTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
                '
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = String$(32, "-")
                '
                With SUBCOTMP
                  Do While Not SUBCOTMP.EOF
                    ORDENTAR = XVALO(!orden_tar)
                    DESCRITAR$ = SAFETEXT(!descri_tar)
                    FRECUTAR$ = XVALO(!Frecu_tar)
                    NUIFRECU = SAFETEXT(!Uni_Frecu)
                    OBSERTAR$ = SAFETEXT(!Obse_Tar)
                    '
                    CAITAB1% = CAITAB1% + 1
                    TETABU1$(1, CAITAB1%) = ""
                    TETABU1$(2, CAITAB1%) = ""
                    TETABU1$(3, CAITAB1%) = DESCRITAR$
                    TETABU1$(4, CAITAB1%) = FRECUTAR$
                    TETABU1$(5, CAITAB1%) = NUIFRECU
                    TETABU1$(6, CAITAB1%) = OBSERTAR$
                    '
                  SUBCOTMP.MoveNext
                  Loop
                End With
                SUBCOTMP.Close
                Set SUBCOTMP = Nothing
             .MoveNext
             Loop
          End With
          EQUIPTMP.Close
          Set EQUIPTMP = Nothing
          '
       .MoveNext
       Loop
    End With
    '
    CAITAB1% = CAITAB1% + 1
    TETABU1$(1, CAITAB1%) = String$(32, "-")
    '
    LINEQUTMP.Close
    Set LINEQUTMP = Nothing
    '
    Call PRINTDOC("FOR-PMP")   ' PLAN MANTENIMIENTO PREVENTIVO
    '
End Sub

Private Sub Command24_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

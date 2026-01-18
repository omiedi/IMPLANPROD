VERSION 5.00
Begin VB.Form PRODUCT07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Productividad de Mano de Obra"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   8205
      Left            =   11025
      ScaleHeight     =   8145
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   -315
      Width           =   1170
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
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
         Top             =   630
         Width           =   640
      End
      Begin VB.CommandButton Command29 
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
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1320
         Width           =   640
      End
      Begin VB.CommandButton Command19 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceso a Opciones Complementarias"
         Top             =   5730
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   4305
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Listado por Rango de Fechas"
         Top             =   3105
         Width           =   640
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   5040
         Width           =   640
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PRODUCT07.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Impresión - Vuelco de Pantalla"
         Top             =   2415
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "PRODUCT07.frx":1514
         Top             =   7455
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   2620
      Width           =   10550
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Resumen de Productividad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   10
      Top             =   210
      Width           =   6840
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         Index           =   7
         Left            =   4515
         TabIndex        =   26
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   5565
         TabIndex        =   25
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rendim. General: "
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
         Index           =   6
         Left            =   4935
         TabIndex        =   24
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5775
         TabIndex        =   23
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   3620
         TabIndex        =   22
         Top             =   1260
         Width           =   1250
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   3620
         TabIndex        =   21
         Top             =   840
         Width           =   1250
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Difer.en Tiempo:"
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
         Index           =   5
         Left            =   1350
         TabIndex        =   20
         Top             =   1320
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Tot.Empl:"
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
         Index           =   4
         Left            =   1350
         TabIndex        =   19
         Top             =   900
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Tot.Estim:"
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
         Index           =   3
         Left            =   1350
         TabIndex        =   18
         Top             =   480
         Width           =   2220
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   3620
         TabIndex        =   17
         Top             =   420
         Width           =   1250
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   1260
         TabIndex        =   16
         Top             =   1260
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Negativas:"
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
         Index           =   2
         Left            =   -1050
         TabIndex        =   15
         Top             =   1320
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   1260
         TabIndex        =   14
         Top             =   840
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Positivas:"
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
         Index           =   1
         Left            =   -1050
         TabIndex        =   13
         Top             =   900
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   0
         Left            =   1260
         TabIndex        =   12
         Top             =   420
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Ordenes:"
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
         Left            =   -1050
         TabIndex        =   11
         Top             =   480
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   120
      Left            =   210
      TabIndex        =   9
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ultimo Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7035
      TabIndex        =   0
      Top             =   210
      Width           =   3620
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2730
         Picture         =   "PRODUCT07.frx":3436
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Re-Cálculo de Resumen Mensual para el Mes"
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Left            =   1680
         TabIndex        =   8
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   315
         TabIndex        =   7
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Left            =   315
         TabIndex        =   6
         Top             =   630
         Width           =   2115
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         Left            =   1420
         TabIndex        =   4
         Top             =   1220
         Width           =   225
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Left            =   2460
         TabIndex        =   3
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Left            =   105
         TabIndex        =   2
         Top             =   990
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   105
         TabIndex        =   1
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   5145
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   12000
      Y1              =   7630
      Y2              =   7630
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operario            Cat. Ords  Pos  Neg    Estimado    Empleado        Total  Rendim"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   2360
      Width           =   10410
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                        Total  Ord. Ord.  Tiemp.Tot.  Tiemp.Tot.     Diferen."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   28
      Top             =   2100
      Width           =   10550
   End
End
Attribute VB_Name = "PRODUCT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPRODUCT.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   OPCIPRO07.Show 1
   Command19.Enabled = True
End Sub


Private Sub Command2_Click()
   '
   Command2.Enabled = False
   FORIPRE$ = TRIM$(CONTROL$("", "FORPROGE"))
   If FORIPRE$ = "" Then
        PRINTFORM
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Productividad General"
        '
        CODPARAM$(1) = "FECH-LAN"
        CODPARAM$(2) = "FECH-VEN"
        CODPARAM$(3) = "CUEN-CLI"
        CODPARAM$(4) = "RASO-CLI"
        CODPARAM$(5) = "CIVA-CLI"
        CODPARAM$(6) = "PRE-LIST"
        CODPARAM$(7) = "TOTA-CAN"
        CODPARAM$(8) = "CANT-DIS"
        CODPARAM$(9) = "CANT-FAL"
        CODPARAM$(10) = "TIE-PREP"
        CODPARAM$(11) = "TIE-PROD"
        CODPARAM$(12) = "TIE-TOTA"
        CODPARAM$(13) = "CAKILOS"
        CAPARDOC% = 13
        '
        DOCPARAM$(1) = TRIM$(Left$(Label1.Caption, 8))
        DOCPARAM$(2) = TRIM$(Mid$(Label1.Caption, 12, 8))
        TTX0$ = PRODUCTI.List1.TEXT
        DOCPARAM$(3) = ""
        DOCPARAM$(4) = ""
        DOCPARAM$(5) = ""
        DOCPARAM$(6) = ""
        DOCPARAM$(7) = TRIM$(Label6(0).Caption)
        DOCPARAM$(8) = TRIM$(Label6(1).Caption)
        DOCPARAM$(9) = TRIM$(Label6(2).Caption)
        DOCPARAM$(10) = AJUSTD$(Label7(0).Caption, 12)
        DOCPARAM$(11) = AJUSTD$(Label7(1).Caption, 12)
        DOCPARAM$(12) = AJUSTD$(Label7(2).Caption, 12)
        DOCPARAM$(13) = TRIM$(Label8.Caption)
        '
        CODTABU1$(1) = "REF-MAT1"
        CACOLTAB1% = 1
        '
        CAITAB1% = 0
        For U& = 1 To List1.ListCount
           Z$ = List1.List(U& - 1)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = Z$
        Next U&
        '
        Call PRINTDOC("FORPROGE")
        Screen.MousePointer = 1
        '
      End If
      '
   Command2.Enabled = True
   '
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("EFIMOBRA")
End Sub


Private Sub Command4_Click()
    '
    Command4.Enabled = False
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Dim LEGA%(512)
    Dim NOME$(512)
    Dim CATE%(512)
    Dim TORD&(512)
    Dim ORPO&(512)
    Dim ORNE&(512)
    Dim TISTA(512)
    Dim TIEMP(512)
    '
    Call BLOQARCH("COMOBRA")
    Call BLOQARCH("RECOMOB")
    '
    For U& = 1 To ULTREG&("PADRON")
      XX$ = REGLEIDO$("PADRON", U&)
      LEGA%(U&) = CVI(Left$(XX$, 2))
      NOME$(U&) = Mid$(XX$, 3, 30)
      CATE%(U&) = Val(Mid$(XX$, 54, 1))
      CAOPE% = U&
      If U& >= 512 Then GoTo 20
    Next U&
    '
20      JJ& = 0: SUVADIF = 0
    'Nuevo
'    DES1# = CDbl(CVDAT(DES%))
'    HAS1# = CDbl(CVDAT(HAS%))
    DESD1$ = FECHATEX$(DES%)
    HAST1$ = FECHATEX$(DIPOST(HAS%))
    SQLX$ = "SELECT * FROM INFOFAB "
    SQLX$ = SQLX$ + " WHERE FeHorLib >= '" & DESD1$ & "' "
    SQLX$ = SQLX$ + " AND FeHorLib < '" & HAST1$ & "' "
    Call ABREORFAB
    'Set OPERTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim OPERTMP As ADODB.Recordset
    Set OPERTMP = New ADODB.Recordset
    OPERTMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With OPERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
         On Error GoTo 0
         Call COMUNI("No se ha Podido Generar el Informe")
         GoTo 99
      End If
      'On Error GoTo 0
      '
      Do While .EOF = False
        LEGOPER$ = !LegaOps
        MAQUI$ = !Maquina
        FEHORLIB1$ = TRIM$(Mid(!fehorlib, 11, 29))
        HFIN% = HORANUM(FEHORLIB1$)
        NOFA$ = !IdSubOr
        CIXI% = !Cod_Inte
        NOPE% = !NumeOper
        DOPE$ = !DescOper
        CAPED# = !CanOrPro
        CAREA# = !CantReal
        TREAL = 60 * !TiemEjec
        TSTAN = 60 * !TiemStan
        RENDIM = 999: If TREAL > 0.01 Then RENDIM = 100 * TSTAN / TREAL
        If RENDIM > 999 Then RENDIM = 999
          '
        ZZ1$ = REGBLAN$("COMOBRA")
        Call REPLA(ZZ1$, NOFA$, 3, 12)
        Call REPLA(ZZ1$, MKI$(CIXI%), 15, 2)
        Call REPLA(ZZ1$, MKI$(NOPE%), 17, 2)
        Call REPLA(ZZ1$, DOPE$, 19, 48)
        Call REPLA(ZZ1$, MKD$(CAREA#), 67, 8)
        Call REPLA(ZZ1$, MKI$(CVINT(!fehorlib)), 75, 2)
        Call REPLA(ZZ1$, MKI$(HFIN%), 77, 2)
        Call REPLA(ZZ1$, MKS$(TREAL), 81, 4)
        Call REPLA(ZZ1$, MKS$(TSTAN), 85, 4)
        Call REPLA(ZZ1$, MKS$(RENDIM), 89, 4)
        Call REPLA(ZZ1$, MAQUI$, 101, 6)
        Call REPLA(ZZ1$, MKD$(CAPED#), 107, 8)
        '
        For KK% = 1 To Len(LEGOPER$) Step 2
          LEGX% = CVI(Mid$(LEGOPER$, KK%, 2))
          If LEGX% > 0 Then
            ZZ$ = ZZ1$
            Call REPLA(ZZ$, MKI$(LEGX%), 1, 2)
            Call REPLA(ZZ$, MKI$(LEGX%), 79, 2)
            JJ& = JJ& + 1
            Call GRAREG("COMOBRA", ZZ$, JJ&)
            '
            For MM% = 1 To CAOPE%
              If LEGA%(MM%) = LEGX% Then
                TORD&(MM%) = TORD&(MM%) + 1
                If RENDIM >= 100 Then
                    ORPO&(MM%) = ORPO&(MM%) + 1
                  Else
                    ORNE&(MM%) = ORNE&(MM%) + 1
                End If
                TISTA(MM%) = TISTA(MM%) + TSTAN
                TIEMP(MM%) = TIEMP(MM%) + TREAL
                CATOPX% = CATE%(MM%)
                VALODIF = VALHORA(CATOPX%) * (TSTAN - TREAL) / 3600
                Call REPLA(ZZ$, MKS$(VALODIF), 115, 4)
                Call GRAREG("COMOBRA", ZZ$, JJ&)
                SUVADIF = SUVADIF + VALODIF
                GoTo 39
              End If
            Next MM%
          End If
39      Next KK%
       .MoveNext
      Loop
    End With
    On Error GoTo 0
    Set OPERTMP = Nothing
    Call SETULTREG("COMOBRA", JJ&)
    Call CIERRARCH("COMOBRA")

    '///***ORIGINAL***///
'20  JJ& = 0: SUVADIF = 0
'    FIN& = ULTREG&("OPERFAB")
'    For U& = 1 To FIN&
'       Call TRACE(20, U&, FIN&)
'       X$ = REGLEIDO$("OPERFAB", U&)
'       FFIN% = CVI(Mid$(X$, 157, 2))
'       If FFIN% >= DES% Then
'         If FFIN% <= HAS% Then
'           LEGOPER$ = Mid$(X$, 165, 8)
'           MAQUI$ = Mid$(X$, 151, 6)
'           HFIN% = CVI(Mid$(X$, 159, 2))
'           NOFA$ = Mid$(X$, 1, 12)
'           CIXI% = CVI(Mid$(X$, 13, 2))
'           NOPE% = CVI(Mid$(X$, 15, 2))
'           DOPE$ = Mid$(X$, 17, 48)
'           CAPED# = CVD(Mid$(X$, 79, 8))
'           CAREA# = CVD(Mid$(X$, 87, 8))
'           TREAL = 60 * CVS(Mid$(X$, 161, 4)) ' PASA A SEGUNDOS
'           TSTAN = 60 * CVS(Mid$(X$, 173, 4)) ' PASA A SEGUNDOS
'           RENDIM = 999: If TREAL > 0.01 Then RENDIM = 100 * TSTAN / TREAL
'           If RENDIM > 999 Then RENDIM = 999
'           '
'           ZZ1$ = REGBLAN$("COMOBRA")
'           Call REPLA(ZZ1$, NOFA$, 3, 12)
'           Call REPLA(ZZ1$, MKI$(CIXI%), 15, 2)
'           Call REPLA(ZZ1$, MKI$(NOPE%), 17, 2)
'           Call REPLA(ZZ1$, DOPE$, 19, 48)
'           Call REPLA(ZZ1$, MKD$(CAREA#), 67, 8)
'           Call REPLA(ZZ1$, MKI$(FFIN%), 75, 2)
'           Call REPLA(ZZ1$, MKI$(HFIN%), 77, 2)
'           Call REPLA(ZZ1$, MKS$(TREAL), 81, 4)
'           Call REPLA(ZZ1$, MKS$(TSTAN), 85, 4)
'           Call REPLA(ZZ1$, MKS$(RENDIM), 89, 4)
'           Call REPLA(ZZ1$, MAQUI$, 101, 6)
'           Call REPLA(ZZ1$, MKD$(CAPED#), 107, 8)
'           '
'           For KK% = 1 To Len(LEGOPER$) Step 2
'             LEGX% = CVI(Mid$(LEGOPER$, KK%, 2))
'             If LEGX% > 0 Then
'               ZZ$ = ZZ1$
'               Call REPLA(ZZ$, MKI$(LEGX%), 1, 2)
'               Call REPLA(ZZ$, MKI$(LEGX%), 79, 2)
'               JJ& = JJ& + 1
'               Call GRAREG("COMOBRA", ZZ$, JJ&)
'               '
'               For MM% = 1 To CAOPE%
'                 If LEGA%(MM%) = LEGX% Then
'                   TORD&(MM%) = TORD&(MM%) + 1
'                   If RENDIM >= 100 Then
'                       ORPO&(MM%) = ORPO&(MM%) + 1
'                     Else
'                       ORNE&(MM%) = ORNE&(MM%) + 1
'                   End If
'                   TISTA(MM%) = TISTA(MM%) + TSTAN
'                   TIEMP(MM%) = TIEMP(MM%) + TREAL
'                   CATOPX% = CATE%(MM%)
'                   VALODIF = VALHORA(CATOPX%) * (TSTAN - TREAL) / 3600
'                   Call REPLA(ZZ$, MKS$(VALODIF), 115, 4)
'                   Call GRAREG("COMOBRA", ZZ$, JJ&)
'                   SUVADIF = SUVADIF + VALODIF
'                   GoTo 39
'                 End If
'               Next MM%
'             End If
'39         Next KK%
'         End If
'       End If
'    Next U&
'    '
'    Call SETULTREG("COMOBRA", JJ&)
'    Call CIERRARCH("COMOBRA")
'    '
    List1.Clear
    JJ& = 0
    For U& = 1 To CAOPE%
      '
      TTXX$ = AJUSTI$(NOME$(U&), 20)
      If TRIM$(TTXX$) = "" Then TTXX$ = AJUSTI$("Legajo Nro." + TRIM$(Str$(LEGA%(U&))), 20)
      TTXX$ = TTXX$ + "  " + AJUSTI$(Str$(CATE%(U&)), 1)
      TTXX$ = TTXX$ + CSTRING$(MKS$(TORD&(U&)), 3, 6, 0, 2)
      TTXX$ = TTXX$ + CSTRING$(MKS$(ORPO&(U&)), 3, 5, 0, 2)
      TTXX$ = TTXX$ + CSTRING$(MKS$(ORNE&(U&)), 3, 5, 0, 2)
      TTXX$ = TTXX$ + AJUSTD$(HORMINSE$(TISTA(U&)), 12)
      TTXX$ = TTXX$ + AJUSTD$(HORMINSE$(TIEMP(U&)), 12)
      DIFETI = TISTA(U&) - TIEMP(U&)
      DITI$ = AJUSTD$(HORMINSE$(Abs(DIFETI)), 13)
      If DIFETI < 0 Then Call REPLA(DITI$, "-", 3, 1)
      TTXX$ = TTXX$ + DITI$
      RENDIM = 999: If TIEMP(U&) > 0.01 Then RENDIM = 100 * TISTA(U&) / TIEMP(U&)
      If RENDIM > 999 Then RENDIM = 999
      RDXX$ = CSTRING$(MKS$(RENDIM), 3, 8, 1, 2)
      If TIEMP(U&) < 1 Then RDXX$ = Space$(8)
      TTXX$ = TTXX$ + RDXX$
      TTXX$ = Left$(TTXX$ + Space$(122), 122)
      TTXX$ = TTXX$ + CSTRING$(MKI$(LEGA%(U&)), 1, 6, 0, 2)
      List1.AddItem TTXX$
      JJ& = JJ& + 1
      Call GRAREG("RECOMOB", TTXX$, JJ&)
      '
      TORD&(0) = TORD&(0) + TORD&(U&)
      ORPO&(0) = ORPO&(0) + ORPO&(U&)
      ORNE&(0) = ORNE&(0) + ORNE&(U&)
      TISTA(0) = TISTA(0) + TISTA(U&)
      TIEMP(0) = TIEMP(0) + TIEMP(U&)
      '
    Next U&
    List1.Refresh
    Label6(0).Caption = TRIM$(Str$(TORD&(0)))
    Label6(1).Caption = TRIM$(Str$(ORPO&(0)))
    Label6(2).Caption = TRIM$(Str$(ORNE&(0)))
    Label7(0).Caption = TRIM$(HORMINSE$(TISTA(0)))
    Label7(1).Caption = TRIM$(HORMINSE$(TIEMP(0)))
    DIFETI = TISTA(0) - TIEMP(0)
    DITI$ = TRIM$(HORMINSE$(Abs(DIFETI)))
    Label7(2).ForeColor = RGB(0, 0, 0)
    If DIFETI < 0 Then Label7(2).ForeColor = RGB(255, 0, 0)
    Label7(2).Caption = DITI$
    RENDIM = 999: If TIEMP(0) > 0.01 Then RENDIM = 100 * TISTA(0) / TIEMP(0)
    If RENDIM > 999 Then RENDIM = 999
    Label8.ForeColor = RGB(0, 0, 0)
    If RENDIM < 100 Then Label8.ForeColor = RGB(255, 0, 0)
    Label8.Caption = TRIM$(CSTRING$(MKS$(RENDIM), 3, 8, 1, 2))
    Label9.ForeColor = RGB(0, 0, 0)
    If SUVADIF < 0 Then Label9.ForeColor = RGB(255, 0, 0)
    Label9.Caption = TRIM$(CSTRING$(MKS$(Abs(SUVADIF)), 3, 10, 2, 2))
    '
    TTX0$ = Space$(128)
    Call REPLA(TTX0$, TRIM$(Str$(TORD&(0))), 1, 6)
    Call REPLA(TTX0$, TRIM$(Str$(ORPO&(0))), 7, 6)
    Call REPLA(TTX0$, TRIM$(Str$(ORNE&(0))), 13, 6)
    Call REPLA(TTX0$, CSTRING$(MKS$(TISTA(0)), 4, 12, 1, 2), 19, 12)
    Call REPLA(TTX0$, CSTRING$(MKS$(TIEMP(0)), 4, 12, 1, 2), 31, 12)
    Call REPLA(TTX0$, CSTRING$(MKS$(SUVADIF), 4, 12, 1, 2), 93, 12)
    '
    HOII% = HOY(HOS$)
    Label1.Caption = PERIO$
    Label3.Caption = HOS$
    Label4.Caption = Left$(Time$, 5)
    '
    Call REPLA(TTX0$, PERIO$, 53, 24)
    Call REPLA(TTX0$, HOS$, 77, 8)
    Call REPLA(TTX0$, Time$, 85, 8)
    JJ& = JJ& + 1
    Call GRAREG("RECOMOB", TTX0$, JJ&)
    '
    Call SETULTREG("RECOMOB", JJ&)
    Call CIERRARCH("RECOMOB")
    '
99  Screen.MousePointer = 1
    Command4.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call FINAL("*SOCOMOB")
    Command5.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    DoEvents
    'Show
    'Refresh
    Call PRECOMOB
    '
End Sub

Sub PRECOMOB()
    '
    Screen.MousePointer = 11
    List1.Clear
    JJ& = 0
    For U& = 1 To ULTREG&("RECOMOB") - 1
      '
      TTXX$ = REGLEIDO$("RECOMOB", U&)
      List1.AddItem TTXX$
      '
    Next U&
    List1.Refresh
    
    URE& = ULTREG&("RECOMOB")
    TTX0$ = REGLEIDO$("RECOMOB", URE&)
    TORD0& = Val(Left$(TTX0$, 6))
    ORPO0& = Val(Mid$(TTX0$, 7, 6))
    ORNE0& = Val(Mid$(TTX0$, 13, 6))
    TISTA0 = Val(Mid$(TTX0$, 19, 12))
    TIEMP0 = Val(Mid$(TTX0$, 31, 12))
    PERIO$ = TRIM$(Mid$(TTX0$, 53, 24))
    FEGEN$ = TRIM$(Mid$(TTX0$, 77, 8))
    HOGEN$ = TRIM$(Mid$(TTX0$, 85, 8))
    SUVADIF = Val(Mid$(TTX0$, 93, 12))
    '
    Label6(0).Caption = TRIM$(Str$(TORD0&))
    Label6(1).Caption = TRIM$(Str$(ORPO0&))
    Label6(2).Caption = TRIM$(Str$(ORNE0&))
    Label7(0).Caption = TRIM$(HORMINSE$(TISTA0))
    Label7(1).Caption = TRIM$(HORMINSE$(TIEMP0))
    DIFETI = TISTA0 - TIEMP0
    DITI$ = TRIM$(HORMINSE$(Abs(DIFETI)))
    Label7(2).ForeColor = RGB(0, 0, 0)
    If DIFETI < 0 Then Label7(2).ForeColor = RGB(255, 0, 0)
    Label7(2).Caption = DITI$
    RENDIM = 999: If TIEMP0 > 0.01 Then RENDIM = 100 * TISTA0 / TIEMP0
    If RENDIM > 999 Then RENDIM = 999
    Label8.ForeColor = RGB(0, 0, 0)
    If RENDIM < 100 Then Label8.ForeColor = RGB(255, 0, 0)
    Label8.Caption = TRIM$(CSTRING$(MKS$(RENDIM), 3, 8, 1, 2))
    Label9.ForeColor = RGB(0, 0, 0)
    If SUVADIF < 0 Then Label9.ForeColor = RGB(255, 0, 0)
    Label9.Caption = TRIM$(CSTRING$(MKS$(Abs(SUVADIF)), 3, 10, 2, 2))
    '
    Label1.Caption = PERIO$
    Label3.Caption = FEGEN$
    Label4.Caption = Left$(HOGEN$, 5)
    '
99  Screen.MousePointer = 1
    '
End Sub

Private Sub List1_DblClick()
   PRODUOP07.Show 1
End Sub
Sub VESARRAS(CIXI%)
   Exit Sub
   '
   If CIXI% > 0 Then
     '
     SCMA = Screen.MousePointer
     Screen.MousePointer = 11
     '
     Call CARDEPOS
     For KI% = 1 To 64
       CATEO#(KI%) = 0
     Next KI%
     '
     Call APERBASDAT("STOCKS")
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI% & " "
     '
     Dim MOVVVISTO As ADODB.Recordset
     Set MOVVVISTO = New ADODB.Recordset
     MOVVVISTO.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     CAMOVI& = 0
     With MOVVVISTO
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CAMOVI& = CAMOVI& + 1
           CANTI# = !CantIngre - !CantSalid
           DEP% = !DepoMovi
           '
           COLUX% = 1
           For KI% = 1 To NUSU%
             If DEP% = COSU%(KI%) Then
               COLUX% = KI%
               Exit For
             End If
           Next KI%
           '
           CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
         .MoveNext
         Loop
       End If
     End With
     MOVVVISTO.Close
     Set MOVVVISTO = Nothing
     '
     VESAL$ = String$(256, 0)
     For KI% = 1 To NUSU%
       Mid$(VESAL$, 4 * KI% - 3, 4) = MKS$(CATEO#(KI%))
     Next KI%
     '
     SQLX$ = "SELECT * FROM SalStock "
     SQLX$ = SQLX$ + " WHERE CodiInte = " & CIXI% & ""
     With SALLLSTOCK 'SalStock
       '.FindFirst "CodiInte = " & CIXI%
       'If .NoMatch = False Then
       If Not (.EOF And .BOF) Then
           '.Edit
         Else
           .AddNew   ' YA CONVERTIDO
           ON ERROR RESUME NEXT
           !CodiEmpr = Codiemp%
           ON ERROR GOTO 0
           !CodiInte = CIXI%
       End If
       !VeSalSto = VESAL$
       .Update
     End With
'     SalStock.Close
      SALLLSTOCK.Close
      Set SALLLSTOCK = Nothing
     '
   End If
   '
99 End Sub
'




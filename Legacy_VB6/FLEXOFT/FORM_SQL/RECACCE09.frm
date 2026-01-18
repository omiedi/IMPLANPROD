VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RECACCE09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion de Equipos para Reparación (2)"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   8640
      ScaleHeight     =   4485
      ScaleWidth      =   1350
      TabIndex        =   24
      Top             =   -120
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   29
         Top             =   4000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "RECACCE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   975
         Width           =   690
      End
      Begin VB.CommandButton Command11 
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
         Left            =   120
         Picture         =   "RECACCE09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "RECACCE09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Tabla de Accesorios"
         Top             =   2760
         Width           =   690
      End
      Begin VB.CommandButton Command5 
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
         Left            =   120
         Picture         =   "RECACCE09.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Muestra el Historial del Equipo Elegido"
         Top             =   1890
         Width           =   690
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<<< Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7980
      Picture         =   "RECACCE09.frx":0CE0
      TabIndex        =   20
      ToolTipText     =   "Pantalla Anterior"
      Top             =   7280
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1470
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   6825
      Width           =   6420
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "ACCESORIOS ENTREGADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   105
      TabIndex        =   18
      Top             =   4620
      Width           =   9360
      Begin MSFlexGridLib.MSFlexGrid GRID 
         Height          =   1710
         Left            =   105
         TabIndex        =   22
         Top             =   315
         Width           =   9100
         _ExtentX        =   16060
         _ExtentY        =   3016
         _Version        =   393216
         Rows            =   8
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "DEFECTOS INFORMADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   105
      TabIndex        =   12
      Top             =   1785
      Width           =   8415
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2625
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1995
         Width           =   5685
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   2625
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   315
         Width           =   5685
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Fallas Tipificadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -1365
         TabIndex        =   14
         Top             =   350
         Width           =   3960
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Otras Fallas (describir)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   105
         TabIndex        =   13
         Top             =   2100
         Width           =   2460
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente >>>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7980
      Picture         =   "RECACCE09.frx":0FEA
      TabIndex        =   8
      ToolTipText     =   "Pantalla Siguiente"
      Top             =   6825
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICACION DEL EQUIPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   8415
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   315
         Width           =   2625
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Talón"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   4305
         TabIndex        =   10
         Top             =   405
         Width           =   2760
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7185
         TabIndex        =   9
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Garantia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6800
         TabIndex        =   7
         Top             =   825
         Width           =   870
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7800
         TabIndex        =   6
         Top             =   750
         Width           =   465
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   940
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -1575
         TabIndex        =   4
         Top             =   360
         Width           =   3960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2370
         TabIndex        =   3
         Top             =   825
         Width           =   870
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3330
         TabIndex        =   2
         Top             =   750
         Width           =   2625
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   800
         TabIndex        =   1
         Top             =   1150
         Width           =   7470
      End
   End
   Begin VB.Label APRO 
      Height          =   330
      Left            =   7770
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   -525
      TabIndex        =   17
      Top             =   6825
      Width           =   1935
   End
End
Attribute VB_Name = "RECACCE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call Command3_Click
End Sub

Private Sub Command11_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
    '
    APRO = "PREV"
    Hide
    '
End Sub

Private Sub Command3_Click()
   '
   If DERACCE%("CARTACCE", "Actualizar Tabla de Accesorios") >= 2 Then
      Call COPYSTRU("INDIPEDI", "TACCESOR")
      Call SETULTREG("!TATIPFA", 0)
      JI& = 0
      For i& = 1 To ULTREG&("TACCESOR")
         X$ = REGLEIDO$("TACCESOR", i&)
         If CVS(Left$(X$, 4)) > 0 Then
            JI& = JI& + 1
            Call REPLA(X$, MKS$(JI&), 1, 4)
            Call REGAPP("!TATIPFA", X$)
         End If
      Next i&
      Call CIERRARCH("TACCESOR")
      '
      VTB% = VENTABU%("TATIPFA/NC/TITUPAN=Tabla de Accesorios")
      If VTB% < 0 Then GoTo 99
      J& = 0
      For i& = 1 To ULTREG&("!TATIPFA")
         X$ = REGLEIDO$("!TATIPFA", i&)
         If TRIM$(Mid$(X$, 5)) <> "" Then
             J& = J& + 1
             Call REPLA(X$, MKS$(J&), 1, 4)
             Call GRAREG("TACCESOR", X$, J&)
         End If
      Next i&
      '
      Call SETULTREG("TACCESOR", J&)
      Call CIERRARCH("TACCESOR")
      Call FILESORT("TACCESOR", "", 2, 2, "ASC")
      '
      Call ACGRIDACC
   End If
   '
99
   '
End Sub

Sub ACGRIDACC()
   '
   GRID.Clear
   GRID.TextMatrix(0, 0) = "Accesorio"
   GRID.TextMatrix(0, 1) = "Cant"
   GRID.TextMatrix(0, 2) = "Observaciones)"
   '
   For i& = 1 To ULTREG&("TACCESOR")
      X$ = REGLEIDO$("TACCESOR", i&)
      If CVS(Left$(X$, 4)) > 0 Then
         JI& = JI& + 1
10       If JI& >= GRID.Rows Then
            GRID.Rows = GRID.Rows + 1
            GoTo 10
         End If
         GRID.TextMatrix(JI&, 0) = Mid$(X$, 5, 48)
      End If
   Next i&
   '
End Sub
'
Private Sub Command4_Click()
   '
   Command4.Enabled = False
   If VALIDATOS% < 1 Then GoTo 99
   '
   Command4.Enabled = True
   APRO = "NEXT"
   Hide
   '
99 Command4.Enabled = True

End Sub

Function VALIDATOS%()
   '
   VALIDATOS% = 1
   '
End Function

Private Sub Command5_Click()
    Command5.Enabled = False
    CONUSE07.Label7 = Label14
    CONUSE07.Show 1
    Command5.Enabled = True
End Sub
'
Sub Form_Load()
   '
   List1.Clear
   For i& = 1 To ULTREG&("TATIPFA")
       X$ = REGLEIDO$("TATIPFA", i&)
       If CVS(Left$(X$, 4)) > 0 Then
           TTXX$ = FORMATNUM$(CVS(Left$(X$, 4)), "F4") + " - " + Mid$(X$, 5, 44)
           List1.AddItem TTXX$
       End If
   Next i&
   Call CIERRARCH("TATIPFA")
   '
   GRID.ColWidth(0) = 3500
   GRID.ColWidth(1) = 800
   GRID.ColWidth(2) = 4400
   '
   Call ACGRIDACC

End Sub

Private Sub GRID_DblClick()
    '
    NUMEAC& = GRID.Row
    DEAC$ = GRID.TextMatrix(NUMEAC&, 0)
    CANU& = XVALO(GRID.TextMatrix(NUMEAC&, 1))
    OBSE$ = GRID.TextMatrix(NUMEAC&, 2)
    '
    VDEF$ = MKS$(NUMEAC&) + AJUSTI$(DEAC$, 32) + MKS$(CANU&) + AJUSTI$(OBSE$, 48)
    OBX$ = OBJPLA$("CARGACCE", VDEF$)
    If OBX$ <> "" Then
       CANU& = CVS(Mid$(OBX$, 37, 4))
       OBSE$ = TRIM$(Mid$(OBX$, 41, 48))
       GRID.TextMatrix(NUMEAC&, 1) = CANU&
       GRID.TextMatrix(NUMEAC&, 2) = OBSE$
    End If
    '
End Sub


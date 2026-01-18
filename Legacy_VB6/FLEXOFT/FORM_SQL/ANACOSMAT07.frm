VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ANACOSMAT07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Costos de Materiales"
   ClientHeight    =   6300
   ClientLeft      =   135
   ClientTop       =   2070
   ClientWidth     =   11865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5805
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   10239
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   840
      Picture         =   "ANACOSMAT07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado (Tabla) de Costo de Suministros"
      Top             =   5880
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   210
      Picture         =   "ANACOSMAT07.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ficha sw Análisis de Costo de Suministros"
      Top             =   5880
      Width           =   540
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   0
      TabIndex        =   2
      Top             =   315
      Width           =   11880
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12000
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"ANACOSMAT07.frx":024C
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   40
         TabIndex        =   4
         Top             =   50
         Width           =   11880
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANACOSMAT07.frx":0303
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "ANACOSMAT07.frx":0537
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   11320
      X2              =   11320
      Y1              =   5900
      Y2              =   6200
   End
   Begin VB.Line Line1 
      X1              =   6520
      X2              =   11330
      Y1              =   6200
      Y2              =   6200
   End
   Begin VB.Label Label5 
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
      Height          =   330
      Left            =   10120
      TabIndex        =   6
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   8925
      TabIndex        =   7
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label4 
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
      Height          =   330
      Left            =   7730
      TabIndex        =   5
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Left            =   6530
      TabIndex        =   0
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5475
      TabIndex        =   1
      Top             =   5955
      Width           =   825
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10500
      Y1              =   5060
      Y2              =   5060
   End
End
Attribute VB_Name = "ANACOSMAT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
   ANACOSTO.COMMLAB.Caption = "PRISUMI"
   ANACOSTO07.COMMLAB.Caption = "PRISUMI"
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
  '\\\OT-06-0133-WAR-06/05/06///
  Command2.Enabled = False
  Screen.MousePointer = 11
  JJ& = 0
  Y$ = REGBLAN$("LICOSMAT")
  For J% = 0 To List1.ListCount - 1
    x$ = List1.List(J%)
    '
    CE1$ = Mid$(x$, 1, 15)
    Ci% = CODINT%(CE1$)
    Q1 = Val(Mid$(x$, 40, 10))
    UM1$ = Mid$(x$, 52, 3)
    PUN1# = Val(Mid$(x$, 169, 10))
    MON1$ = Mid$(x$, 179, 10)
    PUND1# = Val(Mid$(x$, 189, 10))
    CMAPRI1# = Val(Mid$(x$, 55, 10))
    CPIE1# = Val(Mid$(x$, 65, 10))
    CTER1# = Val(Mid$(x$, 75, 10))
    CTOT1# = Val(Mid$(x$, 85, 10))
    '
    Call REPLA(Y$, MKI$(Ci%), 1, 2)
    Call REPLA(Y$, MKS$(Q1), 3, 4)
    Call REPLA(Y$, UM1$, 7, 4)
    Call REPLA(Y$, MKD$(PUN1#), 11, 8)
    Call REPLA(Y$, MON1$, 19, 10)
    Call REPLA(Y$, MKD$(PUND1#), 29, 8)
    Call REPLA(Y$, MKD$(CMAPRI1#), 37, 8)
    Call REPLA(Y$, MKD$(CPIE1#), 45, 8)
    Call REPLA(Y$, MKD$(CTER1#), 53, 8)
    Call REPLA(Y$, MKD$(CTOT1#), 61, 8)
    JJ& = JJ& + 1
    Call GRAREG("LICOSMAT", Y$, JJ&)
    '
  Next J%
  Screen.MousePointer = 1
  Call SETULTREG("LICOSMAT", JJ&)
  Call HEADERS("LICOSMAT", "")
  On Error Resume Next
     TTT1$ = ANACOSTO.Text1 + " - " + ANACOSTO.Text3
     If Err Then
        On Error GoTo 0
        TTT1$ = ANACOSTO07.Text1 + " - " + ANACOSTO07.Text3
     End If
  On Error GoTo 0
  '
  Call HEADERS("LICOSMAT", "Articulo: " + TTT1$)
  Call FINAL("*LICOSMAT")
  Command2.Enabled = True
  '\\\OT-06-0133-WAR-FIN///
End Sub


Private Sub Form_Activate()
   'CARGA GRILLA DE COSTOS DE SUMINISTROS
   CAMPOS$ = "Codigo/A14;Descripción/A26;Cant/F10.1;U-M/A3;M.Prima/F11.5;Piezas/F11.5;3ros/F11.5;Total/F12.5"
   Call CARGA_ENCABEZADO(ANACOSMAT07.GRID, CAMPOS$, ANACOSMAT07, , , ANTOT)
    'LE PASO LAS POSICIONES DEL LIST
   Call PASAR_LIST_a_Grilla(Me.List1, Me.GRID, "1,15;17,23;40,10;52,4;55,10;65,10;75,10;85,10") ' POSICION 16,1 ES PARA SALTEAR UNA COLUMNA
End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   Hide
End Sub

Private Sub GRID_DblClick()

    NroFila& = GRID.Row
    List1.ListIndex = NroFila& - 1
    Call List1_DblClick
    
End Sub

Private Sub List1_DblClick()
   '
   CODICOM$ = TRIM$(Left$(List1.TEXT, 16))
   If CODICOM$ = String$(16, "-") Or CODICOM$ = String$(16, "=") Then
     Exit Sub
   End If
   '
   On Error Resume Next
   TTXX1$ = ""
   TTXX1$ = TTXX1$ + ANACOSTO.Text1
   TTXX1$ = TTXX1$ + ANACOSTO07.Text1
   On Error GoTo 0
   '
   If TRIM$(TTXX1$) <> "" Then
     LLAMACOSTO$ = TTXX1$
     ORILLAMA$ = "MAT"
     REGILLAMA% = List1.ListIndex
     TOPELLAMA% = List1.TopIndex
   End If
   '
   COD$ = TRIM$(Left$(List1.TEXT, 16))
   If COD$ <> "" Then
     Hide
     On Error Resume Next
     ANACOSTO.SSTab1.Tab = 0
     ANACOSTO07.SSTab1.Tab = 0
     DoEvents
     ANACOSTO.Text1.TEXT = COD$
     ANACOSTO07.Text1.TEXT = COD$
     On Error GoTo 0
   End If
   '
End Sub


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ANACOSMOD07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Costos de Mano de Obra"
   ClientHeight    =   6330
   ClientLeft      =   135
   ClientTop       =   2070
   ClientWidth     =   14595
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5820
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   10266
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
      Picture         =   "ANACOSMOD07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado (Tabla) de Costo de Mano de Obra"
      Top             =   5880
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   210
      Picture         =   "ANACOSMOD07.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   540
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
      Height          =   5460
      Left            =   -30
      TabIndex        =   2
      Top             =   330
      Width           =   11755
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
         Caption         =   $"ANACOSMOD07.frx":024C
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
      OleObjectBlob   =   "ANACOSMOD07.frx":0303
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "ANACOSMOD07.frx":0537
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
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   13125
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
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
      Left            =   11445
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
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
      Left            =   10005
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
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
      Left            =   7815
      TabIndex        =   5
      Top             =   5880
      Width           =   1845
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   11865
      Y1              =   5775
      Y2              =   5775
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
      Left            =   6405
      TabIndex        =   1
      Top             =   5955
      Width           =   1305
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10500
      Y1              =   5060
      Y2              =   5060
   End
End
Attribute VB_Name = "ANACOSMOD07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
   ANACOSTO.COMMLAB.Caption = "PRIMOD"
   ANACOSTO07.COMMLAB.Caption = "PRIMOD"
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
 '\\\OT-06-0161-WAR-28/05/06///
  Command2.Enabled = False
  Screen.MousePointer = 11
  JJ& = 0
  Y$ = REGBLAN$("LICOSMOD")
  For J% = 0 To List1.ListCount - 1
    x$ = List1.List(J%)
    '
    CE1$ = Mid$(x$, 1, 15)
    CIA% = Ci%
    Ci% = CODINT%(CE1$)
    If TRIM$(CE1$) = "" And Ci% = 0 Then
     Ci% = CIA%
    End If
    ' OPERA1$ = Mid$(X$, 39, 15)
    OPERA1$ = Mid$(x$, 21, 30)
    If OPERA1$ = "------------------------------" Or OPERA1$ = "==============================" Then
      GoTo 10
    End If
    If TRIM$(Left$(x$, 16)) <> "" Then OPERA1$ = ""
    CAOP1 = Val(Mid$(x$, 199, 5))
    MINUT1$ = Mid$(x$, 55, 10)
    '
    SEG1 = Val(Right$(MINUT1$, 3))
    HOR1 = Val(Left$(MINUT1$, 3))
    MINUT2 = Val(Mid$(MINUT1$, 5, 2))
    '
    SEG2 = SEG1 / 60 'PASO SEGUNDOS A MINUTOS
    HOR2 = HOR1 * 60 'PASO HORAS A MINUTOS
    MINUTOT1 = SEG2 + HOR2 + MINUT2
    '
    PUNPES1# = Val(Mid$(x$, 204, 10))
    CMOD1# = Val(Mid$(x$, 65, 10))
    CEQUI1# = Val(Mid$(x$, 75, 10))
    CTOT1# = Val(Mid$(x$, 85, 10))
    '
    Call REPLA(Y$, MKI$(Ci%), 1, 2)
    Call REPLA(Y$, OPERA1$, 3, 19)
    Call REPLA(Y$, MKS$(CAOP1), 21, 4)
    Call REPLA(Y$, MKS$(MINUTOT1), 25, 4)
    Call REPLA(Y$, MKD$(PUNPES1#), 29, 8)
    Call REPLA(Y$, MKD$(CMOD1#), 37, 8)
    Call REPLA(Y$, MKD$(CEQUI1#), 45, 8)
    Call REPLA(Y$, MKD$(CTOT1#), 53, 8)
    JJ& = JJ& + 1
    Call GRAREG("LICOSMOD", Y$, JJ&)
    '
10 Next J%
  Screen.MousePointer = 1
  Call SETULTREG("LICOSMOD", JJ&)
  Call HEADERS("LICOSMOD", "")
  On Error Resume Next
     TTT1$ = ANACOSTO.Text1 + " - " + ANACOSTO.Text3
     If Err Then
        On Error GoTo 0
        TTT1$ = ANACOSTO07.Text1 + " - " + ANACOSTO07.Text3
     End If
  On Error GoTo 0
  '
  Call HEADERS("LICOSMOD", "Articulo: " + TTT1$)
  Call FINAL("*LICOSMOD")
  Command2.Enabled = True
  '\\\OT-06-0161-WAR-FIN///
End Sub

Private Sub Form_Activate()
   'CARGA GRILLA DE COSTOS DE SUMINISTROS
    CAMPOS$ = "Codigo/A16;Descripción/A30;Operacion/A16;hh:mm:ss.nn/H12;M.O.D./F14.2;C.Equip./F14.2;Total/F14.2"
   Call CARGA_ENCABEZADO(ANACOSMOD07.GRID, CAMPOS$, ANACOSMOD07, , , ANTOT)

   For i% = 1 To Me.List1.ListCount - 1
        Z$ = List1.List(i% - 1)
        CODI$ = TRIM$(Mid(Z$, 1, 15))
        CI1% = CODINT(CODI$)
        If CI1% > 0 Then
          ANACOSMOD07.GRID.Rows = ANACOSMOD07.GRID.Rows + 1
          ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 1) = TRIM$(Mid(Z$, 1, 15))
          ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 2) = DESCRIT0$(CI1%)
        Else
          ANACOSMOD07.GRID.Rows = ANACOSMOD07.GRID.Rows + 1
          ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 3) = TRIM$(Mid(Z$, 21, 30))
        End If
        
        ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 4) = TRIM$(Mid(Z$, 51, 14))
        ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 5) = TRIM$(Mid(Z$, 65, 10))
        ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 6) = TRIM$(Mid(Z$, 75, 10))
        ANACOSMOD07.GRID.TextMatrix(ANACOSMOD07.GRID.Rows - 1, 7) = TRIM$(Mid(Z$, 85, 10))
   Next i%
   
'    'LE PASO LAS POSICIONES DEL LIST
'   Call PASAR_LIST_a_Grilla(Me.List1, Me.GRID, "1,15;17,35;21,30;51,14;65,10;75,10;85,10")
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
     ORILLAMA$ = "MOD"
     REGILLAMA% = List1.ListIndex
     TOPELLAMA% = List1.TopIndex
   End If
   '
   COPER% = -1
   For KKK% = List1.ListIndex To 0 Step -1
     COD$ = TRIM$(Left$(List1.List(KKK%), 16))
     If COD$ <> "" Then
       If COD$ <> String$(16, "-") Then
         On Error Resume Next
           ANACOSTO.SSTab1.Tab = 2
           ANACOSTO07.SSTab1.Tab = 2
           DoEvents
           ANACOSTO.Text1.TEXT = COD$
           ANACOSTO07.Text1.TEXT = COD$
           DoEvents
           If COPER% < 0 Then COPER% = 0
           ANACOSTO.List3.ListIndex = COPER%
           ANACOSTO07.List3.ListIndex = COPER%
         On Error GoTo 0
         Hide
         Exit Sub
       End If
     End If
     COPER% = COPER% + 1
   Next KKK%

End Sub

VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELCHE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Cheques a Depositar"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   7080
      ScaleHeight     =   3870
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   130
         Picture         =   "SELCHE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Búsqueda por Número, Banco, Importe, etc."
         Top             =   2220
         Width           =   680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   130
         Picture         =   "SELCHE07.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Conmuta Ordenamiento por Fecha / Importe"
         Top             =   1520
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
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
         Height          =   700
         Left            =   130
         Picture         =   "SELCHE07.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   680
      End
      Begin VB.CommandButton BOTREC 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   130
         Picture         =   "SELCHE07.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   810
         Width           =   680
      End
   End
   Begin VB.ListBox LICHESEL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "SELCHE07.frx":1AC4
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELCHE07.frx":1CF8
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ************"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL VALORES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   12
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECCIONADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1845
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ************"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "SELCHE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODORDE%
Dim MUESTRACAR%

Private Sub BOTEXIT_Click()
    LICHESEL.Clear
    Label2.Caption = ""
    Hide
End Sub
'

Private Sub BOTREC_Click()
    SUANT# = XVALO(Label2)
    Call SUCHESEL
    If XVALO(Label2) = SUANT# Then
        Hide
      Else
        Call MENSERR(24, "Revise Valores Seleccionados")
    End If
End Sub

Private Sub Command1_Click()
    If MODORDE% = 1 Then
      MODORDE% = 2
        For KI% = 1 To LICHESEL.ListCount
10        IMPOI# = XVALO(Mid$(LICHESEL.List(KI% - 1), 41, 12))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            IMPOJ# = XVALO(Mid$(LICHESEL.List(KJ% - 1), 41, 12))
            If IMPOJ# < IMPOI# Then
              TTXX$ = LICHESEL.List(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 10
            End If
          Next KJ%
        Next KI%
      ElseIf MODORDE% = 2 Then
        MODORDE% = 1
        For KI% = 1 To LICHESEL.ListCount
20        FECHI% = FECHANUM(Mid$(LICHESEL.List(KI% - 1), 33, 8))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            FECHJ% = FECHANUM(Mid$(LICHESEL.List(KJ% - 1), 33, 8))
            If FECHJ% < FECHI% Then
              TTXX$ = LICHESEL.List(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 20
            End If
          Next KJ%
        Next KI%
      
    End If
    '
End Sub

Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.List(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Form_Activate()
    Call SUCHESEL
    MODORDE% = 1
    'LICHESEL.SetFocus
End Sub

Sub SUCHESEL()
    SUCHESE# = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.List(i& - 1), 41, 12))
       End If
    Next i&
    Label2.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    Label2.Refresh
End Sub

Private Sub Form_Load()
    '
    MODORDE% = 1
    MUESTRACAR% = 0
    If DERACCE%("VECARCHE/NOMESS", "Visualizar Cartera de Cheques y Valores") >= 2 Then
      MUESTRACAR% = 1
      Label5.Visible = False
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
End Sub

Private Sub LICHESEL_Click()
   Call SUCHESEL
End Sub





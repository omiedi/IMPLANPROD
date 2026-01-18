VERSION 5.00
Begin VB.Form CIERRCTELMDB 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Consultas Telefonicas"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9555
      TabIndex        =   13
      Top             =   3990
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9555
      TabIndex        =   12
      Top             =   3675
      Width           =   645
   End
   Begin VB.Frame Frame4 
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1420
      Left            =   8025
      TabIndex        =   9
      Top             =   2860
      Width           =   1260
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Procesa Cierre General"
         Top             =   735
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   315
         Width           =   960
      End
   End
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
      Height          =   630
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancela los Datos Ingresados"
      Top             =   2940
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Consultas pendientes de cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   10095
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
         Height          =   2100
         Left            =   210
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   9650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "  Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   105
      TabIndex        =   4
      Top             =   2835
      Width           =   7785
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   315
         Width           =   7380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Caption         =   "  Diagnóstico e Instrucciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3430
      Left            =   105
      TabIndex        =   2
      Top             =   4335
      Width           =   10095
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   315
         Width           =   9650
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000013&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8400
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   1485
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   225
      Left            =   5250
      TabIndex        =   17
      Top             =   2730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   3990
      TabIndex        =   16
      Top             =   2730
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "P:"
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
      Left            =   9240
      TabIndex        =   15
      Top             =   4060
      Width           =   300
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "T:"
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
      Left            =   9240
      TabIndex        =   14
      Top             =   3750
      Width           =   300
   End
End
Attribute VB_Name = "CIERRCTELMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUCETE$
Dim ABUSQUE$

Private Sub BOTEXIT_Click()
    Unload Me
End Sub

Private Sub BOTREC_Click()
   '
   Screen.MousePointer = 11
   For KU& = 1 To List1.ListCount
     If List1.Selected(KU& - 1) = True Then
        NUCOTE$ = Left$(List1.List(KU& - 1), 12)
        CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
        FEXI% = FECHANUM(Text1)
        FENUE$ = FETEXCOR1$(FEXI%)
        Call ACTUREGISTRO("COTELEF", "Fech_Cierre", CONDI$, FENUE$, RAFE&, "NOMESS")
     End If
29 Next KU&
   '
   Screen.MousePointer = 1
   Call COMUNI("Cierre Procesado")
   Call RECARLISTA
   '
End Sub

Private Sub Form_Load()
    Text1.TEXT = HOS$
End Sub

Private Sub Label7_Change()
   Call CARDATCOTE
End Sub

Private Sub List1_Click()
   Label7 = Left$(List1.TEXT, 12)
End Sub

Private Sub List1_DblClick()
    CTACT$ = TRIM$(Left$(List1.TEXT, 12))
    If COMALTER%("¿ Está Seguro de Eliminar la C-Tel '" + CTACT$ + "' ?\\ No, Cancelar\Si, Eliminar") = 2 Then
       Screen.MousePointer = 11
       CONDI$ = "Nume_Cotel = '" & CTACT$ & "'"
       Call BORRAREGISTROS("COTELEF", CONDI$, RAFE&, "NOMESS")
       '
       Call RECARLISTA
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("SELFECHA")
    Text1 = VALACT1$("SELFECHA")
    
End Sub

Sub CARDATCOTE()
   '
   NUCOTE$ = TRIM$(Label7)
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   Text22 = VALOBADA("COTELEF", "DetaConsu", CONDI$, "NOMESS")
   Text23 = VALOBADA("COTELEF", "DiagInstru", CONDI$, "NOMESS")
   '
End Sub

Private Sub Text2_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text22_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text23_GotFocus()
   BOTREC.SetFocus
End Sub

Private Sub Text3_GotFocus()
   BOTREC.SetFocus
End Sub

Sub RECARLISTA()
   '
   Screen.MousePointer = 11
   DES% = XVALO(Label2)
     DESDE = FETEXCOR1$(DES%)
   HAS% = XVALO(Label3)
     HASTA = FETEXCOR1$(HAS%)
   List1.Clear     ' EPB - NECESARIO PORQUE SI NO AL VOLVER A ENTRAR DUPLICA LA CARGA
   CONDI0$ = "FechEmis >= '" & DESDE & "' "
   CONDI0$ = CONDI0$ & " AND FechEmis <= '" & HASTA & "' "
   CONDI$ = CONDI0$ & " AND (Fech_Cierre <= '30/NOV/1999' OR Fech_Cierre IS NULL)"
   Call CARGALISTA(List1, "COTELEF", "Nume_Cotel/A12;Referencia/A66", CONDI$)
   '
   Text2 = CantRegistros("COTELEF", CONDI0$)
   Text3 = CantRegistros("COTELEF", CONDI$)
   Screen.MousePointer = 1
   '
End Sub

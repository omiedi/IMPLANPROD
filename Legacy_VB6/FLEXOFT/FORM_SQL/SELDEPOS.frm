VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELDEPOS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBCA9D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Depósitos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7755
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LISNUM 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde para seleccionar"
      Top             =   120
      Width           =   6705
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   5775
      Left            =   6840
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
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
         Height          =   640
         Left            =   120
         Picture         =   "SELDEPOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "SELDEPOS.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4425
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "SELDEPOS.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   3555
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELDEPOS.frx":059E
         Top             =   4770
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "SELDEPOS.frx":24C0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "SELDEPOS.frx":26F4
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5500
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line2 
      X1              =   5505
      X2              =   5505
      Y1              =   495
      Y2              =   3245
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   7
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELDEPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Public FORMNAME As Form
Public TXT As TextBox


Private Sub BOTREC_Click()
     BOTREC.Enabled = False
     '
'     STRINGDEPOS$ = ""
'     List1.Clear
'     For U& = 1 To LISNUM.ListCount
'       If LISNUM.Selected(U& - 1) = True Then
'        TXTX$ = LISNUM.List(U& - 1)
'        List1.AddItem TXTX$
'        NDEP$ = TRIM$(str$(Mid$(TXTX$, 1, 6)))
''        CANTI = CANTI + 1
''        If CANTI < LISNUM.ListCount Then
'           STRINGDEPOS$ = STRINGDEPOS$ + NDEP$ + ","
''       End If
'      End If
'     Next U&
'     If TRIM$(STRINGDEPOS$) <> "" Then
'        Text1 = Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1) 'para quitar la coma
'     Else
'        Text1 = ""
'     End If
     Hide
     BOTREC.Enabled = True
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LISNUM.TopIndex
   For U& = 1 To LISNUM.ListCount
     If LISNUM.Selected(U& - 1) = True Then
         LISNUM.Selected(U& - 1) = False
       Else
         LISNUM.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LISNUM.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   LISNUM.Clear
   Unload Me
End Sub




Private Sub Command6_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label1 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Form_Load()
     Screen.MousePointer = 11
     LISNUM.Clear
     FIN& = ULTREG&("SUCURSAL")
      For U& = 1 To FIN&
         If U& <= 62 Then
            TX$ = REGBLAN$("SUCURSAL")
            X$ = REGLEIDO$("SUCURSAL", U&)
            NUSU% = U&
            If Len(X$) >= 64 Then
               COSU%(U&) = Asc(Mid$(X$, 64, 1))
               SUCU$(U&) = Left$(X$, 20)
               Call REPLA(TX$, TRIM$(Str$(COSU%(U&))), 1, 6)
               Call REPLA(TX$, SUCU$(U&), 8, 24)
               If UCase$(Mid$(X$, 62, 2) = "NO") Then
                  Call REPLA(TX$, UCase$(Mid$(X$, 62, 2)), 34, 2)
               End If
               LISNUM.AddItem TX$
            End If
         End If
     Next U&
     Screen.MousePointer = 1
     '
End Sub

Private Sub label1_Change()
   Label8 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label1)))
End Sub

Private Sub Label3_Change()
   Label2 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label3)))
End Sub



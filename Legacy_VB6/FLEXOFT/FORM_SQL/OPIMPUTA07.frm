VERSION 5.00
Begin VB.Form OPIMPUTA 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Imputación Débitos / Créditos"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   1485
      Left            =   4305
      ScaleHeight     =   1425
      ScaleWidth      =   1530
      TabIndex        =   7
      Top             =   0
      Width           =   1590
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
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
         Left            =   105
         Picture         =   "OPIMPUTA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aplica Imputación Contable Según Selección"
         Top             =   735
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   105
         Picture         =   "OPIMPUTA07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Mantiene Imputación Contable General"
         Top             =   150
         Width           =   645
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Imputación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   800
      Left            =   105
      TabIndex        =   0
      Top             =   1470
      Width           =   5010
      Begin VB.TextBox Eco21 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1890
         TabIndex        =   3
         Text            =   " "
         Top             =   1130
         Width           =   3270
      End
      Begin VB.TextBox TEXT21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1155
         TabIndex        =   2
         Text            =   " "
         Top             =   1130
         Width           =   435
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1575
         TabIndex        =   1
         Top             =   1130
         Width           =   180
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   345
         Width           =   1275
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   -3045
         TabIndex        =   6
         Top             =   1205
         Width           =   4110
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Index           =   5
         Left            =   -780
         TabIndex        =   5
         Top             =   400
         Width           =   1590
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   4
         Top             =   345
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Motivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   0
      TabIndex        =   10
      Top             =   105
      Width           =   4215
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
         Height          =   735
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   4005
      End
   End
End
Attribute VB_Name = "OPIMPUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'\\\OT-06-0403-MH-19/09/06///
   CUEVE$ = FORFASQL.Text16.TEXT
   QQQ% = CUECOINT%(CUEVE$)
   Call COMUNI("El Documento se Contabiliza\en la Cuenta Contable\'" + TRIM$(CUEVE$) + " - " + TRIM$(DENOCUE$(QQQ%)) + "'")
   Hide
End Sub
'\\\OT-06-0403-MH-FIN///

'\\\OT-06-0403-MH-18/09/06///
Private Sub Command12_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
     Text16 = VDEVU$
   End If
End Sub

Private Sub Command14_Click()
   CUEVE$ = Label1.Caption
   QQQ% = CUECOINT%(CUEVE$)
   If TRIM$(CUEVE$) <> "" Then
      If QQQ% > 0 Then
         Call COMUNI("El Documento se Contabiliza\en la Cuenta Contable\'" + TRIM$(CUEVE$) + " - " + TRIM$(DENOCUE$(QQQ%)) + "'")
         On Error Resume Next
         FORFASQL.Text16 = CUEVE$
         On Error GoTo 0
         Hide
         Exit Sub
      End If
   End If
   Call MENSERR(24, "Cuenta Invalida")
99 End Sub
'
Private Sub Form_Load()
'    CUEVE$ = CONTROL$("SVD451S", "CUEVE-8")
'    Text16.TEXT = CUEVE$
   List1.Clear
   TX$ = ""
   FIN = ULTREG("IMPDECRE")
   For I& = 1 To FIN
     X$ = REGLEIDO$("IMPDECRE", I&)
     MTV$ = Mid$(X$, 3, 40)
     CUCNT% = CVI(Mid$(X$, 43, 2))
     DNMNCION$ = Left$(ECOARCH("CUECON", MKI(CUCNT%)), 20)
     
     Call REPLA(TX$, MTV$, 1, 30)
     Call REPLA(TX$, Str$(CUCNT%), 60, 4)
     List1.AddItem TX$
   Next I&
   
End Sub
'
'Sub Option1_Click()
'    CUEVE$ = CONTROL$("SVD451S", "CUEVE-8")
'    Text16.TEXT = CUEVE$
'End Sub
'
' Sub Option2_Click()
'    CUEVE$ = CONTROL$("SVD451S", "CUEVE-9")
'    Text16.TEXT = CUEVE$
'End Sub
'
' Sub Option3_Click()
'    CUEVE$ = CONTROL$("SVD451S", "CUEVE-11")
'    Text16.TEXT = CUEVE$
'End Sub

Private Sub List1_Click()
  CODINTERNO% = XVALO(Mid$(List1.TEXT, 60, 4))
  Label1 = CODCUE(CODINTERNO%)
  Label16 = DENOCUE(CODINTERNO%)
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Label1), 20)
End Sub

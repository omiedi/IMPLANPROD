VERSION 5.00
Begin VB.Form OPIMPUTB 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Imputación Débitos / Créditos"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   600
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
      Begin VB.PictureBox Picture16 
         Height          =   1450
         Left            =   0
         ScaleHeight     =   1395
         ScaleWidth      =   1020
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   1450
            Left            =   0
            Picture         =   "OPIMPUTB.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
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
         Picture         =   "OPIMPUTB.frx":2398
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
         Picture         =   "OPIMPUTB.frx":26A2
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
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aceptar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPIMPUTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aceptar_Click()
   Call Command14_Click
End Sub

Private Sub Command1_Click()
'\\\OT-06-0403-MH-19/09/06///
   CUEVE$ = FORFACCS.Text16.TEXT
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
         FORFACCS.Text16 = CUEVE$
         On Error GoTo 0
         Hide
         Exit Sub
      End If
   End If
   Call MENSERR(24, "Cuenta Invalida")
99 End Sub
'
Private Sub Form_Load()
   '
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3600
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      Set Command14.Container = Me.Picture16
      Command14.ZOrder 0
    End If
  
   List1.Clear
   TX$ = ""
   FIN& = ULTREG&("IMPDECRE")
   If FIN& < 3 Then
      CUEVE$ = Control$("SVD451S", "CUEVE-8")
      CUEII% = CUEINT%(CUEVE$)
      TTXX$ = MKI$(1) + AJUSTI$("Cheques Rechazados", 40) + MKI$(CUEII%)
      Call GRAREG("IMPDECRE", TTXX$, 1)
      '
      CUEVE$ = Control$("SVD451S", "CUEVE-9")
      CUEII% = CUEINT%(CUEVE$)
      TTXX$ = MKI$(1) + AJUSTI$("Diferencia de Cambio", 40) + MKI$(CUEII%)
      Call GRAREG("IMPDECRE", TTXX$, 2)
      '
      CUEVE$ = Control$("SVD451S", "CUEVE-11")
      CUEII% = CUEINT%(CUEVE$)
      TTXX$ = MKI$(1) + AJUSTI$("Debitos / Creditos Varios", 40) + MKI$(CUEII%)
      Call GRAREG("IMPDECRE", TTXX$, 3)
      Call CIERRARCH("IMPDECRE")
   End If
   '
   For I& = 1 To FIN
     X$ = REGLEIDO$("IMPDECRE", I&)
     MTV$ = Mid$(X$, 3, 40)
     CUCNT% = CVI(Mid$(X$, 43, 2))
     DNMNCION$ = Left$(ECOARCH("CUECON", MKI(CUCNT%)), 20)
     
     Call REPLA(TX$, MTV$, 1, 30)
     ' ///OT-8-637 - 15/08/08
     CTA$ = FORMATNUM$(CUCNT%, "F4.0")
     Call REPLA(TX$, CTA$, 60, 4)
     ' ///OT-8-637 - FIN
     List1.AddItem TX$
   Next I&
   
End Sub

'
Private Sub List1_Click()
  CODINTERNO% = XVALO(Mid$(List1.TEXT, 60))
  Label1 = CODCUE(CODINTERNO%)
  Label16 = DENOCUE(CODINTERNO%)
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Label1), 20)
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

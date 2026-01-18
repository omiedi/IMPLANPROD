VERSION 5.00
Begin VB.Form FINCONFIS 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Comprobante"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3735
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Coranza Contado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   210
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         TabIndex        =   11
         Top             =   420
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   945
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   7
         Top             =   945
         Width           =   225
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Importe:"
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
         Left            =   315
         TabIndex        =   12
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Left            =   210
         TabIndex        =   10
         Top             =   1365
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Medio de Cobranza:"
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
         Left            =   315
         TabIndex        =   9
         Top             =   1050
         Width           =   1845
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   2430
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         TabIndex        =   4
         Top             =   420
         Width           =   225
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Depósito:"
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
         Left            =   420
         TabIndex        =   5
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   840
         Width           =   2010
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aprobar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1220
      Left            =   2730
      Picture         =   "FINCONFIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   795
   End
End
Attribute VB_Name = "FINCONFIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '\\\OT-07-0023-RG-24/01/07///
   DEP% = XVALO(Text7)
   If ECOARCH$("TADEPOSI", Chr(DEP%)) = "" Then
      Call MENSERR(24, "Falta Ingresar el Depósito de Consumo")
      Text7 = TRIM$(Str$(DEPOEXPE%))
      Text7.SetFocus
      Exit Sub
   End If
   '\\\OT-07-0023-RG-FIN///
   If Abs(XVALO(Text2)) >= 0.005 Then
      MECO% = XVALO(Text1)
      If MECO% < 1 Or CUEMECOB%(MECO%) < 1 Or TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%))) = "" Then
         Call MENSERR(24, "Medio de Cobranza Faltante o No Válido.")
         Text2.SetFocus
         Exit Sub
      End If
      Call GRACONTROL("MECOBRA", "CONTADO", TRIM$(Text1))
   End If
   '
   FINFACTURA.Label5.Caption = "APRO"
   Hide
   '
End Sub

Private Sub Command2_Click()
   Call SELECHO("MECOBRA")
   VDEVU$ = TRIM$(VALACT1$("MECOBRA"))
   If VDEVU$ <> "" Then
      Text1 = VDEVU$
   End If
End Sub

Private Sub Command3_Click()
    '\\\OT-07-0023-RG-24/01/07///
    Call SELECHO("TADEPOSI")
    DEP$ = VALACT1$("TADEPOSI")
    If DEP$ <> "" Then
       Text7.TEXT = DEP$
    End If
    '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FINFACTURA.Label5.Caption = "CANCE"
End Sub

Private Sub Text1_Change()
   Label2 = ECOARCH$("MECOBRA", MKI$(XVALO(Text1)))
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub

Private Sub TEXT2_LOSTFOCUS()
   NT2 = XVALO(Text2)
   Text2 = FORMATNUM$(NT2, "F10.2")
End Sub

Private Sub Text7_Change()
  '\\\OT-07-0023-RG-24/01/07///
  If Text7 = "" Then Label8.Caption = ""
     On Error Resume Next
     If Text7 <> "" Then
        DEP% = XVALO(Text7)
        Label8.Caption = Mid$(ECOARCH$("TADEPOSI", Chr(DEP%)), 1, 13)
     End If
  '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Text7_DblClick()
   '\\\OT-07-0023-RG-24/01/07///
   Call Command3_Click
   '\\\OT-07-0023-RG-FIN///
End Sub

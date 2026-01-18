VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INIMAY07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mayor por Cuenta Contable"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   8600
      Left            =   5775
      ScaleHeight     =   8565
      ScaleWidth      =   1245
      TabIndex        =   16
      Top             =   -160
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         Height          =   195
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   4080
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   24
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command15 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INIMAY07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   3240
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INIMAY07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1065
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "INIMAY07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   380
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0095C5E1&
      Caption         =   "UNIDADES DE NEGOCIO"
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
      Left            =   105
      TabIndex        =   13
      Top             =   3240
      Width           =   5500
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Consolidar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4095
         TabIndex        =   15
         Top             =   420
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unificado Empresa (Todas las U.N's)"
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
         Left            =   420
         TabIndex        =   14
         Top             =   450
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Caption         =   "PRESENTACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   5500
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Detalle según Sub-Sistemas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3360
         TabIndex        =   12
         Top             =   315
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Resumen según Contabilidad General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   630
         TabIndex        =   11
         Top             =   315
         Value           =   -1  'True
         Width           =   2235
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0095C5E1&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "Doble-Click para Editar Rango de Fechas"
      Top             =   1080
      Width           =   5500
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5100
         TabIndex        =   5
         Top             =   405
         Width           =   175
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   4
         Top             =   405
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Left            =   3990
         TabIndex        =   9
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label label5 
         Alignment       =   2  'Center
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
         Left            =   1260
         TabIndex        =   8
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1155
         TabIndex        =   7
         Top             =   500
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3045
         TabIndex        =   6
         Top             =   500
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "INIMAY07.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   240
      OleObjectBlob   =   "INIMAY07.frx":0992
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label8 
      Height          =   330
      Left            =   210
      TabIndex        =   20
      Top             =   105
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Left            =   2460
      TabIndex        =   2
      Top             =   300
      Width           =   1275
   End
   Begin VB.Label Ltext2 
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
      Left            =   3820
      TabIndex        =   1
      Top             =   210
      Width           =   1770
   End
   Begin VB.Label Label4 
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
      Left            =   105
      TabIndex        =   0
      Top             =   630
      Width           =   5490
   End
End
Attribute VB_Name = "INIMAY07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
   If Check1.Value = 1 Then
       Label7 = "Unificado Empresa (Todas las U.N's)"
     Else
       NUNE% = UNINEGO%
       DENU$ = TRIM$(VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%))
       Label7 = DENU$
   End If
   Label7.Refresh
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command15_Click()
   Label8 = "OK"
   Hide
End Sub

Private Sub Command21_Click()
   Call CommSel("SELFECHA", Command21, Label5)
End Sub

Private Sub Command4_Click()
   Call CommSel("SELFECHA", Command4, Label6)
End Sub

Private Sub Form_Load()
   Label5 = FECHATEX$(COMEJ%)
   Label6 = FECHATEX$(FINEJ%)
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   Height = Frame2.Top + Height - ScaleHeight
   CANUNE% = CantRegistros&("UNEGOCIOS")
   If CANUNE% > 0 Then
      Height = 4700
      Check1.Value = 0
      NUNE% = UNINEGO%
      DENU$ = TRIM$(VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%))
      Label7 = DENU$
   End If
   Command15.Top = Height - 1100
   Screen.MousePointer = 1
End Sub

Private Sub Frame2_DblClick()
   Check1.Value = 0
   NUNE% = UNINEGO%
   DENU$ = TRIM$(VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%))
   Label7 = DENU$
End Sub

Private Sub Frame6_DblClick()
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ <> "" Then
      Label5 = FECHATEX$(Des%)
      Label6 = FECHATEX$(Has%)
   End If
End Sub

Private Sub label5_DblClick()
   Call Command21_Click
End Sub

Private Sub Label6_DblClick()
   Call Command4_Click
End Sub

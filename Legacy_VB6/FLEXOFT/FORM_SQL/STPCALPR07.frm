VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STPCALPR07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   14
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
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
      Left            =   6120
      Picture         =   "STPCALPR07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Registra y Guarda la Configuración"
      Top             =   1155
      Width           =   650
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ESCALA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3060
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "0 a 10 puntos"
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
         Height          =   450
         Left            =   1890
         TabIndex        =   11
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "0 a 5 puntos"
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
         Height          =   450
         Left            =   525
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3270
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puntaje Máximo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   315
         TabIndex        =   8
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1340
         TabIndex        =   7
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "TABLA DE INDICES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5685
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   210
         TabIndex        =   6
         Top             =   735
         Width           =   5250
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         ScaleHeight     =   300
         ScaleWidth      =   5220
         TabIndex        =   4
         Top             =   420
         Width           =   5250
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "N. Indice de Calidad                Pond."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   5
            Top             =   0
            Width           =   5000
         End
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   6120
      Picture         =   "STPCALPR07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1995
      Width           =   650
   End
   Begin VB.CommandButton command1 
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
      Left            =   6120
      Picture         =   "STPCALPR07.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   650
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4260
      Left            =   6000
      ScaleHeight     =   4230
      ScaleWidth      =   1005
      TabIndex        =   13
      Top             =   0
      Width           =   1035
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "STPCALPR07.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STPCALPR07.frx":0992
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "STPCALPR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   ESCACALI = 5: If Option2.Value = True Then ESCACALI = 10
   Call GRACONTROL("CALIPROV", "ESCACALI", TRIM$(Str$(ESCACALI)))
   Call GRACONTROL("CALIPROV", "PUMAXCAL", TRIM$(Label2))
   Command2.Enabled = True
End Sub
'

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
       Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Form_Load()
   '
   If ULTREG&("ICALPROV") < 1 Then
     REBLA$ = REGBLAN$("ICALPROV")
     zz$ = REBLA$: Call REPLA(zz$, Chr$(1) + "Calidad Propiamente Dicha", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(2) + "Embalajes", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(3) + "Identificación", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(4) + "Entrega de Certificados", 1, 32): Call REGAPP("ICALPROV", zz$)
     Call CIERRARCH("ICALPROV")
   End If
   '
   If Val(CONTROL$("CALIPROV", "ESCACALI")) = 10 Then
      Option2.Value = True
   End If
   '
   List1.Clear: PUNMAX = 0
   FIN& = ULTREG&("ICALPROV")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("ICALPROV", U&)
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, FORMATNUM$(U&, "I2"), 1, 2)
     Call REPLA(TTXX$, Mid$(XX$, 2, 31), 5, 30)
     Call REPLA(TTXX$, FORMATNUM$(CVS(Mid$(XX$, 33, 4)), "F6.2"), 36, 6)
     PUNMAX = PUNMAX + CVS(Mid$(XX$, 33, 4))
     List1.AddItem TTXX$
   Next U&
   '
   Label2.Caption = TRIM$(FORMATNUM$(PUNMAX, "F6.2"))
   '
   Call APLICACIONSKINS(Me, Picture2)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = rasocli$(NCMIN%)
   End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Text8_Change()
   Label8 = ""
   NCMAX% = Val(Text8.TEXT)
   If Val(NCMAX%) > 0 Then
     Label8 = rasocli$(NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub List1_DblClick()
   '
   Call TRALOCAL("ICALPROV", "")
   VTB% = VENTABU%("ICALPROV")
   If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!ICALPROV")
       XX$ = REGLEIDO$("!ICALPROV", U&)
       If U& < 256 Then
         Call REPLA(XX$, Chr$(U&), 1, 1)
         JJ& = JJ& + 1
         Call GRAREG("ICALPROV", XX$, JJ&)
       End If
     Next U&
     Call SETULTREG("ICALPROV", JJ&)
     Call CIERRARCH("ICALPROV")
     '
     List1.Clear
     FIN& = ULTREG&("ICALPROV")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("ICALPROV", U&)
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, FORMATNUM$(U&, "I2"), 1, 2)
       Call REPLA(TTXX$, Mid$(XX$, 2, 31), 5, 30)
       Call REPLA(TTXX$, FORMATNUM$(CVS(Mid$(XX$, 33, 4)), "F6.2"), 36, 6)
       List1.AddItem TTXX$
       PUNMAX = PUNMAX + CVS(Mid$(XX$, 33, 4))
     Next U&
     '
     Label2.Caption = TRIM$(FORMATNUM$(PUNMAX, "F6.2"))
   End If
   '
End Sub


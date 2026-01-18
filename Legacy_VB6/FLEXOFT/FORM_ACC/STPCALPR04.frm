VERSION 5.00
Begin VB.Form STPCALPR04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7035
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
   ScaleHeight     =   3900
   ScaleWidth      =   7035
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture15 
      Height          =   3930
      Left            =   5950
      ScaleHeight     =   3870
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   3900
         Left            =   0
         Picture         =   "STPCALPR04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6090
      Picture         =   "STPCALPR04.frx":2398
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Registra y Guarda la Configuración"
      Top             =   1155
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Escala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   210
      TabIndex        =   9
      Top             =   2835
      Width           =   3060
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "0 a 10 puntos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1890
         TabIndex        =   11
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "0 a 5 puntos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   525
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3360
      TabIndex        =   3
      Top             =   2835
      Width           =   2535
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puntaje Máximo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   315
         TabIndex        =   8
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   330
         Left            =   1340
         TabIndex        =   7
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Tabla de Indices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   5685
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
         Height          =   1635
         Left            =   210
         TabIndex        =   6
         Top             =   735
         Width           =   5250
      End
      Begin VB.PictureBox Picture1 
         Height          =   330
         Left            =   210
         ScaleHeight     =   270
         ScaleWidth      =   5190
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
      Begin VB.Line Line2 
         X1              =   5460
         X2              =   5460
         Y1              =   420
         Y2              =   2370
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   5460
         Y1              =   2370
         Y2              =   2370
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6090
      Picture         =   "STPCALPR04.frx":26A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1995
      Width           =   750
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6090
      Picture         =   "STPCALPR04.frx":29AC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5490
      Picture         =   "STPCALPR04.frx":2AF6
      Top             =   3210
      Width           =   2010
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "STPCALPR04"
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
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture15.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Image2.Visible = False
      MARCOBARRA.Top = 5800
      Set MARCOBARRA.Container = Me.Picture15
    End If

   If ULTREG&("ICALPROV") < 1 Then
     REBLA$ = REGBLAN$("ICALPROV")
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(1) + "Calidad Propiamente Dicha", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(2) + "Embalajes", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(3) + "Identificación", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(4) + "Entrega de Certificados", 1, 32): Call REGAPP("ICALPROV", ZZ$)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = RASOCLI$(NCMIN%)
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
     Label8 = RASOCLI$(NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub grabar_Click()
   Call Command2_Click
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

Private Sub salir_Click()
   Call Command1_Click
End Sub

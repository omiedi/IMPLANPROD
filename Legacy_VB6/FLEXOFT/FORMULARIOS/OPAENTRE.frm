VERSION 5.00
Begin VB.Form OPAENTRE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizaciones de Entregas - Configuración y Formularios"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7575
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
   ScaleHeight     =   2895
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tbls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2510
      Left            =   5565
      TabIndex        =   16
      Top             =   240
      Width           =   855
      Begin VB.CommandButton Command12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPAENTRE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Tabla de Embalajes Codificados"
         Top             =   1670
         Width           =   595
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPAENTRE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Tasas de I.V.A."
         Top             =   345
         Width           =   595
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "OPAENTRE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Tabla de Tipo de Cambio"
         Top             =   1010
         Width           =   595
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Autorizacion de Entregas"
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
      Width           =   5265
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2100
         TabIndex        =   13
         Top             =   1260
         Width           =   1800
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "App.Revision"
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
            Left            =   210
            TabIndex        =   15
            Top             =   575
            Width           =   1485
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nro. Original"
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
            Left            =   210
            TabIndex        =   14
            Top             =   260
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   210
         TabIndex        =   11
         Top             =   1260
         Width           =   1800
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Incluir Espe- cificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   210
            TabIndex        =   12
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3780
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   4095
         Picture         =   "OPAENTRE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   400
         Width           =   960
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3990
         TabIndex        =   3
         Top             =   1260
         Width           =   1080
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   4
            Top             =   630
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Man"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   5
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   2325
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
      Height          =   645
      Left            =   6615
      Picture         =   "OPAENTRE.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1050
      Width           =   700
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
      Height          =   645
      Left            =   6615
      Picture         =   "OPAENTRE.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   700
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6010
      Picture         =   "OPAENTRE.frx":0EBC
      Top             =   2220
      Width           =   2010
   End
End
Attribute VB_Name = "OPAENTRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARTAMONED
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command11.Enabled = False
   Call CARTAEMBAL
   Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMAENT", Text1.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEAENT", MONUME$)
    '
    MONUME$ = "NO"
    If Check1.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("AUTENTRE", "INCLUSPE", MONUME$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("AUTENTRE", "REPLAORI", MONUME$)
    '
    MONUME$ = "NO"
    If Check3.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("AUTENTRE", "REVICOD", MONUME$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'
Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command8_Click()
    Call CARDATCENCALC
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FORMAENT")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
    If CONTROL$("", "NUMEAENT") = "MANUAL" Then Option1.Value = True
    '
    If CONTROL$("AUTENTRE", "INCLUSPE") = "SI" Then Check1.Value = 1
    If CONTROL$("AUTENTRE", "REPLAORI") = "SI" Then Check2.Value = 1
    If CONTROL$("AUTENTRE", "REVICOD") = "SI" Then Check3.Value = 1
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      JU& = 0
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         If TRIM$(Mid$(X$, 2, 27)) <> "" Then
           JU& = JU& + 1
           Call REPLA(X$, Chr$(JU&), 1, 1)
           Call GRAREG("TAMONED", X$, JU&)
         End If
      Next U&
      Call SETULTREG("TAMONED", JU&)
   End If
   Call CIERRARCH("TAMONED")
End Sub

Sub CARTAEMBAL()
   Call TRALOCAL("TAEMBALA", "")
   VTB% = VENTABU%("TAEMBALA/NC/TITUPAN=Embalajes Tipificados")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAEMBALA")
         X$ = REGLEIDO$("!TAEMBALA", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAEMBALA", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAEMBALA")
      Call SETULTREG("TAEMBALA", FIN&)
   End If
   Call CIERRARCH("TAEMBALA")
End Sub


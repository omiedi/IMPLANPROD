VERSION 5.00
Begin VB.Form ANUREMI 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reversión Movimientos de Stock"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
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
   ScaleHeight     =   2820
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4220
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   4
      Top             =   1890
      Width           =   875
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   160
      TabIndex        =   3
      Top             =   210
      Width           =   3900
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1365
         TabIndex        =   6
         Text            =   " "
         Top             =   525
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   10
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Fecha de Ajuste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   160
      TabIndex        =   2
      Top             =   1470
      Width           =   3900
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1365
         TabIndex        =   9
         Text            =   " "
         Top             =   630
         Width           =   2010
      End
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
         Height          =   360
         Left            =   3360
         TabIndex        =   8
         Top             =   630
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   735
         Width           =   1170
      End
   End
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
      Height          =   585
      Left            =   4220
      Picture         =   "ANUREMI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton command10 
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
      Height          =   800
      Left            =   4220
      Picture         =   "ANUREMI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4050
      Picture         =   "ANUREMI.frx":0454
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   4725
   End
End
Attribute VB_Name = "ANUREMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   COREMI$ = TRIM$(MESAN.Text)
   If COREMI$ = "" Then GoTo 99
   '
   FAJUI% = FECHANUM(Text1.Text)
   FIN& = ULTREG&("MOVISTO")
   Call BLOQARCH("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MOVISTO", U&)
     If TRIM$(Mid$(X$, 23, 10)) = COREMI$ Then
       ENTRA$ = Mid$(X$, 88, 8)
       SALID$ = Mid$(X$, 96, 8)
       Call REPLA(X$, SALID$, 88, 8)
       Call REPLA(X$, ENTRA$, 96, 8)
       Call REPLA(X$, "Anulacion " + COREMI$, 43, 30)
       If FAJUI% > 0 Then
         Call REPLA(X$, MKI$(FAJUI%), 1, 2)
       End If
       Call REGAPP("MOVISTO", X$)
     End If
   Next U&
   Call CIERRARCH("MOVISTO")
   Call FINAL("")
   '
99 Command12.Enabled = True
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   Text1 = VALACT1$("SELFECHA")
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


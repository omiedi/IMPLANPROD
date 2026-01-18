VERSION 5.00
Begin VB.Form DESCACUECO 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depuracion Cuenta Corriente Acreedores"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   525
      ScaleHeight     =   75
      ScaleWidth      =   3315
      TabIndex        =   6
      Top             =   2205
      Width           =   3375
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -105
         TabIndex        =   7
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
      Height          =   585
      Left            =   3570
      Picture         =   "DESCACUECORR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   1000
      Width           =   750
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
      Height          =   620
      Left            =   3570
      Picture         =   "DESCACUECORR.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.CommandButton Command1 
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
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   1260
      Width           =   175
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
      Height          =   330
      Left            =   1155
      TabIndex        =   2
      Top             =   1260
      Width           =   1065
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
      Left            =   210
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mediante  este  utilitario  se eliminan au- tomaticamente los movimientos de cuen- ta corriente anteriores a la fecha elegida."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   2730
      Width           =   4320
   End
End
Attribute VB_Name = "DESCACUECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub BOTREC_Click()
    '
    HOO% = HOY(HO$)
    FLIM% = FECHANUM(Text1.TEXT)
    F11 = FLIM%: F22 = HOO%
    If FLIM% < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text1.TEXT = ""
        Text1.SetFocus
        Exit Sub
      ElseIf DIENTRE%(F11, F22) < 100 Then
        Call MENSERR(24, "No se Pueden Eliminar\Movimientos Tan Cercanos.\  \Elija Otra Fecha o Cancele el Proceso.")
        Text1.TEXT = ""
        Text1.SetFocus
        Exit Sub
      Else
        If InStr(EMPREAC$, "FERVI") < 1 Then
           Call OPNOIN("Descarga de Cuenta Corriente Acreedores")
           Exit Sub
        End If
        Call DESCACU
    End If
    '
End Sub

Private Sub Command1_Click()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Form_Activate()
    If TRIM$(Text1.TEXT) = "" Then
        Text1.TEXT = FECHATEX$(HOY(HO$))
    End If
    Text1.SetFocus
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
        BOTREC.SetFocus
    End If
End Sub

VERSION 5.00
Begin VB.Form SUSTICO 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sustitucion de Codigo"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8010
   ControlBox      =   0   'False
   FillColor       =   &H0080FF80&
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   390
      Left            =   4305
      TabIndex        =   11
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Sustico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Sustico.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Height          =   855
      Left            =   6960
      Picture         =   "Sustico.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   6960
      Picture         =   "Sustico.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Left            =   1995
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   4215
   End
   Begin VB.TextBox TEXT3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1995
      TabIndex        =   2
      Text            =   " "
      Top             =   1840
      Width           =   2325
   End
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
      Height          =   390
      Left            =   1995
      TabIndex        =   0
      Top             =   210
      Width           =   2325
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   420
      TabIndex        =   10
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo Código:"
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
      Left            =   315
      TabIndex        =   9
      Top             =   2005
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código Actual:"
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
      Left            =   -315
      TabIndex        =   8
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Sustico.frx":0A68
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   315
      TabIndex        =   7
      Top             =   3000
      Width           =   5895
   End
End
Attribute VB_Name = "SUSTICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COD$, CI%

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Text1_DBLCLICK()
    Text1.Text = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub
'

Private Sub COMMAND4_CLICK()
    Text1.Text = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub
'

Private Sub Text1_GotFocus()
    Refresh
    If Text1.Text = "" Then
        Text1.Text = VALACT1$("MASTER")
    End If
    CODX$ = TRIM$(Text1.Text)
    cix% = CODINT%(CODX$)
    If cix% > 0 Then
         COD$ = CODX$: CI% = cix%
         Text1.Text = CODX$
         Text2.Text = DESCRIT$(cix%)
         If TRIM$(TEXT3.Text) = "" Then
            TEXT3.SetFocus
         End If
       Else
         Text1.Text = COD$
    End If
End Sub
'

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CODX$ = TRIM$(Text1.Text)
        cix% = CODINT%(CODX$)
        If cix% > 0 Then
            COD$ = CODX$: CI% = cix%
            Text1.Text = CODX$
            Text2.Text = DESCRIT$(cix%)
            COD$ = CODX$: CI% = cix%
            Text2.Text = DESCRIT$(CI%)
            TEXT3.SetFocus
          Else
            Call MENSERR(24, "Código no Válido")
            Text1.Text = COD$
        End If
    End If
End Sub
'

Private Sub Text2_GotFocus()
    TEXT3.SetFocus
End Sub
'

Private Sub Text3_GotFocus()
    CODX$ = TRIM$(Text1.Text)
    cix% = CODINT%(CODX$)
    If cix% > 0 Then
          COD$ = CODX$: CI% = cix%
          Text2.Text = DESCRIT$(CI%)
          TEXT3.SetFocus
        Else
          Call MENSERR(24, "Código no Válido")
          Text1.Text = COD$
          Text1.SetFocus
    End If
End Sub
'

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TRIM$(TEXT3.Text) <> "" Then
            Command1.SetFocus
        End If
    End If
End Sub
'

Private Sub Command1_Click()
    '
    cod1$ = Text1.Text
    CODI% = CODINT%(cod1$)
    If CODI% = 0 Then
        SUSTICO.Refresh
        Call MENSERR(21, "Codigo Origen\Inexistente")
        Text1.Text = COD$
        Text1.SetFocus
        Exit Sub
      ElseIf CODI% < 0 Then
        SUSTICO.Refresh
        Call MENSERR(21, "Código Origen\de Baja")
        Text1.Text = COD$
        Text1.SetFocus
        Exit Sub
    End If
    '
    cod2$ = TEXT3.Text
    If TRIM$(cod2$) = "" Then
        SUSTICO.Refresh
        Call MENSERR(21, "Código Destino no Válido")
        TEXT3.SetFocus
        Exit Sub
    End If
    '
    CODI% = CODINT%(cod2$)
    If CODI% <> 0 Then
        SUSTICO.Refresh
        Call MENSERR(21, "Código Destino\ya Utilizado")
        TEXT3.SetFocus
        Exit Sub
    End If
    '
    Call GRACONTROL("", "AMASTO", "ABORT " + HOS$ + " - " + Time$ + " hs.")
    '
    Call SUSTICOX(cod1$, cod2$, SUCCESS%)
    Screen.MousePointer = 1
    '
    If SUCCESS% <> 1 Then
        SUSTICO.Refresh
        Call COMUNI("La Sustitución de Código\no pudo Completarse\   \Ejecute la Opción de Control.")
        Text1.SetFocus
        Exit Sub
      Else
        Call COMUNI("Sustitución de Código\'" + TRIM$(cod1$) + "' por '" + TRIM$(cod2$) + "'\Completa y Verificada")
        Call CIERRARCH("*.*")
        Text1.Text = TEXT3.Text
        Text2.Text = ""
        TEXT3.Text = ""
        Call GRACONTROL("", "AMASTO", "NORMAL")
        Call FRESHECHO("MASTER")
        Text1.SetFocus
    End If
    '
End Sub
'

Private Sub Command2_Click()
    TXT$ = "Ingrese  en la  Primera  Caja  de Entrada\"
    TXT$ = TXT$ + "el CODIGO ACTUAL del Item a Modificar.\"
    TXT$ = TXT$ + "Para  Seleccionar o  Buscar  de  la  Lista\"
    TXT$ = TXT$ + "'pinche'  con el  Mouse  sobre la Caja de\"
    TXT$ = TXT$ + "entrada.\"
    TXT$ = TXT$ + "   \"
    TXT$ = TXT$ + "Ingrese en la  Segunda  Caja de Entrada\"
    TXT$ = TXT$ + "el NUEVO CODIGO.\"
    TXT$ = TXT$ + "   \"
    TXT$ = TXT$ + "Pinche sobre <GRABAR> para Confirmar\"
    TXT$ = TXT$ + "la Sustitución de Código."
    Call COMUNI(TXT$)
    '
End Sub
'

Private Sub Command3_Click()
    Text1.Text = COD$
    TEXT3.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command9_Click()
    Call CIERRARCH("*.*")
    Hide
    Call FRESHALL
End Sub

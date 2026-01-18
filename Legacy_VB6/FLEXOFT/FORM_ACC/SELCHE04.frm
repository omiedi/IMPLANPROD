VERSION 5.00
Begin VB.Form SELCHE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Cheques a Depositar"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00008000&
      Height          =   3900
      Left            =   7140
      ScaleHeight     =   3840
      ScaleWidth      =   1845
      TabIndex        =   7
      Top             =   0
      Width           =   1905
      Begin VB.PictureBox Picture16 
         Height          =   3900
         Left            =   0
         ScaleHeight     =   3840
         ScaleWidth      =   1140
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         Begin VB.Image Image3 
            Height          =   3950
            Left            =   0
            Picture         =   "SELCHE04.frx":0000
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "SELCHE04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Búsqueda por Número, Banco, Importe, etc."
         Top             =   2310
         Width           =   680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "SELCHE04.frx":36FE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Conmuta Ordenamiento por Fecha / Importe"
         Top             =   1575
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "SELCHE04.frx":3848
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   680
      End
      Begin VB.CommandButton BOTREC 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "SELCHE04.frx":3B52
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   840
         Width           =   680
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   8
         Top             =   3100
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "SELCHE04.frx":3E5C
         Top             =   3360
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   540
      Left            =   105
      ScaleHeight     =   480
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   3220
      Width           =   6945
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ************"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4790
         TabIndex        =   14
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4790
         TabIndex        =   6
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Valores:"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   150
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1470
         TabIndex        =   4
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccionados:"
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
         Left            =   105
         TabIndex        =   3
         Top             =   150
         Width           =   1695
      End
   End
   Begin VB.ListBox LICHESEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   8500
      Y1              =   3765
      Y2              =   3765
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu ordenar 
         Caption         =   "Ordenar"
      End
      Begin VB.Menu continuar 
         Caption         =   "Continuar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
      Begin VB.Menu buscar 
         Caption         =   "Buscar"
      End
   End
End
Attribute VB_Name = "SELCHE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODORDE%
Dim MUESTRACAR%

Private Sub BOTEXIT_Click()
    LICHESEL.Clear
    Label2.Caption = ""
    Hide
End Sub
'

Private Sub BOTREC_Click()
    SUANT# = XVALO(Label2)
    Call SUCHESEL
    If XVALO(Label2) = SUANT# Then
        Hide
      Else
        Call MENSERR(24, "Revise Valores Seleccionados")
    End If
End Sub

Private Sub buscar_Click()
   Call Command2_Click
End Sub

Private Sub Command1_Click()
    If MODORDE% = 1 Then
      MODORDE% = 2
        For KI% = 1 To LICHESEL.ListCount
10        IMPOI# = XVALO(Mid$(LICHESEL.List(KI% - 1), 41, 12))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            IMPOJ# = XVALO(Mid$(LICHESEL.List(KJ% - 1), 41, 12))
            If IMPOJ# < IMPOI# Then
              TTXX$ = LICHESEL.List(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 10
            End If
          Next KJ%
        Next KI%
      ElseIf MODORDE% = 2 Then
        MODORDE% = 1
        For KI% = 1 To LICHESEL.ListCount
20        FECHI% = FECHANUM(Mid$(LICHESEL.List(KI% - 1), 33, 8))
          For KJ% = KI% + 1 To LICHESEL.ListCount
            FECHJ% = FECHANUM(Mid$(LICHESEL.List(KJ% - 1), 33, 8))
            If FECHJ% < FECHI% Then
              TTXX$ = LICHESEL.List(KJ% - 1)
              MARCA% = 0: If LICHESEL.Selected(KJ% - 1) = True Then MARCA% = 1
              LICHESEL.RemoveItem KJ% - 1
              LICHESEL.AddItem TTXX$, KI% - 1
              If MARCA% = 1 Then LICHESEL.Selected(KI% - 1) = True
              GoTo 20
            End If
          Next KJ%
        Next KI%
      
    End If
    '
End Sub

Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.List(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub continuar_Click()
   Call BOTREC_Click
End Sub

Private Sub Form_Activate()
    Call SUCHESEL
    MODORDE% = 1
    LICHESEL.SetFocus
End Sub

Sub SUCHESEL()
    SUCHESE# = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.List(i& - 1), 41, 12))
       End If
    Next i&
    Label2.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    Label2.Refresh
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3600
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    MODORDE% = 1
    MUESTRACAR% = 0
    If DERACCE%("VECARCHE/NOMESS", "Visualizar Cartera de Cheques y Valores") >= 2 Then
      MUESTRACAR% = 1
      Label5.Visible = False
    End If
    '
End Sub

Private Sub LICHESEL_Click()
   Call SUCHESEL
End Sub




Private Sub ordenar_Click()
    Call Command1_Click
End Sub

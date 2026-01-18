VERSION 5.00
Begin VB.Form DATEMPRE05 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de la Empresa"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000040&
      Height          =   6135
      Left            =   6165
      ScaleHeight     =   6075
      ScaleWidth      =   915
      TabIndex        =   11
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   90
         Picture         =   "DATEMPRE05.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ayuda en Línea"
         Top             =   810
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   90
         Picture         =   "DATEMPRE05.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Height          =   690
         Left            =   90
         Picture         =   "DATEMPRE05.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Guardar "
         Top             =   1680
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DATEMPRE05.frx":075E
         Top             =   3960
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   645
         Left            =   5670
         TabIndex        =   23
         Top             =   2520
         Width           =   330
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   8
         Top             =   4080
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   2235
         MaxLength       =   24
         TabIndex        =   7
         Top             =   3600
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   6
         Top             =   3120
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   5
         Top             =   2640
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2235
         MaxLength       =   24
         TabIndex        =   4
         Top             =   2160
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2235
         MaxLength       =   24
         TabIndex        =   3
         Top             =   1680
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2235
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2235
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   3450
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2235
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3450
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dato Complementario 2:"
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
         Index           =   8
         Left            =   0
         TabIndex        =   22
         Top             =   4185
         Width           =   2145
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dato Complementario 1:"
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
         Index           =   7
         Left            =   0
         TabIndex        =   21
         Top             =   3705
         Width           =   2145
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Inscripción:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   3225
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ex-Caja de Jubilación:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2745
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2265
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1785
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1305
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Fantasía:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   825
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Razon Social:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   345
         Width           =   1905
      End
   End
End
Attribute VB_Name = "DATEMPRE05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODIFIC%

Private Sub Command1_Click()
  '
  For J& = 1 To ULTREG("DATEMPRE")
  X$ = REGLEIDO$("DATEMPRE", J&)
    Select Case J&
      Case 1
        RASOEM$ = TRIM$(Mid$(X$, 1, 30))
        LOCAEM$ = TRIM$(Mid$(X$, 31, 24))
        CAJUEM$ = TRIM$(Mid$(X$, 55, 16))
      Case 2
        FANTEM$ = TRIM$(Mid$(X$, 1, 30))
        TELEM$ = TRIM$(Mid$(X$, 31, 24))
        NUIMEM$ = TRIM$(Mid$(X$, 55, 16))
      Case 3
        DOMIEM$ = TRIM$(Mid$(X$, 1, 30))
        DAT1EM$ = TRIM$(Mid$(X$, 31, 24))
        DAT2EM$ = TRIM$(Mid$(X$, 55, 16))
    End Select
  Next J&
  '
  Call REPLA(VDEF$, RASOEM$, 1, 30)
  Call REPLA(VDEF$, FANTEM$, 31, 30)
  Call REPLA(VDEF$, DOMIEM$, 61, 30)
  Call REPLA(VDEF$, LOCAEM$, 91, 24)
  Call REPLA(VDEF$, TELEM$, 115, 24)
  Call REPLA(VDEF$, CAJUEM$, 139, 16)
  Call REPLA(VDEF$, NUIMEM$, 155, 24)
  Call REPLA(VDEF$, DAT1EM$, 171, 24)
  Call REPLA(VDEF$, DAT2EM$, 195, 24)
  Z$ = OBJPLA$("DATAEMPRE", VDEF$)
  '
End Sub

Private Sub Command3_Click()
  '
  If MODIFIC% = 1 Then
    If COMALTER%("Se han Modificado Datos de la Empresa.\\Registrar Cambios\Cancelar") = 1 Then
        Screen.MousePointer = 11
        MODIFIC% = 0
        Call Command9_Click
        Screen.MousePointer = 1
      End If
  End If
  Call CIERRARCH("*.*")
  Unload Me
  '
End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  '
  RASOEM$ = Text25(0)
  FANTEM$ = Text25(1)
  DOMIEM$ = Text25(2)
  LOCAEM$ = Text25(3)
  TELEM$ = Text25(4)
  CAJUEM$ = Text25(5)
  NUIMEM$ = Text25(6)
  DAT1EM$ = Text25(7)
  DAT2EM$ = Text25(8)
  '
  Call BLANARCH("DATEMPRE")
  YY$ = Space$(128)
  Call REPLA(YY$, RASOEM$, 1, 30)
  Call REPLA(YY$, LOCAEM$, 31, 24)
  Call REPLA(YY$, CAJUEM$, 55, 16)
  Call REGAPP("DATEMPRE", YY$)
  '
  YY$ = Space$(128)
  Call REPLA(YY$, FANTEM$, 1, 30)
  Call REPLA(YY$, TELEM$, 31, 24)
  Call REPLA(YY$, NUIMEM$, 55, 16)
  Call REGAPP("DATEMPRE", YY$)
  '
  YY$ = Space$(128)
  Call REPLA(YY$, DOMIEM$, 1, 30)
  Call REPLA(YY$, DAT1EM$, 31, 24)
  Call REPLA(YY$, DAT2EM$, 55, 16)
  Call REGAPP("DATEMPRE", YY$)
  '
  Call CIERRARCH("DATEMPRE")
  MODIFIC% = 0
  '
  Command9.Enabled = True
End Sub

Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  MODIFIC% = 0
  '
End Sub
Private Sub Text25_Change(Index As Integer)
  '
  MODIFIC% = 1
  '
End Sub

Private Sub Text25_GotFocus(Index As Integer)
  '
  Text25(Index).SelStart = 0
  Text25(Index).SelLength = (Len(Text25(Index)))
  '
End Sub

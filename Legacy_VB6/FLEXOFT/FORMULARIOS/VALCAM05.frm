VERSION 5.00
Begin VB.Form VALCAM05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valores de Cambio para Pago"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   6135
      Left            =   3195
      ScaleHeight     =   6075
      ScaleWidth      =   915
      TabIndex        =   12
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
         Picture         =   "VALCAM05.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda en Línea"
         Top             =   930
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
         Picture         =   "VALCAM05.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit"
         Top             =   240
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
         Picture         =   "VALCAM05.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guardar "
         Top             =   1800
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "VALCAM05.frx":075E
         Top             =   4420
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   -680
      TabIndex        =   11
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   330
         Left            =   3570
         TabIndex        =   25
         Top             =   1890
         Width           =   330
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   9
         Top             =   4560
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
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
         Width           =   1290
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro.10:"
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
         Index           =   9
         Left            =   0
         TabIndex        =   24
         Top             =   4665
         Width           =   2145
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 9:"
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
         TabIndex        =   23
         Top             =   4185
         Width           =   2145
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 8:"
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
         TabIndex        =   22
         Top             =   3705
         Width           =   2145
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 7:"
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
         TabIndex        =   21
         Top             =   3225
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 6:"
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
         TabIndex        =   20
         Top             =   2745
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 5:"
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
         TabIndex        =   19
         Top             =   2265
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 4:"
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
         TabIndex        =   18
         Top             =   1785
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 3:"
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
         TabIndex        =   17
         Top             =   1305
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 2:"
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
         TabIndex        =   16
         Top             =   825
         Width           =   1905
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Nro. 1:"
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
         TabIndex        =   15
         Top             =   345
         Width           =   1905
      End
   End
End
Attribute VB_Name = "VALCAM05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODIFIC%

Private Sub Command1_Click()
  '
  Dim VALBILL(9)
  '
  For J& = 1 To ULTREG("BILLETES")
  X$ = REGLEIDO$("BILLETES", J&)
    '
    For JJ& = 1 To 40 Step 4
      VAL1 = CVS(Mid$(X$, JJ&, 4)) 'CADA 4 BYTE UN BILLETE
      VALBILL(CONTI1%) = VAL1
      CONTI1% = CONTI1% + 1
      ZJ% = Int(JJ&)
      Call REPLA(VDEF$, MKS$(VAL1), ZJ%, 4)
    Next JJ&
    '
  Next J&
  '
  OBJ1$ = OBJPLA$("VALCAMPA", VDEF$)
  '
  VALBILL(0) = Val(CVS(Mid$(OBJ1$, 1, 4)))
  VALBILL(1) = Val(CVS(Mid$(OBJ1$, 5, 4)))
  VALBILL(2) = Val(CVS(Mid$(OBJ1$, 9, 4)))
  VALBILL(3) = Val(CVS(Mid$(OBJ1$, 13, 4)))
  VALBILL(4) = Val(CVS(Mid$(OBJ1$, 17, 4)))
  VALBILL(5) = Val(CVS(Mid$(OBJ1$, 21, 4)))
  VALBILL(6) = Val(CVS(Mid$(OBJ1$, 25, 4)))
  VALBILL(7) = CVS(Mid$(OBJ1$, 29, 4))
  VALBILL(8) = CVS(Mid$(OBJ1$, 33, 4))
  VALBILL(9) = CVS(Mid$(OBJ1$, 37, 4))
  '
  Call BLANARCH("BILLETES")
  YY$ = Space$(40)
  'CADA 4 BYTES GRABO EL VALOR DE UN BILLETE
  For JJ& = 1 To 40 Step 4
    Call REPLA(YY$, MKS$(VALBILL(CONTI)), Int(JJ& * 1), 4)
    CONTI = CONTI + 1
  Next JJ&
  Call REGAPP("BILLETES", YY$)
  '
  Call CIERRARCH("BILLETES")
  '
   
End Sub

Private Sub Command3_Click()
  '
  If MODIFIC% = 1 Then
    If COMALTER%("Se han Modificado Datos de Billetes de Cambio para Pago.\\Registrar Cambios\Cancelar") = 1 Then
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
  Dim VALBILL(9)
  '
  Command9.Enabled = False
  '
  For J% = 0 To 9
    VALBILL(J%) = Val(Text25(J%))
  Next J%
  '
  Call BLANARCH("BILLETES")
  YY$ = Space$(40)
  'CADA 4 BYTES GRABO EL VALOR DE UN BILLETE
  For JJ& = 1 To 40 Step 4
    Call REPLA(YY$, MKS$(VALBILL(CONTI)), Int(JJ& * 1), 4)
    CONTI = CONTI + 1
  Next JJ&
  Call REGAPP("BILLETES", YY$)
  '
  Call CIERRARCH("BILLETES")
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

Private Sub Text25_LostFocus(Index As Integer)
  '
  Text25(Index) = FORMATNUM(Val(Text25(Index)), "F10.2")
  '
End Sub
Private Sub Text25_GotFocus(Index As Integer)
  '
  Text25(Index).SelStart = 0
  Text25(Index).SelLength = (Len(Text25(Index)))
  '
End Sub

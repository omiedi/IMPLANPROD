VERSION 5.00
Begin VB.Form ALTAITEM 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta Nuevo Item de Stock"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
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
      Left            =   10150
      TabIndex        =   47
      Top             =   4095
      Width           =   175
   End
   Begin VB.CommandButton Command23 
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
      Left            =   7455
      TabIndex        =   46
      Top             =   4095
      Width           =   175
   End
   Begin VB.CommandButton Command18 
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
      Left            =   8505
      TabIndex        =   45
      Top             =   1785
      Width           =   175
   End
   Begin VB.CommandButton Command20 
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
      Left            =   2520
      TabIndex        =   44
      Top             =   4095
      Width           =   175
   End
   Begin VB.CommandButton Command26 
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
      Left            =   2205
      TabIndex        =   43
      Top             =   2940
      Width           =   175
   End
   Begin VB.CommandButton Command16 
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
      Left            =   5145
      TabIndex        =   42
      Top             =   1785
      Width           =   175
   End
   Begin VB.CommandButton Command14 
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
      Left            =   2205
      TabIndex        =   41
      Top             =   1785
      Width           =   175
   End
   Begin VB.TextBox Text20 
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
      Left            =   8925
      TabIndex        =   12
      Top             =   2940
      Width           =   1410
   End
   Begin VB.TextBox Text19 
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
      Left            =   8925
      TabIndex        =   13
      Top             =   3360
      Width           =   1410
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Especificaciones"
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
      Left            =   3675
      Picture         =   "ALTAITEM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5250
      Width           =   1800
   End
   Begin VB.TextBox Text18 
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
      Left            =   5670
      TabIndex        =   11
      Top             =   2940
      Width           =   1410
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox Text16 
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
      Left            =   1575
      TabIndex        =   9
      Text            =   "  "
      Top             =   2940
      Width           =   675
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7140
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4620
      Width           =   3195
   End
   Begin VB.TextBox Text14 
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
      Left            =   8925
      TabIndex        =   19
      Top             =   4095
      Width           =   1250
   End
   Begin VB.TextBox Text13 
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
      Left            =   7140
      TabIndex        =   17
      Text            =   " "
      Top             =   4095
      Width           =   360
   End
   Begin VB.TextBox Text12 
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
      Left            =   4410
      TabIndex        =   16
      Top             =   4095
      Width           =   1620
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4620
      Width           =   4455
   End
   Begin VB.TextBox Text10 
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
      Left            =   1575
      TabIndex        =   14
      Text            =   " "
      Top             =   4095
      Width           =   990
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8190
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2205
      Width           =   2145
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4830
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   " "
      Top             =   2205
      Width           =   2250
   End
   Begin VB.TextBox Text8 
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
      Left            =   8190
      TabIndex        =   7
      Text            =   " "
      Top             =   1785
      Width           =   360
   End
   Begin VB.TextBox Text6 
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
      Left            =   4830
      TabIndex        =   5
      Text            =   " "
      Top             =   1785
      Width           =   360
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2205
      Width           =   2250
   End
   Begin VB.TextBox Text4 
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
      Left            =   1575
      TabIndex        =   3
      Text            =   "  "
      Top             =   1785
      Width           =   675
   End
   Begin VB.TextBox Text3 
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
      Left            =   1575
      TabIndex        =   2
      Text            =   " "
      Top             =   1050
      Width           =   8760
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Left            =   105
      TabIndex        =   35
      Top             =   5145
      Width           =   2535
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   210
         Picture         =   "ALTAITEM.frx":030A
         Stretch         =   -1  'True
         Top             =   105
         Width           =   8925
      End
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
      Left            =   6300
      Picture         =   "ALTAITEM.frx":13470
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5250
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7350
      Picture         =   "ALTAITEM.frx":1377A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5250
      Width           =   855
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
      Height          =   855
      Left            =   9450
      Picture         =   "ALTAITEM.frx":13A84
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5250
      Width           =   855
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8400
      Picture         =   "ALTAITEM.frx":13BCE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5250
      Width           =   855
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
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   2565
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9135
      TabIndex        =   1
      Top             =   315
      Width           =   1170
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Máximo:"
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
      Left            =   7140
      TabIndex        =   40
      Top             =   3045
      Width           =   1665
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Económico:"
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
      Left            =   7035
      TabIndex        =   39
      Top             =   3465
      Width           =   1770
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Contable:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   38
      Top             =   2835
      Width           =   1290
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      TabIndex        =   37
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
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
      Left            =   5880
      TabIndex        =   36
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Criterio de Reposicion."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6720
      TabIndex        =   34
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Material:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3360
      TabIndex        =   33
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   120
      Left            =   3900
      TabIndex        =   32
      Top             =   4305
      Width           =   15
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4410
      TabIndex        =   31
      Top             =   2835
      Width           =   1185
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Medida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   30
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Costeo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7665
      TabIndex        =   29
      Top             =   3990
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Costo de Reposición:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2940
      TabIndex        =   28
      Top             =   3990
      Width           =   1290
   End
   Begin VB.Label Label3 
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
      Left            =   0
      TabIndex        =   27
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   7875
      TabIndex        =   26
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      Left            =   585
      TabIndex        =   25
      Top             =   420
      Width           =   795
   End
End
Attribute VB_Name = "ALTAITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PODESCARGA%

Private Sub Command1_Click()
    '
    FEX = HOY(HOS$)
    Call GRACONTROL("", "AMASTO", "ABORT " + HOS$ + " - " + Time$ + " hs.")
    '
    COD$ = AJUSTI$(Text1.Text, 16)
    If TRIM$(COD$) = "" Then
        Call MENSERR(24, "Código no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    If CODINT%(COD$) <> 0 Then
        Call MENSERR(24, "Codigo Existente en Maestro")
        Text1.SetFocus
        Exit Sub
    End If
    '
    DE$ = TRIM$(TEXT3.Text)
    If DE$ = "" Then
        Call MENSERR(24, "Falta Descripción del Item")
        TEXT3.SetFocus
        Exit Sub
    End If
    '
    VALOR = Int(Val(Text6.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 15
    TIMA% = Val(Text6.Text)
    If TIMA% > 0 Then
        If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
15          TIMA% = 0
            Text6.Text = ""
            Text7.Text = ""
            Call MENSERR(24, "Tipo de Material no Válido")
            Text6.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text8.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 25
    CRIT% = Val(Text8.Text)
    If CRIT% > 0 Then
        If ECOARCH$("TACRIRE", Chr$(CRIT%)) = "" Then
25          CRIT% = 0
            Text8.Text = ""
            Text9.Text = ""
            Call MENSERR(24, "Criterio de Reposición no Válido")
            Text8.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text13.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 35
    MONEI% = Val(Text13.Text)
    If MONEI% > 0 Then
        If ECOARCH$("TAMONED", Chr$(MONEI%)) = "" Then
35          MONEI% = 1
            Text13.Text = "1"
            Text15.Text = ECOARCH$("TAMONED", Chr$(1))
            Call MENSERR(24, "Moneda de Costeo no Válida")
            Text13.SetFocus
            Exit Sub
        End If
    End If
    '
    STMIN = Val(Text18.Text)
    If STMIN <= 30000 Then
          STMI& = STMIN
        ElseIf STMIN <= 3000000 Then
          STMI& = (-1) * STMIN / 100
        Else
          Call MENSERR(24, "Stock Mínimo fuera de Rango")
          Text18.Text = ""
          Text18.SetFocus
          Exit Sub
    End If
    STMII% = STMI&
    '
    LOREP = Val(Text19.Text)
    If LOREP <= 30000 Then
          LORE& = LOREP
        ElseIf LOREP <= 3000000 Then
          LORE& = (-1) * LOREP / 100
        Else
          Call MENSERR(24, "Lote de Reposición fuera de Rango")
          Text19.Text = ""
          Text19.SetFocus
          Exit Sub
    End If
    LOREI% = LORE&
    '
    LOTEMA = Val(Text20.Text)
    If LOTEMA <= 30000 Then
          LOMA& = LOTEMA
        ElseIf LOTEMA <= 3000000 Then
          LOMA& = (-1) * LOTEMA / 100
        Else
          Call MENSERR(24, "Lote de Reposición fuera de Rango")
          Text20.Text = ""
          Text20.SetFocus
          Exit Sub
    End If
    LOTEMAI% = LOMA&
    '
    UNID$ = Text4.Text
    COU = Val(Text12.Text)
    FECO% = FECHANUM(Text14.Text)
    '
    VALOR = Int(Val(Text10.Text))
    If VALOR < 0 Or VALOR > 32767 Then GoTo 85
    NPRO% = Val(Text10.Text)
    If NPRO% > 0 Then
        If ECOARCH$("PROVED1", MKI$(NPRO%)) = "" Then
85          NPRO% = 0
            Text10.Text = ""
            Text11.Text = ""
            Call MENSERR(24, "Proveedor Inexistente")
            Text10.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text16.Text))
    If VALOR < 0 Or VALOR > 32767 Then GoTo 95
    GRUCO% = Val(Text16.Text)
    If GRUCO% > 0 Then
        If ECOARCH$("GRUCOVEN", MKI$(GRUCO%)) = "" Then
95          GRUCO% = 0
            Text16.Text = ""
            Text17.Text = ""
            Call MENSERR(24, "Grupo Contable de Venta no Válido")
            Text16.SetFocus
            Exit Sub
        End If
    End If
    '
    Call ALTAMAS
    Text1.Text = ""
    Text1.SetFocus
    '
End Sub

Private Sub Command14_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("UMEDIDA")
    Text4.Text = VALACT1$("UMEDIDA")
    Text4.Refresh
    Text5.Text = VALACT2$("UMEDIDA")
    Text5.Refresh
End Sub

Private Sub Command16_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TATIMAT")
    Text6.Text = VALACT1$("TATIMAT")
    Text6.Refresh
    Text7.Text = VALACT2$("TATIMAT")
    Text7.Refresh
End Sub

Private Sub Command18_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACRIRE")
    Text8.Text = VALACT1$("TACRIRE")
    Text8.Refresh
    Text9.Text = VALACT2$("TACRIRE")
    Text9.Refresh
End Sub

Private Sub Command2_Click()
    Call CLEARFIELDS
    Label22.Caption = Str$(ULTREG&("MASTER"))
    Text1.SetFocus
End Sub

Private Sub Command20_Click()
    If VALICOD% <> 1 Then
       Text1.SetFocus
       Exit Sub
    End If
    Call SELECHO("PROVED1")
    Text10.Text = VALACT1$("PROVED1")
    Text10.Refresh
    Text11.Text = VALACT2$("PROVED1")
    Text11.Refresh
End Sub

Private Sub Command23_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TAMONED")
    Text13.Text = VALACT1$("TAMONED")
    Text13.Refresh
    Text15.Text = VALACT2$("TAMONED")
    Text15.Refresh
End Sub

Private Sub Command26_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("GRUCOVEN")
    Text16.Text = VALACT1$("GRUCOVEN")
    Text16.Refresh
    Text17.Text = VALACT2$("GRUCOVEN")
    Text17.Refresh
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    '
    Hide
    Call FRESHALL
    '
    Call BLOQARCH("MASTER")
    Call BLOQARCH("INVERT")
    Screen.MousePointer = 11
    '
    NM% = ULTREG&("MASTER") - 1
    If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
        X$ = REGLEIDO$("MASTER", 1&)
        Call REPLA(X$, MKI$(NM%), 1, 2)
        Call GRAREG("MASTER", X$, 1&)
    End If
    '
    Screen.MousePointer = 11
    '
    FINVE% = FILENBR%("INVERT")
    Dim VL As String * 18
    FIECO% = FILENBR%("ECOMAST")
    Dim EL As String * 64
    FICIA% = FILENBR%("CIACTI")
    Dim CL As String * 2
    FIMAS% = FILENBR%("MASTER")
    Dim ML As String * 128
    '       '
    J1& = 1: J2& = 0: J3& = 1: ATX% = 0
    FIN& = AMASTOK.List1.ListCount
    For II& = 1 To FIN&
        Call TRACE(1, II&, FIN&)
        CODD$ = Mid$(AMASTOK.List1.List(II& - 1), 2, 16)
        CIIX% = Val(Mid$(AMASTOK.List1.List(II& - 1), 18))
        If CODD$ <> Left$(VL$, 16) Then ATX% = 1
        If CIIX% < 0 Then ATX% = 1
        J1& = J1& + 1
        If ATX% = 1 Then
           ZZ1$ = Space$(18)
           Call REPLA(ZZ1$, CODD$, 1, 16)
           Call REPLA(ZZ1$, MKI$(Abs(CIIX%)), 17, 2)
           LSet VL$ = ZZ1$: Put #FINVE%, J1&, VL$
        End If
        '
        If Mid$(CODD$, 16, 1) <> Chr$(96) Then
           If TRIM$(CODD$) <> "" Then
              J2& = J2& + 1
              J3& = J3& + 1
              If J2& > LOF(FIECO%) / 64 Then ATX% = 1
              '
              If ATX% = 1 Then
                 CI% = Abs(CIIX%)
                 REGMA& = CI% + 1
                 Get #FIMAS%, REGMA&, ML$
                 DENO$ = Mid$(ML$, 5, 64)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, MKI$(CI%), 63, 2)
                 Call REPLA(EC$, CODD$, 1, 16)
                 Call REPLA(EC$, DENO$, 17, 46)
                 '
                 LSet EL$ = EC$: Put #FIECO%, J2&, EL$
                 LSet CL$ = MKI$(CI%): Put #FICIA%, J3&, CL$
                 '
              End If
           End If
        End If
    Next II&
    '
    ZZ$ = String$(18, 0)
    NVI% = AMASTOK.List1.ListCount
    Call REPLA(ZZ$, MKI$(NVI%), 17, 2)
    LSet VL$ = ZZ$: Put #FINVE%, 1, VL$
    Call SETULTREG("INVERT", J1&)
    Call CIERRARCH("INVERT")
    '
    Call SETULTREG("ECOMAST", J2&)
    Call CIERRARCH("ECOMAST")
    '
    NVJ% = J2&
    LSet CL$ = MKI$(NVJ%): Put #FICIA%, 1, CL$
    Call SETULTREG("CIACTI", J3&)
    Call CIERRARCH("CIACTI")
    '
    Call GRACONTROL("", "AMASTO", "NORMAL")
    '
    Call FRESHECHO("MASTER")
    Call FRESHECHO("ECOMAST")
    '
    Call LIBARCH("AMASTO")
    Screen.MousePointer = 1
    '
    If AGREPRECIO% = 1 Then
      Call COMUNI("Recuerde Ordenar Oportunamente\la Lista de Precios Principal.")
      AGREPRECIO% = 0
    End If
    '
    PODESCARGA% = 0
    Unload ALTAITEM
    '
    If COMMA$ = "NEWITEM" Then
        Call FINAL("")
      Else
        Call FRESHALL
    End If
    '
End Sub

Private Sub COMMAND4_CLICK()
    Call HELPONLINE("AMASTO")
End Sub

Private Sub Command5_Click()
    Call OPNOIN("")
End Sub

Private Sub Command6_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
         Text14.Text = VDEVU$
       Else
         Text14.Text = HO$
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    Cancel = PODESCARGA%
    '
End Sub

Private Sub Text1_GotFocus()
    VALICOD% = 0
    If TRIM$(Text1.Text) = "" Then
        Call CLEARFIELDS
    End If
    Label22.Caption = Str$(ULTREG&("MASTER"))
    VACONTRO% = 0
End Sub

Private Sub Text1_LostFocus()
    CODY$ = Text1.Text
    COD$ = AJUSTI$(CODY$, 16)
    If TRIM$(CODY$) = "" Then
         Command3.SetFocus
    End If
End Sub
Private Sub Text10_DblClick()
    If VALICOD% <> 1 Then
       Text1.SetFocus
       Exit Sub
    End If
    Call SELECHO("PROVED1")
    Text10.Text = VALACT1$("PROVED1")
    Text10.Refresh
    Text11.Text = VALACT2$("PROVED1")
    Text11.Refresh
End Sub

Private Sub Text11_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text10.SetFocus
End Sub

Private Sub Text14_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
         Text14.Text = VDEVU$
    End If
End Sub

Private Sub Text15_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text13.SetFocus
End Sub

Private Sub Text16_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("GRUCOVEN")
    Text16.Text = VALACT1$("GRUCOVEN")
    Text16.Refresh
    Text17.Text = VALACT2$("GRUCOVEN")
    Text17.Refresh
End Sub

Private Sub Text17_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text16.SetFocus
End Sub

Private Sub Text19_Change()
    If Val(Text19.Text) > Val(Text20.Text) Then
        Text20.Text = Text19.Text
    End If
End Sub

Private Sub Text3_GotFocus()
    Call ALTAIT(COD$)
End Sub

Private Sub Text3_CLICK()
    Call ALTAIT(COD$)
End Sub

Private Sub Text4_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text6_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text8_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text12_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text13_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text14_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text16_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text18_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text19_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text10_Click()
    Call ALTAIT(COD$)
End Sub

Private Sub Text4_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("UMEDIDA")
    ALTAITEM.Refresh
    Text4.Text = VALACT1$("UMEDIDA")
    Text4.Refresh
    Text5.Text = VALACT2$("UMEDIDA")
    Text5.Refresh
End Sub

Private Sub Text5_click()
    Call ALTAIT(COD$)
    Text4.SetFocus
End Sub

Private Sub Text6_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TATIMAT")
    Text6.Text = VALACT1$("TATIMAT")
    Text6.Refresh
    Text7.Text = VALACT2$("TATIMAT")
    Text7.Refresh
End Sub

Private Sub Text7_click()
    Call ALTAIT(COD$)
    Text6.SetFocus
End Sub

Private Sub Text8_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACRIRE")
    Text8.Text = VALACT1$("TACRIRE")
    Text8.Refresh
    Text9.Text = VALACT2$("TACRIRE")
    Text9.Refresh
End Sub
Private Sub Text13_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TAMONED")
    Text13.Text = VALACT1$("TAMONED")
    Text13.Refresh
    Text15.Text = VALACT2$("TAMONED")
    Text15.Refresh
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    PODESCARGA% = 1
    Call BLOQARCH("AMASTO")
    '
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Call FINAL("")
    End If
    '
    Screen.MousePointer = 11
    AMASTOK.List1.Clear
    Dim II As String * 18
    N1% = FILENBR%("INVERT")
    FIN& = NUINV%
    For U& = 1 To FIN&
      Call TRACE(1, U&, FIN&)
      Get #N1%, U& + 1, II$
      ZZ$ = "A" + Left$(II$, 16) + CSTRING$(Mid$(II$, 17, 2), 1, 7, 0, 2)
      'AMASTOK.List1.AddItem ZZ$
      Call AGRELISTA(ZZ$)
    Next U&
    Screen.MousePointer = 1
    '
    Show
    Refresh
    '
End Sub


Private Sub Text9_click()
    Call ALTAIT(COD$)
    Text8.SetFocus
End Sub

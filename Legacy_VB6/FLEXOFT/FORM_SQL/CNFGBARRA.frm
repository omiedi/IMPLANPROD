VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CNFGBARRA 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Lectura de Códigos de Barra"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   3360
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   4890
      Left            =   3990
      ScaleHeight     =   4830
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command4 
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
         Height          =   650
         Left            =   120
         Picture         =   "CNFGBARRA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   945
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   8
         Top             =   4020
         Width           =   650
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
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   120
         Picture         =   "CNFGBARRA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "CNFGBARRA.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4380
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Configuración Lectura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4780
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   3855
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Opción de Lectura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   105
         TabIndex        =   22
         Top             =   3960
         Width           =   3585
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Long. Fija"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   24
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Por Caracter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1785
            TabIndex        =   23
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Caracter Separador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   105
         TabIndex        =   20
         Top             =   3135
         Width           =   3585
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   4
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Caracter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   21
            Top             =   315
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Longitud Fija"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   105
         TabIndex        =   13
         Top             =   300
         Width           =   3600
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Lote (Opcional)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   3375
            Begin VB.TextBox Text7 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2625
               TabIndex        =   27
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1050
               TabIndex        =   26
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Hasta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1995
               TabIndex        =   29
               Top             =   400
               Width           =   555
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Desde:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   315
               TabIndex        =   28
               Top             =   400
               Width           =   660
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   105
            TabIndex        =   17
            Top             =   1020
            Width           =   3375
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1050
               TabIndex        =   2
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2625
               TabIndex        =   3
               Top             =   315
               Width           =   615
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Desde:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   315
               TabIndex        =   19
               Top             =   400
               Width           =   660
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Hasta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1995
               TabIndex        =   18
               Top             =   400
               Width           =   555
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   105
            TabIndex        =   14
            Top             =   240
            Width           =   3375
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2625
               TabIndex        =   1
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1050
               TabIndex        =   0
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Hasta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1995
               TabIndex        =   16
               Top             =   400
               Width           =   555
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Desde:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   315
               TabIndex        =   15
               Top             =   405
               Width           =   660
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CNFGBARRA.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CNFGBARRA.frx":25AA
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "CNFGBARRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Unload Me
   Command1.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   'VALIDACIONES
   If Option1.Value = False And Option5.Value = False Then
     Call COMUNI("Falta Seleccionar una Opción de Lectura")
     GoTo 99
   End If
   '
   If Option1.Value = True Then
     If XVALO(Text1) < 1 Or XVALO(Text1) > 32767 Then Call COMUNI("Posición Inicial de Código No Válida"): GoTo 99
     If XVALO(Text2) < 1 Or XVALO(Text1) > 32767 Then Call COMUNI("Posición Final de Código No Válida"): GoTo 99
     If XVALO(Text4) < 1 Or XVALO(Text1) > 32767 Then Call COMUNI("Posición Inicial de Cantidad No Válida"): GoTo 99
     If XVALO(Text3) < 1 Or XVALO(Text1) > 32767 Then Call COMUNI("Posición Final de Cantidad No Válida"): GoTo 99
     If TRIM$(Text6) <> "" And (XVALO(Text1) < 1 Or XVALO(Text1) > 32767) Then Call COMUNI("Posición Inicial de Lote No Válida"): GoTo 99
     If TRIM$(Text7) <> "" And (XVALO(Text2) < 1 Or XVALO(Text1) > 32767) Then Call COMUNI("Posición Final de Lote No Válida"): GoTo 99

     '
     If XVALO(Text1) > XVALO(Text2) Then Call COMUNI("Posición de Código No Válida. Desde es Mayor a Hasta"): GoTo 99
     If XVALO(Text4) > XVALO(Text3) Then Call COMUNI("Posición de Cantidad No Válida. Desde es Mayor a Hasta"): GoTo 99
     If TRIM$(Text6) <> "" Then
      If XVALO(Text6) > XVALO(Text7) Then Call COMUNI("Posición de Lorw No Válida. Desde es Mayor a Hasta"): GoTo 99
     End If
   End If
   '
   If Option5.Value = True Then
     If TRIM$(Text5) = "" Then Call COMUNI("Falta Ingresar Caracter Separador"): GoTo 99
   End If
   '
   'POSICIONES DE LECTURA DE CODIGO Y CANTIDAD
   Call GRACONTROL("LECCODBA", "CODIINIC", Text1)
   Call GRACONTROL("LECCODBA", "CODIGFIN", Text2)
   Call GRACONTROL("LECCODBA", "CANTINIC", Text4)
   Call GRACONTROL("LECCODBA", "CANTIFIN", Text3)
   Call GRACONTROL("LECCODBA", "LOTEINIC", Text6)
   Call GRACONTROL("LECCODBA", "LOTEIFIN", Text7)
   
   '
   'POSICION DE CARACTER.
   Call GRACONTROL("LECCODBA", "CARASEPA", TRIM$(Text5))
   '
   'OPCION DE LECTURA
   OPCIONX$ = "LONGFIJA"
   If Option5.Value = True Then OPCIONX$ = "CARACTER"
   Call GRACONTROL("LECCODBA", "OPLECTUR", OPCIONX$)
   '
   Unload Me
99 Command4.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call APLICACIONSKINS(Me, Picture1)
   Text1 = CONTROL("LECCODBA", "CODIINIC") 'DESDE
   Text2 = CONTROL("LECCODBA", "CODIGFIN") 'HASTA
   '
   Text4 = CONTROL("LECCODBA", "CANTINIC") 'DESDE
   Text3 = CONTROL("LECCODBA", "CANTIFIN") 'HASTA
   '
   Text6 = CONTROL("LECCODBA", "LOTEINIC")
   Text7 = CONTROL("LECCODBA", "LOTEIFIN")
   
   Text5 = TRIM$(CONTROL("LECCODBA", "CARASEPA"))
   '
   OPCIONX$ = CONTROL("LECCODBA", "OPLECTUR")  'OPCION DE LECTURA
   If OPCIONX$ = "LONGFIJA" Then Option1.Value = True
   If OPCIONX$ = "CARACTER" Then Option5.Value = True
   
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        SendKeys "{TAB}"
      End If

End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        Text1.SetFocus
      End If
End Sub


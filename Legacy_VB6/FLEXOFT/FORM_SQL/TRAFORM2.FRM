VERSION 5.00
Begin VB.Form traform2 
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   525
      Left            =   8310
      TabIndex        =   67
      Top             =   45
      Width           =   915
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Contactos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   8430
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10620
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   6705
         Left            =   45
         TabIndex        =   1
         Top             =   525
         Width           =   10485
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   540
            Left            =   1560
            TabIndex        =   68
            Top             =   3090
            Width           =   825
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   960
            LargeChange     =   7
            Left            =   10360
            TabIndex        =   12
            Top             =   0
            Width           =   120
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            Height          =   5595
            Left            =   30
            TabIndex        =   2
            Top             =   0
            Width           =   10350
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   4
               Left            =   9180
               ScaleHeight     =   2700
               ScaleWidth      =   4230
               TabIndex        =   57
               Top             =   0
               Width           =   4230
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   66
                  Top             =   2400
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   65
                  Top             =   2100
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   64
                  Top             =   1800
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   63
                  Top             =   1500
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   62
                  Top             =   1200
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   61
                  Top             =   900
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   60
                  Top             =   0
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   59
                  Top             =   300
                  Width           =   4100
               End
               Begin VB.TextBox Text30 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   58
                  Top             =   600
                  Width           =   4100
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   3
               Left            =   7290
               ScaleHeight     =   2700
               ScaleWidth      =   4230
               TabIndex        =   47
               Top             =   0
               Width           =   4230
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   56
                  Top             =   2400
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   55
                  Top             =   2100
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   54
                  Top             =   1800
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   53
                  Top             =   1500
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   52
                  Top             =   1200
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   51
                  Top             =   900
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   50
                  Top             =   0
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   49
                  Top             =   300
                  Width           =   4100
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   48
                  Top             =   600
                  Width           =   4100
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   2
               Left            =   5085
               ScaleHeight     =   2700
               ScaleWidth      =   5175
               TabIndex        =   37
               Top             =   0
               Width           =   5175
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   46
                  Top             =   2400
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   45
                  Top             =   2100
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   44
                  Top             =   1800
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   43
                  Top             =   1500
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   42
                  Top             =   1200
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   41
                  Top             =   900
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  TabIndex        =   40
                  ToolTipText     =   "Doble-Click para Enviar un Mail"
                  Top             =   0
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   39
                  Top             =   300
                  Width           =   5050
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   38
                  Top             =   600
                  Width           =   5050
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   1
               Left            =   3615
               ScaleHeight     =   2700
               ScaleWidth      =   5175
               TabIndex        =   27
               Top             =   0
               Width           =   5175
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   2
                  Left            =   105
                  TabIndex        =   36
                  Top             =   600
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   1
                  Left            =   105
                  TabIndex        =   35
                  Top             =   300
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   0
                  Left            =   105
                  TabIndex        =   34
                  Top             =   0
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   3
                  Left            =   105
                  TabIndex        =   33
                  Top             =   900
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   4
                  Left            =   105
                  TabIndex        =   32
                  Top             =   1200
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   5
                  Left            =   105
                  TabIndex        =   31
                  Top             =   1500
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   6
                  Left            =   105
                  TabIndex        =   30
                  Top             =   1800
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   7
                  Left            =   105
                  TabIndex        =   29
                  Top             =   2100
                  Width           =   5000
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Index           =   8
                  Left            =   105
                  TabIndex        =   28
                  Top             =   2400
                  Width           =   5000
               End
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   2700
               Index           =   0
               Left            =   250
               ScaleHeight     =   2700
               ScaleWidth      =   6120
               TabIndex        =   16
               Top             =   0
               Width           =   6120
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  TabIndex        =   25
                  Top             =   600
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  TabIndex        =   24
                  Top             =   300
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   23
                  ToolTipText     =   "Doble-Click para Enviar un Mail"
                  Top             =   0
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   3
                  Left            =   0
                  TabIndex        =   22
                  Top             =   900
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   4
                  Left            =   0
                  TabIndex        =   21
                  Top             =   1200
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   5
                  Left            =   0
                  TabIndex        =   20
                  Top             =   1500
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   6
                  Left            =   0
                  TabIndex        =   19
                  Top             =   1800
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   7
                  Left            =   0
                  TabIndex        =   18
                  Top             =   2100
                  Width           =   6000
               End
               Begin VB.TextBox Text6 
                  Height          =   285
                  Index           =   8
                  Left            =   0
                  TabIndex        =   17
                  Top             =   2400
                  Width           =   6000
               End
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "3)"
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
               Index           =   2
               Left            =   -1365
               TabIndex        =   11
               Top             =   630
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "2)"
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
               Index           =   1
               Left            =   -1365
               TabIndex        =   10
               Top             =   330
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "1)"
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
               Left            =   -1365
               TabIndex        =   9
               Top             =   45
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "6)"
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
               Index           =   6
               Left            =   -1365
               TabIndex        =   8
               Top             =   1530
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "5)"
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
               Index           =   7
               Left            =   -1365
               TabIndex        =   7
               Top             =   1230
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "4)"
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
               Index           =   8
               Left            =   -1365
               TabIndex        =   6
               Top             =   930
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "9)"
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
               Index           =   9
               Left            =   -1365
               TabIndex        =   5
               Top             =   2430
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "8)"
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
               Index           =   10
               Left            =   -1365
               TabIndex        =   4
               Top             =   2130
               Width           =   1590
            End
            Begin VB.Label LEFE 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "7)"
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
               Index           =   11
               Left            =   -1365
               TabIndex        =   3
               Top             =   1830
               Width           =   1590
            End
         End
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo / Función - Sector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Persona de Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1905
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9285
         TabIndex        =   14
         Top             =   270
         Width           =   2010
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6210
         TabIndex        =   13
         Top             =   270
         Width           =   960
      End
   End
End
Attribute VB_Name = "traform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARRABOR%

Private Sub PICTURE1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub PICTURE1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture1.MousePointer = 7
   If ARRABOR% = 1 Then
      Picture1.Left = Picture1.Left + X
   End If
End Sub

Private Sub PICTURE1_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
   Picture1.MousePointer = 1
End Sub

Private Sub Command1_Click()
   '
   INTERVER = 50
   NULINE% = Text6.Count

   Text6(0).Parent.SetFocus
   Load Text6(NULINE%)
   Text6(NULINE%).Top = Text6(NULINE% - 1).Top + Text6(NULINE% - 1).Height + INTERVER
'AAA = Text6(NULINE%).Top
 'Text6(NULINE%).Top = 3500
   Text6(NULINE%).Parent = Text6(NULINE% - 1).Parent.hWnd
   '
   Text6(NULINE%).Visible = True
   'Text6(NULINE%).ZOrder 0
   
 '  Text6(NULINE%).Top = Text6(NULINE% - 1).Top + Text6(NULINE% - 1).Height
 '  Text6(NULINE%).Visible = True
   Text6(NULINE%).Text = ""

 
  LEFE(NULINE%).Visible = True
  LEFE(NULINE%) = "10"





End Sub

Private Sub Command2_Click()
Form1.Show 1
End Sub

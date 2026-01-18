VERSION 5.00
Begin VB.Form LIBARRI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación por Barrido"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   9810
      Left            =   8295
      ScaleHeight     =   9750
      ScaleWidth      =   1335
      TabIndex        =   17
      Top             =   0
      Width           =   1395
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
         Picture         =   "LIBARRI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "LIBARRI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Height          =   690
         Left            =   90
         Picture         =   "LIBARRI.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guardar "
         Top             =   4515
         Width           =   640
      End
      Begin VB.Label Label11 
         Height          =   330
         Left            =   105
         TabIndex        =   26
         Top             =   3885
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -330
         Picture         =   "LIBARRI.frx":075E
         Top             =   5400
         Width           =   1515
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5460
      TabIndex        =   16
      Top             =   2205
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4095
      TabIndex        =   15
      Top             =   2205
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6825
      TabIndex        =   14
      Top             =   2205
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5460
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4095
      TabIndex        =   9
      Top             =   1890
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6825
      TabIndex        =   8
      Top             =   1890
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
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
      Left            =   6825
      TabIndex        =   25
      Top             =   1365
      Width           =   1290
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
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
      Left            =   5460
      TabIndex        =   24
      Top             =   1365
      Width           =   1290
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Horas"
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
      Left            =   4095
      TabIndex        =   23
      Top             =   1365
      Width           =   1290
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Left            =   1050
      TabIndex        =   22
      Top             =   1365
      Width           =   825
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   210
      TabIndex        =   21
      Top             =   1365
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   8190
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   1050
      TabIndex        =   13
      Top             =   2235
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   2235
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1155
      TabIndex        =   11
      Top             =   630
      Width           =   6950
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Index           =   0
      Left            =   1050
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
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
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1155
      TabIndex        =   5
      Top             =   180
      Width           =   825
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3795
      TabIndex        =   4
      Top             =   180
      Width           =   4305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1890
      TabIndex        =   3
      Top             =   285
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
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
      Left            =   -315
      TabIndex        =   2
      Top             =   540
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Legajo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   -105
      TabIndex        =   1
      Top             =   105
      Width           =   1140
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3045
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "LIBARRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
  Label11.Caption = ""
  Unload Me
End Sub

Private Sub Command9_Click()
   Label11.Caption = "OK"
   Hide
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index))
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index))
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then
       Call Text1_KeyPress(Index, 13)
     ElseIf KeyCode = 38 Then
       Call Text1_KeyPress(Index, 11)
   End If
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then
       Call Text2_KeyPress(Index, 13)
     ElseIf KeyCode = 38 Then
       Call Text2_KeyPress(Index, 11)
   End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then
       Call Text3_KeyPress(Index, 13)
     ElseIf KeyCode = 38 Then
       Call Text3_KeyPress(Index, 11)
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
        If Index >= Text1.UBOUND Then
           Command9.SetFocus
           Exit Sub
        End If
        '
        If Text1(Index + 1).Visible = True Then
             Text1(Index + 1).SetFocus
          ElseIf Text2(Index + 1).Visible = True Then
             Text2(Index + 1).SetFocus
          ElseIf Text3(Index + 1).Visible = True Then
             Text3(Index + 1).SetFocus
        End If
      ElseIf KeyAscii% = 11 Then
        If Index <= 0 Then Exit Sub
        If Text1(Index - 1).Visible = True Then
             Text1(Index - 1).SetFocus
          ElseIf Text2(Index - 1).Visible = True Then
             Text2(Index - 1).SetFocus
          ElseIf Text3(Index - 1).Visible = True Then
             Text3(Index - 1).SetFocus
        End If
      ElseIf KeyAscii% = 27 Then
        Command9.SetFocus
   End If
   

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
        If Index >= Text1.UBOUND Then
           Command9.SetFocus
           Exit Sub
        End If
        '
        If Text1(Index + 1).Visible = True Then
             Text1(Index + 1).SetFocus
          ElseIf Text2(Index + 1).Visible = True Then
             Text2(Index + 1).SetFocus
          ElseIf Text3(Index + 1).Visible = True Then
             Text3(Index + 1).SetFocus
        End If
      ElseIf KeyAscii% = 11 Then
        If Index <= 0 Then Exit Sub
        If Text1(Index - 1).Visible = True Then
             Text1(Index - 1).SetFocus
          ElseIf Text2(Index - 1).Visible = True Then
             Text2(Index - 1).SetFocus
          ElseIf Text3(Index - 1).Visible = True Then
             Text3(Index - 1).SetFocus
        End If
      ElseIf KeyAscii% = 27 Then
        Command9.SetFocus
   End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
        If Index >= Text1.UBOUND Then
           Command9.SetFocus
           Exit Sub
        End If
        '
        If Text1(Index + 1).Visible = True Then
             Text1(Index + 1).SetFocus
          ElseIf Text2(Index + 1).Visible = True Then
             Text2(Index + 1).SetFocus
          ElseIf Text3(Index + 1).Visible = True Then
             Text3(Index + 1).SetFocus
        End If
      ElseIf KeyAscii% = 11 Then
        If Index <= 0 Then Exit Sub
        If Text1(Index - 1).Visible = True Then
             Text1(Index - 1).SetFocus
          ElseIf Text2(Index - 1).Visible = True Then
             Text2(Index - 1).SetFocus
          ElseIf Text3(Index - 1).Visible = True Then
             Text3(Index - 1).SetFocus
        End If
      ElseIf KeyAscii% = 27 Then
        Command9.SetFocus
   End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)
  Text1(Index).TEXT = CSTRING$(MKD$(Val(Text1(Index).TEXT)), 3, 7, 2, 2)
End Sub

Private Sub Text2_LostFocus(Index As Integer)
  Text2(Index).TEXT = CSTRING$(MKD$(Val(Text2(Index).TEXT)), 3, 10, 2, 2)
End Sub

Private Sub Text3_LostFocus(Index As Integer)
  Text3(Index).TEXT = CSTRING$(MKD$(Val(Text3(Index).TEXT)), 3, 10, 2, 2)
End Sub


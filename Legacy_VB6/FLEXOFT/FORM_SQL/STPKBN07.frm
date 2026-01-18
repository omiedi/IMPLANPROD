VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STPKBN04 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Configuración de Lectura de Código de Barra"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Parametros de Lectura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   4635
      Begin VB.PictureBox Picture999 
         Height          =   330
         Index           =   0
         Left            =   1470
         ScaleHeight     =   270
         ScaleWidth      =   1110
         TabIndex        =   18
         Top             =   525
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   12
         Top             =   2100
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   11
         Top             =   2100
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   10
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   9
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   8
         Top             =   1050
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   7
         Top             =   1050
         Width           =   750
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2625
         OleObjectBlob   =   "STPKBN07.frx":0000
         Top             =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "STPKBN07.frx":0234
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nº Tarjeta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   17
         Top             =   2100
         Width           =   1380
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
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
         Height          =   330
         Left            =   420
         TabIndex        =   16
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
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
         Height          =   330
         Left            =   420
         TabIndex        =   15
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HASTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   14
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DESDE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   13
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   4425
      Left            =   4890
      ScaleHeight     =   4365
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   -105
      Width           =   855
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   2625
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "STPKBN07.frx":0292
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "STPKBN07.frx":03DC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1005
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   105
         Picture         =   "STPKBN07.frx":06E6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Graba la Configuración "
         Top             =   1890
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -345
         Picture         =   "STPKBN07.frx":09F0
         Top             =   2835
         Width           =   1515
      End
   End
End
Attribute VB_Name = "STPKBN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Command1.Enabled = False
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Unload Me
    Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    Screen.MousePointer = 11
    If XVALO(Text1.TEXT) <= XVALO(Text2.TEXT) And XVALO(Text3.TEXT) <= XVALO(Text4.TEXT) And XVALO(Text5.TEXT) <= XVALO(Text6.TEXT) Then
        If XVALO(Text2.TEXT) - XVALO(Text1.TEXT) <= 15 And XVALO(Text4.TEXT) - XVALO(Text3.TEXT) <= 5 And XVALO(Text6.TEXT) - XVALO(Text5.TEXT) <= 3 Then
           Call GRACONTROL("REMITO", "CODESDE", TRIM$(Text1))
           Call GRACONTROL("REMITO", "CODHASTA", TRIM$(Text2))
           Call GRACONTROL("REMITO", "CANDESDE", TRIM$(Text3))
           Call GRACONTROL("REMITO", "CANHASTA", TRIM$(Text4))
           Call GRACONTROL("REMITO", "TARDESDE", TRIM$(Text5))
           Call GRACONTROL("REMITO", "TARHASTA", TRIM$(Text6))
           Screen.MousePointer = 1
           '
        Else
           'MsgBox ("El Número de Caracteres Sugerido es Superior al Permitodo")
           Screen.MousePointer = 1
          Call MENSERR(24, "El Número de Caracteres Sugerido es Superior al Permitido")
          
        End If
    Else
    
         'MsgBox ("El Campo DESDE no Puede ser Superior al Campo HASTA")
         Screen.MousePointer = 1
        Call MENSERR(24, "El Campo DESDE no Puede ser Superior al Campo HASTA")
    End If
      
    
    
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)

    Text1 = CONTROL$("REMITO", "CODESDE")
    Text2 = CONTROL$("REMITO", "CODHASTA")
    Text3 = CONTROL$("REMITO", "CANDESDE")
    Text4 = CONTROL$("REMITO", "CANHASTA")
    Text5 = CONTROL$("REMITO", "TARDESDE")
    Text6 = CONTROL$("REMITO", "TARHASTA")
    '
End Sub
